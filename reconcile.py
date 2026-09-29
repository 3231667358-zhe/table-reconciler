#!/usr/bin/env python3
"""
Table Reconciler H1 —— Excel / CSV 两表匹配与对账差异检测

把左右两张表按**用户显式指定的**匹配键对齐，输出一份对账结果工作簿 + 一份文本报告：

    左表（例如客户对账单）        右表（例如内部台账）
        订单号 = 单号    ←  匹配键，必须显式指定（单列，一对一）
        应收金额 = 金额  ←  比较字段，必须显式指定
              ↓
    reconciliation.xlsx
        summary      六项统计 + 本次运行参数回显
        matched      键两边都有，且所有比较字段都一致
        differences  键两边都有，但有字段值不同（明确给出左值和右值）
        only_left    只有左表有的记录
        only_right   只有右表有的记录
    report.txt       运行参数、文件指纹、统计、公式策略、警告

============================================================================
H1 的硬边界（写死在代码里，不靠自觉）
============================================================================

1. 匹配键和比较字段**必须由用户显式指定**，程序不猜、不自动挑列。
2. **只支持单列匹配键**（左一列 = 右一列）。复合键留到 H2。
3. **只支持一对一匹配**。任意一侧匹配键重复 → 硬失败，绝不偷偷做一对多。
4. 匹配键为空 → **无条件硬失败**（没有任何参数可以绕过）。
   空键报错必须讲清：是左表还是右表、空键有多少行、对应哪些原始行号。
   理由：把空键排除掉再产出结果，会得到一份"不完整但看起来成功"的对账结果 ——
   那正是 H1 最不能接受的东西。
5. **匹配键按文本处理**：`00123` 绝不会被当成 `123`。
6. **默认口径就是最严格的**：除"去除首尾空白"外不做任何变换 ——
   不忽略大小写、不做全角半角转换、不去千分位、不做数值容差、不做日期识别、不做相似匹配。
   `ABC` 和 `abc` 就是不同的键。
   宁可把业务上本来相同的数据列进 differences，也绝不因为过度标准化产生"假匹配"。
   任何宽松规则都属于 H2，将来加也必须显式开启，不能默认开启。
7. 支持 `.csv` / `.txt` / `.xlsx`；`.xls` / `.xlsm` / `.xlsb` 一律硬失败。
8. **原始文件绝对只读**。写入路径（--output / --report / work/staging）都不允许指向
   --left / --right 那两个文件，按 resolve() 判，与 --force 无关。
9. 输出**先在 work/staging/ 完整生成并校验**，通过后两个文件一起提交到 output/；
   提交中任一步失败就回滚，output/ 绝不留下半套结果。
10. 输出已存在 → 默认拒绝；`--force` 才重跑，旧结果**归档**到 work/previous/（不删除、
    不覆盖已有历史）。判断"上次成功了吗"就看 output/ 里有没有 reconciliation.xlsx。

============================================================================
Excel 公式策略（明确设计，不无声使用可能过期的缓存）
============================================================================

- 读 Excel 一律 `data_only=False`：只拿得到公式串，**结构上不可能误用公式缓存值**。
- 公式出现在**参与列**（匹配键列 / 比较字段列）→ **硬失败**并列出坐标。
  理由：那时我们手里的"值"其实是 `=A1&B1` 这种公式文本，拿它当键或比较值必然错。
- 公式出现在**不参与的列** → 不读、不影响结果，但报告里明确写出"忽略了 N 处"。
- 以 `=` 开头的**文本**单元格（不是公式）→ 当文本正常接受，但报告里给出警告。
  这是对姊妹项目策略的有意收窄：它们在**重写**工作簿时分不清公式和文本，所以整文件拒绝；
  本工具只读输入、另建工作簿输出，不存在把文本误写成公式的风险。
- 我们自己写出的 xlsx 只写文本和数字，**永不写公式**。

用法示例：
    python3 reconcile.py \\
        --left  客户对账单.csv --left-header-row 3 --left-key 订单号 \\
        --right 内部台账.xlsx --right-sheet 内部台账 --right-key 单号 \\
        --compare 应收金额=金额,状态

看详细帮助：python3 reconcile.py -h
"""

import argparse
import csv
import datetime
import hashlib
import os
import re
import shutil
import sys
from decimal import Decimal
from pathlib import Path

VERSION_NAME = "Table Reconciler H1"

PROJECT_ROOT = Path(__file__).resolve().parent

# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------

CSV_SUFFIXES = {".csv", ".txt"}
EXCEL_SUFFIXES = {".xlsx"}
SUPPORTED_SUFFIXES = CSV_SUFFIXES | EXCEL_SUFFIXES
# 明确不支持：见到就硬失败（对账只处理用户指定的两个文件，没有"跳过继续"的语义）
UNSUPPORTED_SUFFIXES = {".xls", ".xlsm", ".xlsb"}

# 读 CSV 按顺序试编码：先 UTF-8（含 BOM），再 GB18030（GBK 超集，兼容老系统导出的中文表）
CSV_INPUT_ENCODINGS = ("utf-8-sig", "gb18030")
REPLACEMENT_CHAR = "�"

SHEET_SUMMARY = "summary"
SHEET_MATCHED = "matched"
SHEET_DIFFERENCES = "differences"
SHEET_ONLY_LEFT = "only_left"
SHEET_ONLY_RIGHT = "only_right"
DATA_SHEETS = (SHEET_MATCHED, SHEET_DIFFERENCES, SHEET_ONLY_LEFT, SHEET_ONLY_RIGHT)

COL_KEY = "匹配键"
COL_DIFF_FIELDS = "差异字段"
COL_ROW_LEFT = "左表行号"
COL_ROW_RIGHT = "右表行号"

OUTPUT_XLSX_NAME = "reconciliation.xlsx"
REPORT_NAME = "report.txt"
FAILED_REPORT_NAME = "failed_report.txt"

WORK_SUBDIR_STAGING = "staging"
WORK_SUBDIR_REPORTS = "reports"
WORK_SUBDIR_PREVIOUS = "previous"
# work/ 下由本工具独占、每次运行前清空的子目录。
# previous/ 不在其中 —— 它是历史留存，任何清理逻辑都不许删它。
CLEARABLE_WORK_SUBDIRS = (WORK_SUBDIR_STAGING, WORK_SUBDIR_REPORTS)

REPORT_TITLE = VERSION_NAME + " 对账报告"

# 「看不出来但会让键对不上」的字符。注意 U+00A0 / U+3000 本身是 Unicode 空白，
# str.strip() 会去掉首尾的它们；这里要抓的是**出现在中间**的，以及 strip() 不处理的零宽字符。
INVISIBLE_CHARS = " 　​‌‍⁠﻿"
INVISIBLE_NAMES = {
    " ": "U+00A0 不换行空格",
    "　": "U+3000 全角空格",
    "​": "U+200B 零宽空格",
    "‌": "U+200C 零宽非连接符",
    "‍": "U+200D 零宽连接符",
    "⁠": "U+2060 词连接符",
    "﻿": "U+FEFF 零宽不换行空格(BOM)",
}

# 输出 xlsx 的时间戳固定住：同样的输入产出同样的文件，便于 diff 和回归测试。
# 真实运行时间记录在 report.txt 里。
FIXED_XLSX_TIMESTAMP = datetime.datetime(2000, 1, 1, 0, 0, 0)


class ReconcileError(Exception):
    """本工具能对用户说清楚的问题。捕获后只打印提示，不打印 Python 堆栈。"""


# --------------------------------------------------------------------------
# 文本渲染与规范化（对账正确性的地基，改这里之前先读注释）
# --------------------------------------------------------------------------

def render_cell_text(value):
    """
    把一个单元格的值**无损地**渲染成文本。这是匹配键和比较值的唯一入口。

    规则（唯一目标：不让渲染本身制造假差异或假匹配）：
      - None         → ""             （空键由调用方按"空键"处理，不在这里吞掉）
      - str          → **原样返回**   （不去空白、不转全角半角、不动大小写）
      - bool         → TRUE / FALSE
      - int          → 十进制文本
      - float 整数值 → 整数文本：1200.0 写成 "1200"
                       这不是"数值容差"，而是把同一个数写成同一种形式；
                       否则 Excel 里的 1200 与文本 "1200" 会对不上，那是纯粹的假差异。
      - float 非整数 → 用 Decimal 定点展开，避免 "1e-05" 这种科学计数法；
                       尾差照样保留（0.1+0.2 → "0.30000000000000004"，严格比较会报差异，
                       这正是 H1 要的：不做数值容差）
      - 日期/时间    → ISO 文本；零点整的 datetime 只写日期（"2026-08-01"）
      - 其它         → str(value)
    """
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value != value:  # NaN
            return "NaN"
        if value in (float("inf"), float("-inf")):
            return "inf" if value > 0 else "-inf"
        if value.is_integer():
            return str(int(value))
        return format(Decimal(repr(value)), "f")
    if isinstance(value, datetime.datetime):
        if value.time() == datetime.time(0, 0):
            return value.date().isoformat()
        return value.isoformat(sep=" ")
    if isinstance(value, (datetime.date, datetime.time)):
        return value.isoformat()
    return str(value)


def trim_text(text):
    """
    去掉首尾空白 —— H1 **唯一**允许的匹配/比较规范化。

    str.strip() 是 Unicode 感知的，会去掉 U+00A0（不换行空格）、U+3000（全角空格）
    这类"看不见的"空白，所以从系统里粘贴出来的带空白键能对上。
    它不会去掉零宽字符 —— 那些字符一律**保留**（可能是数据的一部分），只检测并告警。
    """
    return text.strip()


def normalize_match_text(value):
    """匹配与比较共用的规范化：渲染成文本 + 去首尾空白。就这两步，没有第三步。"""
    return trim_text(render_cell_text(value))


def find_invisible_chars(text):
    """返回 text 里出现的不可见字符（去重，保持出现顺序）。"""
    found = []
    for char in text:
        if char in INVISIBLE_CHARS and char not in found:
            found.append(char)
    return found


def describe_char(char):
    return INVISIBLE_NAMES.get(char, "U+%04X" % ord(char))


def sha256_of(path):
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_openpyxl():
    """按需导入 openpyxl，没装就给一条能直接照抄的命令。"""
    try:
        import openpyxl
    except ImportError:
        raise ReconcileError(
            "处理 .xlsx 需要 openpyxl，但当前 Python 里没有这个库。\n"
            "  Ubuntu 安装（先跑这一条）：sudo apt install python3-openpyxl\n"
            "  装完重新运行即可。只想对 CSV 的话，让 --left/--right 都指向 CSV 就行。"
        )
    return openpyxl


# --------------------------------------------------------------------------
# 读表
# --------------------------------------------------------------------------

class Table:
    """一张读进来的表，外加为"明确报告"采集的各种痕迹。"""

    def __init__(self, role, path, kind):
        self.role = role                  # "左" / "右"，用于报错信息
        self.path = path
        self.kind = kind                  # "csv" / "excel"
        self.encoding = None              # CSV 实际使用的编码
        self.delimiter = None
        self.sheet_title = None
        self.sheet_names = []
        self.header_row = 1
        self.header = []                  # 原始表头文本
        self.data = []                    # [(物理行号, [渲染后的原始值...])]
        self.blank_rows_skipped = 0
        self.ragged_rows = 0              # 列数与表头不一致的数据行数
        self.formulas_by_col = {}         # {列下标: [(坐标, 行号, 公式串), ...]}
        self.equal_text_cells = []        # [("=" 开头文本单元格的 (坐标, 行号, 文本))]
        self.hidden_rows = set()
        self.has_autofilter = False
        self.replacement_rows = []        # 解出 U+FFFD 的行号（编码猜错）
        # 每一列里"数值单元格 / 文本单元格"各有多少个 —— 用来诊断
        # "比较字段两边类型不一致"（严格文本比较下这类差异几乎全是写法不同造成的）
        self.numeric_by_col = {}
        self.text_by_col = {}
        self.bytes_size = 0
        self.sha256 = ""

    def note_cell_type(self, column, raw_value):
        if isinstance(raw_value, bool) or raw_value is None:
            return
        if isinstance(raw_value, (int, float)):
            self.numeric_by_col[column] = self.numeric_by_col.get(column, 0) + 1
        elif isinstance(raw_value, str) and raw_value.strip() != "":
            self.text_by_col[column] = self.text_by_col.get(column, 0) + 1

    @property
    def flag_prefix(self):
        return "left-" if self.role == "左" else "right-"


def detect_kind(path, role):
    """按后缀判断类型；不支持的直接硬失败。"""
    suffix = path.suffix.lower()
    if suffix in UNSUPPORTED_SUFFIXES:
        hint = ""
        if suffix == ".xlsm":
            hint = "\\n（宏文件请先在 Excel 里另存为 .xlsx —— 本工具会重建表，宏无法安全保留）"
        raise ReconcileError(
            "{0}表是不支持的格式：{1}\\n"
            "H1 支持：{2}\\n"
            "不支持：{3}{4}".format(
                role, path.name,
                "、".join(sorted(SUPPORTED_SUFFIXES)),
                "、".join(sorted(UNSUPPORTED_SUFFIXES)),
                hint,
            )
        )
    if suffix in EXCEL_SUFFIXES:
        return "excel"
    if suffix in CSV_SUFFIXES:
        return "csv"
    raise ReconcileError(
        "{0}表是不支持的格式：{1}\\nH1 支持：{2}".format(
            role, suffix or "（没有后缀）", "、".join(sorted(SUPPORTED_SUFFIXES))
        )
    )


def read_csv_table(path, delimiter, header_row, role):
    """读 CSV/TXT，返回 Table。"""
    table = Table(role, path, "csv")
    table.header_row = header_row
    table.delimiter = delimiter
    if len(delimiter) != 1:
        raise ReconcileError(
            "{}表的分隔符必须是**一个**字符，现在给的是 {!r}。".format(role, delimiter)
        )

    last_error = None
    text = None
    for encoding in CSV_INPUT_ENCODINGS:
        try:
            text = path.read_text(encoding=encoding)
            table.encoding = encoding
            break
        except UnicodeDecodeError as exc:
            last_error = exc
    if text is None:
        raise ReconcileError(
            "读不了这个文件：{}\n按 {} 依次尝试都解码失败。\n原始报错：{}".format(
                path, " / ".join(CSV_INPUT_ENCODINGS), last_error
            )
        )

    # 记下每条记录在文件里的物理起始行号（引号里含换行时，一条记录会跨多行）
    reader = csv.reader(text.splitlines(True), delimiter=delimiter)
    records = []
    consumed = 0
    for record in reader:
        records.append((consumed + 1, record))
        consumed = reader.line_num

    if len(records) < header_row:
        raise ReconcileError(
            "{}表的表头行号超出范围：--{}header-row {}，但该文件一共只有 {} 行。".format(
                role, table.flag_prefix, header_row, len(records)
            )
        )

    table.header = [render_cell_text(value) for value in records[header_row - 1][1]]
    width = len(table.header)
    for _, record in records[header_row:]:
        width = max(width, len(record))

    for line_no, record in records[header_row:]:
        if len(record) != len(table.header):
            table.ragged_rows += 1
        values = [render_cell_text(value) for value in record]
        if any(REPLACEMENT_CHAR in value for value in values):
            table.replacement_rows.append(line_no)
        padded = values + [""] * (width - len(values))
        if all(value == "" for value in padded):
            table.blank_rows_skipped += 1
            continue
        for index, raw_value in enumerate(record):
            table.note_cell_type(index, raw_value)
        table.data.append((line_no, padded))
    return table


def read_excel_table(path, sheet_name, header_row, role):
    """
    读 .xlsx 的指定工作表，返回 Table。

    一律 data_only=False：只拿得到公式串，从结构上杜绝"用到可能过期的公式缓存值"。
    文件句柄由我们自己持有 —— openpyxl 的 Workbook.close() 在非 read_only 模式下是空操作，
    不自己拿着句柄就会漏文件描述符（ResourceWarning），这是姊妹项目踩过的坑。
    """
    openpyxl = load_openpyxl()
    table = Table(role, path, "excel")
    table.header_row = header_row

    try:
        handle = path.open("rb")
    except OSError as exc:
        raise ReconcileError("打不开这个 Excel：{}\n原始报错：{}".format(path, exc))

    with handle:
        try:
            workbook = openpyxl.load_workbook(handle, data_only=False)
        except Exception as exc:
            raise ReconcileError(
                "打不开这个 Excel：{}\n原始报错：{}\n"
                "（如果它其实是 .xls / .xlsb，请先在 Excel 里另存为 .xlsx）".format(path, exc)
            )
        try:
            table.sheet_names = list(workbook.sheetnames)
            if not table.sheet_names:
                raise ReconcileError("这个 Excel 里一张工作表都没有：{}".format(path))
            if sheet_name is None:
                title = table.sheet_names[0]
            else:
                if sheet_name not in table.sheet_names:
                    raise ReconcileError(
                        "{}表里找不到工作表「{}」。\n"
                        "这个文件里的工作表有：{}\n"
                        "请用 --{}sheet 指定正确的工作表名。".format(
                            role, sheet_name,
                            "、".join("「{}」".format(name) for name in table.sheet_names),
                            table.flag_prefix,
                        )
                    )
                title = sheet_name
            table.sheet_title = title
            sheet = workbook[title]

            table.hidden_rows = {
                number
                for number, dimension in sheet.row_dimensions.items()
                if getattr(dimension, "hidden", False)
            }
            table.has_autofilter = bool(getattr(sheet.auto_filter, "ref", None))

            raw_rows = []
            for cells in sheet.iter_rows():
                if not cells:
                    continue
                physical_row = cells[0].row
                raw_rows.append((physical_row, cells))
                for index, cell in enumerate(cells):
                    value = cell.value
                    if isinstance(value, str) and value.startswith("="):
                        if cell.data_type == "f":
                            table.formulas_by_col.setdefault(index, []).append(
                                (cell.coordinate, physical_row, value)
                            )
                        else:
                            table.equal_text_cells.append(
                                (cell.coordinate, physical_row, value)
                            )

            if len(raw_rows) < header_row:
                raise ReconcileError(
                    "{}表的工作表「{}」行数不够：--{}header-row {}，但整个表只有 {} 行。".format(
                        role, title, table.flag_prefix, header_row, len(raw_rows)
                    )
                )

            table.header = [
                render_cell_text(cell.value) for cell in raw_rows[header_row - 1][1]
            ]
            width = len(table.header)
            for _, cells in raw_rows[header_row:]:
                width = max(width, len(cells))

            for physical_row, cells in raw_rows[header_row:]:
                values = [render_cell_text(cell.value) for cell in cells]
                if len(values) != len(table.header):
                    table.ragged_rows += 1
                padded = values + [""] * (width - len(values))
                if all(value == "" for value in padded):
                    table.blank_rows_skipped += 1
                    continue
                for index, cell in enumerate(cells):
                    table.note_cell_type(index, cell.value)
                table.data.append((physical_row, padded))
        finally:
            workbook.close()
    return table


# --------------------------------------------------------------------------
# 列表头与列定位
# --------------------------------------------------------------------------

def build_header_index(header):
    """把表头整理成 {去空白后的列名: [列下标...]}，并给出重名描述。"""
    index = {}
    for position, raw in enumerate(header):
        index.setdefault(trim_text(raw), []).append(position)
    duplicates = {
        name: positions for name, positions in index.items() if name and len(positions) > 1
    }
    if not duplicates:
        return index, ""
    detail = "；".join(
        "「{}」在第 {} 列".format(name, "、".join(str(p + 1) for p in positions))
        for name, positions in duplicates.items()
    )
    return index, detail


def resolve_column(table, index, name, what):
    """
    按用户给的列名定位列下标。找不到 / 有歧义 → 硬失败，并把该文件真实表头打出来
    （用引号包起来，方便看出多余空格）。
    """
    where = "{}{}".format(
        table.role, "表" if table.sheet_title is None else "表（工作表「{}」）".format(table.sheet_title)
    )
    if name not in index:
        shown = "、".join(
            repr(trim_text(raw)) if trim_text(raw) else "（空表头）" for raw in table.header
        )
        hint = ""
        header_formulas = [
            item
            for hits in table.formulas_by_col.values()
            for item in hits
            if item[1] <= table.header_row
        ]
        if header_formulas:
            hint = (
                "\n注意：表头行里有公式单元格（例如 {}），列名可能是公式算出来的。"
                "\n      请先在 Excel 里把表头选择性粘贴成数值再试。".format(header_formulas[0][0])
            )
        raise ReconcileError(
            "{0}的{1}找不到指定的列：{2!r}\n"
            "该文件第 {3} 行（表头行）实际读到的是：{4}{5}".format(
                where, what, name, table.header_row, shown, hint
            )
        )
    positions = index[name]
    if len(positions) > 1:
        raise ReconcileError(
            "{0}的{1}列名 {2!r} 在表头里出现了 {3} 次（第 {4} 列），无法确定用哪一列。\n"
            "对账必须精确到列，请先在源文件里把重复的列名改掉。".format(
                where, what, name, len(positions),
                "、".join(str(p + 1) for p in positions),
            )
        )
    return positions[0]


def parse_compare_specs(raw_items):
    """
    解析 --compare。两种写法：
        金额            两边列名相同
        应收金额=金额    左边列名=右边列名
    返回 [(左列名, 右列名, 展示名)]，顺序即用户给的顺序。
    """
    specs = []
    seen = set()
    for raw in raw_items:
        for piece in raw.replace("，", ",").split(","):
            item = piece.strip()
            if not item:
                continue
            if "=" in item:
                left_name, _, right_name = item.partition("=")
                left_name, right_name = left_name.strip(), right_name.strip()
                if not left_name or not right_name:
                    raise ReconcileError(
                        "--compare 的写法是「列名」或「左列名=右列名」，这一项看不懂：{!r}\n"
                        "（列名里如果真的有 = 号，请先在源文件里把列名改掉）".format(item)
                    )
            else:
                left_name = right_name = item
            key = (left_name, right_name)
            if key in seen:
                raise ReconcileError(
                    "--compare 里重复指定了同一个比较字段：{!r}\n"
                    "重复比较没有意义，而且会让「{}」列出现重名。".format(item, COL_DIFF_FIELDS)
                )
            seen.add(key)
            display = left_name if left_name == right_name else "{}={}".format(left_name, right_name)
            specs.append((left_name, right_name, display))
    if not specs:
        raise ReconcileError("--compare 至少要指定一个比较字段。")
    return specs


# --------------------------------------------------------------------------
# 建键索引（空键 / 重复键都在这里拦下）
# --------------------------------------------------------------------------

class KeyIndex:
    def __init__(self):
        self.map = {}            # 规范化键 -> (物理行号, 行值, 键原值)
        self.first_row = {}      # 规范化键 -> 第一次出现的物理行号
        self.duplicates = {}     # 规范化键 -> [物理行号, ...]（升序，含首次出现）
        self.empty_rows = []     # [(物理行号, 键原值)]
        self.invisible_hits = [] # [(物理行号, [字符...])]
        self.trimmed_rows = []   # [(物理行号, 键原值, 规范化后)]
        self.numeric_keys = 0    # 键来自数值单元格的条数
        self.digit_text_keys = 0 # 键是"纯数字文本"的条数（比如 00123 这类）


def build_key_index(table, column):
    """
    按匹配键建索引，同时收集两类必须拦下的问题：空键、重复键。

    重复判定在**规范化之后**的键上做 —— "123" 和 " 123 " 去空白后是同一个键，
    不在这层判重就会变成悄悄的一对多。
    """
    index = KeyIndex()
    for physical_row, values in table.data:
        raw_value = values[column]
        text = render_cell_text(raw_value)
        normalized = trim_text(text)

        if normalized == "":
            index.empty_rows.append((physical_row, text))
            continue

        invisible = find_invisible_chars(normalized)
        if invisible:
            index.invisible_hits.append((physical_row, invisible))
        if text != normalized:
            index.trimmed_rows.append((physical_row, text, normalized))

        if isinstance(raw_value, bool):
            pass
        elif isinstance(raw_value, (int, float)):
            index.numeric_keys += 1
        elif isinstance(raw_value, str) and raw_value.strip().isdigit():
            index.digit_text_keys += 1

        if normalized in index.map:
            index.duplicates.setdefault(normalized, [index.first_row[normalized]])
            if physical_row not in index.duplicates[normalized]:
                index.duplicates[normalized].append(physical_row)
            continue

        index.map[normalized] = (physical_row, values, text)
        index.first_row[normalized] = physical_row

    for rows in index.duplicates.values():
        rows.sort()
    return index


def describe_key_problems(left_index, right_index, limit=5, extra_row_limit=45):
    """
    把两侧的空键、重复键拼成一段人话报错；没问题就返回空字符串。

    空键和重复键**都**无条件硬失败 —— H1 是严格的一对一对账工具，
    没有任何参数能把它们降级成"排除掉继续跑"。
    理由：排掉空键再产出结果，会得到一份"不完整但看起来成功"的对账结果，
    那正是 H1 最不能接受的东西。

    空键这段报错必须讲清三件事，让人能直接回到源文件改：
      1. 是左表还是右表
      2. 空键有几行
      3. 对应的原始行号（列不下时也会把剩余行号补上）
    """
    parts = []

    for role, index in (("左", left_index), ("右", right_index)):
        if not index.empty_rows:
            continue
        total = len(index.empty_rows)
        shown = "、".join(
            "第 {} 行（原值 {!r}）".format(row, raw) for row, raw in index.empty_rows[:limit]
        )
        more = ""
        if total > limit:
            rest = index.empty_rows[limit:limit + extra_row_limit]
            more = "；另外还有 {} 行，行号：{}".format(
                total - limit, "、".join(str(row) for row, _raw in rest)
            )
            if total > limit + extra_row_limit:
                more += " ……"
        parts.append(
            "{}表有 {} 行的匹配键是空的：{}{}\n"
            "  常见原因：表尾的「合计」行、表头下面的空行、或者键列用了合并单元格"
            "（合并后只有左上角那个格子有值）。".format(role, total, shown, more)
        )

    for role, index in (("左", left_index), ("右", right_index)):
        if not index.duplicates:
            continue
        items = sorted(index.duplicates.items())
        shown = "；".join(
            "{!r} 出现在第 {} 行".format(key, "、".join(str(row) for row in rows))
            for key, rows in items[:limit]
        )
        more = ""
        if len(items) > limit:
            more = "；另外还有 {} 个键".format(len(items) - limit)
        parts.append(
            "{}表有 {} 个匹配键是重复的：{}{}\n"
            "  对账是一对一的：一个键在一侧出现两次，就无法确定该和对面哪一行比。\n"
            "  本工具不会替你挑一行、也不会做成一对多，所以直接停下。".format(
                role, len(items), shown, more)
        )

    if not parts:
        return ""
    return (
        "匹配键检查没通过，任务已中止，没有生成任何输出文件。\n\n"
        + "\n\n".join(parts)
        + "\n\n怎么改：修好源文件（删掉合计行、去掉重复记录、换成真正唯一的键列）再重新运行。"
        "\n键列里如果确实有合并单元格，请先取消合并并把值补齐。"
    )


# --------------------------------------------------------------------------
# 对账
# --------------------------------------------------------------------------

def reconcile(left_index, right_index, compare):
    """
    一对一匹配并分类。每一项都把整行原值一起带上，写表时直接用，不需要旁路查表。

    返回：
      matched     [(键, 左行号, 右行号, 左行值, 右行值), ...]
      differences [(键, 左行号, 右行号, [(展示名, 左值, 右值)...], 左行值, 右行值), ...]
      only_left   [(键, 左行号, 左行值), ...]
      only_right  [(键, 右行号, 右行值), ...]
    """
    matched = []
    differences = []
    only_left = []
    only_right = []
    field_diff_counts = {}   # 展示名 -> 该字段造成差异的次数（只用于诊断提示）

    for key, (left_row, left_values, _raw) in left_index.map.items():
        if key not in right_index.map:
            only_left.append((key, left_row, left_values))
            continue
        right_row, right_values, _raw2 = right_index.map[key]
        field_diffs = []
        for left_col, right_col, _left_name, _right_name, display in compare:
            left_text = normalize_match_text(left_values[left_col])
            right_text = normalize_match_text(right_values[right_col])
            if left_text != right_text:
                field_diffs.append((display, left_text, right_text))
                field_diff_counts[display] = field_diff_counts.get(display, 0) + 1
        if field_diffs:
            differences.append((key, left_row, right_row, field_diffs, left_values, right_values))
        else:
            matched.append((key, left_row, right_row, left_values, right_values))

    for key, (right_row, right_values, _raw) in right_index.map.items():
        if key not in left_index.map:
            only_right.append((key, right_row, right_values))

    # 输出顺序按物理行号排，人回原始文件对照时最顺手
    matched.sort(key=lambda item: item[1])
    differences.sort(key=lambda item: item[1])
    only_left.sort(key=lambda item: item[1])
    only_right.sort(key=lambda item: item[1])
    return {
        "matched": matched,
        "differences": differences,
        "only_left": only_left,
        "only_right": only_right,
        "field_diff_counts": field_diff_counts,
    }


# --------------------------------------------------------------------------
# 写工作簿
# --------------------------------------------------------------------------

def append_row(sheet, values):
    """
    往工作表追加一行，并保证「以 = 开头的字符串」不会变成公式。

    openpyxl 默认会把以 = 开头的字符串当成公式写进文件 —— 那会把用户的原始文本
    （例如订单号 "=A-1"）变成公式，Excel 打开显示 #NAME?，也直接违反
    「输出里永不写公式」的承诺。这里对这种单元格强制标成文本类型，原样写出。
    这条和读入侧的策略是对称的：读的时候我们把 "=" 开头的文本当文本，写的时候也一样。
    """
    sheet.append(list(values))
    for cell in sheet[sheet.max_row]:
        if isinstance(cell.value, str) and cell.value.startswith("="):
            cell.data_type = "s"


def build_workbook(result, left, right, left_key_col, right_key_col, compare, config_rows):
    """
    生成对账结果工作簿。只写文本和数字，**永不写公式** ——
    输出是我们全新造的表，不存在公式引用被破坏的问题。

    固定五张表：summary / matched / differences / only_left / only_right。
    （H1 没有 excluded 表：匹配键为空是硬失败，不会产出"排掉一部分"的结果。）
    """
    openpyxl = load_openpyxl()
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    # 固定时间戳：同样的输入产出同样的文件（运行时间记录在 report.txt 和 summary 表里）
    workbook.properties.created = FIXED_XLSX_TIMESTAMP
    workbook.properties.modified = FIXED_XLSX_TIMESTAMP

    summary = workbook.create_sheet(SHEET_SUMMARY)
    append_row(summary, ["项目", "内容"])
    for name, value in config_rows:
        append_row(summary, [name, value])

    compare_pair_header = []
    for _left_col, _right_col, left_name, right_name, _display in compare:
        compare_pair_header.extend(["{}_左".format(left_name), "{}_右".format(right_name)])

    # matched：匹配成功的行，比较字段两边都摆出来（不藏，便于发现"规范化后恰好相等"）
    sheet = workbook.create_sheet(SHEET_MATCHED)
    append_row(sheet, [COL_KEY, COL_ROW_LEFT, COL_ROW_RIGHT] + compare_pair_header)
    for key, left_row, right_row, left_values, right_values in result["matched"]:
        row = [key, left_row, right_row]
        for left_col, right_col, _ln, _rn, _display in compare:
            row.extend([left_values[left_col], right_values[right_col]])
        append_row(sheet, row)

    # differences：明确给出左值和右值
    sheet = workbook.create_sheet(SHEET_DIFFERENCES)
    append_row(sheet, [COL_KEY, COL_DIFF_FIELDS, COL_ROW_LEFT, COL_ROW_RIGHT] + compare_pair_header)
    for key, left_row, right_row, field_diffs, left_values, right_values in result["differences"]:
        row = [key, "、".join(diff[0] for diff in field_diffs), left_row, right_row]
        for left_col, right_col, _ln, _rn, _display in compare:
            row.extend([left_values[left_col], right_values[right_col]])
        append_row(sheet, row)

    # only_left / only_right：带上该侧自己的列名和整行参与列的值
    # （键原值单独出一列：它和规范化后的「匹配键」可能不同，摆出来才不会有信息被藏掉）
    sheet = workbook.create_sheet(SHEET_ONLY_LEFT)
    append_row(sheet, [COL_KEY, COL_ROW_LEFT, trim_text(left.header[left_key_col])]
               + [trim_text(left.header[left_col]) for left_col, _rc, _ln, _rn, _d in compare])
    for key, row_number, values in result["only_left"]:
        append_row(sheet, [key, row_number, values[left_key_col]]
                   + [values[left_col] for left_col, _rc, _ln, _rn, _d in compare])

    sheet = workbook.create_sheet(SHEET_ONLY_RIGHT)
    append_row(sheet, [COL_KEY, COL_ROW_RIGHT, trim_text(right.header[right_key_col])]
               + [trim_text(right.header[right_col]) for _lc, right_col, _ln, _rn, _d in compare])
    for key, row_number, values in result["only_right"]:
        append_row(sheet, [key, row_number, values[right_key_col]]
                   + [values[right_col] for _lc, right_col, _ln, _rn, _d in compare])

    return workbook


def write_workbook(path, workbook):
    """把工作簿写到磁盘。文件句柄自己持有，任何路径下都会关闭。"""
    try:
        handle = path.open("wb")
    except OSError as exc:
        raise ReconcileError("写不了结果文件：{}\n原始报错：{}".format(path, exc))
    with handle:
        workbook.save(handle)


# --------------------------------------------------------------------------
# 报告
# --------------------------------------------------------------------------

def get_counts(left, right, result):
    """六项统计。报告、summary 表、提交前校验都用这一份，避免三处算法漂移。"""
    return {
        "left_rows": len(left.data),
        "right_rows": len(right.data),
        "matched": len(result["matched"]),
        "differences": len(result["differences"]),
        "only_left": len(result["only_left"]),
        "only_right": len(result["only_right"]),
    }


def build_report(stamp, left, right, left_arg, right_arg, left_key_name, right_key_name,
                 compare, result, warnings, formula_notes,
                 output_xlsx, output_report):
    counts = get_counts(left, right, result)
    lines = [
        REPORT_TITLE,
        "=" * 68,
        "版本：{}".format(VERSION_NAME),
        "运行时间：{}".format(stamp),
        "",
        "一、运行参数",
        "  左表：{}".format(left_arg),
        "  右表：{}".format(right_arg),
        "  匹配键：左「{}」= 右「{}」（单列，一对一）".format(left_key_name, right_key_name),
        "  比较字段：{}".format(
            "、".join(
                "左「{}」= 右「{}」".format(left_name, right_name)
                for _lc, _rc, left_name, right_name, _display in compare
            )
        ),
        "  左表工作表：{}".format(left.sheet_title or "（CSV，无工作表）"),
        "  右表工作表：{}".format(right.sheet_title or "（CSV，无工作表）"),
        "  左表表头行：第 {} 行".format(left.header_row),
        "  右表表头行：第 {} 行".format(right.header_row),
        "  左表解析：{}".format(
            "编码 {} / 分隔符 {!r}".format(left.encoding, left.delimiter)
            if left.kind == "csv" else "Excel（openpyxl 直读）"
        ),
        "  右表解析：{}".format(
            "编码 {} / 分隔符 {!r}".format(right.encoding, right.delimiter)
            if right.kind == "csv" else "Excel（openpyxl 直读）"
        ),
        "  匹配/比较口径：严格 —— 只去除首尾空白，其它一律不动",
        "      不做大小写折叠、不做全角半角转换、不去千分位、不做数值容差、",
        "      不做日期识别、不做相似匹配；00123 与 123 视为不同，ABC 与 abc 视为不同。",
        "      唯一的渲染规则：数值单元格 1200.0 写成 1200"
        "（同一个数写成同一种形式，不是容差）",
        "  行号含义：源文件里的物理行号（Excel 里看到的那一行）",
        "",
        "二、文件指纹（原始文件只读，可自行复核）",
        "  左表：{} 字节  SHA256 {}".format(left.bytes_size, left.sha256),
        "  右表：{} 字节  SHA256 {}".format(right.bytes_size, right.sha256),
        "",
        "三、行数统计",
        "  左表数据行数：{}".format(counts["left_rows"]),
        "  右表数据行数：{}".format(counts["right_rows"]),
        "  完全匹配数量：{}".format(counts["matched"]),
        "  有差异数量：{}".format(counts["differences"]),
        "  仅左表存在数量：{}".format(counts["only_left"]),
        "  仅右表存在数量：{}".format(counts["only_right"]),
        "",
        "  一致性校验：",
        "    完全匹配 + 有差异 + 仅左表 = {} + {} + {} = {}".format(
            counts["matched"], counts["differences"], counts["only_left"],
            counts["matched"] + counts["differences"] + counts["only_left"],
        ),
        "    上面这个数必须等于左表数据行数 {}".format(counts["left_rows"]),
        "    完全匹配 + 有差异 + 仅右表 = {} + {} + {} = {}".format(
            counts["matched"], counts["differences"], counts["only_right"],
            counts["matched"] + counts["differences"] + counts["only_right"],
        ),
        "    上面这个数必须等于右表数据行数 {}".format(counts["right_rows"]),
        "    （本工具只支持一对一：匹配键为空或重复都会直接失败，不会有被排除掉的行）",
        "",
        "四、公式与安全策略",
    ]
    lines.extend("  " + note for note in formula_notes)
    lines += [
        "  原始文件保护：写入路径（--output / --report / work/staging）都已确认不是 --left / --right",
        "  输出覆盖保护：已存在时默认拒绝，--force 才重跑，且旧结果先归档到 work/previous/（不删除）",
        "",
        "五、警告",
    ]
    if warnings:
        lines.extend("  - {}".format(warning) for warning in warnings)
    else:
        lines.append("  （无）")
    sheets = "、".join((SHEET_SUMMARY,) + DATA_SHEETS)
    lines += [
        "",
        "六、输出（两个文件同属本次成功运行）",
        "  {}  {}".format(OUTPUT_XLSX_NAME, output_xlsx),
        "  {}  {}".format(REPORT_NAME, output_report),
        "  工作表：{}".format(sheets),
        "",
        "说明",
        "  - 原始文件全程只读，未被修改（上面的 SHA256 可自行复核）。",
        "  - reconciliation.xlsx 与 report.txt 都先在 work/staging/ 完整生成并通过校验，",
        "    然后才一起提交到 output/；提交中任何一步失败都会回滚，output/ 不会出现半套结果。",
        "  - 判断「上次跑成功了吗」就看 output/ 里有没有 {}。".format(OUTPUT_XLSX_NAME),
    ]
    return "\n".join(lines) + "\n"


def build_failure_report(stamp, left_arg, right_arg, stage, reason):
    return "\n".join([
        REPORT_TITLE + "（本次运行失败）",
        "=" * 68,
        "版本：{}".format(VERSION_NAME),
        "运行时间：{}".format(stamp),
        "左表：{}".format(left_arg),
        "右表：{}".format(right_arg),
        "失败发生在：{}".format(stage),
        "",
        "拒绝 / 失败原因",
        reason,
        "",
        "说明",
        "  - 本次运行没有产出 output/ 下的任何文件（要么写盘前就被拒绝，要么提交已回滚）。",
        "  - 原始文件全程只读，未被修改。",
    ]) + "\n"


# --------------------------------------------------------------------------
# 提交前的硬校验
# --------------------------------------------------------------------------

def read_sheet_rows(path, sheet_name):
    """重读我们自己写出来的 xlsx 的某张表，返回二维列表（含表头行）。"""
    openpyxl = load_openpyxl()
    try:
        handle = path.open("rb")
    except OSError as exc:
        raise ReconcileError("校验时读不了结果文件：{}\n原始报错：{}".format(path, exc))
    with handle:
        try:
            workbook = openpyxl.load_workbook(handle, data_only=True)
        except Exception as exc:
            raise ReconcileError("校验时打不开结果文件：{}\n原始报错：{}".format(path, exc))
        try:
            if sheet_name not in workbook.sheetnames:
                raise ReconcileError(
                    "结果工作簿里缺少工作表「{}」（实际有：{}）".format(
                        sheet_name, "、".join(workbook.sheetnames))
                )
            rows = []
            for cells in workbook[sheet_name].iter_rows(values_only=True):
                rows.append(["" if value is None else value for value in cells])
            return rows
        finally:
            workbook.close()


def validate_staged(staged_xlsx, staged_report, report_text, counts, compare_display_names):
    """
    提交前的硬校验闸门。任何一条不过就不许靠近 output/。

    这里做的是"守恒 + 互斥"校验，它才是"没有静默丢行"的真正证据：
      完全匹配 + 有差异 + 仅左表 + 被排除(左) == 左表数据行数
      完全匹配 + 有差异 + 仅右表 + 被排除(右) == 右表数据行数
      四张表的键集合两两不相交
    """
    if not staged_xlsx.exists() or staged_xlsx.stat().st_size == 0:
        raise ReconcileError("暂存的对账结果不存在或是空文件：{}".format(staged_xlsx))
    if not staged_report.exists() or staged_report.stat().st_size == 0:
        raise ReconcileError("暂存的报告没有生成出来：{}".format(staged_report))

    for sheet_name in (SHEET_SUMMARY,) + DATA_SHEETS:
        if not _sheet_exists(staged_xlsx, sheet_name):
            raise ReconcileError(
                "结果工作簿里缺少工作表「{}」：{}".format(sheet_name, staged_xlsx)
            )

    actual = {}
    key_sets = {}
    for sheet_name in DATA_SHEETS:
        rows = read_sheet_rows(staged_xlsx, sheet_name)
        # 「匹配键」列永远非空，用它当"这是一条记录"的判据最稳
        body = [row for row in rows[1:] if row and str(row[0]) != ""]
        actual[sheet_name] = len(body)
        key_sets[sheet_name] = {str(row[0]) for row in body}

    for sheet_name, label in (
        (SHEET_MATCHED, "完全匹配数量"),
        (SHEET_DIFFERENCES, "有差异数量"),
        (SHEET_ONLY_LEFT, "仅左表存在数量"),
        (SHEET_ONLY_RIGHT, "仅右表存在数量"),
    ):
        if actual[sheet_name] != counts[sheet_name]:
            raise ReconcileError(
                "结果文件里的行数和统计对不上，已中止提交：\n"
                "    {}：工作表里实际 {} 行，统计是 {} 行".format(
                    label, actual[sheet_name], counts[sheet_name])
            )

    left_total = (actual[SHEET_MATCHED] + actual[SHEET_DIFFERENCES]
                  + actual[SHEET_ONLY_LEFT])
    if left_total != counts["left_rows"]:
        raise ReconcileError(
            "左表行数守恒校验没过，已中止提交：\n"
            "    完全匹配 + 有差异 + 仅左表 = {} + {} + {} = {}\n"
            "    左表数据行数 = {}\n"
            "两边不相等说明有行被静默丢掉了，这样的结果不能交付。".format(
                actual[SHEET_MATCHED], actual[SHEET_DIFFERENCES], actual[SHEET_ONLY_LEFT],
                left_total, counts["left_rows"],
            )
        )
    right_total = (actual[SHEET_MATCHED] + actual[SHEET_DIFFERENCES]
                   + actual[SHEET_ONLY_RIGHT])
    if right_total != counts["right_rows"]:
        raise ReconcileError(
            "右表行数守恒校验没过，已中止提交：\n"
            "    完全匹配 + 有差异 + 仅右表 = {} + {} + {} = {}\n"
            "    右表数据行数 = {}\n"
            "两边不相等说明有行被静默丢掉了，这样的结果不能交付。".format(
                actual[SHEET_MATCHED], actual[SHEET_DIFFERENCES], actual[SHEET_ONLY_RIGHT],
                right_total, counts["right_rows"],
            )
        )

    for first, second in (
        (SHEET_MATCHED, SHEET_DIFFERENCES),
        (SHEET_MATCHED, SHEET_ONLY_LEFT),
        (SHEET_MATCHED, SHEET_ONLY_RIGHT),
        (SHEET_DIFFERENCES, SHEET_ONLY_LEFT),
        (SHEET_DIFFERENCES, SHEET_ONLY_RIGHT),
        (SHEET_ONLY_LEFT, SHEET_ONLY_RIGHT),
    ):
        overlap = key_sets[first] & key_sets[second]
        if overlap:
            raise ReconcileError(
                "结果文件里同一个键同时出现在「{}」和「{}」里，已中止提交：\n    {}".format(
                    first, second, "、".join(sorted(overlap)[:5])
                )
            )

    diff_rows = read_sheet_rows(staged_xlsx, SHEET_DIFFERENCES)
    if diff_rows:
        header = [str(cell) for cell in diff_rows[0]]
        if COL_DIFF_FIELDS not in header:
            raise ReconcileError("differences 表缺少「{}」列。".format(COL_DIFF_FIELDS))
        position = header.index(COL_DIFF_FIELDS)
        allowed = set(compare_display_names)
        for row in diff_rows[1:]:
            if not row or str(row[0]) == "":
                continue
            for field in str(row[position]).split("、"):
                if field and field not in allowed:
                    raise ReconcileError(
                        "differences 表里出现了没让比较的字段「{}」，已中止提交。".format(field)
                    )

    text = staged_report.read_text(encoding="utf-8")
    if text != report_text:
        raise ReconcileError(
            "暂存的报告内容和刚才校验的那份不一致，已中止（写盘后被改动过？）：\n    {}".format(
                staged_report)
        )
    if REPORT_TITLE not in text:
        raise ReconcileError("暂存的报告内容不完整（缺少标题）：{}".format(staged_report))
    for label, value in (
        ("左表数据行数", counts["left_rows"]),
        ("右表数据行数", counts["right_rows"]),
        ("完全匹配数量", counts["matched"]),
        ("有差异数量", counts["differences"]),
        ("仅左表存在数量", counts["only_left"]),
        ("仅右表存在数量", counts["only_right"]),
    ):
        match = re.search(r"{}[^\d]*(\d+)".format(label), text)
        if not match or int(match.group(1)) != value:
            raise ReconcileError(
                "报告里的「{}」和实际结果对不上，已中止提交：\n"
                "    报告写的：{}\n"
                "    实际是　：{}".format(
                    label, match.group(1) if match else "（没找到）", value)
            )


def _sheet_exists(path, sheet_name):
    openpyxl = load_openpyxl()
    try:
        handle = path.open("rb")
    except OSError:
        return False
    with handle:
        try:
            workbook = openpyxl.load_workbook(handle, read_only=True)
        except Exception:
            return False
        try:
            return sheet_name in workbook.sheetnames
        finally:
            workbook.close()


# --------------------------------------------------------------------------
# 原子提交 + 回滚
# --------------------------------------------------------------------------

def install_file(src, dst):
    """
    把一个暂存文件装到目标位置，且"出现"这个动作是原子的。

    先在目标目录内复制成临时文件，再在同一个目录里改名过去 —— 同目录改名是原子的，
    所以外面永远看不到写了一半的 reconciliation.xlsx。
    失败时清掉自己造的临时文件，不在 output/ 里留下 .commit-* 垃圾。
    """
    tmp = dst.parent / (".commit-{}".format(dst.name))
    if tmp.exists():
        tmp.unlink()
    try:
        shutil.copy2(str(src), str(tmp))
        os.replace(str(tmp), str(dst))
    except OSError:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
        raise


def commit_outputs(staged_xlsx, staged_report, output_xlsx, output_report):
    """
    把暂存好的两个文件一起提交。

    顺序：先 report.txt，最后 reconciliation.xlsx ——
    于是"output/ 里出现 reconciliation.xlsx"成为整条流水线的最后一个动作。
    任何一步失败 → 回滚：删掉本次已经装进去的文件，output/ 不留半套成功结果。
    回滚只删自己刚装的那两个路径，绝不碰 output/ 里别的东西。
    """
    plan = [
        (staged_report, output_report),
        (staged_xlsx, output_xlsx),
    ]
    committed = []
    try:
        for src, dst in plan:
            install_file(src, dst)
            committed.append(dst)
    except OSError as exc:
        leftovers = []
        for path in committed:
            try:
                path.unlink()
            except OSError:
                leftovers.append(str(path))
        message = (
            "提交结果到 output/ 的时候失败了，已经回滚本次提交的文件。\n"
            "原始报错：{}\n"
            "output/ 里的本次产物已清理，不会出现半套结果。".format(exc)
        )
        if leftovers:
            message += (
                "\n注意：下面这些文件回滚删除失败，请手动确认：\n"
                + "\n".join("    " + item for item in leftovers)
            )
        raise ReconcileError(message)
    return committed


# --------------------------------------------------------------------------
# work/ 与 output/ 的准备
# --------------------------------------------------------------------------

def unique_path(directory, filename):
    """在目录里找一个不冲突的文件名，绝不覆盖已有文件（历史留存靠它）。"""
    candidate = directory / filename
    if not candidate.exists():
        return candidate
    stem, suffix = candidate.stem, candidate.suffix
    number = 2
    while True:
        candidate = directory / "{}-{}{}".format(stem, number, suffix)
        if not candidate.exists():
            return candidate
        number += 1


def work_subdir(work_dir, name):
    path = work_dir / name
    path.mkdir(parents=True, exist_ok=True)
    return path


def clear_work_dirs(work_dir):
    """
    清空 work/ 下我们自己拥有的子目录，保证中间产物一定是本次运行产生的。

    绝不能碰 work/previous/ —— 那里放的是 --force 归档的旧结果。有专门用例守着。
    """
    for name in CLEARABLE_WORK_SUBDIRS:
        path = work_dir / name
        if path.exists():
            shutil.rmtree(path)
        path.mkdir(parents=True, exist_ok=True)


def existing_outputs(output_xlsx, output_report):
    return [path for path in (output_xlsx, output_report) if path.exists()]


def move_previous_outputs(output_xlsx, output_report, work_dir, stamp):
    """--force 时把旧结果移到 work/previous/（移动、不删除，且绝不覆盖已有历史）。"""
    previous = work_subdir(work_dir, WORK_SUBDIR_PREVIOUS)
    moved = []
    for path in existing_outputs(output_xlsx, output_report):
        target = unique_path(previous, "{}_{}".format(stamp, path.name))
        shutil.move(str(path), str(target))
        moved.append((path, target))
    return moved


# --------------------------------------------------------------------------
# 路径闸门
# --------------------------------------------------------------------------

def check_paths(left_path, right_path, output_xlsx, output_report):
    for role, path in (("左表", left_path), ("右表", right_path)):
        if not path.exists():
            raise ReconcileError("找不到{}文件：{}".format(role, path))
        if not path.is_file():
            raise ReconcileError("{}不是一个普通文件：{}".format(role, path))
    if left_path.resolve() == right_path.resolve():
        raise ReconcileError(
            "左表和右表指向的是同一个文件：\n    {}\n"
            "自己跟自己比只会得到一份「全部匹配」的假结果，所以直接拒绝。\n"
            "请确认 --left / --right 给的是两份不同的表。".format(left_path.resolve())
        )
    if output_xlsx.suffix.lower() != ".xlsx":
        raise ReconcileError(
            "结果文件必须是 .xlsx，现在给的是「{}」。".format(
                output_xlsx.suffix or "（没有后缀）")
        )
    if output_xlsx.is_dir():
        raise ReconcileError("--output 指到了一个目录：{}，请给一个文件名。".format(output_xlsx))
    if output_xlsx.resolve() == output_report.resolve():
        raise ReconcileError("--output 和 --report 是同一个路径，换一个。")


def check_write_targets(left_path, right_path, output_xlsx, output_report, work_dir):
    """
    硬闸门：要写入的路径绝不能是 --left / --right 指向的那两个原始文件。

    与 --force 无关 —— --force 只放宽"覆盖自己以前的输出"，永不放宽"覆盖原始数据"。
    按 resolve() 后的真实路径比对，符号链接绕不过去。
    这个检查在所有 mkdir / 写文件之前执行。
    """
    protected = {"左表": left_path.resolve(), "右表": right_path.resolve()}
    targets = [
        ("输出文件（--output）", output_xlsx),
        ("报告文件（--report）", output_report),
        ("暂存目录（--work/staging）", work_dir / WORK_SUBDIR_STAGING),
    ]
    for label, path in targets:
        resolved = path.resolve()
        for role, protected_path in protected.items():
            if resolved == protected_path:
                raise ReconcileError(
                    "{}{} 指向的是{}原始文件：\n"
                    "    要写入的路径：{}\n"
                    "    解析后的真实文件：{}\n"
                    "本工具承诺永不覆盖任何原始文件，所以任务已中止，没有改写任何文件。\n"
                    "注意：这个拒绝和 --force 无关。\n"
                    "怎么改：换一个写入路径 —— 输出本来就不该落在输入文件上。".format(
                        label, "（是符号链接）" if path.is_symlink() else "", role,
                        path, resolved,
                    )
                )


# --------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------

def collect_warnings(left, right, left_index, right_index, result,
                     compare, left_key_col, right_key_col, args):
    """把"为什么可能对不上/需要注意"的诊断性提示收集起来。这些都不改变匹配结果。"""
    warnings = []

    only_left_count = len(result["only_left"])
    only_right_count = len(result["only_right"])
    if only_left_count and only_right_count:
        if (left_index.numeric_keys and right_index.digit_text_keys) or \
           (left_index.digit_text_keys and right_index.numeric_keys):
            warnings.append(
                "两侧都有对不上的键，而且左表有 {} 个键来自数值单元格、右表有 {} 个是纯数字文本"
                "（或相反）——\n"
                "    数值 123 和文本 \"00123\" 渲染出来不会是同一个键。这通常说明两边是从不同系统"
                "导出的、键的类型不一致。\n"
                "    请确认：如果键本来是编号，两边都应该按文本导出（保留前导零）。".format(
                    left_index.numeric_keys + right_index.numeric_keys,
                    left_index.digit_text_keys + right_index.digit_text_keys,
                )
            )

    big_numbers = []
    for role, table, column in (("左", left, left_key_col), ("右", right, right_key_col)):
        for row_number, values in table.data:
            raw = values[column]
            if isinstance(raw, float) and abs(raw) >= 1e15:
                big_numbers.append((role, row_number, render_cell_text(raw)))
            elif isinstance(raw, str) and re.search(r"\d[eE]\+?\d", raw):
                big_numbers.append((role, row_number, raw))
    if big_numbers:
        warnings.append(
            "有 {} 个匹配键是超大数字或含科学计数法（例如 {}）——\n"
            "    Excel 的数字只能可靠保存 15 位有效数字，超过就会丢精度而且无法恢复。\n"
            "    请从源头把这类键导出为文本。".format(
                len(big_numbers),
                "；".join("{}表第 {} 行 {!r}".format(role, row, text)
                          for role, row, text in big_numbers[:3]),
            )
        )

    invisible = []
    for role, index in (("左", left_index), ("右", right_index)):
        for row_number, chars in index.invisible_hits:
            invisible.append(
                "{}表第 {} 行（{}）".format(
                    role, row_number, "、".join(describe_char(ch) for ch in chars))
            )
    if invisible:
        warnings.append(
            "有 {} 个匹配键里含看不见的字符 —— {}\n"
            "    这些字符肉眼看不出来，但会让键对不上。请从源头清理后再对账。\n"
            "    注意：本工具只去除首尾空白，不做任何其它清理。".format(
                len(invisible), "；".join(invisible[:5])
            )
        )

    trimmed = [
        "{}表第 {} 行 {!r} → {!r}".format(role, row, raw, normalized)
        for role, index in (("左", left_index), ("右", right_index))
        for row, raw, normalized in index.trimmed_rows[:3]
    ]
    trimmed_total = len(left_index.trimmed_rows) + len(right_index.trimmed_rows)
    if trimmed_total:
        warnings.append(
            "有 {} 个匹配键是靠「去掉首尾空白」才对上的（这是 H1 唯一允许的规范化）：{}".format(
                trimmed_total, "；".join(trimmed)
            )
        )

    # 匹配键的值和表头名一模一样 → 很可能这个表里上下堆了两块数据：
    # 第二块的表头被当成了一条普通记录。这条不猜、不改匹配，只提醒。
    for role, table, column in (
        ("左", left, left_key_col), ("右", right, right_key_col)
    ):
        header_name = trim_text(table.header[column])
        if not header_name:
            continue
        hits = [row for row, values in table.data if values[column] == header_name]
        if hits:
            warnings.append(
                "{}表有 {} 行的匹配键和表头名 {!r} 一模一样（第 {} 行）——\n"
                "    这通常意味着这张表里**上下堆了两块数据**，第二块的表头被当成了一条普通记录。\n"
                "    本工具不猜表结构，请把数据拆成两个工作表、或用 --{}header-row 指定真正的表头行。".format(
                    role, len(hits), header_name,
                    "、".join(str(row) for row in hits[:5]), table.flag_prefix,
                )
            )

    # 比较字段两边类型不一致（一侧全是数值、另一侧全是文本）：
    # 严格文本比较下这类差异几乎全是"写法不同"造成的，必须提醒，
    # 但**不会**因此放宽比较口径 —— 结果里它们仍然是差异。
    for left_col, right_col, _left_name, _right_name, display in compare:
        if not result["field_diff_counts"].get(display):
            continue
        left_numeric = left.numeric_by_col.get(left_col, 0)
        left_text = left.text_by_col.get(left_col, 0)
        right_numeric = right.numeric_by_col.get(right_col, 0)
        right_text = right.text_by_col.get(right_col, 0)
        one_side_numbers = (left_numeric and not left_text) or (right_numeric and not right_text)
        other_side_texts = (left_text and not left_numeric) or (right_text and not right_numeric)
        if one_side_numbers and other_side_texts:
            warnings.append(
                "比较字段「{}」两边类型不一致：左表 {} 个数值单元格 / {} 个文本单元格，"
                "右表 {} 个数值 / {} 个文本。\n"
                "    严格文本比较下，这类值很可能全部被报成差异 —— 例如数值 1200 渲染成 \"1200\"，"
                "而文本 \"1200.00\" 保持原样，两者不相等。\n"
                "    本工具不会自动做数值等价判断（那是 H2 的事），"
                "请把两侧导出成同一种格式，或先在源侧统一小数位。".format(
                    display, left_numeric, left_text, right_numeric, right_text
                )
            )

    for role, table, sheet_arg in (("左", left, args.left_sheet), ("右", right, args.right_sheet)):
        if table.sheet_names and sheet_arg is None and len(table.sheet_names) > 1:
            warnings.append(
                "{}表没有指定 --{}sheet，按默认用了**第一个**工作表「{}」。"
                "该文件一共 {} 张工作表：{}\n"
                "    如果内容不对，请用 --{}sheet 显式指定。".format(
                    role, table.flag_prefix, table.sheet_title, len(table.sheet_names),
                    "、".join("「{}」".format(name) for name in table.sheet_names),
                    table.flag_prefix,
                )
            )
        hidden_in_data = [row for row, _values in table.data if row in table.hidden_rows]
        if hidden_in_data:
            warnings.append(
                "{}表有 {} 行是**隐藏行**，它们仍然参与了对账（不跳过 —— 跳过就是静默丢数据）：{}".format(
                    role, len(hidden_in_data),
                    "、".join(str(row) for row in hidden_in_data[:5]),
                )
            )
        if table.has_autofilter:
            warnings.append(
                "{}表当前处于筛选状态：读到的仍然是**全部行**，不是筛选后可见的行。".format(role)
            )
        if table.header_row > 1:
            warnings.append(
                "{}表按 --{}header-row {} 读取，表头上面那 {} 行已作为标题忽略。".format(
                    role, table.flag_prefix, table.header_row, table.header_row - 1)
            )
        if table.blank_rows_skipped:
            warnings.append(
                "{}表有 {} 个**整行空白**的行被忽略（一个值都没有的行不算记录）。".format(
                    role, table.blank_rows_skipped)
            )
        if table.ragged_rows:
            warnings.append(
                "{}表有 {} 行的列数和表头不一致（可能是分隔符或引号问题），"
                "多出来的列没有参与对账。".format(role, table.ragged_rows)
            )

    for left_col, right_col, left_name, right_name, _display in compare:
        if left_col == left_key_col:
            warnings.append(
                "左表的比较字段「{}」和匹配键是同一列 —— 匹配成功的行里它必然一致，"
                "放在比较字段里没有意义。".format(left_name)
            )
        if right_col == right_key_col:
            warnings.append(
                "右表的比较字段「{}」和匹配键是同一列 —— 匹配成功的行里它必然一致，"
                "放在比较字段里没有意义。".format(right_name)
            )
    return warnings


def execute(args):
    left_path = Path(args.left).expanduser().resolve()
    right_path = Path(args.right).expanduser().resolve()
    work_dir = Path(args.work).expanduser().resolve()
    # 写入目标先**不要** resolve：闸门要用用户给的原路径来判断"这是不是一个符号链接"，
    # 才好把话说明白（"你给的是个链接，它实际指向左表"）。闸门过了之后再统一 resolve。
    output_xlsx_raw = Path(args.output).expanduser()
    output_report_raw = Path(args.report).expanduser()
    stamp_readable = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    stamp_file = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")

    # ---- 路径闸门（都在写盘之前）----
    # 顺序有意为之：先查「写入目标是不是原始文件」，再查格式。
    # 如果有人把 -o 指到了自己的源文件上（哪怕是绕道符号链接），最该听到的是
    # "这会覆盖原始数据"，而不是"后缀不对"（后者会让人以为换个后缀就行了）。
    check_write_targets(left_path, right_path, output_xlsx_raw, output_report_raw, work_dir)
    output_xlsx = output_xlsx_raw.resolve()
    output_report = output_report_raw.resolve()
    check_paths(left_path, right_path, output_xlsx, output_report)

    previous = existing_outputs(output_xlsx, output_report)
    if previous and not args.force:
        raise ReconcileError(
            "输出位置已经有上一次的对账结果了：\n{}\n"
            "本工具默认不覆盖已经存在的最终结果。\n"
            "确实要重新跑就加 --force —— 加了之后这些文件会被移到\n"
            "    {}\n"
            "里保留下来（不删除），而不是就地覆盖。".format(
                "\n".join("    " + str(path) for path in previous),
                work_dir / WORK_SUBDIR_PREVIOUS,
            )
        )

    work_dir.mkdir(parents=True, exist_ok=True)
    clear_work_dirs(work_dir)
    staging_dir = work_subdir(work_dir, WORK_SUBDIR_STAGING)
    moved_previous = move_previous_outputs(
        output_xlsx, output_report, work_dir, stamp_file) if previous else []

    warnings = []
    if output_xlsx.parent == left_path.parent or output_xlsx.parent == right_path.parent:
        warnings.append(
            "结果文件所在目录就是输入文件所在目录 —— 已确认输出路径不是那两个输入文件本身，"
            "但建议把输出放到单独的目录，避免下次误把结果当成输入。"
        )

    # ---- 读两表 ----
    left = (read_csv_table(left_path, args.left_delimiter, args.left_header_row, "左")
            if detect_kind(left_path, "左") == "csv"
            else read_excel_table(left_path, args.left_sheet, args.left_header_row, "左"))
    right = (read_csv_table(right_path, args.right_delimiter, args.right_header_row, "右")
             if detect_kind(right_path, "右") == "csv"
             else read_excel_table(right_path, args.right_sheet, args.right_header_row, "右"))
    for table in (left, right):
        table.bytes_size = table.path.stat().st_size
        table.sha256 = sha256_of(table.path)
        if table.replacement_rows:
            raise ReconcileError(
                "{}表解码后出现替换字符（U+FFFD），说明编码猜错了：{}\n"
                "受影响的行：{}\n"
                "乱码的键不可能匹配上。请把源文件另存为 UTF-8 或 GB18030 再试。".format(
                    table.role, table.path,
                    "、".join(str(row) for row in table.replacement_rows[:10]),
                )
            )
        if not table.data:
            extra = ""
            if table.sheet_names:
                extra = "（这个文件的工作表有：{}）\n".format(
                    "、".join("「{}」".format(name) for name in table.sheet_names))
            raise ReconcileError(
                "{}表在表头行（第 {} 行）下面一行数据都没有：{}\n{}"
                "对账两边都必须有数据，本工具不会拿空表算出一份「全部只在另一侧」的结果。\n"
                "请检查 --left/--right 是否指对了文件、--{}sheet 是否选对了工作表、"
                "--{}header-row 是否指对了表头行。".format(
                    table.role, table.header_row, table.path, extra,
                    table.flag_prefix, table.flag_prefix,
                )
            )

    # ---- 列定位 ----
    left_index, left_dups = build_header_index(left.header)
    right_index, right_dups = build_header_index(right.header)
    left_key_col = resolve_column(left, left_index, args.left_key, "匹配键")
    right_key_col = resolve_column(right, right_index, args.right_key, "匹配键")
    compare = []
    for left_name, right_name, display in parse_compare_specs(args.compare):
        left_col = resolve_column(left, left_index, left_name, "比较字段")
        right_col = resolve_column(right, right_index, right_name, "比较字段")
        compare.append((left_col, right_col, left_name, right_name, display))
    if left_dups:
        warnings.append("左表表头里有重名的列（没有参与本次对账，只提示）：{}".format(left_dups))
    if right_dups:
        warnings.append("右表表头里有重名的列（没有参与本次对账，只提示）：{}".format(right_dups))

    # ---- 公式策略 ----
    formula_notes = [
        "读 Excel 的方式：openpyxl data_only=False —— 只拿得到公式串，",
        "结构上不可能用到可能过期的公式缓存值。",
        "我们自己写出的 {} 只写文本和数字，永不写公式。".format(OUTPUT_XLSX_NAME),
    ]
    participating = {
        "左": {left_key_col} | {left_col for left_col, _rc, _ln, _rn, _d in compare},
        "右": {right_key_col} | {right_col for _lc, right_col, _ln, _rn, _d in compare},
    }
    for table in (left, right):
        used = participating[table.role]
        in_use, ignored = [], []
        for column, hits in table.formulas_by_col.items():
            (in_use if column in used else ignored).extend(hits)
        if in_use:
            shown = "\n".join(
                "    [{}] 第 {} 行 {}：{}".format(table.sheet_title or "", row, coord, text)
                for coord, row, text in in_use[:5]
            )
            more = "" if len(in_use) <= 5 else "\n    …… 另外还有 {} 处".format(len(in_use) - 5)
            raise ReconcileError(
                "{0}表的参与列里有公式单元格，无法安全对账，任务已中止，没有生成任何输出文件。\n"
                "文件：{1}\n工作表：{2}\n{3} 处公式：\n{4}{5}\n\n"
                "为什么不能继续：本工具读 Excel 用 data_only=False，拿到的只有公式串"
                "（例如 =A1&B1），拿它当匹配键或比较值必然是错的；\n"
                "改用公式缓存值又可能拿到过期结果 —— 两条路都不能走，所以只能停下。\n"
                "怎么改：在 Excel 里把这些列的公式选择性粘贴成数值，另存为新的 .xlsx 再对账。".format(
                    table.role, table.path, table.sheet_title or "（CSV）",
                    len(in_use), shown, more,
                )
            )
        if ignored:
            formula_notes.append(
                "{}表：不参与对账的列里有 {} 处公式，已**明确忽略**（不读、不影响结果）：{}".format(
                    table.role, len(ignored),
                    "；".join("第 {} 行 {}".format(row, coord) for coord, row, _t in ignored[:3])
                    + ("" if len(ignored) <= 3 else " 等"),
                )
            )
        else:
            formula_notes.append("{}表：公式单元格 0 处。".format(table.role))
        if table.equal_text_cells:
            formula_notes.append(
                "{}表：有 {} 个以 = 开头的**文本**单元格（不是公式），已按文本正常读取，只做提示：{}".format(
                    table.role, len(table.equal_text_cells),
                    "；".join("第 {} 行 {}".format(row, coord)
                              for coord, row, _t in table.equal_text_cells[:3]),
                )
            )

    # ---- 建键索引：空键 / 重复键在这里无条件拦下（没有任何参数可以绕过）----
    left_keys = build_key_index(left, left_key_col)
    right_keys = build_key_index(right, right_key_col)
    problems = describe_key_problems(left_keys, right_keys)
    if problems:
        raise ReconcileError(problems)

    # ---- 对账 ----
    result = reconcile(left_keys, right_keys, compare)
    warnings.extend(collect_warnings(
        left, right, left_keys, right_keys, result, compare,
        left_key_col, right_key_col, args,
    ))

    counts = get_counts(left, right, result)
    config_rows = [
        ("版本", VERSION_NAME),
        ("运行时间", stamp_readable),
        ("左表数据行数", counts["left_rows"]),
        ("右表数据行数", counts["right_rows"]),
        ("完全匹配数量", counts["matched"]),
        ("有差异数量", counts["differences"]),
        ("仅左表存在数量", counts["only_left"]),
        ("仅右表存在数量", counts["only_right"]),
        ("比较字段数", len(compare)),
        ("左表文件", str(left_path)),
        ("右表文件", str(right_path)),
        ("左表 SHA256", left.sha256),
        ("右表 SHA256", right.sha256),
        ("左表字节数", left.bytes_size),
        ("右表字节数", right.bytes_size),
        ("匹配键（左）", args.left_key),
        ("匹配键（右）", args.right_key),
        ("比较字段", "、".join(
            "左「{}」= 右「{}」".format(left_name, right_name)
            for _lc, _rc, left_name, right_name, _d in compare)),
        ("左表工作表", left.sheet_title or "（CSV）"),
        ("右表工作表", right.sheet_title or "（CSV）"),
        ("左表表头行", left.header_row),
        ("右表表头行", right.header_row),
        ("左表编码", left.encoding or "（Excel）"),
        ("右表编码", right.encoding or "（Excel）"),
        ("左表分隔符", repr(left.delimiter) if left.delimiter else "（Excel）"),
        ("右表分隔符", repr(right.delimiter) if right.delimiter else "（Excel）"),
        ("匹配/比较口径", "严格：只去除首尾空白（不做大小写/全半角/千分位/容差/日期识别）"),
        ("行号含义", "源文件里的物理行号"),
        ("输出文件", str(output_xlsx)),
        ("报告文件", str(output_report)),
    ]

    staged_xlsx = staging_dir / OUTPUT_XLSX_NAME
    staged_report = staging_dir / REPORT_NAME
    workbook = build_workbook(
        result, left, right, left_key_col, right_key_col, compare, config_rows,
    )
    write_workbook(staged_xlsx, workbook)

    report_text = build_report(
        stamp_readable, left, right, left_path, right_path, args.left_key, args.right_key,
        compare, result, warnings, formula_notes, output_xlsx, output_report,
    )
    staged_report.write_text(report_text, encoding="utf-8")

    # ---- 校验闸门 + 原子提交 ----
    validate_staged(staged_xlsx, staged_report, report_text, counts,
                    [display for _lc, _rc, _ln, _rn, display in compare])
    # 输出目录的父目录放在这里才建 —— 一定要在校验通过之后，
    # 这样"被拒绝时磁盘上不多出任何东西（连父目录都不会被创建）"才是真的。
    for directory in {output_xlsx.parent, output_report.parent}:
        directory.mkdir(parents=True, exist_ok=True)
    commit_outputs(staged_xlsx, staged_report, output_xlsx, output_report)

    if moved_previous:
        report_text += "\n本次用 --force 归档的旧结果（保留，未删除）：\n" + "\n".join(
            "  {} → {}".format(old, target) for old, target in moved_previous
        ) + "\n"
    return report_text


# --------------------------------------------------------------------------
# 命令行
# --------------------------------------------------------------------------

def build_parser():
    parser = argparse.ArgumentParser(
        description=(
            "{0}：Excel / CSV 两表匹配与对账差异检测。"
            "匹配键和比较字段必须显式指定，只支持一对一匹配，原始文件只读。".format(VERSION_NAME)
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python3 reconcile.py \\\n"
            "      --left  客户对账单.csv --left-header-row 3 --left-key 订单号 \\\n"
            "      --right 内部台账.xlsx --right-sheet 内部台账 --right-key 单号 \\\n"
            "      --compare 应收金额=金额,状态\n"
        ),
    )
    parser.add_argument("--left", required=True, help="左表文件（.csv / .txt / .xlsx）")
    parser.add_argument("--right", required=True, help="右表文件（.csv / .txt / .xlsx）")
    parser.add_argument("--left-key", required=True, help="左表的匹配键列名（单列）")
    parser.add_argument("--right-key", required=True, help="右表的匹配键列名（单列）")
    parser.add_argument(
        "--compare", required=True, action="append",
        help="要比较的字段，逗号分隔；写法「列名」或「左列名=右列名」。可重复给多次。",
    )
    parser.add_argument("--left-sheet", help="左表 .xlsx 的工作表名（默认第一个工作表）")
    parser.add_argument("--right-sheet", help="右表 .xlsx 的工作表名（默认第一个工作表）")
    parser.add_argument("--left-header-row", type=int, default=1, help="左表表头在第几行（默认 1）")
    parser.add_argument("--right-header-row", type=int, default=1, help="右表表头在第几行（默认 1）")
    parser.add_argument("--left-delimiter", default=",", help="左表 CSV 分隔符（默认逗号）")
    parser.add_argument("--right-delimiter", default=",", help="右表 CSV 分隔符（默认逗号）")
    parser.add_argument(
        "--output", default=str(PROJECT_ROOT / "output" / OUTPUT_XLSX_NAME),
        help="结果工作簿路径（默认：output/{}）".format(OUTPUT_XLSX_NAME),
    )
    parser.add_argument(
        "--report", default=str(PROJECT_ROOT / "output" / REPORT_NAME),
        help="文本报告路径（默认：output/{}）".format(REPORT_NAME),
    )
    parser.add_argument(
        "--work", default=str(PROJECT_ROOT / "work"),
        help="中间目录（默认：项目里的 work/）",
    )
    parser.add_argument(
        "--force", action="store_true",
        help="输出已存在时允许重跑；旧结果会先归档到 work/previous/（保留，不删除）",
    )
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)

    for name in ("left_header_row", "right_header_row"):
        if getattr(args, name) < 1:
            print(
                "错误：--{} 必须从 1 开始（第 1 行就是表头）。".format(name.replace("_", "-")),
                file=sys.stderr,
            )
            return 1

    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        report = execute(args)
    except ReconcileError as exc:
        print("错误：{}".format(exc), file=sys.stderr)
        try:
            work_dir = Path(args.work).expanduser().resolve()
            left_path = Path(args.left).expanduser().resolve()
            right_path = Path(args.right).expanduser().resolve()
            # 失败报告也不许影响原始文件：如果输入文件就在 work/ 里面，就什么都不写。
            # （否则为了留一份失败报告，会在放着原始文件的目录里凭空造出 work/reports/）
            safe = not (left_path.is_relative_to(work_dir)
                        or right_path.is_relative_to(work_dir))
            if safe:
                reports_dir = work_subdir(work_dir, WORK_SUBDIR_REPORTS)
                (reports_dir / FAILED_REPORT_NAME).write_text(
                    build_failure_report(stamp, args.left, args.right, "对账流程", str(exc)),
                    encoding="utf-8",
                )
        except Exception:
            pass
        return 1
    print(report)
    return 0


if __name__ == "__main__":
    sys.exit(main())

