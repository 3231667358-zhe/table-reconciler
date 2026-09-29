#!/usr/bin/env python3
"""
Table Reconciler H1 自检脚本

跑法（在项目目录里）：
    python3 -W error::ResourceWarning selftest.py

只用标准库 unittest，不需要 pytest。所有临时文件都建在系统临时目录里，跑完自己删掉 ——
不会碰项目里真实的 input/ work/ output/ 和 examples/。

这里测的是**契约**，不是实现细节。重点守这几条：

1. 文本保真：`00123` 绝不能被当成 `123`；不做大小写 / 全半角 / 千分位 / 容差 / 日期识别。
   默认口径只有"去除首尾空白"这一条规范化。
2. 一对一：任意一侧匹配键重复 → 硬失败；任意一侧匹配键为空 → 也**无条件**硬失败。
   空键没有任何参数可以绕过（没有 excluded 表、没有"排除后继续"的口子），
   报错必须说清是左表还是右表、空了几行、对应哪些原始行号。
3. 分类正确 + 不丢行：
   左行数 == matched + differences + only_left + 被排除(左)，
   右行数 == matched + differences + only_right + 被排除(右)，
   四张表的键集合两两不相交。
4. Excel 公式：参与列有公式 → 硬失败；非参与列有公式 → 忽略且明确报告。
5. 原始文件只读：所有成功和失败场景之后，左右文件的 sha256 都必须一模一样。
6. 输出保护：默认拒绝覆盖；--force 归档到 work/previous/（不删除、不覆盖历史）；
   提交失败必须回滚，output/ 不留半套结果。
"""

import contextlib
import csv
import hashlib
import io
import shutil
import subprocess
import sys
import tempfile
import unittest
import unittest.mock
from pathlib import Path

import reconcile

try:
    import openpyxl

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

PROJECT_ROOT = reconcile.PROJECT_ROOT
DEMO_LEFT = PROJECT_ROOT / "examples" / "客户对账单.csv"
DEMO_RIGHT = PROJECT_ROOT / "examples" / "内部台账.xlsx"


# --------------------------------------------------------------------------
# 小工具
# --------------------------------------------------------------------------

def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read_csv_rows(path):
    with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
        return [row for row in csv.reader(handle)]


def make_xlsx(path, sheets):
    """sheets = [(表名, [行, 行, ...]), ...]"""
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    for title, rows in sheets:
        sheet = workbook.create_sheet(title=title)
        for row in rows:
            sheet.append(row)
    workbook.save(path)
    workbook.close()


def read_xlsx_rows(path, sheet_name):
    with Path(path).open("rb") as handle:
        workbook = openpyxl.load_workbook(handle, data_only=False)
    try:
        sheet = workbook[sheet_name]
        return [
            ["" if cell.value is None else cell.value for cell in row]
            for row in sheet.iter_rows()
        ]
    finally:
        workbook.close()


def sheet_names(path):
    with Path(path).open("rb") as handle:
        workbook = openpyxl.load_workbook(handle)
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def make_formula_xlsx(path, sheet_title, rows, formula_cell="B2"):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = sheet_title
    for row in rows:
        sheet.append(row)
    sheet[formula_cell] = "=SUM(1,2)"
    workbook.save(path)
    workbook.close()


class TempCase(unittest.TestCase):
    """每个用例一套独立的临时文件与 work/output 目录。"""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="reconciler-selftest-"))
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.work = self.tmp / "work"
        self.output_xlsx = self.tmp / "output" / "reconciliation.xlsx"
        self.output_report = self.tmp / "output" / "report.txt"

    # ---- 造数据 ----

    def write_csv(self, name, text, encoding="utf-8"):
        path = self.tmp / name
        path.write_text(text, encoding=encoding)
        return path

    def left_csv(self, rows, header="订单号,客户,金额,状态", name="left.csv", encoding="utf-8"):
        lines = [header] + [",".join(str(cell) for cell in row) for row in rows]
        return self.write_csv(name, "\n".join(lines) + "\n", encoding=encoding)

    def right_xlsx(self, rows, header=("单号", "金额", "状态"), sheet="内部台账",
                   name="right.xlsx"):
        path = self.tmp / name
        make_xlsx(path, [(sheet, [list(header)] + [list(row) for row in rows])])
        return path

    def right_csv(self, rows, header="单号,金额,状态", name="right.csv"):
        lines = [header] + [",".join(str(cell) for cell in row) for row in rows]
        return self.write_csv(name, "\n".join(lines) + "\n")

    # ---- 跑 ----

    def run_main(self, *positional, left=None, right=None, left_key="订单号", right_key="单号",
                 compare=("金额",), output=None, report=None, work=None, extra=()):
        argv = [
            "--left", str(left if left is not None else self.left),
            "--right", str(right if right is not None else self.right),
            "--left-key", left_key,
            "--right-key", right_key,
            "--work", str(work if work is not None else self.work),
            "--output", str(output if output is not None else self.output_xlsx),
            "--report", str(report if report is not None else self.output_report),
        ]
        for item in compare:
            argv += ["--compare", item]
        argv += list(positional) + list(extra)
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer), contextlib.redirect_stderr(buffer):
            code = reconcile.main(argv)
        return code, buffer.getvalue()

    def run_cli(self, argv):
        completed = subprocess.run(
            [sys.executable, str(PROJECT_ROOT / "reconcile.py")] + list(argv),
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        return completed.returncode, (completed.stdout or "") + (completed.stderr or "")

    # ---- 断言 ----

    def assert_no_output(self):
        self.assertFalse(
            self.output_xlsx.exists(), "output/ 里出现了 reconciliation.xlsx，这是假成功")
        self.assertFalse(
            self.output_report.exists(), "output/ 里出现了 report.txt，这是假成功")

    def assert_output_ok(self):
        self.assertTrue(self.output_xlsx.exists())
        self.assertTrue(self.output_report.exists())

    def counts_from_xlsx(self):
        result = {}
        for sheet in reconcile.DATA_SHEETS:
            rows = read_xlsx_rows(self.output_xlsx, sheet)
            result[sheet] = len([row for row in rows[1:] if str(row[0]) != ""])
        return result

    def assert_conservation(self, left_rows, right_rows):
        actual = self.counts_from_xlsx()
        self.assertEqual(
            actual[reconcile.SHEET_MATCHED] + actual[reconcile.SHEET_DIFFERENCES]
            + actual[reconcile.SHEET_ONLY_LEFT],
            left_rows, "左表行数守恒不成立")
        self.assertEqual(
            actual[reconcile.SHEET_MATCHED] + actual[reconcile.SHEET_DIFFERENCES]
            + actual[reconcile.SHEET_ONLY_RIGHT],
            right_rows, "右表行数守恒不成立")
        return actual


# ==========================================================================
# T1 文本渲染（对账正确性的地基）
# ==========================================================================

class TestCellRendering(unittest.TestCase):
    def test_text_keeps_leading_zeros(self):
        self.assertEqual(reconcile.render_cell_text("00123"), "00123")
        self.assertEqual(reconcile.normalize_match_text("00123"), "00123")

    def test_int(self):
        self.assertEqual(reconcile.render_cell_text(123), "123")

    def test_integral_float_renders_as_integer(self):
        # 同一个数写成同一种形式：Excel 的 1200 与文本 "1200" 要对得上
        self.assertEqual(reconcile.render_cell_text(1200.0), "1200")
        self.assertEqual(reconcile.render_cell_text(-7.0), "-7")

    def test_non_integral_float(self):
        self.assertEqual(reconcile.render_cell_text(850.5), "850.5")

    def test_float_tail_is_kept_not_tolerated(self):
        # 不做数值容差：0.1+0.2 就如实渲染成它真实存的值
        self.assertEqual(reconcile.render_cell_text(0.1 + 0.2), "0.30000000000000004")

    def test_small_float_no_scientific_notation(self):
        text = reconcile.render_cell_text(1e-5)
        self.assertNotIn("e", text.lower())
        self.assertEqual(text, "0.00001")

    def test_large_integral_float_no_scientific_notation(self):
        self.assertEqual(reconcile.render_cell_text(1e16), "10000000000000000")

    def test_bool(self):
        self.assertEqual(reconcile.render_cell_text(True), "TRUE")
        self.assertEqual(reconcile.render_cell_text(False), "FALSE")

    def test_none(self):
        self.assertEqual(reconcile.render_cell_text(None), "")

    def test_datetime_midnight_becomes_date(self):
        import datetime
        self.assertEqual(
            reconcile.render_cell_text(datetime.datetime(2026, 8, 1, 0, 0)), "2026-08-01")

    def test_datetime_with_time(self):
        import datetime
        self.assertEqual(
            reconcile.render_cell_text(datetime.datetime(2026, 8, 1, 13, 5, 6)),
            "2026-08-01 13:05:06")

    def test_date_and_time(self):
        import datetime
        self.assertEqual(reconcile.render_cell_text(datetime.date(2026, 8, 1)), "2026-08-01")
        self.assertEqual(reconcile.render_cell_text(datetime.time(1, 2, 3)), "01:02:03")

    def test_nan_and_inf(self):
        self.assertEqual(reconcile.render_cell_text(float("nan")), "NaN")
        self.assertEqual(reconcile.render_cell_text(float("inf")), "inf")
        self.assertEqual(reconcile.render_cell_text(float("-inf")), "-inf")


class TestTextVersusNumberKeys(TempCase):
    """`00123`（文本）和 123（数值）永远不能被当成同一个键。"""

    def test_text_key_and_number_key_do_not_match(self):
        self.left = self.left_csv([["00123", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["123", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 0, out)
        actual = self.counts_from_xlsx()
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 0)
        self.assertEqual(actual[reconcile.SHEET_ONLY_LEFT], 1)
        self.assertEqual(actual[reconcile.SHEET_ONLY_RIGHT], 1)
        # 报告要提醒"两侧键的类型可能不一致"
        self.assertIn("数值", out)

    def test_int_key_matches_same_digit_text_key(self):
        # 123（数值）与 "123"（文本）渲染后都是 "123"，属于同一个键 —— 这不是假匹配，
        # 而是"同一个数写成同一种形式"
        self.left = self.left_csv([["123", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["123", "10", "已发"]])
        code, _ = self.run_main()
        self.assertEqual(code, 0)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)


# ==========================================================================
# T2 规范化边界（只有"去首尾空白"这一条）
# ==========================================================================

class TestNormalization(TempCase):
    def match_single(self, left_key, right_key):
        self.left = self.left_csv([[left_key, "甲", "10", "已发"]])
        self.right = self.right_xlsx([[right_key, "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 0, out)
        return self.counts_from_xlsx()

    def test_leading_trailing_space_trimmed(self):
        actual = self.match_single("  A-1  ", "A-1")
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 1)

    def test_nbsp_trimmed(self):
        actual = self.match_single(" A-1 ", "A-1")
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 1)

    def test_ideographic_space_trimmed(self):
        actual = self.match_single("　A-1　", "A-1")
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 1)

    def test_tab_trimmed(self):
        actual = self.match_single("\tA-1\t", "A-1")
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 1)

    def test_interior_nbsp_is_not_removed_and_warns(self):
        self.left = self.left_csv([["A 1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A 1", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 0, out)
        actual = self.counts_from_xlsx()
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 0)
        self.assertIn("看不见的字符", out)

    def test_zero_width_space_not_removed_and_warns(self):
        self.left = self.left_csv([["A​1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A1", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 0, out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 0)
        self.assertIn("U+200B", out)

    def test_case_is_not_folded(self):
        actual = self.match_single("abc", "ABC")
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 0)
        self.assertEqual(actual[reconcile.SHEET_ONLY_LEFT], 1)

    def test_fullwidth_is_not_converted(self):
        actual = self.match_single("ＡＢＣ", "ABC")
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 0)

    def test_thousand_separator_is_not_removed(self):
        actual = self.match_single("1,200", "1200")
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 0)

    def test_trim_only_matching_is_reported(self):
        self.left = self.left_csv([["  A-1  ", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 0)
        self.assertIn("去掉首尾空白", out)


class TestCompareValueStrictness(TempCase):
    def diff_for(self, left_amount, right_amount):
        self.left = self.left_csv([["A-1", "甲", left_amount, "已发"]])
        self.right = self.right_xlsx([["A-1", right_amount, "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 0, out)
        return self.counts_from_xlsx(), out

    def test_same_text_matches(self):
        actual, _ = self.diff_for("1200.00", "1200.00")
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 1)

    def test_text_amount_vs_number_cell_is_a_difference(self):
        # 左 "1200.00"（文本）vs 右 1200（数值）→ 渲染成 "1200.00" / "1200"，严格比较就是差异
        actual, out = self.diff_for("1200.00", 1200)
        self.assertEqual(actual[reconcile.SHEET_DIFFERENCES], 1)
        self.assertIn("类型不一致", out)

    def test_compare_value_surrounding_space_trimmed(self):
        actual, _ = self.diff_for("  1200.00  ", "1200.00")
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 1)

    def test_empty_versus_zero_is_a_difference(self):
        actual, _ = self.diff_for("", "0")
        self.assertEqual(actual[reconcile.SHEET_DIFFERENCES], 1)

    def test_date_cell_versus_iso_text_matches(self):
        import datetime
        path = self.tmp / "right.xlsx"
        make_xlsx(path, [("内部台账", [
            ["单号", "金额", "状态"],
            ["A-1", "10", datetime.datetime(2026, 8, 1, 0, 0)],
        ])])
        left = self.left_csv([["A-1", "甲", "10", "2026-08-01"]])
        code, out = self.run_main(left=left, right=path, compare=("状态",))
        self.assertEqual(code, 0, out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)


# ==========================================================================
# T3 唯一性与空键
# ==========================================================================

class TestKeyUniqueness(TempCase):
    def test_duplicate_key_on_left_hard_fails(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"], ["A-1", "乙", "20", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        before = (sha256(self.left), sha256(self.right))
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assert_no_output()
        self.assertIn("重复", out)
        self.assertIn("A-1", out)
        self.assertEqual((sha256(self.left), sha256(self.right)), before)

    def test_duplicate_key_on_right_hard_fails(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"], ["A-1", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assert_no_output()

    def test_duplicate_after_trimming_hard_fails(self):
        # "A-1" 与 "  A-1 " 去空白后是同一个键 → 仍然要判为重复，不能变成悄悄的一对多
        self.left = self.left_csv([["A-1", "甲", "10", "已发"], ["  A-1 ", "乙", "20", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assertIn("重复", out)

    def test_empty_key_on_left_hard_fails(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"], ["", "合计", "10", ""]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        before = (sha256(self.left), sha256(self.right))
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assert_no_output()
        self.assertIn("空", out)
        self.assertIn("合计", out)
        self.assertEqual((sha256(self.left), sha256(self.right)), before)

    def test_empty_key_on_left_reports_side_count_and_row_number(self):
        # 报错必须讲清三件事：哪一侧、空了几行、对应哪些原始行号
        self.left = self.left_csv([
            ["A-1", "甲", "10", "已发"],
            ["", "合计", "10", ""],
            ["A-2", "乙", "20", "已发"],
            ["", "小计", "20", ""],
        ])
        self.right = self.right_xlsx([["A-1", "10", "已发"], ["A-2", "20", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assert_no_output()
        self.assertIn("左表有 2 行的匹配键是空的", out)
        self.assertIn("第 3 行", out)   # 第一条空键的物理行号
        self.assertIn("第 5 行", out)   # 第二条空键的物理行号
        self.assertNotIn("右表有", out)  # 右表没有空键，不该提它

    def test_empty_key_on_right_hard_fails(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"], ["", "0", ""]])
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assert_no_output()

    def test_empty_key_on_right_reports_side_count_and_row_number(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([
            ["A-1", "10", "已发"],
            ["", "0", ""],
            ["A-2", "20", "已发"],
            ["", "0", ""],
            ["", "0", ""],
        ])
        before = (sha256(self.left), sha256(self.right))
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assert_no_output()
        self.assertIn("右表有 3 行的匹配键是空的", out)
        self.assertIn("第 3 行", out)
        self.assertIn("第 5 行", out)
        self.assertIn("第 6 行", out)
        self.assertNotIn("左表有", out)
        self.assertEqual((sha256(self.left), sha256(self.right)), before)

    def test_empty_keys_on_both_sides_reported_together(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"], ["", "合计", "10", ""]])
        self.right = self.right_xlsx([["A-1", "10", "已发"], ["", "0", ""]])
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assert_no_output()
        self.assertIn("左表有 1 行的匹配键是空的", out)
        self.assertIn("右表有 1 行的匹配键是空的", out)

    def test_empty_key_cannot_be_bypassed_by_any_argument(self):
        # H1 没有任何参数能把空键降级成"排除掉继续跑"：--allow-empty-keys 这个参数根本不存在，
        # 传进来应该被 argparse 直接拒绝（退出码非 0），而不是被当成一个开关生效。
        self.left = self.left_csv([["A-1", "甲", "10", "已发"], ["", "合计", "10", ""]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_cli([
            "--left", str(self.left), "--right", str(self.right),
            "--left-key", "订单号", "--right-key", "单号", "--compare", "金额",
            "--work", str(self.work), "--output", str(self.output_xlsx),
            "--report", str(self.output_report),
            "--allow-empty-keys",
        ])
        self.assertNotEqual(code, 0, out)
        self.assertIn("unrecognized arguments", out)
        self.assert_no_output()

    def test_no_excluded_sheet_is_ever_produced(self):
        # 成功运行的输出里只有固定的五张表，不存在 excluded
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, _ = self.run_main()
        self.assertEqual(code, 0)
        self.assertEqual(
            sheet_names(self.output_xlsx),
            ["summary", "matched", "differences", "only_left", "only_right"],
        )

    def test_empty_and_duplicate_reported_together(self):
        self.left = self.left_csv([
            ["A-1", "甲", "10", "已发"],
            ["A-1", "乙", "20", "已发"],
            ["", "合计", "30", ""],
        ])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assertIn("空", out)
        self.assertIn("重复", out)

    def test_all_keys_empty_hard_fails(self):
        self.left = self.left_csv([["", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assert_no_output()


# ==========================================================================
# T4 分类正确性
# ==========================================================================

class TestClassification(TempCase):
    def build(self, left_rows, right_rows):
        self.left = self.left_csv(left_rows)
        self.right = self.right_xlsx(right_rows)
        return self.run_main()

    def test_all_matched(self):
        code, out = self.build(
            [["A-1", "甲", "10", "已发"], ["A-2", "乙", "20", "待发"]],
            [["A-1", "10", "已发"], ["A-2", "20", "待发"]],
        )
        self.assertEqual(code, 0, out)
        actual = self.counts_from_xlsx()
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 2)
        self.assertEqual(actual[reconcile.SHEET_DIFFERENCES], 0)
        self.assertEqual(actual[reconcile.SHEET_ONLY_LEFT], 0)
        self.assertEqual(actual[reconcile.SHEET_ONLY_RIGHT], 0)

    def test_all_different(self):
        code, _ = self.build(
            [["A-1", "甲", "10", "已发"], ["A-2", "乙", "20", "待发"]],
            [["A-1", "99", "已发"], ["A-2", "88", "待发"]],
        )
        self.assertEqual(code, 0)
        actual = self.counts_from_xlsx()
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 0)
        self.assertEqual(actual[reconcile.SHEET_DIFFERENCES], 2)

    def test_only_left_and_only_right(self):
        code, _ = self.build(
            [["A-1", "甲", "10", "已发"]],
            [["A-9", "10", "已发"]],
        )
        self.assertEqual(code, 0)
        actual = self.counts_from_xlsx()
        self.assertEqual(actual[reconcile.SHEET_ONLY_LEFT], 1)
        self.assertEqual(actual[reconcile.SHEET_ONLY_RIGHT], 1)
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 0)
        self.assertEqual(actual[reconcile.SHEET_DIFFERENCES], 0)

    def test_mixed_classification_and_conservation(self):
        code, out = self.build(
            [
                ["A-1", "甲", "10", "已发"],   # 完全匹配
                ["A-2", "乙", "20", "待发"],   # 金额不同
                ["A-3", "丙", "33", "已发"],   # 金额不同（本用例只比金额）
                ["A-5", "戊", "50", "已发"],   # 仅左
            ],
            [
                ["A-1", "10", "已发"],
                ["A-2", "200", "待发"],
                ["A-3", "30", "已发"],
                ["A-6", "60", "已发"],         # 仅右
            ],
        )
        self.assertEqual(code, 0, out)
        actual = self.assert_conservation(left_rows=4, right_rows=4)
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 1)
        self.assertEqual(actual[reconcile.SHEET_DIFFERENCES], 2)
        self.assertEqual(actual[reconcile.SHEET_ONLY_LEFT], 1)
        self.assertEqual(actual[reconcile.SHEET_ONLY_RIGHT], 1)
        self.assertIn("完全匹配数量：1", out)
        self.assertIn("有差异数量：2", out)
        self.assertIn("仅左表存在数量：1", out)
        self.assertIn("仅右表存在数量：1", out)

    def test_key_sets_are_disjoint(self):
        code, _ = self.build(
            [["A-1", "甲", "10", "已发"], ["A-2", "乙", "20", "待发"], ["A-5", "戊", "50", "已发"]],
            [["A-1", "10", "已发"], ["A-2", "200", "待发"], ["A-6", "60", "已发"]],
        )
        self.assertEqual(code, 0)
        sets = {}
        for sheet in reconcile.DATA_SHEETS:
            rows = read_xlsx_rows(self.output_xlsx, sheet)
            sets[sheet] = {str(row[0]) for row in rows[1:] if str(row[0]) != ""}
        names = list(sets)
        for i in range(len(names)):
            for j in range(i + 1, len(names)):
                self.assertEqual(
                    sets[names[i]] & sets[names[j]], set(),
                    "{} 与 {} 的键集合相交了".format(names[i], names[j]))

    def test_summary_numbers_equal_actual_sheet_rows(self):
        code, _ = self.build(
            [["A-1", "甲", "10", "已发"], ["A-2", "乙", "20", "待发"], ["A-5", "戊", "50", "已发"]],
            [["A-1", "10", "已发"], ["A-2", "200", "待发"], ["A-6", "60", "已发"]],
        )
        self.assertEqual(code, 0)
        summary = read_xlsx_rows(self.output_xlsx, reconcile.SHEET_SUMMARY)
        values = {str(row[0]): row[1] for row in summary[1:] if row and str(row[0])}
        actual = self.counts_from_xlsx()
        self.assertEqual(int(values["完全匹配数量"]), actual[reconcile.SHEET_MATCHED])
        self.assertEqual(int(values["有差异数量"]), actual[reconcile.SHEET_DIFFERENCES])
        self.assertEqual(int(values["仅左表存在数量"]), actual[reconcile.SHEET_ONLY_LEFT])
        self.assertEqual(int(values["仅右表存在数量"]), actual[reconcile.SHEET_ONLY_RIGHT])
        self.assertEqual(int(values["左表数据行数"]), 3)
        self.assertEqual(int(values["右表数据行数"]), 3)

    def test_row_numbers_point_back_to_source_rows(self):
        code, _ = self.build(
            [["A-1", "甲", "10", "已发"], ["A-2", "乙", "20", "待发"]],
            [["A-1", "10", "已发"], ["A-2", "20", "待发"]],
        )
        self.assertEqual(code, 0)
        rows = read_xlsx_rows(self.output_xlsx, reconcile.SHEET_MATCHED)
        header = [str(cell) for cell in rows[0]]
        left_pos = header.index(reconcile.COL_ROW_LEFT)
        right_pos = header.index(reconcile.COL_ROW_RIGHT)
        # left.csv 表头在第 1 行 → 第一条数据在第 2 行
        self.assertEqual(int(rows[1][left_pos]), 2)
        self.assertEqual(int(rows[2][left_pos]), 3)
        self.assertEqual(int(rows[1][right_pos]), 2)

    def test_output_order_follows_source_rows(self):
        code, _ = self.build(
            [["A-3", "丙", "30", "已发"], ["A-1", "甲", "10", "已发"], ["A-2", "乙", "20", "已发"]],
            [["A-1", "10", "已发"], ["A-2", "20", "已发"], ["A-3", "30", "已发"]],
        )
        self.assertEqual(code, 0)
        rows = read_xlsx_rows(self.output_xlsx, reconcile.SHEET_MATCHED)
        self.assertEqual([str(row[0]) for row in rows[1:]], ["A-3", "A-1", "A-2"])


# ==========================================================================
# T5 比较语义
# ==========================================================================

class TestCompareSemantics(TempCase):
    def test_differences_shows_left_and_right_values(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "99", "已发"]])
        code, _ = self.run_main()
        self.assertEqual(code, 0)
        rows = read_xlsx_rows(self.output_xlsx, reconcile.SHEET_DIFFERENCES)
        header = [str(cell) for cell in rows[0]]
        self.assertIn(reconcile.COL_KEY, header)
        self.assertIn(reconcile.COL_DIFF_FIELDS, header)
        self.assertIn("金额_左", header)
        self.assertIn("金额_右", header)
        row = dict(zip(header, [str(cell) for cell in rows[1]]))
        self.assertEqual(row["金额_左"], "10")
        self.assertEqual(row["金额_右"], "99")
        self.assertEqual(row[reconcile.COL_DIFF_FIELDS], "金额")

    def test_partial_field_difference_lists_only_that_field(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "待发"]])
        code, _ = self.run_main(compare=("金额", "状态"))
        self.assertEqual(code, 0)
        rows = read_xlsx_rows(self.output_xlsx, reconcile.SHEET_DIFFERENCES)
        header = [str(cell) for cell in rows[0]]
        row = dict(zip(header, [str(cell) for cell in rows[1]]))
        self.assertEqual(row[reconcile.COL_DIFF_FIELDS], "状态")

    def test_same_name_compare(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]], header="订单号,客户,金额,状态")
        self.right = self.right_xlsx([["A-1", "10", "已发"]], header=("单号", "金额", "状态"))
        code, out = self.run_main(compare=("金额",))
        self.assertEqual(code, 0, out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)

    def test_left_equals_right_compare(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]], header="订单号,客户,应收金额,状态")
        self.right = self.right_xlsx([["A-1", "10", "已发"]], header=("单号", "金额", "状态"))
        code, out = self.run_main(compare=("应收金额=金额",))
        self.assertEqual(code, 0, out)
        self.assertIn("左「应收金额」= 右「金额」", out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)

    def test_missing_compare_column_on_left_fails_with_header_dump(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_main(compare=("不存在的列",))
        self.assertEqual(code, 1, out)
        self.assert_no_output()
        self.assertIn("不存在的列", out)
        self.assertIn("订单号", out)  # 必须把实际表头打出来帮用户改

    def test_missing_compare_column_on_right_fails(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_main(compare=("客户=客户",))
        self.assertEqual(code, 1, out)
        self.assert_no_output()
        self.assertIn("单号", out)

    def test_duplicate_compare_spec_fails(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_main(compare=("金额", "金额"))
        self.assertEqual(code, 1, out)
        self.assertIn("重复", out)

    def test_compare_equal_to_key_warns(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_main(compare=("订单号=单号",))
        self.assertEqual(code, 0, out)
        self.assertIn("同一列", out)

    def test_missing_compare_argument_rejected_by_argparse(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_cli([
            "--left", str(self.left), "--right", str(self.right),
            "--left-key", "订单号", "--right-key", "单号",
            "--work", str(self.work), "--output", str(self.output_xlsx),
            "--report", str(self.output_report),
        ])
        self.assertNotEqual(code, 0, out)
        self.assertIn("--compare", out)

    def test_matching_key_column_ambiguous_fails(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]], header="订单号,订单号,金额,状态")
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assertIn("出现了", out)


# ==========================================================================
# T6 公式策略
# ==========================================================================

@unittest.skipUnless(HAS_OPENPYXL, "没装 openpyxl")
class TestFormulaPolicy(TempCase):
    def make_right_with_formula(self, formula_cell, rows=None):
        path = self.tmp / "right_formula.xlsx"
        rows = rows or [["单号", "金额", "状态"], ["A-1", "10", "已发"]]
        make_formula_xlsx(path, "内部台账", rows, formula_cell=formula_cell)
        return path

    def test_formula_in_key_column_hard_fails(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.make_right_with_formula("A2")  # A 列 = 单号（匹配键列）
        before = (sha256(self.left), sha256(self.right))
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assert_no_output()
        self.assertIn("公式", out)
        self.assertIn("A2", out)
        self.assertIn("data_only=False", out)
        self.assertEqual((sha256(self.left), sha256(self.right)), before)

    def test_formula_in_compare_column_hard_fails(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.make_right_with_formula("B2")  # B 列 = 金额（比较字段列）
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assert_no_output()
        self.assertIn("参与列里有公式", out)

    def test_formula_in_unrelated_column_is_ignored_and_reported(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        # D 列是"备注"，不参与对账
        path = self.tmp / "right.xlsx"
        make_xlsx(path, [("内部台账", [
            ["单号", "金额", "状态", "备注"],
            ["A-1", "10", "已发", None],
        ])])
        with Path(path).open("rb") as handle:
            workbook = openpyxl.load_workbook(handle)
        try:
            workbook["内部台账"]["D2"] = "=1+1"
            workbook.save(path)
        finally:
            workbook.close()

        code, out = self.run_main(right=path)
        self.assertEqual(code, 0, out)
        self.assert_output_ok()
        self.assertIn("明确忽略", out)
        self.assertIn("不参与对账的列里有 1 处公式", out)

    def test_equal_prefixed_text_cell_accepted_with_warning(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        path = self.tmp / "right.xlsx"
        make_xlsx(path, [("内部台账", [
            ["单号", "金额", "状态"],
            ["A-1", "10", "=待确认"],
        ])])
        with Path(path).open("rb") as handle:
            workbook = openpyxl.load_workbook(handle)
        try:
            cell = workbook["内部台账"]["C2"]
            cell.value = "=待确认"
            cell.data_type = "s"
            workbook.save(path)
        finally:
            workbook.close()

        code, out = self.run_main(left=self.left, right=path, compare=("状态",))
        self.assertEqual(code, 0, out)
        self.assertIn("以 = 开头的**文本**单元格", out)

    def test_output_workbook_contains_no_formulas(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, _ = self.run_main()
        self.assertEqual(code, 0)
        with self.output_xlsx.open("rb") as handle:
            workbook = openpyxl.load_workbook(handle, data_only=False)
        try:
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        self.assertNotEqual(
                            cell.data_type, "f",
                            "输出里出现了公式：{}!{}".format(sheet.title, cell.coordinate))
        finally:
            workbook.close()

    def test_csv_equal_prefixed_text_is_kept_as_text_in_output(self):
        # 用户数据里本来就有 "=" 开头的文本时：读进来当文本，写出去也必须还是文本，
        # 绝不能让 openpyxl 把它变成公式（Excel 打开会显示 #NAME?）
        self.left = self.left_csv([["=A-1", "甲", "10", "已发"]])
        self.right = self.right_csv([["=A-1", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 0, out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)

        with self.output_xlsx.open("rb") as handle:
            workbook = openpyxl.load_workbook(handle, data_only=False)
        try:
            sheet = workbook[reconcile.SHEET_MATCHED]
            keys = [sheet.cell(row=row, column=1) for row in range(1, sheet.max_row + 1)]
            self.assertEqual(keys[1].value, "=A-1")
            self.assertEqual(keys[1].data_type, "s", "输出里 '=' 开头的文本被当成公式了")
        finally:
            workbook.close()


# ==========================================================================
# T7 工作表与表头行
# ==========================================================================

@unittest.skipUnless(HAS_OPENPYXL, "没装 openpyxl")
class TestSheetAndHeaderRow(TempCase):
    def test_single_sheet_auto_selected_and_reported(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]], sheet="唯一表")
        code, out = self.run_main()
        self.assertEqual(code, 0, out)
        self.assertIn("右表工作表：唯一表", out)

    def test_multiple_sheets_defaults_to_first_with_warning(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        path = self.tmp / "multi.xlsx"
        make_xlsx(path, [
            ("第一张", [["单号", "金额", "状态"], ["A-1", "10", "已发"]]),
            ("第二张", [["单号", "金额", "状态"], ["A-1", "999", "已发"]]),
        ])
        code, out = self.run_main(right=path)
        self.assertEqual(code, 0, out)
        self.assertIn("第一个", out)
        self.assertIn("第二张", out)  # 必须把可选的工作表名都列出来
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)

    def test_explicit_sheet_selects_the_right_one(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        path = self.tmp / "multi.xlsx"
        make_xlsx(path, [
            ("第一张", [["单号", "金额", "状态"], ["A-1", "999", "已发"]]),
            ("第二张", [["单号", "金额", "状态"], ["A-1", "10", "已发"]]),
        ])
        code, out = self.run_main(right=path, extra=["--right-sheet", "第二张"])
        self.assertEqual(code, 0, out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)

    def test_missing_sheet_fails_and_lists_available(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        path = self.tmp / "multi.xlsx"
        make_xlsx(path, [("第一张", [["单号", "金额", "状态"], ["A-1", "10", "已发"]])])
        code, out = self.run_main(right=path, extra=["--right-sheet", "不存在"])
        self.assertEqual(code, 1, out)
        self.assertIn("找不到工作表", out)
        self.assertIn("第一张", out)

    def test_header_row_offset(self):
        left = self.write_csv(
            "left.csv",
            "客户对账单\n制表日期：2026-09-01\n订单号,客户,金额,状态\nA-1,甲,10,已发\n",
        )
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_main(left=left, extra=["--left-header-row", "3"])
        self.assertEqual(code, 0, out)
        self.assertIn("左表表头行：第 3 行", out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)
        rows = read_xlsx_rows(self.output_xlsx, reconcile.SHEET_MATCHED)
        header = [str(cell) for cell in rows[0]]
        # 数据在第 4 行（物理行号）
        self.assertEqual(int(rows[1][header.index(reconcile.COL_ROW_LEFT)]), 4)

    def test_header_row_beyond_file_fails(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_main(extra=["--left-header-row", "99"])
        self.assertEqual(code, 1, out)
        self.assertIn("超出范围", out)

    def test_header_row_zero_rejected(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_cli([
            "--left", str(self.left), "--right", str(self.right),
            "--left-key", "订单号", "--right-key", "单号", "--compare", "金额",
            "--left-header-row", "0",
            "--work", str(self.work), "--output", str(self.output_xlsx),
            "--report", str(self.output_report),
        ])
        self.assertEqual(code, 1, out)
        self.assertIn("必须从 1 开始", out)


# ==========================================================================
# T8 编码与分隔符
# ==========================================================================

class TestEncodingAndDelimiter(TempCase):
    def test_gb18030_csv(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]], encoding="gb18030")
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 0, out)
        self.assertIn("gb18030", out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)

    def test_bom_csv(self):
        path = self.tmp / "left.csv"
        path.write_text("订单号,客户,金额,状态\nA-1,甲,10,已发\n", encoding="utf-8-sig")
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main(left=path)
        self.assertEqual(code, 0, out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)

    def test_semicolon_delimiter(self):
        path = self.write_csv("left.csv", "订单号;客户;金额;状态\nA-1;甲;10;已发\n")
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main(left=path, extra=["--left-delimiter", ";"])
        self.assertEqual(code, 0, out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)

    def test_semicolon_without_flag_fails(self):
        path = self.write_csv("left.csv", "订单号;客户;金额;状态\nA-1;甲;10;已发\n")
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main(left=path)
        self.assertEqual(code, 1, out)
        self.assertIn("找不到指定的列", out)

    def test_tab_delimiter(self):
        path = self.write_csv("left.csv", "订单号\t客户\t金额\t状态\nA-1\t甲\t10\t已发\n")
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main(left=path, extra=["--left-delimiter", "\t"])
        self.assertEqual(code, 0, out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 1)

    def test_multi_character_delimiter_rejected(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main(extra=["--left-delimiter", ";;"])
        self.assertEqual(code, 1, out)
        self.assertIn("必须是**一个**字符", out)

    def test_ragged_row_warns(self):
        path = self.write_csv("left.csv", "订单号,客户,金额,状态\nA-1,甲,10,已发\nA-2,乙\n")
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main(left=path)
        self.assertEqual(code, 0, out)
        self.assertIn("列数和表头不一致", out)
        # 缺的格子按空值处理 → A-2 的金额是空，不会崩，但键正常
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_ONLY_LEFT], 1)

    def test_uncertain_encoding_replacement_char_fails(self):
        # 造一个"解出来带替换字符"的文件：用 UTF-8 写入后再按 GB18030 读的典型症状
        path = self.tmp / "left.csv"
        path.write_bytes("订单号,客户,金额,状态\nA-1,甲,�,已发\n".encode("utf-8"))
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main(left=path)
        self.assertEqual(code, 1, out)
        self.assertIn("替换字符", out)


# ==========================================================================
# T9 输入安全
# ==========================================================================

class TestInputSafety(TempCase):
    def test_same_file_on_both_sides_fails(self):
        path = self.left_csv([["A-1", "甲", "10", "已发"]])
        code, out = self.run_main(left=path, right=path)
        self.assertEqual(code, 1, out)
        self.assertIn("同一个文件", out)
        self.assert_no_output()

    def test_output_pointing_at_left_input_is_refused(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_csv([["A-1", "10", "已发"]])
        before = sha256(self.left)
        code, out = self.run_main(output=self.left, extra=["--force"])
        self.assertEqual(code, 1, out)
        self.assertIn("永不覆盖", out)
        self.assertEqual(sha256(self.left), before)

    def test_report_pointing_at_right_input_is_refused(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_csv([["A-1", "10", "已发"]])
        before = sha256(self.right)
        code, out = self.run_main(report=self.right)
        self.assertEqual(code, 1, out)
        self.assertIn("永不覆盖", out)
        self.assertEqual(sha256(self.right), before)

    def test_symlink_to_input_is_refused(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_csv([["A-1", "10", "已发"]])
        link = self.tmp / "看着无害.xlsx"
        try:
            link.symlink_to(self.left)
        except (OSError, NotImplementedError):
            self.skipTest("当前文件系统不支持符号链接")
        before = sha256(self.left)
        code, out = self.run_main(output=link, extra=["--force"])
        self.assertEqual(code, 1, out)
        self.assertIn("符号链接", out)
        self.assertEqual(sha256(self.left), before)

    def test_inputs_unchanged_after_success(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        before = (sha256(self.left), sha256(self.right))
        code, _ = self.run_main()
        self.assertEqual(code, 0)
        self.assertEqual((sha256(self.left), sha256(self.right)), before)

    def test_inputs_unchanged_after_duplicate_failure(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"], ["A-1", "乙", "20", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        before = (sha256(self.left), sha256(self.right))
        code, _ = self.run_main()
        self.assertEqual(code, 1)
        self.assertEqual((sha256(self.left), sha256(self.right)), before)

    def test_inputs_unchanged_after_empty_key_failure(self):
        self.left = self.left_csv([["", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        before = (sha256(self.left), sha256(self.right))
        code, _ = self.run_main()
        self.assertEqual(code, 1)
        self.assertEqual((sha256(self.left), sha256(self.right)), before)

    def test_inputs_unchanged_after_missing_column_failure(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        before = (sha256(self.left), sha256(self.right))
        code, _ = self.run_main(compare=("没有这一列",))
        self.assertEqual(code, 1)
        self.assertEqual((sha256(self.left), sha256(self.right)), before)

    def test_unsupported_suffix_xls(self):
        left = self.write_csv("old.xls", "订单号,客户,金额,状态\nA-1,甲,10,已发\n")
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main(left=left)
        self.assertEqual(code, 1, out)
        self.assertIn("不支持的格式", out)

    def test_unsupported_suffix_xlsm_mentions_macro(self):
        left = self.write_csv("book.xlsm", "内容随便")
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main(left=left)
        self.assertEqual(code, 1, out)
        self.assertIn("宏", out)

    def test_missing_file(self):
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main(left=self.tmp / "不存在.csv")
        self.assertEqual(code, 1, out)
        self.assertIn("找不到左表文件", out)

    def test_empty_side_fails(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        right = self.right_xlsx([], header=("单号", "金额", "状态"))
        code, out = self.run_main(right=right)
        self.assertEqual(code, 1, out)
        self.assertIn("一行数据都没有", out)
        self.assert_no_output()

    def test_header_only_on_left_fails(self):
        self.left = self.left_csv([])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assertIn("一行数据都没有", out)

    def test_directory_as_input_fails(self):
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_main(left=self.tmp)
        self.assertEqual(code, 1, out)
        self.assertIn("不是一个普通文件", out)


# ==========================================================================
# T10 输出保护与提交
# ==========================================================================

class TestOutputProtection(TempCase):
    def setUp(self):
        super().setUp()
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])

    def test_existing_output_refused_without_force(self):
        self.output_xlsx.parent.mkdir(parents=True, exist_ok=True)
        self.output_xlsx.write_text("上一次的结果", encoding="utf-8")
        self.output_report.write_text("上一次的报告", encoding="utf-8")
        code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assertIn("--force", out)
        self.assertEqual(self.output_xlsx.read_text(encoding="utf-8"), "上一次的结果")
        self.assertEqual(self.output_report.read_text(encoding="utf-8"), "上一次的报告")

    def test_force_archives_previous_results(self):
        self.output_xlsx.parent.mkdir(parents=True, exist_ok=True)
        self.output_xlsx.write_text("上一次的结果", encoding="utf-8")
        self.output_report.write_text("上一次的报告", encoding="utf-8")
        code, out = self.run_main("--force")
        self.assertEqual(code, 0, out)
        self.assert_output_ok()
        previous = self.work / "previous"
        contents = {path.read_text(encoding="utf-8") for path in previous.iterdir()}
        self.assertIn("上一次的结果", contents)
        self.assertIn("上一次的报告", contents)
        self.assertIn("归档", out)

    def archived_texts(self):
        """work/previous/ 里所有**文本**文件的内容（归档里还有 xlsx，不能当文本读）。"""
        previous = self.work / "previous"
        return {
            path.read_text(encoding="utf-8")
            for path in previous.iterdir()
            if path.suffix == ".txt"
        }

    def test_previous_history_survives_repeated_force_in_same_second(self):
        self.output_xlsx.parent.mkdir(parents=True, exist_ok=True)
        self.output_xlsx.write_text("第一轮", encoding="utf-8")
        self.output_report.write_text("第一轮报告", encoding="utf-8")
        self.assertEqual(self.run_main("--force")[0], 0)
        previous = self.work / "previous"
        first_batch = sorted(path.name for path in previous.iterdir())
        self.assertEqual(len(first_batch), 2)

        self.assertEqual(self.run_main("--force")[0], 0)
        after = sorted(path.name for path in previous.iterdir())
        self.assertTrue(set(first_batch) <= set(after), "旧历史被覆盖了")
        self.assertIn("第一轮报告", self.archived_texts())

    def test_previous_survives_a_failing_run(self):
        self.output_xlsx.parent.mkdir(parents=True, exist_ok=True)
        self.output_xlsx.write_text("上一轮", encoding="utf-8")
        self.assertEqual(self.run_main("--force")[0], 0)
        previous = self.work / "previous"
        before = sorted(path.name for path in previous.iterdir())

        bad_left = self.left_csv([["A-1", "甲", "10", "已发"], ["A-1", "乙", "20", "已发"]],
                                 name="bad.csv")
        code, out = self.run_main("--force", left=bad_left)
        self.assertEqual(code, 1, out)
        self.assert_no_output()
        self.assertTrue(set(before) <= set(path.name for path in previous.iterdir()))

    def test_clear_work_dirs_never_touches_previous(self):
        previous = self.work / "previous"
        previous.mkdir(parents=True)
        keeper = previous / "20200101-000000_reconciliation.xlsx"
        keeper.write_text("历史结果", encoding="utf-8")
        (self.work / "staging").mkdir(parents=True)
        (self.work / "staging" / "junk.txt").write_text("x", encoding="utf-8")
        reconcile.clear_work_dirs(self.work)
        self.assertTrue(keeper.exists())
        self.assertEqual(list((self.work / "staging").iterdir()), [])

    def test_previous_not_in_clearable_list(self):
        self.assertNotIn("previous", reconcile.CLEARABLE_WORK_SUBDIRS)

    def test_commit_failure_rolls_back_everything(self):
        self.output_xlsx.parent.mkdir(parents=True, exist_ok=True)
        calls = {"n": 0}
        real_install = reconcile.install_file

        def flaky(src, dst):
            calls["n"] += 1
            if calls["n"] == 2:
                raise OSError("模拟第二个文件提交失败")
            return real_install(src, dst)

        with unittest.mock.patch.object(reconcile, "install_file", flaky):
            code, out = self.run_main()
        self.assertEqual(code, 1, out)
        self.assertIn("回滚", out)
        self.assert_no_output()
        self.assertEqual(list(self.output_xlsx.parent.iterdir()), [])

    def test_successful_commit_leaves_exactly_two_files(self):
        code, _ = self.run_main()
        self.assertEqual(code, 0)
        self.assertEqual(
            sorted(path.name for path in self.output_xlsx.parent.iterdir()),
            ["reconciliation.xlsx", "report.txt"],
        )

    def test_failed_install_leaves_no_temp_litter(self):
        source = self.tmp / "src.bin"
        source.write_bytes(b"data")
        self.output_xlsx.parent.mkdir(parents=True, exist_ok=True)
        with unittest.mock.patch.object(reconcile.os, "replace",
                                        side_effect=OSError("模拟改名失败")):
            with self.assertRaises(OSError):
                reconcile.install_file(source, self.output_xlsx)
        self.assertFalse(self.output_xlsx.exists())
        self.assertEqual(list(self.output_xlsx.parent.iterdir()), [])

    def test_refusal_does_not_create_output_parent_dir(self):
        # 空键会让流程在写盘前就被拒绝 → 连 output 目录都不该被创建
        bad_left = self.left_csv([["", "甲", "10", "已发"]], name="bad.csv")
        code, _ = self.run_main(left=bad_left)
        self.assertEqual(code, 1)
        self.assertFalse(self.output_xlsx.parent.exists())

    def test_unique_path_never_returns_existing_name(self):
        directory = self.tmp / "hist"
        directory.mkdir()
        first = directory / "x.txt"
        first.write_text("一", encoding="utf-8")
        second = reconcile.unique_path(directory, "x.txt")
        self.assertNotEqual(second, first)
        second.write_text("二", encoding="utf-8")
        third = reconcile.unique_path(directory, "x.txt")
        self.assertNotIn(third, (first, second))
        self.assertEqual(first.read_text(encoding="utf-8"), "一")

    def test_failed_run_writes_failure_report_into_work(self):
        bad_left = self.left_csv([["", "甲", "10", "已发"]], name="bad.csv")
        code, out = self.run_main(left=bad_left)
        self.assertEqual(code, 1)
        self.assert_no_output()
        failed = self.work / "reports" / reconcile.FAILED_REPORT_NAME
        self.assertTrue(failed.exists())
        self.assertIn(reconcile.VERSION_NAME, failed.read_text(encoding="utf-8"))


# ==========================================================================
# T11 报告
# ==========================================================================

class TestReport(TempCase):
    def setUp(self):
        super().setUp()
        self.left = self.left_csv([["A-1", "甲", "10", "已发"], ["A-2", "乙", "20", "待发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"], ["A-2", "200", "待发"]])

    def test_stdout_equals_report_file(self):
        code, out = self.run_main()
        self.assertEqual(code, 0)
        text = self.output_report.read_text(encoding="utf-8")
        self.assertEqual(out.strip(), text.strip())

    def test_report_contains_required_sections(self):
        code, out = self.run_main()
        self.assertEqual(code, 0)
        for needed in [
            reconcile.VERSION_NAME,
            "一、运行参数", "二、文件指纹", "三、行数统计", "四、公式与安全策略",
            "五、警告", "六、输出",
            "匹配键：左「订单号」= 右「单号」",
            "比较字段：左「金额」= 右「金额」",
            "左表数据行数：2", "右表数据行数：2",
            "SHA256", "严格",
        ]:
            self.assertIn(needed, out, "报告缺少：{}".format(needed))

    def test_report_reports_warning_section_empty_when_nothing_to_warn(self):
        code, out = self.run_main()
        self.assertEqual(code, 0)
        self.assertIn("五、警告", out)
        self.assertIn("（无）", out)

    def test_report_and_xlsx_numbers_agree(self):
        code, out = self.run_main()
        self.assertEqual(code, 0)
        actual = self.counts_from_xlsx()
        summary = read_xlsx_rows(self.output_xlsx, reconcile.SHEET_SUMMARY)
        values = {str(row[0]): row[1] for row in summary[1:] if row and str(row[0])}
        self.assertEqual(int(values["完全匹配数量"]), actual[reconcile.SHEET_MATCHED])
        self.assertIn("完全匹配数量：{}".format(actual[reconcile.SHEET_MATCHED]), out)

    def test_summary_sheet_echoes_run_config(self):
        code, _ = self.run_main(compare=("金额", "状态"))
        self.assertEqual(code, 0)
        summary = read_xlsx_rows(self.output_xlsx, reconcile.SHEET_SUMMARY)
        labels = {str(row[0]) for row in summary[1:] if row}
        for needed in ["版本", "左表文件", "右表文件", "匹配键（左）", "匹配键（右）",
                       "比较字段", "左表 SHA256", "右表 SHA256", "匹配/比较口径",
                       "行号含义", "输出文件", "报告文件"]:
            self.assertIn(needed, labels, "summary 缺少：{}".format(needed))
        self.assertEqual([str(cell) for cell in summary[0]], ["项目", "内容"])

    def test_summary_lists_all_five_required_sheets(self):
        code, _ = self.run_main()
        self.assertEqual(code, 0)
        names = sheet_names(self.output_xlsx)
        self.assertEqual(names, ["summary", "matched", "differences", "only_left", "only_right"])


# ==========================================================================
# T12 端到端与确定性
# ==========================================================================

class TestEndToEnd(TempCase):
    def test_demo_data_end_to_end(self):
        if not DEMO_LEFT.exists() or not DEMO_RIGHT.exists():
            self.skipTest("演示数据不在，跳过")
        before = (sha256(DEMO_LEFT), sha256(DEMO_RIGHT))
        code, out = self.run_main(
            left=DEMO_LEFT, right=DEMO_RIGHT,
            left_key="订单号", right_key="单号",
            compare=("应收金额=金额", "状态"),
            extra=["--left-header-row", "3", "--right-sheet", "内部台账"],
        )
        self.assertEqual(code, 0, out)
        self.assert_output_ok()
        actual = self.assert_conservation(left_rows=5, right_rows=5)
        self.assertEqual(actual[reconcile.SHEET_MATCHED], 2)
        self.assertEqual(actual[reconcile.SHEET_DIFFERENCES], 2)
        self.assertEqual(actual[reconcile.SHEET_ONLY_LEFT], 1)
        self.assertEqual(actual[reconcile.SHEET_ONLY_RIGHT], 1)
        self.assertEqual((sha256(DEMO_LEFT), sha256(DEMO_RIGHT)), before)

    def test_same_input_produces_identical_data_sheets(self):
        # 四张数据表必须逐格一致（summary 里记了运行时间，所以整文件不参与比较）
        self.left = self.left_csv([["A-1", "甲", "10", "已发"], ["A-2", "乙", "20", "待发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"], ["A-2", "200", "待发"]])
        first = self.tmp / "out1.xlsx"
        second = self.tmp / "out2.xlsx"
        self.assertEqual(self.run_main(output=first, report=self.tmp / "r1.txt")[0], 0)
        self.assertEqual(self.run_main(output=second, report=self.tmp / "r2.txt")[0], 0)
        for sheet in reconcile.DATA_SHEETS:
            self.assertEqual(
                read_xlsx_rows(first, sheet), read_xlsx_rows(second, sheet),
                "工作表 {} 两次运行结果不一致".format(sheet))

    def test_same_input_produces_identical_report_except_time(self):
        # 同一组输入、同一个输出位置跑两次（第二次 --force）→ 除了运行时间那一行，报告应完全一致
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_xlsx([["A-1", "10", "已发"]])
        first_code, first = self.run_main()
        self.assertEqual(first_code, 0)
        second_code, second = self.run_main("--force")
        self.assertEqual(second_code, 0)

        def strip(text):
            # 两次运行之间只有两处**应该**不同：
            #   - 运行时间那一行
            #   - --force 才有的"旧结果归档到哪了"那一段
            text = text.split("本次用 --force 归档的旧结果")[0]
            return [line for line in text.splitlines() if "运行时间" not in line]

        self.assertEqual(strip(first), strip(second))

    def test_cli_exit_code_zero_on_success(self):
        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_cli([
            "--left", str(self.left), "--right", str(self.right),
            "--left-key", "订单号", "--right-key", "单号", "--compare", "金额",
            "--work", str(self.work), "--output", str(self.output_xlsx),
            "--report", str(self.output_report),
        ])
        self.assertEqual(code, 0, out)
        self.assert_output_ok()

    def test_cli_exit_code_nonzero_on_failure(self):
        left = self.left_csv([["A-1", "甲", "10", "已发"], ["A-1", "乙", "20", "已发"]],
                             name="dup.csv")
        self.right = self.right_csv([["A-1", "10", "已发"]])
        code, out = self.run_cli([
            "--left", str(left), "--right", str(self.right),
            "--left-key", "订单号", "--right-key", "单号", "--compare", "金额",
            "--work", str(self.work), "--output", str(self.output_xlsx),
            "--report", str(self.output_report),
        ])
        self.assertEqual(code, 1, out)
        self.assert_no_output()


# ==========================================================================
# T13 边界数据
# ==========================================================================

@unittest.skipUnless(HAS_OPENPYXL, "没装 openpyxl")
class TestEdgeData(TempCase):
    def test_blank_rows_are_ignored_and_counted(self):
        path = self.write_csv(
            "left.csv", "订单号,客户,金额,状态\nA-1,甲,10,已发\n,,\nA-2,乙,20,已发\n,,\n")
        self.right = self.right_xlsx([["A-1", "10", "已发"], ["A-2", "20", "已发"]])
        code, out = self.run_main(left=path)
        self.assertEqual(code, 0, out)
        self.assertIn("整行空白", out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 2)
        self.assert_conservation(left_rows=2, right_rows=2)

    def test_hidden_rows_still_participate_with_warning(self):
        path = self.tmp / "right.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "内部台账"
        sheet.append(["单号", "金额", "状态"])
        sheet.append(["A-1", "10", "已发"])
        sheet.append(["A-2", "20", "已发"])
        sheet.row_dimensions[3].hidden = True
        workbook.save(path)
        workbook.close()

        self.left = self.left_csv([["A-1", "甲", "10", "已发"], ["A-2", "乙", "20", "已发"]])
        code, out = self.run_main(right=path)
        self.assertEqual(code, 0, out)
        self.assertIn("隐藏行", out)
        self.assertEqual(self.counts_from_xlsx()[reconcile.SHEET_MATCHED], 2)

    def test_autofilter_state_is_reported(self):
        path = self.tmp / "right.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "内部台账"
        sheet.append(["单号", "金额", "状态"])
        sheet.append(["A-1", "10", "已发"])
        sheet.auto_filter.ref = "A1:C2"
        workbook.save(path)
        workbook.close()

        self.left = self.left_csv([["A-1", "甲", "10", "已发"]])
        code, out = self.run_main(right=path)
        self.assertEqual(code, 0, out)
        self.assertIn("筛选状态", out)

    def test_stacked_tables_in_one_sheet_are_detected_and_reported(self):
        # 一个 sheet 里上下堆了两块表：第二块的表头会变成一条普通记录（键="单号"）。
        # 工具不猜表结构，但必须把"这块看起来是堆叠的"明确指出来。
        path = self.tmp / "right.xlsx"
        make_xlsx(path, [("内部台账", [
            ["单号", "金额", "状态"],
            ["A-1", "10", "已发"],
            ["", "", ""],
            ["单号", "金额", "状态"],
            ["A-2", "20", "已发"],
        ])])
        left = self.left_csv([["A-1", "甲", "10", "已发"]])
        code, out = self.run_main(left=left, right=path)
        self.assertEqual(code, 0, out)
        self.assertIn("上下堆了两块数据", out)
        self.assertIn("整行空白", out)


if __name__ == "__main__":
    unittest.main(verbosity=2)
