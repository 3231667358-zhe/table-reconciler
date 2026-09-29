# PROJECT CONTEXT

> 给 AI 助手读的项目上下文。新会话先看这份文件，不要重新问一遍已经定下来的事。

## 是什么

**Table Reconciler H1** —— Excel / CSV 两表匹配与对账差异检测。
命令行工具：按用户显式指定的匹配键把左右两表对齐，输出对账结果工作簿 + 文本报告。原文件永不改动。

**命名约定：版本一律叫 H1 / H2 / H3……，不再使用 V1 / V2。**

## 版本

H1，2026-09-29 完成。122 项自检用例全过，0 ResourceWarning
（`python3 -W error::ResourceWarning selftest.py`）。

## 位置与环境（实测）

- 目录：`/home/mei/ai-automation/table-reconciler`
- Ubuntu 24.04 / Python 3.12.3（`/usr/bin/python3`）
- 本机没有 pip、没有 pandas；openpyxl 3.1.2 已装（apt 的 `python3-openpyxl`）
- 不引入任何第三方依赖：标准库 `csv/argparse/pathlib/decimal/datetime`；只有读写 `.xlsx` 才用 openpyxl
- 姊妹项目（**本工具不依赖它们的运行时，也不修改它们**）：
  `/home/mei/ai-automation/` 下的 `excel-cleaner`、`table-merger`、`excel-workflow`

## H1 的硬边界（小环明确要求，不许放宽）

1. **匹配键、比较字段必须由用户显式指定**，程序不猜、不自动挑列、不做 AI 字段识别。
2. **只支持单列匹配键**（左一列 = 右一列）。复合键明确留到 H2。
3. **只支持一对一匹配**。任一侧键重复 → 硬失败，绝不偷偷做一对多。
4. **任一侧匹配键为空 → 整批硬失败**：退出码非 0，不生成 `reconciliation.xlsx`、不生成成功 report.txt。
   **不允许任何参数绕过**，**不产出 `excluded` 工作表**，也不"排除掉继续跑"。
   （曾短暂存在过的 `--allow-empty-keys` 与 excluded 逻辑已彻底删除，**不许再加回来**。）
5. **匹配键按文本处理**：`00123` 绝不等于 `123`。
6. **默认口径就是最严格的**：除"去除首尾空白"外不做任何变换。
   不忽略大小写、不做全半角、不去千分位、不做数值容差、不做日期识别、不做相似匹配。
   宁可把业务上本相同的数据列进 differences，也绝不过度标准化产生"假匹配"。
   一切宽松规则属于 H2，将来加也必须显式开启。
7. 支持 `.csv` / `.txt` / `.xlsx`；`.xls` / `.xlsm` / `.xlsb` 硬失败。
8. 原始文件绝对只读；写入路径不许指向 `--left` / `--right`，与 `--force` 无关。
9. 两个交付物先在 `work/staging/` 完整生成并校验，再一起提交；提交失败回滚，不留半套。
10. 输出已存在 → 默认拒绝；`--force` 归档到 `work/previous/`（不删除、不覆盖历史）。
11. 不加 H2 功能：多对多、复合键、模糊匹配、容差、日期识别、GUI/Web/DB/OCR/EXE。

## 已定的架构决策（含理由）

1. **默认工作表 = 第一张、默认表头 = 第 1 行**，严格照默认值走，不猜。
   多张工作表且未指定时**仍用第一张**，但报告里必须警告并列出所有工作表名（透明，不猜）。
2. **只有一个规范化**：`trim_text()` = `str.strip()`（Python 的 strip 是 Unicode 感知的，
   能去掉 U+00A0 / U+3000 这类"看不见的空白"）。
   零宽字符（U+200B/U+200C/U+200D/U+2060/U+FEFF）与**中间的** NBSP/全角空格**一律保留**，
   只检测 + 在报告里告警（去掉它们有可能改数据含义）。
3. **`render_cell_text()` 是匹配与比较的唯一入口**，规则：
   `None→""`；`str→原样`；`bool→TRUE/FALSE`；`int→十进制`；
   **`float` 整数值→整数文本**（1200.0→"1200"，这是"同一个数写成同一种形式"，不是容差）；
   `float` 非整数→`Decimal(repr(v))` 定点展开（绝不出科学计数法，尾差如实保留）；
   零点整的 `datetime`→只写日期。
4. **比较字段写法的两种形态**：`金额`（两边同名）/ `应收金额=金额`（左名=右名）。
   解析结果统一成 5 元组 `(左列下标, 右列下标, 左列名, 右列名, 展示名)`，
   全流程都用这一个结构 —— 早期版本用 3 元组时，报告里曾把列下标当列名打出来。
5. **关键校验闸门 `validate_staged()`**（提交前，任何一个不过就禁止靠近 output/）：
   - 五张表齐全（summary + 四张数据表）
   - 四张表的行数 == 统计数字
   - **守恒**：`matched+differences+only_left == 左表行数`，右侧同理
     （空键是整批硬失败，所以守恒式里**没有** excluded 项）
   - **互斥**：四张表的键集合两两不相交
   - `differences` 的「差异字段」取值必须都在 `--compare` 指定的集合里
   - 报告里的六项数字与实测一致；报告内容与刚校验的那份逐字节相同
6. **提交是"目标目录内先复制成临时文件、再同目录改名"**（`install_file()`），
   每个文件的"出现"是原子的；顺序 = 先 `report.txt`、最后 `reconciliation.xlsx`；
   任一步失败 → 回滚删掉本次已装的，绝不留半套；失败时清掉自己的 `.commit-*` 临时文件。
7. **`work/previous/` 用"时间戳+原名"，冲突再加 `-2`/`-3`**（`unique_path()`）。
   `CLEARABLE_WORK_SUBDIRS = (staging, reports)` —— previous/ **绝不在清理名单里**，有用例守着。
8. **输出目录的父目录在校验通过之后才创建**，这样"被拒绝时连父目录都不创建"才是真的。
9. **两次运行的数据表内容要一致**：xlsx 的 `created/modified` 固定为 2000-01-01
   （真实运行时间记录在 report.txt 与 summary 表里）。
10. **行号一律用源文件里的物理行号**（Excel 里看到的那一行），便于回去核对。

## 历史 bug 记录（别改回去）

| bug | 根因 | 修法 |
|---|---|---|
| 输出工作簿里 `=` 开头的文本变成了公式（Excel 显示 #NAME?） | openpyxl 把以 `=` 开头的字符串当公式写 | 新增 `append_row()`：追加后把这类单元格的 `data_type` 强制标成 `"s"`。所有写单元格的地方都必须走它 |
| `-o` 指向源文件 / 符号链接指向源文件时，报的是"后缀必须 .xlsx" | `check_paths()` 的格式检查排在写入目标闸门前面 | 改成 `check_write_targets()` **先**执行；另外写入目标**先不要 resolve**（否则 `is_symlink()` 永远为假，"是符号链接"的提示不会出现） |
| **范围漂移：擅自加了 `--allow-empty-keys` + `excluded` 工作表** | 违反 H1 硬边界"空键必须整批硬失败、不留任何绕过口子" | 彻底删除参数与 excluded 逻辑。`selftest.py` 加了负向用例：传 `--allow-empty-keys` 必须被 argparse 以 `unrecognized arguments` 拒绝。**不许恢复** |
| 输出目录不存在时提交报 `No such file or directory: output/.commit-*.txt` | 提交前没建输出目录的父目录 | 在 `validate_staged()` 通过之后、`commit_outputs()` 之前 `mkdir(parents=True)` |
| 报告里打出 `左「2」= 右「1」`（列下标当列名） | compare 元组用 3 元组、按位置解包错了 | compare 统一成 5 元组，全流程显式命名解包 |

## 文件清单

| 文件 | 作用 |
|---|---|
| `reconcile.py` | 唯一入口，约 1800 行 |
| `selftest.py` | 自检 122 用例（T1 渲染 / T2 规范化 / T3 唯一性与空键 / T4 分类 / T5 比较语义 / T6 公式策略 / T7 工作表与表头行 / T8 编码与分隔符 / T9 输入安全 / T10 输出与提交 / T11 报告 / T12 端到端 / T13 边界数据） |
| `examples/客户对账单.csv` | 演示左表：前两行标题、表头在第 3 行、`  A-1001 ` 带首尾空格、金额列有一处带空格 |
| `examples/内部台账.xlsx` | 演示右表：工作表「内部台账」，含"备注/经办人"两个不参与对账的列 |
| `output/` `work/` `input/` | 运行时目录（`.gitkeep` 占位） |

演示数据预期结果：左 5 行 / 右 5 行 → 完全匹配 2、有差异 2、仅左 1、仅右 1。

## 验证方式

```bash
cd /home/mei/ai-automation/table-reconciler
python3 -W error::ResourceWarning selftest.py     # 122 用例，看到 OK 即全过
python3 reconcile.py \
  --left examples/客户对账单.csv --left-header-row 3 --left-key 订单号 \
  --right examples/内部台账.xlsx --right-sheet 内部台账 --right-key 单号 \
  --compare 应收金额=金额,状态
```

## H1 里刻意"不做"的（别当成 bug）

- 不做数值等价判断：`"1200.00"`（文本）vs `1200`（数值）**会**报差异，只在报告里提醒类型不一致。
- 不跳过隐藏行（跳过就是静默丢数据），只告警。
- 不猜表结构：一个工作表里上下堆两块数据时，只提醒"某行的键和表头名一样"。
- 单侧 0 数据行 → 硬失败（几乎总是选错了文件/工作表/表头行），不产出"全部只在另一侧"的结果。

## H2 候选（未决定，别自己加）

复合键 · 多对多聚合对账 · 金额容差与四舍五入规则 · 日期智能识别 · 忽略大小写/全半角的显式开关 ·
"I 只差大小写/全半角"的配对诊断（不改匹配结果，只提示）· 跨多文件对账 · 增量对账 ·
多工作表批量 · 三方以上对账 · 差异分类与容忍规则 · 调整分录生成 · 配置存档（月度重复对账）·
GUI / Web / 数据库 / OCR / AI 字段识别 / Windows EXE
