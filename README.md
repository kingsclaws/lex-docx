# lex_docx

DOCX 自动化工具库，适用于法律尽调报告、合同审阅等场景。
支持中英文文档，供 Agent 调用，无外部 API 依赖，纯本地操作。

---

## 安装依赖

```bash
pip install python-docx lxml
```

---

## 快速开始

```python
from docx import Document
from lex_docx import DocConfig

# 律所级配置（一次性设定，传入各函数）
cfg = DocConfig(
    author="JT",
    note_prefix="[JT Note: ",
    note_suffix="]",
    note_highlight="yellow",
    header_shading="D9E2F3",
)

doc = Document("report.docx")
```

---

## 模块说明

### 1. `format_brush` — 格式刷

修复 TC INS 注入后新段落丢失缩进、段间距、样式的问题。

```python
from lex_docx import format_brush

# 从参考段落复制格式到目标段落
format_brush.apply(
    doc,
    target_indices=[177, 178, 180],
    reference_index=171,
    copy=["indent", "spacing", "style"],
)

# 按 style name 批量修复
format_brush.auto_fix(doc, para_range=(175, 190))

# 查看段落格式摘要（调试用）
summaries = format_brush.get_pPr_summary(doc, para_range=(170, 185))

# 从 styles.xml 提取全量 style rPr 映射（供 tc_utils 的 Style-Aware 注入使用）
style_map = format_brush.extract_style_rPr_map(doc)
# 返回 {"Normal": <rPr element>, "Heading 2": <rPr element>, ...}
```

---

### 2. `jt_note` — 律所批注注入

将律所 Note 以 Track Changes INS 形式注入文档，支持任意律所格式。

```python
from lex_docx import jt_note

# 方式 A：在已有段落末尾追加
jt_note.append_to_paragraph(doc, paragraph_index=180,
    note_text="待确认是否存在其他未披露的仲裁案件。", cfg=cfg)

# 方式 B：插入独立 Note 段落
jt_note.insert_paragraph(doc, after_index=180,
    note_text="待获取征信报告后替换本段。",
    inherit_style_from=180, cfg=cfg)

# 方式 C：普通文字 + Note 混排
jt_note.create_mixed_paragraph(doc, after_index=176,
    segments=[
        ("根据本所律师查询，", False),
        ("待获取征信报告。", True),
    ], cfg=cfg)

# 修复存量 Note 格式
fixed = jt_note.fix_note_format(doc, cfg=cfg)
```

**Note 格式效果**：`[JT Note: 待确认...]`，粗体 + 斜体 + 黄色高亮，包裹在 `w:ins`。

---

### 3. `tc_utils` — Track Changes 底层工具

所有 TC INS / DEL XML 构造的统一入口。

#### 段落级 TC DEL + INS

```python
from lex_docx.tc_utils import tc_del_paragraph, tc_ins_text, tc_ins_mixed, next_tc_id

tc_id = next_tc_id(doc)

# 删除整段（返回原 run 的 rPr，供后续继承）
saved_rPr = tc_del_paragraph(para, tc_id, "JT")

# 插入新文字，inherit_rPr 继承策略：
tc_ins_text(para, "新文字", tc_id+1, "JT",
    inherit_rPr=True)           # 从本段 DEL run 继承字体字号（最常用）

tc_ins_text(para, "新文字", tc_id+1, "JT",
    inherit_rPr="style")        # 不设 rPr，让 Word 从 pStyle 继承

tc_ins_text(para, "新文字", tc_id+1, "JT",
    inherit_rPr="auto",         # 按段落 pStyle 查 style_rPr_map
    style_rPr_map={"Normal": {"eastAsia": "仿宋_GB2312", "sz": "24"}})

tc_ins_text(para, "新文字", tc_id+1, "JT",
    inherit_rPr=doc.paragraphs[335])  # 从指定段落第一个 run 复制

# position 参数：
tc_ins_text(para, "前置文字", tc_id, "JT", position="start")
tc_ins_text(para, "末尾文字", tc_id, "JT", position="end")   # 默认
tc_ins_text(para, "插在第2个run后", tc_id, "JT", position=2)
```

#### 混排注入（普通文字 + Note）

```python
# 在段落末尾插入混合内容
tc_ins_mixed(para, [
    ("根据本所律师查询，", False),     # 普通文字（继承 rPr）
    ("待确认。", True),                # Note（自动加 prefix/suffix + B+I+HL）
], tc_id, "JT", cfg=cfg, inherit_rPr=True)
```

#### rPr 配置字典构造

```python
from lex_docx.tc_utils import make_rPr_from_dict

rPr = make_rPr_from_dict({
    "eastAsia": "仿宋_GB2312",
    "ascii":    "Times New Roman",
    "sz":       "24",       # 12pt（half-points）
    "b":        True,
    "highlight": "yellow",
})
```

---

### 4. `defined_terms` — 定义术语加粗

自动检测并加粗首次定义的术语，支持中英文合同。

| 模式 | 示例 |
|------|------|
| 中文括号 | `（"漕河泾总公司"）` / `（"甲方"或"乙方"）` |
| 中文行内 | `以下简称"甲方"` / `合称"各方"` / `下称"目标公司"` |
| 英文括号 | `("Borrower")` / `(the "Lender")` / `(each a "Party")` |
| 英文行内 | `hereinafter "ABC"` / `collectively the "Parties"` |

```python
from lex_docx import defined_terms

terms = defined_terms.auto_bold(doc, paragraph_index=39)
defined_terms.bold_terms(doc, paragraph_index=39, terms=["中信金资", "贵司"])
results = defined_terms.scan_terms(doc, para_range=(0, 100))
```

---

### 5. `table_ops` — 表格操作

#### 提取

```python
from lex_docx import table_ops

data = table_ops.extract_table("autodocs.docx",
    near_text="目前公司的股东情况", output="list_of_dicts")
data = table_ops.extract_table("autodocs.docx", table_index=3)
```

#### 填充普通表格

```python
table_ops.fill_table(
    doc, table_index=12, data=data,
    column_mapping={
        "序号": "auto_number",
        "股东姓名/名称": "股东名称",
        "出资额（万元）": "认缴出资额\n（人民币/万元）",  # 含换行也能匹配
    },
    auto_add_rows=True,   # 数据行不足时自动 TC INS 增行
    cfg=cfg,
)
```

#### 填充 KV 表（基本信息表）

```python
table_ops.fill_kv_table(
    doc, table_index=8,
    data={
        "企业名称": "上海漕河泾新兴技术开发区总公司",
        # compound key：文档中任意一种写法均可匹配
        "法定代表人/负责人/执行事务合伙人": "桂恩亮",
        "注册资本": "45,325万元",
    },
    key_columns=[0, 2],   # 四列布局（key|val|key|val），默认 None = 两列
    cfg=cfg,
)
```

#### 增删行 / 格式统一

```python
table_ops.adjust_rows(doc, table_index=12, target_data_rows=5, cfg=cfg)

table_ops.format_table(doc, table_index=12,
    header_shading="D9E2F3", borders="single",
    column_widths=[800, 4000, 2000, 2000],
    column_alignments=["center", "left", "right", "right"],
    header_row_index=0, cfg=cfg)
```

**所有 `table_ops` 函数均支持 `header_row_index`**（标题行不在第 0 行时使用）。

#### 跨文档表格复制（FR-9）

从 AutoDocs 或其他源文档直接搬运表格到报告，保留完整格式（字体/字号/边框/合并单元格）。

```python
from lex_docx import table_ops

# 方式 A：完整复制（保留源表格格式 + 数据）
table_ops.copy_table(
    src_doc="autodocs.docx",
    src_table_index=3,           # 或 src_near_text="变更时间"
    dst_doc=doc,
    dst_position="after_para:241",   # P241 之后插入
    # dst_position="replace_table:15",  # 或替换目标文档第 15 个表格
    as_tc_ins=True,              # 整张表标记为 TC INS（默认 True）
    cfg=cfg,
)

# 方式 B：复制时变换数据（过滤列/行、重命名表头、限制行数）
table_ops.copy_table(
    src_doc="autodocs.docx",
    src_table_index=3,
    dst_doc=doc,
    dst_position="replace_table:15",
    transform={
        "columns": [0, 1, 2, 3],           # 只保留前 4 列
        "rename_headers": {"变更时间": "日期", "变更类型": "变更事项"},
        "filter_rows": lambda row: row["变更类型"] in ["股东股权变更", "负责人变更"],
        "max_rows": 30,
    },
    cfg=cfg,
)

# 方式 C：仅复制格式，不复制数据（给现有表格套格式）
table_ops.copy_table_format(
    src_doc="template.docx",
    src_table_index=8,
    dst_doc=doc,
    dst_table_index=15,
    # 复制：列宽、边框、底色、字体、对齐
    # 保留：原有单元格内容不变
)
```

**常见用途**：变更记录表（28 行）、股东表、诉讼表、行政处罚表从 AutoDocs 原样搬运到报告。

> ⚠️ `src_near_text` 仅搜索段落文本，若关键词只出现在表格 header 单元格内则无法定位。此时改用 `src_table_index` 直接指定序号即可。

---

### 6. `cleanup` — 注入后清理

```python
from lex_docx import cleanup

# 一次性清理空段落 + 孤儿编号
result = cleanup.cleanup_all(doc, para_range=(206, 265), cfg=cfg)
print(result)  # {"empty": [208, 212, ...], "orphan_numbering": [220]}

# 分开调用
cleanup.remove_empty_paragraphs(doc,
    as_tc_del=True,          # True = TC DEL 标记（推荐）；False = 直接删除
    keep_styles=["Heading 1"],  # 保留这些 style 的空段落
    cfg=cfg)

cleanup.remove_orphan_numbering(doc, cfg=cfg)
```

---

### 7. `inject_engine` — 批量注入引擎

将整个章节的注入工作封装为结构化计划，一键执行。

```python
from lex_docx import inject_engine, DocConfig

cfg = DocConfig(author="JT", entity_names={"allowed": ["临港资管"]})

plan = inject_engine.InjectPlan(
    doc_path="report.docx",
    out_path="report_reviewed.docx",
    target_range=(266, 332),       # 段落范围（用于 cleanup + lint）
    tables=[
        inject_engine.TableFill(
            table_index=8,
            data={"企业名称": "临港资管", "法定代表人": "张三", "注册资本": "1060万"},
            mode="kv",
        ),
        inject_engine.TableFill(
            table_index=12,
            data=[{"股东": "临港集团", "出资额": "1060万", "比例": "100%"}],
            mode="rows",
            auto_adjust=True,      # 自动增行/删行
        ),
    ],
    jt_notes={
        180: "待获取完整工商登记资料后确认。",     # int → 段落索引
        "治理结构": "待获取公司章程后核实。",       # str → 按文字定位段落
    },
    auto_cleanup=True,
    run_lint=True,
)

result = inject_engine.execute(plan, cfg)
print(result.summary())
# 表格: 2 张, 共填充 4 行/格
# JT Note: 注入 2 条
# 清理: 空段落 3 处, 孤儿编号 1 处
# Lint: 13/13 条规则通过
```

---

### 8. `lint` — 格式验证

```python
from lex_docx import lint

results = lint.check("report.docx", config=cfg,
    rules=["jt_note_format", "spelling"],   # 可选：只跑部分规则
    custom_rules={"my_rule": my_fn})

for r in results:
    print("✅" if r.passed else "❌", r.rule, r.detail)
```

**内置 13 条规则：**

| 规则 | 检查内容 | 需要配置 |
|------|---------|---------|
| `jt_note_format` | Note 必须有 B + I + 高亮 | `note_prefix` |
| `jt_note_brackets` | Note 必须有前后包裹标记 | `note_prefix` |
| `no_forbidden_text` | 无草稿残留字（`待填写` / `TBD` 等） | 可选 `forbidden_draft_patterns` |
| `no_old_project_refs` | 无旧项目/客户名称出现 | `entity_names.forbidden` |
| `entity_name_consistency` | 主体名称统一 | `entity_names` |
| `tc_author_check` | TC INS/DEL 作者一致 | `tc_author` |
| `indent_consistency` | 同级标题缩进一致 | — |
| `defined_terms_bold` | 定义术语均已加粗 | — |
| `table_header_format` | 标题行有底色 + 加粗 | `expected_header_shading` |
| `table_borders` | 表格有边框 | — |
| `table_data_not_empty` | 数据行无空白单元格 | — |
| `spelling` | 无 `common_typos` 中的错别字 | `common_typos` |
| `rPr_consistency` | TC INS run 字体/字号与 pStyle 一致 | `style_rPr_map` |

---

## `DocConfig` — 律所/项目级配置

```python
from lex_docx import DocConfig

cfg = DocConfig(
    # TC 作者
    author="JT",

    # Note 格式
    note_prefix="[JT Note: ",
    note_suffix="]",
    note_highlight="yellow",

    # 表格默认格式
    header_shading="D9E2F3",
    border_style="single",
    border_width=4,
    border_color="000000",

    # Lint 配置
    entity_names={
        "allowed":   ["漕河泾总公司", "联合发展公司"],
        "forbidden": ["目标公司", "新元房产"],
    },
    common_typos=["漕河泵", "演河泵"],

    # Style-Aware 注入（FR-6）
    # dict 格式，或由 format_brush.extract_style_rPr_map(doc) 提供的 lxml 元素
    style_rPr_map={
        "Normal":   {"eastAsia": "仿宋_GB2312", "sz": "24"},
        "Heading 2": {"eastAsia": "楷体", "sz": "32"},
        "Heading 3": {"eastAsia": "仿宋_GB2312", "sz": "24"},
    },

    # 其他
    extra_term_patterns=[r'下简称[「""]([^"""」]{1,30})[""」]'],
    custom_lint_rules={"my_rule": my_fn},
)
```

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `author` | `"JT"` | TC INS/DEL 作者 |
| `note_prefix` | `"[JT Note: "` | Note 开头标记 |
| `note_suffix` | `"]"` | Note 结尾标记 |
| `note_highlight` | `"yellow"` | Word 高亮颜色 |
| `header_shading` | `"D9E2F3"` | 表格标题行底色 |
| `border_style` | `"single"` | 表格边框样式 |
| `border_width` | `4` | 边框宽度（1/8 磅） |
| `border_color` | `"000000"` | 边框颜色 |
| `entity_names` | `{}` | `{"allowed": [...], "forbidden": [...]}` |
| `common_typos` | `[]` | 错别字列表 |
| `style_rPr_map` | `{}` | style 字体/字号映射，用于 rPr 继承和 lint |
| `extra_term_patterns` | `[]` | 额外定义术语正则 |
| `custom_lint_rules` | `{}` | 自定义 lint 规则 |

---

## 典型工作流

### 方式 A — 逐步调用

```python
from docx import Document
from lex_docx import DocConfig, format_brush, jt_note, table_ops, cleanup, lint
from lex_docx.format_brush import extract_style_rPr_map

cfg = DocConfig(
    author="JT",
    entity_names={"allowed": ["漕河泾总公司"], "forbidden": ["目标公司"]},
    common_typos=["漕河泵"],
)

doc = Document("report_draft.docx")

# Step 0: 提取 style 映射（用于 rPr 继承）
cfg.style_rPr_map = extract_style_rPr_map(doc)

# Step 1: 填充表格
shareholders = table_ops.extract_table("autodocs.docx", near_text="目前公司的股东情况")
table_ops.fill_table(doc, table_index=12, data=shareholders,
    column_mapping={"股东姓名/名称": "股东名称", "出资额（万元）": "认缴出资额\n（人民币/万元）"},
    auto_add_rows=True, cfg=cfg)
table_ops.format_table(doc, table_index=12, cfg=cfg)

# Step 2: 注入 Note
jt_note.append_to_paragraph(doc, 180, "待获取完整工商登记资料后确认。", cfg=cfg)

# Step 3: 清理空段落
cleanup.cleanup_all(doc, para_range=(175, 200), cfg=cfg)

# Step 4: 修复格式
format_brush.auto_fix(doc, para_range=(175, 200))

# Step 5: Lint 验证
results = lint.check("report_draft.docx", config=cfg)
for r in results:
    if not r.passed:
        print(f"❌ {r.rule}: {r.detail}")

doc.save("report_reviewed.docx")
```

### 方式 B — inject_engine 一键执行

```python
from lex_docx import inject_engine, DocConfig

plan = inject_engine.InjectPlan(
    doc_path="report.docx",
    out_path="report_reviewed.docx",
    tables=[
        inject_engine.TableFill(table_index=8,
            data={"企业名称": "临港资管", "法定代表人/负责人/执行事务合伙人": "张三"},
            mode="kv"),
        inject_engine.TableFill(table_index=12,
            data=[{"股东": "临港集团", "出资额": "1060万", "比例": "100%"}],
            mode="rows", auto_adjust=True),
    ],
    jt_notes={180: "待获取完整工商登记资料后确认。"},
    auto_cleanup=True,
    run_lint=True,
)
result = inject_engine.execute(plan, DocConfig())
print(result.summary())
```

### CLI 用法

```bash
# ── 查询 / 检查 ───────────────────────────────────────────────────────────────
lex_docx lint report.docx --fmt text
lex_docx lint report.docx --rules jt_note_format,spelling --fmt json
lex_docx extract autodocs.docx --table 3 > data.json
lex_docx extract autodocs.docx --near "目前公司的股东情况" > data.json

# ── 表格填充 ──────────────────────────────────────────────────────────────────
lex_docx fill-kv report.docx --table 8 --data kv.json --key-cols 0,2 --out out.docx
lex_docx fill-table report.docx --table 12 --data rows.json --auto-del --out out.docx

# ── 表格格式 ──────────────────────────────────────────────────────────────────
lex_docx format-table report.docx --table 12 --shading D9E2F3 --borders single
lex_docx copy-table autodocs.docx report.docx --src-table 3 --dst-pos after_para:241 --out out.docx
lex_docx copy-table autodocs.docx report.docx --src-table 3 --dst-pos replace_table:15 --max-rows 30 --out out.docx

# ── Track Changes ─────────────────────────────────────────────────────────────
# 在段落 180 末尾插入文字（继承原 run 字体字号）
lex_docx tc-insert report.docx --para 180 --text "（待核实）" --out out.docx
# 插入时指定格式
lex_docx tc-insert report.docx --para 180 --text "注：" --bold --highlight yellow --pos start --out out.docx
# rPr 继承策略：true=继承首run（默认）| style=跟pStyle | auto=按style_rPr_map
lex_docx tc-insert report.docx --para 180 --text "..." --inherit-rpr style --out out.docx

# 将单段落标记为 TC DEL
lex_docx tc-delete report.docx --para 180 --out out.docx
# 批量 TC DEL（含两端）
lex_docx tc-delete report.docx --range 180,195 --out out.docx

# ── 高亮 ──────────────────────────────────────────────────────────────────────
lex_docx highlight report.docx --para 180 --out out.docx
lex_docx highlight report.docx --range 200,210 --color yellow --out out.docx

# ── 格式刷 ────────────────────────────────────────────────────────────────────
# 从段落 171 复制格式到 177、178、180
lex_docx format-brush report.docx --ref 171 --target 177,178,180 --out out.docx
# 复制格式到连续范围
lex_docx format-brush report.docx --ref 171 --range 175,185 --out out.docx
# 只复制部分格式项
lex_docx format-brush report.docx --ref 171 --target 177 --copy indent,spacing --out out.docx

# ── 清理 ──────────────────────────────────────────────────────────────────────
lex_docx cleanup report.docx --range 206,265 --mode tc-del --out out.docx
lex_docx cleanup report.docx --mode report    # 只报告，不修改

# ── 术语 / Lint ───────────────────────────────────────────────────────────────
lex_docx bold-terms report.docx --scan
lex_docx bold-terms report.docx --para 39 --out out.docx

# ── 一键注入（inject_engine）─────────────────────────────────────────────────
lex_docx inject plan.json --cfg jt.json --out report_reviewed.docx
```

**`inject` 的 `plan.json` 格式**：

```json
{
  "doc_path": "report.docx",
  "out_path": "report_reviewed.docx",
  "target_range": [200, 300],
  "tables": [
    {"table_index": 8, "data": {"企业名称": "临港资管"}, "mode": "kv", "key_columns": [0, 2]},
    {"table_index": 12, "data": [{"股东": "临港集团", "出资额": "1060万"}], "mode": "rows", "auto_adjust": true}
  ],
  "jt_notes": {"180": "待获取工商资料后确认", "治理结构": "待核实"},
  "auto_cleanup": true,
  "run_lint": true
}
```

---

## 文件结构

```
lex_docx/
├── __init__.py          # 顶层导出（含 DocConfig, PRESET_JT）
├── config.py            # DocConfig 配置类
├── constants.py         # 内置默认常量
├── tc_utils.py          # TC INS/DEL XML 底层工具
├── format_brush.py      # 格式刷 + extract_style_rPr_map
├── jt_note.py           # 律所批注注入
├── defined_terms.py     # 定义术语加粗（中英文）
├── table_ops.py         # 表格操作
├── cleanup.py           # 空段落 / 孤儿编号清理
├── inject_engine.py     # 批量注入引擎
├── lint.py              # 格式验证（13 条规则）
└── cli.py               # 命令行入口
```

---

## 与 Adeu 的关系

| | Adeu | lex_docx |
|---|---|---|
| 职责 | 文本级 Track Changes（读取/注入/accept/reject） | 格式控制、表格操作、Note 注入、lint 验证 |
| 适合场景 | 文字内容的增删改 | 格式修复、数据填充、审阅标注、批量注入 |
| 互补关系 | 读取文档结构 → | → lex_docx 处理格式与表格 |

两者可以在同一个文档上协同工作，不互相冲突。
