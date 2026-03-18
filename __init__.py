"""
lex_docx — DOCX 自动化工具库（中英文合同 / 尽调报告通用）

模块：
  format_brush   — 格式刷（段落缩进/间距/样式复制）
  jt_note        — 律所批注注入（B+I+HL，支持任意律所 prefix/suffix）
  defined_terms  — 首次定义术语加粗（中英文均支持）
  table_ops      — 表格操作（提取/填充/增删行/格式统一）
  lint           — 格式验证（12 条规则，支持自定义规则）
  tc_utils       — Track Changes XML 底层工具
  constants      — 内置默认常量
  config         — DocConfig 律所/项目级配置

与 Adeu 的关系：
  Adeu 负责文本级 Track Changes（read / apply edits / accept）；
  lex_docx 负责格式控制、表格操作和 Note 注入。

典型工作流：
  from lex_docx import DocConfig
  cfg = DocConfig(author="JT", note_prefix="[JT Note: ")

  1. adeu.read_docx()                     了解文档结构
  2. lex_docx.table_ops.fill_table()    填充报告表格
  3. lex_docx.jt_note.append_to_paragraph(cfg=cfg)
  4. lex_docx.format_brush.apply()      修复格式
  5. lex_docx.lint.check(cfg=cfg)       验证全文
"""

from . import format_brush   # noqa: F401
from . import jt_note        # noqa: F401
from . import lint            # noqa: F401
from . import table_ops      # noqa: F401
from . import defined_terms  # noqa: F401
from . import tc_utils       # noqa: F401
from . import constants      # noqa: F401
from . import config         # noqa: F401
from . import cleanup        # noqa: F401
from . import inject_engine  # noqa: F401

from .config import DocConfig, PRESET_JT   # noqa: F401

__version__ = "0.4.0"
__all__ = [
    # modules
    "format_brush",
    "jt_note",
    "lint",
    "table_ops",
    "defined_terms",
    "tc_utils",
    "constants",
    "config",
    "cleanup",
    "inject_engine",
    # top-level convenience
    "DocConfig",
    "PRESET_JT",
]
