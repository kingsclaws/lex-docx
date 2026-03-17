"""
config.py — lex_docx 项目/律所级配置

使用方式：
    from lex_docx import DocConfig

    # 竞天公诚（默认）
    cfg = DocConfig()

    # 其他律所（举例）
    cfg = DocConfig(
        author="XX",
        note_prefix="[XX Note: ",
        note_suffix="]",
        note_highlight="cyan",
        header_shading="F2F2F2",
    )

    # 传入各函数
    jt_note.append_to_paragraph(doc, 180, "待确认...", cfg=cfg)
    lint.check("report.docx", config=cfg)
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable


@dataclass
class DocConfig:
    """
    律所 / 项目级配置，集中管理所有可定制参数。

    所有函数均可接受 cfg=DocConfig(...) 来覆盖默认值；
    不传则退回 constants.py 中的内置默认值（向后兼容）。
    """

    # ── TC 作者 ──────────────────────────────────────────────────────────────
    author: str = "JT"

    # ── 批注注释格式（律所 Note）──────────────────────────────────────────────
    note_prefix: str = "[JT Note: "    # 开头标记，含律所名
    note_suffix: str = "]"             # 结尾标记
    note_highlight: str = "yellow"     # Word highlight 颜色值

    # ── 表格默认格式 ─────────────────────────────────────────────────────────
    header_shading: str = "D9E2F3"    # 标题行底色（十六进制，不含 #）
    header_row_index: int = 0          # 标题行在表格中的位置（0=首行）
    border_style: str = "single"
    border_width: int = 4              # 单位：eighths of a point
    border_color: str = "000000"

    # ── Lint 规则配置 ─────────────────────────────────────────────────────────
    # entity_names: {"allowed": [...], "forbidden": [...]}
    entity_names: dict[str, list[str]] = field(default_factory=dict)
    # common_typos: 常见错别字列表
    common_typos: list[str] = field(default_factory=list)
    # forbidden_draft_patterns: 草稿禁用文字（None = 用 constants 内置列表）
    forbidden_draft_patterns: list[str] | None = None
    # 自定义 lint 规则：{rule_name: fn(doc, config_dict, check_range) -> LintResult}
    custom_lint_rules: dict[str, Callable] = field(default_factory=dict)

    # ── 定义术语检测 ──────────────────────────────────────────────────────────
    # 额外正则模式（字符串），每个 pattern 应有一个捕获组对应术语文字
    extra_term_patterns: list[str] = field(default_factory=list)

    # ── Style-Aware 注入（FR-6）──────────────────────────────────────────────
    # style_rPr_map: {style_name: dict_or_rPr_element}
    # dict 格式: {"eastAsia": "仿宋_GB2312", "sz": "24"}
    # 或由 format_brush.extract_style_rPr_map(doc) 返回的 lxml element
    style_rPr_map: dict = field(default_factory=dict)

    # ── Lint tc_author 独立设置（None = 同 author）────────────────────────────
    tc_author: str | None = None

    def __post_init__(self):
        if self.tc_author is None:
            self.tc_author = self.author
        if self.forbidden_draft_patterns is None:
            from .constants import FORBIDDEN_DRAFT_PATTERNS
            self.forbidden_draft_patterns = FORBIDDEN_DRAFT_PATTERNS

    def to_lint_config(self, extra: dict | None = None) -> dict:
        """
        将 DocConfig 转为 lint.check() 的 config dict 格式。
        可传入 extra 字典覆盖特定键（如 check_range）。
        """
        d = {
            "tc_author": self.tc_author,
            "entity_names": self.entity_names,
            "common_typos": self.common_typos,
            "forbidden_draft_patterns": self.forbidden_draft_patterns,
            "note_prefix": self.note_prefix,
            "note_name": self.note_prefix.strip("[ :"),   # e.g. "JT Note"
            "expected_header_shading": self.header_shading,
            "custom_rules": self.custom_lint_rules,
            "style_rPr_map": self.style_rPr_map,
        }
        if extra:
            d.update(extra)
        return d

    @property
    def note_name(self) -> str:
        """从 note_prefix 提取可读名称，如 '[JT Note: ' → 'JT Note'。"""
        return self.note_prefix.strip(" [:（[")


# 内置预设配置
PRESET_JT = DocConfig(
    author="JT",
    note_prefix="[JT Note: ",
    note_suffix="]",
    note_highlight="yellow",
    header_shading="D9E2F3",
)
