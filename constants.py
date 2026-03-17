"""lex_docx — 常量定义"""

# Track Changes 默认作者
JT_AUTHOR = "JT"

# JT Note 格式
HIGHLIGHT_YELLOW = "yellow"
JT_NOTE_PREFIX = "[JT Note: "
JT_NOTE_SUFFIX = "]"

# 表格标题行底色（十六进制，不含 #）
HEADER_SHADING_HEX = "D9E2F3"

# 表格边框默认值
DEFAULT_BORDER_COLOR = "000000"
DEFAULT_BORDER_WIDTH = 4   # 单位：eighths of a point（4 = 半磅）
DEFAULT_BORDER_STYLE = "single"

# 常见草稿禁用文字（no_forbidden_text 规则）
FORBIDDEN_DRAFT_PATTERNS = [
    "（删除）",
    "（已删除）",
    "（待填写）",
    "（待补充）",
    "（TBD）",
    "（tbd）",
    "TODO",
    "FIXME",
    "（此处）",
    "XXXXX",
]
