"""
para_query.py — 全文格式检索

按段落样式、字体、字号、大纲级别、粗体、斜体等属性过滤段落。
字体检测走完整继承链：run 显式值 → 段落样式定义 → 文档默认样式。

用法：
    from lex_docx import para_query
    results = para_query.query(doc, font="仿宋")
    results = para_query.query(doc, outline_level=[1, 2])
    results = para_query.query(doc, style=["Heading 1", "Heading 2"])
    results = para_query.query(doc, font="仿宋", font_size=12.0)
"""
from __future__ import annotations

from docx.oxml.ns import qn


# --------------------------------------------------------------------------- #
# 样式继承辅助                                                                  #
# --------------------------------------------------------------------------- #

def _build_style_map(doc) -> dict[str, dict]:
    """
    从 styles.xml 构建 {style_name: {font_eastasia, font_ascii, font_size, bold, italic, color}} 映射。
    包含 basedOn 链的继承合并（父 → 子方向覆盖）。
    """
    raw: dict[str, dict] = {}   # style_name -> raw attrs (不含继承)
    based_on: dict[str, str] = {}  # style_name -> parent_name

    for style in doc.styles:
        name = style.name
        el = style.element
        parent_el = el.find(qn("w:basedOn"))
        if parent_el is not None:
            based_on[name] = parent_el.get(qn("w:val"), "")

        attrs: dict = {}
        rPr = el.find(qn("w:rPr"))
        if rPr is not None:
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is not None:
                attrs["font_eastasia"] = rFonts.get(qn("w:eastAsia"))
                attrs["font_ascii"]    = rFonts.get(qn("w:ascii"))
                attrs["font_hAnsi"]   = rFonts.get(qn("w:hAnsi"))
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                try:
                    attrs["font_size"] = int(sz.get(qn("w:val"), 0)) / 2
                except (ValueError, TypeError):
                    pass
            if rPr.find(qn("w:b")) is not None:
                attrs["bold"] = True
            if rPr.find(qn("w:i")) is not None:
                attrs["italic"] = True
            color_el = rPr.find(qn("w:color"))
            if color_el is not None:
                attrs["color"] = color_el.get(qn("w:val"))
        raw[name] = attrs

    # 解析继承链（最多 10 层防循环）
    def _resolve(name: str, visited: set | None = None) -> dict:
        if visited is None:
            visited = set()
        if name in visited:
            return dict(raw.get(name, {}))
        visited.add(name)
        own = dict(raw.get(name, {}))
        parent_name = based_on.get(name)
        if not parent_name:
            return own
        parent = _resolve(parent_name, visited)
        merged = {**parent, **{k: v for k, v in own.items() if v is not None}}
        return merged

    return {name: _resolve(name) for name in raw}


def _get_run_rpr_val(run_el, tag: str) -> str | None:
    rPr = run_el.find(qn("w:rPr"))
    if rPr is None:
        return None
    el = rPr.find(qn(tag))
    if el is None:
        return None
    return el.get(qn("w:val"))


def _run_has_flag(run_el, tag: str) -> bool | None:
    """检查 run 是否有指定开关标记（b/i/strike 等）。返回 None 表示未显式设置。"""
    rPr = run_el.find(qn("w:rPr"))
    if rPr is None:
        return None
    el = rPr.find(qn(tag))
    if el is None:
        return None
    val = el.get(qn("w:val"))
    # <w:b/> 或 <w:b w:val="true"> → True；<w:b w:val="false"> → False
    if val is None or val.lower() in ("true", "1", "on"):
        return True
    return False


def _run_font(run_el) -> dict:
    """提取 run 的显式字体设置。"""
    rPr = run_el.find(qn("w:rPr"))
    if rPr is None:
        return {}
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        return {}
    return {
        "font_eastasia": rFonts.get(qn("w:eastAsia")),
        "font_ascii":    rFonts.get(qn("w:ascii")),
        "font_hAnsi":   rFonts.get(qn("w:hAnsi")),
    }


def _run_font_size(run_el) -> float | None:
    rPr = run_el.find(qn("w:rPr"))
    if rPr is None:
        return None
    sz = rPr.find(qn("w:sz"))
    if sz is None:
        return None
    try:
        return int(sz.get(qn("w:val"), 0)) / 2
    except (ValueError, TypeError):
        return None


# --------------------------------------------------------------------------- #
# 段落属性提取                                                                  #
# --------------------------------------------------------------------------- #

def _para_outline_level(para) -> int | None:
    """读取 w:outlineLvl，转换为用户级别 1-9，None 表示未设置（正文）。"""
    pPr = para._element.find(qn("w:pPr"))
    if pPr is None:
        return None
    ol = pPr.find(qn("w:outlineLvl"))
    if ol is None:
        return None
    try:
        val = int(ol.get(qn("w:val"), 9))
        return None if val >= 9 else val + 1
    except (ValueError, TypeError):
        return None


def _para_style_name(para) -> str:
    return para.style.name if para.style else "Normal"


def _collect_para_attrs(para, style_map: dict) -> dict:
    """
    收集段落的有效属性（含继承链）。
    返回 {style, outline_level, font_eastasia, font_ascii, font_size, bold, italic, color}。
    字体/字号/粗斜体优先取段落内 run 的显式值，再回落到样式继承值。
    """
    style_name = _para_style_name(para)
    outline_level = _para_outline_level(para)

    # 收集所有 run 的显式属性
    run_fonts_ea: set[str] = set()
    run_fonts_ascii: set[str] = set()
    run_sizes: set[float] = set()
    run_bold: set[bool] = set()
    run_italic: set[bool] = set()
    run_colors: set[str] = set()

    for run_el in para._element.findall(qn("w:r")):
        rf = _run_font(run_el)
        if rf.get("font_eastasia"):
            run_fonts_ea.add(rf["font_eastasia"])
        if rf.get("font_ascii"):
            run_fonts_ascii.add(rf["font_ascii"])
        sz = _run_font_size(run_el)
        if sz:
            run_sizes.add(sz)
        b = _run_has_flag(run_el, "w:b")
        if b is not None:
            run_bold.add(b)
        i = _run_has_flag(run_el, "w:i")
        if i is not None:
            run_italic.add(i)
        c = _get_run_rpr_val(run_el, "w:color")
        if c and c.lower() != "auto":
            run_colors.add(c.upper())

    # 从样式继承链取默认值
    style_attrs = style_map.get(style_name, {})

    return {
        "style":          style_name,
        "outline_level":  outline_level,
        # 字体：run 显式值（可能多个）+ 样式继承值
        "font_eastasia":  sorted(run_fonts_ea) or ([style_attrs.get("font_eastasia")] if style_attrs.get("font_eastasia") else []),
        "font_ascii":     sorted(run_fonts_ascii) or ([style_attrs.get("font_ascii")] if style_attrs.get("font_ascii") else []),
        "font_size":      sorted(run_sizes) or ([style_attrs.get("font_size")] if style_attrs.get("font_size") else []),
        "bold":           (True in run_bold) or style_attrs.get("bold", False),
        "italic":         (True in run_italic) or style_attrs.get("italic", False),
        "colors":         sorted(run_colors),
    }


# --------------------------------------------------------------------------- #
# 主查询接口                                                                    #
# --------------------------------------------------------------------------- #

def query(
    doc,
    *,
    style: list[str] | None = None,
    font: str | None = None,
    font_size: float | None = None,
    outline_level: list[int] | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: str | None = None,
    para_range: tuple[int, int] | None = None,
    text_preview_len: int = 60,
) -> list[dict]:
    """
    全文格式检索，返回满足所有指定条件（AND）的段落列表。

    Args:
        doc:             python-docx Document
        style:           段落样式名列表（OR 匹配），如 ["Heading 1", "标题1"]
        font:            字体名（部分匹配，不区分大小写），如 "仿宋"
        font_size:       字号（pt），精确匹配
        outline_level:   大纲级别列表（1-9，OR 匹配）
        bold:            True=只要有粗体 run，False=无粗体
        italic:          True=只要有斜体 run，False=无斜体
        color:           字体颜色十六进制，如 "FF0000"（不区分大小写）
        para_range:      扫描范围 (start, end)，默认全文
        text_preview_len: 文本预览截断长度

    Returns:
        list of {
            "index": int,
            "style": str,
            "outline_level": int | None,
            "font_eastasia": list[str],
            "font_ascii": list[str],
            "font_size": list[float],
            "bold": bool,
            "italic": bool,
            "colors": list[str],
            "text": str,
            "matched_on": list[str],
        }
    """
    style_map = _build_style_map(doc)
    paras = doc.paragraphs
    start, end = para_range if para_range else (0, len(paras))
    font_lower = font.lower() if font else None
    color_upper = color.upper() if color else None

    results = []
    for i in range(start, min(end, len(paras))):
        para = paras[i]
        attrs = _collect_para_attrs(para, style_map)
        matched_on: list[str] = []

        # -- style filter
        if style is not None:
            if attrs["style"] not in style:
                continue
            matched_on.append("style")

        # -- outline_level filter
        if outline_level is not None:
            if attrs["outline_level"] not in outline_level:
                continue
            matched_on.append("outline_level")

        # -- font filter（东亚字体 + ascii 字体都检查，部分匹配）
        if font_lower is not None:
            all_fonts = [f.lower() for f in (attrs["font_eastasia"] + attrs["font_ascii"]) if f]
            if not any(font_lower in f for f in all_fonts):
                continue
            matched_on.append("font")

        # -- font_size filter
        if font_size is not None:
            if not attrs["font_size"] or not any(abs(s - font_size) < 0.1 for s in attrs["font_size"]):
                continue
            matched_on.append("font_size")

        # -- bold filter
        if bold is not None:
            if attrs["bold"] != bold:
                continue
            matched_on.append("bold")

        # -- italic filter
        if italic is not None:
            if attrs["italic"] != italic:
                continue
            matched_on.append("italic")

        # -- color filter
        if color_upper is not None:
            if color_upper not in attrs["colors"]:
                continue
            matched_on.append("color")

        text = para.text
        results.append({
            "index":         i,
            "style":         attrs["style"],
            "outline_level": attrs["outline_level"],
            "font_eastasia": attrs["font_eastasia"],
            "font_ascii":    attrs["font_ascii"],
            "font_size":     attrs["font_size"],
            "bold":          attrs["bold"],
            "italic":        attrs["italic"],
            "colors":        attrs["colors"],
            "text":          text[:text_preview_len] + ("…" if len(text) > text_preview_len else ""),
            "matched_on":    matched_on,
        })

    return results
