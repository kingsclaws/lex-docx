"""
tc_utils.py — Track Changes XML 底层工具

所有 w:ins / w:del 构造统一从这里发出，避免各模块重复实现。
"""
from __future__ import annotations

from copy import deepcopy
from datetime import datetime, timezone

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --------------------------------------------------------------------------- #
# 内部辅助                                                                      #
# --------------------------------------------------------------------------- #

def _utc_now() -> str:
    """返回 Word 兼容的 ISO 8601 UTC 时间戳，如 2026-03-17T10:30:00Z"""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def next_tc_id(doc) -> int:
    """
    扫描文档中所有 w:ins / w:del / w:comment* 的 w:id，返回 max+1。
    保证不与现有 TC 冲突。
    """
    body = doc.element.body
    max_id = 0
    for tag in (qn("w:ins"), qn("w:del"),
                qn("w:commentRangeStart"), qn("w:commentRangeEnd")):
        for el in body.iter(tag):
            try:
                max_id = max(max_id, int(el.get(qn("w:id"), 0)))
            except (ValueError, TypeError):
                pass
    return max_id + 1


# --------------------------------------------------------------------------- #
# w:rPr 构造                                                                    #
# --------------------------------------------------------------------------- #

def make_rPr(
    bold: bool = False,
    italic: bool = False,
    highlight: str | None = None,
    base_rPr=None,          # 可选：从现有 w:rPr element deepcopy 作为基础
) -> OxmlElement:
    """
    构造 w:rPr 元素。
    - bold=True  → 同时设 w:b + w:bCs（中文加粗必须两者都设）
    - italic=True → w:i + w:iCs
    - highlight  → w:highlight val="yellow"（或其他颜色值）
    - base_rPr   → deepcopy 现有 rPr 后追加/覆盖

    元素顺序遵循 OOXML 规范：b > bCs > i > iCs > highlight
    """
    rPr = deepcopy(base_rPr) if base_rPr is not None else OxmlElement("w:rPr")

    if bold:
        if rPr.find(qn("w:b")) is None:
            b = OxmlElement("w:b")
            rPr.append(b)
        if rPr.find(qn("w:bCs")) is None:
            bCs = OxmlElement("w:bCs")
            rPr.append(bCs)

    if italic:
        if rPr.find(qn("w:i")) is None:
            i = OxmlElement("w:i")
            rPr.append(i)
        if rPr.find(qn("w:iCs")) is None:
            iCs = OxmlElement("w:iCs")
            rPr.append(iCs)

    if highlight:
        existing_hl = rPr.find(qn("w:highlight"))
        if existing_hl is not None:
            rPr.remove(existing_hl)
        hl = OxmlElement("w:highlight")
        hl.set(qn("w:val"), highlight)
        rPr.append(hl)

    return rPr


def make_rPr_from_dict(d: dict) -> OxmlElement:
    """
    从配置字典构造 w:rPr。

    支持键（均为字符串值）：
      ascii, eastAsia, hAnsi, cs  → w:rFonts 属性（字体）
      sz, szCs                    → 字号（单位 half-points，如 "24" = 12pt）
      b, i                        → True/False，加粗/斜体
      color                       → 十六进制颜色，如 "FF0000"
      highlight                   → Word highlight 颜色名，如 "yellow"

    示例：
      make_rPr_from_dict({"eastAsia": "仿宋_GB2312", "sz": "24"})
    """
    rPr = OxmlElement("w:rPr")

    # 字体
    font_attrs = {k: d[k] for k in ("ascii", "eastAsia", "hAnsi", "cs") if k in d}
    if font_attrs:
        rFonts = OxmlElement("w:rFonts")
        _attr_map = {
            "ascii": qn("w:ascii"),
            "eastAsia": qn("w:eastAsia"),
            "hAnsi": qn("w:hAnsi"),
            "cs": qn("w:cs"),
        }
        for k, v in font_attrs.items():
            rFonts.set(_attr_map[k], v)
        rPr.append(rFonts)

    # 加粗
    if d.get("b"):
        rPr.append(OxmlElement("w:b"))
        rPr.append(OxmlElement("w:bCs"))

    # 斜体
    if d.get("i"):
        rPr.append(OxmlElement("w:i"))
        rPr.append(OxmlElement("w:iCs"))

    # 字号
    if "sz" in d:
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(d["sz"]))
        rPr.append(sz)
    if "szCs" in d:
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(d["szCs"]))
        rPr.append(szCs)

    # 颜色
    if "color" in d:
        color_el = OxmlElement("w:color")
        color_el.set(qn("w:val"), d["color"].lstrip("#").upper())
        rPr.append(color_el)

    # 高亮
    if "highlight" in d:
        hl = OxmlElement("w:highlight")
        hl.set(qn("w:val"), d["highlight"])
        rPr.append(hl)

    return rPr


def _resolve_rPr(
    para_el,
    inherit_rPr,
    style_rPr_map: dict | None,
) -> OxmlElement | None:
    """
    根据 inherit_rPr 策略解析出 base_rPr。

    inherit_rPr 取值：
      False / None  → 不继承（返回 None）
      True          → 从段落中第一个 w:del > w:r 的 rPr 复制
      "style"       → 纯样式继承，不设 rPr（返回 _STYLE_SENTINEL）
      "auto"        → 从段落 pStyle 查 style_rPr_map（需传入 style_rPr_map）
      Paragraph     → 从 python-docx Paragraph 的第一个 run 取 rPr
      Run           → 从 python-docx Run 取 rPr
      lxml element  → 直接 deepcopy 该 rPr 元素
    """
    if not inherit_rPr:
        return None

    if inherit_rPr is _STYLE_SENTINEL or inherit_rPr == "style":
        return _STYLE_SENTINEL   # 哨兵：调用方不设 rPr

    if inherit_rPr is True:
        # 从段落中已有 w:del 的第一个 run 取 rPr
        for del_el in para_el.iter(qn("w:del")):
            for r_el in del_el.findall(qn("w:r")):
                rPr = r_el.find(qn("w:rPr"))
                if rPr is not None:
                    return deepcopy(rPr)
        return None   # 找不到 → 不继承

    if inherit_rPr == "auto":
        # 从 pStyle 查 style_rPr_map
        if not style_rPr_map:
            return None
        pPr = para_el.find(qn("w:pPr"))
        pStyle = pPr.find(qn("w:pStyle")) if pPr is not None else None
        style_name = pStyle.get(qn("w:val"), "") if pStyle is not None else "Normal"
        entry = style_rPr_map.get(style_name)
        if entry is None:
            return None
        if isinstance(entry, dict):
            return make_rPr_from_dict(entry)
        return deepcopy(entry)   # lxml element

    # python-docx Paragraph
    if hasattr(inherit_rPr, "paragraphs"):  # Document — not supported
        return None
    if hasattr(inherit_rPr, "runs"):        # Paragraph
        para_obj = inherit_rPr
        for run in para_obj.runs:
            rPr = run._element.find(qn("w:rPr"))
            if rPr is not None:
                return deepcopy(rPr)
        return None
    if hasattr(inherit_rPr, "_element") and hasattr(inherit_rPr, "text"):  # Run
        rPr = inherit_rPr._element.find(qn("w:rPr"))
        return deepcopy(rPr) if rPr is not None else None

    # lxml element（直接传入 rPr）
    if hasattr(inherit_rPr, "tag"):
        return deepcopy(inherit_rPr)

    return None


# 哨兵对象：表示"不加 rPr，让 Word 从 style 继承"
_STYLE_SENTINEL = object()


# --------------------------------------------------------------------------- #
# w:r 构造                                                                      #
# --------------------------------------------------------------------------- #

def make_run(
    text: str,
    bold: bool = False,
    italic: bool = False,
    highlight: str | None = None,
    base_rPr=None,
) -> OxmlElement:
    """
    构造 w:r 元素。
    空文本返回空 run（调用方自行决定是否跳过）。
    """
    r = OxmlElement("w:r")

    # 只有需要格式时才加 rPr
    if bold or italic or highlight or base_rPr is not None:
        rPr = make_rPr(bold=bold, italic=italic, highlight=highlight,
                       base_rPr=base_rPr)
        r.append(rPr)

    t = OxmlElement("w:t")
    t.text = text or ""
    # 头尾有空格时需要 xml:space="preserve"，否则 Word 会吞掉空格
    if text and (text[0] == " " or text[-1] == " "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)

    return r


# --------------------------------------------------------------------------- #
# w:ins / w:del tag 构造                                                        #
# --------------------------------------------------------------------------- #

def make_tc_tag(
    tag_name: str,      # "w:ins" 或 "w:del"
    tc_id: int,
    author: str,
    date: str | None = None,
) -> OxmlElement:
    """
    构造 w:ins 或 w:del 标签（不含子元素）。
    date 默认为当前 UTC 时间。
    """
    el = OxmlElement(tag_name)
    el.set(qn("w:id"), str(tc_id))
    el.set(qn("w:author"), author)
    el.set(qn("w:date"), date or _utc_now())
    return el


def make_ins_run(
    text: str,
    tc_id: int,
    author: str,
    bold: bool = False,
    italic: bool = False,
    highlight: str | None = None,
    date: str | None = None,
    base_rPr=None,
) -> OxmlElement:
    """
    构造完整的 w:ins > w:r > w:t 结构（文本级 Track Changes）。
    """
    ins = make_tc_tag("w:ins", tc_id, author, date)
    r = make_run(text, bold=bold, italic=italic, highlight=highlight,
                 base_rPr=base_rPr)
    ins.append(r)
    return ins


# --------------------------------------------------------------------------- #
# 表格行级 TC（w:trPr > w:ins / w:del）                                        #
# --------------------------------------------------------------------------- #

def mark_row_as_inserted(tr_element, tc_id: int, author: str, date: str | None = None):
    """
    将 w:tr 标记为 TC INS（行级插入）。
    正确方式：在 w:trPr 内添加 w:ins，而非用 w:ins 包裹整个 w:tr。
    """
    trPr = tr_element.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr_element.insert(0, trPr)
    ins = make_tc_tag("w:ins", tc_id, author, date)
    trPr.append(ins)


def mark_row_as_deleted(tr_element, tc_id: int, author: str, date: str | None = None):
    """将 w:tr 标记为 TC DEL（行级删除）。"""
    trPr = tr_element.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr_element.insert(0, trPr)
    del_el = make_tc_tag("w:del", tc_id, author, date)
    trPr.append(del_el)


# --------------------------------------------------------------------------- #
# 段落级 TC                                                                     #
# --------------------------------------------------------------------------- #

def tc_del_paragraph(
    para_el,
    tc_id: int,
    author: str,
    date: str | None = None,
) -> OxmlElement | None:
    """
    将段落标记为 TC DEL（段落级删除）。

    操作：
    1. 每个 w:r 的 w:t 改为 w:delText
    2. 每个 w:r 用 w:del 包裹
    3. w:pPr > w:rPr > w:del 标记段落结束符被删除

    para_el: python-docx Paragraph 对象 或 w:p lxml element

    Returns:
        删除前第一个 run 的 rPr（deepcopy），供 tc_ins_text(inherit_rPr=True) 使用。
        若段落无 run 则返回 None。
    """
    if hasattr(para_el, "_element"):
        para_el = para_el._element

    _date = date or _utc_now()

    # 预先保存第一个 run 的 rPr（在修改前）
    first_rPr: OxmlElement | None = None
    first_r = para_el.find(qn("w:r"))
    if first_r is not None:
        rPr_el = first_r.find(qn("w:rPr"))
        if rPr_el is not None:
            first_rPr = deepcopy(rPr_el)

    # ── 1. 将所有 w:r 包裹进 w:del，w:t → w:delText ── #
    for r_el in list(para_el.findall(qn("w:r"))):
        for t_el in r_el.findall(qn("w:t")):
            t_el.tag = qn("w:delText")
        del_wrap = make_tc_tag("w:del", tc_id, author, _date)
        idx = list(para_el).index(r_el)
        para_el.remove(r_el)
        del_wrap.append(r_el)
        para_el.insert(idx, del_wrap)
        tc_id += 1   # 每个 run 用独立 id（Word 规范要求 del 内每个 run 唯一）

    # ── 2. 标记段落结束符删除（¶ mark deleted） ── #
    pPr = para_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        para_el.insert(0, pPr)
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    if rPr.find(qn("w:del")) is None:
        del_mark = make_tc_tag("w:del", tc_id, author, _date)
        rPr.append(del_mark)

    return first_rPr


def tc_ins_text(
    para_el,
    text: str,
    tc_id: int,
    author: str,
    position: str | int = "end",
    bold: bool = False,
    italic: bool = False,
    highlight: str | None = None,
    date: str | None = None,
    base_rPr=None,
    inherit_rPr=False,
    style_rPr_map: dict | None = None,
) -> OxmlElement:
    """
    在段落内以 TC INS 形式插入文字。

    Args:
        para_el:        python-docx Paragraph 或 w:p lxml element
        text:           插入文字
        tc_id:          Track Changes ID
        author:         作者
        position:       "end"（末尾）| "start"（开头）| 整数（第 n 个 run 之后）
        bold/italic/highlight: 额外格式（叠加在继承的 rPr 之上）
        base_rPr:       显式传入 rPr element（优先级最高）
        inherit_rPr:    rPr 继承策略（见下）
        style_rPr_map:  {style_name: rPr_el_or_dict}，配合 inherit_rPr="auto" 使用

    inherit_rPr 取值：
        False       → 不继承（默认）
        True        → 从段落中第一个 w:del > w:r 复制 rPr（配合 tc_del_paragraph 使用）
        "style"     → 不设 rPr，让 Word 从 pStyle 继承字体字号
        "auto"      → 从 style_rPr_map 按段落 pStyle 查找 rPr
        Paragraph   → 从指定 python-docx Paragraph 的第一个 run 复制 rPr
        Run         → 从指定 python-docx Run 复制 rPr
        lxml el     → 直接 deepcopy 该 rPr element

    Returns:
        插入的 w:ins element
    """
    if hasattr(para_el, "_element"):
        para_el = para_el._element

    # 解析 base_rPr
    if base_rPr is None and inherit_rPr is not False:
        resolved = _resolve_rPr(para_el, inherit_rPr, style_rPr_map)
        if resolved is _STYLE_SENTINEL:
            # 纯样式继承：构造无 rPr 的 run
            ins = make_tc_tag("w:ins", tc_id, author, date)
            r = OxmlElement("w:r")
            t = OxmlElement("w:t")
            t.text = text or ""
            if text and (text[0] == " " or text[-1] == " "):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.append(t)
            ins.append(r)
            _insert_at(para_el, ins, position)
            return ins
        base_rPr = resolved

    ins = make_ins_run(text, tc_id, author,
                       bold=bold, italic=italic, highlight=highlight,
                       date=date, base_rPr=base_rPr)
    _insert_at(para_el, ins, position)
    return ins


def _insert_at(para_el, ins_el, position) -> None:
    """将 ins_el 插入 para_el 的指定位置。"""
    runs = para_el.findall(qn("w:r"))

    if position == "end" or not runs:
        para_el.append(ins_el)
    elif position == "start":
        first_content = (para_el.findall(qn("w:r")) or
                         para_el.findall(qn("w:ins")))
        if first_content:
            idx = list(para_el).index(first_content[0])
            para_el.insert(idx, ins_el)
        else:
            para_el.append(ins_el)
    elif isinstance(position, int):
        all_runs = para_el.findall(qn("w:r"))
        if position < len(all_runs):
            ref = all_runs[position]
            idx = list(para_el).index(ref) + 1
            para_el.insert(idx, ins_el)
        else:
            para_el.append(ins_el)
    else:
        para_el.append(ins_el)


def tc_ins_mixed(
    para_el,
    segments: list[tuple[str, bool]],
    tc_id: int,
    author: str,
    cfg=None,
    inherit_rPr=False,
    style_rPr_map: dict | None = None,
    date: str | None = None,
) -> list[OxmlElement]:
    """
    在段落末尾以 TC INS 形式插入混合内容：普通文字 + 律所 Note。

    Args:
        segments:  [(text, is_note), ...]
                   is_note=True → 自动叠加 B+I+HL，并加 note_prefix/suffix
        cfg:       DocConfig（提供 note_prefix/suffix/highlight）
        inherit_rPr: 普通文字 run 的 rPr 继承策略（同 tc_ins_text）

    Returns:
        插入的 w:ins elements 列表
    """
    if hasattr(para_el, "_element"):
        para_el = para_el._element

    note_prefix  = (getattr(cfg, "note_prefix",  "[JT Note: ") if cfg else "[JT Note: ")
    note_suffix  = (getattr(cfg, "note_suffix",  "]")          if cfg else "]")
    note_hl      = (getattr(cfg, "note_highlight","yellow")     if cfg else "yellow")
    _date        = date or _utc_now()

    # 解析 base_rPr（普通文字用）
    base_rPr = None
    if inherit_rPr is not False:
        resolved = _resolve_rPr(para_el, inherit_rPr, style_rPr_map)
        if resolved is not _STYLE_SENTINEL:
            base_rPr = resolved

    inserted = []
    for text, is_note in segments:
        if not text:
            continue
        if is_note:
            full_text = f"{note_prefix}{text}{note_suffix}"
            ins = make_ins_run(full_text, tc_id, author,
                               bold=True, italic=True, highlight=note_hl,
                               date=_date, base_rPr=base_rPr)
        else:
            ins = make_ins_run(text, tc_id, author,
                               date=_date, base_rPr=base_rPr)
        para_el.append(ins)
        inserted.append(ins)
        tc_id += 1

    return inserted
