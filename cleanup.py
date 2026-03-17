"""
cleanup.py — 注入后文档清理

remove_empty_paragraphs  — 扫描并以 TC DEL 标记（或直接删除）空段落
remove_orphan_numbering  — 清理有列表编号但无内容的段落
"""
from __future__ import annotations

from docx.oxml.ns import qn

from .constants import JT_AUTHOR
from .tc_utils import next_tc_id, tc_del_paragraph, _utc_now


def _get_cfg_attr(cfg, attr: str, fallback):
    if cfg is None:
        return fallback
    return getattr(cfg, attr, fallback)


def _para_text(para_el) -> str:
    """提取段落全文（含 w:ins 内文字，忽略 w:del 内文字）。"""
    text = []
    for el in para_el.iter():
        if el.tag == qn("w:delText"):
            continue
        if el.tag == qn("w:t") and el.text:
            text.append(el.text)
    return "".join(text).strip()


def _has_numbering(para_el) -> bool:
    """段落是否带有列表编号（w:pPr > w:numPr）。"""
    pPr = para_el.find(qn("w:pPr"))
    if pPr is None:
        return False
    return pPr.find(qn("w:numPr")) is not None


def _already_deleted(para_el) -> bool:
    """段落是否已被 TC DEL 标记（w:pPr > w:rPr > w:del 存在）。"""
    pPr = para_el.find(qn("w:pPr"))
    if pPr is None:
        return False
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        return False
    return rPr.find(qn("w:del")) is not None


def _in_range(idx: int, para_range) -> bool:
    if para_range is None:
        return True
    start, end = para_range
    return start <= idx < end


# =========================================================================== #
# remove_empty_paragraphs                                                      #
# =========================================================================== #

def remove_empty_paragraphs(
    doc,
    as_tc_del: bool = True,
    author: str | None = None,
    para_range: tuple[int, int] | None = None,
    cfg=None,
    keep_styles: list[str] | None = None,
) -> list[int]:
    """
    扫描文档，找出空段落并标记删除或直接移除。

    空段落定义：段落文本为空（去掉 TC DEL 内文字后），且不含图片/对象/域代码。

    Args:
        as_tc_del:    True = TC DEL 标记（保留修订历史，推荐）；
                      False = 直接从 XML 移除（不可撤销）
        author:       TC DEL 作者
        para_range:   只处理该段落索引范围 (start, end)，None = 全文
        keep_styles:  保留这些 style 的空段落不处理（如 ["Heading 1"]）
        cfg:          DocConfig

    Returns:
        被处理的段落索引列表
    """
    author = author or _get_cfg_attr(cfg, "author", JT_AUTHOR)
    keep_styles = set(keep_styles or [])
    tc_id = next_tc_id(doc)
    date = _utc_now()
    affected: list[int] = []

    paragraphs = doc.paragraphs
    body = doc.element.body

    for idx, para in enumerate(paragraphs):
        if not _in_range(idx, para_range):
            continue
        para_el = para._element
        if _already_deleted(para_el):
            continue
        # 跳过受保护的 style
        if para.style and para.style.name in keep_styles:
            continue
        # 跳过含图片/嵌入对象的段落
        if para_el.find(qn("w:drawing")) is not None:
            continue
        if para_el.find(qn("w:object")) is not None:
            continue
        if _para_text(para_el):
            continue

        if as_tc_del:
            tc_del_paragraph(para_el, tc_id, author, date)
            tc_id += 2   # tc_del_paragraph 内部可能消耗多个 id
        else:
            body.remove(para_el)

        affected.append(idx)

    return affected


# =========================================================================== #
# remove_orphan_numbering                                                      #
# =========================================================================== #

def remove_orphan_numbering(
    doc,
    as_tc_del: bool = True,
    author: str | None = None,
    para_range: tuple[int, int] | None = None,
    cfg=None,
) -> list[int]:
    """
    清理有列表编号（w:numPr）但文本为空的孤儿段落。

    这类段落通常是章节裁剪后残留的列表项占位符，
    在 Word 里显示为一个孤立的编号点（"1."、"(一)"等）后面没有文字。

    Args:
        as_tc_del:  True = TC DEL 标记；False = 直接移除
        author:     TC DEL 作者
        para_range: 只处理该索引范围
        cfg:        DocConfig

    Returns:
        被处理的段落索引列表
    """
    author = author or _get_cfg_attr(cfg, "author", JT_AUTHOR)
    tc_id = next_tc_id(doc)
    date = _utc_now()
    affected: list[int] = []

    paragraphs = doc.paragraphs
    body = doc.element.body

    for idx, para in enumerate(paragraphs):
        if not _in_range(idx, para_range):
            continue
        para_el = para._element
        if _already_deleted(para_el):
            continue
        if not _has_numbering(para_el):
            continue
        if _para_text(para_el):
            continue   # 有文字，不动

        if as_tc_del:
            tc_del_paragraph(para_el, tc_id, author, date)
            tc_id += 2
        else:
            body.remove(para_el)

        affected.append(idx)

    return affected


# =========================================================================== #
# convenience: run both                                                        #
# =========================================================================== #

def cleanup_all(
    doc,
    as_tc_del: bool = True,
    author: str | None = None,
    para_range: tuple[int, int] | None = None,
    cfg=None,
    keep_styles: list[str] | None = None,
) -> dict:
    """
    一次性运行 remove_empty_paragraphs + remove_orphan_numbering。

    Returns:
        {"empty": [idx, ...], "orphan_numbering": [idx, ...]}
    """
    author = author or _get_cfg_attr(cfg, "author", JT_AUTHOR)
    empty = remove_empty_paragraphs(
        doc, as_tc_del=as_tc_del, author=author,
        para_range=para_range, cfg=cfg, keep_styles=keep_styles,
    )
    orphan = remove_orphan_numbering(
        doc, as_tc_del=as_tc_del, author=author,
        para_range=para_range, cfg=cfg,
    )
    return {"empty": empty, "orphan_numbering": orphan}
