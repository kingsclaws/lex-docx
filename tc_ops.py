"""
tc_ops.py — Track Changes accept/reject and comment/header cleanup

Functions:
  list_tc(doc, author_filter)      → list all w:ins / w:del entries
  accept_all(doc, author_filter)   → accept all tracked changes
  reject_all(doc, author_filter)   → reject all tracked changes
  clean_comments(doc)              → remove comment annotations from body + comments part
  clean_headers(doc)               → clear all header text content
"""
from __future__ import annotations

from docx.oxml.ns import qn


# ────────────────────────────────────────────────────────────────────────────── #
# Helpers                                                                         #
# ────────────────────────────────────────────────────────────────────────────── #

_COMMENTS_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
_HEADERS_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
)


def _author_ok(el, author_filter: str | None) -> bool:
    return author_filter is None or el.get(qn("w:author")) == author_filter


def _collect_body_ins(body, author_filter):
    """Collect all w:ins not inside w:trPr (text-level / pPr-rPr excluded)."""
    result = []
    for ins_el in body.iter(qn("w:ins")):
        parent = ins_el.getparent()
        if parent is not None and parent.tag == qn("w:trPr"):
            continue
        if _author_ok(ins_el, author_filter):
            result.append(ins_el)
    return result


def _collect_body_del(body, author_filter, *, skip_para_mark: bool = True):
    """Collect all w:del not inside w:trPr, optionally skipping pPr>rPr>del."""
    result = []
    for del_el in body.iter(qn("w:del")):
        parent = del_el.getparent()
        if parent is None:
            continue
        if parent.tag == qn("w:trPr"):
            continue
        if skip_para_mark and parent.tag == qn("w:rPr"):
            gp = parent.getparent()
            if gp is not None and gp.tag == qn("w:pPr"):
                continue
        if _author_ok(del_el, author_filter):
            result.append(del_el)
    return result


# ────────────────────────────────────────────────────────────────────────────── #
# list_tc                                                                         #
# ────────────────────────────────────────────────────────────────────────────── #

def list_tc(doc, author_filter: str | None = None) -> list[dict]:
    """
    List all tracked changes (w:ins / w:del) in the document body.

    Args:
        doc:           python-docx Document
        author_filter: only return entries by this author; None = all

    Returns:
        List of dicts with keys: id, type, author, date, level, text
          level: "text" (run-level), "row" (table row), "para_mark" (¶ mark)
    """
    body = doc.element.body
    items: list[dict] = []

    for el in body.iter():
        if el.tag not in (qn("w:ins"), qn("w:del")):
            continue
        if not _author_ok(el, author_filter):
            continue

        tc_type = "ins" if el.tag == qn("w:ins") else "del"
        tc_id = el.get(qn("w:id"), "?")
        author = el.get(qn("w:author"), "")
        date = el.get(qn("w:date"), "")

        parent = el.getparent()
        if parent is not None and parent.tag == qn("w:trPr"):
            level = "row"
        elif parent is not None and parent.tag == qn("w:rPr"):
            gp = parent.getparent()
            if gp is not None and gp.tag == qn("w:pPr"):
                level = "para_mark"
            else:
                level = "text"
        else:
            level = "text"

        text_parts: list[str] = []
        text_tag = qn("w:t") if tc_type == "ins" else qn("w:delText")
        for t_el in el.iter(text_tag):
            if t_el.text:
                text_parts.append(t_el.text)

        items.append({
            "id": tc_id,
            "type": tc_type,
            "author": author,
            "date": date,
            "level": level,
            "text": "".join(text_parts)[:120],
        })

    return items


# ────────────────────────────────────────────────────────────────────────────── #
# accept_all                                                                      #
# ────────────────────────────────────────────────────────────────────────────── #

def accept_all(doc, author_filter: str | None = None) -> dict:
    """
    Accept all tracked changes in the document.

    - w:ins (text-level): unwrap → keep the inserted runs
    - w:del (text-level): remove → discard deleted runs
    - w:trPr > w:ins (row INS): remove marker, keep row
    - w:trPr > w:del (row DEL): remove the entire row
    - w:pPr > w:rPr > w:del (¶ mark DEL): remove the marker

    Args:
        doc:           python-docx Document
        author_filter: only accept changes by this author; None = accept all

    Returns:
        dict with counts: ins_accepted, del_accepted, row_ins_accepted,
                          row_del_accepted, para_mark_cleaned
    """
    body = doc.element.body
    stats = {
        "ins_accepted": 0,
        "del_accepted": 0,
        "row_ins_accepted": 0,
        "row_del_accepted": 0,
        "para_mark_cleaned": 0,
    }

    # ── 1. Table-row TC ───────────────────────────────────────────────── #
    for tbl in list(body.iter(qn("w:tbl"))):
        rows_to_remove: list = []
        for tr in list(tbl.findall(qn("w:tr"))):
            trPr = tr.find(qn("w:trPr"))
            if trPr is None:
                continue
            # Row INS → keep row, remove marker
            for ins_el in list(trPr.findall(qn("w:ins"))):
                if _author_ok(ins_el, author_filter):
                    trPr.remove(ins_el)
                    stats["row_ins_accepted"] += 1
            # Row DEL → remove row
            del_els = [e for e in trPr.findall(qn("w:del")) if _author_ok(e, author_filter)]
            if del_els:
                rows_to_remove.append(tr)
                stats["row_del_accepted"] += 1
        for tr in rows_to_remove:
            parent = tr.getparent()
            if parent is not None:
                parent.remove(tr)

    # ── 2. Text-level w:ins → unwrap (keep children) ─────────────────── #
    for ins_el in _collect_body_ins(body, author_filter):
        parent = ins_el.getparent()
        if parent is None or ins_el.getparent() is None:
            continue
        idx = list(parent).index(ins_el)
        children = list(ins_el)
        for child in children:
            ins_el.remove(child)
        for i, child in enumerate(children):
            parent.insert(idx + i, child)
        parent.remove(ins_el)
        stats["ins_accepted"] += 1

    # ── 3. Text-level w:del → remove entirely ────────────────────────── #
    for del_el in _collect_body_del(body, author_filter, skip_para_mark=True):
        parent = del_el.getparent()
        if parent is not None and del_el.getparent() is parent:
            parent.remove(del_el)
            stats["del_accepted"] += 1

    # ── 4. Paragraph-mark deletion (pPr > rPr > del) → clean marker ──── #
    for p_el in list(body.iter(qn("w:p"))):
        pPr = p_el.find(qn("w:pPr"))
        if pPr is None:
            continue
        rPr = pPr.find(qn("w:rPr"))
        if rPr is None:
            continue
        for del_el in list(rPr.findall(qn("w:del"))):
            if _author_ok(del_el, author_filter):
                rPr.remove(del_el)
                stats["para_mark_cleaned"] += 1

    return stats


# ────────────────────────────────────────────────────────────────────────────── #
# reject_all                                                                      #
# ────────────────────────────────────────────────────────────────────────────── #

def reject_all(doc, author_filter: str | None = None) -> dict:
    """
    Reject all tracked changes in the document.

    - w:ins (text-level): remove → discard inserted runs
    - w:del (text-level): unwrap + restore → convert w:delText→w:t, keep runs
    - w:trPr > w:ins (row INS): remove the entire row
    - w:trPr > w:del (row DEL): remove marker, keep row
    - w:pPr > w:rPr > w:del (¶ mark DEL): remove the marker (restore paragraph break)

    Args:
        doc:           python-docx Document
        author_filter: only reject changes by this author; None = reject all

    Returns:
        dict with counts: ins_rejected, del_rejected, row_ins_rejected,
                          row_del_rejected, para_mark_cleaned
    """
    body = doc.element.body
    stats = {
        "ins_rejected": 0,
        "del_rejected": 0,
        "row_ins_rejected": 0,
        "row_del_rejected": 0,
        "para_mark_cleaned": 0,
    }

    # ── 1. Table-row TC ───────────────────────────────────────────────── #
    for tbl in list(body.iter(qn("w:tbl"))):
        rows_to_remove: list = []
        for tr in list(tbl.findall(qn("w:tr"))):
            trPr = tr.find(qn("w:trPr"))
            if trPr is None:
                continue
            # Row INS → remove row
            ins_els = [e for e in trPr.findall(qn("w:ins")) if _author_ok(e, author_filter)]
            if ins_els:
                rows_to_remove.append(tr)
                stats["row_ins_rejected"] += 1
            # Row DEL → keep row, remove marker
            for del_el in list(trPr.findall(qn("w:del"))):
                if _author_ok(del_el, author_filter):
                    trPr.remove(del_el)
                    stats["row_del_rejected"] += 1
        for tr in rows_to_remove:
            parent = tr.getparent()
            if parent is not None:
                parent.remove(tr)

    # ── 2. Text-level w:ins → remove entirely ────────────────────────── #
    for ins_el in _collect_body_ins(body, author_filter):
        parent = ins_el.getparent()
        if parent is not None and ins_el.getparent() is parent:
            parent.remove(ins_el)
            stats["ins_rejected"] += 1

    # ── 3. Text-level w:del → unwrap (restore deleted text) ──────────── #
    for del_el in _collect_body_del(body, author_filter, skip_para_mark=True):
        parent = del_el.getparent()
        if parent is None or del_el.getparent() is not parent:
            continue
        # Convert w:delText → w:t in child runs
        for r_el in del_el.findall(qn("w:r")):
            for dt_el in r_el.findall(qn("w:delText")):
                dt_el.tag = qn("w:t")
        # Unwrap: move run children before del_el, remove del_el
        idx = list(parent).index(del_el)
        children = list(del_el)
        for child in children:
            del_el.remove(child)
        for i, child in enumerate(children):
            parent.insert(idx + i, child)
        parent.remove(del_el)
        stats["del_rejected"] += 1

    # ── 4. Paragraph-mark deletion (pPr > rPr > del) → restore ──────── #
    for p_el in list(body.iter(qn("w:p"))):
        pPr = p_el.find(qn("w:pPr"))
        if pPr is None:
            continue
        rPr = pPr.find(qn("w:rPr"))
        if rPr is None:
            continue
        for del_el in list(rPr.findall(qn("w:del"))):
            if _author_ok(del_el, author_filter):
                rPr.remove(del_el)
                stats["para_mark_cleaned"] += 1

    return stats


# ────────────────────────────────────────────────────────────────────────────── #
# clean_comments                                                                  #
# ────────────────────────────────────────────────────────────────────────────── #

def clean_comments(doc) -> dict:
    """
    Remove all comment annotations from the document.

    Removes from body:
      - w:commentRangeStart / w:commentRangeEnd elements
      - Runs containing w:commentReference, w:annotationRef, or
        rStyle val="CommentReference"

    Clears the comments OPC part if present (removes all w:comment children).

    Returns:
        dict with counts: range_starts, range_ends, ref_runs, comments_cleared
    """
    body = doc.element.body
    stats = {
        "range_starts": 0,
        "range_ends": 0,
        "ref_runs": 0,
        "comments_cleared": 0,
    }

    # ── 1. Remove w:commentRangeStart / w:commentRangeEnd ────────────── #
    for tag, key in (
        (qn("w:commentRangeStart"), "range_starts"),
        (qn("w:commentRangeEnd"), "range_ends"),
    ):
        for el in list(body.iter(tag)):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
                stats[key] += 1

    # ── 2. Remove comment-reference runs ─────────────────────────────── #
    for r_el in list(body.iter(qn("w:r"))):
        is_comment_run = (
            r_el.find(qn("w:commentReference")) is not None
            or r_el.find(qn("w:annotationRef")) is not None
        )
        if not is_comment_run:
            rPr = r_el.find(qn("w:rPr"))
            if rPr is not None:
                rStyle = rPr.find(qn("w:rStyle"))
                if rStyle is not None and rStyle.get(qn("w:val")) == "CommentReference":
                    is_comment_run = True
        if is_comment_run:
            parent = r_el.getparent()
            if parent is not None:
                parent.remove(r_el)
                stats["ref_runs"] += 1

    # ── 3. Clear comments OPC part ────────────────────────────────────── #
    try:
        for rel in doc.part.rels.values():
            if _COMMENTS_REL in rel.reltype:
                comments_root = rel.target_part._element
                for comment_el in list(comments_root):
                    comments_root.remove(comment_el)
                stats["comments_cleared"] += 1
                break
    except Exception:
        pass

    return stats


# ────────────────────────────────────────────────────────────────────────────── #
# clean_headers                                                                   #
# ────────────────────────────────────────────────────────────────────────────── #

def clean_headers(doc, *, clear_text: bool = True, remove_refs: bool = False) -> dict:
    """
    Clean document headers.

    Args:
        doc:         python-docx Document
        clear_text:  Clear all text content from header parts (default True)
        remove_refs: Also remove w:headerReference elements from w:sectPr (default False)

    Returns:
        dict with counts: headers_cleared, header_refs_removed
    """
    stats = {"headers_cleared": 0, "header_refs_removed": 0}

    # ── 1. Clear header text content ─────────────────────────────────── #
    if clear_text:
        try:
            for rel in doc.part.rels.values():
                if _HEADERS_REL in rel.reltype:
                    hdr_root = rel.target_part._element
                    # Clear all w:t text in the header part
                    for t_el in hdr_root.iter(qn("w:t")):
                        t_el.text = ""
                    stats["headers_cleared"] += 1
        except Exception:
            pass

    # ── 2. Remove header references from w:sectPr ────────────────────── #
    if remove_refs:
        body = doc.element.body
        for sect_el in list(body.iter(qn("w:sectPr"))):
            for href in list(sect_el.findall(qn("w:headerReference"))):
                sect_el.remove(href)
                stats["header_refs_removed"] += 1

    return stats
