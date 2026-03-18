"""
footer_ops.py — Footer 检查与替换工具

提供对所有 footer OPC parts 的枚举（含奇偶页/首页）、
文本内容提取（含 textbox）、以及批量查找替换。

用法：
    from lex_docx import footer_ops

    # 审查所有 footer parts
    parts = footer_ops.audit_footers(doc)
    for p in parts:
        print(p["rId"], p["footer_type"], p["text"], p["has_textbox"])

    # 批量替换 footer 文本
    count = footer_ops.fill_footer(doc, find="Auspicious Linkage", replace="Rokid HK Ltd")
"""
from __future__ import annotations

from docx.oxml.ns import qn

_FOOTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"

# w:footerReference 的 r:id 属性使用 relationships 命名空间
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


# --------------------------------------------------------------------------- #
# 辅助                                                                          #
# --------------------------------------------------------------------------- #

def _build_rid_type_map(doc) -> dict[str, str]:
    """
    扫描文档所有 w:sectPr 中的 w:footerReference，
    构建 {rId: footer_type} 映射（type 为 "default"/"even"/"first"）。
    """
    rid_type: dict[str, str] = {}
    for sect_pr in doc.element.body.iter(qn("w:sectPr")):
        for ref in sect_pr.findall(qn("w:footerReference")):
            rid = ref.get(_R_ID)
            ftype = ref.get(qn("w:type"), "default")
            if rid:
                rid_type[rid] = ftype
    return rid_type


# --------------------------------------------------------------------------- #
# 公开接口                                                                      #
# --------------------------------------------------------------------------- #

def audit_footers(doc) -> list[dict]:
    """
    枚举文档所有 footer OPC parts，返回结构化信息列表。

    每个元素包含：
        rId:         关系 ID（如 "rId3"）
        part_name:   OPC part 路径（如 "/word/footer1.xml"）
        footer_type: "default" | "even" | "first" | "unknown"
                     default = 奇数页/通用；even = 偶数页；first = 首页
        text:        footer 全部文本（含 textbox 内 w:t）
        has_textbox: 是否含 w:txbxContent

    Args:
        doc: python-docx Document

    Returns:
        按 rId 排序的 footer 信息列表
    """
    rid_type = _build_rid_type_map(doc)
    results = []

    for rId, rel in doc.part.rels.items():
        if rel.reltype != _FOOTER_REL:
            continue
        part = rel.target_part
        ftr_el = part._element

        # 提取全部文本（含 textbox 内的 w:t）
        text = "".join(el.text or "" for el in ftr_el.iter(qn("w:t")))

        # 检测 textbox
        has_textbox = ftr_el.find(".//" + qn("w:txbxContent")) is not None

        results.append({
            "rId":         rId,
            "part_name":   str(part.partname),
            "footer_type": rid_type.get(rId, "unknown"),
            "text":        text,
            "has_textbox": has_textbox,
        })

    results.sort(key=lambda x: x["rId"])
    return results


def fill_footer(doc, find: str, replace: str) -> int:
    """
    在所有 footer parts（奇偶页/首页）中查找并替换文本。
    同时处理 textbox（w:txbxContent）内的 w:t 元素。

    Args:
        doc:     python-docx Document
        find:    要查找的字符串
        replace: 替换目标字符串

    Returns:
        实际替换次数（w:t 元素级别）
    """
    count = 0
    for rId, rel in doc.part.rels.items():
        if rel.reltype != _FOOTER_REL:
            continue
        ftr_el = rel.target_part._element
        for t_el in ftr_el.iter(qn("w:t")):
            if t_el.text and find in t_el.text:
                t_el.text = t_el.text.replace(find, replace)
                # 确保含空格的文本保留 xml:space="preserve"
                if " " in t_el.text:
                    t_el.set(_XML_SPACE, "preserve")
                count += 1
    return count
