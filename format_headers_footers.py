#!/usr/bin/env python3
"""
北京邮电大学学位论文页眉页脚格式化脚本
- 奇数页页眉：本章标题（STYLEREF 域自动获取）
- 偶数页页眉：北京邮电大学硕士学位论文
- 页眉格式：宋体/TNR，小五号，居中，下方单实线
- 前置部分页码：大写罗马数字 Ⅰ Ⅱ Ⅲ
- 正文部分页码：阿拉伯数字从1开始
- 页码格式：TNR，五号，居中，无修饰线

用法：python3 format_headers_footers.py <unpacked_dir>
"""

import lxml.etree as ET
import os
import re
import sys

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
HEADER_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
FOOTER_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'

HDR_NS_DECL = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'mc:Ignorable="w14 w15 wp14"'
)

RPR_HEADER = (
    '<w:rPr>'
    '<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
    '<w:sz w:val="18"/>'
    '<w:szCs w:val="18"/>'
    '</w:rPr>'
)

RPR_FOOTER = (
    '<w:rPr>'
    '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
    '<w:sz w:val="21"/>'
    '<w:szCs w:val="21"/>'
    '</w:rPr>'
)


def make_even_header():
    """偶数页页眉：北京邮电大学硕士学位论文"""
    return f'''<?xml version="1.0" encoding="utf-8"?>
<w:hdr {HDR_NS_DECL}>
  <w:p>
    <w:pPr>
      <w:pBdr>
        <w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/>
      </w:pBdr>
      <w:jc w:val="center"/>
      {RPR_HEADER}
    </w:pPr>
    <w:r>
      {RPR_HEADER}
      <w:t>北京邮电大学硕士学位论文</w:t>
    </w:r>
  </w:p>
</w:hdr>'''


def make_odd_header_styleref():
    """奇数页页眉：STYLEREF 域自动获取当前章标题"""
    return f'''<?xml version="1.0" encoding="utf-8"?>
<w:hdr {HDR_NS_DECL}>
  <w:p>
    <w:pPr>
      <w:pBdr>
        <w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/>
      </w:pBdr>
      <w:jc w:val="center"/>
      {RPR_HEADER}
    </w:pPr>
    <w:r>
      {RPR_HEADER}
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      {RPR_HEADER}
      <w:instrText xml:space="preserve"> STYLEREF "heading 1" \\* MERGEFORMAT </w:instrText>
    </w:r>
    <w:r>
      {RPR_HEADER}
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      {RPR_HEADER}
      <w:t>章标题</w:t>
    </w:r>
    <w:r>
      {RPR_HEADER}
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:hdr>'''


def make_odd_header_static(text):
    """奇数页页眉：静态文本（用于前置部分如摘要、目录）"""
    return f'''<?xml version="1.0" encoding="utf-8"?>
<w:hdr {HDR_NS_DECL}>
  <w:p>
    <w:pPr>
      <w:pBdr>
        <w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/>
      </w:pBdr>
      <w:jc w:val="center"/>
      {RPR_HEADER}
    </w:pPr>
    <w:r>
      {RPR_HEADER}
      <w:t>{text}</w:t>
    </w:r>
  </w:p>
</w:hdr>'''


def make_empty_header():
    """空页眉（封面用）"""
    return f'''<?xml version="1.0" encoding="utf-8"?>
<w:hdr {HDR_NS_DECL}>
  <w:p>
    <w:pPr>
      <w:pBdr>
        <w:bottom w:val="none" w:sz="0" w:space="1" w:color="auto"/>
      </w:pBdr>
    </w:pPr>
  </w:p>
</w:hdr>'''


def make_footer_with_page():
    """页脚：居中页码，TNR 五号"""
    return f'''<?xml version="1.0" encoding="utf-8"?>
<w:ftr {HDR_NS_DECL}>
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      {RPR_FOOTER}
    </w:pPr>
    <w:r>
      {RPR_FOOTER}
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      {RPR_FOOTER}
      <w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText>
    </w:r>
    <w:r>
      {RPR_FOOTER}
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      {RPR_FOOTER}
      <w:t>1</w:t>
    </w:r>
    <w:r>
      {RPR_FOOTER}
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:ftr>'''


def make_empty_footer():
    """空页脚（封面用）"""
    return f'''<?xml version="1.0" encoding="utf-8"?>
<w:ftr {HDR_NS_DECL}>
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
    </w:pPr>
  </w:p>
</w:ftr>'''


def get_paragraph_text(p):
    texts = []
    for t in p.findall(f'.//{W}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts).strip()


BODY_SPECIAL_TITLES = {
    '参考文献', '致谢', '致 谢', '附录', '附 录',
    '攻读学位期间发表的学术论文', '攻读学位期间发表的学术论文目录',
    '个人简历', '作者简介',
}


def normalize_spaces(text):
    """Remove all whitespace for flexible matching."""
    return re.sub(r'\s+', '', text)


def detect_front_matter(text):
    """Check if text is a front matter title, return normalized title or None."""
    norm = normalize_spaces(text)
    if norm == '摘要':
        return '摘要'
    if text.strip().upper() == 'ABSTRACT':
        return 'ABSTRACT'
    if norm in ('目录',):
        return '目录'
    if norm in ('符号说明', '符号和缩略语说明', '主要符号对照表', '插图和附表清单', '缩略语表'):
        return '符号说明'
    return None


def get_section_content_type(paragraphs_in_section):
    """Determine section type from all paragraphs in this section."""
    for p_text in paragraphs_in_section:
        t = p_text.strip()
        if not t:
            continue
        fm = detect_front_matter(t)
        if fm:
            return 'front', fm
        if re.match(r'^第[一二三四五六七八九十百]+章', t):
            return 'body', t
        if t in BODY_SPECIAL_TITLES:
            return 'body', t
    return 'unknown', ''


def parse_rels(rels_path):
    """Parse document.xml.rels, return dict of rId → (Type, Target)."""
    tree = ET.parse(rels_path)
    root = tree.getroot()
    ns = REL_PKG
    rels = {}
    for rel in root.findall(f'{{{ns}}}Relationship'):
        rid = rel.get('Id')
        rtype = rel.get('Type')
        target = rel.get('Target')
        rels[rid] = (rtype, target)
    return rels, tree, root


def get_max_rid(rels):
    """Find the maximum numeric rId value."""
    max_id = 0
    for rid in rels:
        m = re.match(r'rId(\d+)', rid)
        if m:
            max_id = max(max_id, int(m.group(1)))
    return max_id


def add_relationship(rels_root, rid, rtype, target):
    """Add a relationship to the rels XML."""
    ns = REL_PKG
    el = ET.SubElement(rels_root, f'{{{ns}}}Relationship')
    el.set('Id', rid)
    el.set('Type', rtype)
    el.set('Target', target)


def ensure_content_type(ct_path, partname):
    """Ensure the [Content_Types].xml includes the partname."""
    tree = ET.parse(ct_path)
    root = tree.getroot()
    ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    for override in root.findall(f'{{{ct_ns}}}Override'):
        if override.get('PartName') == partname:
            return
    el = ET.SubElement(root, f'{{{ct_ns}}}Override')
    el.set('PartName', partname)
    if 'header' in partname:
        el.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
    else:
        el.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
    tree.write(ct_path, xml_declaration=True, encoding='UTF-8', standalone=True)


def main():
    if len(sys.argv) < 2:
        print("用法: python3 format_headers_footers.py <unpacked_dir>")
        sys.exit(1)

    unpacked_dir = sys.argv[1]
    word_dir = os.path.join(unpacked_dir, 'word')
    doc_path = os.path.join(word_dir, 'document.xml')
    rels_path = os.path.join(word_dir, '_rels', 'document.xml.rels')
    ct_path = os.path.join(unpacked_dir, '[Content_Types].xml')
    settings_path = os.path.join(word_dir, 'settings.xml')

    doc_tree = ET.parse(doc_path)
    doc_root = doc_tree.getroot()
    rels, rels_tree, rels_root = parse_rels(rels_path)
    next_rid = get_max_rid(rels) + 1
    next_hdr_num = 30
    next_ftr_num = 30

    body = doc_root.find(f'{W}body')
    if body is None:
        print("错误: 找不到 w:body")
        sys.exit(1)

    # Collect all section properties and the paragraphs in each section
    sections = []
    current_paragraphs = []
    is_first_section = True

    for child in body:
        if child.tag == f'{W}p':
            text = get_paragraph_text(child)
            current_paragraphs.append(text)

            pPr = child.find(f'{W}pPr')
            if pPr is not None:
                sectPr = pPr.find(f'{W}sectPr')
                if sectPr is not None:
                    sections.append({
                        'sectPr': sectPr,
                        'paragraphs': current_paragraphs[:],
                        'is_first': is_first_section,
                    })
                    current_paragraphs = []
                    is_first_section = False
        elif child.tag == f'{W}tbl':
            current_paragraphs.append('[table]')

    final_sectPr = body.find(f'{W}sectPr')
    if final_sectPr is not None:
        sections.append({
            'sectPr': final_sectPr,
            'paragraphs': current_paragraphs[:],
            'is_first': is_first_section,
        })

    print(f"找到 {len(sections)} 个节")

    files_written = []
    cover_done = False

    for idx, sec in enumerate(sections):
        sectPr = sec['sectPr']
        paras = sec['paragraphs']

        pgNumType = sectPr.find(f'{W}pgNumType')
        is_roman = False
        if pgNumType is not None:
            fmt = pgNumType.get(f'{W}fmt', '')
            if 'Roman' in fmt or 'roman' in fmt:
                is_roman = True

        content_type, title = get_section_content_type(paras)

        # Roman sections are always front matter (TOC entries may contain chapter titles)
        if is_roman and content_type == 'body':
            content_type = 'front'
            title = '目录'
        if content_type == 'unknown' and is_roman:
            content_type = 'front'
        if content_type == 'unknown' and not is_roman:
            content_type = 'body'

        # Determine if this is a cover section (first 1-2 sections before any front matter)
        is_cover = False
        if not cover_done:
            non_empty = [p for p in paras if p.strip()]
            has_chapter = any(re.match(r'^第[一二三四五六七八九十百]+章', p.strip()) for p in non_empty)
            has_abstract = any(
                normalize_spaces(p) in ('摘要', 'ABSTRACT') for p in non_empty
            )
            has_toc = any(normalize_spaces(p) == '目录' for p in non_empty)
            if not has_chapter and not has_abstract and not has_toc:
                # Check if this section has any meaningful thesis content
                cover_keywords = ['学位论文', '题目', '学号', '姓名', '学科', '导师',
                                  'Master', 'Thesis', 'BEIJING', 'UNIVERSITY',
                                  '硕士', '博士', '密级', '日期', 'Date']
                is_cover_content = any(
                    any(kw in p for kw in cover_keywords)
                    for p in non_empty[:20]
                )
                if is_cover_content or len(non_empty) < 5:
                    is_cover = True
            if has_abstract or has_toc or has_chapter:
                cover_done = True

        # For front matter with empty title, try to guess from position
        if content_type == 'front' and not title:
            prev_types = [s.get('_type', '') for s in sections[:idx] if '_type' in s]
            if 'front-摘要' not in prev_types:
                title = '摘要'
            elif 'front-目录' not in prev_types:
                title = '目录'

        sec['_type'] = f"{content_type}-{title}" if content_type == 'front' else content_type

        if is_cover:
            section_label = f"封面 (节 {idx+1})"
        elif content_type == 'front':
            cover_done = True
            section_label = f"前置-{title} (节 {idx+1})"
        else:
            cover_done = True
            section_label = f"正文-{title[:20] if title else '续'} (节 {idx+1})"

        print(f"  {section_label}: {'罗马' if is_roman else '阿拉伯'}页码")

        # Get existing header/footer references
        existing_refs = {}
        for href in sectPr.findall(f'{W}headerReference'):
            htype = href.get(f'{W}type')
            rid = href.get(f'{R_NS}id')
            existing_refs[('header', htype)] = (href, rid)

        for fref in sectPr.findall(f'{W}footerReference'):
            ftype = fref.get(f'{W}type')
            rid = fref.get(f'{R_NS}id')
            existing_refs[('footer', ftype)] = (fref, rid)

        def write_hdr_ftr(kind, hf_type, content):
            """Write a header/footer file. Returns the rId."""
            nonlocal next_rid, next_hdr_num, next_ftr_num

            key = (kind, hf_type)
            if key in existing_refs:
                _, rid = existing_refs[key]
                if rid in rels:
                    target = rels[rid][1]
                    filepath = os.path.join(word_dir, target)
                    with open(filepath, 'w', encoding='utf-8') as f:
                        f.write(content)
                    files_written.append(target)
                    return rid

            if kind == 'header':
                filename = f'header{next_hdr_num}.xml'
                next_hdr_num += 1
                rel_type = HEADER_TYPE
            else:
                filename = f'footer{next_ftr_num}.xml'
                next_ftr_num += 1
                rel_type = FOOTER_TYPE

            filepath = os.path.join(word_dir, filename)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            files_written.append(filename)

            rid = f'rId{next_rid}'
            next_rid += 1
            add_relationship(rels_root, rid, rel_type, filename)
            ensure_content_type(ct_path, f'/word/{filename}')

            ref_el = ET.SubElement(sectPr, f'{W}headerReference' if kind == 'header' else f'{W}footerReference')
            ref_el.set(f'{W}type', hf_type)
            ref_el.set(f'{R_NS}id', rid)

            # Move the new reference to the beginning of sectPr
            sectPr.remove(ref_el)
            sectPr.insert(0, ref_el)

            return rid

        if is_cover:
            write_hdr_ftr('header', 'default', make_empty_header())
            write_hdr_ftr('header', 'even', make_empty_header())
            write_hdr_ftr('header', 'first', make_empty_header())
            write_hdr_ftr('footer', 'default', make_empty_footer())
            write_hdr_ftr('footer', 'even', make_empty_footer())
            write_hdr_ftr('footer', 'first', make_empty_footer())
        elif content_type == 'front':
            odd_title = title if title else '摘要'
            write_hdr_ftr('header', 'default', make_odd_header_static(odd_title))
            write_hdr_ftr('header', 'even', make_even_header())
            write_hdr_ftr('footer', 'default', make_footer_with_page())
            write_hdr_ftr('footer', 'even', make_footer_with_page())

            if pgNumType is None:
                pgNumType = ET.SubElement(sectPr, f'{W}pgNumType')
                pgNumType.set(f'{W}fmt', 'upperRoman')
        else:
            write_hdr_ftr('header', 'default', make_odd_header_styleref())
            write_hdr_ftr('header', 'even', make_even_header())
            write_hdr_ftr('footer', 'default', make_footer_with_page())
            write_hdr_ftr('footer', 'even', make_footer_with_page())

    # Ensure evenAndOddHeaders is in settings.xml
    if os.path.exists(settings_path):
        settings_tree = ET.parse(settings_path)
        settings_root = settings_tree.getroot()
        if settings_root.find(f'{W}evenAndOddHeaders') is None:
            el = ET.SubElement(settings_root, f'{W}evenAndOddHeaders')
            settings_root.insert(0, el)
            settings_tree.write(settings_path, xml_declaration=True, encoding='UTF-8', standalone=True)
            print("\n已在 settings.xml 中启用奇偶页页眉")

    # Save document.xml and rels
    doc_tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    rels_tree.write(rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    print(f"\n已写入/更新 {len(files_written)} 个页眉页脚文件")
    print("\n页眉页脚规范:")
    print("  偶数页页眉: 北京邮电大学硕士学位论文")
    print("  奇数页页眉: 本章标题（STYLEREF 域自动获取）")
    print("  页眉格式:   宋体/TNR 小五号 居中 单实线底边框")
    print("  前置页码:   大写罗马数字 居中 TNR 五号")
    print("  正文页码:   阿拉伯数字从1开始 居中 TNR 五号")


if __name__ == '__main__':
    main()
