from __future__ import annotations
from typing import List, Optional
from io import BytesIO
import aspose.words as aw
from core.utils.docs_util import ensure_path, move_builder, resolve_heading_style_identifier, resolve_outline_level, hex_to_color, find_paragraph_indices_by_anchor

def find_heading_style_by_name(doc: aw.Document, level: int):
    lvl = level
    preferred = None
    fallback = None
    for s in doc.styles:
        if s.type != aw.StyleType.PARAGRAPH:
            continue
        name = (s.name or '').strip()
        low = name.lower()
        digit = None
        for tok in low.replace('_', ' ').split():
            if tok.isdigit():
                digit = int(tok)
                break
        if digit == lvl:
            if 'heading' in low:
                preferred = s
                break
            if fallback is None:
                fallback = s
    return preferred or fallback

def get_heading_style_object(doc: aw.Document, level: int):
    sid = resolve_heading_style_identifier(level)
    style_obj = doc.styles.get_by_style_identifier(sid)
    if style_obj is None:
        style_obj = doc.styles.get_by_name(f'Heading {level}')
        if style_obj is None:
            style_obj = find_heading_style_by_name(doc, level)
    return style_obj

def insert_text(doc_id: str, text: str='', where: str='end', paragraph_index: Optional[int]=None) -> bool:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    builder.write(text)
    doc.save(str(file_path))
    return True

def replace_text(doc_id: str, search: str='', replace: str='', replace_all: bool=True, case_sensitive: bool=False, whole_word: bool=False, use_regex: bool=False) -> int:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    options = aw.replacing.FindReplaceOptions()
    options.match_case = bool(case_sensitive)
    options.direction = aw.replacing.FindReplaceDirection.FORWARD
    options.find_whole_words_only = bool(whole_word)
    if not replace_all:
        options.max_matches = 1

    if use_regex:
        import re
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern = re.compile(search, flags)
        count = doc.range.replace(pattern, replace, options)
    else:
        count = doc.range.replace(search, replace, options)
    doc.save(str(file_path))
    return int(count)

def read_paragraphs(doc_id: str, start: Optional[int]=None, end: Optional[int]=None):
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    nodes = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    texts: List[str] = []
    for i in range(nodes.count):
        para = nodes[i]
        t = para.to_string(aw.SaveFormat.TEXT) or ''
        texts.append(t)
    s = start or 0
    e = end if end is not None else len(texts)
    s = max(0, min(s, len(texts)))
    e = max(s, min(e, len(texts)))
    return texts[s:e]

def insert_image(doc_id: str, image_bytes: bytes=b'', where: str='end', width_points: Optional[float]=None, height_points: Optional[float]=None, keep_aspect: bool=False) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, None)
    with BytesIO(image_bytes) as buf:
        if width_points is not None and height_points is not None:
            builder.insert_image(buf, float(width_points), float(height_points))
        else:
            shape = builder.insert_image(buf)
            if keep_aspect and (width_points is not None) != (height_points is not None):
                if width_points is not None:
                    ratio = shape.height / shape.width if shape.width else 1.0
                    shape.width = float(width_points)
                    shape.height = float(width_points) * ratio
                else:
                    ratio = shape.width / shape.height if shape.height else 1.0
                    shape.height = float(height_points)
                    shape.width = float(height_points) * ratio
            else:
                if width_points is not None:
                    shape.width = float(width_points)
                if height_points is not None:
                    shape.height = float(height_points)
    doc.save(str(path))
    return True

def insert_html(doc_id: str, html: str='', where: str='end', paragraph_index: Optional[int]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    builder.insert_html(html or '')
    doc.save(str(path))
    return True

def insert_markdown(doc_id: str, markdown: str='', where: str='end', paragraph_index: Optional[int]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    builder.write(markdown or '')
    doc.save(str(path))
    return True

def add_heading(doc_id: str, text: str='', level: int=1, font_name: Optional[str]=None, font_size: Optional[float]=None, bold: Optional[bool]=None, italic: Optional[bool]=None, border_bottom: Optional[bool]=None, where: str='end', paragraph_index: Optional[int]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    builder.insert_paragraph()
    style_id = resolve_heading_style_identifier(level)
    builder.paragraph_format.style_identifier = style_id
    builder.paragraph_format.outline_level = resolve_outline_level(level)
    if font_name:
        builder.font.name = font_name
    if font_size:
        builder.font.size = font_size
    if bold is not None:
        builder.font.bold = bold
    if italic is not None:
        builder.font.italic = italic
    if border_bottom:
        builder.paragraph_format.borders.bottom.line_style = aw.LineStyle.SINGLE
    builder.writeln(text or '')
    doc.save(str(path))
    return True

def add_paragraph(doc_id: str, text: str='', style: Optional[str]=None, font_name: Optional[str]=None, font_size: Optional[float]=None, bold: Optional[bool]=None, italic: Optional[bool]=None, color_hex: Optional[str]=None, where: str='end', paragraph_index: Optional[int]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    if style:
        s = doc.styles.get_by_name(style)
        if s is not None:
            builder.paragraph_format.style = s
    if font_name:
        builder.font.name = font_name
    if font_size:
        builder.font.size = font_size
    if bold is not None:
        builder.font.bold = bold
    if italic is not None:
        builder.font.italic = italic
    col = hex_to_color(color_hex)
    if col is not None:
        builder.font.color = col
    builder.writeln(text or '')
    doc.save(str(path))
    return True

def add_page_break(doc_id: str, where: str='end', paragraph_index: Optional[int]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    builder.insert_break(aw.BreakType.PAGE_BREAK)
    doc.save(str(path))
    return True

def insert_list(doc_id: str, items: List[str]=None, kind: str='bullet', where: str='end', paragraph_index: Optional[int]=None) -> bool:
    if items is None:
        items = []
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    # Применяем список по умолчанию как в примерах: bullet или numbered
    if (kind or '').lower().startswith('number'):
        builder.list_format.apply_number_default()
    else:
        builder.list_format.apply_bullet_default()
    for it in items or []:
        builder.writeln(str(it))
    builder.list_format.remove_numbers()
    doc.save(str(path))
    return True

def insert_near_text(doc_id: str, target_text: Optional[str]=None, target_paragraph_index: Optional[int]=None, position: str='after', content_type: str='paragraph', text: Optional[str]=None, level: Optional[int]=None, items: Optional[List[str]]=None, kind: Optional[str]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    pidx = None
    if target_paragraph_index is not None and 0 <= target_paragraph_index < paras.count:
        pidx = int(target_paragraph_index)
    elif target_text:
        matches = find_paragraph_indices_by_anchor(doc, target_text)
        if matches:
            pidx = matches[0]
    if pidx is None:
        pidx = max(0, paras.count - 1)
    builder = aw.DocumentBuilder(doc)
    if (position or 'after') == 'before':
        builder.move_to_paragraph(pidx, 0)
    else:
        builder.move_to_paragraph(pidx, -1)
    ctype = (content_type or 'paragraph').lower()
    if ctype == 'heading':
        add_heading(doc_id=doc_id, text=text or '', level=level or 1, where='paragraph', paragraph_index=pidx)
    elif ctype == 'list':
        insert_list(doc_id=doc_id, items=items or [], kind=kind or 'bullet', where='paragraph', paragraph_index=pidx)
    else:
        builder.writeln(text or '')
        doc.save(str(path))
    return True

def format_range(doc_id: str, paragraph_index: int=0, start: int=0, end: int=0, bold: Optional[bool]=None, italic: Optional[bool]=None, underline: Optional[bool]=None, color_hex: Optional[str]=None, font_name: Optional[str]=None, font_size: Optional[float]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    if paragraph_index < 0 or paragraph_index >= paras.count:
        raise IndexError('paragraph_index out of range')
    para = paras[paragraph_index]
    acc = 0
    p_obj = para.as_paragraph()
    runs = p_obj.runs
    count = runs.count
    for i in range(count):
        run = runs[i]
        text = run.text or ''
        length = len(text)
        r_start = acc
        r_end = acc + length
        acc = r_end
        if length == 0:
            continue
        if r_end <= start or r_start >= end:
            continue
        font = run.font
        if bold is not None:
            font.bold = bold
        if italic is not None:
            font.italic = italic
        if underline is not None:
            font.underline = aw.Underline.SINGLE if underline else aw.Underline.NONE
        col = hex_to_color(color_hex)
        if col is not None:
            font.color = col
        if font_name:
            font.name = font_name
        if font_size:
            font.size = font_size
    doc.save(str(path))
    return True

def delete_paragraph(doc_id: str, paragraph_index: int=0) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    if paragraph_index < 0 or paragraph_index >= paras.count:
        raise IndexError('paragraph_index out of range')
    node = paras[paragraph_index]
    node.remove()
    doc.save(str(path))
    return True
