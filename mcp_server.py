import base64
import logging
import os
from pathlib import Path
from typing import Optional

from fastmcp import FastMCP

from core import comments as _comments
from core import content as _content
from core import export as _export
from core import io as _io
from core import layout as _layout
from core import links as _links
from core import notes as _notes
from core import properties as _properties
from core import protection as _protection
from core import reading as _reading
from core import styles as _styles
from core import tables as _tables
from core import watermarks as _watermarks
from core.store import document_store
from core.utils import license as _license

mcp = FastMCP('Aspose.Words MCP')


def _setup_logging():
    level = os.getenv('LOG_LEVEL', '')
    logging.basicConfig(
        level=getattr(logging, level, logging.INFO),
        format='%(asctime)s - %(levelname)s - %(message)s',
    )
    return logging.getLogger('mcp')


def _ensure_data_dir_initialized():
    from core.utils import docs_util as _docs_mod

    data_dir = Path(os.getenv('DOCS_DATA_DIR', './data'))
    _docs_mod.init_data_dir(data_dir)


def _store_add_mapping(doc_id: str, name: Optional[str]) -> None:
    if name:
        document_store.add_document(doc_id, name)


def _find_doc_id_by_name(name: str) -> Optional[str]:
    items = document_store.get_all_documents()
    for did, info in items.items():
        if info.get('name') == name:
            return did
    return None


def tool_create_document(name: Optional[str] = None):
    doc_id, fname = _io.create_document(name)
    _store_add_mapping(doc_id, fname)
    return {'docId': doc_id, 'name': fname}


def tool_insert_text_end(doc_id: str, text: str):
    _content.insert_text(doc_id=doc_id, text=text, where='end', paragraph_index=None)
    return {}


def tool_insert_text_start(doc_id: str, text: str):
    _content.insert_text(doc_id=doc_id, text=text, where='start', paragraph_index=None)
    return {}


def tool_insert_text_at_paragraph(doc_id: str, text: str, paragraph_index: int):
    _content.insert_text(
        doc_id=doc_id, text=text, where='paragraph', paragraph_index=paragraph_index
    )
    return {}


def tool_read_paragraphs(doc_id: str, start: Optional[int] = None, end: Optional[int] = None):
    paras = _reading.read_paragraphs(doc_id=doc_id, start=start, end=end)
    return {'paragraphs': paras}


def tool_export_base64(doc_id: str, fmt: str = 'docx'):
    data, mime, ext = _export.export(doc_id, fmt=fmt)
    b64 = base64.b64encode(data).decode('utf-8')
    return {'base64': b64, 'mime': mime, 'ext': ext}


def tool_get_info(doc_id: str):
    return _reading.get_info(doc_id=doc_id)


def tool_add_heading(
    doc_id: str,
    text: str,
    level: int = 1,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    border_bottom: Optional[bool] = None,
):
    _content.add_heading(
        doc_id=doc_id,
        text=text,
        level=level,
        font_name=font_name,
        font_size=font_size,
        bold=bold,
        italic=italic,
        border_bottom=border_bottom,
    )
    return {}


def tool_add_paragraph(
    doc_id: str,
    text: str,
    style: Optional[str] = None,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color: Optional[str] = None,
):
    _content.add_paragraph(
        doc_id=doc_id,
        text=text,
        style=style,
        font_name=font_name,
        font_size=font_size,
        bold=bold,
        italic=italic,
        color_hex=color,
    )
    return {}


def tool_add_page_break_end(doc_id: str):
    _content.add_page_break(doc_id=doc_id, where='end', paragraph_index=None)
    return {}


def tool_add_page_break_start(doc_id: str):
    _content.add_page_break(doc_id=doc_id, where='start', paragraph_index=None)
    return {}


def tool_add_page_break_at_paragraph(doc_id: str, paragraph_index: int):
    _content.add_page_break(doc_id=doc_id, where='paragraph', paragraph_index=paragraph_index)
    return {}


def tool_get_outline(doc_id: str):
    items = _reading.get_outline(doc_id=doc_id)
    return {'outline': items}


def tool_replace_text(
    doc_id: str,
    find_text: str,
    replace_text: str,
    replace_all: bool = True,
    case_sensitive: bool = False,
):
    count = _content.replace_text(
        doc_id=doc_id,
        search=find_text,
        replace=replace_text,
        replace_all=replace_all,
        case_sensitive=case_sensitive,
    )
    return {'count': int(count)}


def tool_find_text(doc_id: str, text: str, match_case: bool = False, whole_word: bool = False):
    matches = _reading.find_text(
        doc_id=doc_id, text=text, match_case=match_case, whole_word=whole_word
    )
    return {'matches': matches}


def tool_delete_paragraph(doc_id: str, paragraph_index: int):
    _content.delete_paragraph(doc_id=doc_id, paragraph_index=paragraph_index)
    return {}


def tool_create_style(
    doc_id: str,
    style_name: str,
    base_style: Optional[str] = None,
    font_size: Optional[float] = None,
    font_name: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color: Optional[str] = None,
):
    s = _styles.create_style(
        doc_id=doc_id,
        style_name=style_name,
        base_style=base_style,
        font_size=font_size,
        font_name=font_name,
        bold=bold,
        italic=italic,
        color_hex=color,
    )
    return {'style': s}


def tool_add_table_end(
    doc_id: str, rows: int, cols: int, data: Optional[list] = None, has_header_row: bool = False
):
    idx = _tables.add_table(
        doc_id,
        rows,
        cols,
        data=data,
        has_header_row=has_header_row,
        where='end',
        paragraph_index=None,
    )
    return {'tableIndex': int(idx)}


def tool_add_table_start(
    doc_id: str, rows: int, cols: int, data: Optional[list] = None, has_header_row: bool = False
):
    idx = _tables.add_table(
        doc_id,
        rows,
        cols,
        data=data,
        has_header_row=has_header_row,
        where='start',
        paragraph_index=None,
    )
    return {'tableIndex': int(idx)}


def tool_add_table_at_paragraph(
    doc_id: str,
    rows: int,
    cols: int,
    paragraph_index: int,
    data: Optional[list] = None,
    has_header_row: bool = False,
):
    idx = _tables.add_table(
        doc_id,
        rows,
        cols,
        data=data,
        has_header_row=has_header_row,
        where='paragraph',
        paragraph_index=paragraph_index,
    )
    return {'tableIndex': int(idx)}


def tool_insert_list_end(doc_id: str, items: list, kind: str = 'bullet'):
    _content.insert_list(doc_id=doc_id, items=items, kind=kind, where='end', paragraph_index=None)
    return {}


def tool_insert_list_start(doc_id: str, items: list, kind: str = 'bullet'):
    _content.insert_list(doc_id=doc_id, items=items, kind=kind, where='start', paragraph_index=None)
    return {}


def tool_insert_list_at_paragraph(
    doc_id: str, items: list, paragraph_index: int, kind: str = 'bullet'
):
    _content.insert_list(
        doc_id=doc_id, items=items, kind=kind, where='paragraph', paragraph_index=paragraph_index
    )
    return {}


def tool_insert_header_near_text(
    doc_id: str,
    target_text: Optional[str] = None,
    header_title: Optional[str] = None,
    position: str = 'after',
    level: int = 1,
    target_paragraph_index: Optional[int] = None,
):
    _content.insert_near_text(
        doc_id=doc_id,
        target_text=target_text,
        target_paragraph_index=target_paragraph_index,
        position=position,
        content_type='heading',
        text=header_title,
        level=level,
    )
    return {}


def tool_insert_line_or_paragraph_near_text(
    doc_id: str,
    target_text: Optional[str] = None,
    line_text: Optional[str] = None,
    position: str = 'after',
    target_paragraph_index: Optional[int] = None,
):
    _content.insert_near_text(
        doc_id=doc_id,
        target_text=target_text,
        target_paragraph_index=target_paragraph_index,
        position=position,
        content_type='paragraph',
        text=line_text,
    )
    return {}


def tool_insert_numbered_list_near_text(
    doc_id: str,
    target_text: Optional[str] = None,
    list_items: Optional[list] = None,
    position: str = 'after',
    target_paragraph_index: Optional[int] = None,
    bullet_type: str = 'bullet',
):
    kind = 'number' if (bullet_type or '').lower() == 'number' else 'bullet'
    _content.insert_near_text(
        doc_id=doc_id,
        target_text=target_text,
        target_paragraph_index=target_paragraph_index,
        position=position,
        content_type='list',
        items=list_items or [],
        kind=kind,
    )
    return {}


def tool_add_picture_base64_end(
    doc_id: str,
    image_base64: str,
    width_points: Optional[float] = None,
    height_points: Optional[float] = None,
    keep_aspect: bool = False,
):
    data = base64.b64decode(image_base64)
    _content.insert_image(
        doc_id=doc_id,
        image_bytes=data,
        where='end',
        width_points=width_points,
        height_points=height_points,
        keep_aspect=keep_aspect,
    )
    return {}


def tool_add_picture_base64_start(
    doc_id: str,
    image_base64: str,
    width_points: Optional[float] = None,
    height_points: Optional[float] = None,
    keep_aspect: bool = False,
):
    data = base64.b64decode(image_base64)
    _content.insert_image(
        doc_id=doc_id,
        image_bytes=data,
        where='start',
        width_points=width_points,
        height_points=height_points,
        keep_aspect=keep_aspect,
    )
    return {}


def tool_add_header_text(doc_id: str, text: str, primary: bool = True):
    _layout.header_add_text(doc_id, text, primary=bool(primary))
    return {}


def tool_add_footer_text(doc_id: str, text: str, primary: bool = True):
    _layout.footer_add_text(doc_id, text, primary=bool(primary))
    return {}


def tool_add_page_numbering(doc_id: str, format_string: str = 'Page {PAGE} of {NUMPAGES}'):
    _layout.add_page_numbering(doc_id, format_string=format_string)
    return {}


def tool_set_different_first_page_header_footer(doc_id: str, enabled: bool = True):
    _layout.set_different_first_page_header_footer(doc_id, enabled=bool(enabled))
    return {}


def tool_set_page_setup(
    doc_id: str,
    margins: Optional[dict] = None,
    orientation: Optional[str] = None,
    paper: Optional[str] = None,
    columns: Optional[int] = None,
):
    _layout.set_page_setup(
        doc_id, margins=margins, orientation=orientation, paper=paper, columns=columns
    )
    return {}


def tool_insert_section_break(doc_id: str, kind: str = 'nextPage'):
    _layout.insert_section_break(doc_id, kind=kind)
    return {}


def tool_insert_html_end(doc_id: str, html: str):
    _content.insert_html(doc_id=doc_id, html=html, where='end', paragraph_index=None)
    return {}


def tool_insert_html_start(doc_id: str, html: str):
    _content.insert_html(doc_id=doc_id, html=html, where='start', paragraph_index=None)
    return {}


def tool_insert_html_at_paragraph(doc_id: str, html: str, paragraph_index: int):
    _content.insert_html(
        doc_id=doc_id, html=html, where='paragraph', paragraph_index=paragraph_index
    )
    return {}


def tool_insert_markdown_end(doc_id: str, markdown: str):
    _content.insert_markdown(doc_id=doc_id, markdown=markdown, where='end', paragraph_index=None)
    return {}


def tool_insert_markdown_start(doc_id: str, markdown: str):
    _content.insert_markdown(doc_id=doc_id, markdown=markdown, where='start', paragraph_index=None)
    return {}


def tool_insert_markdown_at_paragraph(doc_id: str, markdown: str, paragraph_index: int):
    _content.insert_markdown(
        doc_id=doc_id, markdown=markdown, where='paragraph', paragraph_index=paragraph_index
    )
    return {}


def tool_add_watermark_text(
    doc_id: str,
    text: str,
    font_name: Optional[str] = None,
    size: Optional[float] = None,
    color: Optional[str] = None,
    diagonal: bool = True,
):
    _watermarks.add_watermark_text(
        doc_id, text, font_name=font_name, size=size, color_hex=color, diagonal=bool(diagonal)
    )
    return {}


def tool_add_watermark_image_base64(
    doc_id: str, image_base64: str, scale: Optional[float] = None, washout: bool = True
):
    data = base64.b64decode(image_base64)
    _watermarks.add_watermark_image(doc_id, data, scale=scale, washout=bool(washout))
    return {}


def tool_add_bookmark_at_paragraph(doc_id: str, name: str, paragraph_index: int):
    _links.add_bookmark_at_paragraph(doc_id, name, paragraph_index)
    return {}


def tool_insert_hyperlink_at_paragraph(
    doc_id: str, paragraph_index: int, text: str, target: str, external: bool = True
):
    _links.insert_hyperlink_at_paragraph(doc_id, paragraph_index, text, target, bool(external))
    return {}


def tool_render_page_base64(doc_id: str, page_index: int = 0, fmt: str = 'png', dpi: int = 150):
    data, mime, ext = _export.render_page(doc_id, page_index=page_index, fmt=fmt, dpi=dpi)
    b64 = base64.b64encode(data).decode('utf-8')
    return {'base64': b64, 'mime': mime, 'ext': ext}


def tool_export_base64_advanced(doc_id: str, fmt: str, options: Optional[dict] = None):
    data, mime, ext = _export.export_advanced(doc_id, fmt=fmt, options=options)
    b64 = base64.b64encode(data).decode('utf-8')
    return {'base64': b64, 'mime': mime, 'ext': ext}


def tool_format_text(
    doc_id: str,
    paragraph_index: int,
    start_pos: int,
    end_pos: int,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color: Optional[str] = None,
    font_size: Optional[float] = None,
    font_name: Optional[str] = None,
):
    _content.format_range(
        doc_id=doc_id,
        paragraph_index=paragraph_index,
        start=start_pos,
        end=end_pos,
        bold=bold,
        italic=italic,
        underline=underline,
        color_hex=color,
        font_name=font_name,
        font_size=font_size,
    )
    return {}


def tool_set_table_cell_shading(
    doc_id: str, table_index: int, row_index: int, col_index: int, fill_color: str
):
    _tables.table_set_cell_shading(
        doc_id, table_index, row_index, col_index, fill_color_hex=fill_color
    )
    return {}


def tool_apply_table_alternating_rows(
    doc_id: str, table_index: int, color1: str = 'FFFFFF', color2: str = 'F2F2F2'
):
    _tables.table_apply_alternating_rows(doc_id, table_index, color1_hex=color1, color2_hex=color2)
    return {}


def tool_highlight_table_header(
    doc_id: str, table_index: int, header_color: str = '4472C4', text_color: str = 'FFFFFF'
):
    _tables.table_highlight_header(
        doc_id, table_index, header_color_hex=header_color, text_color_hex=text_color
    )
    return {}


def tool_merge_table_cells(
    doc_id: str, table_index: int, start_row: int, start_col: int, end_row: int, end_col: int
):
    _tables.table_merge_cells(doc_id, table_index, start_row, start_col, end_row, end_col)
    return {}


def tool_merge_table_cells_horizontal(
    doc_id: str, table_index: int, row_index: int, start_col: int, end_col: int
):
    _tables.table_merge_cells(doc_id, table_index, row_index, start_col, row_index, end_col)
    return {}


def tool_merge_table_cells_vertical(
    doc_id: str, table_index: int, col_index: int, start_row: int, end_row: int
):
    _tables.table_merge_cells(doc_id, table_index, start_row, col_index, end_row, col_index)
    return {}


def tool_set_table_cell_alignment(
    doc_id: str,
    table_index: int,
    row_index: int,
    col_index: int,
    horizontal: str = 'left',
    vertical: str = 'top',
):
    _tables.table_set_cell_alignment(
        doc_id, table_index, row_index, col_index, horizontal=horizontal, vertical=vertical
    )
    return {}


def tool_set_table_alignment_all(doc_id: str, table_index: int, horizontal: str = 'left'):
    _tables.table_set_alignment(doc_id, table_index, horizontal=horizontal)
    return {}


def tool_set_table_column_width(
    doc_id: str, table_index: int, col_index: int, width: float, width_type: str = 'points'
):
    _tables.table_set_column_width(doc_id, table_index, col_index, width, unit=width_type)
    return {}


def tool_set_table_column_widths(
    doc_id: str, table_index: int, widths: list, width_type: str = 'points'
):
    if width_type and width_type != 'points':
        raise ValueError('Only width_type="points" is supported')
    _tables.table_set_column_widths(doc_id, table_index, widths)
    return {}


def tool_set_table_width(doc_id: str, table_index: int, width: float, width_type: str = 'points'):
    _tables.table_set_width(doc_id, table_index, width, unit=width_type)
    return {}


def tool_auto_fit_table_columns(doc_id: str, table_index: int):
    _tables.table_auto_fit_columns(doc_id, table_index)
    return {}


def tool_format_table_cell_text(
    doc_id: str,
    table_index: int,
    row_index: int,
    col_index: int,
    text_content: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color: Optional[str] = None,
    font_size: Optional[float] = None,
    font_name: Optional[str] = None,
):
    _tables.table_format_cell_text(
        doc_id,
        table_index,
        row_index,
        col_index,
        text_content=text_content,
        bold=bold,
        italic=italic,
        underline=underline,
        color_hex=color,
        font_size=font_size,
        font_name=font_name,
    )
    return {}


def tool_set_table_cell_padding(
    doc_id: str,
    table_index: int,
    row_index: int,
    col_index: int,
    top: Optional[float] = None,
    bottom: Optional[float] = None,
    left: Optional[float] = None,
    right: Optional[float] = None,
):
    _tables.table_set_cell_padding(
        doc_id, table_index, row_index, col_index, top=top, bottom=bottom, left=left, right=right
    )
    return {}


def tool_format_table(
    doc_id: str,
    table_index: int,
    has_header_row: Optional[bool] = None,
    border_style: Optional[str] = None,
    shading: Optional[list] = None,
):
    if border_style:
        _tables.table_set_borders(doc_id, table_index, line_style=border_style, line_width=None)
    if has_header_row:
        _tables.table_highlight_header(doc_id, table_index)
    if shading:
        for entry in shading:
            r, c, col = entry
            _tables.table_set_cell_shading(
                doc_id, table_index, int(r), int(c), fill_color_hex=str(col)
            )
    return {}


def tool_get_outline_simple(doc_id: str):
    return tool_get_outline(doc_id)


def tool_merge(doc_ids: list):
    new_id = _io.merge(doc_ids)
    first = doc_ids[0] if doc_ids else None
    base = document_store.get_document_name(first) if first else None
    merged_name = f'merged_{base}' if base else 'merged.docx'
    _store_add_mapping(new_id, merged_name)
    return {'docId': new_id}


def tool_save_as_new(doc_id: str, name: str, fmt: str = 'docx'):
    new_id, new_name = _io.save_as_new(doc_id, name, fmt=fmt)
    _store_add_mapping(new_id, new_name)
    return {'docId': new_id, 'name': new_name}


def tool_stats(doc_id: str):
    return _reading.stats(doc_id)


def tool_delete_document(doc_id: str):
    _io.delete(doc_id)
    document_store.remove_document(doc_id)
    return {}


def tool_list_documents():
    return {'docIds': _reading.list_documents()}


def tool_copy_document(doc_id: str):
    new_id = _io.copy(doc_id)
    src_name = document_store.get_document_name(doc_id)
    new_name = f'copy_of_{src_name}' if src_name else 'copy.docx'
    _store_add_mapping(new_id, new_name)
    return {'docId': new_id}


def tool_get_text(doc_id: str):
    return {'text': _reading.get_text(doc_id)}


def tool_get_xml(doc_id: str):
    return {'xml': _reading.get_xml(doc_id)}


def tool_get_document_base64(doc_id: str):
    data = _io.get_document_bytes(doc_id)
    return {'base64': base64.b64encode(data).decode('utf-8')}


def tool_health():
    return {'status': 'ok'}


def tool_properties_get(doc_id: str):
    return _properties.properties_get(doc_id)


def tool_properties_set(
    doc_id: str,
    title: Optional[str] = None,
    author: Optional[str] = None,
    subject: Optional[str] = None,
    keywords: Optional[str] = None,
):
    _properties.properties_set(
        doc_id, title=title, author=author, subject=subject, keywords=keywords
    )
    return _properties.properties_get(doc_id)


def tool_protect_document(doc_id: str, password: Optional[str] = None):
    _protection.protect(doc_id, password=password)
    return {}


def tool_unprotect_document(doc_id: str, password: Optional[str] = None):
    _protection.unprotect(doc_id, password=password)
    return {}


def tool_protect_restrict(
    doc_id: str, password: Optional[str] = None, ranges: Optional[list] = None
):
    _protection.protect_restrict(doc_id, password=password, ranges=ranges)
    return {}


def tool_get_all_comments(doc_id: str):
    return {'comments': _comments.get_comments(doc_id)}


def tool_get_comments_by_author(doc_id: str, author: str):
    return {'comments': _comments.get_comments_by_author(doc_id, author)}


def tool_get_comments_for_paragraph(doc_id: str, paragraph_index: int):
    return {'comments': _comments.get_comments_for_paragraph(doc_id, paragraph_index)}


def tool_add_footnote(doc_id: str, paragraph_index: int, text: str):
    _notes.notes_add_footnote(doc_id, paragraph_index, text)
    return {}


def tool_add_endnote(doc_id: str, paragraph_index: int, text: str):
    _notes.notes_add_endnote(doc_id, paragraph_index, text)
    return {}


def tool_convert_footnotes_to_endnotes(doc_id: str):
    _notes.notes_convert_footnotes_to_endnotes(doc_id)
    return {}


def tool_notes_style(
    doc_id: str,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color: Optional[str] = None,
):
    _notes.notes_style(
        doc_id, font_name=font_name, font_size=font_size, bold=bold, italic=italic, color_hex=color
    )
    return {}


def tool_get_all_notes(doc_id: str):
    return {'notes': _notes.notes_list(doc_id)}


def tool_add_footnote_by_anchor(
    doc_id: str,
    anchor_text: str,
    text: str,
    position: str = 'after',
    occurrence: int = 1,
    match_case: bool = False,
    whole_word: bool = False,
):
    _notes.notes_add_by_anchor(
        doc_id,
        note_type='footnote',
        anchor_text=anchor_text,
        text=text,
        position=position,
        occurrence=occurrence,
        match_case=match_case,
        whole_word=whole_word,
    )
    return {}


def tool_add_endnote_by_anchor(
    doc_id: str,
    anchor_text: str,
    text: str,
    position: str = 'after',
    occurrence: int = 1,
    match_case: bool = False,
    whole_word: bool = False,
):
    _notes.notes_add_by_anchor(
        doc_id,
        note_type='endnote',
        anchor_text=anchor_text,
        text=text,
        position=position,
        occurrence=occurrence,
        match_case=match_case,
        whole_word=whole_word,
    )
    return {}


def tool_delete_footnotes_by_anchor(
    doc_id: str,
    anchor_text: str,
    occurrence: Optional[int] = None,
    remove_all: bool = True,
    match_case: bool = False,
    whole_word: bool = False,
):
    count = _notes.notes_delete_by_anchor(
        doc_id,
        note_type='footnote',
        anchor_text=anchor_text,
        match_case=match_case,
        whole_word=whole_word,
        occurrence=occurrence,
        remove_all=remove_all,
    )
    return {'count': int(count)}


def tool_delete_endnotes_by_anchor(
    doc_id: str,
    anchor_text: str,
    occurrence: Optional[int] = None,
    remove_all: bool = True,
    match_case: bool = False,
    whole_word: bool = False,
):
    count = _notes.notes_delete_by_anchor(
        doc_id,
        note_type='endnote',
        anchor_text=anchor_text,
        match_case=match_case,
        whole_word=whole_word,
        occurrence=occurrence,
        remove_all=remove_all,
    )
    return {'count': int(count)}


def tool_validate_footnotes_by_anchor(
    doc_id: str,
    anchor_text: str,
    min_count: int = 1,
    match_case: bool = False,
    whole_word: bool = False,
):
    return _notes.notes_validate_by_anchor(
        doc_id,
        note_type='footnote',
        anchor_text=anchor_text,
        min_count=min_count,
        match_case=match_case,
        whole_word=whole_word,
    )


def tool_validate_endnotes_by_anchor(
    doc_id: str,
    anchor_text: str,
    min_count: int = 1,
    match_case: bool = False,
    whole_word: bool = False,
):
    return _notes.notes_validate_by_anchor(
        doc_id,
        note_type='endnote',
        anchor_text=anchor_text,
        min_count=min_count,
        match_case=match_case,
        whole_word=whole_word,
    )


def register_tools() -> None:

    @mcp.tool(description='Create a new document and return its ID')
    def create_document(name: Optional[str] = None):
        return tool_create_document(name)

    @mcp.tool(description='Insert text at the end of the document')
    def insert_text_end(doc_id: str, text: str):
        return tool_insert_text_end(doc_id, text)

    @mcp.tool(description='Insert text at the start of the document')
    def insert_text_start(doc_id: str, text: str):
        return tool_insert_text_start(doc_id, text)

    @mcp.tool(description='Insert text into the specified paragraph by index')
    def insert_text_at_paragraph(doc_id: str, text: str, paragraph_index: int):
        return tool_insert_text_at_paragraph(doc_id, text, paragraph_index)

    @mcp.tool(description='Read a range of paragraphs from the document')
    def read_paragraphs(doc_id: str, start: Optional[int] = None, end: Optional[int] = None):
        return tool_read_paragraphs(doc_id, start=start, end=end)

    @mcp.tool(description='Export the document to the specified format and return as base64')
    def export_base64(doc_id: str, fmt: str = 'docx'):
        return tool_export_base64(doc_id, fmt=fmt)

    @mcp.tool(description='Get general document information')
    def get_info(doc_id: str):
        return tool_get_info(doc_id)

    @mcp.tool(description='Add a heading with level and font options')
    def add_heading(
        doc_id: str,
        text: str,
        level: int = 1,
        font_name: Optional[str] = None,
        font_size: Optional[float] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        border_bottom: Optional[bool] = None,
    ):
        return tool_add_heading(
            doc_id,
            text,
            level=level,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            border_bottom=border_bottom,
        )

    @mcp.tool(description='Add a paragraph with optional style and formatting')
    def add_paragraph(
        doc_id: str,
        text: str,
        style: Optional[str] = None,
        font_name: Optional[str] = None,
        font_size: Optional[float] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[str] = None,
    ):
        return tool_add_paragraph(
            doc_id,
            text,
            style=style,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=color,
        )

    @mcp.tool(description='Insert a page break at the end of the document')
    def add_page_break_end(doc_id: str):
        return tool_add_page_break_end(doc_id)

    @mcp.tool(description='Insert a page break at the start of the document')
    def add_page_break_start(doc_id: str):
        return tool_add_page_break_start(doc_id)

    @mcp.tool(description='Insert a page break before the specified paragraph')
    def add_page_break_at_paragraph(doc_id: str, paragraph_index: int):
        return tool_add_page_break_at_paragraph(doc_id, paragraph_index)

    @mcp.tool(description='Get the document outline (headings structure)')
    def get_outline(doc_id: str):
        return tool_get_outline(doc_id)

    @mcp.tool(description='Replace text in the document with case/scope options')
    def replace_text(
        doc_id: str,
        search_text: str,
        replacement_text: str,
        replace_all: bool = True,
        case_sensitive: bool = False,
    ):
        return tool_replace_text(
            doc_id,
            search_text,
            replacement_text,
            replace_all=replace_all,
            case_sensitive=case_sensitive,
        )

    @mcp.tool(description='Find text occurrences in the document')
    def find_text(doc_id: str, text: str, match_case: bool = False, whole_word: bool = False):
        return tool_find_text(doc_id, text, match_case=match_case, whole_word=whole_word)

    @mcp.tool(description='Delete a paragraph by index')
    def delete_paragraph(doc_id: str, paragraph_index: int):
        return tool_delete_paragraph(doc_id, paragraph_index)

    @mcp.tool(description='Create a custom style with font properties')
    def create_custom_style(
        doc_id: str,
        style_name: str,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        font_size: Optional[float] = None,
        font_name: Optional[str] = None,
        color: Optional[str] = None,
        base_style: Optional[str] = None,
    ):
        return tool_create_style(
            doc_id,
            style_name,
            base_style=base_style,
            font_size=font_size,
            font_name=font_name,
            bold=bold,
            italic=italic,
            color=color,
        )

    @mcp.tool(description='Add a table at the end of the document')
    def add_table_end(
        doc_id: str, rows: int, cols: int, data: Optional[list] = None, has_header_row: bool = False
    ):
        return tool_add_table_end(doc_id, rows, cols, data=data, has_header_row=has_header_row)

    @mcp.tool(description='Add a table at the start of the document')
    def add_table_start(
        doc_id: str, rows: int, cols: int, data: Optional[list] = None, has_header_row: bool = False
    ):
        return tool_add_table_start(doc_id, rows, cols, data=data, has_header_row=has_header_row)

    @mcp.tool(description='Add a table at the specified paragraph')
    def add_table_at_paragraph(
        doc_id: str,
        rows: int,
        cols: int,
        paragraph_index: int,
        data: Optional[list] = None,
        has_header_row: bool = False,
    ):
        return tool_add_table_at_paragraph(
            doc_id, rows, cols, paragraph_index, data=data, has_header_row=has_header_row
        )

    @mcp.tool(description='Insert a bulleted or numbered list at the end')
    def insert_list_end(doc_id: str, items: list, kind: str = 'bullet'):
        return tool_insert_list_end(doc_id, items, kind=kind)

    @mcp.tool(description='Insert a bulleted or numbered list at the start')
    def insert_list_start(doc_id: str, items: list, kind: str = 'bullet'):
        return tool_insert_list_start(doc_id, items, kind=kind)

    @mcp.tool(description='Insert a list at the specified paragraph')
    def insert_list_at_paragraph(
        doc_id: str, items: list, paragraph_index: int, kind: str = 'bullet'
    ):
        return tool_insert_list_at_paragraph(doc_id, items, paragraph_index, kind=kind)

    @mcp.tool(description='Insert a heading near the found text or paragraph index')
    def insert_header_near_text(
        doc_id: str,
        target_text: Optional[str] = None,
        header_title: Optional[str] = None,
        position: str = 'after',
        level: int = 1,
        target_paragraph_index: Optional[int] = None,
    ):
        return tool_insert_header_near_text(
            doc_id, target_text, header_title, position, level, target_paragraph_index
        )

    @mcp.tool(description='Insert a line/paragraph near the found text')
    def insert_line_or_paragraph_near_text(
        doc_id: str,
        target_text: Optional[str] = None,
        line_text: Optional[str] = None,
        position: str = 'after',
        target_paragraph_index: Optional[int] = None,
    ):
        return tool_insert_line_or_paragraph_near_text(
            doc_id, target_text, line_text, position, target_paragraph_index
        )

    @mcp.tool(description='Create a numbered or bulleted list near text')
    def insert_numbered_list_near_text(
        doc_id: str,
        target_text: Optional[str] = None,
        list_items: Optional[list] = None,
        position: str = 'after',
        target_paragraph_index: Optional[int] = None,
        bullet_type: str = 'bullet',
    ):
        return tool_insert_numbered_list_near_text(
            doc_id, target_text, list_items, position, target_paragraph_index, bullet_type
        )

    @mcp.tool(description='Add an image (base64) at the end of the document')
    def add_picture_base64_end(
        doc_id: str,
        image_base64: str,
        width_points: Optional[float] = None,
        height_points: Optional[float] = None,
        keep_aspect: bool = False,
    ):
        return tool_add_picture_base64_end(
            doc_id,
            image_base64,
            width_points=width_points,
            height_points=height_points,
            keep_aspect=keep_aspect,
        )

    @mcp.tool(description='Add an image (base64) at the start of the document')
    def add_picture_base64_start(
        doc_id: str,
        image_base64: str,
        width_points: Optional[float] = None,
        height_points: Optional[float] = None,
        keep_aspect: bool = False,
    ):
        return tool_add_picture_base64_start(
            doc_id,
            image_base64,
            width_points=width_points,
            height_points=height_points,
            keep_aspect=keep_aspect,
        )

    @mcp.tool(description='Format text within a paragraph by character range')
    def format_text(
        doc_id: str,
        paragraph_index: int,
        start_pos: int,
        end_pos: int,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[bool] = None,
        color: Optional[str] = None,
        font_size: Optional[float] = None,
        font_name: Optional[str] = None,
    ):
        return tool_format_text(
            doc_id,
            paragraph_index,
            start_pos,
            end_pos,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            font_size=font_size,
            font_name=font_name,
        )

    @mcp.tool(description='Apply composite formatting to a table')
    def format_table(
        doc_id: str,
        table_index: int,
        has_header_row: bool = None,
        border_style: str = None,
        shading: list = None,
    ):
        return tool_format_table(
            doc_id,
            table_index,
            has_header_row=has_header_row,
            border_style=border_style,
            shading=shading,
        )

    @mcp.tool(description='Set table cell shading to the specified color')
    def set_table_cell_shading(
        doc_id: str, table_index: int, row_index: int, col_index: int, fill_color: str
    ):
        return tool_set_table_cell_shading(doc_id, table_index, row_index, col_index, fill_color)

    @mcp.tool(description='Apply alternating row colors to a table')
    def apply_table_alternating_rows(
        doc_id: str, table_index: int, color1: str = 'FFFFFF', color2: str = 'F2F2F2'
    ):
        return tool_apply_table_alternating_rows(doc_id, table_index, color1=color1, color2=color2)

    @mcp.tool(description='Highlight the table header row')
    def highlight_table_header(
        doc_id: str, table_index: int, header_color: str = '4472C4', text_color: str = 'FFFFFF'
    ):
        return tool_highlight_table_header(
            doc_id, table_index, header_color=header_color, text_color=text_color
        )

    @mcp.tool(description='Merge a range of table cells')
    def merge_table_cells(
        doc_id: str, table_index: int, start_row: int, start_col: int, end_row: int, end_col: int
    ):
        return tool_merge_table_cells(doc_id, table_index, start_row, start_col, end_row, end_col)

    @mcp.tool(description='Merge cells horizontally in a row')
    def merge_table_cells_horizontal(
        doc_id: str, table_index: int, row_index: int, start_col: int, end_col: int
    ):
        return tool_merge_table_cells_horizontal(doc_id, table_index, row_index, start_col, end_col)

    @mcp.tool(description='Merge cells vertically in a column')
    def merge_table_cells_vertical(
        doc_id: str, table_index: int, col_index: int, start_row: int, end_row: int
    ):
        return tool_merge_table_cells_vertical(doc_id, table_index, col_index, start_row, end_row)

    @mcp.tool(description='Set table cell horizontal/vertical alignment')
    def set_table_cell_alignment(
        doc_id: str,
        table_index: int,
        row_index: int,
        col_index: int,
        horizontal: str = 'left',
        vertical: str = 'top',
    ):
        return tool_set_table_cell_alignment(
            doc_id, table_index, row_index, col_index, horizontal=horizontal, vertical=vertical
        )

    @mcp.tool(description='Align the entire table relative to the page')
    def set_table_alignment_all(doc_id: str, table_index: int, horizontal: str = 'left'):
        return tool_set_table_alignment_all(doc_id, table_index, horizontal=horizontal)

    @mcp.tool(description='Set a table column width')
    def set_table_column_width(
        doc_id: str, table_index: int, col_index: int, width: float, width_type: str = 'points'
    ):
        return tool_set_table_column_width(
            doc_id, table_index, col_index, width, width_type=width_type
        )

    @mcp.tool(description='Set all table column widths')
    def set_table_column_widths(
        doc_id: str, table_index: int, widths: list, width_type: str = 'points'
    ):
        return tool_set_table_column_widths(doc_id, table_index, widths, width_type=width_type)

    @mcp.tool(description='Set the overall table width')
    def set_table_width(doc_id: str, table_index: int, width: float, width_type: str = 'points'):
        return tool_set_table_width(doc_id, table_index, width, width_type=width_type)

    @mcp.tool(description='Auto fit table columns to content')
    def auto_fit_table_columns(doc_id: str, table_index: int):
        return tool_auto_fit_table_columns(doc_id, table_index)

    @mcp.tool(description='Set text and formatting of a specific table cell')
    def format_table_cell_text(
        doc_id: str,
        table_index: int,
        row_index: int,
        col_index: int,
        text_content: Optional[str] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[bool] = None,
        color: Optional[str] = None,
        font_size: Optional[float] = None,
        font_name: Optional[str] = None,
    ):
        return tool_format_table_cell_text(
            doc_id,
            table_index,
            row_index,
            col_index,
            text_content=text_content,
            bold=bold,
            italic=italic,
            underline=underline,
            color=color,
            font_size=font_size,
            font_name=font_name,
        )

    @mcp.tool(description='Set table cell padding')
    def set_table_cell_padding(
        doc_id: str,
        table_index: int,
        row_index: int,
        col_index: int,
        top: Optional[float] = None,
        bottom: Optional[float] = None,
        left: Optional[float] = None,
        right: Optional[float] = None,
    ):
        return tool_set_table_cell_padding(
            doc_id,
            table_index,
            row_index,
            col_index,
            top=top,
            bottom=bottom,
            left=left,
            right=right,
        )

    @mcp.tool(description='Merge multiple documents into one')
    def merge_documents(source_doc_ids: list):
        return tool_merge(source_doc_ids)

    @mcp.tool(description='Save a copy of the document with a new name and format')
    def save_as_new(doc_id: str, name: str, fmt: str = 'docx'):
        return tool_save_as_new(doc_id, name, fmt=fmt)

    @mcp.tool(description='Get document statistics (pages, words, etc.)')
    def stats(doc_id: str):
        return tool_stats(doc_id)

    @mcp.tool(description='Delete a document from the store')
    def delete_document(doc_id: str):
        return tool_delete_document(doc_id)

    @mcp.tool(description='List all available documents')
    def list_documents():
        return tool_list_documents()

    @mcp.tool(description='Create a copy of the document and return the new ID')
    def copy_document(doc_id: str):
        return tool_copy_document(doc_id)

    @mcp.tool(description='Get the full document text as a single block')
    def get_text(doc_id: str):
        return tool_get_text(doc_id)

    @mcp.tool(description='Get the document XML')
    def get_xml(doc_id: str):
        return tool_get_xml(doc_id)

    @mcp.tool(description='Get the document binary content as base64')
    def get_document_base64(doc_id: str):
        return tool_get_document_base64(doc_id)

    @mcp.tool(description='Health check')
    def health():
        return tool_health()

    @mcp.tool(description='Read document properties (Title, Author, etc.)')
    def get_properties(doc_id: str):
        return tool_properties_get(doc_id)

    @mcp.tool(description='Set document properties (Title, Author, etc.)')
    def set_properties(
        doc_id: str,
        title: Optional[str] = None,
        author: Optional[str] = None,
        subject: Optional[str] = None,
        keywords: Optional[str] = None,
    ):
        return tool_properties_set(
            doc_id, title=title, author=author, subject=subject, keywords=keywords
        )

    @mcp.tool(description='Protect the document with a password')
    def protect_document(doc_id: str, password: Optional[str] = None):
        return tool_protect_document(doc_id, password=password)

    @mcp.tool(description='Unprotect the document')
    def unprotect_document(doc_id: str, password: Optional[str] = None):
        return tool_unprotect_document(doc_id, password=password)

    @mcp.tool(description='Restrict editing ranges in the document')
    def protect_restrict(
        doc_id: str, password: Optional[str] = None, ranges: Optional[list] = None
    ):
        return tool_protect_restrict(doc_id, password=password, ranges=ranges)

    @mcp.tool(description='Get all document comments')
    def get_all_comments(doc_id: str):
        return tool_get_all_comments(doc_id)

    @mcp.tool(description='Get comments by author')
    def get_comments_by_author(doc_id: str, author: str):
        return tool_get_comments_by_author(doc_id, author)

    @mcp.tool(description='Get comments for the specified paragraph')
    def get_comments_for_paragraph(doc_id: str, paragraph_index: int):
        return tool_get_comments_for_paragraph(doc_id, paragraph_index)

    @mcp.tool(description='Add a footnote to a paragraph')
    def add_footnote(doc_id: str, paragraph_index: int, text: str):
        return tool_add_footnote(doc_id, paragraph_index, text)

    @mcp.tool(description='Add an endnote to a paragraph')
    def add_endnote(doc_id: str, paragraph_index: int, text: str):
        return tool_add_endnote(doc_id, paragraph_index, text)

    @mcp.tool(description='Convert all footnotes to endnotes')
    def convert_footnotes_to_endnotes(doc_id: str):
        return tool_convert_footnotes_to_endnotes(doc_id)

    @mcp.tool(description='Set text style for footnotes/endnotes')
    def notes_style(
        doc_id: str,
        font_name: Optional[str] = None,
        font_size: Optional[float] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[str] = None,
    ):
        return tool_notes_style(
            doc_id, font_name=font_name, font_size=font_size, bold=bold, italic=italic, color=color
        )

    @mcp.tool(description='Get a list of all document notes')
    def get_all_notes(doc_id: str):
        return tool_get_all_notes(doc_id)

    @mcp.tool(description='Add a footnote after/before the anchor text')
    def add_footnote_by_anchor(
        doc_id: str,
        anchor_text: str,
        text: str,
        position: str = 'after',
        occurrence: int = 1,
        match_case: bool = False,
        whole_word: bool = False,
    ):
        return tool_add_footnote_by_anchor(
            doc_id,
            anchor_text,
            text,
            position=position,
            occurrence=occurrence,
            match_case=match_case,
            whole_word=whole_word,
        )

    @mcp.tool(description='Add an endnote after/before the anchor text')
    def add_endnote_by_anchor(
        doc_id: str,
        anchor_text: str,
        text: str,
        position: str = 'after',
        occurrence: int = 1,
        match_case: bool = False,
        whole_word: bool = False,
    ):
        return tool_add_endnote_by_anchor(
            doc_id,
            anchor_text,
            text,
            position=position,
            occurrence=occurrence,
            match_case=match_case,
            whole_word=whole_word,
        )

    @mcp.tool(description='Delete footnotes by anchor text (with filter options)')
    def delete_footnotes_by_anchor(
        doc_id: str,
        anchor_text: str,
        occurrence: Optional[int] = None,
        remove_all: bool = True,
        match_case: bool = False,
        whole_word: bool = False,
    ):
        return tool_delete_footnotes_by_anchor(
            doc_id,
            anchor_text,
            occurrence=occurrence,
            remove_all=remove_all,
            match_case=match_case,
            whole_word=whole_word,
        )

    @mcp.tool(description='Delete endnotes by anchor text (with filter options)')
    def delete_endnotes_by_anchor(
        doc_id: str,
        anchor_text: str,
        occurrence: Optional[int] = None,
        remove_all: bool = True,
        match_case: bool = False,
        whole_word: bool = False,
    ):
        return tool_delete_endnotes_by_anchor(
            doc_id,
            anchor_text,
            occurrence=occurrence,
            remove_all=remove_all,
            match_case=match_case,
            whole_word=whole_word,
        )

    @mcp.tool(
        description='Validate that at least the specified number of footnotes exist by anchor'
    )
    def validate_footnotes_by_anchor(
        doc_id: str,
        anchor_text: str,
        min_count: int = 1,
        match_case: bool = False,
        whole_word: bool = False,
    ):
        return tool_validate_footnotes_by_anchor(
            doc_id, anchor_text, min_count=min_count, match_case=match_case, whole_word=whole_word
        )

    @mcp.tool(description='Validate that at least the specified number of endnotes exist by anchor')
    def validate_endnotes_by_anchor(
        doc_id: str,
        anchor_text: str,
        min_count: int = 1,
        match_case: bool = False,
        whole_word: bool = False,
    ):
        return tool_validate_endnotes_by_anchor(
            doc_id, anchor_text, min_count=min_count, match_case=match_case, whole_word=whole_word
        )

    @mcp.tool(description='Add text to the header')
    def add_header_text(doc_id: str, text: str, primary: bool = True):
        return tool_add_header_text(doc_id, text, primary=primary)

    @mcp.tool(description='Add text to the footer')
    def add_footer_text(doc_id: str, text: str, primary: bool = True):
        return tool_add_footer_text(doc_id, text, primary=primary)

    @mcp.tool(description='Insert page numbering into headers/footers')
    def add_page_numbering(doc_id: str, format_string: str = 'Page {PAGE} of {NUMPAGES}'):
        return tool_add_page_numbering(doc_id, format_string=format_string)

    @mcp.tool(description='Use different first-page header/footer')
    def set_different_first_page_header_footer(doc_id: str, enabled: bool = True):
        return tool_set_different_first_page_header_footer(doc_id, enabled=enabled)

    @mcp.tool(description='Set page setup: margins, orientation, paper, columns')
    def set_page_setup(
        doc_id: str,
        margins: Optional[dict] = None,
        orientation: Optional[str] = None,
        paper: Optional[str] = None,
        columns: Optional[int] = None,
    ):
        return tool_set_page_setup(
            doc_id, margins=margins, orientation=orientation, paper=paper, columns=columns
        )

    @mcp.tool(description='Insert a section break of the specified type')
    def insert_section_break(doc_id: str, kind: str = 'nextPage'):
        return tool_insert_section_break(doc_id, kind=kind)

    @mcp.tool(description='Insert HTML at the end of the document')
    def insert_html_end(doc_id: str, html: str):
        return tool_insert_html_end(doc_id, html)

    @mcp.tool(description='Insert HTML at the start of the document')
    def insert_html_start(doc_id: str, html: str):
        return tool_insert_html_start(doc_id, html)

    @mcp.tool(description='Insert HTML at the specified paragraph')
    def insert_html_at_paragraph(doc_id: str, html: str, paragraph_index: int):
        return tool_insert_html_at_paragraph(doc_id, html, paragraph_index)

    @mcp.tool(description='Insert Markdown at the end of the document')
    def insert_markdown_end(doc_id: str, markdown: str):
        return tool_insert_markdown_end(doc_id, markdown)

    @mcp.tool(description='Insert Markdown at the start of the document')
    def insert_markdown_start(doc_id: str, markdown: str):
        return tool_insert_markdown_start(doc_id, markdown)

    @mcp.tool(description='Insert Markdown at the specified paragraph')
    def insert_markdown_at_paragraph(doc_id: str, markdown: str, paragraph_index: int):
        return tool_insert_markdown_at_paragraph(doc_id, markdown, paragraph_index)

    @mcp.tool(description='Add a text watermark')
    def add_watermark_text(
        doc_id: str,
        text: str,
        font_name: Optional[str] = None,
        size: Optional[float] = None,
        color: Optional[str] = None,
        diagonal: bool = True,
    ):
        return tool_add_watermark_text(
            doc_id, text, font_name=font_name, size=size, color=color, diagonal=diagonal
        )

    @mcp.tool(description='Add an image watermark from base64')
    def add_watermark_image_base64(
        doc_id: str, image_base64: str, scale: Optional[float] = None, washout: bool = True
    ):
        return tool_add_watermark_image_base64(doc_id, image_base64, scale=scale, washout=washout)

    @mcp.tool(description='Add a bookmark to the specified paragraph')
    def add_bookmark_at_paragraph(doc_id: str, name: str, paragraph_index: int):
        return tool_add_bookmark_at_paragraph(doc_id, name, paragraph_index)

    @mcp.tool(description='Insert a hyperlink at the specified paragraph')
    def insert_hyperlink_at_paragraph(
        doc_id: str, paragraph_index: int, text: str, target: str, external: bool = True
    ):
        return tool_insert_hyperlink_at_paragraph(doc_id, paragraph_index, text, target, external)

    @mcp.tool(description='Render a document page to an image/base64')
    def render_page_base64(doc_id: str, page_index: int = 0, fmt: str = 'png', dpi: int = 150):
        return tool_render_page_base64(doc_id, page_index=page_index, fmt=fmt, dpi=dpi)

    @mcp.tool(description='Advanced export with additional format options')
    def export_base64_advanced(doc_id: str, fmt: str, options: Optional[dict] = None):
        return tool_export_base64_advanced(doc_id, fmt=fmt, options=options)

    return None


def run_server(
    transport: Optional[str] = None,
    host: str = '0.0.0.0',
    port: int = 8080,
    path: str = '/mcp',
    license_path: Optional[str] = None,
):
    logger = _setup_logging()
    _ensure_data_dir_initialized()
    register_tools()
    _license.apply_license(license_path or os.getenv('ASPOSE_WORDS_LICENSE_PATH'))
    tr = (
        (transport or os.getenv('MCP_TRANSPORT') or os.getenv('TRANSPORT') or 'stdio')
        .strip()
        .lower()
    )
    host_env = os.getenv('MCP_HOST') or os.getenv('HOST') or host
    port_env = int(os.getenv('MCP_PORT') or os.getenv('PORT') or port)
    path_http_env = os.getenv('MCP_PATH') or path
    path_sse_env = os.getenv('MCP_SSE_PATH') or '/sse'
    logger.info('Starting Aspose.Words MCP (FastMCP)...')
    logger.info('Transport: %s', tr)
    if tr in {'streamable-http', 'sse'}:
        path_for_tr = path_sse_env if tr == 'sse' else path_http_env
        logger.info('Listening on http://%s:%s%s (%s)', host_env, port_env, path_for_tr, tr)
        mcp.run(transport=tr, host=host_env, port=port_env, path=path_for_tr)
    else:
        mcp.run(transport='stdio')


if __name__ == '__main__':
    run_server()
