from __future__ import annotations
from typing import Optional, List
import aspose.words as aw
from core.utils.docs_util import ensure_path, move_builder, hex_to_color

def add_table(doc_id: str, rows: int, cols: int, data: Optional[List[List[str]]]=None, has_header_row: bool=False, where: str='end', paragraph_index: Optional[int]=None) -> int:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    builder.start_table()
    for r in range(rows):
        for c in range(cols):
            builder.insert_cell()
            val = ''
            if data and r < len(data) and (c < len(data[r])):
                val = str(data[r][c])
            if has_header_row and r == 0:
                builder.font.bold = True
            else:
                builder.font.bold = False
            builder.write(val)
        builder.end_row()
    builder.end_table()
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    tbl_index = max(0, tables.count - 1)
    doc.save(str(path))
    return int(tbl_index)

def table_set_cell_shading(doc_id: str, table_index: int, row_index: int, col_index: int, fill_color_hex: str) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    if row_index < 0 or row_index >= table_obj.rows.count:
        raise IndexError('row_index out of range')
    row = table_obj.rows[row_index]
    if col_index < 0 or col_index >= row.cells.count:
        raise IndexError('col_index out of range')
    cell = row.cells[col_index]
    col = hex_to_color(fill_color_hex)
    if col is not None:
        cell.cell_format.shading.background_pattern_color = col
    doc.save(str(path))
    return True

def table_apply_alternating_rows(doc_id: str, table_index: int, color1_hex: str='FFFFFF', color2_hex: str='F2F2F2') -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    c1 = hex_to_color(color1_hex)
    c2 = hex_to_color(color2_hex)
    for r_i in range(table_obj.rows.count):
        row = table_obj.rows[r_i]
        color = c1 if r_i % 2 == 0 else c2
        if color is None:
            continue
        for c in range(row.cells.count):
            row.cells[c].cell_format.shading.background_pattern_color = color
    doc.save(str(path))
    return True

def table_highlight_header(doc_id: str, table_index: int, header_color_hex: str='4472C4', text_color_hex: str='FFFFFF') -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    if table_obj.rows.count > 0:
        header_row = table_obj.rows[0]
        bg = hex_to_color(header_color_hex)
        fg = hex_to_color(text_color_hex)
        for c in range(header_row.cells.count):
            if bg is not None:
                header_row.cells[c].cell_format.shading.background_pattern_color = bg
            if fg is not None and header_row.cells[c].first_paragraph.runs.count > 0:
                header_row.cells[c].first_paragraph.runs[0].font.color = fg
    doc.save(str(path))
    return True

def table_set_cell_padding(doc_id: str, table_index: int, row_index: int, col_index: int, top: Optional[float]=None, bottom: Optional[float]=None, left: Optional[float]=None, right: Optional[float]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    if row_index < 0 or row_index >= table_obj.rows.count:
        raise IndexError('row_index out of range')
    row = table_obj.rows[row_index]
    if col_index < 0 or col_index >= row.cells.count:
        raise IndexError('col_index out of range')
    cell = row.cells[col_index]
    cf = cell.cell_format
    if top is not None:
        cf.top_padding = float(top)
    if bottom is not None:
        cf.bottom_padding = float(bottom)
    if left is not None:
        cf.left_padding = float(left)
    if right is not None:
        cf.right_padding = float(right)
    doc.save(str(path))
    return True

def table_set_column_widths(doc_id: str, table_index: int, widths: List[float]) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    if table_obj.rows.count > 0:
        row0 = table_obj.rows[0]
        for idx in range(min(len(widths), row0.cells.count)):
            row0.cells[idx].cell_format.width = float(widths[idx])
    doc.save(str(path))
    return True

def table_merge_cells(doc_id: str, table_index: int, start_row: int, start_col: int, end_row: int, end_col: int) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    for r in range(start_row, end_row + 1):
        if r < 0 or r >= table_obj.rows.count:
            continue
        row = table_obj.rows[r]
        for c in range(start_col, end_col + 1):
            if c < 0 or c >= row.cells.count:
                continue
            cell = row.cells[c]
            if cell is None:
                continue
            if r == start_row:
                cell.cell_format.vertical_merge = aw.tables.CellMerge.FIRST
            else:
                cell.cell_format.vertical_merge = aw.tables.CellMerge.PREVIOUS
    for r in range(start_row, end_row + 1):
        if r < 0 or r >= table_obj.rows.count:
            continue
        row = table_obj.rows[r]
        if start_col < 0 or start_col >= row.cells.count:
            continue
        if row.cells[start_col] is not None:
            row.cells[start_col].cell_format.horizontal_merge = aw.tables.CellMerge.FIRST
        for c in range(start_col + 1, end_col + 1):
            if c < 0 or c >= row.cells.count:
                continue
            if row.cells[c] is None:
                continue
            row.cells[c].cell_format.horizontal_merge = aw.tables.CellMerge.PREVIOUS
    doc.save(str(path))
    return True

def table_set_cell_alignment(doc_id: str, table_index: int, row_index: int, col_index: int, horizontal: Optional[str]=None, vertical: Optional[str]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    if row_index < 0 or row_index >= table_obj.rows.count:
        raise IndexError('row_index out of range')
    row = table_obj.rows[row_index]
    if col_index < 0 or col_index >= row.cells.count:
        raise IndexError('col_index out of range')
    cell = row.cells[col_index]
    if horizontal:
        h = horizontal.lower()
        if h.startswith('center'):
            pa = aw.ParagraphAlignment.CENTER
        elif h.startswith('right'):
            pa = aw.ParagraphAlignment.RIGHT
        else:
            pa = aw.ParagraphAlignment.LEFT
        for i in range(cell.paragraphs.count):
            p = cell.paragraphs[i]
            p.paragraph_format.alignment = pa
    if vertical:
        v = vertical.lower()
        if v.startswith('middle') or v.startswith('center'):
            cell.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER
        elif v.startswith('bottom'):
            cell.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.BOTTOM
        else:
            cell.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.TOP
    doc.save(str(path))
    return True

def table_set_alignment(doc_id: str, table_index: int, horizontal: Optional[str]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    if horizontal:
        h = horizontal.lower()
        if h.startswith('center'):
            table_obj.alignment = aw.tables.TableAlignment.CENTER
        elif h.startswith('right'):
            table_obj.alignment = aw.tables.TableAlignment.RIGHT
        else:
            table_obj.alignment = aw.tables.TableAlignment.LEFT
    doc.save(str(path))
    return True

def table_set_borders(doc_id: str, table_index: int, line_style: Optional[str]=None, line_width: Optional[float]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    style = (line_style or 'single').lower()
    lw = float(line_width or 1.0)
    if style in ('single', 'solid'):
        ls = aw.LineStyle.SINGLE
    elif style in ('dashed',):
        ls = aw.LineStyle.DASHED
    else:
        ls = aw.LineStyle.SINGLE
    for r in range(table_obj.rows.count):
        bc = table_obj.rows[r].row_format.borders
        for s in ('left', 'right', 'top', 'bottom', 'horizontal', 'vertical'):
            b = getattr(bc, s)
            b.line_style = ls
            b.line_width = lw
    doc.save(str(path))
    return True

def table_set_style(doc_id: str, table_index: int, style_name: str) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    table_obj.style = doc.styles[style_name]
    doc.save(str(path))
    return True

def table_format_cell_text(doc_id: str, table_index: int, row_index: int, col_index: int, text_content: Optional[str]=None, bold: Optional[bool]=None, italic: Optional[bool]=None, underline: Optional[bool]=None, color_hex: Optional[str]=None, font_size: Optional[float]=None, font_name: Optional[str]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    if row_index < 0 or row_index >= table_obj.rows.count:
        raise IndexError('row_index out of range')
    row = table_obj.rows[row_index]
    if col_index < 0 or col_index >= row.cells.count:
        raise IndexError('col_index out of range')
    cell = row.cells[col_index]
    if text_content is not None:
        while cell.first_paragraph.runs.count > 0:
            cell.first_paragraph.runs[0].remove()
        builder = aw.DocumentBuilder(doc)
        builder.move_to(cell.first_paragraph)
        builder.write(text_content)
    font = cell.first_paragraph.runs[0].font
    if font_name:
        font.name = font_name
    if font_size:
        font.size = font_size
    if bold is not None:
        font.bold = bold
    if italic is not None:
        font.italic = italic
    if underline is not None:
        font.underline = aw.Underline.SINGLE if underline else aw.Underline.NONE
    col = hex_to_color(color_hex)
    if col is not None:
        font.color = col
    doc.save(str(path))
    return True

def table_set_column_width(doc_id: str, table_index: int, col_index: int, width: float, unit: str='points') -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    if table_obj.rows.count == 0:
        return True
    row0 = table_obj.rows[0]
    if col_index < 0 or col_index >= row0.cells.count:
        raise IndexError('col_index out of range')
    w = float(width)
    if unit and unit.lower().startswith('percent'):
        page_w = page_width_points(doc)
        w = max(0.0, min(page_w, page_w * (float(width) / 100.0)))
    row0.cells[col_index].cell_format.width = w
    doc.save(str(path))
    return True

def table_set_width(doc_id: str, table_index: int, width: float, unit: str='points') -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    w = float(width)
    if unit and unit.lower().startswith('percent'):
        page_w = page_width_points(doc)
        w = max(0.0, min(page_w, page_w * (float(width) / 100.0)))
    table_obj.preferred_width = aw.tables.PreferredWidth.from_points(w)
    doc.save(str(path))
    return True

def table_auto_fit_columns(doc_id: str, table_index: int) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    tables = doc.get_child_nodes(aw.NodeType.TABLE, True)
    if table_index < 0 or table_index >= tables.count:
        raise IndexError('table_index out of range')
    table_obj = tables[table_index].as_table()
    table_obj.auto_fit(aw.tables.AutoFitBehavior.AUTO_FIT_TO_CONTENTS)
    doc.save(str(path))
    return True

def page_width_points(doc: aw.Document) -> float:
    return float(doc.first_section.page_setup.page_width)
