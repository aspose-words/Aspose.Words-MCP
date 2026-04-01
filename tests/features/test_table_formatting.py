import pytest

pytest.importorskip('aspose.words')
import aspose.words as aw

import mcp_server as srv
from core.utils import docs_util as _docs


def test_table_formatting_suite():
    r = srv.tool_create_document('tbl.docx')
    did = r['docId']
    t = srv.tool_add_table_end(
        did, 3, 3, data=[['H1', 'H2', 'H3'], ['A', 'B', 'C'], ['D', 'E', 'F']], has_header_row=True
    )
    tidx = t['tableIndex']
    srv.tool_format_table(
        did, tidx, has_header_row=True, border_style='single', shading=[[1, 1, 'FFEEEE']]
    )
    srv.tool_apply_table_alternating_rows(did, tidx, color1='FFFFFF', color2='F2F2F2')
    srv.tool_set_table_cell_shading(did, tidx, 2, 2, fill_color='EEEFFF')
    srv.tool_merge_table_cells_horizontal(did, tidx, row_index=1, start_col=0, end_col=1)
    srv.tool_merge_table_cells_vertical(did, tidx, col_index=2, start_row=1, end_row=2)
    srv.tool_set_table_cell_alignment(did, tidx, 0, 0, horizontal='center', vertical='middle')
    srv.tool_set_table_alignment_all(did, tidx, horizontal='center')
    srv.tool_set_table_column_widths(did, tidx, widths=[50.0, 60.0, 70.0])
    srv.tool_set_table_column_width(did, tidx, 0, 55.0)
    srv.tool_set_table_width(did, tidx, 300.0)
    srv.tool_auto_fit_table_columns(did, tidx)
    srv.tool_format_table_cell_text(did, tidx, 1, 0, text_content='X', bold=True)
    srv.tool_set_table_cell_padding(did, tidx, 1, 0, top=2.0, bottom=2.0, left=1.0, right=1.0)
    paras = srv.tool_read_paragraphs(did)['paragraphs']
    assert isinstance(paras, list) and len(paras) >= 1


def test_add_table_start_and_at_paragraph_and_merge_cells():
    r = srv.tool_create_document('tblpos.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'P0')
    srv.tool_add_paragraph(did, 'P1')
    t1 = srv.tool_add_table_start(did, 2, 2, data=[['a', 'b'], ['c', 'd']], has_header_row=False)
    assert isinstance(t1['tableIndex'], int)
    paras1 = srv.tool_read_paragraphs(did)['paragraphs']
    assert isinstance(paras1, list) and len(paras1) >= 1
    idx_para = 1 if len(paras1) > 1 else 0
    t2 = srv.tool_add_table_at_paragraph(
        did,
        2,
        3,
        paragraph_index=idx_para,
        data=[['x', 'y', 'z'], ['u', 'v', 'w']],
        has_header_row=False,
    )
    assert isinstance(t2['tableIndex'], int)
    srv.tool_merge_table_cells(
        did, table_index=t2['tableIndex'], start_row=0, start_col=0, end_row=0, end_col=1
    )
    paras2 = srv.tool_read_paragraphs(did)['paragraphs']
    assert isinstance(paras2, list) and len(paras2) >= 1


def _assert_row_border_members(row_borders, expected_line_style, expected_line_width):
    border_members = (
        row_borders.left,
        row_borders.right,
        row_borders.top,
        row_borders.bottom,
        row_borders.horizontal,
        row_borders.vertical,
    )
    for border_member in border_members:
        assert border_member.line_style == expected_line_style
        assert border_member.line_width == expected_line_width


def test_format_table_sets_all_explicit_border_members_for_each_row():
    created_document = srv.tool_create_document('tbl-borders-single.docx')
    doc_id = created_document['docId']
    added_table = srv.tool_add_table_end(
        doc_id, 3, 3, data=[['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h', 'i']]
    )
    table_index = added_table['tableIndex']

    srv.tool_format_table(doc_id, table_index, border_style='single')

    doc_path = _docs.ensure_path(doc_id)
    document = aw.Document(str(doc_path))
    table_nodes = document.get_child_nodes(aw.NodeType.TABLE, True)
    table_obj = table_nodes[table_index].as_table()

    for row_position in range(table_obj.rows.count):
        row_borders = table_obj.rows[row_position].row_format.borders
        _assert_row_border_members(row_borders, aw.LineStyle.SINGLE, 1.0)


def test_format_table_sets_solid_alias_for_all_explicit_border_members():
    created_document = srv.tool_create_document('tbl-borders-solid.docx')
    doc_id = created_document['docId']
    added_table = srv.tool_add_table_end(doc_id, 2, 2, data=[['h1', 'h2'], ['v1', 'v2']])
    table_index = added_table['tableIndex']

    srv.tool_format_table(doc_id, table_index, border_style='solid')

    doc_path = _docs.ensure_path(doc_id)
    document = aw.Document(str(doc_path))
    table_nodes = document.get_child_nodes(aw.NodeType.TABLE, True)
    table_obj = table_nodes[table_index].as_table()

    for row_position in range(table_obj.rows.count):
        row_borders = table_obj.rows[row_position].row_format.borders
        _assert_row_border_members(row_borders, aw.LineStyle.SINGLE, 1.0)
