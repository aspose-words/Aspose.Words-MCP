import pytest

pytest.importorskip('aspose.words')
import mcp_server as srv


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
