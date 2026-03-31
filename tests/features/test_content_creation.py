import pytest

pytest.importorskip('aspose.words')
import mcp_server as srv


def test_add_headings_and_get_outline():
    res = srv.tool_create_document('outline.docx')
    doc_id = res['docId']
    srv.tool_add_heading(doc_id, 'Section 1', level=1)
    srv.tool_add_paragraph(doc_id, 'Body under section 1')
    srv.tool_add_heading(doc_id, 'Subsection 1.1', level=2)
    srv.tool_add_paragraph(doc_id, 'Details')
    srv.tool_add_heading(doc_id, 'Section 2', level=1)
    outline = srv.tool_get_outline(doc_id)['outline']
    assert isinstance(outline, list)
    texts = [i['text'] for i in outline]
    for expected in ['Section 1', 'Subsection 1.1', 'Section 2']:
        assert expected in texts
    levels = {i['level'] for i in outline}
    assert 1 in levels
    assert 2 in levels


def test_read_paragraphs_and_add_paragraph_via_tools():
    res = srv.tool_create_document('content.docx')
    doc_id = res['docId']
    srv.tool_add_paragraph(doc_id, 'First paragraph')
    srv.tool_add_heading(doc_id, 'Header', level=1)
    srv.tool_insert_text_end(doc_id, ' Second paragraph')
    out = srv.tool_read_paragraphs(doc_id)
    paras = out['paragraphs']
    assert isinstance(paras, list)
    assert any(('First paragraph' in p for p in paras))
    assert any(('Header' in p for p in paras))
    out2 = srv.tool_read_paragraphs(doc_id, start=0, end=max(1, len(paras) - 1))
    assert isinstance(out2['paragraphs'], list)


def test_page_break_positions_via_manager():
    res = srv.tool_create_document('breaks.docx')
    doc_id = res['docId']
    srv.tool_add_page_break_end(doc_id)
    srv.tool_add_page_break_start(doc_id)
    srv.tool_add_page_break_at_paragraph(doc_id, paragraph_index=0)
    info = srv.tool_get_info(doc_id)
    assert info['paragraphs'] >= 1


def test_add_page_break_via_tool():
    res = srv.tool_create_document('break-tool.docx')
    doc_id = res['docId']
    srv.tool_add_paragraph(doc_id, 'Intro')
    out_before = srv.tool_read_paragraphs(doc_id)
    assert len(out_before['paragraphs']) >= 1
    srv.tool_add_page_break_end(doc_id)
    out_after = srv.tool_read_paragraphs(doc_id)
    assert len(out_after['paragraphs']) >= len(out_before['paragraphs'])


def _png_1x1_b64():
    return 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8AABQMBgYJ4V7wAAAAASUVORK5CYII='


def test_near_text_insertions_and_lists():
    res = srv.tool_create_document('near.docx')
    did = res['docId']
    srv.tool_add_paragraph(did, 'Hello')
    srv.tool_insert_header_near_text(
        did, target_text='Hello', header_title='H1', position='after', level=1
    )
    srv.tool_insert_line_or_paragraph_near_text(
        did, target_text='Hello', line_text='Line', position='after'
    )
    srv.tool_insert_numbered_list_near_text(
        did, target_text='Hello', list_items=['A', 'B'], position='after', bullet_type='number'
    )
    srv.tool_insert_list_end(did, items=['I1', 'I2'], kind='bullet')
    paras = srv.tool_read_paragraphs(did)['paragraphs']
    joined = '\n'.join(paras)
    for s in ('H1', 'Line', 'A', 'B', 'I1', 'I2'):
        assert s in joined


def test_add_picture_table_and_text_formatting():
    res = srv.tool_create_document('media.docx')
    did = res['docId']
    srv.tool_add_paragraph(did, 'abcdef')
    srv.tool_add_picture_base64_end(did, _png_1x1_b64(), keep_aspect=True)
    t = srv.tool_add_table_end(did, 2, 2, data=[['h1', 'h2'], ['v1', 'v2']], has_header_row=True)
    assert isinstance(t['tableIndex'], int)
    style = srv.tool_create_style(did, 'MyStyle', font_name='Times New Roman', bold=True)
    assert style['style'] == 'MyStyle'
    matches = srv.tool_find_text(did, text='Hello')
    assert isinstance(matches['matches'], list)
    paras = srv.tool_read_paragraphs(did)['paragraphs']
    pidx = max(0, len(paras) - 1)
    srv.tool_format_text(did, paragraph_index=pidx, start_pos=0, end_pos=min(3, len(paras[pidx])))
    rep = srv.tool_replace_text(did, find_text='Hello', replace_text='Hi')
    assert isinstance(rep['count'], int)


def test_insert_text_start_and_at_paragraph_and_delete():
    r = srv.tool_create_document('textpos.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'P0')
    srv.tool_add_paragraph(did, 'P1')
    srv.tool_insert_text_start(did, 'START ')
    paras = srv.tool_read_paragraphs(did)['paragraphs']
    assert any((p.strip().startswith('START') for p in paras))
    idx = max(0, len(paras) - 1)
    srv.tool_insert_text_at_paragraph(did, text=' END', paragraph_index=idx)
    paras2 = srv.tool_read_paragraphs(did)['paragraphs']
    assert paras2[idx].strip().endswith('END')
    before = srv.tool_read_paragraphs(did)['paragraphs']
    del_idx = 0
    for i, p in enumerate(before):
        if 'P0' in p:
            del_idx = i
            break
    srv.tool_delete_paragraph(did, paragraph_index=del_idx)
    after = srv.tool_read_paragraphs(did)['paragraphs']
    assert len(after) <= len(before)
    joined = '\n'.join(after)
    assert 'P0' not in joined


def test_insert_lists_start_and_at_paragraph():
    r = srv.tool_create_document('lists.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'A')
    srv.tool_add_paragraph(did, 'B')
    srv.tool_insert_list_start(did, items=['S1', 'S2'], kind='bullet')
    paras = srv.tool_read_paragraphs(did)['paragraphs']
    assert any(('S1' in p for p in paras))
    target_idx = max(0, len(paras) - 1)
    srv.tool_insert_list_at_paragraph(
        did, items=['E1', 'E2'], paragraph_index=target_idx, kind='number'
    )
    paras2 = srv.tool_read_paragraphs(did)['paragraphs']
    j = '\n'.join(paras2)
    for s in ('E1', 'E2'):
        assert s in j


def test_add_picture_start_and_outline_simple_and_stats():
    r = srv.tool_create_document('startpic.docx')
    did = r['docId']
    srv.tool_add_heading(did, 'H', level=1)
    srv.tool_add_paragraph(did, 'T')
    srv.tool_add_picture_base64_start(did, _png_1x1_b64(), keep_aspect=True)
    out = srv.tool_export_base64(did, fmt='docx')
    assert isinstance(out['base64'], str) and len(out['base64']) > 0
    simple = srv.tool_get_outline_simple(did)['outline']
    assert any((i['text'] == 'H' and i['level'] == 1 for i in simple))
    st = srv.tool_stats(did)
    assert isinstance(st, dict)
    for k in ('words', 'paragraphs', 'pages'):
        assert k in st and isinstance(st[k], int)
