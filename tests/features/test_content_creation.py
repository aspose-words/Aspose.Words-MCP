import pytest

pytest.importorskip('aspose.words')
import aspose.words as aw

import mcp_server as srv
from core.utils import docs_util as _docs


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
    assert info['pages'] >= 2


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
    return (
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8AABQMB'
        'gYJ4V7wAAAAASUVORK5CYII='
    )


def _doc_with_text(name: str, text: str) -> str:
    doc_id = srv.tool_create_document(name)['docId']
    srv.tool_insert_text_end(doc_id, text)
    return doc_id


def _doc_with_adjacent_same_format_runs(name: str) -> tuple[str, str]:
    doc_id = srv.tool_create_document(name)['docId']
    doc_path = _docs.ensure_path(doc_id)
    document = aw.Document(str(doc_path))
    body = document.first_section.body
    body.remove_all_children()

    marker_text = 'join-marker:'
    paragraph = aw.Paragraph(document)
    paragraph.append_child(aw.Run(document, marker_text))
    paragraph.append_child(aw.Run(document, 'token'))
    body.append_child(paragraph)

    document.save(str(doc_path))
    return doc_id, marker_text


def _doc_with_spacing_difference_between_adjacent_runs(name: str) -> tuple[str, str]:
    doc_id = srv.tool_create_document(name)['docId']
    doc_path = _docs.ensure_path(doc_id)
    document = aw.Document(str(doc_path))
    body = document.first_section.body
    body.remove_all_children()

    marker_text = 'spacing-marker:'
    paragraph = aw.Paragraph(document)
    marker_run = aw.Run(document, marker_text)
    token_run = aw.Run(document, 'token')
    token_run.font.spacing = 2.0
    paragraph.append_child(marker_run)
    paragraph.append_child(token_run)
    body.append_child(paragraph)

    document.save(str(doc_path))
    return doc_id, marker_text


def _doc_with_bold_whitespace_trailing_run(name: str) -> tuple[str, str]:
    doc_id = srv.tool_create_document(name)['docId']
    doc_path = _docs.ensure_path(doc_id)
    document = aw.Document(str(doc_path))
    body = document.first_section.body
    body.remove_all_children()

    marker_text = 'insignificant-marker:'
    paragraph = aw.Paragraph(document)
    paragraph.append_child(aw.Run(document, marker_text))
    paragraph.append_child(aw.Run(document, 'token'))

    trailing_whitespace_run = aw.Run(document, ' ')
    trailing_whitespace_run.font.bold = True
    paragraph.append_child(trailing_whitespace_run)
    body.append_child(paragraph)

    document.save(str(doc_path))
    return doc_id, marker_text


def _paragraph_with_marker(doc_id: str, marker_text: str) -> aw.Paragraph:
    doc_path = _docs.ensure_path(doc_id)
    document = aw.Document(str(doc_path))
    paragraph_nodes = document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    for paragraph_index in range(paragraph_nodes.count):
        paragraph = paragraph_nodes[paragraph_index].as_paragraph()
        paragraph_text = paragraph.to_string(aw.SaveFormat.TEXT)
        if marker_text in paragraph_text:
            return paragraph
    raise AssertionError(f'Paragraph with marker not found: {marker_text}')


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


def test_replace_text_regex_plain_literal_and_zero_match_behavior():
    doc_id = _doc_with_text('replace-surface.docx', 'token cat cot cut')

    plain = srv.tool_replace_text(doc_id, find_text='token', replace_text='_', use_regex=False)
    assert plain['count'] == 1
    plain_text = srv.tool_get_text(doc_id)['text']
    assert 'token' not in plain_text
    assert '_' in plain_text

    regex = srv.tool_replace_text(
        doc_id,
        search_text='c.t',
        replacement_text='dog',
        use_regex=True,
    )
    assert regex['count'] == 3

    before = srv.tool_get_text(doc_id)['text']
    zero = srv.tool_replace_text(doc_id, find_text='zzz', replace_text='x')
    after = srv.tool_get_text(doc_id)['text']
    assert zero['count'] == 0
    assert after == before


def test_replace_text_default_omits_run_joining_for_direct_tool_flow():
    doc_id, marker_text = _doc_with_adjacent_same_format_runs('replace-default-join.docx')

    replacement_response = srv.tool_replace_text(doc_id, find_text='token', replace_text='TOKEN')

    assert set(replacement_response.keys()) == {'count'}
    assert replacement_response['count'] == 1

    target_paragraph = _paragraph_with_marker(doc_id, marker_text)
    assert target_paragraph.runs.count == 2
    assert target_paragraph.runs[0].text == marker_text
    assert target_paragraph.runs[1].text == 'TOKEN'


def test_replace_text_join_runs_merges_adjacent_same_format_runs_for_direct_tool_flow():
    doc_id, marker_text = _doc_with_adjacent_same_format_runs('replace-opt-in-join.docx')

    replacement_response = srv.tool_replace_text(
        doc_id,
        find_text='token',
        replace_text='TOKEN',
        join_runs=True,
    )

    assert set(replacement_response.keys()) == {'count'}
    assert replacement_response['count'] == 1

    target_paragraph = _paragraph_with_marker(doc_id, marker_text)
    assert target_paragraph.runs.count == 1
    assert target_paragraph.runs[0].text == f'{marker_text}TOKEN'


def test_replace_text_join_runs_zero_match_leaves_run_structure_unchanged_for_direct_tool_flow():
    doc_id, marker_text = _doc_with_adjacent_same_format_runs('replace-opt-in-join-zero-match.docx')

    before_paragraph = _paragraph_with_marker(doc_id, marker_text)
    before_run_count = before_paragraph.runs.count
    before_run_texts = [before_paragraph.runs[index].text for index in range(before_run_count)]

    replacement_response = srv.tool_replace_text(
        doc_id,
        find_text='absent-token',
        replace_text='TOKEN',
        join_runs=True,
    )

    assert set(replacement_response.keys()) == {'count'}
    assert replacement_response['count'] == 0

    after_paragraph = _paragraph_with_marker(doc_id, marker_text)
    after_run_count = after_paragraph.runs.count
    after_run_texts = [after_paragraph.runs[index].text for index in range(after_run_count)]
    assert after_run_count == before_run_count
    assert after_run_texts == before_run_texts


def test_replace_text_ignore_spacing_changes_join_behavior_for_direct_tool_flow():
    without_ignore_doc_id, without_ignore_marker = (
        _doc_with_spacing_difference_between_adjacent_runs('replace-join-spacing-off.docx')
    )
    without_ignore_response = srv.tool_replace_text(
        without_ignore_doc_id,
        find_text='token',
        replace_text='TOKEN',
        join_runs=True,
    )

    assert set(without_ignore_response.keys()) == {'count'}
    assert without_ignore_response['count'] == 1

    without_ignore_paragraph = _paragraph_with_marker(without_ignore_doc_id, without_ignore_marker)
    assert without_ignore_paragraph.runs.count == 2
    assert without_ignore_paragraph.runs[0].text == without_ignore_marker
    assert without_ignore_paragraph.runs[1].text == 'TOKEN'

    with_ignore_doc_id, with_ignore_marker = _doc_with_spacing_difference_between_adjacent_runs(
        'replace-join-spacing-on.docx'
    )
    with_ignore_response = srv.tool_replace_text(
        with_ignore_doc_id,
        find_text='token',
        replace_text='TOKEN',
        join_runs=True,
        ignore_spacing=True,
    )

    assert set(with_ignore_response.keys()) == {'count'}
    assert with_ignore_response['count'] == 1

    with_ignore_paragraph = _paragraph_with_marker(with_ignore_doc_id, with_ignore_marker)
    assert with_ignore_paragraph.runs.count == 1
    assert with_ignore_paragraph.runs[0].text == f'{with_ignore_marker}TOKEN'


def test_replace_text_ignore_insignificant_changes_join_behavior_for_direct_tool_flow():
    without_ignore_doc_id, without_ignore_marker = _doc_with_bold_whitespace_trailing_run(
        'replace-join-insignificant-off.docx'
    )
    without_ignore_response = srv.tool_replace_text(
        without_ignore_doc_id,
        find_text='token',
        replace_text='TOKEN',
        join_runs=True,
    )

    assert set(without_ignore_response.keys()) == {'count'}
    assert without_ignore_response['count'] == 1

    without_ignore_paragraph = _paragraph_with_marker(without_ignore_doc_id, without_ignore_marker)
    assert without_ignore_paragraph.runs.count == 2
    assert without_ignore_paragraph.runs[0].text == f'{without_ignore_marker}TOKEN'
    assert without_ignore_paragraph.runs[1].text == ' '

    with_ignore_doc_id, with_ignore_marker = _doc_with_bold_whitespace_trailing_run(
        'replace-join-insignificant-on.docx'
    )
    with_ignore_response = srv.tool_replace_text(
        with_ignore_doc_id,
        find_text='token',
        replace_text='TOKEN',
        join_runs=True,
        ignore_insignificant=True,
    )

    assert set(with_ignore_response.keys()) == {'count'}
    assert with_ignore_response['count'] == 1

    with_ignore_paragraph = _paragraph_with_marker(with_ignore_doc_id, with_ignore_marker)
    assert with_ignore_paragraph.runs.count == 1
    assert with_ignore_paragraph.runs[0].text == f'{with_ignore_marker}TOKEN '


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


def test_insert_numbered_list_can_remove_list_level_tab_stop():
    r = srv.tool_create_document('list-remove-tab-stop.docx')
    did = r['docId']

    srv.tool_insert_list_end(
        did,
        items=['Numbered list item 1', 'Numbered list item 2'],
        kind='number',
        remove_list_level_tab_stop=True,
    )

    paras = srv.tool_read_paragraphs(did)['paragraphs']
    joined = '\n'.join(paras)
    assert 'Numbered list item 1' in joined
    assert 'Numbered list item 2' in joined


def test_insert_numbered_list_near_text_can_remove_list_level_tab_stop():
    r = srv.tool_create_document('near-list-remove-tab-stop.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'Anchor')

    srv.tool_insert_numbered_list_near_text(
        did,
        target_text='Anchor',
        list_items=['Near numbered item'],
        position='after',
        bullet_type='number',
        remove_list_level_tab_stop=True,
    )

    paras = srv.tool_read_paragraphs(did)['paragraphs']
    assert any('Near numbered item' in paragraph for paragraph in paras)


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
