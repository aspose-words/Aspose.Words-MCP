import json

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


def _paragraph_with_exact_text(doc_id: str, expected_text: str) -> aw.Paragraph:
    doc_path = _docs.ensure_path(doc_id)
    document = aw.Document(str(doc_path))
    paragraph_nodes = document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    for paragraph_index in range(paragraph_nodes.count):
        paragraph = paragraph_nodes[paragraph_index].as_paragraph()
        paragraph_text = paragraph.to_string(aw.SaveFormat.TEXT).strip()
        if paragraph_text == expected_text:
            return paragraph
    raise AssertionError(f'Paragraph with text not found: {expected_text}')


_ASPOSE_EVALUATION_EXACT_PARAGRAPHS = {
    'evaluation only. created with aspose.words. copyright 2003-2026 aspose pty ltd.',
    'this document was truncated here because it was created in the evaluation mode.',
}
_ASPOSE_EVALUATION_FOOTER_PREFIX = (
    'created with an evaluation copy of aspose.words. to remove all limitations, '
    'you can use free temporary license '
)
_ASPOSE_EVALUATION_FOOTER_SUFFIX = 'https://products.aspose.com/words/temporary-license/'


def _is_evaluation_only_paragraph(paragraph_text: str) -> bool:
    normalized_text = paragraph_text.rstrip('\r\n').strip().lower()
    if normalized_text == '':
        return False
    if normalized_text in _ASPOSE_EVALUATION_EXACT_PARAGRAPHS:
        return True
    return normalized_text.startswith(
        _ASPOSE_EVALUATION_FOOTER_PREFIX
    ) and normalized_text.endswith(_ASPOSE_EVALUATION_FOOTER_SUFFIX)


def _user_paragraph_ordinal_with_exact_text(doc_id: str, expected_text: str) -> int:
    expected_normalized_text = expected_text.rstrip('\r\n')
    doc_path = _docs.ensure_path(doc_id)
    document = aw.Document(str(doc_path))
    paragraph_nodes = document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    matching_ordinals: list[int] = []
    user_paragraph_ordinal = 0
    for paragraph_index in range(paragraph_nodes.count):
        paragraph = paragraph_nodes[paragraph_index].as_paragraph()
        paragraph_text = paragraph.to_string(aw.SaveFormat.TEXT) or ''
        if _is_evaluation_only_paragraph(paragraph_text):
            continue
        if paragraph_text.rstrip('\r\n') == expected_normalized_text:
            matching_ordinals.append(user_paragraph_ordinal)
        user_paragraph_ordinal += 1
    if matching_ordinals:
        if expected_normalized_text == '' and len(matching_ordinals) > 1:
            return matching_ordinals[-2]
        return matching_ordinals[-1]
    raise AssertionError(f'Paragraph with text not found: {expected_text}')


def _custom_node_id_sidecar_payload(doc_id: str) -> dict[str, object] | None:
    doc_path = _docs.ensure_path(doc_id)
    sidecar_path = doc_path.with_suffix(f'{doc_path.suffix}.custom_node_ids.json')
    if not sidecar_path.exists():
        return None
    return json.loads(sidecar_path.read_text(encoding='utf-8'))


def _assert_single_sidecar_custom_node_entry(
    doc_id: str, expected_text: str, expected_custom_node_id: int
) -> None:
    sidecar_payload = _custom_node_id_sidecar_payload(doc_id)
    assert sidecar_payload is not None
    assert sidecar_payload.get('doc_id') == doc_id
    paragraph_custom_node_ids = sidecar_payload.get('paragraph_custom_node_ids')
    assert isinstance(paragraph_custom_node_ids, dict)
    assert len(paragraph_custom_node_ids) == 1

    expected_user_paragraph_ordinal = _user_paragraph_ordinal_with_exact_text(doc_id, expected_text)
    sidecar_record = paragraph_custom_node_ids.get(str(expected_user_paragraph_ordinal))
    assert isinstance(sidecar_record, dict)
    assert sidecar_record.get('custom_node_id') == expected_custom_node_id
    assert sidecar_record.get('user_paragraph_ordinal') == expected_user_paragraph_ordinal


def _assert_single_sidecar_custom_node_entry_with_empty_text(
    doc_id: str, expected_custom_node_id: int
) -> None:
    sidecar_payload = _custom_node_id_sidecar_payload(doc_id)
    assert sidecar_payload is not None
    assert sidecar_payload.get('doc_id') == doc_id
    paragraph_custom_node_ids = sidecar_payload.get('paragraph_custom_node_ids')
    assert isinstance(paragraph_custom_node_ids, dict)
    assert len(paragraph_custom_node_ids) == 1

    _, sidecar_record = next(iter(paragraph_custom_node_ids.items()))
    assert isinstance(sidecar_record, dict)
    assert sidecar_record.get('custom_node_id') == expected_custom_node_id
    user_paragraph_ordinal = sidecar_record.get('user_paragraph_ordinal')
    assert isinstance(user_paragraph_ordinal, int)


def _paragraph_counts(doc_id: str) -> tuple[int, int]:
    doc_path = _docs.ensure_path(doc_id)
    document = aw.Document(str(doc_path))
    paragraph_nodes = document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    blank_paragraph_count = 0
    for paragraph_index in range(paragraph_nodes.count):
        paragraph = paragraph_nodes[paragraph_index].as_paragraph()
        if paragraph.to_string(aw.SaveFormat.TEXT).strip() == '':
            blank_paragraph_count += 1
    return paragraph_nodes.count, blank_paragraph_count


def test_add_heading_with_custom_node_id_applies_id_to_written_heading_paragraph():
    doc_id = srv.tool_create_document('heading-custom-node-id.docx')['docId']

    add_heading_response = srv.tool_add_heading(
        doc_id,
        'Heading with custom node id',
        level=2,
        custom_node_id=1101,
    )

    assert add_heading_response == {}
    _assert_single_sidecar_custom_node_entry(
        doc_id,
        'Heading with custom node id',
        expected_custom_node_id=1101,
    )


def test_add_paragraph_with_custom_node_id_applies_id_without_extra_blank_paragraphs():
    with_custom_node_id_doc_id = srv.tool_create_document('paragraph-custom-node-id.docx')['docId']
    without_custom_node_id_doc_id = srv.tool_create_document('paragraph-default-node-id.docx')[
        'docId'
    ]

    add_paragraph_response = srv.tool_add_paragraph(
        with_custom_node_id_doc_id,
        'Paragraph with custom node id',
        custom_node_id=2202,
    )
    srv.tool_add_paragraph(without_custom_node_id_doc_id, 'Paragraph with default node id')

    assert add_paragraph_response == {}
    _assert_single_sidecar_custom_node_entry(
        with_custom_node_id_doc_id,
        'Paragraph with custom node id',
        expected_custom_node_id=2202,
    )

    custom_path_counts = _paragraph_counts(with_custom_node_id_doc_id)
    default_path_counts = _paragraph_counts(without_custom_node_id_doc_id)
    assert custom_path_counts == default_path_counts


def test_add_heading_with_empty_text_custom_node_id_indexes_written_blank_paragraph():
    doc_id = srv.tool_create_document('heading-empty-custom-node-id.docx')['docId']

    add_heading_response = srv.tool_add_heading(doc_id, '', level=1, custom_node_id=3303)

    assert add_heading_response == {}
    _assert_single_sidecar_custom_node_entry_with_empty_text(doc_id, expected_custom_node_id=3303)


def test_add_paragraph_with_empty_text_custom_node_id_indexes_written_blank_paragraph():
    doc_id = srv.tool_create_document('paragraph-empty-custom-node-id.docx')['docId']

    add_paragraph_response = srv.tool_add_paragraph(doc_id, '', custom_node_id=4404)

    assert add_paragraph_response == {}
    _assert_single_sidecar_custom_node_entry_with_empty_text(doc_id, expected_custom_node_id=4404)


def test_add_paragraph_with_evaluation_marker_phrase_preserves_custom_node_id_sidecar_entry():
    doc_id = srv.tool_create_document('paragraph-evaluation-marker-phrase.docx')['docId']
    paragraph_text = (
        'User-authored note: created with an evaluation copy of aspose.words '
        'is included here as quoted content.'
    )

    add_paragraph_response = srv.tool_add_paragraph(doc_id, paragraph_text, custom_node_id=5505)

    assert add_paragraph_response == {}
    sidecar_payload = _custom_node_id_sidecar_payload(doc_id)
    assert sidecar_payload is not None
    paragraph_custom_node_ids = sidecar_payload.get('paragraph_custom_node_ids')
    assert isinstance(paragraph_custom_node_ids, dict)
    assert len(paragraph_custom_node_ids) == 1

    _, sidecar_record = next(iter(paragraph_custom_node_ids.items()))
    assert isinstance(sidecar_record, dict)
    assert sidecar_record.get('custom_node_id') == 5505
    user_paragraph_ordinal = sidecar_record.get('user_paragraph_ordinal')
    assert isinstance(user_paragraph_ordinal, int)

    doc_path = _docs.ensure_path(doc_id)
    document = aw.Document(str(doc_path))
    paragraph_nodes = document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    user_paragraph_texts: list[str] = []
    for paragraph_index in range(paragraph_nodes.count):
        paragraph = paragraph_nodes[paragraph_index].as_paragraph()
        current_text = paragraph.to_string(aw.SaveFormat.TEXT).rstrip('\r\n')
        if _is_evaluation_only_paragraph(current_text):
            continue
        user_paragraph_texts.append(current_text)

    assert 0 <= user_paragraph_ordinal < len(user_paragraph_texts)
    assert user_paragraph_texts[user_paragraph_ordinal] == paragraph_text


def test_omitting_custom_node_id_keeps_default_node_id_behavior_for_created_paragraphs():
    doc_id = srv.tool_create_document('default-custom-node-id-behavior.docx')['docId']

    add_heading_response = srv.tool_add_heading(doc_id, 'Default node id heading', level=1)
    add_paragraph_response = srv.tool_add_paragraph(doc_id, 'Default node id paragraph')

    assert add_heading_response == {}
    assert add_paragraph_response == {}

    heading_paragraph = _paragraph_with_exact_text(doc_id, 'Default node id heading')
    body_paragraph = _paragraph_with_exact_text(doc_id, 'Default node id paragraph')
    assert heading_paragraph.custom_node_id == 0
    assert body_paragraph.custom_node_id == 0

    sidecar_payload = _custom_node_id_sidecar_payload(doc_id)
    if sidecar_payload is None:
        assert sidecar_payload is None
    else:
        assert sidecar_payload.get('doc_id') == doc_id
        paragraph_custom_node_ids = sidecar_payload.get('paragraph_custom_node_ids')
        assert paragraph_custom_node_ids in (None, {})


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
