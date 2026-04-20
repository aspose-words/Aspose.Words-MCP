import asyncio
import base64
import json
import os
from contextlib import asynccontextmanager
from pathlib import Path

import aspose.words as aw
from fastmcp import Client


@asynccontextmanager
async def _client_session(mcp_client_config):
    client = Client(mcp_client_config)
    async with client:
        yield client


def _run_and_assert_file(result_file_path, run_client):
    asyncio.run(run_client())
    assert os.path.isfile(result_file_path)
    assert os.path.getsize(result_file_path) > 0


async def _create_document_and_get_id(client, arguments=None):
    if arguments is None:
        arguments = {}
    resp = await client.call_tool(name='create_document', arguments=arguments)
    assert hasattr(resp, 'data')
    assert isinstance(resp.data, dict)
    assert 'docId' in resp.data
    return resp.data['docId']


def _runtime_docx_path(doc_id):
    docs_data_dir = os.environ.get('DOCS_DATA_DIR')
    if not docs_data_dir:
        raise RuntimeError('DOCS_DATA_DIR must be set for integration tests')
    return Path(docs_data_dir) / f'{doc_id}.docx'


def _has_struct_tree_root(pdf_bytes):
    return b'/StructTreeRoot' in pdf_bytes or b'/Type/StructTreeRoot' in pdf_bytes


def _has_marked_true(pdf_bytes):
    return b'/MarkInfo<</Marked true>>' in pdf_bytes or b'/Marked true' in pdf_bytes


def _set_adjacent_same_format_runs(doc_id, marker_text):
    doc_path = _runtime_docx_path(doc_id)
    document = aw.Document(str(doc_path))
    body = document.first_section.body
    body.remove_all_children()

    paragraph = aw.Paragraph(document)
    paragraph.append_child(aw.Run(document, marker_text))
    paragraph.append_child(aw.Run(document, 'token'))
    body.append_child(paragraph)

    document.save(str(doc_path))


def _set_adjacent_runs_with_spacing_difference(doc_id, marker_text):
    doc_path = _runtime_docx_path(doc_id)
    document = aw.Document(str(doc_path))
    body = document.first_section.body
    body.remove_all_children()

    paragraph = aw.Paragraph(document)
    marker_run = aw.Run(document, marker_text)
    token_run = aw.Run(document, 'token')
    token_run.font.spacing = 2.0
    paragraph.append_child(marker_run)
    paragraph.append_child(token_run)
    body.append_child(paragraph)

    document.save(str(doc_path))


def _set_runs_with_bold_whitespace_trailing(doc_id, marker_text):
    doc_path = _runtime_docx_path(doc_id)
    document = aw.Document(str(doc_path))
    body = document.first_section.body
    body.remove_all_children()

    paragraph = aw.Paragraph(document)
    paragraph.append_child(aw.Run(document, marker_text))
    paragraph.append_child(aw.Run(document, 'token'))

    trailing_whitespace_run = aw.Run(document, ' ')
    trailing_whitespace_run.font.bold = True
    paragraph.append_child(trailing_whitespace_run)
    body.append_child(paragraph)

    document.save(str(doc_path))


def _paragraph_with_marker(doc_id, marker_text):
    document = aw.Document(str(_runtime_docx_path(doc_id)))
    paragraph_nodes = document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    for paragraph_index in range(paragraph_nodes.count):
        paragraph = paragraph_nodes[paragraph_index].as_paragraph()
        paragraph_text = paragraph.to_string(aw.SaveFormat.TEXT)
        if marker_text in paragraph_text:
            return paragraph
    raise AssertionError(f'Paragraph with marker not found: {marker_text}')


async def _save_exported_to_file(result_file_path, client, doc_id):
    exported = await client.call_tool(
        name='export_base64', arguments={'doc_id': doc_id, 'fmt': 'docx'}
    )
    assert hasattr(exported, 'data')
    assert isinstance(exported.data, dict)
    assert 'base64' in exported.data
    raw = base64.b64decode(exported.data['base64'])
    assert isinstance(raw, (bytes, bytearray))
    assert len(raw) > 0
    with open(result_file_path, 'wb') as f:
        f.write(raw)


def test_client_creates_document(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Created document: OK'}
            )
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_insert_text_start_end(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client, {'name': 'start.docx'})
            await client.call_tool(
                name='insert_text_start', arguments={'doc_id': doc_id, 'text': 'Start '}
            )
            await client.call_tool(
                name='insert_text_end', arguments={'doc_id': doc_id, 'text': ' End'}
            )
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_add_paragraph_and_read(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Line 1'}
            )
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Line 2'}
            )
            resp = await client.call_tool(name='read_paragraphs', arguments={'doc_id': doc_id})
            assert hasattr(resp, 'data')
            assert isinstance(resp.data, dict)
            assert 'paragraphs' in resp.data
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_replace_text(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='insert_text_end',
                arguments={'doc_id': doc_id, 'text': 'a.b Item-10 Item-22 Item-X'},
            )

            plain = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': doc_id,
                    'find_text': 'Item-X',
                    'replace_text': 'Item-Y',
                    'use_regex': False,
                },
            )
            assert hasattr(plain, 'data')
            assert isinstance(plain.data, dict)
            assert plain.data['count'] == 1

            regex = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': doc_id,
                    'search_text': r'Item-\d+',
                    'replacement_text': 'Matched',
                    'use_regex': True,
                },
            )
            assert hasattr(regex, 'data')
            assert isinstance(regex.data, dict)
            assert regex.data['count'] == 2

            zero = await client.call_tool(
                name='replace_text',
                arguments={'doc_id': doc_id, 'find_text': 'zzz', 'replace_text': 'noop'},
            )
            assert hasattr(zero, 'data')
            assert isinstance(zero.data, dict)
            assert zero.data['count'] == 0

            marker_text = 'join-marker:'
            join_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-client.docx'}
            )
            _set_adjacent_same_format_runs(join_doc_id, marker_text)

            joined = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': join_doc_id,
                    'find_text': 'token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                    'ignore_redundant': True,
                    'ignore_insignificant': True,
                    'ignore_spacing': True,
                },
            )
            assert hasattr(joined, 'data')
            assert isinstance(joined.data, dict)
            assert set(joined.data.keys()) == {'count'}
            assert joined.data['count'] == 1

            target_paragraph = _paragraph_with_marker(join_doc_id, marker_text)
            assert target_paragraph.runs.count == 1
            assert target_paragraph.runs[0].text == f'{marker_text}TOKEN'

            spacing_without_ignore_marker_text = 'join-spacing-off-marker:'
            spacing_without_ignore_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-spacing-off-client.docx'}
            )
            _set_adjacent_runs_with_spacing_difference(
                spacing_without_ignore_doc_id, spacing_without_ignore_marker_text
            )

            spacing_without_ignore = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': spacing_without_ignore_doc_id,
                    'find_text': 'token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                },
            )
            assert hasattr(spacing_without_ignore, 'data')
            assert isinstance(spacing_without_ignore.data, dict)
            assert set(spacing_without_ignore.data.keys()) == {'count'}
            assert spacing_without_ignore.data['count'] == 1

            spacing_without_ignore_paragraph = _paragraph_with_marker(
                spacing_without_ignore_doc_id, spacing_without_ignore_marker_text
            )
            assert spacing_without_ignore_paragraph.runs.count == 2
            assert (
                spacing_without_ignore_paragraph.runs[0].text == spacing_without_ignore_marker_text
            )
            assert spacing_without_ignore_paragraph.runs[1].text == 'TOKEN'

            spacing_with_ignore_marker_text = 'join-spacing-on-marker:'
            spacing_with_ignore_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-spacing-on-client.docx'}
            )
            _set_adjacent_runs_with_spacing_difference(
                spacing_with_ignore_doc_id, spacing_with_ignore_marker_text
            )

            spacing_with_ignore = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': spacing_with_ignore_doc_id,
                    'find_text': 'token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                    'ignore_spacing': True,
                },
            )
            assert hasattr(spacing_with_ignore, 'data')
            assert isinstance(spacing_with_ignore.data, dict)
            assert set(spacing_with_ignore.data.keys()) == {'count'}
            assert spacing_with_ignore.data['count'] == 1

            spacing_with_ignore_paragraph = _paragraph_with_marker(
                spacing_with_ignore_doc_id, spacing_with_ignore_marker_text
            )
            assert spacing_with_ignore_paragraph.runs.count == 1
            assert (
                spacing_with_ignore_paragraph.runs[0].text
                == f'{spacing_with_ignore_marker_text}TOKEN'
            )

            insignificant_without_ignore_marker_text = 'join-insignificant-off-marker:'
            insignificant_without_ignore_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-insignificant-off-client.docx'}
            )
            _set_runs_with_bold_whitespace_trailing(
                insignificant_without_ignore_doc_id, insignificant_without_ignore_marker_text
            )

            insignificant_without_ignore = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': insignificant_without_ignore_doc_id,
                    'find_text': 'token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                },
            )
            assert hasattr(insignificant_without_ignore, 'data')
            assert isinstance(insignificant_without_ignore.data, dict)
            assert set(insignificant_without_ignore.data.keys()) == {'count'}
            assert insignificant_without_ignore.data['count'] == 1

            insignificant_without_ignore_paragraph = _paragraph_with_marker(
                insignificant_without_ignore_doc_id, insignificant_without_ignore_marker_text
            )
            assert insignificant_without_ignore_paragraph.runs.count == 2
            assert (
                insignificant_without_ignore_paragraph.runs[0].text
                == f'{insignificant_without_ignore_marker_text}TOKEN'
            )
            assert insignificant_without_ignore_paragraph.runs[1].text == ' '

            insignificant_with_ignore_marker_text = 'join-insignificant-on-marker:'
            insignificant_with_ignore_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-insignificant-on-client.docx'}
            )
            _set_runs_with_bold_whitespace_trailing(
                insignificant_with_ignore_doc_id, insignificant_with_ignore_marker_text
            )

            insignificant_with_ignore = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': insignificant_with_ignore_doc_id,
                    'find_text': 'token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                    'ignore_insignificant': True,
                },
            )
            assert hasattr(insignificant_with_ignore, 'data')
            assert isinstance(insignificant_with_ignore.data, dict)
            assert set(insignificant_with_ignore.data.keys()) == {'count'}
            assert insignificant_with_ignore.data['count'] == 1

            insignificant_with_ignore_paragraph = _paragraph_with_marker(
                insignificant_with_ignore_doc_id, insignificant_with_ignore_marker_text
            )
            assert insignificant_with_ignore_paragraph.runs.count == 1
            assert (
                insignificant_with_ignore_paragraph.runs[0].text
                == f'{insignificant_with_ignore_marker_text}TOKEN '
            )

            zero_match_marker_text = 'join-zero-marker:'
            zero_match_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-zero-match-client.docx'}
            )
            _set_adjacent_same_format_runs(zero_match_doc_id, zero_match_marker_text)

            before_zero_match_paragraph = _paragraph_with_marker(
                zero_match_doc_id, zero_match_marker_text
            )
            before_zero_match_run_count = before_zero_match_paragraph.runs.count
            before_zero_match_run_texts = [
                before_zero_match_paragraph.runs[run_index].text
                for run_index in range(before_zero_match_run_count)
            ]

            zero_match = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': zero_match_doc_id,
                    'find_text': 'absent-token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                },
            )
            assert hasattr(zero_match, 'data')
            assert isinstance(zero_match.data, dict)
            assert set(zero_match.data.keys()) == {'count'}
            assert zero_match.data['count'] == 0

            after_zero_match_paragraph = _paragraph_with_marker(
                zero_match_doc_id, zero_match_marker_text
            )
            after_zero_match_run_count = after_zero_match_paragraph.runs.count
            after_zero_match_run_texts = [
                after_zero_match_paragraph.runs[run_index].text
                for run_index in range(after_zero_match_run_count)
            ]
            assert after_zero_match_run_count == before_zero_match_run_count
            assert after_zero_match_run_texts == before_zero_match_run_texts

            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_list_documents_and_get_info(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            lst = await client.call_tool(name='list_documents')
            assert hasattr(lst, 'data')
            assert isinstance(lst.data, dict)
            info = await client.call_tool(name='get_info', arguments={'doc_id': doc_id})
            assert hasattr(info, 'data')
            assert isinstance(info.data, dict)
            assert 'pages' in info.data and 'paragraphs' in info.data
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_add_heading_and_get_outline(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_heading', arguments={'doc_id': doc_id, 'text': 'Title', 'level': 1}
            )
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Body'}
            )
            outline = await client.call_tool(name='get_outline', arguments={'doc_id': doc_id})
            assert hasattr(outline, 'data')
            assert isinstance(outline.data, dict)
            assert 'outline' in outline.data
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_add_heading_and_paragraph_with_custom_node_id(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)

            heading_text = 'Heading custom node id marker'
            paragraph_text = 'Paragraph custom node id marker'

            add_heading_response = await client.call_tool(
                name='add_heading',
                arguments={
                    'doc_id': doc_id,
                    'text': heading_text,
                    'level': 2,
                    'custom_node_id': 1101,
                },
            )
            assert hasattr(add_heading_response, 'data')
            assert hasattr(add_heading_response, 'structured_content')
            assert add_heading_response.structured_content == {}
            assert add_heading_response.data is None or add_heading_response.data == {}

            add_paragraph_response = await client.call_tool(
                name='add_paragraph',
                arguments={
                    'doc_id': doc_id,
                    'text': paragraph_text,
                    'custom_node_id': 2202,
                },
            )
            assert hasattr(add_paragraph_response, 'data')
            assert hasattr(add_paragraph_response, 'structured_content')
            assert add_paragraph_response.structured_content == {}
            assert add_paragraph_response.data is None or add_paragraph_response.data == {}

            await _save_exported_to_file(result_file_path, client, doc_id)

            exported_document = aw.Document(str(result_file_path))
            exported_paragraph_nodes = exported_document.get_child_nodes(
                aw.NodeType.PARAGRAPH, True
            )

            exported_heading_paragraph = None
            exported_body_paragraph = None

            for paragraph_index in range(exported_paragraph_nodes.count):
                paragraph = exported_paragraph_nodes[paragraph_index].as_paragraph()
                paragraph_text_content = paragraph.to_string(aw.SaveFormat.TEXT)
                if heading_text in paragraph_text_content:
                    exported_heading_paragraph = paragraph
                if paragraph_text in paragraph_text_content:
                    exported_body_paragraph = paragraph

            assert exported_heading_paragraph is not None
            assert exported_body_paragraph is not None

    _run_and_assert_file(result_file_path, run_client)


def test_find_text(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='insert_text_end', arguments={'doc_id': doc_id, 'text': 'Foo bar baz'}
            )
            found = await client.call_tool(
                name='find_text', arguments={'doc_id': doc_id, 'text': 'bar'}
            )
            assert hasattr(found, 'data')
            assert isinstance(found.data, dict)
            assert 'matches' in found.data
            assert isinstance(found.data['matches'], list)
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_delete_paragraph(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Line 1'}
            )
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Line 2'}
            )
            await client.call_tool(
                name='delete_paragraph', arguments={'doc_id': doc_id, 'paragraph_index': 1}
            )
            resp = await client.call_tool(name='read_paragraphs', arguments={'doc_id': doc_id})
            assert hasattr(resp, 'data')
            assert isinstance(resp.data, dict)
            assert 'paragraphs' in resp.data
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_page_breaks(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Page 1 Start'}
            )
            await client.call_tool(name='add_page_break_start', arguments={'doc_id': doc_id})
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Middle'}
            )
            await client.call_tool(name='add_page_break_end', arguments={'doc_id': doc_id})
            await client.call_tool(
                name='add_page_break_at_paragraph',
                arguments={'doc_id': doc_id, 'paragraph_index': 0},
            )
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Page 2 End'}
            )
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_insert_list_and_table(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='insert_list_end',
                arguments={'doc_id': doc_id, 'items': ['a', 'b', 'c'], 'kind': 'bullet'},
            )
            tbl = await client.call_tool(
                name='add_table_end',
                arguments={
                    'doc_id': doc_id,
                    'rows': 2,
                    'cols': 2,
                    'data': [['1', '2'], ['3', '4']],
                },
            )
            assert hasattr(tbl, 'data')
            assert isinstance(tbl.data, dict)
            assert 'tableIndex' in tbl.data
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_export_base64_advanced_docling(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Docling export payload'}
            )

            exported = await client.call_tool(
                name='export_base64_advanced',
                arguments={'doc_id': doc_id, 'fmt': 'docling'},
            )
            assert hasattr(exported, 'data')
            assert isinstance(exported.data, dict)
            assert exported.data['ext'] == 'json'
            assert exported.data['mime'] == 'application/json'
            assert 'base64' in exported.data

            raw = base64.b64decode(exported.data['base64'])
            assert isinstance(raw, (bytes, bytearray))
            assert len(raw) > 0

            payload = json.loads(raw.decode('utf-8'))
            assert isinstance(payload, (dict, list))
            assert 'Docling export payload' in json.dumps(payload)

            with open(result_file_path, 'wb') as f:
                f.write(raw)

    _run_and_assert_file(result_file_path, run_client)


def test_export_base64_advanced_pdf_text_shaping(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph',
                arguments={'doc_id': doc_id, 'text': 'PDF advanced export payload'},
            )

            pdf_default_export = await client.call_tool(
                name='export_base64_advanced',
                arguments={'doc_id': doc_id, 'fmt': 'pdf'},
            )
            assert hasattr(pdf_default_export, 'data')
            assert isinstance(pdf_default_export.data, dict)
            assert set(pdf_default_export.data) >= {'base64', 'mime', 'ext'}
            assert pdf_default_export.data['ext'] == 'pdf'
            assert pdf_default_export.data['mime'] == 'application/pdf'

            decoded_default_pdf_bytes = base64.b64decode(pdf_default_export.data['base64'])
            assert isinstance(decoded_default_pdf_bytes, (bytes, bytearray))
            assert len(decoded_default_pdf_bytes) > 0

            pdf_shaped_export = await client.call_tool(
                name='export_base64_advanced',
                arguments={
                    'doc_id': doc_id,
                    'fmt': 'pdf',
                    'options': {'enable_text_shaping': True},
                },
            )
            assert hasattr(pdf_shaped_export, 'data')
            assert isinstance(pdf_shaped_export.data, dict)
            assert set(pdf_shaped_export.data) >= {'base64', 'mime', 'ext'}
            assert pdf_shaped_export.data['ext'] == 'pdf'
            assert pdf_shaped_export.data['mime'] == 'application/pdf'

            decoded_shaped_pdf_bytes = base64.b64decode(pdf_shaped_export.data['base64'])
            assert isinstance(decoded_shaped_pdf_bytes, (bytes, bytearray))
            assert len(decoded_shaped_pdf_bytes) > 0

            with open(result_file_path, 'wb') as f:
                f.write(decoded_shaped_pdf_bytes)

    _run_and_assert_file(result_file_path, run_client)


def test_export_base64_advanced_pdf_tagged_custom_node_id_control(
    mcp_client_config, result_file_path
):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_heading',
                arguments={
                    'doc_id': doc_id,
                    'text': 'Tagged PDF heading marker',
                    'level': 1,
                    'custom_node_id': 3101,
                },
            )
            await client.call_tool(
                name='add_paragraph',
                arguments={
                    'doc_id': doc_id,
                    'text': 'Tagged PDF paragraph marker',
                    'custom_node_id': 3202,
                },
            )

            pdf_default_export = await client.call_tool(
                name='export_base64_advanced',
                arguments={'doc_id': doc_id, 'fmt': 'pdf'},
            )
            assert hasattr(pdf_default_export, 'data')
            assert isinstance(pdf_default_export.data, dict)
            assert set(pdf_default_export.data.keys()) == {'base64', 'mime', 'ext'}
            assert pdf_default_export.data['ext'] == 'pdf'
            assert pdf_default_export.data['mime'] == 'application/pdf'

            decoded_default_pdf_bytes = base64.b64decode(pdf_default_export.data['base64'])
            assert isinstance(decoded_default_pdf_bytes, (bytes, bytearray))
            assert len(decoded_default_pdf_bytes) > 0

            pdf_tagged_export = await client.call_tool(
                name='export_base64_advanced',
                arguments={
                    'doc_id': doc_id,
                    'fmt': 'pdf',
                    'tagged_pdf_preserve_custom_node_ids': True,
                },
            )
            assert hasattr(pdf_tagged_export, 'data')
            assert isinstance(pdf_tagged_export.data, dict)
            assert set(pdf_tagged_export.data.keys()) == {'base64', 'mime', 'ext'}
            assert pdf_tagged_export.data['ext'] == 'pdf'
            assert pdf_tagged_export.data['mime'] == 'application/pdf'

            decoded_tagged_pdf_bytes = base64.b64decode(pdf_tagged_export.data['base64'])
            assert isinstance(decoded_tagged_pdf_bytes, (bytes, bytearray))
            assert len(decoded_tagged_pdf_bytes) > 0

            assert _has_struct_tree_root(decoded_default_pdf_bytes) is False
            assert _has_marked_true(decoded_default_pdf_bytes) is False

            assert _has_struct_tree_root(decoded_tagged_pdf_bytes) is True
            assert _has_marked_true(decoded_tagged_pdf_bytes) is True

            with open(result_file_path, 'wb') as file_handle:
                file_handle.write(decoded_tagged_pdf_bytes)

    _run_and_assert_file(result_file_path, run_client)


def test_merge_documents_with_and_without_new_page(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            first_doc_id = await _create_document_and_get_id(client)
            second_doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph',
                arguments={'doc_id': first_doc_id, 'text': 'Default merge first source'},
            )
            await client.call_tool(
                name='add_paragraph',
                arguments={'doc_id': second_doc_id, 'text': 'Default merge second source'},
            )

            second_doc_path = _runtime_docx_path(second_doc_id)
            second_doc = aw.Document(str(second_doc_path))
            second_doc.first_section.page_setup.section_start = aw.SectionStart.ODD_PAGE
            second_doc.save(str(second_doc_path))

            second_doc_reloaded = aw.Document(str(second_doc_path))
            assert (
                second_doc_reloaded.first_section.page_setup.section_start
                == aw.SectionStart.ODD_PAGE
            )

            merged_default = await client.call_tool(
                name='merge_documents',
                arguments={'source_doc_ids': [first_doc_id, second_doc_id]},
            )
            assert hasattr(merged_default, 'data')
            assert isinstance(merged_default.data, dict)
            assert 'docId' in merged_default.data
            merged_default_id = merged_default.data['docId']
            assert isinstance(merged_default_id, str)
            assert merged_default_id

            merged_default_text = await client.call_tool(
                name='get_text', arguments={'doc_id': merged_default_id}
            )
            assert hasattr(merged_default_text, 'data')
            assert isinstance(merged_default_text.data, dict)
            assert 'text' in merged_default_text.data
            assert 'Default merge first source' in merged_default_text.data['text']
            assert 'Default merge second source' in merged_default_text.data['text']

            merged_default_info = await client.call_tool(
                name='get_info', arguments={'doc_id': merged_default_id}
            )
            assert hasattr(merged_default_info, 'data')
            assert isinstance(merged_default_info.data, dict)
            assert 'pages' in merged_default_info.data

            merged_preserve_source_section_start = await client.call_tool(
                name='merge_documents',
                arguments={
                    'source_doc_ids': [first_doc_id, second_doc_id],
                    'append_document_with_new_page': False,
                },
            )
            assert hasattr(merged_preserve_source_section_start, 'data')
            assert isinstance(merged_preserve_source_section_start.data, dict)
            assert 'docId' in merged_preserve_source_section_start.data
            merged_preserve_source_section_start_id = merged_preserve_source_section_start.data[
                'docId'
            ]
            assert isinstance(merged_preserve_source_section_start_id, str)
            assert merged_preserve_source_section_start_id

            merged_preserve_source_section_start_text = await client.call_tool(
                name='get_text', arguments={'doc_id': merged_preserve_source_section_start_id}
            )
            assert hasattr(merged_preserve_source_section_start_text, 'data')
            assert isinstance(merged_preserve_source_section_start_text.data, dict)
            assert 'text' in merged_preserve_source_section_start_text.data
            assert (
                'Default merge first source'
                in merged_preserve_source_section_start_text.data['text']
            )
            assert (
                'Default merge second source'
                in merged_preserve_source_section_start_text.data['text']
            )

            merged_preserve_source_section_start_info = await client.call_tool(
                name='get_info', arguments={'doc_id': merged_preserve_source_section_start_id}
            )
            assert hasattr(merged_preserve_source_section_start_info, 'data')
            assert isinstance(merged_preserve_source_section_start_info.data, dict)
            assert 'pages' in merged_preserve_source_section_start_info.data

            merged_default_doc = aw.Document(str(_runtime_docx_path(merged_default_id)))
            merged_preserve_source_section_start_doc = aw.Document(
                str(_runtime_docx_path(merged_preserve_source_section_start_id))
            )
            assert merged_default_doc.sections.count >= 2
            assert merged_preserve_source_section_start_doc.sections.count >= 2
            assert (
                merged_default_doc.sections[1].page_setup.section_start == aw.SectionStart.NEW_PAGE
            )
            assert (
                merged_preserve_source_section_start_doc.sections[1].page_setup.section_start
                == aw.SectionStart.ODD_PAGE
            )

            await _save_exported_to_file(
                result_file_path, client, merged_preserve_source_section_start_id
            )

    _run_and_assert_file(result_file_path, run_client)
