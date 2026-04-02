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
