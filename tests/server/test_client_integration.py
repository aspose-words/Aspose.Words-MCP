import asyncio
import base64
import os
from contextlib import asynccontextmanager

import pytest
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


def test_replace_text_literal_compatibility(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='insert_text_end',
                arguments={'doc_id': doc_id, 'text': 'Hello World world'},
            )
            replaced = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': doc_id,
                    'search_text': 'World',
                    'replacement_text': 'MCP',
                    'case_sensitive': True,
                },
            )
            assert hasattr(replaced, 'data')
            assert replaced.data == {'count': 1}
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_replace_text_regex_options(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='insert_text_end',
                arguments={'doc_id': doc_id, 'text': 'cat scatter Cat'},
            )
            replaced = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': doc_id,
                    'search_text': 'cat',
                    'replacement_text': 'dog',
                    'use_regex': True,
                    'whole_word': True,
                    'case_sensitive': False,
                },
            )
            assert hasattr(replaced, 'data')
            assert replaced.data == {'count': 2}
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_replace_text_regex_invalid_pattern_is_explicit(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='insert_text_end',
                arguments={'doc_id': doc_id, 'text': 'sample content'},
            )
            with pytest.raises(Exception):
                await client.call_tool(
                    name='replace_text',
                    arguments={
                        'doc_id': doc_id,
                        'search_text': '(',
                        'replacement_text': 'x',
                        'use_regex': True,
                    },
                )
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_replace_regex_to_images_base64_success_and_shape(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='insert_text_end',
                arguments={'doc_id': doc_id, 'text': 'Item-100 Item-200'},
            )
            replaced = await client.call_tool(
                name='replace_regex_to_images_base64',
                arguments={
                    'doc_id': doc_id,
                    'pattern': r'Item-\d+',
                    'replacement_text': 'Updated',
                    'fmt': 'png',
                    'dpi': 120,
                    'case_sensitive': True,
                },
            )
            assert hasattr(replaced, 'data')
            assert isinstance(replaced.data, dict)
            assert 'images' in replaced.data
            assert isinstance(replaced.data['images'], list)
            assert len(replaced.data['images']) > 0

            for image_payload in replaced.data['images']:
                assert set(image_payload.keys()) == {'base64', 'mime', 'ext'}
                assert image_payload['mime'] == 'image/png'
                assert image_payload['ext'] == 'png'
                decoded = base64.b64decode(image_payload['base64'])
                assert isinstance(decoded, (bytes, bytearray))
                assert len(decoded) > 0

            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_replace_regex_to_images_base64_svg_is_explicitly_unsupported(
    mcp_client_config, result_file_path
):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='insert_text_end',
                arguments={'doc_id': doc_id, 'text': 'Item-100 Item-200'},
            )
            with pytest.raises(Exception, match='Unsupported replace-to-images format: svg'):
                await client.call_tool(
                    name='replace_regex_to_images_base64',
                    arguments={
                        'doc_id': doc_id,
                        'pattern': r'Item-\d+',
                        'replacement_text': 'Updated',
                        'fmt': 'svg',
                    },
                )

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
