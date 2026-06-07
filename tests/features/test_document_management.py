import base64
import types
import uuid
from pathlib import Path

import pytest

pytest.importorskip('aspose.words')
import mcp_server as srv
from core import content as _content
from core import export as _export
from core import io as _io
from core import reading as _reading
from core.utils import docs_util as _docs


def _assert_canonical_uuid_text(value: str) -> None:
    parsed = uuid.UUID(value)
    assert str(parsed) == value


def test_document_id_generators_return_canonical_uuid_text(monkeypatch, tmp_path):
    generated_ids = iter(
        [
            uuid.UUID('11111111-1111-4111-8111-111111111111'),
            uuid.UUID('22222222-2222-4222-8222-222222222222'),
            uuid.UUID('33333333-3333-4333-8333-333333333333'),
            uuid.UUID('44444444-4444-4444-8444-444444444444'),
            uuid.UUID('55555555-5555-4555-8555-555555555555'),
        ]
    )
    saved_paths = []

    class FakeDocument:
        def __init__(self, path: str | None = None):
            self.path = path
            self.append_calls = []

        def append_document(self, *args):
            self.append_calls.append(args)

        def save(self, path: str, *_args):
            saved_paths.append(path)

    source_file = tmp_path / 'source file.docx'
    source_file.write_bytes(b'docx placeholder')
    monkeypatch.setattr(_io.uuid, 'uuid4', lambda: next(generated_ids))
    monkeypatch.setattr(_io.aw, 'Document', FakeDocument)
    monkeypatch.setattr(_io.aw, 'SaveFormat', types.SimpleNamespace(DOCX='DOCX'))
    monkeypatch.setattr(
        _io.aw,
        'ImportFormatMode',
        types.SimpleNamespace(KEEP_SOURCE_FORMATTING='KEEP_SOURCE_FORMATTING'),
    )
    monkeypatch.setattr(_io, 'ensure_path', lambda doc_id: tmp_path / f'{doc_id}.docx')
    monkeypatch.setattr(_io, 'docx_path', lambda doc_id: tmp_path / f'{doc_id}.docx')

    created_id, created_name = _io.create_document('created.docx')
    imported_id, imported_name = _io.import_from_file(str(source_file))
    copied_id = _io.copy('source-doc')
    saved_id, saved_name = _io.save_as_new('source-doc', name='saved.docx', fmt='docx')
    merged_id = _io.merge(['source-a', 'source-b'])

    assert [created_id, imported_id, copied_id, saved_id, merged_id] == [
        '11111111-1111-4111-8111-111111111111',
        '22222222-2222-4222-8222-222222222222',
        '33333333-3333-4333-8333-333333333333',
        '44444444-4444-4444-8444-444444444444',
        '55555555-5555-4555-8555-555555555555',
    ]
    for generated_id in [created_id, imported_id, copied_id, saved_id, merged_id]:
        _assert_canonical_uuid_text(generated_id)
    assert created_name == 'created.docx'
    assert imported_name == 'source file.docx'
    assert saved_name == 'saved.docx'
    assert saved_paths == [
        str(tmp_path / f'{created_id}.docx'),
        str(tmp_path / f'{imported_id}.docx'),
        str(tmp_path / f'{copied_id}.docx'),
        str(tmp_path / f'{saved_id}.docx'),
        str(tmp_path / f'{merged_id}.docx'),
    ]


def test_create_and_get_info():
    doc_id, name = _io.create_document('hello.docx')
    assert doc_id and name.endswith('.docx')
    info = _reading.get_info(doc_id)
    assert isinstance(info, dict)
    assert info['paragraphs'] >= 1


def test_get_info_via_tool():
    res = srv.tool_create_document('info.docx')
    doc_id = res['docId']
    srv.tool_insert_text_end(doc_id, ' Some content here')
    info = srv.tool_get_info(doc_id)
    assert isinstance(info, dict)
    for key in ('sizeBytes', 'words', 'paragraphs'):
        assert key in info
    assert info['words'] >= 2
    assert info['paragraphs'] >= 1


def test_export_base64_tool_smoke():
    res = srv.tool_create_document('file.docx')
    doc_id = res['docId']
    srv.tool_insert_text_end(doc_id, ' Some text')
    out = srv.tool_export_base64(doc_id, fmt='docx')
    assert isinstance(out, dict)
    assert out['ext'] == 'docx'
    assert out['mime'].startswith('application/')
    data = base64.b64decode(out['base64'])
    assert isinstance(data, (bytes, bytearray)) and len(data) > 0


@pytest.mark.parametrize('fmt,ext_prefix', [('docx', 'docx'), ('pdf', 'pdf'), ('rtf', 'rtf')])
def test_export_multiple_formats_via_manager(fmt, ext_prefix):
    doc_id, _ = _io.create_document('a.docx')
    _content.insert_text(doc_id, ' More text')
    data, mime, ext = _export.export(doc_id, fmt=fmt)
    assert isinstance(data, (bytes, bytearray)) and len(data) > 0
    assert isinstance(mime, str) and len(mime) > 0
    assert ext == ext_prefix


def test_merge_documents_via_manager():
    a_id, _ = _io.create_document('a.docx')
    b_id, _ = _io.create_document('b.docx')
    _content.insert_text(a_id, ' AAA')
    _content.insert_text(b_id, ' BBB')
    merged_id = _io.merge([a_id, b_id])
    assert isinstance(merged_id, str) and merged_id != a_id and (merged_id != b_id)
    data, _, _ = _export.export(merged_id, fmt='docx')
    assert len(data) > 0


def test_merge_invalid_ids_raises():
    a_id, _ = _io.create_document('a.docx')
    with pytest.raises(FileNotFoundError):
        _io.merge([a_id, 'missing-id'])


def test_merge_omitted_new_page_option_uses_legacy_append_signature(monkeypatch, tmp_path):
    class FakeImportFormatOptions:
        def __init__(self):
            self.append_document_with_new_page = None

    class FakeDocument:
        instances = []

        def __init__(self, path: str):
            self.path = path
            self.append_calls = []
            self.saved_path = None
            FakeDocument.instances.append(self)

        def append_document(self, *args):
            self.append_calls.append(args)

        def save(self, path: str):
            self.saved_path = path

    monkeypatch.setattr(_io, 'ensure_path', lambda sid: Path(tmp_path / f'{sid}.docx'))
    monkeypatch.setattr(_io, 'docx_path', lambda sid: Path(tmp_path / f'{sid}.docx'))
    monkeypatch.setattr(
        _io.uuid, 'uuid4', lambda: uuid.UUID('00000000-0000-0000-0000-000000000123')
    )
    monkeypatch.setattr(_io.aw, 'Document', FakeDocument)
    monkeypatch.setattr(_io.aw, 'ImportFormatOptions', FakeImportFormatOptions)
    monkeypatch.setattr(
        _io.aw,
        'ImportFormatMode',
        types.SimpleNamespace(KEEP_SOURCE_FORMATTING='KEEP_SOURCE_FORMATTING'),
    )

    merged_id = _io.merge(['source-a', 'source-b'])

    assert merged_id == '00000000-0000-0000-0000-000000000123'
    receiving_docs = [instance for instance in FakeDocument.instances if instance.append_calls]
    assert len(receiving_docs) == 1
    result_doc = receiving_docs[0]
    assert len(result_doc.append_calls) == 1
    append_args = result_doc.append_calls[0]
    assert len(append_args) == 2
    assert append_args[1] == 'KEEP_SOURCE_FORMATTING'


@pytest.mark.parametrize('append_with_new_page', [True, False])
def test_merge_explicit_new_page_option_passes_import_format_option(
    monkeypatch,
    tmp_path,
    append_with_new_page,
):
    class FakeImportFormatOptions:
        def __init__(self):
            self.append_document_with_new_page = None

    class FakeDocument:
        instances = []

        def __init__(self, path: str):
            self.path = path
            self.append_calls = []
            FakeDocument.instances.append(self)

        def append_document(self, *args):
            self.append_calls.append(args)

        def save(self, path: str):
            self.saved_path = path

    monkeypatch.setattr(_io, 'ensure_path', lambda sid: Path(tmp_path / f'{sid}.docx'))
    monkeypatch.setattr(_io, 'docx_path', lambda sid: Path(tmp_path / f'{sid}.docx'))
    monkeypatch.setattr(
        _io.uuid, 'uuid4', lambda: uuid.UUID('00000000-0000-0000-0000-000000000456')
    )
    monkeypatch.setattr(_io.aw, 'Document', FakeDocument)
    monkeypatch.setattr(_io.aw, 'ImportFormatOptions', FakeImportFormatOptions)
    monkeypatch.setattr(
        _io.aw,
        'ImportFormatMode',
        types.SimpleNamespace(KEEP_SOURCE_FORMATTING='KEEP_SOURCE_FORMATTING'),
    )

    merged_id = _io.merge(
        ['source-a', 'source-b'], append_document_with_new_page=append_with_new_page
    )

    assert merged_id == '00000000-0000-0000-0000-000000000456'
    receiving_docs = [instance for instance in FakeDocument.instances if instance.append_calls]
    assert len(receiving_docs) == 1
    result_doc = receiving_docs[0]
    assert len(result_doc.append_calls) == 1
    append_args = result_doc.append_calls[0]
    assert len(append_args) == 3
    assert append_args[1] == 'KEEP_SOURCE_FORMATTING'
    assert append_args[2].append_document_with_new_page is append_with_new_page


def test_import_header_footer_node_forwards_footer_primary_options(monkeypatch, tmp_path):
    source_node = object()
    imported_node = object()
    source_path = Path(tmp_path / 'source-doc.docx')
    destination_path = Path(tmp_path / 'destination-doc.docx')

    class FakeImportFormatOptions:
        instances = []

        def __init__(self):
            self.resolve_theme_colors = None
            FakeImportFormatOptions.instances.append(self)

    class FakeSourceHeadersFooters:
        def __init__(self):
            self.requested_header_footer_type = None

        def get_by_header_footer_type(self, header_footer_type):
            self.requested_header_footer_type = header_footer_type
            return source_node

    class FakeDestinationHeadersFooters:
        def __init__(self):
            self.added_nodes = []

        def add(self, imported_header_footer_node):
            self.added_nodes.append(imported_header_footer_node)

    class FakeSection:
        def __init__(self, headers_footers):
            self.headers_footers = headers_footers

    class FakeDocument:
        instances = []

        def __init__(self, path: str):
            self.path = path
            self.import_calls = []
            self.saved_path = None
            if path == str(source_path):
                self.first_section = FakeSection(FakeSourceHeadersFooters())
            else:
                self.first_section = FakeSection(FakeDestinationHeadersFooters())
            FakeDocument.instances.append(self)

        def import_node(
            self,
            *,
            src_node,
            is_import_children,
            import_format_mode,
            import_format_options,
        ):
            self.import_calls.append(
                {
                    'src_node': src_node,
                    'is_import_children': is_import_children,
                    'import_format_mode': import_format_mode,
                    'import_format_options': import_format_options,
                }
            )
            return imported_node

        def save(self, path: str):
            self.saved_path = path

    monkeypatch.setattr(
        _io,
        'ensure_path',
        lambda document_id: source_path if document_id == 'source-doc' else destination_path,
    )
    monkeypatch.setattr(_io.aw, 'Document', FakeDocument)
    monkeypatch.setattr(_io.aw, 'ImportFormatOptions', FakeImportFormatOptions)
    monkeypatch.setattr(
        _io.aw,
        'HeaderFooterType',
        types.SimpleNamespace(
            HEADER_PRIMARY='HEADER_PRIMARY',
            HEADER_FIRST='HEADER_FIRST',
            HEADER_EVEN='HEADER_EVEN',
            FOOTER_PRIMARY='FOOTER_PRIMARY',
            FOOTER_FIRST='FOOTER_FIRST',
            FOOTER_EVEN='FOOTER_EVEN',
        ),
        raising=False,
    )
    monkeypatch.setattr(
        _io.aw,
        'ImportFormatMode',
        types.SimpleNamespace(KEEP_SOURCE_FORMATTING='KEEP_SOURCE_FORMATTING'),
    )

    returned_doc_id = _io.import_header_footer_node(
        'source-doc',
        'destination-doc',
        'footer_primary',
        resolve_theme_colors=True,
    )

    assert returned_doc_id == 'destination-doc'
    source_document = FakeDocument.instances[0]
    destination_document = FakeDocument.instances[1]
    assert (
        source_document.first_section.headers_footers.requested_header_footer_type
        == 'FOOTER_PRIMARY'
    )
    assert len(FakeImportFormatOptions.instances) == 1
    options = FakeImportFormatOptions.instances[0]
    assert options.resolve_theme_colors is True
    assert destination_document.import_calls == [
        {
            'src_node': source_node,
            'is_import_children': True,
            'import_format_mode': 'KEEP_SOURCE_FORMATTING',
            'import_format_options': options,
        }
    ]
    assert destination_document.first_section.headers_footers.added_nodes == [imported_node]
    assert destination_document.saved_path == str(destination_path)


def test_import_header_footer_node_invalid_selector_raises_value_error(monkeypatch):
    monkeypatch.setattr(
        _io.aw,
        'HeaderFooterType',
        types.SimpleNamespace(
            HEADER_PRIMARY='HEADER_PRIMARY',
            HEADER_FIRST='HEADER_FIRST',
            HEADER_EVEN='HEADER_EVEN',
            FOOTER_PRIMARY='FOOTER_PRIMARY',
            FOOTER_FIRST='FOOTER_FIRST',
            FOOTER_EVEN='FOOTER_EVEN',
        ),
        raising=False,
    )

    with pytest.raises(ValueError, match='sidebar') as error:
        _io.import_header_footer_node('source-doc', 'destination-doc', 'sidebar')

    assert 'footer_primary' in str(error.value)


def test_tool_import_header_footer_node_forwards_arguments(monkeypatch):
    forwarded_arguments = []

    def fake_import_header_footer_node(
        source_doc_id,
        destination_doc_id,
        header_footer_type,
        resolve_theme_colors=False,
    ):
        forwarded_arguments.append(
            {
                'source_doc_id': source_doc_id,
                'destination_doc_id': destination_doc_id,
                'header_footer_type': header_footer_type,
                'resolve_theme_colors': resolve_theme_colors,
            }
        )
        return destination_doc_id

    monkeypatch.setattr(
        srv._io,
        'import_header_footer_node',
        fake_import_header_footer_node,
    )

    wrapper_response = srv.tool_import_header_footer_node(
        'source-doc',
        'destination-doc',
        'footer_primary',
        resolve_theme_colors=True,
    )

    assert wrapper_response == {'docId': 'destination-doc'}
    assert forwarded_arguments == [
        {
            'source_doc_id': 'source-doc',
            'destination_doc_id': 'destination-doc',
            'header_footer_type': 'footer_primary',
            'resolve_theme_colors': True,
        }
    ]


def test_tool_merge_forwards_none_new_page_flag_when_omitted(monkeypatch):
    calls = []
    mapped = []

    def fake_merge(doc_ids, append_document_with_new_page=None):
        calls.append((doc_ids, append_document_with_new_page))
        return 'merged-id-omitted'

    monkeypatch.setattr(srv._io, 'merge', fake_merge)
    monkeypatch.setattr(srv.document_store, 'get_document_name', lambda doc_id: 'source.docx')
    monkeypatch.setattr(
        srv, '_store_add_mapping', lambda doc_id, name: mapped.append((doc_id, name))
    )

    result = srv.tool_merge(['doc-1', 'doc-2'])

    assert result == {'docId': 'merged-id-omitted'}
    assert calls == [(['doc-1', 'doc-2'], None)]
    assert mapped == [('merged-id-omitted', 'merged_source.docx')]


@pytest.mark.parametrize('append_with_new_page', [True, False])
def test_tool_merge_forwards_explicit_new_page_flag(monkeypatch, append_with_new_page):
    calls = []
    mapped = []

    def fake_merge(doc_ids, append_document_with_new_page=None):
        calls.append((doc_ids, append_document_with_new_page))
        return 'merged-id-1'

    monkeypatch.setattr(srv._io, 'merge', fake_merge)
    monkeypatch.setattr(srv.document_store, 'get_document_name', lambda doc_id: 'source.docx')
    monkeypatch.setattr(
        srv, '_store_add_mapping', lambda doc_id, name: mapped.append((doc_id, name))
    )

    result = srv.tool_merge(['doc-1', 'doc-2'], append_document_with_new_page=append_with_new_page)

    assert result == {'docId': 'merged-id-1'}
    assert calls == [(['doc-1', 'doc-2'], append_with_new_page)]
    assert mapped == [('merged-id-1', 'merged_source.docx')]


def test_list_copy_text_xml_save_delete_merge():
    r1 = srv.tool_create_document('a.docx')
    r2 = srv.tool_create_document('b.docx')
    id1, id2 = (r1['docId'], r2['docId'])
    srv.tool_insert_text_end(id1, ' Alpha')
    lst = srv.tool_list_documents()
    assert id1 in lst['docIds'] and id2 in lst['docIds']
    t = srv.tool_get_text(id1)
    assert isinstance(t['text'], str) and 'Alpha' in t['text']
    x = srv.tool_get_xml(id1)
    assert isinstance(x['xml'], str) and len(x['xml']) > 0
    cp = srv.tool_copy_document(id1)
    id3 = cp['docId']
    assert id3 and id3 != id1
    assert _docs.document_exists(id3)
    m = srv.tool_merge([id1, id2])
    mid = m['docId']
    assert isinstance(mid, str) and mid not in (id1, id2)
    sn = srv.tool_save_as_new(id1, name='copy.docx', fmt='docx')
    sn_id = sn['docId']
    assert _docs.document_exists(sn_id)
    b64 = srv.tool_get_document_base64(id1)
    raw = base64.b64decode(b64['base64'])
    assert isinstance(raw, (bytes, bytearray)) and len(raw) > 0
    srv.tool_delete_document(id2)
    with pytest.raises(FileNotFoundError):
        _docs.ensure_path(id2)


def test_properties_get_set():
    r = srv.tool_create_document('meta.docx')
    did = r['docId']
    updated = srv.tool_properties_set(did, title='T', author='A', subject='S', keywords='K')
    assert updated['title'] == 'T'
    props = srv.tool_properties_get(did)
    assert props['title'] == 'T'
    assert props['author'] == 'A'
