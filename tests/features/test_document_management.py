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


def test_merge_false_resolve_theme_colors_uses_legacy_append_signature(monkeypatch, tmp_path):
    class FakeImportFormatOptions:
        def __init__(self):
            self.append_document_with_new_page = None
            self.resolve_theme_colors = None

    class FakeDocument:
        instances = []

        def __init__(self, path: str):
            self.path = path
            self.append_calls = []
            self.import_node_calls = []
            self.saved_path = None
            FakeDocument.instances.append(self)

        def append_document(self, *args):
            self.append_calls.append(args)

        def import_node(self, **kwargs):
            self.import_node_calls.append(kwargs)
            return 'unexpected-imported-section'

        def save(self, path: str):
            self.saved_path = path

    monkeypatch.setattr(_io, 'ensure_path', lambda sid: Path(tmp_path / f'{sid}.docx'))
    monkeypatch.setattr(_io, 'docx_path', lambda sid: Path(tmp_path / f'{sid}.docx'))
    monkeypatch.setattr(
        _io.uuid, 'uuid4', lambda: uuid.UUID('00000000-0000-0000-0000-000000000789')
    )
    monkeypatch.setattr(_io.aw, 'Document', FakeDocument)
    monkeypatch.setattr(_io.aw, 'ImportFormatOptions', FakeImportFormatOptions)
    monkeypatch.setattr(
        _io.aw,
        'ImportFormatMode',
        types.SimpleNamespace(KEEP_SOURCE_FORMATTING='KEEP_SOURCE_FORMATTING'),
    )

    merged_id = _io.merge(['source-a', 'source-b'], resolve_theme_colors=False)

    assert merged_id == '00000000-0000-0000-0000-000000000789'
    receiving_docs = [instance for instance in FakeDocument.instances if instance.append_calls]
    assert len(receiving_docs) == 1
    result_doc = receiving_docs[0]
    assert len(result_doc.append_calls) == 1
    append_args = result_doc.append_calls[0]
    assert len(append_args) == 2
    assert append_args[1] == 'KEEP_SOURCE_FORMATTING'
    assert result_doc.import_node_calls == []


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


def test_merge_true_resolve_theme_colors_imports_sections_with_named_options(
    monkeypatch,
    tmp_path,
):
    class FakeImportFormatOptions:
        def __init__(self):
            self.append_document_with_new_page = None
            self.resolve_theme_colors = None

    class FakeSectionCollection:
        def __init__(self, sections=None):
            self._sections = list(sections or [])
            self.added_sections = []

        @property
        def count(self):
            return len(self._sections)

        def __getitem__(self, index):
            return self._sections[index]

        def add(self, section):
            self.added_sections.append(section)

    class FakeDocument:
        instances = []

        def __init__(self, path: str):
            self.path = path
            self.append_calls = []
            self.import_node_calls = []
            self.saved_path = None
            if path.endswith('source-b.docx'):
                self.sections = FakeSectionCollection(['section-b-1', 'section-b-2'])
            else:
                self.sections = FakeSectionCollection()
            FakeDocument.instances.append(self)

        def append_document(self, *args):
            self.append_calls.append(args)

        def import_node(self, **kwargs):
            self.import_node_calls.append(kwargs)
            return f'imported-{kwargs["src_node"]}'

        def save(self, path: str):
            self.saved_path = path

    monkeypatch.setattr(_io, 'ensure_path', lambda sid: Path(tmp_path / f'{sid}.docx'))
    monkeypatch.setattr(_io, 'docx_path', lambda sid: Path(tmp_path / f'{sid}.docx'))
    monkeypatch.setattr(
        _io.uuid, 'uuid4', lambda: uuid.UUID('00000000-0000-0000-0000-000000000abc')
    )
    monkeypatch.setattr(_io.aw, 'Document', FakeDocument)
    monkeypatch.setattr(_io.aw, 'ImportFormatOptions', FakeImportFormatOptions)
    monkeypatch.setattr(
        _io.aw,
        'ImportFormatMode',
        types.SimpleNamespace(KEEP_SOURCE_FORMATTING='KEEP_SOURCE_FORMATTING'),
    )

    merged_id = _io.merge(['source-a', 'source-b'], resolve_theme_colors=True)

    assert merged_id == '00000000-0000-0000-0000-000000000abc'
    result_doc = FakeDocument.instances[0]
    assert result_doc.append_calls == []
    assert len(result_doc.import_node_calls) == 2
    assert result_doc.sections.added_sections == ['imported-section-b-1', 'imported-section-b-2']
    for expected_section, import_node_call in zip(
        ['section-b-1', 'section-b-2'], result_doc.import_node_calls
    ):
        assert set(import_node_call) == {
            'src_node',
            'is_import_children',
            'import_format_mode',
            'import_format_options',
        }
        assert import_node_call['src_node'] == expected_section
        assert import_node_call['is_import_children'] is True
        assert import_node_call['import_format_mode'] == 'KEEP_SOURCE_FORMATTING'
        assert import_node_call['import_format_options'].resolve_theme_colors is True


def test_tool_merge_forwards_none_new_page_flag_when_omitted(monkeypatch):
    calls = []
    mapped = []

    def fake_merge(
        doc_ids,
        append_document_with_new_page=None,
        resolve_theme_colors=None,
    ):
        calls.append((doc_ids, append_document_with_new_page, resolve_theme_colors))
        return 'merged-id-omitted'

    monkeypatch.setattr(srv._io, 'merge', fake_merge)
    monkeypatch.setattr(srv.document_store, 'get_document_name', lambda doc_id: 'source.docx')
    monkeypatch.setattr(
        srv, '_store_add_mapping', lambda doc_id, name: mapped.append((doc_id, name))
    )

    merge_result = srv.tool_merge(['doc-1', 'doc-2'])

    assert merge_result == {'docId': 'merged-id-omitted'}
    assert calls == [(['doc-1', 'doc-2'], None, None)]
    assert mapped == [('merged-id-omitted', 'merged_source.docx')]


@pytest.mark.parametrize('append_with_new_page', [True, False])
def test_tool_merge_forwards_explicit_new_page_flag(monkeypatch, append_with_new_page):
    calls = []
    mapped = []

    def fake_merge(
        doc_ids,
        append_document_with_new_page=None,
        resolve_theme_colors=None,
    ):
        calls.append((doc_ids, append_document_with_new_page, resolve_theme_colors))
        return 'merged-id-1'

    monkeypatch.setattr(srv._io, 'merge', fake_merge)
    monkeypatch.setattr(srv.document_store, 'get_document_name', lambda doc_id: 'source.docx')
    monkeypatch.setattr(
        srv, '_store_add_mapping', lambda doc_id, name: mapped.append((doc_id, name))
    )

    merge_result = srv.tool_merge(
        ['doc-1', 'doc-2'], append_document_with_new_page=append_with_new_page
    )

    assert merge_result == {'docId': 'merged-id-1'}
    assert calls == [(['doc-1', 'doc-2'], append_with_new_page, None)]
    assert mapped == [('merged-id-1', 'merged_source.docx')]


def test_tool_merge_forwards_resolve_theme_colors(monkeypatch):
    calls = []
    mapped = []

    def fake_merge(
        doc_ids,
        append_document_with_new_page=None,
        resolve_theme_colors=None,
    ):
        calls.append((doc_ids, append_document_with_new_page, resolve_theme_colors))
        return 'merged-id-theme'

    monkeypatch.setattr(srv._io, 'merge', fake_merge)
    monkeypatch.setattr(srv.document_store, 'get_document_name', lambda doc_id: 'source.docx')
    monkeypatch.setattr(
        srv, '_store_add_mapping', lambda doc_id, name: mapped.append((doc_id, name))
    )

    merge_result = srv.tool_merge(['doc-1', 'doc-2'], resolve_theme_colors=True)

    assert merge_result == {'docId': 'merged-id-theme'}
    assert calls == [(['doc-1', 'doc-2'], None, True)]
    assert mapped == [('merged-id-theme', 'merged_source.docx')]


def test_registered_merge_documents_forwards_resolve_theme_colors(monkeypatch):
    captured_tool_functions = {}
    calls = []

    class FakeMcp:
        def tool(self, description=None):
            def capture_tool(function_to_register):
                captured_tool_functions[function_to_register.__name__] = function_to_register
                return function_to_register

            return capture_tool

    def fake_tool_merge(
        source_doc_ids,
        append_document_with_new_page=None,
        resolve_theme_colors=None,
    ):
        calls.append((source_doc_ids, append_document_with_new_page, resolve_theme_colors))
        return {'docId': 'merged-id-wrapper'}

    monkeypatch.setattr(srv, 'mcp', FakeMcp())
    monkeypatch.setattr(srv, 'tool_merge', fake_tool_merge)

    srv.register_tools()
    merge_documents = captured_tool_functions['merge_documents']
    merge_result = merge_documents(
        ['doc-1', 'doc-2'],
        append_document_with_new_page=False,
        resolve_theme_colors=True,
    )

    assert merge_result == {'docId': 'merged-id-wrapper'}
    assert calls == [(['doc-1', 'doc-2'], False, True)]


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


def test_remove_customizations_core_calls_document_api(monkeypatch, tmp_path):
    calls = []

    class FakeDocument:
        def __init__(self, path: str):
            self.path = path

        def remove_customizations(self):
            calls.append(('remove_customizations', self.path))

        def save(self, path: str):
            calls.append(('save', path))

    monkeypatch.setattr(srv._properties, 'ensure_path', lambda doc_id: tmp_path / f'{doc_id}.docx')
    monkeypatch.setattr(srv._properties.aw, 'Document', FakeDocument)

    assert srv._properties.remove_customizations('customized') is True

    expected_path = str(tmp_path / 'customized.docx')
    assert calls == [
        ('remove_customizations', expected_path),
        ('save', expected_path),
    ]


def test_tool_remove_customizations_delegates(monkeypatch):
    calls = []

    def fake_remove_customizations(doc_id):
        calls.append(doc_id)
        return True

    monkeypatch.setattr(srv._properties, 'remove_customizations', fake_remove_customizations)

    assert srv.tool_remove_customizations('doc-custom') == {}
    assert calls == ['doc-custom']


def test_registered_remove_customizations_delegates(monkeypatch):
    captured_tool_functions = {}
    calls = []

    class FakeMcp:
        def tool(self, description=None):
            def capture_tool(function_to_register):
                captured_tool_functions[function_to_register.__name__] = function_to_register
                return function_to_register

            return capture_tool

    def fake_tool_remove_customizations(doc_id):
        calls.append(doc_id)
        return {}

    monkeypatch.setattr(srv, 'mcp', FakeMcp())
    monkeypatch.setattr(srv, 'tool_remove_customizations', fake_tool_remove_customizations)

    srv.register_tools()
    remove_customizations = captured_tool_functions['remove_customizations']
    result = remove_customizations('registered-doc')

    assert result == {}
    assert calls == ['registered-doc']
