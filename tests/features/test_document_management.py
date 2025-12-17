import base64
import pytest
pytest.importorskip('aspose.words')
import mcp_server as srv
from core import content as _content
from core import io as _io
from core import export as _export
from core import reading as _reading

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
from core.utils import docs_util as _docs

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
