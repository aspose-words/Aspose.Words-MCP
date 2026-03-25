import pytest

pytest.importorskip('aspose.words')
import mcp_server as srv


def test_protect_unprotect_and_restrict():
    res = srv.tool_create_document('secure.docx')
    did = res['docId']
    srv.tool_protect_document(did, password='p1')
    srv.tool_protect_restrict(did, password='p1', ranges=[{'start': 0, 'end': 5}])
    srv.tool_unprotect_document(did, password='p1')
    srv.tool_insert_text_end(did, ' after')
    paras = srv.tool_read_paragraphs(did)['paragraphs']
    assert any(('after' in (p or '') for p in paras))


def test_comments_and_notes_smoke():
    res = srv.tool_create_document('notes.docx')
    did = res['docId']
    srv.tool_add_paragraph(did, 'para')
    srv.tool_add_footnote(did, paragraph_index=0, text='foot')
    srv.tool_add_endnote(did, paragraph_index=0, text='end')
    srv.tool_convert_footnotes_to_endnotes(did)
    srv.tool_notes_style(did, font_name='Arial', font_size=10.0)
    allc = srv.tool_get_all_comments(did)
    assert 'comments' in allc and isinstance(allc['comments'], list)
    by_author = srv.tool_get_comments_by_author(did, author='Unknown')
    assert 'comments' in by_author and isinstance(by_author['comments'], list)
    by_para = srv.tool_get_comments_for_paragraph(did, paragraph_index=0)
    assert 'comments' in by_para and isinstance(by_para['comments'], list)
    notes = srv.tool_get_all_notes(did)
    assert isinstance(notes.get('notes', []), list) and len(notes['notes']) >= 1
