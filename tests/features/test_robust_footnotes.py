import pytest
pytest.importorskip('aspose.words')
import mcp_server as srv

def test_footnotes_by_anchor_add_validate_delete():
    r = srv.tool_create_document('anchors.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'Intro')
    srv.tool_add_paragraph(did, 'This line has ANCHOR text')
    srv.tool_add_paragraph(did, 'Another ANCHOR appears here')
    srv.tool_add_footnote_by_anchor(did, anchor_text='ANCHOR', text='FN1', position='after', occurrence=1)
    v = srv.tool_validate_footnotes_by_anchor(did, anchor_text='ANCHOR', min_count=1)
    assert v['ok'] is True
    assert v['count'] >= 1
    srv.tool_add_footnote_by_anchor(did, anchor_text='ANCHOR', text='FN2', position='before', occurrence=2)
    v2 = srv.tool_validate_footnotes_by_anchor(did, anchor_text='ANCHOR', min_count=2)
    assert v2['ok'] is True
    assert v2['count'] >= 2
    notes = srv.tool_get_all_notes(did)
    assert isinstance(notes, dict) and isinstance(notes.get('notes', []), list)
    assert len(notes['notes']) >= 2
    del_one = srv.tool_delete_footnotes_by_anchor(did, anchor_text='ANCHOR', occurrence=1, remove_all=False)
    assert isinstance(del_one['count'], int) and del_one['count'] >= 1
    v3 = srv.tool_validate_footnotes_by_anchor(did, anchor_text='ANCHOR', min_count=1)
    assert v3['ok'] is True
    assert v3['count'] >= 1
    del_all = srv.tool_delete_footnotes_by_anchor(did, anchor_text='ANCHOR', remove_all=True)
    assert isinstance(del_all['count'], int) and del_all['count'] >= 0
    v4 = srv.tool_validate_footnotes_by_anchor(did, anchor_text='ANCHOR', min_count=1)
    assert isinstance(v4['count'], int)

def test_endnotes_by_anchor_smoke():
    r = srv.tool_create_document('endnotes.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'Prelude')
    srv.tool_add_paragraph(did, 'Anchor HERE for endnote')
    srv.tool_add_endnote_by_anchor(did, anchor_text='HERE', text='EN1', position='after')
    v = srv.tool_validate_endnotes_by_anchor(did, anchor_text='HERE', min_count=1)
    assert v['ok'] is True
    assert v['count'] >= 1
    removed = srv.tool_delete_endnotes_by_anchor(did, anchor_text='HERE', remove_all=True)
    assert isinstance(removed['count'], int) and removed['count'] >= 0

def test_add_footnote_by_anchor_missing_anchor_raises():
    r = srv.tool_create_document('miss.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'No anchors here')
    with pytest.raises(ValueError):
        srv.tool_add_footnote_by_anchor(did, anchor_text='MISSING', text='X')
