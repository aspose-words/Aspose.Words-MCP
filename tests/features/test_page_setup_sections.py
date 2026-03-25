import pytest

pytest.importorskip('aspose.words')
import mcp_server as srv


def test_page_setup_and_section_breaks():
    res = srv.tool_create_document('p0-setup.docx')
    did = res['docId']
    xml_before = srv.tool_get_xml(did)['xml']
    c1 = xml_before.count('sectPr')
    srv.tool_set_page_setup(
        did,
        margins={'top': 36.0, 'bottom': 36.0, 'left': 36.0, 'right': 36.0},
        orientation='landscape',
        paper='A4',
        columns=2,
    )
    xml_mid = srv.tool_get_xml(did)['xml']
    assert isinstance(xml_mid, str) and xml_mid != xml_before
    srv.tool_insert_section_break(did, kind='nextPage')
    xml_after = srv.tool_get_xml(did)['xml']
    c2 = xml_after.count('sectPr')
    assert c2 >= c1
