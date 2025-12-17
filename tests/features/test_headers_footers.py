import pytest
pytest.importorskip('aspose.words')
import mcp_server as srv

def test_headers_footers_and_numbering():
    res = srv.tool_create_document('p0-hf.docx')
    did = res['docId']
    srv.tool_add_header_text(did, 'HEADER')
    srv.tool_add_footer_text(did, 'FOOTER')
    srv.tool_set_different_first_page_header_footer(did, enabled=True)
    srv.tool_add_page_numbering(did, format_string='Page {PAGE} of {NUMPAGES}')
    xml = srv.tool_get_xml(did)['xml']
    assert isinstance(xml, str) and len(xml) > 0
    assert 'HEADER' in xml or 'HEADER' in xml.upper()
    assert 'FOOTER' in xml or 'FOOTER' in xml.upper()
    assert 'PAGE' in xml or 'NUMPAGES' in xml
