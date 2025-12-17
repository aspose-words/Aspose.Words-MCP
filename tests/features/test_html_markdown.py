import pytest
pytest.importorskip('aspose.words')
import mcp_server as srv

def test_insert_html_variants():
    r = srv.tool_create_document('p0-html.docx')
    did = r['docId']
    srv.tool_insert_html_end(did, '<p>Html End</p>')
    srv.tool_insert_html_start(did, '<p>Html Start</p>')
    srv.tool_insert_html_at_paragraph(did, '<p>Html P0</p>', paragraph_index=0)
    paras = srv.tool_read_paragraphs(did)['paragraphs']
    joined = '\n'.join(paras)
    for needle in ('Html End', 'Html Start', 'Html P0'):
        assert needle in joined

def test_insert_markdown_variants():
    r = srv.tool_create_document('p0-md.docx')
    did = r['docId']
    srv.tool_insert_markdown_end(did, '# MD End\n\nText E')
    srv.tool_insert_markdown_start(did, '# MD Start\n\nText S')
    srv.tool_insert_markdown_at_paragraph(did, '# MD P0\n\nText P', paragraph_index=0)
    paras = srv.tool_read_paragraphs(did)['paragraphs']
    joined = '\n'.join(paras)
    for token in ('MD End', 'MD Start', 'MD P0'):
        assert token in joined
