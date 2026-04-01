import pytest

pytest.importorskip('aspose.words')
import aspose.words as aw

import core.layout as layout
import mcp_server as srv
from core.utils import docs_util as _docs


@pytest.mark.parametrize(
    ('paper_input', 'expected_paper_size'),
    [
        ('A3', aw.PaperSize.A3),
        ('a4', aw.PaperSize.A4),
        (' A5 ', aw.PaperSize.A5),
        ('legal', aw.PaperSize.LEGAL),
        ('LETTER', aw.PaperSize.LETTER),
    ],
)
def test_resolve_paper_size_allowlist_mappings(paper_input, expected_paper_size):
    assert layout._resolve_paper_size(paper_input) == expected_paper_size


@pytest.mark.parametrize(
    ('paper_value', 'expected_message'),
    [
        ('B6', "Unsupported paper size 'B6'. Supported values: A3, A4, A5, LEGAL, LETTER"),
        ('', "Unsupported paper size ''. Supported values: A3, A4, A5, LEGAL, LETTER"),
    ],
)
def test_set_page_setup_with_invalid_paper_raises_clear_error(paper_value, expected_message):
    res = srv.tool_create_document('p0-invalid-paper.docx')
    did = res['docId']

    with pytest.raises(ValueError) as exc:
        srv.tool_set_page_setup(did, paper=paper_value)

    assert str(exc.value) == expected_message


def test_page_setup_and_section_breaks_with_valid_mapped_paper():
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
    document_path = _docs.ensure_path(did)
    document = aw.Document(str(document_path))
    assert document.sections[0].page_setup.paper_size == aw.PaperSize.A4
    xml_mid = srv.tool_get_xml(did)['xml']
    assert xml_mid != xml_before
    assert 'w:orient="landscape"' in xml_mid
    srv.tool_insert_section_break(did, kind='nextPage')
    xml_after = srv.tool_get_xml(did)['xml']
    c2 = xml_after.count('sectPr')
    assert c2 >= c1
