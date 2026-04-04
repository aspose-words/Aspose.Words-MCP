from pathlib import Path

import pytest

pytest.importorskip('aspose.words')
import aspose.words as aw

import mcp_server as srv
from core.utils import docs_util as _docs


def _create_run_joining_target(name: str) -> tuple[str, Path, str]:
    doc_id = srv.tool_create_document(name)['docId']
    doc_path = _docs.ensure_path(doc_id)
    document = aw.Document(str(doc_path))
    body = document.first_section.body
    body.remove_all_children()

    marker_text = 'adversarial-join-marker:'
    paragraph = aw.Paragraph(document)
    paragraph.append_child(aw.Run(document, marker_text))
    paragraph.append_child(aw.Run(document, 'token'))
    body.append_child(paragraph)

    document.save(str(doc_path))
    return doc_id, doc_path, marker_text


def _run_texts_for_marker(doc_path: Path, marker_text: str) -> list[str]:
    document = aw.Document(str(doc_path))
    paragraph_nodes = document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    for index in range(paragraph_nodes.count):
        paragraph = paragraph_nodes[index].as_paragraph()
        run_texts = [paragraph.runs[i].text for i in range(paragraph.runs.count)]
        if marker_text in ''.join(run_texts):
            return run_texts
    raise AssertionError(f'Paragraph with marker not found: {marker_text}')


def test_replace_text_rejects_search_alias_misuse_without_mutation():
    doc_id, doc_path, _ = _create_run_joining_target('replace-attack-search-alias.docx')
    before_bytes = doc_path.read_bytes()

    with pytest.raises(ValueError, match='Provide exactly one search term alias'):
        srv.tool_replace_text(
            doc_id,
            find_text='token',
            search_text='token',
            replace_text='TOKEN',
            join_runs=True,
        )

    assert doc_path.read_bytes() == before_bytes


def test_replace_text_rejects_replacement_alias_misuse_without_mutation():
    doc_id, doc_path, _ = _create_run_joining_target('replace-attack-replacement-alias.docx')
    before_bytes = doc_path.read_bytes()

    with pytest.raises(ValueError, match='Provide exactly one replacement term alias'):
        srv.tool_replace_text(
            doc_id,
            find_text='token',
            replace_text='TOKEN',
            replacement_text='TOKEN',
            join_runs=True,
        )

    assert doc_path.read_bytes() == before_bytes


def test_replace_text_rejects_missing_aliases_without_mutation():
    doc_id, doc_path, _ = _create_run_joining_target('replace-attack-missing-aliases.docx')
    before_bytes = doc_path.read_bytes()

    with pytest.raises(ValueError, match='Provide exactly one search term alias'):
        srv.tool_replace_text(doc_id, replace_text='TOKEN', join_runs=True)

    with pytest.raises(ValueError, match='Provide exactly one replacement term alias'):
        srv.tool_replace_text(doc_id, find_text='token', join_runs=True)

    assert doc_path.read_bytes() == before_bytes


def test_replace_text_zero_match_with_join_runs_preserves_runs_and_contract():
    doc_id, doc_path, marker_text = _create_run_joining_target(
        'replace-attack-zero-match-join.docx'
    )
    before_bytes = doc_path.read_bytes()
    before_runs = _run_texts_for_marker(doc_path, marker_text)

    response = srv.tool_replace_text(
        doc_id,
        find_text='absent-token',
        replace_text='TOKEN',
        join_runs=True,
        ignore_redundant=True,
        ignore_insignificant=True,
        ignore_spacing=True,
    )

    assert response == {'count': 0}
    assert _run_texts_for_marker(doc_path, marker_text) == before_runs
    assert doc_path.read_bytes() == before_bytes


@pytest.mark.parametrize('subset_pattern', ['(?=token)', '(token)'])
def test_replace_text_rejects_regex_patterns_outside_public_subset_without_mutation(
    subset_pattern: str,
):
    doc_id, doc_path, _ = _create_run_joining_target(
        f'replace-attack-subset-regex-{len(subset_pattern)}.docx'
    )
    before_bytes = doc_path.read_bytes()

    with pytest.raises(ValueError, match='regex'):
        srv.tool_replace_text(
            doc_id,
            find_text=subset_pattern,
            replace_text='TOKEN',
            use_regex=True,
            join_runs=True,
        )

    assert doc_path.read_bytes() == before_bytes


@pytest.mark.parametrize('malformed_pattern', [r'\1', 'a**'])
def test_replace_text_rejects_malformed_regex_inputs_without_mutation(malformed_pattern: str):
    doc_id, doc_path, _ = _create_run_joining_target(
        f'replace-attack-malformed-regex-{len(malformed_pattern)}.docx'
    )
    before_bytes = doc_path.read_bytes()

    with pytest.raises(ValueError, match='regex'):
        srv.tool_replace_text(
            doc_id,
            find_text=malformed_pattern,
            replace_text='TOKEN',
            use_regex=True,
            join_runs=True,
        )

    assert doc_path.read_bytes() == before_bytes


def test_replace_text_regex_zero_match_is_idempotent_and_mutation_free_with_join_runs():
    doc_id, doc_path, marker_text = _create_run_joining_target(
        'replace-attack-regex-zero-idempotent.docx'
    )
    before_bytes = doc_path.read_bytes()
    before_runs = _run_texts_for_marker(doc_path, marker_text)
    replacement_payload = '<script>alert(1)</script>${x}../\\x00🙂'

    first = srv.tool_replace_text(
        doc_id,
        search_text='does[-_ ]not[-_ ]exist',
        replacement_text=replacement_payload,
        use_regex=True,
        join_runs=True,
    )
    second = srv.tool_replace_text(
        doc_id,
        search_text='does[-_ ]not[-_ ]exist',
        replacement_text=replacement_payload,
        use_regex=True,
        join_runs=True,
    )

    assert first == {'count': 0}
    assert second == {'count': 0}
    assert _run_texts_for_marker(doc_path, marker_text) == before_runs
    assert doc_path.read_bytes() == before_bytes


def test_replace_text_join_runs_success_path_keeps_count_only_contract():
    doc_id, doc_path, marker_text = _create_run_joining_target('replace-attack-contract-drift.docx')

    response = srv.tool_replace_text(
        doc_id,
        search_text='token',
        replacement_text='TOKEN',
        join_runs=True,
        ignore_redundant=False,
        ignore_insignificant=False,
        ignore_spacing=False,
    )

    assert response == {'count': 1}
    assert _run_texts_for_marker(doc_path, marker_text) == [f'{marker_text}TOKEN']
