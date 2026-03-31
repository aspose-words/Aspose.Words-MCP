import uuid
from types import SimpleNamespace

import pytest

pytest.importorskip('aspose.words')
import aspose.words as aw

from core.content import join_paragraph_runs
from core.utils.docs_util import docx_path, init_data_dir


@pytest.fixture(autouse=True)
def isolated_data_dir(tmp_path):
    init_data_dir(tmp_path)


def _create_doc_with_runs(run_texts: list[str]) -> str:
    doc_id = f'join-runs-{uuid.uuid4().hex}'
    doc = aw.Document()
    para = aw.Paragraph(doc)
    for text in run_texts:
        para.append_child(aw.Run(doc, text))
    doc.first_section.body.append_child(para)
    doc.save(str(docx_path(doc_id)))
    return doc_id


def _find_paragraph_index_with_text(doc_id: str, text: str) -> int:
    doc = aw.Document(str(docx_path(doc_id)))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    for i in range(paras.count):
        para = paras[i].as_paragraph()
        paragraph_text = para.to_string(aw.SaveFormat.TEXT) or ''
        if text in paragraph_text:
            return int(i)
    raise AssertionError(f'Could not find paragraph containing marker: {text}')


def _get_paragraph_run_count(doc_id: str, paragraph_index: int) -> int:
    doc = aw.Document(str(docx_path(doc_id)))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    return int(paras[paragraph_index].as_paragraph().runs.count)


def test_join_paragraph_runs_merges_same_formatting_and_returns_metadata():
    marker = f'MARK_{uuid.uuid4().hex}'
    doc_id = _create_doc_with_runs([marker, 'Hello', ' ', 'world'])
    paragraph_index = _find_paragraph_index_with_text(doc_id, marker)

    result = join_paragraph_runs(doc_id, paragraph_index=paragraph_index)

    assert result == {
        'paragraph_index': paragraph_index,
        'run_count_before': 4,
        'run_count_after': 1,
        'changed': True,
    }
    assert _get_paragraph_run_count(doc_id, paragraph_index) == 1


def test_join_paragraph_runs_is_idempotent_after_first_merge():
    marker = f'MARK_{uuid.uuid4().hex}'
    doc_id = _create_doc_with_runs([marker, 'A', 'B', 'C'])
    paragraph_index = _find_paragraph_index_with_text(doc_id, marker)

    first = join_paragraph_runs(doc_id, paragraph_index=paragraph_index)
    second = join_paragraph_runs(doc_id, paragraph_index=paragraph_index)

    assert first == {
        'paragraph_index': paragraph_index,
        'run_count_before': 4,
        'run_count_after': 1,
        'changed': True,
    }
    assert second == {
        'paragraph_index': paragraph_index,
        'run_count_before': 1,
        'run_count_after': 1,
        'changed': False,
    }


@pytest.mark.parametrize('paragraph_index', [-1, 10_000])
def test_join_paragraph_runs_rejects_out_of_range_indices(paragraph_index):
    doc_id = _create_doc_with_runs(['one'])

    with pytest.raises(IndexError, match='paragraph_index out of range'):
        join_paragraph_runs(doc_id, paragraph_index=paragraph_index)


def test_join_paragraph_runs_sets_options_from_optional_booleans(monkeypatch):
    captured = {}

    class FakeOptions:
        def __init__(self):
            self.ignore_redundant = None
            self.ignore_insignificant = None
            self.ignore_spacing = None

    class FakeParagraph:
        def __init__(self):
            self.runs = SimpleNamespace(count=4)

        def join_runs_with_same_formatting(self, options):
            captured['options'] = options
            self.runs.count = 2

    class FakeNode:
        def __init__(self, paragraph):
            self._paragraph = paragraph

        def as_paragraph(self):
            return self._paragraph

    class FakeParas:
        def __init__(self, nodes):
            self._nodes = nodes
            self.count = len(nodes)

        def __getitem__(self, index):
            return self._nodes[index]

    class FakeDoc:
        def __init__(self):
            self._paragraph = FakeParagraph()
            self._paras = FakeParas([FakeNode(self._paragraph)])
            self.saved_path = None

        def get_child_nodes(self, _node_type, _is_deep):
            return self._paras

        def save(self, path):
            self.saved_path = path

    fake_doc = FakeDoc()
    monkeypatch.setattr('core.content.ensure_path', lambda _doc_id: 'fake.docx')
    monkeypatch.setattr('core.content.aw.Document', lambda _path: fake_doc)
    monkeypatch.setattr('core.content.aw.JoinRunsOptions', FakeOptions)

    result = join_paragraph_runs(
        'doc-id',
        paragraph_index=0,
        ignore_redundant=1,
        ignore_insignificant=0,
        ignore_spacing='yes',
    )

    assert captured['options'].ignore_redundant is True
    assert captured['options'].ignore_insignificant is False
    assert captured['options'].ignore_spacing is True
    assert result == {
        'paragraph_index': 0,
        'run_count_before': 4,
        'run_count_after': 2,
        'changed': True,
    }
    assert fake_doc.saved_path == 'fake.docx'


def test_join_paragraph_runs_raises_when_indexed_node_is_not_paragraph(monkeypatch):
    class FakeNode:
        def as_paragraph(self):
            return None

    class FakeParas:
        def __init__(self):
            self.count = 1

        def __getitem__(self, index):
            assert index == 0
            return FakeNode()

    class FakeDoc:
        def get_child_nodes(self, _node_type, _is_deep):
            return FakeParas()

    monkeypatch.setattr('core.content.ensure_path', lambda _doc_id: 'fake.docx')
    monkeypatch.setattr('core.content.aw.Document', lambda _path: FakeDoc())

    with pytest.raises(ValueError, match='node at paragraph_index is not a paragraph'):
        join_paragraph_runs('doc-id', paragraph_index=0)
