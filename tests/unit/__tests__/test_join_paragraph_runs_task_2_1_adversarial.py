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


def _create_doc_with_paragraph_runs(paragraph_runs: list[list[str]]) -> str:
    doc_id = f'join-runs-adv-{uuid.uuid4().hex}'
    doc = aw.Document()
    body = doc.first_section.body
    for run_texts in paragraph_runs:
        para = aw.Paragraph(doc)
        for text in run_texts:
            para.append_child(aw.Run(doc, text))
        body.append_child(para)
    doc.save(str(docx_path(doc_id)))
    return doc_id


def _paragraph_run_count(doc_id: str, paragraph_index: int) -> int:
    doc = aw.Document(str(docx_path(doc_id)))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    return int(paras[paragraph_index].as_paragraph().runs.count)


def _find_paragraph_index_with_text(doc_id: str, text: str) -> int:
    doc = aw.Document(str(docx_path(doc_id)))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    for i in range(paras.count):
        para = paras[i].as_paragraph()
        paragraph_text = para.to_string(aw.SaveFormat.TEXT) or ''
        if text in paragraph_text:
            return int(i)
    raise AssertionError(f'Could not find paragraph containing marker: {text}')


def _last_paragraph_index(doc_id: str) -> int:
    doc = aw.Document(str(docx_path(doc_id)))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    return int(paras.count - 1)


@pytest.mark.parametrize('bad_index', [-1, 9999])
def test_join_paragraph_runs_rejects_index_boundaries(bad_index):
    doc_id = _create_doc_with_paragraph_runs([['A', 'B']])

    with pytest.raises(IndexError, match='paragraph_index out of range'):
        join_paragraph_runs(doc_id, paragraph_index=bad_index)


@pytest.mark.parametrize('bad_index', ['0', None, 1.5])
def test_join_paragraph_runs_rejects_malformed_index_types(bad_index):
    doc_id = _create_doc_with_paragraph_runs([['A', 'B']])

    with pytest.raises(TypeError):
        join_paragraph_runs(doc_id, paragraph_index=bad_index)


def test_join_paragraph_runs_no_op_on_single_run_paragraph():
    marker = f'SINGLE_{uuid.uuid4().hex}'
    doc_id = _create_doc_with_paragraph_runs([[marker]])
    paragraph_index = _find_paragraph_index_with_text(doc_id, marker)

    result = join_paragraph_runs(doc_id, paragraph_index=paragraph_index)

    assert result == {
        'paragraph_index': paragraph_index,
        'run_count_before': 1,
        'run_count_after': 1,
        'changed': False,
    }
    assert _paragraph_run_count(doc_id, paragraph_index) == 1


def test_join_paragraph_runs_no_op_on_empty_paragraph():
    doc_id = _create_doc_with_paragraph_runs([[]])
    paragraph_index = _last_paragraph_index(doc_id)
    before_count = _paragraph_run_count(doc_id, paragraph_index)

    result = join_paragraph_runs(doc_id, paragraph_index=paragraph_index)

    assert result['paragraph_index'] == paragraph_index
    assert result['run_count_before'] == before_count
    assert result['run_count_after'] == before_count
    assert result['changed'] is False
    assert _paragraph_run_count(doc_id, paragraph_index) == before_count


def test_join_paragraph_runs_option_value_misuse_is_bool_coerced(monkeypatch):
    captured = {}

    class FakeOptions:
        def __init__(self):
            self.ignore_redundant = None
            self.ignore_insignificant = None
            self.ignore_spacing = None

    class FakeParagraph:
        def __init__(self):
            self.runs = SimpleNamespace(count=3)

        def join_runs_with_same_formatting(self, options):
            captured['options'] = options
            self.runs.count = 1

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
        ignore_redundant='0',
        ignore_insignificant='',
        ignore_spacing='<script>alert(1)</script>',
    )

    assert captured['options'].ignore_redundant is True
    assert captured['options'].ignore_insignificant is False
    assert captured['options'].ignore_spacing is True
    assert result == {
        'paragraph_index': 0,
        'run_count_before': 3,
        'run_count_after': 1,
        'changed': True,
    }
    assert fake_doc.saved_path == 'fake.docx'


@pytest.mark.parametrize(
    'adversarial_text',
    [
        '<script>alert(1)</script>',
        '../etc/passwd',
        'zero\x00byte',
        'emoji-🧪-مرحبا',
        'x' * 12_000,
    ],
)
def test_join_paragraph_runs_idempotent_for_adversarial_text_inputs(adversarial_text):
    marker = f'MARK_{uuid.uuid4().hex}'
    doc_id = _create_doc_with_paragraph_runs([[marker, adversarial_text, 'tail']])
    paragraph_index = _find_paragraph_index_with_text(doc_id, marker)

    first = join_paragraph_runs(doc_id, paragraph_index=paragraph_index)
    second = join_paragraph_runs(doc_id, paragraph_index=paragraph_index)

    assert first['run_count_before'] == 3
    assert first['run_count_after'] == 1
    assert first['changed'] is True
    assert second['run_count_before'] == 1
    assert second['run_count_after'] == 1
    assert second['changed'] is False
