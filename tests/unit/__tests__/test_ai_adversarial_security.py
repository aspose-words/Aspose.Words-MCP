from pathlib import Path
from types import SimpleNamespace

import pytest

import core.ai as ai


class _FakeSummarizeOptions:
    def __init__(self):
        self.summary_length = None


class _FakeSummaryDocument:
    def __init__(self):
        self.saved = []

    def save(self, path: str, fmt):
        self.saved.append((path, fmt))


def _patch_runtime(monkeypatch, tmp_path: Path, source_filename: str = 'source.docx'):
    source_path = tmp_path / source_filename
    source_path.write_text('placeholder', encoding='utf-8')
    source_doc = SimpleNamespace(kind='source-doc')

    monkeypatch.setattr(ai, 'ensure_path', lambda _doc_id: source_path)
    monkeypatch.setattr(ai.aw, 'Document', lambda _path: source_doc)
    monkeypatch.setattr(ai.aw.ai, 'SummarizeOptions', _FakeSummarizeOptions)

    return SimpleNamespace(source_path=source_path, source_doc=source_doc)


@pytest.mark.parametrize(
    ('doc_id', 'error_message'),
    [
        ('.', 'doc_id must be a simple identifier, not a path'),
        ('./', 'doc_id must be a simple identifier, not a path'),
        ('.\\', 'doc_id must be a simple identifier, not a path'),
        ('../escape', 'doc_id must not contain path traversal segments'),
        ('folder/../escape', 'doc_id must not contain path traversal segments'),
        ('folder\\..\\escape', 'doc_id must be a simple identifier, not a path'),
    ],
)
def test_rejects_path_traversal_and_current_directory_doc_ids(doc_id, error_message):
    with pytest.raises(ValueError, match=error_message):
        ai.summarize_document(doc_id=doc_id, model_name='gpt-4.1', api_key='k')


@pytest.mark.parametrize(
    ('kwargs', 'error_message'),
    [
        ({'doc_id': None, 'model_name': 'gpt-4.1', 'api_key': 'k'}, 'doc_id is required'),
        ({'doc_id': 'doc1', 'model_name': '', 'api_key': 'k'}, 'model_name is required'),
        ({'doc_id': 'doc1', 'model_name': 'gpt-4.1', 'api_key': '   '}, 'api_key is required'),
    ],
)
def test_rejects_malformed_required_inputs(kwargs, error_message):
    with pytest.raises(ValueError, match=error_message):
        ai.summarize_document(**kwargs)


def test_rejects_output_overwrite_before_document_model_and_summarize(monkeypatch, tmp_path):
    runtime = _patch_runtime(monkeypatch, tmp_path, source_filename='source.docx')
    summary_doc = _FakeSummaryDocument()
    document_calls = []
    model_factory_calls = []
    summarize_calls = []

    def document_ctor(path: str):
        document_calls.append(path)
        return runtime.source_doc

    monkeypatch.setattr(ai.aw, 'Document', document_ctor)

    def model_factory(_name: str, _api_key: str):
        model_factory_calls.append(1)
        return SimpleNamespace(summarize=lambda **_kwargs: summary_doc)

    def summarize_call(_model, _source_doc, _options):
        summarize_calls.append(1)
        return summary_doc

    with pytest.raises(
        ValueError,
        match=(
            'output_name must not match the source document filename; '
            'choose a different output_name'
        ),
    ):
        ai.summarize_document(
            doc_id='doc1',
            model_name='gpt-4.1',
            api_key='k',
            output_name='../../source.docx',
            model_factory=model_factory,
            summarize_call=summarize_call,
        )

    assert document_calls == []
    assert model_factory_calls == []
    assert summarize_calls == []
    assert summary_doc.saved == []


@pytest.mark.parametrize('output_name', ['source.docx', 'SOURCE.DOCX', 'SoUrCe.DoCx'])
def test_case_insensitive_source_name_collisions_are_always_rejected(
    output_name, monkeypatch, tmp_path
):
    _patch_runtime(monkeypatch, tmp_path, source_filename='source.docx')

    with pytest.raises(
        ValueError,
        match=(
            'output_name must not match the source document filename; '
            'choose a different output_name'
        ),
    ):
        ai.summarize_document(
            doc_id='doc1',
            model_name='gpt-4.1',
            api_key='k',
            output_name=output_name,
            model_factory=lambda _n, _k: SimpleNamespace(),
            summarize_call=lambda _m, _s, _o: _FakeSummaryDocument(),
        )


def test_rejects_oversized_invalid_summary_length_before_model_construction():
    oversized_invalid_summary_length = 'x' * 12000
    model_factory_calls = []

    def model_factory(_name: str, _api_key: str):
        model_factory_calls.append(1)
        return SimpleNamespace()

    with pytest.raises(ValueError, match='Unsupported summary_length'):
        ai.summarize_document(
            doc_id='doc1',
            model_name='gpt-4.1',
            api_key='k',
            summary_length=oversized_invalid_summary_length,
            model_factory=model_factory,
        )

    assert model_factory_calls == []


def test_oversized_output_name_is_sanitized_to_filename_and_saved(monkeypatch, tmp_path):
    runtime = _patch_runtime(monkeypatch, tmp_path, source_filename='report.docx')
    summary_doc = _FakeSummaryDocument()

    oversized_name = 'A' * 10240
    result = ai.summarize_document(
        doc_id='doc-safe',
        model_name='gpt-4.1',
        api_key='k',
        output_name=f'../{oversized_name}',
        model_factory=lambda _n, _k: SimpleNamespace(summarize=lambda **_kwargs: summary_doc),
    )

    expected_filename = f'{oversized_name}.docx'
    expected_output_path = runtime.source_path.parent / expected_filename
    assert result['outputFilename'] == expected_filename
    assert result['outputPath'] == str(expected_output_path)
    assert summary_doc.saved == [(str(expected_output_path), ai.aw.SaveFormat.DOCX)]


def test_whitespace_summary_length_defaults_to_short_and_returns_metadata_after_save(
    monkeypatch, tmp_path
):
    runtime = _patch_runtime(monkeypatch, tmp_path, source_filename='source.docx')
    summary_doc = _FakeSummaryDocument()

    result = ai.summarize_document(
        doc_id='doc-safe-id',
        model_name='gpt-4.1-mini',
        api_key='secret',
        summary_length='   ',
        output_name='safe-output',
        model_factory=lambda _n, _k: SimpleNamespace(summarize=lambda **_kwargs: summary_doc),
    )

    expected_output_path = runtime.source_path.parent / 'safe-output.docx'
    assert summary_doc.saved == [(str(expected_output_path), ai.aw.SaveFormat.DOCX)]
    assert result == {
        'sourceDocId': 'doc-safe-id',
        'sourcePath': str(runtime.source_path),
        'outputFilename': 'safe-output.docx',
        'outputPath': str(expected_output_path),
        'modelName': 'gpt-4.1-mini',
        'summaryLength': 'short',
    }


def test_default_model_factory_is_direct_openai_constructor_reference():
    default_model_factory = ai.summarize_document.__defaults__[2]
    assert default_model_factory == ai.aw.ai.OpenAiModel
