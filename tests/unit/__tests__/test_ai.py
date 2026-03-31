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


class _FakeModel:
    def __init__(self, summary_doc):
        self.summary_doc = summary_doc
        self.calls = []

    def summarize(self, *, source_document, options):
        self.calls.append((source_document, options))
        return self.summary_doc


def _patch_runtime(monkeypatch, tmp_path: Path, source_filename: str = 'source.docx'):
    source_path = tmp_path / source_filename
    source_path.write_text('placeholder', encoding='utf-8')
    source_doc = SimpleNamespace(kind='source-doc')

    monkeypatch.setattr(ai, 'ensure_path', lambda _doc_id: source_path)
    monkeypatch.setattr(ai.aw, 'Document', lambda _path: source_doc)
    monkeypatch.setattr(ai.aw.ai, 'SummarizeOptions', _FakeSummarizeOptions)

    return SimpleNamespace(source_path=source_path, source_doc=source_doc)


def test_summarize_document_happy_path_constructs_model_summarizes_and_saves(monkeypatch, tmp_path):
    runtime = _patch_runtime(monkeypatch, tmp_path)
    summary_doc = _FakeSummaryDocument()
    constructed = []

    def model_factory(name: str, api_key: str):
        model = _FakeModel(summary_doc)
        constructed.append((name, api_key, model))
        return model

    result = ai.summarize_document(
        doc_id='invoice_2026',
        model_name='gpt-4.1-mini',
        api_key='test-key',
        summary_length='very-short',
        model_factory=model_factory,
    )

    assert len(constructed) == 1
    assert constructed[0][0] == 'gpt-4.1-mini'
    assert constructed[0][1] == 'test-key'
    model = constructed[0][2]
    assert len(model.calls) == 1
    assert model.calls[0][0] == runtime.source_doc
    assert model.calls[0][1].summary_length == ai._SUMMARY_LENGTH_MAP['very_short']

    expected_output_path = runtime.source_path.parent / 'source.summary.very_short.docx'
    assert summary_doc.saved == [(str(expected_output_path), ai.aw.SaveFormat.DOCX)]
    assert result == {
        'sourceDocId': 'invoice_2026',
        'sourcePath': str(runtime.source_path),
        'outputFilename': 'source.summary.very_short.docx',
        'outputPath': str(expected_output_path),
        'modelName': 'gpt-4.1-mini',
        'summaryLength': 'very_short',
    }


def test_summarize_document_default_model_factory_is_openai_constructor():
    default_model_factory = ai.summarize_document.__defaults__[2]
    assert default_model_factory == ai.aw.ai.OpenAiModel


def test_summarize_document_uses_custom_summarize_call(monkeypatch, tmp_path):
    runtime = _patch_runtime(monkeypatch, tmp_path)
    summary_doc = _FakeSummaryDocument()
    model = _FakeModel(summary_doc)
    summarize_calls = []

    def model_factory(_name: str, _api_key: str):
        return model

    def summarize_call(model_arg, source_doc_arg, options_arg):
        summarize_calls.append((model_arg, source_doc_arg, options_arg))
        return summary_doc

    result = ai.summarize_document(
        doc_id='doc_123',
        model_name='gpt-4.1',
        api_key='abc123',
        summary_length='medium',
        output_name='custom-output',
        model_factory=model_factory,
        summarize_call=summarize_call,
    )

    assert len(model.calls) == 0
    assert len(summarize_calls) == 1
    assert summarize_calls[0][0] == model
    assert summarize_calls[0][1] == runtime.source_doc
    assert summarize_calls[0][2].summary_length == ai._SUMMARY_LENGTH_MAP['medium']
    assert summary_doc.saved == [
        (str(runtime.source_path.parent / 'custom-output.docx'), ai.aw.SaveFormat.DOCX)
    ]
    assert result['outputFilename'] == 'custom-output.docx'


def test_summarize_document_propagates_constructor_failure_unchanged(monkeypatch, tmp_path):
    _patch_runtime(monkeypatch, tmp_path)
    original_error = RuntimeError('document-constructor-failure')

    def failing_document_ctor(_path: str):
        raise original_error

    monkeypatch.setattr(ai.aw, 'Document', failing_document_ctor)

    with pytest.raises(RuntimeError) as exc:
        ai.summarize_document(
            doc_id='doc-1',
            model_name='gpt-4.1-mini',
            api_key='api-key',
            model_factory=lambda _name, _api_key: _FakeModel(_FakeSummaryDocument()),
        )

    assert exc.value is original_error


@pytest.mark.parametrize(
    ('value', 'expected_normalized'),
    [
        ('short', 'short'),
        ('SHORT', 'short'),
        ('very-short', 'very_short'),
        (' very_long ', 'very_long'),
    ],
)
def test_summary_length_normalization_invariant(monkeypatch, tmp_path, value, expected_normalized):
    _patch_runtime(monkeypatch, tmp_path)
    summary_doc = _FakeSummaryDocument()
    captured = {}

    def summarize_call(model_arg, source_doc_arg, options_arg):
        captured['model'] = model_arg
        captured['source_doc'] = source_doc_arg
        captured['options'] = options_arg
        return summary_doc

    ai.summarize_document(
        doc_id='doc1',
        model_name='gpt-4.1',
        api_key='key',
        summary_length=value,
        model_factory=lambda _n, _k: SimpleNamespace(),
        summarize_call=summarize_call,
    )

    assert captured['options'].summary_length == ai._SUMMARY_LENGTH_MAP[expected_normalized]


@pytest.mark.parametrize(
    ('kwargs', 'error_message'),
    [
        (
            {'doc_id': ' ', 'model_name': 'gpt-4.1', 'api_key': 'k'},
            'doc_id is required',
        ),
        (
            {'doc_id': 'doc1', 'model_name': ' ', 'api_key': 'k'},
            'model_name is required',
        ),
        (
            {'doc_id': 'doc1', 'model_name': 'gpt-4.1', 'api_key': ' '},
            'api_key is required',
        ),
    ],
)
def test_required_inputs_are_validated(kwargs, error_message):
    with pytest.raises(ValueError, match=error_message):
        ai.summarize_document(**kwargs)


@pytest.mark.parametrize(
    ('doc_id', 'error_message'),
    [
        ('../escape', 'doc_id must not contain path traversal segments'),
        ('/absolute/path', 'doc_id must not contain path traversal segments'),
        ('nested/path', 'doc_id must be a simple identifier, not a path'),
        ('nested\\path', 'doc_id must be a simple identifier, not a path'),
    ],
)
def test_doc_id_rejects_path_like_values(doc_id, error_message):
    with pytest.raises(ValueError, match=error_message):
        ai.summarize_document(doc_id=doc_id, model_name='gpt-4.1', api_key='k')


def test_unsupported_summary_length_raises_specific_error():
    with pytest.raises(ValueError, match='Unsupported summary_length: gigantic'):
        ai.summarize_document(
            doc_id='doc1',
            model_name='gpt-4.1',
            api_key='k',
            summary_length='gigantic',
        )


@pytest.mark.parametrize('colliding_output_name', ['source.docx', 'SOURCE.DOCX'])
def test_output_filename_collision_is_rejected(monkeypatch, tmp_path, colliding_output_name):
    runtime = _patch_runtime(monkeypatch, tmp_path)
    summary_doc = _FakeSummaryDocument()
    document_calls = []
    model_factory_calls = []
    summarize_call_calls = []

    def document_ctor(path: str):
        document_calls.append(path)
        return runtime.source_doc

    monkeypatch.setattr(ai.aw, 'Document', document_ctor)

    def model_factory(_name: str, _api_key: str):
        model_factory_calls.append(1)
        return _FakeModel(summary_doc)

    def summarize_call(_model, _source_doc, _options):
        summarize_call_calls.append(1)
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
            output_name=colliding_output_name,
            model_factory=model_factory,
            summarize_call=summarize_call,
        )

    assert summary_doc.saved == []
    assert document_calls == []
    assert model_factory_calls == []
    assert summarize_call_calls == []


@pytest.mark.parametrize(
    ('doc_id', 'error_message'),
    [
        ('..', 'doc_id must not contain path traversal segments'),
        ('.', 'doc_id must be a simple identifier, not a path'),
        ('folder/../doc', 'doc_id must not contain path traversal segments'),
    ],
)
def test_doc_id_rejects_explicit_traversal_segments(doc_id, error_message):
    with pytest.raises(ValueError, match=error_message):
        ai.summarize_document(doc_id=doc_id, model_name='gpt-4.1', api_key='k')


def test_doc_id_with_windows_separator_and_dotdot_is_rejected_as_path():
    with pytest.raises(ValueError, match='doc_id must be a simple identifier, not a path'):
        ai.summarize_document(doc_id='folder\\..\\doc', model_name='gpt-4.1', api_key='k')


def test_output_name_path_traversal_attempt_is_sandboxed_to_source_directory(monkeypatch, tmp_path):
    runtime = _patch_runtime(monkeypatch, tmp_path, source_filename='report.docx')
    summary_doc = _FakeSummaryDocument()
    model = _FakeModel(summary_doc)

    result = ai.summarize_document(
        doc_id='doc-safe',
        model_name='gpt-4.1',
        api_key='k',
        output_name='../../outside-target',
        model_factory=lambda _n, _k: model,
    )

    expected_output_path = runtime.source_path.parent / 'outside-target.docx'
    assert summary_doc.saved == [(str(expected_output_path), ai.aw.SaveFormat.DOCX)]
    assert result['outputFilename'] == 'outside-target.docx'
    assert result['outputPath'] == str(expected_output_path)


def test_oversized_summary_length_is_rejected_before_model_construction(monkeypatch):
    oversized = 'x' * 12000
    model_factory_calls = []

    def never_called_model_factory(_name: str, _api_key: str):
        model_factory_calls.append(1)
        return _FakeModel(_FakeSummaryDocument())

    with pytest.raises(ValueError, match='Unsupported summary_length'):
        ai.summarize_document(
            doc_id='doc1',
            model_name='gpt-4.1',
            api_key='k',
            summary_length=oversized,
            model_factory=never_called_model_factory,
        )

    assert model_factory_calls == []


def test_blank_summary_length_defaults_to_short(monkeypatch, tmp_path):
    _patch_runtime(monkeypatch, tmp_path)
    summary_doc = _FakeSummaryDocument()
    captured = {}

    def summarize_call(_model, _source_doc, options):
        captured['summary_length'] = options.summary_length
        return summary_doc

    result = ai.summarize_document(
        doc_id='doc1',
        model_name='gpt-4.1',
        api_key='k',
        summary_length='',
        model_factory=lambda _n, _k: SimpleNamespace(),
        summarize_call=summarize_call,
    )

    assert captured['summary_length'] == ai._SUMMARY_LENGTH_MAP['short']
    assert result['summaryLength'] == 'short'


def test_whitespace_only_summary_length_defaults_to_short(monkeypatch, tmp_path):
    _patch_runtime(monkeypatch, tmp_path)
    summary_doc = _FakeSummaryDocument()
    captured = {}

    def summarize_call(_model, _source_doc, options):
        captured['summary_length'] = options.summary_length
        return summary_doc

    result = ai.summarize_document(
        doc_id='doc1',
        model_name='gpt-4.1',
        api_key='k',
        summary_length='   ',
        model_factory=lambda _n, _k: SimpleNamespace(),
        summarize_call=summarize_call,
    )

    assert captured['summary_length'] == ai._SUMMARY_LENGTH_MAP['short']
    assert result['summaryLength'] == 'short'


def test_output_name_with_trailing_spaces_is_normalized_and_save_occurs_once(monkeypatch, tmp_path):
    runtime = _patch_runtime(monkeypatch, tmp_path, source_filename='source.docx')
    summary_doc = _FakeSummaryDocument()
    model = _FakeModel(summary_doc)

    assert summary_doc.saved == []

    result = ai.summarize_document(
        doc_id='doc-sec',
        model_name='gpt-4.1-mini',
        api_key='k',
        output_name=' final-summary  ',
        model_factory=lambda _n, _k: model,
    )

    expected_output_path = runtime.source_path.parent / 'final-summary.docx'
    assert len(summary_doc.saved) == 1
    assert summary_doc.saved[0] == (str(expected_output_path), ai.aw.SaveFormat.DOCX)
    assert result['outputFilename'] == 'final-summary.docx'
