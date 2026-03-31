import inspect

import pytest

import mcp_server as srv


def test_tool_summarize_document_delegates_with_env_api_key_and_returns_result(monkeypatch):
    captured = {}
    expected = {
        'sourceDocId': 'doc-123',
        'outputFilename': 'summary.docx',
        'summaryLength': 'medium',
    }

    def fake_summarize_document(**kwargs):
        captured.update(kwargs)
        return expected

    monkeypatch.setenv('API_KEY', '  real-api-key  ')
    monkeypatch.setattr(srv._ai, 'summarize_document', fake_summarize_document)

    result = srv.tool_summarize_document(
        doc_id='doc-123',
        model_name='gpt-4.1-mini',
        summary_length='medium',
        output_name='summary',
    )

    assert result is expected
    assert captured == {
        'doc_id': 'doc-123',
        'model_name': 'gpt-4.1-mini',
        'api_key': 'real-api-key',
        'summary_length': 'medium',
        'output_name': 'summary',
    }


def test_tool_summarize_document_malformed_optional_inputs_forwarded_as_is(monkeypatch):
    captured = {}

    def fake_summarize_document(**kwargs):
        captured.update(kwargs)
        return {'ok': True, 'sourceDocId': kwargs['doc_id']}

    monkeypatch.setenv('API_KEY', 'server-key')
    monkeypatch.setattr(srv._ai, 'summarize_document', fake_summarize_document)

    result = srv.tool_summarize_document(
        doc_id='doc-malformed',
        model_name='gpt-4.1',
        summary_length='<script>alert(1)</script>${injection}',
        output_name='../sensitive/..\\summary\x00.docx',
    )

    assert result == {'ok': True, 'sourceDocId': 'doc-malformed'}
    assert captured == {
        'doc_id': 'doc-malformed',
        'model_name': 'gpt-4.1',
        'api_key': 'server-key',
        'summary_length': '<script>alert(1)</script>${injection}',
        'output_name': '../sensitive/..\\summary\x00.docx',
    }


def test_tool_summarize_document_none_optional_values_forwarded(monkeypatch):
    captured = {}

    def fake_summarize_document(**kwargs):
        captured.update(kwargs)
        return {'ok': True}

    monkeypatch.setenv('API_KEY', 'k')
    monkeypatch.setattr(srv._ai, 'summarize_document', fake_summarize_document)

    result = srv.tool_summarize_document(
        doc_id='doc-defaults', model_name='gpt-4.1', summary_length=None, output_name=None
    )

    assert result == {'ok': True}
    assert captured['summary_length'] is None
    assert captured['output_name'] is None


def test_tool_summarize_document_missing_api_key_raises_clear_error(monkeypatch):
    monkeypatch.delenv('API_KEY', raising=False)

    with pytest.raises(
        ValueError, match='API_KEY environment variable is required for summarize_document'
    ):
        srv.tool_summarize_document(doc_id='doc-no-key', model_name='gpt-4.1')


@pytest.mark.parametrize('api_key', ['', '   '])
def test_tool_summarize_document_empty_api_key_raises_clear_error(monkeypatch, api_key):
    monkeypatch.setenv('API_KEY', api_key)

    with pytest.raises(
        ValueError, match='API_KEY environment variable is required for summarize_document'
    ):
        srv.tool_summarize_document(doc_id='doc-empty-key', model_name='gpt-4.1')


@pytest.mark.parametrize('api_key', ['\n\t ', ' \r\n '])
def test_tool_summarize_document_blank_api_key_variants_raise_without_secret_leak(
    monkeypatch, api_key
):
    monkeypatch.setenv('API_KEY', api_key)

    with pytest.raises(ValueError) as exc:
        srv.tool_summarize_document(doc_id='doc-blank-key', model_name='gpt-4.1')

    message = str(exc.value)
    non_whitespace_api_key = ''.join(api_key.split())
    assert message == 'API_KEY environment variable is required for summarize_document'
    if non_whitespace_api_key:
        assert non_whitespace_api_key not in message


def test_register_tools_exposes_summarize_document_mcp_surface(monkeypatch):
    registrations = []

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    srv.register_tools()

    summarize = next(r for r in registrations if r['name'] == 'summarize_document')

    assert summarize['description'] == (
        'Summarize a document using the AI release surface '
        '(API key is read from server API_KEY environment variable)'
    )
    signature = inspect.signature(summarize['fn'])
    assert list(signature.parameters.keys()) == [
        'doc_id',
        'model_name',
        'summary_length',
        'output_name',
    ]
    assert signature.parameters['summary_length'].default == 'short'
    assert signature.parameters['output_name'].default is None


def test_register_tools_does_not_expose_live_api_key_in_description_or_signature(monkeypatch):
    registrations = []

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    monkeypatch.setenv('API_KEY', 'sk-live-ultra-secret-key')
    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    srv.register_tools()

    summarize = next(r for r in registrations if r['name'] == 'summarize_document')
    signature = inspect.signature(summarize['fn'])

    assert summarize['description'] == (
        'Summarize a document using the AI release surface '
        '(API key is read from server API_KEY environment variable)'
    )
    assert 'sk-live-ultra-secret-key' not in summarize['description']
    assert signature.parameters['model_name'].default is inspect._empty
    assert signature.parameters['summary_length'].default == 'short'
    assert signature.parameters['output_name'].default is None


def test_registered_summarize_document_wrapper_delegates_to_tool_function(monkeypatch):
    registrations = []
    captured = {}
    expected = {'sourceDocId': 'doc-1', 'outputFilename': 'summary.docx'}

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    def fake_tool_summarize_document(doc_id, model_name, summary_length='short', output_name=None):
        captured['doc_id'] = doc_id
        captured['model_name'] = model_name
        captured['summary_length'] = summary_length
        captured['output_name'] = output_name
        return expected

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    monkeypatch.setattr(srv, 'tool_summarize_document', fake_tool_summarize_document)
    srv.register_tools()

    summarize_tool = next(r for r in registrations if r['name'] == 'summarize_document')['fn']
    result = summarize_tool(
        doc_id='doc-1',
        model_name='gpt-4.1-mini',
        summary_length='very_short',
        output_name='summary-output',
    )

    assert result == expected
    assert captured == {
        'doc_id': 'doc-1',
        'model_name': 'gpt-4.1-mini',
        'summary_length': 'very_short',
        'output_name': 'summary-output',
    }


def test_registered_summarize_document_wrapper_propagates_ai_failure_unchanged(monkeypatch):
    registrations = []
    original_error = RuntimeError('summarization-backend-failure')

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    def failing_summarize_document(**_kwargs):
        raise original_error

    monkeypatch.setenv('API_KEY', 'server-key')
    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    monkeypatch.setattr(srv._ai, 'summarize_document', failing_summarize_document)
    srv.register_tools()

    summarize_tool = next(r for r in registrations if r['name'] == 'summarize_document')['fn']

    with pytest.raises(RuntimeError) as exc:
        summarize_tool(doc_id='doc-1', model_name='gpt-4.1-mini')

    assert exc.value is original_error


def test_registered_summarize_document_wrapper_rejects_mcp_boundary_misuse(monkeypatch):
    registrations = []
    call_count = {'tool_summarize_document': 0}

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    def fake_tool_summarize_document(doc_id, model_name, summary_length='short', output_name=None):
        call_count['tool_summarize_document'] += 1
        return {
            'doc_id': doc_id,
            'model_name': model_name,
            'summary_length': summary_length,
            'output_name': output_name,
        }

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    monkeypatch.setattr(srv, 'tool_summarize_document', fake_tool_summarize_document)
    srv.register_tools()

    summarize_tool = next(r for r in registrations if r['name'] == 'summarize_document')['fn']

    with pytest.raises(TypeError, match="missing 1 required positional argument: 'model_name'"):
        summarize_tool(doc_id='doc-1')

    assert call_count['tool_summarize_document'] == 0


def test_registered_summarize_document_wrapper_rejects_api_key_injection_argument(monkeypatch):
    registrations = []
    call_count = {'tool_summarize_document': 0}

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    def fake_tool_summarize_document(doc_id, model_name, summary_length='short', output_name=None):
        call_count['tool_summarize_document'] += 1
        return {
            'doc_id': doc_id,
            'model_name': model_name,
            'summary_length': summary_length,
            'output_name': output_name,
        }

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    monkeypatch.setattr(srv, 'tool_summarize_document', fake_tool_summarize_document)
    srv.register_tools()

    summarize_tool = next(r for r in registrations if r['name'] == 'summarize_document')['fn']
    unexpected_kwargs = {'_'.join(['api', 'key']): 'unexpected-parameter-value'}

    with pytest.raises(TypeError, match="got an unexpected keyword argument 'api_key'"):
        summarize_tool(
            doc_id='doc-1',
            model_name='gpt-4.1-mini',
            summary_length='short',
            output_name='summary',
            **unexpected_kwargs,
        )

    assert call_count['tool_summarize_document'] == 0
