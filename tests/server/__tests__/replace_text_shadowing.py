import inspect

import pytest

import mcp_server as srv


def test_register_tools_keeps_replace_text_wrapper_signature_and_routing(monkeypatch):
    registered = {}
    call = {}

    def fake_tool(*, description):
        def decorator(fn):
            registered[fn.__name__] = fn
            return fn

        return decorator

    def fake_tool_replace_text(**kwargs):
        call.update(kwargs)
        return {'count': 7}

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    monkeypatch.setattr(srv, 'tool_replace_text', fake_tool_replace_text)

    before_count = len(registered)
    srv.register_tools()
    after_count = len(registered)

    assert before_count == 0
    assert after_count > 0
    assert 'replace_text' in registered

    wrapper = registered['replace_text']
    assert wrapper is srv._replace_text_wrapper
    assert srv._replace_text_wrapper.__name__ == 'replace_text'
    assert list(inspect.signature(wrapper).parameters.keys()) == [
        'doc_id',
        'find_text',
        'replace_text',
        'search_text',
        'replacement_text',
        'replace_all',
        'case_sensitive',
        'whole_word',
        'use_regex',
    ]

    result = wrapper(
        doc_id='doc-1',
        find_text='alpha',
        replace_text='beta',
        replace_all=False,
        case_sensitive=True,
        whole_word=True,
        use_regex=True,
    )

    assert result == {'count': 7}
    assert call == {
        'doc_id': 'doc-1',
        'find_text': 'alpha',
        'replace_text': 'beta',
        'search_text': None,
        'replacement_text': None,
        'replace_all': False,
        'case_sensitive': True,
        'whole_word': True,
        'use_regex': True,
    }


def test_replace_text_wrapper_rejects_unexpected_argument_name(monkeypatch):
    registered = {}

    def fake_tool(*, description):
        def decorator(fn):
            registered[fn.__name__] = fn
            return fn

        return decorator

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    srv.register_tools()

    wrapper = registered['replace_text']
    with pytest.raises(TypeError, match="unexpected keyword argument 'replace_txt'"):
        wrapper(doc_id='doc-2', find_text='a', replace_txt='b')


def test_replace_text_wrapper_forwards_adversarial_alias_payloads_without_bypass(monkeypatch):
    registered = {}
    call = {}

    def fake_tool(*, description):
        def decorator(fn):
            registered[fn.__name__] = fn
            return fn

        return decorator

    def fake_tool_replace_text(**kwargs):
        call.update(kwargs)
        return {'count': 1}

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    monkeypatch.setattr(srv, 'tool_replace_text', fake_tool_replace_text)
    srv.register_tools()

    wrapper = registered['replace_text']
    attack_search = '<script>alert(1)</script>${payload}../\x00' + ('A' * 11001)
    attack_replacement = 'safe🧪\u200b\u202e'

    result = wrapper(
        doc_id='doc-3',
        search_text=attack_search,
        replacement_text=attack_replacement,
        replace_all=True,
        case_sensitive=False,
        whole_word=False,
        use_regex=False,
    )

    assert result == {'count': 1}
    assert call['search_text'] == attack_search
    assert call['replacement_text'] == attack_replacement
    assert call['find_text'] is None
    assert call['replace_text'] is None


def test_replace_text_wrapper_preserves_boundary_alias_validation_errors():
    with pytest.raises(
        ValueError,
        match='Provide exactly one search term alias: either "find_text" or "search_text".',
    ):
        srv._replace_text_wrapper(
            doc_id='doc-4',
            find_text='alpha',
            search_text='beta',
            replace_text='gamma',
        )

    with pytest.raises(
        ValueError,
        match='Provide exactly one replacement term alias: either "replace_text" or "replacement_text".',
    ):
        srv._replace_text_wrapper(
            doc_id='doc-5',
            find_text='alpha',
            replace_text='beta',
            replacement_text='gamma',
        )

    with pytest.raises(ValueError, match='search text must be a non-empty string'):
        srv._replace_text_wrapper(doc_id='doc-6', find_text='', replace_text='ok')


def test_replace_text_wrapper_forwards_non_boolean_boundary_values_unchanged(monkeypatch):
    call = {}

    def fake_tool_replace_text(**kwargs):
        call.update(kwargs)
        return {'count': 3}

    monkeypatch.setattr(srv, 'tool_replace_text', fake_tool_replace_text)

    result = srv._replace_text_wrapper(
        doc_id='doc-7',
        find_text='token',
        replace_text='value',
        replace_all='yes',
        case_sensitive=-0.0,
        whole_word=float('inf'),
        use_regex='false',
    )

    assert result == {'count': 3}
    assert call['replace_all'] == 'yes'
    assert call['case_sensitive'] == -0.0
    assert call['whole_word'] == float('inf')
    assert call['use_regex'] == 'false'


def test_replace_text_wrapper_rejects_malformed_kwargs_and_positional_overflow():
    with pytest.raises(TypeError, match="unexpected keyword argument 'replace__text'"):
        srv._replace_text_wrapper(
            doc_id='doc-8', find_text='alpha', replace_text='beta', replace__text='evil'
        )

    with pytest.raises(TypeError, match='takes from 1 to 9 positional arguments but 10 were given'):
        srv._replace_text_wrapper(
            'doc-9', 'a', 'b', None, None, True, False, False, False, 'overflow'
        )
