import datetime as dt
import inspect
import types
from pathlib import Path
from typing import Any, cast

import pytest

pytest.importorskip('aspose.words')
import mcp_server as srv
from core import comparison as _comparison


def test_compare_documents_defaults_to_ms_word_list_definition_behavior(monkeypatch, tmp_path):
    source_path = tmp_path / 'source.docx'
    target_path = tmp_path / 'target.docx'
    events: list[tuple[str, object]] = []

    class FakeAdvancedCompareOptions:
        def __setattr__(self, name: str, value: object) -> None:
            events.append((f'assign:{name}', value))
            object.__setattr__(self, name, value)

    class FakeCompareOptions:
        def __init__(self) -> None:
            events.append(('compare_options', None))
            self.advanced_options = FakeAdvancedCompareOptions()

    class FakeDocument:
        def __init__(self, path: str) -> None:
            events.append(('document', path))
            self.path = path

        def compare(self, **kwargs) -> None:
            events.append(('compare', kwargs))

        def save(self, path: str) -> None:
            events.append(('save', path))

    def fake_ensure_path(doc_id: str) -> Path:
        return source_path if doc_id == 'source-id' else target_path

    monkeypatch.setattr(_comparison, 'ensure_path', fake_ensure_path)
    monkeypatch.setattr(_comparison.aw, 'Document', FakeDocument)
    monkeypatch.setattr(
        _comparison.aw,
        'comparing',
        types.SimpleNamespace(CompareOptions=FakeCompareOptions),
    )

    result = _comparison.compare_documents(
        'source-id',
        'target-id',
        date_time_iso='2026-09-15T10:11:12',
    )

    assert result is True
    compare_event = next(event for event in events if event[0] == 'compare')
    compare_kwargs = cast(dict[str, Any], compare_event[1])
    assert compare_kwargs['author'] == 'MCP'
    assert compare_kwargs['date_time'] == dt.datetime(2026, 9, 15, 10, 11, 12)
    assert compare_kwargs['options'].advanced_options.compare_list_definitions is False
    assert events == [
        ('document', str(source_path)),
        ('document', str(target_path)),
        ('compare_options', None),
        ('assign:compare_list_definitions', False),
        ('compare', compare_kwargs),
        ('save', str(source_path)),
    ]


def test_compare_documents_assigns_list_definition_option_and_passes_options(monkeypatch, tmp_path):
    source_path = tmp_path / 'source.docx'
    target_path = tmp_path / 'target.docx'
    documents = []
    options_instances = []
    compare_calls = []

    class FakeAdvancedCompareOptions:
        def __init__(self) -> None:
            self.compare_list_definitions = None

    class FakeCompareOptions:
        def __init__(self) -> None:
            self.advanced_options = FakeAdvancedCompareOptions()
            options_instances.append(self)

    class FakeDocument:
        def __init__(self, path: str) -> None:
            self.path = path
            self.saved_path = None
            documents.append(self)

        def compare(self, **kwargs) -> None:
            compare_calls.append((self, kwargs))

        def save(self, path: str) -> None:
            self.saved_path = path

    monkeypatch.setattr(
        _comparison,
        'ensure_path',
        lambda doc_id: source_path if doc_id == 'source-id' else target_path,
    )
    monkeypatch.setattr(_comparison.aw, 'Document', FakeDocument)
    monkeypatch.setattr(
        _comparison.aw,
        'comparing',
        types.SimpleNamespace(CompareOptions=FakeCompareOptions),
    )

    _comparison.compare_documents(
        doc_id='source-id',
        compare_to_doc_id='target-id',
        author='Reviewer',
        date_time_iso='2026-09-15T10:11:12Z',
        compare_list_definitions=True,
    )

    assert len(documents) == 2
    source_doc, target_doc = documents
    assert len(options_instances) == 1
    options = options_instances[0]
    assert options.advanced_options.compare_list_definitions is True
    assert compare_calls == [
        (
            source_doc,
            {
                'document': target_doc,
                'author': 'Reviewer',
                'date_time': dt.datetime(2026, 9, 15, 10, 11, 12, tzinfo=dt.timezone.utc),
                'options': options,
            },
        )
    ]
    assert source_doc.saved_path == str(source_path)


@pytest.mark.parametrize(
    'author',
    ['', '   ', None],
)
def test_compare_documents_rejects_empty_author_before_aspose_calls(monkeypatch, author):
    aspose_calls = []

    class FakeDocument:
        def __init__(self, path: str) -> None:
            aspose_calls.append(path)

    monkeypatch.setattr(_comparison.aw, 'Document', FakeDocument)

    with pytest.raises(ValueError, match='author must be a non-empty string'):
        _comparison.compare_documents('source-id', 'target-id', author=author)

    assert aspose_calls == []


@pytest.mark.parametrize(
    ('date_time_iso', 'expected_message'),
    [
        ('', 'date_time_iso must be a non-empty ISO 8601 datetime string'),
        ('   ', 'date_time_iso must be a non-empty ISO 8601 datetime string'),
        ('not-a-date', 'date_time_iso must be a valid ISO 8601 datetime string'),
    ],
)
def test_compare_documents_rejects_invalid_date_before_aspose_calls(
    monkeypatch, date_time_iso, expected_message
):
    aspose_calls = []

    class FakeDocument:
        def __init__(self, path: str) -> None:
            aspose_calls.append(path)

    monkeypatch.setattr(_comparison.aw, 'Document', FakeDocument)

    with pytest.raises(ValueError, match=expected_message):
        _comparison.compare_documents('source-id', 'target-id', date_time_iso=date_time_iso)

    assert aspose_calls == []


@pytest.mark.parametrize('compare_list_definitions', ['true', '', None, 1])
def test_compare_documents_rejects_non_boolean_list_definition_option_before_aspose_calls(
    monkeypatch, compare_list_definitions
):
    aspose_calls = []

    class FakeDocument:
        def __init__(self, path: str) -> None:
            aspose_calls.append(path)

    monkeypatch.setattr(_comparison.aw, 'Document', FakeDocument)

    with pytest.raises(ValueError, match='compare_list_definitions must be a boolean'):
        _comparison.compare_documents(
            'source-id',
            'target-id',
            compare_list_definitions=compare_list_definitions,
        )

    assert aspose_calls == []


def test_compare_documents_does_not_suppress_aspose_compare_failures(monkeypatch, tmp_path):
    source_path = tmp_path / 'source.docx'
    target_path = tmp_path / 'target.docx'

    class FakeCompareOptions:
        def __init__(self) -> None:
            self.advanced_options = types.SimpleNamespace(compare_list_definitions=None)

    class FakeDocument:
        def __init__(self, path: str) -> None:
            self.path = path
            self.saved = False

        def compare(self, **kwargs) -> None:
            raise RuntimeError('compare failed visibly')

        def save(self, path: str) -> None:
            self.saved = True

    monkeypatch.setattr(
        _comparison,
        'ensure_path',
        lambda doc_id: source_path if doc_id == 'source-id' else target_path,
    )
    monkeypatch.setattr(_comparison.aw, 'Document', FakeDocument)
    monkeypatch.setattr(
        _comparison.aw,
        'comparing',
        types.SimpleNamespace(CompareOptions=FakeCompareOptions),
    )

    with pytest.raises(RuntimeError, match='compare failed visibly'):
        _comparison.compare_documents(
            'source-id',
            'target-id',
            date_time_iso='2026-09-15T10:11:12',
        )


def test_tool_compare_documents_forwards_all_parameters(monkeypatch):
    compare_calls = []

    def fake_compare_documents(
        doc_id: str,
        compare_to_doc_id: str,
        author: str = 'MCP',
        date_time_iso: str | None = None,
        compare_list_definitions: bool = False,
    ) -> bool:
        compare_calls.append(
            {
                'doc_id': doc_id,
                'compare_to_doc_id': compare_to_doc_id,
                'author': author,
                'date_time_iso': date_time_iso,
                'compare_list_definitions': compare_list_definitions,
            }
        )
        return True

    monkeypatch.setattr(srv._comparison, 'compare_documents', fake_compare_documents)

    response = srv.tool_compare_documents(
        doc_id='source-id',
        compare_to_doc_id='target-id',
        author='Reviewer',
        date_time_iso='2026-09-15T10:11:12',
        compare_list_definitions=True,
    )

    assert response == {}
    assert compare_calls == [
        {
            'doc_id': 'source-id',
            'compare_to_doc_id': 'target-id',
            'author': 'Reviewer',
            'date_time_iso': '2026-09-15T10:11:12',
            'compare_list_definitions': True,
        }
    ]


def test_registered_compare_documents_forwards_all_parameters(monkeypatch):
    captured_tool_functions = {}
    tool_compare_calls = []

    class FakeMcp:
        def tool(self, description=None):
            def capture_tool(function_to_register):
                captured_tool_functions[function_to_register.__name__] = function_to_register
                return function_to_register

            return capture_tool

    def fake_tool_compare_documents(
        doc_id: str,
        compare_to_doc_id: str,
        author: str = 'MCP',
        date_time_iso: str | None = None,
        compare_list_definitions: bool = False,
    ):
        tool_compare_calls.append(
            {
                'doc_id': doc_id,
                'compare_to_doc_id': compare_to_doc_id,
                'author': author,
                'date_time_iso': date_time_iso,
                'compare_list_definitions': compare_list_definitions,
            }
        )
        return {'compared': True}

    monkeypatch.setattr(srv, 'mcp', FakeMcp())
    monkeypatch.setattr(srv, 'tool_compare_documents', fake_tool_compare_documents)

    srv.register_tools()
    registered_compare_documents = captured_tool_functions['compare_documents']
    registered_response = registered_compare_documents(
        doc_id='source-id',
        compare_to_doc_id='target-id',
        author='Reviewer',
        date_time_iso='2026-09-15T10:11:12',
        compare_list_definitions=True,
    )

    assert registered_response == {'compared': True}
    assert tool_compare_calls == [
        {
            'doc_id': 'source-id',
            'compare_to_doc_id': 'target-id',
            'author': 'Reviewer',
            'date_time_iso': '2026-09-15T10:11:12',
            'compare_list_definitions': True,
        }
    ]


def test_document_comparison_uses_explicit_aspose_api_access():
    comparison_source = inspect.getsource(_comparison)

    assert 'getattr(' not in comparison_source
    assert 'hasattr(' not in comparison_source
