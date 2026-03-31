import inspect

import pytest

import mcp_server as srv
from core.utils.docs_util import docx_path, init_data_dir


def test_tool_join_paragraph_runs_delegates_and_returns_metadata_unchanged(monkeypatch):
    captured = {}
    expected = {
        'paragraph_index': 7,
        'run_count_before': 5,
        'run_count_after': 2,
        'changed': True,
    }

    def fake_join_paragraph_runs(**kwargs):
        captured.update(kwargs)
        return expected

    monkeypatch.setattr(srv._content, 'join_paragraph_runs', fake_join_paragraph_runs)

    result = srv.tool_join_paragraph_runs(
        doc_id='doc-123',
        paragraph_index=7,
        ignore_redundant=True,
        ignore_insignificant=False,
        ignore_spacing=None,
    )

    assert result is expected
    assert captured == {
        'doc_id': 'doc-123',
        'paragraph_index': 7,
        'ignore_redundant': True,
        'ignore_insignificant': False,
        'ignore_spacing': None,
    }


def test_tool_join_paragraph_runs_forwards_adversarial_optional_values_as_is(monkeypatch):
    captured = {}

    def fake_join_paragraph_runs(**kwargs):
        captured.update(kwargs)
        return {'ok': True, 'paragraph_index': kwargs['paragraph_index']}

    monkeypatch.setattr(srv._content, 'join_paragraph_runs', fake_join_paragraph_runs)

    result = srv.tool_join_paragraph_runs(
        doc_id='',
        paragraph_index=0,
        ignore_redundant='0',
        ignore_insignificant='',
        ignore_spacing='<script>alert(1)</script>${injection}\x00',
    )

    assert result == {'ok': True, 'paragraph_index': 0}
    assert captured == {
        'doc_id': '',
        'paragraph_index': 0,
        'ignore_redundant': '0',
        'ignore_insignificant': '',
        'ignore_spacing': '<script>alert(1)</script>${injection}\x00',
    }


@pytest.mark.parametrize('paragraph_index', [-1, 0, (2**53) - 1])
def test_tool_join_paragraph_runs_forwards_boundary_paragraph_index_values(
    monkeypatch, paragraph_index
):
    captured = {}

    def fake_join_paragraph_runs(**kwargs):
        captured.update(kwargs)
        return {'paragraph_index': kwargs['paragraph_index'], 'changed': False}

    monkeypatch.setattr(srv._content, 'join_paragraph_runs', fake_join_paragraph_runs)

    result = srv.tool_join_paragraph_runs(
        doc_id='doc-boundary',
        paragraph_index=paragraph_index,
        ignore_redundant=None,
        ignore_insignificant=None,
        ignore_spacing=None,
    )

    assert result == {'paragraph_index': paragraph_index, 'changed': False}
    assert captured['doc_id'] == 'doc-boundary'
    assert captured['paragraph_index'] == paragraph_index


def test_tool_join_paragraph_runs_forwards_oversized_and_type_confused_optionals(monkeypatch):
    captured = {}
    oversized = 'A' * 12000
    unicode_payload = '\u202ezero\u200bwidth\x00😀e\u0301'
    object_payload = {'path': '../', 'template': '${injection}'}

    def fake_join_paragraph_runs(**kwargs):
        captured.update(kwargs)
        return {'ok': True}

    monkeypatch.setattr(srv._content, 'join_paragraph_runs', fake_join_paragraph_runs)

    result = srv.tool_join_paragraph_runs(
        doc_id='doc-malformed',
        paragraph_index=1,
        ignore_redundant=oversized,
        ignore_insignificant=unicode_payload,
        ignore_spacing=object_payload,
    )

    assert result == {'ok': True}
    assert captured == {
        'doc_id': 'doc-malformed',
        'paragraph_index': 1,
        'ignore_redundant': oversized,
        'ignore_insignificant': unicode_payload,
        'ignore_spacing': object_payload,
    }


def test_register_tools_exposes_join_paragraph_runs_surface_and_schema(monkeypatch):
    registrations = []

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    srv.register_tools()

    registration = next(r for r in registrations if r['name'] == 'join_paragraph_runs')
    signature = inspect.signature(registration['fn'])

    assert registration['description'] == (
        'Join runs with the same formatting in a selected paragraph by index'
    )
    assert list(signature.parameters.keys()) == [
        'doc_id',
        'paragraph_index',
        'ignore_redundant',
        'ignore_insignificant',
        'ignore_spacing',
    ]
    assert signature.parameters['doc_id'].default is inspect._empty
    assert signature.parameters['paragraph_index'].default is inspect._empty
    assert signature.parameters['ignore_redundant'].default is None
    assert signature.parameters['ignore_insignificant'].default is None
    assert signature.parameters['ignore_spacing'].default is None


def test_registered_join_paragraph_runs_wrapper_delegates_to_tool_function(monkeypatch):
    registrations = []
    captured = {}
    expected = {
        'paragraph_index': 2,
        'run_count_before': 4,
        'run_count_after': 4,
        'changed': False,
    }

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    def fake_tool_join_paragraph_runs(
        doc_id,
        paragraph_index,
        ignore_redundant=None,
        ignore_insignificant=None,
        ignore_spacing=None,
    ):
        captured.update(
            {
                'doc_id': doc_id,
                'paragraph_index': paragraph_index,
                'ignore_redundant': ignore_redundant,
                'ignore_insignificant': ignore_insignificant,
                'ignore_spacing': ignore_spacing,
            }
        )
        return expected

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    monkeypatch.setattr(srv, 'tool_join_paragraph_runs', fake_tool_join_paragraph_runs)
    srv.register_tools()

    registered = next(r for r in registrations if r['name'] == 'join_paragraph_runs')['fn']
    result = registered(
        doc_id='doc-x',
        paragraph_index=2,
        ignore_redundant=False,
        ignore_insignificant=True,
        ignore_spacing=False,
    )

    assert result == expected
    assert captured == {
        'doc_id': 'doc-x',
        'paragraph_index': 2,
        'ignore_redundant': False,
        'ignore_insignificant': True,
        'ignore_spacing': False,
    }


def test_registered_join_paragraph_runs_wrapper_rejects_missing_required_argument(monkeypatch):
    registrations = []
    called = {'tool_join_paragraph_runs': 0}

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    def fake_tool_join_paragraph_runs(
        doc_id,
        paragraph_index,
        ignore_redundant=None,
        ignore_insignificant=None,
        ignore_spacing=None,
    ):
        called['tool_join_paragraph_runs'] += 1
        return {
            'doc_id': doc_id,
            'paragraph_index': paragraph_index,
            'ignore_redundant': ignore_redundant,
            'ignore_insignificant': ignore_insignificant,
            'ignore_spacing': ignore_spacing,
        }

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    monkeypatch.setattr(srv, 'tool_join_paragraph_runs', fake_tool_join_paragraph_runs)
    srv.register_tools()

    registered = next(r for r in registrations if r['name'] == 'join_paragraph_runs')['fn']

    with pytest.raises(
        TypeError, match="missing 1 required positional argument: 'paragraph_index'"
    ):
        registered(doc_id='doc-y')

    assert called['tool_join_paragraph_runs'] == 0


def test_registered_join_paragraph_runs_wrapper_rejects_unexpected_keyword(monkeypatch):
    registrations = []
    called = {'tool_join_paragraph_runs': 0}

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    def fake_tool_join_paragraph_runs(
        doc_id,
        paragraph_index,
        ignore_redundant=None,
        ignore_insignificant=None,
        ignore_spacing=None,
    ):
        called['tool_join_paragraph_runs'] += 1
        return {'ok': True}

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    monkeypatch.setattr(srv, 'tool_join_paragraph_runs', fake_tool_join_paragraph_runs)
    srv.register_tools()

    registered = next(r for r in registrations if r['name'] == 'join_paragraph_runs')['fn']

    with pytest.raises(TypeError, match="unexpected keyword argument 'ignore_space'"):
        registered(doc_id='doc-z', paragraph_index=1, ignore_space=True)

    assert called['tool_join_paragraph_runs'] == 0


def test_registered_join_paragraph_runs_wrapper_invokes_full_path_and_persists_changes(
    monkeypatch, tmp_path
):
    pytest.importorskip('aspose.words')
    import aspose.words as aw

    init_data_dir(tmp_path)
    doc_id = 'join-runs-full-path'

    doc = aw.Document()
    body = doc.first_section.body
    body.remove_all_children()

    paragraph = aw.Paragraph(doc)
    paragraph.append_child(aw.Run(doc, 'Hello'))
    paragraph.append_child(aw.Run(doc, 'world'))
    body.append_child(paragraph)

    paragraph_index = 0
    doc.save(str(docx_path(doc_id)))

    registrations = []

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    srv.register_tools()

    registered = next(r for r in registrations if r['name'] == 'join_paragraph_runs')['fn']
    before_doc = aw.Document(str(docx_path(doc_id)))
    before_paragraph = before_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[
        paragraph_index
    ].as_paragraph()
    run_count_before = int(before_paragraph.runs.count)
    result = registered(
        doc_id=doc_id,
        paragraph_index=paragraph_index,
        ignore_redundant=True,
        ignore_insignificant=True,
        ignore_spacing=True,
    )

    saved_doc = aw.Document(str(docx_path(doc_id)))
    saved_paragraph = saved_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[
        paragraph_index
    ].as_paragraph()
    run_count_after = int(saved_paragraph.runs.count)

    assert result == {
        'paragraph_index': paragraph_index,
        'run_count_before': run_count_before,
        'run_count_after': run_count_after,
        'changed': run_count_after != run_count_before,
    }


def test_join_paragraph_runs_wrapper_handles_malformed_optionals_and_idempotency(
    monkeypatch, tmp_path
):
    pytest.importorskip('aspose.words')
    import aspose.words as aw

    init_data_dir(tmp_path)
    doc_id = 'join-runs-adv-full-path'

    doc = aw.Document()
    body = doc.first_section.body
    body.remove_all_children()

    paragraph = aw.Paragraph(doc)
    paragraph.append_child(aw.Run(doc, '<script>alert(1)</script>'))
    paragraph.append_child(aw.Run(doc, 'zero\x00byte-😀-e\u0301'))
    body.append_child(paragraph)
    doc.save(str(docx_path(doc_id)))

    registrations = []

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    srv.register_tools()

    registered = next(r for r in registrations if r['name'] == 'join_paragraph_runs')['fn']

    before_first_doc = aw.Document(str(docx_path(doc_id)))
    before_first_paragraph = before_first_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[
        0
    ].as_paragraph()
    first_run_count_before = int(before_first_paragraph.runs.count)

    first = registered(
        doc_id=doc_id,
        paragraph_index=0,
        ignore_redundant='0',
        ignore_insignificant='',
        ignore_spacing={'path': '../', 'template': '${injection}'},
    )

    after_first_doc = aw.Document(str(docx_path(doc_id)))
    after_first_paragraph = after_first_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[
        0
    ].as_paragraph()
    first_run_count_after = int(after_first_paragraph.runs.count)

    second = registered(
        doc_id=doc_id,
        paragraph_index=0,
        ignore_redundant='0',
        ignore_insignificant='',
        ignore_spacing={'path': '../', 'template': '${injection}'},
    )

    after_second_doc = aw.Document(str(docx_path(doc_id)))
    after_second_paragraph = after_second_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[
        0
    ].as_paragraph()
    second_run_count_after = int(after_second_paragraph.runs.count)

    assert first == {
        'paragraph_index': 0,
        'run_count_before': first_run_count_before,
        'run_count_after': first_run_count_after,
        'changed': first_run_count_after != first_run_count_before,
    }
    assert second == {
        'paragraph_index': 0,
        'run_count_before': first_run_count_after,
        'run_count_after': second_run_count_after,
        'changed': second_run_count_after != first_run_count_after,
    }
    assert second_run_count_after == second['run_count_after']


def test_registered_join_paragraph_runs_wrapper_full_path_rejects_out_of_range_and_preserves_state(
    monkeypatch, tmp_path
):
    pytest.importorskip('aspose.words')
    import aspose.words as aw

    init_data_dir(tmp_path)
    doc_id = 'join-runs-boundary-full-path'

    doc = aw.Document()
    body = doc.first_section.body
    body.remove_all_children()

    paragraph = aw.Paragraph(doc)
    paragraph.append_child(aw.Run(doc, 'Hello'))
    paragraph.append_child(aw.Run(doc, 'world'))
    body.append_child(paragraph)
    doc.save(str(docx_path(doc_id)))

    registrations = []

    def fake_tool(*, description):
        def decorator(fn):
            registrations.append({'name': fn.__name__, 'description': description, 'fn': fn})
            return fn

        return decorator

    monkeypatch.setattr(srv.mcp, 'tool', fake_tool)
    srv.register_tools()

    registered = next(r for r in registrations if r['name'] == 'join_paragraph_runs')['fn']

    before_doc = aw.Document(str(docx_path(doc_id)))
    paragraphs = before_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    paragraph_count = int(paragraphs.count)
    out_of_range_index = paragraph_count
    before_paragraph = before_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()
    before_run_count = int(before_paragraph.runs.count)

    with pytest.raises(IndexError, match='paragraph_index out of range'):
        registered(
            doc_id=doc_id,
            paragraph_index=out_of_range_index,
            ignore_redundant=None,
            ignore_insignificant=None,
            ignore_spacing=None,
        )

    after_doc = aw.Document(str(docx_path(doc_id)))
    after_paragraph = after_doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()
    after_run_count = int(after_paragraph.runs.count)

    assert after_run_count == before_run_count
