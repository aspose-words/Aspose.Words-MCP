import inspect

import pytest

import mcp_server as srv
from core.utils.docs_util import docx_path, init_data_dir


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


def test_registered_join_paragraph_runs_wrapper_invokes_full_path_and_persists_changes(
    monkeypatch, tmp_path
):
    pytest.importorskip('aspose.words')
    import aspose.words as aw

    init_data_dir(tmp_path)
    doc_id = 'join-runs-full-path-registration'

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
