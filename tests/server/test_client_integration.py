import asyncio
import base64
import json
import os
from contextlib import asynccontextmanager
from io import BytesIO
from pathlib import Path

import aspose.words as aw
import pytest
from aspose.pydrawing import Color
from fastmcp import Client
from fastmcp.exceptions import ToolError


@asynccontextmanager
async def _client_session(mcp_client_config):
    client = Client(mcp_client_config)
    async with client:
        yield client


def _run_and_assert_file(result_file_path, run_client):
    asyncio.run(run_client())
    assert os.path.isfile(result_file_path)
    assert os.path.getsize(result_file_path) > 0


async def _create_document_and_get_id(client, arguments=None):
    if arguments is None:
        arguments = {}
    resp = await client.call_tool(name='create_document', arguments=arguments)
    assert hasattr(resp, 'data')
    assert isinstance(resp.data, dict)
    assert 'docId' in resp.data
    return resp.data['docId']


def _runtime_docx_path(doc_id):
    docs_data_dir = os.environ.get('DOCS_DATA_DIR')
    if not docs_data_dir:
        raise RuntimeError('DOCS_DATA_DIR must be set for integration tests')
    return Path(docs_data_dir) / f'{doc_id}.docx'


def _set_adjacent_same_format_runs(doc_id, marker_text):
    doc_path = _runtime_docx_path(doc_id)
    document = aw.Document(str(doc_path))
    body = document.first_section.body
    body.remove_all_children()

    paragraph = aw.Paragraph(document)
    paragraph.append_child(aw.Run(document, marker_text))
    paragraph.append_child(aw.Run(document, 'token'))
    body.append_child(paragraph)

    document.save(str(doc_path))


def _set_adjacent_runs_with_spacing_difference(doc_id, marker_text):
    doc_path = _runtime_docx_path(doc_id)
    document = aw.Document(str(doc_path))
    body = document.first_section.body
    body.remove_all_children()

    paragraph = aw.Paragraph(document)
    marker_run = aw.Run(document, marker_text)
    token_run = aw.Run(document, 'token')
    token_run.font.spacing = 2.0
    paragraph.append_child(marker_run)
    paragraph.append_child(token_run)
    body.append_child(paragraph)

    document.save(str(doc_path))


def _set_runs_with_bold_whitespace_trailing(doc_id, marker_text):
    doc_path = _runtime_docx_path(doc_id)
    document = aw.Document(str(doc_path))
    body = document.first_section.body
    body.remove_all_children()

    paragraph = aw.Paragraph(document)
    paragraph.append_child(aw.Run(document, marker_text))
    paragraph.append_child(aw.Run(document, 'token'))

    trailing_whitespace_run = aw.Run(document, ' ')
    trailing_whitespace_run.font.bold = True
    paragraph.append_child(trailing_whitespace_run)
    body.append_child(paragraph)

    document.save(str(doc_path))


def _paragraph_with_marker(doc_id, marker_text):
    document = aw.Document(str(_runtime_docx_path(doc_id)))
    paragraph_nodes = document.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    for paragraph_index in range(paragraph_nodes.count):
        paragraph = paragraph_nodes[paragraph_index].as_paragraph()
        paragraph_text = paragraph.to_string(aw.SaveFormat.TEXT)
        if marker_text in paragraph_text:
            return paragraph
    raise AssertionError(f'Paragraph with marker not found: {marker_text}')


async def _save_exported_to_file(result_file_path, client, doc_id):
    exported = await client.call_tool(
        name='export_base64', arguments={'doc_id': doc_id, 'fmt': 'docx'}
    )
    assert hasattr(exported, 'data')
    assert isinstance(exported.data, dict)
    assert 'base64' in exported.data
    raw = base64.b64decode(exported.data['base64'])
    assert isinstance(raw, (bytes, bytearray))
    assert len(raw) > 0
    await asyncio.to_thread(Path(result_file_path).write_bytes, raw)


def test_client_creates_document(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Created document: OK'}
            )
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_chart_create_format_export_and_render(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            tools = {tool.name: tool for tool in await client.list_tools()}
            creation = tools['add_chart_end'].input_schema
            assert set(creation['required']) == {'doc_id', 'chart_type', 'categories', 'series'}
            properties = creation['properties']
            assert properties['categories']['items']['type'] == 'string'
            assert properties['series']['items']['type'] == 'object'
            assert 'name' in properties['series']['description']
            assert 'values' in properties['series']['description']
            assert properties['width_points']['default'] == 432
            assert properties['height_points']['default'] == 252
            assert tools['list_charts'].input_schema['required'] == ['doc_id']
            formatting = tools['format_chart_plot_area'].input_schema
            assert set(formatting['required']) == {'doc_id', 'chart_index'}
            for field in ('fill_color', 'gradient', 'line_color', 'line_width_points'):
                assert formatting['properties'][field]['default'] is None
            assert 'from_center' in formatting['properties']['gradient']['description']

            doc_id = await _create_document_and_get_id(client, {'name': 'chart-client.docx'})
            listed = await client.call_tool('list_charts', {'doc_id': doc_id})
            assert listed.data == {'charts': []}
            added = await client.call_tool(
                'add_chart_end',
                {
                    'doc_id': doc_id,
                    'chart_type': ' COLUMN ',
                    'categories': ['A', 'B', 'C'],
                    'series': [{'name': 'Revenue', 'values': [12, -3, 0]}],
                    'width_points': 320,
                    'height_points': 200,
                },
            )
            assert added.data == {'chartIndex': 0}
            listed = await client.call_tool('list_charts', {'doc_id': doc_id})
            assert listed.data == {
                'charts': [
                    {
                        'chartIndex': 0,
                        'series': [{'seriesIndex': 0, 'name': 'Revenue', 'pointCount': 3}],
                    }
                ]
            }
            solid = await client.call_tool(
                'format_chart_plot_area',
                {'doc_id': doc_id, 'chart_index': 0, 'fill_color': '#ABCDEF'},
            )
            assert solid.data == {}
            formatted = await client.call_tool(
                'format_chart_plot_area',
                {
                    'doc_id': doc_id,
                    'chart_index': 0,
                    'gradient': {
                        'color': '336699',
                        'style': 'Diagonal-Up',
                        'variant': ' VARIANT_2 ',
                        'degree': 0.25,
                    },
                    'line_color': '#112233',
                    'line_width_points': 1.5,
                },
            )
            assert formatted.data == {}
            path = _runtime_docx_path(doc_id)
            before = await asyncio.to_thread(path.read_bytes)
            invalid_calls = [
                ('list_charts', {'doc_id': 'missing-chart-document'}, 'not found'),
                (
                    'add_chart_end',
                    {
                        'doc_id': doc_id,
                        'chart_type': 'pie',
                        'categories': ['A'],
                        'series': [{'name': 'X', 'values': [1]}],
                    },
                    'chart_type must be column or waterfall',
                ),
                (
                    'add_chart_end',
                    {
                        'doc_id': doc_id,
                        'chart_type': 'column',
                        'categories': ['A'],
                        'series': [{'name': 'X', 'values': [1, 2]}],
                    },
                    'values must match categories length',
                ),
                (
                    'format_chart_plot_area',
                    {'doc_id': doc_id, 'chart_index': -1, 'fill_color': '112233'},
                    'chart_index out of range',
                ),
                (
                    'format_chart_plot_area',
                    {'doc_id': doc_id, 'chart_index': True, 'fill_color': '112233'},
                    'integer',
                ),
                (
                    'format_chart_plot_area',
                    {'doc_id': doc_id, 'chart_index': 0, 'line_width_points': True},
                    'number',
                ),
                (
                    'format_chart_plot_area',
                    {'doc_id': doc_id, 'chart_index': 0, 'fill_color': ''},
                    'fill_color must be a six-digit RGB',
                ),
                (
                    'format_chart_plot_area',
                    {
                        'doc_id': doc_id,
                        'chart_index': 0,
                        'gradient': {'color': '112233', 'degree': 2},
                    },
                    'degree must be between 0 and 1',
                ),
            ]
            for tool_name, arguments, message in invalid_calls:
                with pytest.raises(ToolError, match=message):
                    await client.call_tool(tool_name, arguments)
                assert await asyncio.to_thread(path.read_bytes) == before

            exported = await client.call_tool('export_base64', {'doc_id': doc_id, 'fmt': 'docx'})
            raw = base64.b64decode(exported.data['base64'])
            document = aw.Document(BytesIO(raw))
            shape = document.get_child_nodes(aw.NodeType.SHAPE, True)[0].as_shape()
            assert (shape.width, shape.height) == (320, 200)
            chart = shape.chart
            assert chart.series.count == 1
            assert chart.series[0].name == 'Revenue'
            assert [chart.series[0].x_values[i].string_value for i in range(3)] == ['A', 'B', 'C']
            assert [chart.series[0].y_values[i].double_value for i in range(3)] == [12, -3, 0]
            fmt = chart.plot_area.format
            assert fmt.fill.fill_type == aw.drawing.FillType.GRADIENT
            assert fmt.fill.gradient_style == aw.drawing.GradientStyle.DIAGONAL_UP
            assert fmt.fill.gradient_variant == aw.drawing.GradientVariant.VARIANT2
            assert fmt.stroke.fore_color.to_argb() == Color.from_argb(255, 17, 34, 51).to_argb()
            assert fmt.stroke.weight == 1.5
            assert fmt.stroke.visible is True
            rendered = await client.call_tool(
                'render_page_base64',
                {'doc_id': doc_id, 'page_index': 0, 'fmt': 'png', 'dpi': 72},
            )
            assert base64.b64decode(rendered.data['base64']).startswith(b'\x89PNG\r\n\x1a\n')
            await asyncio.to_thread(Path(result_file_path).write_bytes, raw)

    _run_and_assert_file(result_file_path, run_client)


def test_waterfall_create_inspect_export_and_reopen(mcp_client_config, result_file_path):
    flags = [True, False, False, True, False, False, True]
    categories = ['Opening', 'Sales', 'Costs', 'Intermediate', 'Income', 'Expense', 'Final']
    values = [100, 20, -10, 110, 15, -5, 120]
    series = [{'name': 'Balance', 'values': values, 'is_subtotal': flags}]
    points = [
        {
            'pointIndex': i,
            'dataPointIndex': i,
            'isTotal': flag,
            'dataPointIsTotal': flag,
        }
        for i, flag in enumerate(flags)
    ]

    async def run_client():
        async with _client_session(mcp_client_config) as client:
            tools = {tool.name: tool for tool in await client.list_tools()}
            creation = tools['add_chart_end'].input_schema['properties']
            assert 'waterfall' in creation['chart_type']['description']
            assert 'is_subtotal' in creation['series']['description']
            schema = tools['get_waterfall_totals'].input_schema
            assert set(schema['required']) == {'doc_id', 'chart_index'}
            props = schema['properties']
            assert props['doc_id']['type'] == 'string'
            assert props['chart_index']['type'] == 'integer'
            assert props['series_index']['type'] == 'integer'
            assert props['series_index']['default'] == 0
            assert props['point_index']['default'] is None
            assert {choice['type'] for choice in props['point_index']['anyOf']} == {
                'integer',
                'null',
            }
            doc_id = await _create_document_and_get_id(client, {'name': 'waterfall-client.docx'})
            added = await client.call_tool(
                'add_chart_end',
                {
                    'doc_id': doc_id,
                    'chart_type': ' WATERFALL ',
                    'categories': categories,
                    'series': series,
                },
            )
            assert added.data == {'chartIndex': 0}
            listed = await client.call_tool('list_charts', {'doc_id': doc_id})
            assert listed.data == {
                'charts': [
                    {
                        'chartIndex': 0,
                        'series': [{'seriesIndex': 0, 'name': 'Balance', 'pointCount': len(flags)}],
                    }
                ]
            }
            path = _runtime_docx_path(doc_id)
            before = await asyncio.to_thread(path.read_bytes)
            expected = {'chartIndex': 0, 'seriesIndex': 0, 'points': points}
            inspected = await client.call_tool(
                'get_waterfall_totals', {'doc_id': doc_id, 'chart_index': 0}
            )
            assert inspected.data == expected
            explicit_defaults = await client.call_tool(
                'get_waterfall_totals',
                {'doc_id': doc_id, 'chart_index': 0, 'series_index': 0, 'point_index': None},
            )
            assert explicit_defaults.data == expected
            for index, point in enumerate(points):
                inspected = await client.call_tool(
                    'get_waterfall_totals',
                    {'doc_id': doc_id, 'chart_index': 0, 'series_index': 0, 'point_index': index},
                )
                assert inspected.data == {**expected, 'points': [point]}
            assert await asyncio.to_thread(path.read_bytes) == before

            for field in ('chart_index', 'series_index', 'point_index'):
                for value, message in (
                    (-1, f'{field} out of range'),
                    (100, f'{field} out of range'),
                    (True, 'integer'),
                    ('', 'integer'),
                    ('0', 'integer'),
                ):
                    with pytest.raises(ToolError, match=message):
                        await client.call_tool(
                            'get_waterfall_totals',
                            {'doc_id': doc_id, 'chart_index': 0, field: value},
                        )
            with pytest.raises(ToolError, match='Document missing-waterfall not found'):
                await client.call_tool(
                    'get_waterfall_totals', {'doc_id': 'missing-waterfall', 'chart_index': 0}
                )
            invalid_series = [
                ({'name': 'X', 'values': values}, 'must contain only name, values and is_subtotal'),
                ({**series[0], 'is_subtotal': []}, 'is_subtotal must match categories length'),
                ({**series[0], 'is_subtotal': ''}, 'is_subtotal must match categories length'),
                ({**series[0], 'is_subtotal': [1] * 7}, 'is_subtotal must contain only booleans'),
                (
                    {**series[0], 'is_subtotal': ['false'] * 7},
                    'is_subtotal must contain only booleans',
                ),
            ]
            for item, message in invalid_series:
                with pytest.raises(ToolError, match=message):
                    await client.call_tool(
                        'add_chart_end',
                        {
                            'doc_id': doc_id,
                            'chart_type': 'waterfall',
                            'categories': categories,
                            'series': [item],
                        },
                    )
            with pytest.raises(
                ToolError, match='is_subtotal is only supported for waterfall charts'
            ):
                await client.call_tool(
                    'add_chart_end',
                    {
                        'doc_id': doc_id,
                        'chart_type': 'column',
                        'categories': categories,
                        'series': series,
                    },
                )
            assert await asyncio.to_thread(path.read_bytes) == before
            await _save_exported_to_file(result_file_path, client, doc_id)

            await client.call_tool(
                'add_chart_end',
                {
                    'doc_id': doc_id,
                    'chart_type': 'column',
                    'categories': ['X'],
                    'series': [{'name': 'Column', 'values': [1]}],
                },
            )
            before = await asyncio.to_thread(path.read_bytes)
            with pytest.raises(ToolError, match='only supported for waterfall series'):
                await client.call_tool('get_waterfall_totals', {'doc_id': doc_id, 'chart_index': 1})
            assert await asyncio.to_thread(path.read_bytes) == before

    _run_and_assert_file(result_file_path, run_client)
    document = aw.Document(str(result_file_path))
    shape = document.get_child_nodes(aw.NodeType.SHAPE, True)[0].as_shape()
    chart = shape.chart
    assert chart.series.count == 1
    assert (shape.width, shape.height) == (432, 252)
    reopened = chart.series[0]
    assert reopened.series_type == aw.drawing.charts.ChartSeriesType.WATERFALL
    assert reopened.y_values.count == len(flags)
    assert [reopened.x_values[i].string_value for i in range(len(flags))] == categories
    assert [reopened.y_values[i].double_value for i in range(len(flags))] == values
    assert [reopened.is_total(i) for i in range(len(flags))] == flags
    assert [reopened.data_points[i].is_total for i in range(len(flags))] == flags
    assert [reopened.data_points[i].index for i in range(len(flags))] == list(range(len(flags)))


def test_insert_text_start_end(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client, {'name': 'start.docx'})
            await client.call_tool(
                name='insert_text_start', arguments={'doc_id': doc_id, 'text': 'Start '}
            )
            await client.call_tool(
                name='insert_text_end', arguments={'doc_id': doc_id, 'text': ' End'}
            )
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_add_paragraph_and_read(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Line 1'}
            )
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Line 2'}
            )
            resp = await client.call_tool(name='read_paragraphs', arguments={'doc_id': doc_id})
            assert hasattr(resp, 'data')
            assert isinstance(resp.data, dict)
            assert 'paragraphs' in resp.data
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_replace_text(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='insert_text_end',
                arguments={'doc_id': doc_id, 'text': 'a.b Item-10 Item-22 Item-X'},
            )

            plain = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': doc_id,
                    'find_text': 'Item-X',
                    'replace_text': 'Item-Y',
                    'use_regex': False,
                },
            )
            assert hasattr(plain, 'data')
            assert isinstance(plain.data, dict)
            assert plain.data['count'] == 1

            regex = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': doc_id,
                    'search_text': r'Item-\d+',
                    'replacement_text': 'Matched',
                    'use_regex': True,
                },
            )
            assert hasattr(regex, 'data')
            assert isinstance(regex.data, dict)
            assert regex.data['count'] == 2

            zero = await client.call_tool(
                name='replace_text',
                arguments={'doc_id': doc_id, 'find_text': 'zzz', 'replace_text': 'noop'},
            )
            assert hasattr(zero, 'data')
            assert isinstance(zero.data, dict)
            assert zero.data['count'] == 0

            marker_text = 'join-marker:'
            join_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-client.docx'}
            )
            _set_adjacent_same_format_runs(join_doc_id, marker_text)

            joined = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': join_doc_id,
                    'find_text': 'token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                    'ignore_redundant': True,
                    'ignore_insignificant': True,
                    'ignore_spacing': True,
                },
            )
            assert hasattr(joined, 'data')
            assert isinstance(joined.data, dict)
            assert set(joined.data.keys()) == {'count'}
            assert joined.data['count'] == 1

            target_paragraph = _paragraph_with_marker(join_doc_id, marker_text)
            assert target_paragraph.runs.count == 1
            assert target_paragraph.runs[0].text == f'{marker_text}TOKEN'

            spacing_without_ignore_marker_text = 'join-spacing-off-marker:'
            spacing_without_ignore_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-spacing-off-client.docx'}
            )
            _set_adjacent_runs_with_spacing_difference(
                spacing_without_ignore_doc_id, spacing_without_ignore_marker_text
            )

            spacing_without_ignore = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': spacing_without_ignore_doc_id,
                    'find_text': 'token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                },
            )
            assert hasattr(spacing_without_ignore, 'data')
            assert isinstance(spacing_without_ignore.data, dict)
            assert set(spacing_without_ignore.data.keys()) == {'count'}
            assert spacing_without_ignore.data['count'] == 1

            spacing_without_ignore_paragraph = _paragraph_with_marker(
                spacing_without_ignore_doc_id, spacing_without_ignore_marker_text
            )
            assert spacing_without_ignore_paragraph.runs.count == 2
            assert (
                spacing_without_ignore_paragraph.runs[0].text == spacing_without_ignore_marker_text
            )
            assert spacing_without_ignore_paragraph.runs[1].text == 'TOKEN'

            spacing_with_ignore_marker_text = 'join-spacing-on-marker:'
            spacing_with_ignore_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-spacing-on-client.docx'}
            )
            _set_adjacent_runs_with_spacing_difference(
                spacing_with_ignore_doc_id, spacing_with_ignore_marker_text
            )

            spacing_with_ignore = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': spacing_with_ignore_doc_id,
                    'find_text': 'token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                    'ignore_spacing': True,
                },
            )
            assert hasattr(spacing_with_ignore, 'data')
            assert isinstance(spacing_with_ignore.data, dict)
            assert set(spacing_with_ignore.data.keys()) == {'count'}
            assert spacing_with_ignore.data['count'] == 1

            spacing_with_ignore_paragraph = _paragraph_with_marker(
                spacing_with_ignore_doc_id, spacing_with_ignore_marker_text
            )
            assert spacing_with_ignore_paragraph.runs.count == 1
            assert (
                spacing_with_ignore_paragraph.runs[0].text
                == f'{spacing_with_ignore_marker_text}TOKEN'
            )

            insignificant_without_ignore_marker_text = 'join-insignificant-off-marker:'
            insignificant_without_ignore_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-insignificant-off-client.docx'}
            )
            _set_runs_with_bold_whitespace_trailing(
                insignificant_without_ignore_doc_id, insignificant_without_ignore_marker_text
            )

            insignificant_without_ignore = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': insignificant_without_ignore_doc_id,
                    'find_text': 'token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                },
            )
            assert hasattr(insignificant_without_ignore, 'data')
            assert isinstance(insignificant_without_ignore.data, dict)
            assert set(insignificant_without_ignore.data.keys()) == {'count'}
            assert insignificant_without_ignore.data['count'] == 1

            insignificant_without_ignore_paragraph = _paragraph_with_marker(
                insignificant_without_ignore_doc_id, insignificant_without_ignore_marker_text
            )
            assert insignificant_without_ignore_paragraph.runs.count == 2
            assert (
                insignificant_without_ignore_paragraph.runs[0].text
                == f'{insignificant_without_ignore_marker_text}TOKEN'
            )
            assert insignificant_without_ignore_paragraph.runs[1].text == ' '

            insignificant_with_ignore_marker_text = 'join-insignificant-on-marker:'
            insignificant_with_ignore_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-insignificant-on-client.docx'}
            )
            _set_runs_with_bold_whitespace_trailing(
                insignificant_with_ignore_doc_id, insignificant_with_ignore_marker_text
            )

            insignificant_with_ignore = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': insignificant_with_ignore_doc_id,
                    'find_text': 'token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                    'ignore_insignificant': True,
                },
            )
            assert hasattr(insignificant_with_ignore, 'data')
            assert isinstance(insignificant_with_ignore.data, dict)
            assert set(insignificant_with_ignore.data.keys()) == {'count'}
            assert insignificant_with_ignore.data['count'] == 1

            insignificant_with_ignore_paragraph = _paragraph_with_marker(
                insignificant_with_ignore_doc_id, insignificant_with_ignore_marker_text
            )
            assert insignificant_with_ignore_paragraph.runs.count == 1
            assert (
                insignificant_with_ignore_paragraph.runs[0].text
                == f'{insignificant_with_ignore_marker_text}TOKEN '
            )

            zero_match_marker_text = 'join-zero-marker:'
            zero_match_doc_id = await _create_document_and_get_id(
                client, {'name': 'replace-join-zero-match-client.docx'}
            )
            _set_adjacent_same_format_runs(zero_match_doc_id, zero_match_marker_text)

            before_zero_match_paragraph = _paragraph_with_marker(
                zero_match_doc_id, zero_match_marker_text
            )
            before_zero_match_run_count = before_zero_match_paragraph.runs.count
            before_zero_match_run_texts = [
                before_zero_match_paragraph.runs[run_index].text
                for run_index in range(before_zero_match_run_count)
            ]

            zero_match = await client.call_tool(
                name='replace_text',
                arguments={
                    'doc_id': zero_match_doc_id,
                    'find_text': 'absent-token',
                    'replace_text': 'TOKEN',
                    'join_runs': True,
                },
            )
            assert hasattr(zero_match, 'data')
            assert isinstance(zero_match.data, dict)
            assert set(zero_match.data.keys()) == {'count'}
            assert zero_match.data['count'] == 0

            after_zero_match_paragraph = _paragraph_with_marker(
                zero_match_doc_id, zero_match_marker_text
            )
            after_zero_match_run_count = after_zero_match_paragraph.runs.count
            after_zero_match_run_texts = [
                after_zero_match_paragraph.runs[run_index].text
                for run_index in range(after_zero_match_run_count)
            ]
            assert after_zero_match_run_count == before_zero_match_run_count
            assert after_zero_match_run_texts == before_zero_match_run_texts

            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_list_documents_and_get_info(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            lst = await client.call_tool(name='list_documents')
            assert hasattr(lst, 'data')
            assert isinstance(lst.data, dict)
            info = await client.call_tool(name='get_info', arguments={'doc_id': doc_id})
            assert hasattr(info, 'data')
            assert isinstance(info.data, dict)
            assert 'pages' in info.data and 'paragraphs' in info.data
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_add_heading_and_get_outline(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_heading', arguments={'doc_id': doc_id, 'text': 'Title', 'level': 1}
            )
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Body'}
            )
            outline = await client.call_tool(name='get_outline', arguments={'doc_id': doc_id})
            assert hasattr(outline, 'data')
            assert isinstance(outline.data, dict)
            assert 'outline' in outline.data
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_find_text(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='insert_text_end', arguments={'doc_id': doc_id, 'text': 'Foo bar baz'}
            )
            found = await client.call_tool(
                name='find_text', arguments={'doc_id': doc_id, 'text': 'bar'}
            )
            assert hasattr(found, 'data')
            assert isinstance(found.data, dict)
            assert 'matches' in found.data
            assert isinstance(found.data['matches'], list)
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_delete_paragraph(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Line 1'}
            )
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Line 2'}
            )
            await client.call_tool(
                name='delete_paragraph', arguments={'doc_id': doc_id, 'paragraph_index': 1}
            )
            resp = await client.call_tool(name='read_paragraphs', arguments={'doc_id': doc_id})
            assert hasattr(resp, 'data')
            assert isinstance(resp.data, dict)
            assert 'paragraphs' in resp.data
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_page_breaks(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Page 1 Start'}
            )
            await client.call_tool(name='add_page_break_start', arguments={'doc_id': doc_id})
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Middle'}
            )
            await client.call_tool(name='add_page_break_end', arguments={'doc_id': doc_id})
            await client.call_tool(
                name='add_page_break_at_paragraph',
                arguments={'doc_id': doc_id, 'paragraph_index': 0},
            )
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Page 2 End'}
            )
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_insert_list_and_table(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='insert_list_end',
                arguments={'doc_id': doc_id, 'items': ['a', 'b', 'c'], 'kind': 'bullet'},
            )
            tbl = await client.call_tool(
                name='add_table_end',
                arguments={
                    'doc_id': doc_id,
                    'rows': 2,
                    'cols': 2,
                    'data': [['1', '2'], ['3', '4']],
                },
            )
            assert hasattr(tbl, 'data')
            assert isinstance(tbl.data, dict)
            assert 'tableIndex' in tbl.data
            await _save_exported_to_file(result_file_path, client, doc_id)

    _run_and_assert_file(result_file_path, run_client)


def test_export_base64_advanced_docling(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph', arguments={'doc_id': doc_id, 'text': 'Docling export payload'}
            )

            exported = await client.call_tool(
                name='export_base64_advanced',
                arguments={'doc_id': doc_id, 'fmt': 'docling'},
            )
            assert hasattr(exported, 'data')
            assert isinstance(exported.data, dict)
            assert exported.data['ext'] == 'json'
            assert exported.data['mime'] == 'application/json'
            assert 'base64' in exported.data

            raw = base64.b64decode(exported.data['base64'])
            assert isinstance(raw, (bytes, bytearray))
            assert len(raw) > 0

            payload = json.loads(raw.decode('utf-8'))
            assert isinstance(payload, (dict, list))
            assert 'Docling export payload' in json.dumps(payload)

            await asyncio.to_thread(Path(result_file_path).write_bytes, raw)

    _run_and_assert_file(result_file_path, run_client)


def test_export_base64_advanced_pdf_text_shaping(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph',
                arguments={'doc_id': doc_id, 'text': 'PDF advanced export payload'},
            )

            pdf_default_export = await client.call_tool(
                name='export_base64_advanced',
                arguments={'doc_id': doc_id, 'fmt': 'pdf'},
            )
            assert hasattr(pdf_default_export, 'data')
            assert isinstance(pdf_default_export.data, dict)
            assert set(pdf_default_export.data) >= {'base64', 'mime', 'ext'}
            assert pdf_default_export.data['ext'] == 'pdf'
            assert pdf_default_export.data['mime'] == 'application/pdf'

            decoded_default_pdf_bytes = base64.b64decode(pdf_default_export.data['base64'])
            assert isinstance(decoded_default_pdf_bytes, (bytes, bytearray))
            assert len(decoded_default_pdf_bytes) > 0

            pdf_shaped_export = await client.call_tool(
                name='export_base64_advanced',
                arguments={
                    'doc_id': doc_id,
                    'fmt': 'pdf',
                    'options': {'enable_text_shaping': True},
                },
            )
            assert hasattr(pdf_shaped_export, 'data')
            assert isinstance(pdf_shaped_export.data, dict)
            assert set(pdf_shaped_export.data) >= {'base64', 'mime', 'ext'}
            assert pdf_shaped_export.data['ext'] == 'pdf'
            assert pdf_shaped_export.data['mime'] == 'application/pdf'

            decoded_shaped_pdf_bytes = base64.b64decode(pdf_shaped_export.data['base64'])
            assert isinstance(decoded_shaped_pdf_bytes, (bytes, bytearray))
            assert len(decoded_shaped_pdf_bytes) > 0

            await asyncio.to_thread(Path(result_file_path).write_bytes, decoded_shaped_pdf_bytes)

    _run_and_assert_file(result_file_path, run_client)


def test_export_base64_advanced_xps_compression(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph',
                arguments={'doc_id': doc_id, 'text': 'XPS advanced export payload'},
            )

            exported = await client.call_tool(
                name='export_base64_advanced',
                arguments={
                    'doc_id': doc_id,
                    'fmt': 'xps',
                    'options': {'compression_level': 'super_fast'},
                },
            )
            assert hasattr(exported, 'data')
            assert isinstance(exported.data, dict)
            assert set(exported.data) >= {'base64', 'mime', 'ext'}
            assert exported.data['ext'] == 'xps'
            assert exported.data['mime'] == 'application/vnd.ms-xpsdocument'

            decoded_xps_bytes = base64.b64decode(exported.data['base64'])
            assert isinstance(decoded_xps_bytes, (bytes, bytearray))
            assert len(decoded_xps_bytes) > 0

            await asyncio.to_thread(Path(result_file_path).write_bytes, decoded_xps_bytes)

    _run_and_assert_file(result_file_path, run_client)


def test_merge_documents_with_and_without_new_page(mcp_client_config, result_file_path):
    async def run_client():
        async with _client_session(mcp_client_config) as client:
            first_doc_id = await _create_document_and_get_id(client)
            second_doc_id = await _create_document_and_get_id(client)
            await client.call_tool(
                name='add_paragraph',
                arguments={'doc_id': first_doc_id, 'text': 'Default merge first source'},
            )
            await client.call_tool(
                name='add_paragraph',
                arguments={'doc_id': second_doc_id, 'text': 'Default merge second source'},
            )

            second_doc_path = _runtime_docx_path(second_doc_id)
            second_doc = aw.Document(str(second_doc_path))
            second_doc.first_section.page_setup.section_start = aw.SectionStart.ODD_PAGE
            second_doc.save(str(second_doc_path))

            second_doc_reloaded = aw.Document(str(second_doc_path))
            assert (
                second_doc_reloaded.first_section.page_setup.section_start
                == aw.SectionStart.ODD_PAGE
            )

            merged_default = await client.call_tool(
                name='merge_documents',
                arguments={'source_doc_ids': [first_doc_id, second_doc_id]},
            )
            assert hasattr(merged_default, 'data')
            assert isinstance(merged_default.data, dict)
            assert 'docId' in merged_default.data
            merged_default_id = merged_default.data['docId']
            assert isinstance(merged_default_id, str)
            assert merged_default_id

            merged_default_text = await client.call_tool(
                name='get_text', arguments={'doc_id': merged_default_id}
            )
            assert hasattr(merged_default_text, 'data')
            assert isinstance(merged_default_text.data, dict)
            assert 'text' in merged_default_text.data
            assert 'Default merge first source' in merged_default_text.data['text']
            assert 'Default merge second source' in merged_default_text.data['text']

            merged_default_info = await client.call_tool(
                name='get_info', arguments={'doc_id': merged_default_id}
            )
            assert hasattr(merged_default_info, 'data')
            assert isinstance(merged_default_info.data, dict)
            assert 'pages' in merged_default_info.data

            merged_preserve_source_section_start = await client.call_tool(
                name='merge_documents',
                arguments={
                    'source_doc_ids': [first_doc_id, second_doc_id],
                    'append_document_with_new_page': False,
                },
            )
            assert hasattr(merged_preserve_source_section_start, 'data')
            assert isinstance(merged_preserve_source_section_start.data, dict)
            assert 'docId' in merged_preserve_source_section_start.data
            merged_preserve_source_section_start_id = merged_preserve_source_section_start.data[
                'docId'
            ]
            assert isinstance(merged_preserve_source_section_start_id, str)
            assert merged_preserve_source_section_start_id

            merged_preserve_source_section_start_text = await client.call_tool(
                name='get_text', arguments={'doc_id': merged_preserve_source_section_start_id}
            )
            assert hasattr(merged_preserve_source_section_start_text, 'data')
            assert isinstance(merged_preserve_source_section_start_text.data, dict)
            assert 'text' in merged_preserve_source_section_start_text.data
            assert (
                'Default merge first source'
                in merged_preserve_source_section_start_text.data['text']
            )
            assert (
                'Default merge second source'
                in merged_preserve_source_section_start_text.data['text']
            )

            merged_preserve_source_section_start_info = await client.call_tool(
                name='get_info', arguments={'doc_id': merged_preserve_source_section_start_id}
            )
            assert hasattr(merged_preserve_source_section_start_info, 'data')
            assert isinstance(merged_preserve_source_section_start_info.data, dict)
            assert 'pages' in merged_preserve_source_section_start_info.data

            merged_default_doc = aw.Document(str(_runtime_docx_path(merged_default_id)))
            merged_preserve_source_section_start_doc = aw.Document(
                str(_runtime_docx_path(merged_preserve_source_section_start_id))
            )
            assert merged_default_doc.sections.count >= 2
            assert merged_preserve_source_section_start_doc.sections.count >= 2
            assert (
                merged_default_doc.sections[1].page_setup.section_start == aw.SectionStart.NEW_PAGE
            )
            assert (
                merged_preserve_source_section_start_doc.sections[1].page_setup.section_start
                == aw.SectionStart.ODD_PAGE
            )

            await _save_exported_to_file(
                result_file_path, client, merged_preserve_source_section_start_id
            )

    _run_and_assert_file(result_file_path, run_client)
