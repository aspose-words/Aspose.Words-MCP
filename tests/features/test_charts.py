import base64
import io
import tomllib
from pathlib import Path
from zipfile import ZipFile

import aspose.words as aw
import aspose.words.drawing as drawing
import aspose.words.drawing.charts as charts
import pytest
from aspose.pydrawing import Color

import mcp_server as srv
from core.utils.docs_util import ensure_path

CATEGORIES = ['Start', 'Increase', 'Intermediate', 'Decrease', 'Final']
VALUES = [10, 5, 15, -2, 13]
FLAGS = [True, False, True, False, True]


@pytest.fixture
def chart_doc():
    return srv.tool_create_document('charts.docx')['docId']


def _insert(doc_id, kind='column', **kwargs):
    series = [{'name': 'Flow', 'values': VALUES}]
    if kind.strip().lower() == 'waterfall':
        series[0]['is_subtotal'] = FLAGS
    return srv.tool_insert_chart(doc_id, kind, CATEGORIES, series, **kwargs)


def _shapes(doc_id):
    doc = aw.Document(str(ensure_path(doc_id)))
    return [
        node.as_shape()
        for node in doc.get_child_nodes(aw.NodeType.SHAPE, True)
        if node.as_shape().has_chart
    ]


def _series_data(chart):
    return [
        (
            series.name,
            [series.x_values[i].string_value for i in range(series.x_values.count)],
            [series.y_values[i].double_value for i in range(series.y_values.count)],
        )
        for series in chart.series
    ]


def _format_data(chart_format):
    return (
        chart_format.fill.fore_color.to_argb(),
        chart_format.stroke.fore_color.to_argb(),
        chart_format.stroke.weight,
    )


def test_chart_creation_discovery_and_document_order(chart_doc):
    assert srv.tool_list_charts(chart_doc) == {'charts': []}
    srv.tool_insert_text_end(chart_doc, 'Before charts')
    path = ensure_path(chart_doc)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    builder.insert_shape(drawing.ShapeType.RECTANGLE, 20, 20)
    doc.save(str(path), aw.SaveFormat.DOCX)

    assert _insert(chart_doc, ' COLUMN ') == {'chartIndex': 0}
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    builder.move_to_document_end()
    builder.insert_shape(drawing.ShapeType.ELLIPSE, 20, 20)
    doc.save(str(path), aw.SaveFormat.DOCX)
    second = [{'name': 'First', 'values': [1, 2]}, {'name': 'Second', 'values': [3, 4]}]
    assert srv.tool_insert_chart(chart_doc, 'column', ['A', 'B'], second, 300, 200) == {
        'chartIndex': 1
    }
    assert srv.tool_list_charts(chart_doc) == {
        'charts': [
            {
                'chartIndex': 0,
                'seriesCount': 1,
                'series': [{'seriesIndex': 0, 'name': 'Flow', 'seriesType': 'column'}],
            },
            {
                'chartIndex': 1,
                'seriesCount': 2,
                'series': [
                    {'seriesIndex': 0, 'name': 'First', 'seriesType': 'column'},
                    {'seriesIndex': 1, 'name': 'Second', 'seriesType': 'column'},
                ],
            },
        ]
    }
    shapes = _shapes(chart_doc)
    assert (shapes[0].width, shapes[0].height) == (432, 252)
    assert (shapes[1].width, shapes[1].height) == (300, 200)
    assert _series_data(shapes[0].chart) == [('Flow', CATEGORIES, VALUES)]
    assert _series_data(shapes[1].chart) == [
        ('First', ['A', 'B'], [1, 2]),
        ('Second', ['A', 'B'], [3, 4]),
    ]
    nodes = aw.Document(str(path)).get_child_nodes(aw.NodeType.ANY, True)
    text_position = next(
        i
        for i, node in enumerate(nodes)
        if node.node_type == aw.NodeType.RUN and 'Before charts' in node.as_run().text
    )
    chart_position = next(
        i
        for i, node in enumerate(nodes)
        if node.node_type == aw.NodeType.SHAPE and node.as_shape().has_chart
    )
    assert text_position < chart_position


def test_insert_chart_index_accounts_for_footer_charts(chart_doc):
    path = ensure_path(chart_doc)
    document = aw.Document(str(path))
    builder = aw.DocumentBuilder(document)
    builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
    footer = builder.insert_chart(charts.ChartType.COLUMN, 100, 100).chart
    footer.series.clear()
    footer.series.add('Footer', ['A'], [1.0])
    document.save(str(path), aw.SaveFormat.DOCX)

    inserted = _insert(chart_doc)
    listing = srv.tool_list_charts(chart_doc)['charts']
    assert inserted == {'chartIndex': 0}
    assert listing[inserted['chartIndex']]['series'][0]['name'] == 'Flow'
    assert listing[1]['series'][0]['name'] == 'Footer'


@pytest.mark.parametrize(
    'options, expected',
    [
        ({'fill_color': ' #aAbBcC '}, (0xFFAABBCC, 0xFF445566, 1.5)),
        ({'line_color': ' ABCdef '}, (0xFF112233, 0xFFABCDEF, 1.5)),
        ({'line_width': 2.75}, (0xFF112233, 0xFF445566, 2.75)),
        ({'line_color': '#aabbcc', 'line_width': 3}, (0xFF112233, 0xFFAABBCC, 3)),
        (
            {'fill_color': 'abcdef', 'line_color': '123456', 'line_width': 2},
            (0xFFABCDEF, 0xFF123456, 2),
        ),
    ],
)
def test_plot_formatting_persists_without_changing_chart_or_series(chart_doc, options, expected):
    _insert(chart_doc)
    path = ensure_path(chart_doc)
    doc = aw.Document(str(path))
    chart = doc.get_child_nodes(aw.NodeType.SHAPE, True)[0].as_shape().chart
    chart.plot_area.format.fill.solid(Color.from_argb(17, 34, 51))
    chart.plot_area.format.stroke.fore_color = Color.from_argb(68, 85, 102)
    chart.plot_area.format.stroke.weight = 1.5
    chart.format.fill.solid(Color.from_argb(12, 34, 56))
    chart.format.stroke.fore_color = Color.from_argb(78, 90, 12)
    chart.format.stroke.weight = 4
    doc.save(str(path), aw.SaveFormat.DOCX)
    before = _shapes(chart_doc)[0].chart
    before_data = _series_data(before)
    before_format = _format_data(before.format)

    assert srv.tool_format_chart_plot_area(chart_doc, 0, **options) == {}
    after = _shapes(chart_doc)[0].chart
    actual = _format_data(after.plot_area.format)
    assert tuple(value & 0xFFFFFFFF for value in actual[:2]) == expected[:2]
    assert actual[2] == expected[2]
    assert _format_data(after.format) == before_format
    assert _series_data(after) == before_data


def test_waterfall_totals_both_getters_and_read_only_inspection(chart_doc):
    assert _insert(chart_doc, ' WaTeRfAlL ') == {'chartIndex': 0}
    path = ensure_path(chart_doc)
    before = path.read_bytes()
    before_mtime = path.stat().st_mtime_ns
    chart = _shapes(chart_doc)[0].chart
    assert chart.series[0].series_type == charts.ChartSeriesType.WATERFALL
    assert _series_data(chart) == [('Flow', CATEGORIES, VALUES)]
    series = chart.series[0]
    assert series.x_values.count == len(FLAGS)
    assert series.data_points.has_default_format(1)
    assert series.data_points.has_default_format(3)
    assert [series.is_total(i) for i in range(series.x_values.count)] == FLAGS
    assert [series.data_points[i].is_total for i in range(series.x_values.count)] == FLAGS
    assert [series.data_points[i].index for i in range(series.x_values.count)] == list(range(5))
    assert srv.tool_get_waterfall_totals(chart_doc, 0) == {
        'chartIndex': 0,
        'seriesIndex': 0,
        'points': [{'index': i, 'isTotal': flag} for i, flag in enumerate(FLAGS)],
        'totalIndices': [0, 2, 4],
    }
    assert srv.tool_list_charts(chart_doc)['charts'][0]['series'][0]['seriesType'] == 'waterfall'
    assert path.read_bytes() == before
    assert path.stat().st_mtime_ns == before_mtime


def test_waterfall_default_flags_and_explicit_series_index(chart_doc):
    srv.tool_insert_chart(
        chart_doc,
        'waterfall',
        ['A', 'B'],
        [
            {'name': 'Plain', 'values': [1, -1]},
            {'name': 'Totals', 'values': [1, 1], 'is_subtotal': [False, True]},
        ],
    )
    before = ensure_path(chart_doc).read_bytes()
    assert srv.tool_get_waterfall_totals(chart_doc, 0) == {
        'chartIndex': 0,
        'seriesIndex': 0,
        'points': [{'index': 0, 'isTotal': False}, {'index': 1, 'isTotal': False}],
        'totalIndices': [],
    }
    result = srv.tool_get_waterfall_totals(chart_doc, 0, 1)
    assert result['seriesIndex'] == 1
    assert result['points'] == [{'index': 0, 'isTotal': False}, {'index': 1, 'isTotal': True}]
    assert result['totalIndices'] == [1]
    assert ensure_path(chart_doc).read_bytes() == before


@pytest.mark.parametrize('kind', ['column', 'waterfall'])
def test_chart_export_and_render_consumption(chart_doc, kind):
    _insert(chart_doc, kind)
    srv.tool_format_chart_plot_area(chart_doc, 0, fill_color='#ddeeff', line_width=1)
    exported = srv.tool_export_base64(chart_doc)
    raw = base64.b64decode(exported['base64'])
    with ZipFile(io.BytesIO(raw)) as package:
        assert any(
            name.startswith('word/charts/chart') and name.endswith('.xml')
            for name in package.namelist()
        )
    reopened = aw.Document(io.BytesIO(raw))
    assert reopened.get_child_nodes(aw.NodeType.SHAPE, True)[0].as_shape().has_chart
    image = srv.tool_render_page_base64(chart_doc, page_index=0, fmt='png', dpi=72)
    assert base64.b64decode(image['base64']).startswith(b'\x89PNG\r\n\x1a\n')


@pytest.mark.parametrize(
    'change, message',
    [
        ({'chart_type': ''}, 'chart_type must be column or waterfall'),
        ({'chart_type': '  '}, 'chart_type must be column or waterfall'),
        ({'chart_type': 'pie'}, 'chart_type must be column or waterfall'),
        *[
            ({'width': value}, 'width must be a positive finite number')
            for value in [0, -1, float('nan'), float('inf'), True, '12']
        ],
        *[
            ({'height': value}, 'height must be a positive finite number')
            for value in [0, -1, float('nan'), float('inf')]
        ],
        *[
            ({'categories': value}, 'categories must be a nonempty list of nonblank strings')
            for value in [[], [''], ['  '], [1]]
        ],
        ({'series': []}, 'series must be a nonempty list'),
        ({'series': ['bad']}, 'each series must contain name, values, and optional is_subtotal'),
        (
            {'series': [{'name': 'N', 'values': [1], 'extra': 1}]},
            'each series must contain name, values, and optional is_subtotal',
        ),
        *[
            ({'series': [{'name': name, 'values': [1]}]}, 'series name must be a nonblank string')
            for name in ['', ' ', None]
        ],
        *[
            (
                {'series': [{'name': 'N', 'values': values}]},
                'series values must be a list matching categories length',
            )
            for values in [[], None]
        ],
        *[
            (
                {'series': [{'name': 'N', 'values': [value]}]},
                'series values must contain only finite numbers',
            )
            for value in [float('nan'), float('inf'), -float('inf'), True, '1', None]
        ],
        (
            {'series': [{'name': 'N', 'values': [1], 'is_subtotal': [False]}]},
            'is_subtotal is supported only for waterfall charts',
        ),
        *[
            (
                {
                    'chart_type': 'waterfall',
                    'series': [{'name': 'N', 'values': [1], 'is_subtotal': flags}],
                },
                'is_subtotal must be a boolean list matching categories length',
            )
            for flags in [[], [1], ['true'], [False, True], None]
        ],
    ],
)
def test_invalid_chart_input_leaves_document_unchanged(chart_doc, change, message):
    path = ensure_path(chart_doc)
    before = path.read_bytes()
    arguments = {
        'doc_id': chart_doc,
        'chart_type': 'column',
        'categories': ['A'],
        'series': [{'name': 'N', 'values': [1]}],
    }
    arguments.update(change)
    with pytest.raises(ValueError, match=message):
        srv.tool_insert_chart(**arguments)
    assert path.read_bytes() == before


@pytest.mark.parametrize(
    'options, message',
    [
        ({}, 'provide at least one of fill_color, line_color, or line_width'),
        *[
            ({key: value}, f'{key} must be a six-digit hexadecimal color')
            for key in ['fill_color', 'line_color']
            for value in ['', ' ', '#123', '12345678', 'gg0000', '##123456']
        ],
        *[
            ({'line_width': value}, 'line_width must be a positive finite number')
            for value in [0, -1, float('inf'), float('nan'), True, '1']
        ],
        (
            {'fill_color': 'aabbcc', 'line_color': 'invalid'},
            'line_color must be a six-digit hexadecimal color',
        ),
    ],
)
def test_invalid_formatting_leaves_document_unchanged(chart_doc, options, message):
    _insert(chart_doc)
    path = ensure_path(chart_doc)
    before = path.read_bytes()
    with pytest.raises(ValueError, match=message):
        srv.tool_format_chart_plot_area(chart_doc, 0, **options)
    assert path.read_bytes() == before


@pytest.mark.parametrize('index', [-1, 1, 100])
def test_invalid_chart_and_series_indices(chart_doc, index):
    _insert(chart_doc, 'waterfall')
    before = ensure_path(chart_doc).read_bytes()
    with pytest.raises(ValueError, match='chart_index out of range'):
        srv.tool_format_chart_plot_area(chart_doc, index, fill_color='123456')
    with pytest.raises(ValueError, match='chart_index out of range'):
        srv.tool_get_waterfall_totals(chart_doc, index)
    with pytest.raises(ValueError, match='series_index out of range'):
        srv.tool_get_waterfall_totals(chart_doc, 0, index)
    assert ensure_path(chart_doc).read_bytes() == before


def test_non_waterfall_and_empty_document_errors(chart_doc):
    with pytest.raises(ValueError, match='chart_index out of range'):
        srv.tool_get_waterfall_totals(chart_doc, 0)
    _insert(chart_doc)
    before = ensure_path(chart_doc).read_bytes()
    with pytest.raises(ValueError, match='requires a waterfall series'):
        srv.tool_get_waterfall_totals(chart_doc, 0)
    assert ensure_path(chart_doc).read_bytes() == before


@pytest.mark.parametrize(
    'call',
    [
        lambda: _insert('missing-chart-document'),
        lambda: srv.tool_list_charts('missing-chart-document'),
        lambda: srv.tool_format_chart_plot_area('missing-chart-document', 0, line_width=1),
        lambda: srv.tool_get_waterfall_totals('missing-chart-document', 0),
    ],
)
def test_missing_chart_document(call):
    with pytest.raises(FileNotFoundError, match='Document missing-chart-document not found'):
        call()


def test_release_alignment_metadata():
    root = Path(__file__).resolve().parents[2]
    project = tomllib.loads((root / 'pyproject.toml').read_text())['project']
    assert project['version'] == aw.BuildVersionInfo.version == '26.10.0'
    assert 'aspose-words>=26.10.0' in project['dependencies']
    assert 'aspose-words==26.10.0' in (root / 'requirements.txt').read_text()
    assert '`26.10.0`' in (root / 'README.md').read_text()
    assert '## [26.10.0]' in (root / 'CHANGELOG.md').read_text()
