import base64
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import aspose.words as aw
import pytest
from aspose.pydrawing import Color

import mcp_server as srv
from core import charts
from core.utils.docs_util import ensure_path

_CATEGORIES = ['North', 'South', 'West']
_SERIES = [
    {'name': 'Sales', 'values': [10, 0, -3.5]},
    {'name': 'Costs', 'values': [2, 4, 6]},
]
_NS = {
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}


def _new_chart(**kwargs):
    doc_id = srv.tool_create_document('chart.docx')['docId']
    result = srv.tool_add_chart_end(doc_id, 'column', _CATEGORIES, _SERIES, **kwargs)
    assert result == {'chartIndex': 0}
    return doc_id


def _load_chart(doc_id, index=0):
    doc = aw.Document(str(ensure_path(doc_id)))
    return charts._select_chart(doc, index)


def _data(chart):
    return [
        (
            series.name,
            [series.x_values[i].string_value for i in range(series.x_values.count)],
            [series.y_values[i].double_value for i in range(series.y_values.count)],
        )
        for series in chart.series
    ]


def _plot_xml(doc_id, child_path=''):
    with ZipFile(ensure_path(doc_id)) as archive:
        root = ET.fromstring(archive.read('word/charts/chart1.xml'))  # noqa: S314 - generated DOCX
    path = 'c:chart/c:plotArea/c:spPr'
    if child_path:
        path += '/' + child_path
    result = root.find(path, _NS)
    assert result is not None
    return result


@pytest.mark.parametrize('chart_type', ['column', ' COLUMN ', 'Column'])
def test_create_list_and_reopen_chart_data(chart_type):
    doc_id = srv.tool_create_document('chart-data.docx')['docId']
    srv.tool_insert_text_end(doc_id, 'Before chart')
    assert srv.tool_list_charts(doc_id) == {'charts': []}
    assert srv.tool_add_chart_end(doc_id, chart_type, _CATEGORIES, _SERIES) == {'chartIndex': 0}
    doc = aw.Document(str(ensure_path(doc_id)))
    shape = charts._chart_shapes(doc)[0]
    assert shape.width == 432
    assert shape.height == 252
    assert shape.previous_sibling.get_text() == 'Before chart'
    assert _data(shape.chart) == [
        ('Sales', _CATEGORIES, [10, 0, -3.5]),
        ('Costs', _CATEGORIES, [2, 4, 6]),
    ]
    before = ensure_path(doc_id).read_bytes()
    assert srv.tool_list_charts(doc_id) == {
        'charts': [
            {
                'chartIndex': 0,
                'series': [
                    {'seriesIndex': 0, 'name': 'Sales', 'pointCount': 3},
                    {'seriesIndex': 1, 'name': 'Costs', 'pointCount': 3},
                ],
            }
        ]
    }
    assert ensure_path(doc_id).read_bytes() == before


def test_plot_area_partial_updates_preserve_data_and_other_formatting():
    doc_id = _new_chart(width_points=300, height_points=180)
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    shape = charts._chart_shapes(doc)[0]
    assert (shape.width, shape.height) == (300, 180)
    chart = shape.chart
    chart.title.text = 'Preserved title'
    chart.legend.position = aw.drawing.charts.LegendPosition.BOTTOM
    chart.format.fill.solid(Color.from_argb(255, 1, 2, 3))
    chart.series[0].format.fill.solid(Color.from_argb(255, 4, 5, 6))
    chart.plot_area.format.stroke.visible = False
    doc.save(str(path), aw.SaveFormat.DOCX)
    before = path.read_bytes()
    assert srv.tool_format_chart_plot_area(doc_id, 0) == {}
    assert path.read_bytes() == before

    assert srv.tool_format_chart_plot_area(doc_id, 0, fill_color=' #aBcDeF ') == {}
    fmt = _load_chart(doc_id).plot_area.format
    assert fmt.fill.fill_type == aw.drawing.FillType.SOLID
    assert fmt.fill.fore_color.to_argb() == Color.from_argb(255, 171, 205, 239).to_argb()
    assert fmt.stroke.visible is False
    assert _plot_xml(doc_id, 'a:solidFill/a:srgbClr').attrib['val'] == 'ABCDEF'

    assert srv.tool_format_chart_plot_area(doc_id, 0, line_color='123456') == {}
    fmt = _load_chart(doc_id).plot_area.format
    assert fmt.stroke.visible is True
    assert fmt.stroke.fore_color.to_argb() == Color.from_argb(255, 18, 52, 86).to_argb()
    assert fmt.fill.fore_color.to_argb() == Color.from_argb(255, 171, 205, 239).to_argb()
    srv.tool_format_chart_plot_area(doc_id, 0, line_width_points=1.75)
    fmt = _load_chart(doc_id).plot_area.format
    assert fmt.stroke.weight == 1.75
    assert fmt.stroke.fore_color.to_argb() == Color.from_argb(255, 18, 52, 86).to_argb()

    srv.tool_format_chart_plot_area(doc_id, 0, gradient={'color': '336699'})
    fmt = _load_chart(doc_id).plot_area.format
    assert fmt.fill.fill_type == aw.drawing.FillType.GRADIENT
    assert fmt.fill.gradient_style == aw.drawing.GradientStyle.HORIZONTAL
    assert fmt.fill.gradient_variant == aw.drawing.GradientVariant.VARIANT1
    assert fmt.stroke.weight == 1.75
    assert fmt.stroke.fore_color.to_argb() == Color.from_argb(255, 18, 52, 86).to_argb()
    gradient_xml = ET.tostring(_plot_xml(doc_id, 'a:gradFill'))
    srv.tool_format_chart_plot_area(doc_id, 0, line_color='654321', line_width_points=2)
    assert ET.tostring(_plot_xml(doc_id, 'a:gradFill')) == gradient_xml
    line_xml = _plot_xml(doc_id, 'a:ln')
    assert line_xml.attrib['w'] == '25400'
    assert _plot_xml(doc_id, 'a:ln/a:solidFill/a:srgbClr').attrib['val'] == '654321'

    srv.tool_format_chart_plot_area(
        doc_id, 0, fill_color='00FF00', line_color='000000', line_width_points=0.5
    )
    chart = _load_chart(doc_id)
    assert chart.plot_area.format.fill.fill_type == aw.drawing.FillType.SOLID
    assert chart.plot_area.format.fill.fore_color.to_argb() == Color.lime.to_argb()
    assert chart.plot_area.format.stroke.weight == 0.5
    assert chart.title.text == 'Preserved title'
    assert chart.legend.position == aw.drawing.charts.LegendPosition.BOTTOM
    assert chart.format.fill.fore_color.to_argb() == Color.from_argb(255, 1, 2, 3).to_argb()
    assert (
        chart.series[0].format.fill.fore_color.to_argb() == Color.from_argb(255, 4, 5, 6).to_argb()
    )
    assert _data(chart) == [
        ('Sales', _CATEGORIES, [10, 0, -3.5]),
        ('Costs', _CATEGORIES, [2, 4, 6]),
    ]


@pytest.mark.parametrize(
    ('style', 'expected_style'),
    [
        ('horizontal', aw.drawing.GradientStyle.HORIZONTAL),
        (' VERTICAL ', aw.drawing.GradientStyle.VERTICAL),
        ('Diagonal-Up', aw.drawing.GradientStyle.DIAGONAL_UP),
        ('diagonal down', aw.drawing.GradientStyle.DIAGONAL_DOWN),
        ('from_corner', aw.drawing.GradientStyle.FROM_CORNER),
        ('From Center', aw.drawing.GradientStyle.FROM_CENTER),
    ],
)
@pytest.mark.parametrize(
    ('variant', 'expected_variant'),
    [
        ('variant1', aw.drawing.GradientVariant.VARIANT1),
        (' VARIANT_2 ', aw.drawing.GradientVariant.VARIANT2),
        ('Variant-3', aw.drawing.GradientVariant.VARIANT3),
        ('variant 4', aw.drawing.GradientVariant.VARIANT4),
    ],
)
def test_gradient_choices_round_trip(style, expected_style, variant, expected_variant):
    doc_id = _new_chart()
    srv.tool_format_chart_plot_area(
        doc_id,
        0,
        gradient={'color': '#336699', 'style': style, 'variant': variant, 'degree': 0.5},
        line_color='112233',
        line_width_points=1.25,
    )
    fmt = _load_chart(doc_id).plot_area.format
    assert fmt.fill.fill_type == aw.drawing.FillType.GRADIENT
    assert fmt.fill.fore_color.to_argb() == Color.from_argb(255, 51, 102, 153).to_argb()
    assert fmt.fill.gradient_style == expected_style
    # Aspose canonicalizes equivalent center gradients on DOCX round-trip.
    if expected_style == aw.drawing.GradientStyle.FROM_CENTER:
        if expected_variant == aw.drawing.GradientVariant.VARIANT2:
            expected_variant = aw.drawing.GradientVariant.VARIANT1
        elif expected_variant == aw.drawing.GradientVariant.VARIANT4:
            expected_variant = aw.drawing.GradientVariant.VARIANT3
    assert fmt.fill.gradient_variant == expected_variant
    assert fmt.stroke.weight == 1.25
    assert fmt.stroke.visible is True
    assert len(_plot_xml(doc_id, 'a:gradFill/a:gsLst')) >= 2


def test_gradient_degree_bounds_and_default():
    doc_id = _new_chart()
    gradients = []
    for options in ({}, {'degree': 0}, {'degree': 1}, {'degree': 0.5}):
        srv.tool_format_chart_plot_area(doc_id, 0, gradient={'color': '336699', **options})
        gradients.append(ET.tostring(_plot_xml(doc_id, 'a:gradFill')))
    assert gradients[0] == gradients[3]
    assert len(set(gradients[:3])) == 3


def test_chart_indices_filter_pictures_and_include_nested_shapes():
    doc_id = srv.tool_create_document('nested-charts.docx')['docId']
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    # A picture and a chart inside a text box precede the chart appended by the tool.
    png = base64.b64decode(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8AABQMB'
        'gYJ4V7wAAAAASUVORK5CYII='
    )
    builder.insert_image(BytesIO(png))
    box = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 150, 120)
    builder.move_to(box.first_paragraph)
    nested = builder.insert_chart(aw.drawing.charts.ChartType.COLUMN, 100, 80).chart
    nested.series.clear()
    nested.series.add('Nested', ['X'], [7.0])
    builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
    header = builder.insert_chart(aw.drawing.charts.ChartType.COLUMN, 100, 80).chart
    header.series.clear()
    header.series.add('Header', ['X'], [8.0])
    doc.save(str(path), aw.SaveFormat.DOCX)
    # Header/body traversal is determined by the document tree, never shape counts.
    result = srv.tool_add_chart_end(doc_id, 'column', _CATEGORIES, _SERIES)
    listed = srv.tool_list_charts(doc_id)['charts']
    assert len(listed) == 3
    assert [item['chartIndex'] for item in listed] == [0, 1, 2]
    names = [item['series'][0]['name'] for item in listed]
    assert set(names) == {'Nested', 'Header', 'Sales'}
    assert result['chartIndex'] == names.index('Sales')
    for index, name in enumerate(names):
        srv.tool_format_chart_plot_area(doc_id, index, fill_color=f'{index + 1:06X}')
        chart = _load_chart(doc_id, index)
        assert chart.series[0].name == name
        assert (
            chart.plot_area.format.fill.fore_color.to_argb()
            == Color.from_argb(255, 0, 0, index + 1).to_argb()
        )


@pytest.mark.parametrize(
    ('override', 'message'),
    [
        ({'chart_type': ''}, 'chart_type must be column or waterfall'),
        ({'chart_type': '  '}, 'chart_type must be column or waterfall'),
        ({'chart_type': 'pie'}, 'chart_type must be column or waterfall'),
        ({'chart_type': None}, 'chart_type must be column or waterfall'),
        ({'categories': []}, 'categories must be a nonempty list'),
        ({'categories': [' ']}, 'categories must be a nonempty list'),
        ({'categories': [1]}, 'categories must be a nonempty list'),
        ({'categories': 'North'}, 'categories must be a nonempty list'),
        ({'series': []}, 'series must be a nonempty list'),
        ({'series': [{}]}, 'must contain only name and values'),
        ({'series': [{'name': '', 'values': [1, 2, 3]}]}, 'name must be a nonblank string'),
        ({'series': [{'name': ' ', 'values': [1, 2, 3]}]}, 'name must be a nonblank string'),
        ({'series': [{'name': 'X', 'values': [1]}]}, 'values must match categories length'),
        ({'series': [{'name': 'X', 'values': [1, 2, float('nan')]}]}, 'must be a finite number'),
        ({'series': [{'name': 'X', 'values': [1, 2, float('inf')]}]}, 'must be a finite number'),
        ({'series': [{'name': 'X', 'values': [1, 2, True]}]}, 'must be a finite number'),
        ({'series': [{'name': 'X', 'values': [1, 2, '3']}]}, 'must be a finite number'),
        (
            {'series': [{'name': 'X', 'values': [1, 2, 3], 'is_subtotal': [False] * 3}]},
            'is_subtotal is only supported for waterfall charts',
        ),
        ({'width_points': 0}, 'width_points must be greater than zero'),
        ({'width_points': -1}, 'width_points must be greater than zero'),
        ({'width_points': float('inf')}, 'width_points must be a finite number'),
        ({'width_points': True}, 'width_points must be a finite number'),
        ({'height_points': 0}, 'height_points must be greater than zero'),
        ({'height_points': float('nan')}, 'height_points must be a finite number'),
    ],
)
def test_invalid_creation_preserves_document(override, message):
    doc_id = _new_chart()
    path = ensure_path(doc_id)
    before = path.read_bytes()
    arguments = {'chart_type': 'column', 'categories': _CATEGORIES, 'series': _SERIES, **override}
    with pytest.raises(ValueError, match=message):
        srv.tool_add_chart_end(doc_id, **arguments)
    assert path.read_bytes() == before


@pytest.mark.parametrize(
    ('options', 'message'),
    [
        ({'fill_color': '112233', 'gradient': {}}, 'cannot be specified together'),
        ({'fill_color': ''}, 'fill_color must be a six-digit RGB'),
        ({'fill_color': '   '}, 'fill_color must be a six-digit RGB'),
        ({'fill_color': '#123'}, 'fill_color must be a six-digit RGB'),
        ({'fill_color': 'GG0000'}, 'fill_color must be a six-digit RGB'),
        ({'fill_color': '##112233'}, 'fill_color must be a six-digit RGB'),
        ({'fill_color': 'FF112233'}, 'fill_color must be a six-digit RGB'),
        ({'line_color': 'red'}, 'line_color must be a six-digit RGB'),
        ({'line_color': ''}, 'line_color must be a six-digit RGB'),
        ({'gradient': ''}, 'gradient must be an object with color'),
        ({'gradient': {}}, 'gradient must be an object with color'),
        ({'gradient': {'color': ''}}, 'gradient.color must be a six-digit RGB'),
        ({'gradient': {'color': '112233', 'extra': 1}}, 'gradient supports only'),
        ({'gradient': {'color': '112233', 'style': 'radial'}}, 'gradient.style must be one of'),
        ({'gradient': {'color': '112233', 'style': ''}}, 'gradient.style must be one of'),
        ({'gradient': {'color': '112233', 'style': 'none'}}, 'gradient.style must be one of'),
        (
            {'gradient': {'color': '112233', 'variant': 'variant5'}},
            'gradient.variant must be one of',
        ),
        ({'gradient': {'color': '112233', 'variant': ''}}, 'gradient.variant must be one of'),
        ({'gradient': {'color': '112233', 'degree': -0.1}}, 'degree must be between 0 and 1'),
        ({'gradient': {'color': '112233', 'degree': 1.1}}, 'degree must be between 0 and 1'),
        ({'gradient': {'color': '112233', 'degree': float('nan')}}, 'degree must be a finite'),
        ({'gradient': {'color': '112233', 'degree': True}}, 'degree must be a finite'),
        ({'line_width_points': 0}, 'line_width_points must be greater than zero'),
        ({'line_width_points': -1}, 'line_width_points must be greater than zero'),
        ({'line_width_points': float('inf')}, 'line_width_points must be a finite number'),
        ({'line_width_points': float('nan')}, 'line_width_points must be a finite number'),
        ({'line_width_points': ' '}, 'line_width_points must be a finite number'),
        ({'line_width_points': True}, 'line_width_points must be a finite number'),
    ],
)
def test_invalid_formatting_preserves_document(options, message):
    doc_id = _new_chart()
    path = ensure_path(doc_id)
    before = path.read_bytes()
    with pytest.raises(ValueError, match=message):
        srv.tool_format_chart_plot_area(doc_id, 0, **options)
    assert path.read_bytes() == before


@pytest.mark.parametrize('index', [-1, 1, 100, True, 0.5, '0'])
def test_invalid_chart_index_preserves_document(index):
    doc_id = _new_chart()
    path = ensure_path(doc_id)
    before = path.read_bytes()
    with pytest.raises((ValueError, IndexError), match='chart_index'):
        srv.tool_format_chart_plot_area(doc_id, index, fill_color='112233')
    assert path.read_bytes() == before


def test_missing_documents_and_absent_charts():
    with pytest.raises(FileNotFoundError, match='Document missing-chart-document not found'):
        srv.tool_add_chart_end('missing-chart-document', 'column', _CATEGORIES, _SERIES)
    with pytest.raises(FileNotFoundError, match='Document missing-chart-document not found'):
        srv.tool_list_charts('missing-chart-document')
    with pytest.raises(FileNotFoundError, match='Document missing-chart-document not found'):
        srv.tool_format_chart_plot_area('missing-chart-document', 0, fill_color='112233')
    doc_id = srv.tool_create_document()['docId']
    before = ensure_path(doc_id).read_bytes()
    with pytest.raises(IndexError, match='document has 0 charts'):
        srv.tool_format_chart_plot_area(doc_id, 0, fill_color='112233')
    assert ensure_path(doc_id).read_bytes() == before


def test_chart_wrappers_forward_all_parameters(monkeypatch):
    calls = []

    def add(*args, **kwargs):
        calls.append((args, kwargs))
        return 4

    def listing(doc_id):
        calls.append(doc_id)
        return [{'chartIndex': 4, 'series': []}]

    def formatting(*args, **kwargs):
        calls.append((args, kwargs))

    monkeypatch.setattr(charts, 'add_chart_end', add)
    monkeypatch.setattr(charts, 'list_charts', listing)
    monkeypatch.setattr(charts, 'format_chart_plot_area', formatting)
    assert srv.tool_add_chart_end('doc', 'column', _CATEGORIES, _SERIES, 200, 150) == {
        'chartIndex': 4
    }
    assert calls.pop() == (
        ('doc', 'column', _CATEGORIES, _SERIES),
        {'width_points': 200, 'height_points': 150},
    )
    assert srv.tool_list_charts('doc') == {'charts': [{'chartIndex': 4, 'series': []}]}
    assert calls.pop() == 'doc'
    gradient = {'color': '112233'}
    assert srv.tool_format_chart_plot_area('doc', 4, None, gradient, '445566', 2) == {}
    assert calls.pop() == (
        ('doc', 4),
        {'fill_color': None, 'gradient': gradient, 'line_color': '445566', 'line_width_points': 2},
    )


_WATERFALL_CATEGORIES = ['Opening', 'Sales', 'Costs', 'Intermediate', 'Income', 'Expense', 'Final']
_WATERFALL_VALUES = [100, 20, -10, 110, 15, -5, 120]
_WATERFALL_TOTALS = [True, False, False, True, False, False, True]


def _waterfall_series(flags=None):
    return {
        'name': 'Balance',
        'values': _WATERFALL_VALUES,
        'is_subtotal': _WATERFALL_TOTALS if flags is None else flags,
    }


def _expected_totals(flags, chart_index=0, series_index=0):
    return {
        'chartIndex': chart_index,
        'seriesIndex': series_index,
        'points': [
            {
                'pointIndex': index,
                'dataPointIndex': index,
                'isTotal': flag,
                'dataPointIsTotal': flag,
            }
            for index, flag in enumerate(flags)
        ],
    }


@pytest.mark.parametrize('chart_type', ['waterfall', ' WATERFALL ', 'Waterfall'])
def test_waterfall_create_inspect_and_reopen(chart_type):
    doc_id = srv.tool_create_document('waterfall.docx')['docId']
    assert srv.tool_add_chart_end(
        doc_id, chart_type, _WATERFALL_CATEGORIES, [_waterfall_series()]
    ) == {'chartIndex': 0}
    chart = _load_chart(doc_id)
    assert chart.series.count == 1
    series = chart.series[0]
    assert series.series_type == aw.drawing.charts.ChartSeriesType.WATERFALL
    assert _data(chart) == [('Balance', _WATERFALL_CATEGORIES, _WATERFALL_VALUES)]
    assert series.y_values.count == len(_WATERFALL_TOTALS)
    for index, flag in enumerate(_WATERFALL_TOTALS):
        assert series.is_total(index) is flag
        assert series.data_points[index].is_total is flag
        assert series.data_points[index].index == index
        assert series.data_points.has_default_format(index)
    path = ensure_path(doc_id)
    before = path.read_bytes()
    expected = _expected_totals(_WATERFALL_TOTALS)
    assert srv.tool_get_waterfall_totals(doc_id, 0) == expected
    for index, point in enumerate(expected['points']):
        assert srv.tool_get_waterfall_totals(doc_id, 0, point_index=index) == {
            **expected,
            'points': [point],
        }
    # Queries preserve the entire package, including all data and point-format XML.
    assert path.read_bytes() == before
    reopened = _load_chart(doc_id).series[0]
    assert all(reopened.data_points.has_default_format(i) for i in range(reopened.y_values.count))


def test_waterfall_multiple_chart_and_series_selection():
    doc_id = _new_chart()
    flag_sets = [_WATERFALL_TOTALS, [False] * 7, [True] * 7]
    for chart_index in (1, 2):
        flags = flag_sets if chart_index == 1 else list(reversed(flag_sets))
        assert srv.tool_add_chart_end(
            doc_id, 'waterfall', _WATERFALL_CATEGORIES, [_waterfall_series(f) for f in flags]
        ) == {'chartIndex': chart_index}
        before = ensure_path(doc_id).read_bytes()
        for series_index, series_flags in enumerate(flags):
            expected = _expected_totals(series_flags, chart_index, series_index)
            assert srv.tool_get_waterfall_totals(doc_id, chart_index, series_index) == expected
            assert srv.tool_get_waterfall_totals(doc_id, chart_index, series_index, 3) == {
                **expected,
                'points': [expected['points'][3]],
            }
        assert ensure_path(doc_id).read_bytes() == before


@pytest.fixture
def empty_waterfall_document():
    doc_id = srv.tool_create_document('empty-waterfall.docx')['docId']
    doc = aw.Document(str(ensure_path(doc_id)))
    chart = (
        aw.DocumentBuilder(doc).insert_chart(aw.drawing.charts.ChartType.WATERFALL, 432, 252).chart
    )
    chart.series.clear()
    series = chart.series.add('Empty', ['Seed'], [0.0], [False])
    series.clear_values()
    doc.save(str(ensure_path(doc_id)), aw.SaveFormat.DOCX)
    return doc_id


def test_empty_waterfall_fixture(empty_waterfall_document):
    doc_id = empty_waterfall_document
    assert _load_chart(doc_id).series[0].y_values.count == 0
    before = ensure_path(doc_id).read_bytes()
    assert srv.tool_get_waterfall_totals(doc_id, 0) == _expected_totals([])
    with pytest.raises(IndexError, match='point_index out of range: 0; series has 0 points'):
        srv.tool_get_waterfall_totals(doc_id, 0, point_index=0)
    assert ensure_path(doc_id).read_bytes() == before


@pytest.mark.parametrize(
    ('flags', 'message'),
    [
        ([], 'is_subtotal must match categories length'),
        ([True], 'is_subtotal must match categories length'),
        ([True] * 8, 'is_subtotal must match categories length'),
        (None, 'is_subtotal must match categories length'),
        ('', 'is_subtotal must match categories length'),
        ('       ', 'is_subtotal must match categories length'),
        ([True, False, False, True, False, False, 1], 'is_subtotal must contain only booleans'),
        ([True, False, False, True, False, False, 0], 'is_subtotal must contain only booleans'),
        (
            [True, False, False, True, False, False, 'true'],
            'is_subtotal must contain only booleans',
        ),
        ([True, False, False, True, False, False, None], 'is_subtotal must contain only booleans'),
    ],
)
def test_invalid_waterfall_subtotal_arrays_preserve_document(flags, message):
    doc_id = _new_chart()
    before = ensure_path(doc_id).read_bytes()
    series = {**_waterfall_series(), 'is_subtotal': flags}
    with pytest.raises(ValueError, match=message):
        srv.tool_add_chart_end(doc_id, 'waterfall', _WATERFALL_CATEGORIES, [series])
    assert ensure_path(doc_id).read_bytes() == before


@pytest.mark.parametrize(
    ('series', 'message'),
    [
        (
            [{'name': 'X', 'values': _WATERFALL_VALUES}],
            'must contain only name, values and is_subtotal',
        ),
        (
            [{**_waterfall_series(), 'extra': True}],
            'must contain only name, values and is_subtotal',
        ),
        ([{**_waterfall_series(), 'values': [1]}], 'values must match categories length'),
        ([{**_waterfall_series(), 'name': ' '}], 'name must be a nonblank string'),
        ([{**_waterfall_series(), 'values': [True] * 7}], 'values must be a finite number'),
        ([_waterfall_series(), {**_waterfall_series(), 'is_subtotal': [1] * 7}], 'only booleans'),
    ],
)
def test_invalid_waterfall_series_preserve_document(series, message):
    doc_id = _new_chart()
    before = ensure_path(doc_id).read_bytes()
    with pytest.raises(ValueError, match=message):
        srv.tool_add_chart_end(doc_id, 'waterfall', _WATERFALL_CATEGORIES, series)
    assert ensure_path(doc_id).read_bytes() == before


@pytest.mark.parametrize(
    ('field', 'index'),
    [
        (field, index)
        for field in ('chart_index', 'series_index', 'point_index')
        for index in [-1, 100, True, 0.5, '0', '', ' ', 7 if field == 'point_index' else 1]
    ]
    + [('chart_index', None), ('series_index', None)],
)
def test_invalid_waterfall_indices_preserve_document(field, index):
    doc_id = srv.tool_create_document('waterfall-indices.docx')['docId']
    srv.tool_add_chart_end(doc_id, 'waterfall', _WATERFALL_CATEGORIES, [_waterfall_series()])
    before = ensure_path(doc_id).read_bytes()
    kwargs = {'chart_index': 0, field: index}
    with pytest.raises((ValueError, IndexError), match=field):
        srv.tool_get_waterfall_totals(doc_id, **kwargs)
    assert ensure_path(doc_id).read_bytes() == before


def test_waterfall_inspection_rejects_missing_documents_charts_and_non_waterfall_series():
    with pytest.raises(FileNotFoundError, match='Document missing-waterfall not found'):
        srv.tool_get_waterfall_totals('missing-waterfall', 0)
    doc_id = srv.tool_create_document()['docId']
    before = ensure_path(doc_id).read_bytes()
    with pytest.raises(IndexError, match='document has 0 charts'):
        srv.tool_get_waterfall_totals(doc_id, 0)
    assert ensure_path(doc_id).read_bytes() == before
    srv.tool_add_chart_end(doc_id, 'column', _CATEGORIES, _SERIES)
    before = ensure_path(doc_id).read_bytes()
    with pytest.raises(ValueError, match='Total inspection is only supported for waterfall series'):
        srv.tool_get_waterfall_totals(doc_id, 0)
    assert ensure_path(doc_id).read_bytes() == before


def test_waterfall_wrappers_forward_all_parameters(monkeypatch):
    calls = []
    points = _expected_totals(_WATERFALL_TOTALS)['points']

    def inspect(*args, **kwargs):
        calls.append((args, kwargs))
        return points

    def add(*args, **kwargs):
        calls.append((args, kwargs))
        return 4

    monkeypatch.setattr(charts, 'get_waterfall_totals', inspect)
    monkeypatch.setattr(charts, 'add_chart_end', add)
    series = [_waterfall_series()]
    assert srv.tool_add_chart_end('doc', 'waterfall', _WATERFALL_CATEGORIES, series) == {
        'chartIndex': 4
    }
    assert calls.pop() == (
        ('doc', 'waterfall', _WATERFALL_CATEGORIES, series),
        {'width_points': 432, 'height_points': 252},
    )
    assert srv.tool_get_waterfall_totals('doc', 4) == {
        'chartIndex': 4,
        'seriesIndex': 0,
        'points': points,
    }
    assert calls.pop() == (('doc', 4), {'series_index': 0, 'point_index': None})
    assert srv.tool_get_waterfall_totals('doc', 4, 2, 3) == {
        'chartIndex': 4,
        'seriesIndex': 2,
        'points': points,
    }
    assert calls.pop() == (('doc', 4), {'series_index': 2, 'point_index': 3})
