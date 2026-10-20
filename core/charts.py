from __future__ import annotations

import math
import re

import aspose.words as aw
import aspose.words.drawing.charts as charts
from aspose.pydrawing import Color

from core.utils.docs_util import ensure_path

_CHART_TYPES = {'column': charts.ChartType.COLUMN, 'waterfall': charts.ChartType.WATERFALL}


def _positive_number(value: float, parameter: str) -> None:
    if (
        isinstance(value, bool)
        or not isinstance(value, (int, float))
        or not math.isfinite(value)
        or value <= 0
    ):
        raise ValueError(f'{parameter} must be a positive finite number')


def _chart_shapes(document: aw.Document) -> list[aw.drawing.Shape]:
    nodes = document.get_child_nodes(aw.NodeType.SHAPE, True)
    shapes = []
    for index in range(nodes.count):
        shape = nodes[index].as_shape()
        if shape.has_chart:
            shapes.append(shape)
    return shapes


def _validate_index(index: int, count: int, parameter: str) -> None:
    if isinstance(index, bool) or not isinstance(index, int) or index < 0 or index >= count:
        raise ValueError(f'{parameter} out of range: expected a zero-based index below {count}')


def _chart(document: aw.Document, chart_index: int) -> charts.Chart:
    shapes = _chart_shapes(document)
    _validate_index(chart_index, len(shapes), 'chart_index')
    return shapes[chart_index].chart


def insert_chart(
    doc_id: str,
    chart_type: str,
    categories: list[str],
    series: list[dict],
    width: float = 432,
    height: float = 252,
) -> int:
    if not isinstance(chart_type, str) or chart_type.strip().lower() not in _CHART_TYPES:
        raise ValueError('chart_type must be column or waterfall')
    kind = chart_type.strip().lower()
    _positive_number(width, 'width')
    _positive_number(height, 'height')
    if (
        not isinstance(categories, list)
        or not categories
        or any(not isinstance(item, str) or not item.strip() for item in categories)
    ):
        raise ValueError('categories must be a nonempty list of nonblank strings')
    if not isinstance(series, list) or not series:
        raise ValueError('series must be a nonempty list')
    for item in series:
        if not isinstance(item, dict) or set(item) - {'name', 'values', 'is_subtotal'}:
            raise ValueError('each series must contain name, values, and optional is_subtotal')
        name = item.get('name')
        if not isinstance(name, str) or not name.strip():
            raise ValueError('series name must be a nonblank string')
        values = item.get('values')
        if not isinstance(values, list) or len(values) != len(categories):
            raise ValueError('series values must be a list matching categories length')
        if any(
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not math.isfinite(value)
            for value in values
        ):
            raise ValueError('series values must contain only finite numbers')
        if 'is_subtotal' in item:
            if kind != 'waterfall':
                raise ValueError('is_subtotal is supported only for waterfall charts')
            flags = item['is_subtotal']
            if (
                not isinstance(flags, list)
                or len(flags) != len(categories)
                or any(not isinstance(flag, bool) for flag in flags)
            ):
                raise ValueError('is_subtotal must be a boolean list matching categories length')

    path = ensure_path(doc_id)
    document = aw.Document(str(path))
    builder = aw.DocumentBuilder(document)
    builder.move_to_document_end()
    shape = builder.insert_chart(_CHART_TYPES[kind], width, height)
    chart = shape.chart
    chart.series.clear()
    for item in series:
        values = [float(value) for value in item['values']]
        if kind == 'waterfall':
            chart.series.add(
                item['name'], categories, values, item.get('is_subtotal', [False] * len(categories))
            )
        else:
            chart.series.add(item['name'], categories, values)
    # Footer charts can follow the body insertion point in traversal order.
    chart_index = _chart_shapes(document).index(shape)
    document.save(str(path), aw.SaveFormat.DOCX)
    return chart_index


def list_charts(doc_id: str) -> list[dict]:
    document = aw.Document(str(ensure_path(doc_id)))
    result = []
    for index, shape in enumerate(_chart_shapes(document)):
        series = shape.chart.series
        result.append(
            {
                'chartIndex': index,
                'seriesCount': series.count,
                'series': [
                    {
                        'seriesIndex': i,
                        'name': series[i].name,
                        'seriesType': series[i].series_type.name.lower(),
                    }
                    for i in range(series.count)
                ],
            }
        )
    return result


def _color(value: str | None, parameter: str) -> Color | None:
    if value is None:
        return None
    if not isinstance(value, str) or not re.fullmatch(r'#?[0-9a-fA-F]{6}', value.strip()):
        raise ValueError(f'{parameter} must be a six-digit hexadecimal color, optionally with #')
    digits = value.strip().removeprefix('#')
    return Color.from_argb(int(digits[0:2], 16), int(digits[2:4], 16), int(digits[4:6], 16))


def format_chart_plot_area(
    doc_id: str,
    chart_index: int,
    fill_color: str | None = None,
    line_color: str | None = None,
    line_width: float | None = None,
) -> None:
    if fill_color is None and line_color is None and line_width is None:
        raise ValueError('provide at least one of fill_color, line_color, or line_width')
    fill = _color(fill_color, 'fill_color')
    line = _color(line_color, 'line_color')
    if line_width is not None:
        _positive_number(line_width, 'line_width')
    path = ensure_path(doc_id)
    document = aw.Document(str(path))
    plot_format = _chart(document, chart_index).plot_area.format
    if fill is not None:
        plot_format.fill.solid(fill)
    if line is not None:
        plot_format.stroke.fore_color = line
    if line_width is not None:
        plot_format.stroke.weight = line_width
    document.save(str(path), aw.SaveFormat.DOCX)


def get_waterfall_totals(doc_id: str, chart_index: int, series_index: int = 0) -> dict:
    document = aw.Document(str(ensure_path(doc_id)))
    chart = _chart(document, chart_index)
    _validate_index(series_index, chart.series.count, 'series_index')
    series = chart.series[series_index]
    if series.series_type != charts.ChartSeriesType.WATERFALL:
        raise ValueError('get_waterfall_totals requires a waterfall series')
    points = []
    total_indices = []
    for index in range(series.x_values.count):
        if series.is_total(index):
            total_indices.append(index)
        point = series.data_points[index]
        points.append({'index': point.index, 'isTotal': point.is_total})
    return {
        'chartIndex': chart_index,
        'seriesIndex': series_index,
        'points': points,
        'totalIndices': total_indices,
    }
