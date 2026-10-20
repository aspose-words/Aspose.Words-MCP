from __future__ import annotations

import math
import re

import aspose.words as aw
from aspose.pydrawing import Color

from core.utils.docs_util import ensure_path

_GRADIENT_STYLES = {
    'horizontal': aw.drawing.GradientStyle.HORIZONTAL,
    'vertical': aw.drawing.GradientStyle.VERTICAL,
    'diagonal_up': aw.drawing.GradientStyle.DIAGONAL_UP,
    'diagonal_down': aw.drawing.GradientStyle.DIAGONAL_DOWN,
    'from_corner': aw.drawing.GradientStyle.FROM_CORNER,
    'from_center': aw.drawing.GradientStyle.FROM_CENTER,
}
_GRADIENT_VARIANTS = {
    'variant1': aw.drawing.GradientVariant.VARIANT1,
    'variant2': aw.drawing.GradientVariant.VARIANT2,
    'variant3': aw.drawing.GradientVariant.VARIANT3,
    'variant4': aw.drawing.GradientVariant.VARIANT4,
}
_CHART_TYPES = {
    'column': aw.drawing.charts.ChartType.COLUMN,
    'waterfall': aw.drawing.charts.ChartType.WATERFALL,
}


def _normalize(value: str) -> str:
    return value.strip().lower().replace('-', '_').replace(' ', '_')


def _number(value, field: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f'{field} must be a finite number')
    try:
        result = float(value)
    except OverflowError as exc:
        raise ValueError(f'{field} must be a finite number') from exc
    if not math.isfinite(result):
        raise ValueError(f'{field} must be a finite number')
    return result


def _positive_number(value, field: str) -> float:
    result = _number(value, field)
    if result <= 0:
        raise ValueError(f'{field} must be greater than zero')
    return result


def _color(value: str, field: str) -> Color:
    if not isinstance(value, str) or not re.fullmatch(r'#?[0-9a-fA-F]{6}', value.strip()):
        raise ValueError(f'{field} must be a six-digit RGB hex color, optionally prefixed with #')
    rgb = value.strip().removeprefix('#')
    return Color.from_argb(255, int(rgb[:2], 16), int(rgb[2:4], 16), int(rgb[4:], 16))


def _chart_shapes(doc: aw.Document) -> list[aw.drawing.Shape]:
    shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
    return [shape for node in shapes if (shape := node.as_shape()).has_chart]


def _select_chart(doc: aw.Document, chart_index: int) -> aw.drawing.charts.Chart:
    if isinstance(chart_index, bool) or not isinstance(chart_index, int):
        raise ValueError('chart_index must be a zero-based integer')
    shapes = _chart_shapes(doc)
    if chart_index < 0 or chart_index >= len(shapes):
        raise IndexError(
            f'chart_index out of range: {chart_index}; document has {len(shapes)} charts'
        )
    return shapes[chart_index].chart


def add_chart_end(
    doc_id: str,
    chart_type: str,
    categories: list[str],
    series: list[dict],
    width_points: float = 432,
    height_points: float = 252,
) -> int:
    path = ensure_path(doc_id)
    if not isinstance(chart_type, str) or _normalize(chart_type) not in _CHART_TYPES:
        raise ValueError('chart_type must be column or waterfall')
    normalized_type = _normalize(chart_type)
    if not isinstance(categories, list) or not categories:
        raise ValueError('categories must be a nonempty list of nonblank strings')
    if any(not isinstance(item, str) or not item.strip() for item in categories):
        raise ValueError('categories must be a nonempty list of nonblank strings')
    if not isinstance(series, list) or not series:
        raise ValueError('series must be a nonempty list of objects with name and values')
    validated_series = []
    for index, item in enumerate(series):
        fields = {'name', 'values'}
        if normalized_type == 'waterfall':
            fields.add('is_subtotal')
        if isinstance(item, dict) and 'is_subtotal' in item and normalized_type != 'waterfall':
            raise ValueError(f'series[{index}].is_subtotal is only supported for waterfall charts')
        if not isinstance(item, dict) or set(item) != fields:
            expected = (
                'name, values and is_subtotal'
                if normalized_type == 'waterfall'
                else 'name and values'
            )
            raise ValueError(f'series[{index}] must contain only {expected}')
        if not isinstance(item['name'], str) or not item['name'].strip():
            raise ValueError(f'series[{index}].name must be a nonblank string')
        values = item['values']
        if not isinstance(values, list) or len(values) != len(categories):
            raise ValueError(f'series[{index}].values must match categories length')
        subtotals = None
        if normalized_type == 'waterfall':
            subtotals = item['is_subtotal']
            if not isinstance(subtotals, list) or len(subtotals) != len(categories):
                raise ValueError(f'series[{index}].is_subtotal must match categories length')
            if any(not isinstance(flag, bool) for flag in subtotals):
                raise ValueError(f'series[{index}].is_subtotal must contain only booleans')
        validated_series.append(
            (
                item['name'],
                [_number(value, f'series[{index}].values') for value in values],
                subtotals,
            )
        )
    width = _positive_number(width_points, 'width_points')
    height = _positive_number(height_points, 'height_points')
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    builder.move_to_document_end()
    shape = builder.insert_chart(_CHART_TYPES[normalized_type], width, height)
    chart = shape.chart
    chart.series.clear()
    for name, values, subtotals in validated_series:
        if subtotals is None:
            chart.series.add(name, categories, values)
        else:
            chart.series.add(name, categories, values, subtotals)
    chart_index = _chart_shapes(doc).index(shape)
    doc.save(str(path), aw.SaveFormat.DOCX)
    return chart_index


def list_charts(doc_id: str) -> list[dict]:
    doc = aw.Document(str(ensure_path(doc_id)))
    return [
        {
            'chartIndex': chart_index,
            'series': [
                {
                    'seriesIndex': index,
                    'name': shape.chart.series[index].name,
                    'pointCount': shape.chart.series[index].y_values.count,
                }
                for index in range(shape.chart.series.count)
            ],
        }
        for chart_index, shape in enumerate(_chart_shapes(doc))
    ]


def get_waterfall_totals(
    doc_id: str,
    chart_index: int,
    series_index: int = 0,
    point_index: int | None = None,
) -> list[dict]:
    doc = aw.Document(str(ensure_path(doc_id)))
    chart = _select_chart(doc, chart_index)
    if isinstance(series_index, bool) or not isinstance(series_index, int):
        raise ValueError('series_index must be a zero-based integer')
    if series_index < 0 or series_index >= chart.series.count:
        raise IndexError(
            f'series_index out of range: {series_index}; chart has {chart.series.count} series'
        )
    series = chart.series[series_index]
    if series.series_type != aw.drawing.charts.ChartSeriesType.WATERFALL:
        raise ValueError('Total inspection is only supported for waterfall series')
    point_count = series.y_values.count
    if point_index is not None:
        if isinstance(point_index, bool) or not isinstance(point_index, int):
            raise ValueError('point_index must be a zero-based integer or null')
        if point_index < 0 or point_index >= point_count:
            raise IndexError(
                f'point_index out of range: {point_index}; series has {point_count} points'
            )
    indices = range(point_count) if point_index is None else [point_index]
    points = []
    for index in indices:
        point = series.data_points[index]
        points.append(
            {
                'pointIndex': index,
                'dataPointIndex': point.index,
                'isTotal': series.is_total(index),
                'dataPointIsTotal': point.is_total,
            }
        )
    return points


def _gradient_options(gradient: dict):
    if not isinstance(gradient, dict) or 'color' not in gradient:
        raise ValueError(
            'gradient must be an object with color and optional style, variant, degree'
        )
    if set(gradient) - {'color', 'style', 'variant', 'degree'}:
        raise ValueError('gradient supports only color, style, variant, degree')
    color = _color(gradient['color'], 'gradient.color')
    style = gradient.get('style', 'horizontal')
    if not isinstance(style, str) or _normalize(style) not in _GRADIENT_STYLES:
        raise ValueError(f'gradient.style must be one of: {", ".join(_GRADIENT_STYLES)}')
    variant = gradient.get('variant', 'variant1')
    if (
        not isinstance(variant, str)
        or _normalize(variant).replace('_', '') not in _GRADIENT_VARIANTS
    ):
        raise ValueError(f'gradient.variant must be one of: {", ".join(_GRADIENT_VARIANTS)}')
    degree = _number(gradient.get('degree', 0.5), 'gradient.degree')
    if not 0 <= degree <= 1:
        raise ValueError('gradient.degree must be between 0 and 1 inclusive')
    return (
        color,
        _GRADIENT_STYLES[_normalize(style)],
        _GRADIENT_VARIANTS[_normalize(variant).replace('_', '')],
        degree,
    )


def format_chart_plot_area(
    doc_id: str,
    chart_index: int,
    fill_color: str | None = None,
    gradient: dict | None = None,
    line_color: str | None = None,
    line_width_points: float | None = None,
) -> None:
    path = ensure_path(doc_id)
    if fill_color is not None and gradient is not None:
        raise ValueError('fill_color and gradient cannot be specified together')
    solid = _color(fill_color, 'fill_color') if fill_color is not None else None
    gradient_options = _gradient_options(gradient) if gradient is not None else None
    line = _color(line_color, 'line_color') if line_color is not None else None
    width = (
        _positive_number(line_width_points, 'line_width_points')
        if line_width_points is not None
        else None
    )
    doc = aw.Document(str(path))
    chart_format = _select_chart(doc, chart_index).plot_area.format
    if solid is not None:
        chart_format.fill.solid(solid)
    if gradient_options is not None:
        chart_format.fill.one_color_gradient(*gradient_options)
    if line is not None:
        chart_format.stroke.fore_color = line
    if width is not None:
        chart_format.stroke.weight = width
    if line is not None or width is not None:
        chart_format.stroke.visible = True
    if any(value is not None for value in (solid, gradient_options, line, width)):
        doc.save(str(path), aw.SaveFormat.DOCX)
