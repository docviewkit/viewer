// XLS and XLSX share spreadsheet chart geometry and painting. The adapters
// supply their own parsed chart, worksheet anchor and source locator; PPTX/DOCX
// retain their different host layout and source-mapping contracts.
use crate::diagnostic::Fidelity;
use crate::model::{BlendMode, Object, ObjectKind, SourceRef};

fn chart_object_limit_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::ObjectLimit, Phase::Parse, None, message).in_part(part)
}

pub(super) fn push_spreadsheet_chart(
    mut chart: Chart,
    bounds: Rect,
    source: &SourceRef,
    unit_index: u32,
    object_limit: usize,
    state: &mut Vec<Object>,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<(), Diagnostic> {
    chart.resolve_value_axis_layout(bounds);
    let point_count = chart
        .series
        .iter()
        .map(|series| series.values.len())
        .sum::<usize>();
    let legend_object_count = if chart.show_legend {
        let entries = chart.legend_entries();
        entries
            .len()
            .saturating_mul(2)
            .saturating_add(
                entries
                    .iter()
                    .filter(|(index, _, _)| chart.legend_marker(*index).is_some())
                    .count(),
            )
            .saturating_add(usize::from(
                chart.legend_fill.is_some() || chart.legend_stroke.is_some(),
            ))
    } else {
        0
    };
    let title_object_count =
        usize::from(chart.show_title) + usize::from(chart.show_title && chart.title_fill.is_some());
    let data_label_border_count = chart
        .series
        .iter()
        .flat_map(|series| &series.data_labels)
        .filter(|label| label.border_color.is_some())
        .count();
    if state
        .len()
        .checked_add(point_count.saturating_mul(2))
        .and_then(|count| {
            count
                .checked_add(5 + legend_object_count + title_object_count + data_label_border_count)
        })
        .is_none_or(|count| count > object_limit)
    {
        return Err(chart_object_limit_error(
            &chart.source_part,
            "chart objects exceed the configured object limit",
        ));
    }
    let chart_id = u32::try_from(state.len())
        .map_err(|_| format_error(&chart.source_part, "object count exceeds supported range"))?;
    let (area_stroke, area_stroke_width) =
        chart.area_border(bounds, (Paint::Solid(0xd1d5_dbff), 1.0));
    state.push(Object {
        numeric_id: chart_id,
        parent_numeric_id: None,
        stable_id: format!("object:{chart_id}"),
        parent_stable_id: None,
        kind: ObjectKind::Group,
        unit_index,
        bounds,
        z: i32::try_from(chart_id).unwrap_or(i32::MAX),
        text: None,
        source: source.clone(),
        visual: Visual::PaintedShape {
            geometry: Geometry::Rectangle,
            fill: chart
                .chart_area_fill
                .as_ref()
                .map_or(Paint::Solid(0xffff_ffff), |fill| fill.paint(bounds)),
            stroke: area_stroke,
            stroke_width: area_stroke_width,
        },
    });
    if let Some(elements) = chart_extended_elements(&chart, bounds, object_limit)? {
        if state.len().saturating_add(elements.len()) > object_limit {
            return Err(chart_object_limit_error(
                &chart.source_part,
                "chart elements exceed the object limit",
            ));
        }
        for element in elements {
            let numeric_id = state.len() as u32;
            state.push(Object {
                numeric_id,
                parent_numeric_id: Some(chart_id),
                stable_id: format!("object:{numeric_id}"),
                parent_stable_id: Some(format!("object:{chart_id}")),
                kind: if element.text.is_some() {
                    ObjectKind::TextBox
                } else {
                    ObjectKind::Shape
                },
                unit_index,
                bounds: element.bounds,
                z: numeric_id as i32,
                text: element.text,
                source: source.clone(),
                visual: element.visual,
            });
        }
        return Ok(());
    }
    if chart.show_title {
        let title_style = chart.title_text_style();
        let title_bounds = chart.positioned_title_bounds(bounds, chart_title_bounds(
            bounds,
            bounds.y + 6.0,
            28.0,
            chart.title_text(),
            title_style.font_size,
        ));
        if let Some(fill) = chart.title_fill.as_ref() {
            push_spreadsheet_chart_shape(
                state,
                chart_id,
                unit_index,
                title_bounds,
                Geometry::Rectangle,
                fill.paint(title_bounds),
                Paint::None,
                0.0,
                source,
            )?;
        }
        push_spreadsheet_chart_text(
            state,
            chart_id,
            unit_index,
            title_bounds,
            chart.title_text().to_owned(),
            title_style,
            0.0,
            TextOrientation::Horizontal,
            None,
            source,
            usize::MAX,
        )?;
    }
    let has_3d_area = chart
        .series
        .iter()
        .any(|series| series.kind == ChartKind::Area && series.three_d);
    let reserve_axis_titles = !chart.value_axis_options.title.trim().is_empty();
    let reserve_category_labels = !chart.horizontal_axis_options.title.trim().is_empty()
        || has_3d_area
        || chart
            .horizontal_axis_options
            .label_rotation_degrees
            .is_some_and(|rotation| rotation.abs() >= 45.0);
    let stacked_value_title = matches!(
        chart.value_axis_options.title_orientation,
        Some(TextOrientation::StackedRl | TextOrientation::StackedLr)
    );
    let stacked_series_title = chart.series_axis_options.as_ref().is_some_and(|axis| {
        matches!(
            axis.title_orientation,
            Some(TextOrientation::StackedRl | TextOrientation::StackedLr)
        )
    });
    let horizontal_value_title = chart.value_axis_options.title_orientation
        == Some(TextOrientation::Horizontal)
        && chart.value_axis_options.title_rotation_degrees == Some(0.0);
    let compact_horizontal_3d = chart.uses_compact_horizontal_3d_layout();
    let mut plot = Rect {
        x: bounds.x
            + bounds.width
                * if has_3d_area {
                    0.14
                } else if chart.data_table.is_some() {
                    0.17
                } else if horizontal_value_title {
                    0.22
                } else if stacked_value_title {
                    0.18
                } else if reserve_axis_titles {
                    0.14
                } else if compact_horizontal_3d {
                    0.06
                } else {
                    0.10
                },
        y: bounds.y
            + bounds.height
                * if chart.show_title {
                    0.16
                } else if chart.data_table.is_some() {
                    0.05
                } else if compact_horizontal_3d {
                    0.06
                } else {
                    0.08
                },
        width: bounds.width
            * if has_3d_area {
                0.66
            } else if chart.data_table.is_some() {
                0.80
            } else if horizontal_value_title {
                0.72
            } else if stacked_value_title {
                0.76
            } else if reserve_axis_titles {
                0.80
            } else {
                0.84
            },
        height: bounds.height
            * if has_3d_area {
                0.70
            } else if chart.data_table.is_some() {
                0.70
            } else if reserve_category_labels {
                if chart.show_title { 0.60 } else { 0.68 }
            } else if chart.show_title {
                0.74
            } else {
                0.82
            },
    };
    let manual_plot = chart.plot_bounds.is_some();
    plot = chart.plot_area_bounds(bounds, plot);
    if chart.show_legend && !manual_plot && chart.automatic_cartesian_layout(bounds).is_none()
        && chart.constrained_legend(bounds).is_none() {
        match chart.legend_position {
            ChartLegendPosition::Left => {
                let reserve = bounds.width * 0.22;
                plot.x += reserve;
                plot.width -= reserve;
            }
            ChartLegendPosition::Right => {
                plot.width -= bounds.width
                    * if compact_horizontal_3d {
                        0.12
                    } else if stacked_series_title {
                        0.36
                    } else {
                        0.22
                    }
            }
            ChartLegendPosition::Top => {
                plot.y += bounds.height * 0.16;
                plot.height -= bounds.height * 0.16;
            }
            ChartLegendPosition::Bottom => plot.height -= bounds.height * 0.16,
            ChartLegendPosition::TopRight => {}
        }
    }
    if let Some(fill) = chart.plot_area_fill.as_ref() {
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            plot,
            Geometry::Rectangle,
            fill.paint(plot),
            Paint::None,
            0.0,
            source,
        )?;
    } else if let Some(color) = chart.plot_area_color {
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            plot,
            Geometry::Rectangle,
            Paint::Solid(color),
            Paint::None,
            0.0,
            source,
        )?;
    }
    let has_cartesian_series = chart.series.iter().any(|series| {
        matches!(
            series.kind,
            ChartKind::Bar | ChartKind::Waterfall | ChartKind::Line | ChartKind::Area
        )
    });
    if chart.series.iter().any(|series| {
        matches!(
            series.kind,
            ChartKind::Pie | ChartKind::BarOfPie(_) | ChartKind::Doughnut
        )
    }) {
        push_spreadsheet_pie_chart(&chart, chart_id, bounds, plot, source, unit_index, state)?;
    } else if has_cartesian_series {
        push_spreadsheet_chart_axes(&chart, chart_id, bounds, plot, source, unit_index, state)?;
        let (mut bars_rendered, mut areas_rendered, mut lines_rendered) = (false, false, false);
        for series in &chart.series {
            match series.kind {
                ChartKind::Bar | ChartKind::Waterfall if !bars_rendered => {
                    bars_rendered = true;
                    push_spreadsheet_bar_chart(&chart, chart_id, plot, source, unit_index, state)?;
                }
                ChartKind::Area if !areas_rendered => {
                    areas_rendered = true;
                    push_spreadsheet_area_chart(&chart, chart_id, plot, source, unit_index, state)?;
                }
                ChartKind::Line if !lines_rendered => {
                    lines_rendered = true;
                    push_spreadsheet_line_chart(&chart, chart_id, plot, source, unit_index, state)?;
                }
                _ => {}
            }
        }
    } else {
        match chart.series.first().map(|series| series.kind) {
            Some(ChartKind::Bubble) => push_spreadsheet_bubble_chart(
                &chart, chart_id, bounds, plot, source, unit_index, state,
            )?,
            Some(ChartKind::Scatter) => push_spreadsheet_scatter_chart(
                &chart, chart_id, bounds, plot, source, unit_index, state,
            )?,
            Some(ChartKind::BoxWhisker) => push_spreadsheet_box_whisker_chart(
                &chart, chart_id, bounds, plot, source, unit_index, state,
            )?,
            Some(ChartKind::Treemap) | None => {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Render,
                        Fidelity::Omitted,
                        "Chart kind is parsed but has no worksheet renderer",
                    )
                    .in_part(&chart.source_part),
                );
            }
            Some(
                ChartKind::Bar
                | ChartKind::Waterfall
                | ChartKind::Line
                | ChartKind::Area
                | ChartKind::Pie
                | ChartKind::BarOfPie(_)
                | ChartKind::Doughnut
                | ChartKind::Stock
                | ChartKind::Radar
                | ChartKind::Surface
                | ChartKind::SurfaceWireframe
                | ChartKind::Line3D
                | ChartKind::PieOfPie(_)
                | ChartKind::Funnel
                | ChartKind::Sunburst
                | ChartKind::Histogram
                | ChartKind::Pareto,
            ) => unreachable!("handled chart kind"),
        }
    }
    if chart.data_table.is_some() {
        push_spreadsheet_chart_data_table(
            &chart, chart_id, bounds, plot, source, unit_index, state,
        )?;
    }
    if chart.show_legend {
        push_spreadsheet_chart_legend(&chart, chart_id, bounds, source, unit_index, state)?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_scatter_chart(
    chart: &Chart,
    chart_id: u32,
    bounds: Rect,
    plot: Rect,
    source: &SourceRef,
    unit_index: u32,
    state: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    push_spreadsheet_chart_axes(chart, chart_id, bounds, plot, source, unit_index, state)?;
    for (series_index, series) in chart.series.iter().enumerate() {
        let color = series
            .color
            .unwrap_or_else(|| super::office_chart_palette_color(series_index));
        let points = chart.scatter_points(series, plot);
        for (_, element) in chart.scatter_error_bars(series, plot) {
            if let Visual::PaintedShape {
                geometry,
                fill,
                stroke,
                stroke_width,
            } = element.visual
            {
                push_spreadsheet_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    element.bounds,
                    geometry,
                    fill,
                    stroke,
                    stroke_width,
                    source,
                )?;
            }
        }
        if let Some((line_bounds, geometry)) = chart.scatter_line_geometry(series, plot) {
            push_spreadsheet_chart_shape(state, chart_id, unit_index, line_bounds, geometry, Paint::None,
                Paint::Solid(color), series.stroke_width.unwrap_or(2.0), source)?;
        }
        let Some(symbol) = series.marker_symbol.as_deref() else { continue; };
        let marker = series.marker_size.unwrap_or(5.0).clamp(2.0, 20.0);
        let Some((geometry, filled)) = chart_marker_geometry(symbol, marker) else { continue; };
        for (x, y, _) in points {
            push_spreadsheet_chart_shape(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: x - marker / 2.0,
                    y: y - marker / 2.0,
                    width: marker,
                    height: marker,
                },
                geometry.clone(),
                if filled { Paint::Solid(color) } else { Paint::None },
                Paint::Solid(color),
                1.0,
                source,
            )?;
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_box_whisker_chart(
    chart: &Chart,
    chart_id: u32,
    bounds: Rect,
    plot: Rect,
    source: &SourceRef,
    unit_index: u32,
    state: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    push_spreadsheet_chart_axes(chart, chart_id, bounds, plot, source, unit_index, state)?;
    let (minimum, maximum, _) = chart.value_axis();
    let count = chart.series.len();
    let gap_width = chart.category_gap_width.unwrap_or(1.0);
    let label_width = plot.width / count.max(1) as f32;
    for (index, series) in chart.series.iter().enumerate() {
        let Some(box_whisker) = chart_box_whisker(
            &series.values,
            index,
            count,
            plot,
            (minimum, maximum),
            gap_width,
        ) else {
            continue;
        };
        let color = series
            .color
            .unwrap_or_else(|| super::office_chart_palette_color(index));
        let center_x = box_whisker.bounds.x + box_whisker.bounds.width / 2.0;
        for (start, end) in [
            (box_whisker.upper_whisker_y, box_whisker.bounds.y),
            (
                box_whisker.bounds.y + box_whisker.bounds.height,
                box_whisker.lower_whisker_y,
            ),
        ] {
            push_spreadsheet_chart_shape(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: center_x,
                    y: start.min(end),
                    width: 0.0,
                    height: (end - start).abs(),
                },
                Geometry::Line,
                Paint::None,
                Paint::Solid(0x5959_59ff),
                1.0,
                source,
            )?;
        }
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            box_whisker.bounds,
            Geometry::Rectangle,
            Paint::Solid((color & 0xffff_ff00) | 0x80),
            Paint::Solid(color),
            1.5,
            source,
        )?;
        for y in [
            box_whisker.upper_whisker_y,
            box_whisker.lower_whisker_y,
            box_whisker.median_y,
        ] {
            push_spreadsheet_chart_shape(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: box_whisker.bounds.x,
                    y,
                    width: box_whisker.bounds.width,
                    height: 0.0,
                },
                Geometry::Line,
                Paint::None,
                Paint::Solid(0x5959_59ff),
                1.0,
                source,
            )?;
        }
        for y in std::iter::once(box_whisker.mean_y).chain(box_whisker.outlier_y) {
            push_spreadsheet_chart_shape(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: center_x - 3.0,
                    y: y - 3.0,
                    width: 6.0,
                    height: 6.0,
                },
                Geometry::Ellipse,
                Paint::Solid(0xffff_ffff),
                Paint::Solid(color),
                1.0,
                source,
            )?;
        }
        push_spreadsheet_chart_text(
            state,
            chart_id,
            unit_index,
            Rect {
                x: plot.x + label_width * index as f32,
                y: plot.y + plot.height + 6.0,
                width: label_width,
                height: 18.0,
            },
            if series.name.trim().is_empty() {
                (index + 1).to_string()
            } else {
                series.name.clone()
            },
            ChartTextStyle {
                color: 0x5959_59ff,
                font_size: 12.0,
                bold: false,
                align: TextAlign::Center,
                shadow: None,
            },
            0.0,
            TextOrientation::Horizontal,
            None,
            source,
            index,
        )?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_bubble_chart(
    chart: &Chart,
    chart_id: u32,
    bounds: Rect,
    plot: Rect,
    source: &SourceRef,
    unit_index: u32,
    state: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    push_spreadsheet_chart_axes(chart, chart_id, bounds, plot, source, unit_index, state)?;
    for (series_index, series) in chart.series.iter().enumerate() {
        if series.kind != ChartKind::Bubble {
            continue;
        }
        let color = series
            .color
            .unwrap_or_else(|| super::office_chart_palette_color(series_index));
        for (bubble, _) in chart_bubble_bounds(chart, series_index, plot) {
            push_spreadsheet_chart_shape(
                state,
                chart_id,
                unit_index,
                bubble,
                Geometry::Ellipse,
                chart_bubble_paint(bubble, color, series.three_d),
                Paint::Solid(color),
                1.0,
                source,
            )?;
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_chart_data_table(
    chart: &Chart,
    chart_id: u32,
    bounds: Rect,
    plot: Rect,
    source: &SourceRef,
    unit_index: u32,
    state: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    let Some(layout) = chart_data_table_layout(chart, bounds, plot) else {
        return Ok(());
    };
    for line in layout.lines {
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            line.bounds,
            line.geometry.clone(),
            if matches!(line.geometry, Geometry::Rectangle) { Paint::Solid(line.color) } else { Paint::None },
            Paint::Solid(line.color),
            line.width,
            source,
        )?;
    }
    for text in layout.texts {
        push_spreadsheet_chart_text(
            state,
            chart_id,
            unit_index,
            text.bounds,
            text.text,
            ChartTextStyle {
                color: 0x0000_00ff,
                font_size: 10.0 * 96.0 / 72.0,
                bold: false,
                align: text.align,
                shadow: None,
            },
            0.0,
            TextOrientation::Horizontal,
            None,
            source,
            text.series_index.unwrap_or(0),
        )?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_chart_legend(
    chart: &Chart,
    chart_id: u32,
    bounds: Rect,
    source: &SourceRef,
    unit_index: u32,
    state: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    let constrained = chart.constrained_legend(bounds);
    let entries = constrained.as_ref().map_or_else(|| chart.legend_entries(), |(_, _, entries)| entries.clone());
    let mut style = chart.legend_text_style();
    let horizontal = chart.legend_position.horizontal();
    let mut line_height = style.font_size * 1.45;
    let item_width = (bounds.width * 0.8 / entries.len().max(1) as f32).min(96.0);
    let legend_width = if horizontal {
        item_width * entries.len() as f32
    } else {
        bounds.width
            * if chart.uses_compact_horizontal_3d_layout() {
                0.12
            } else {
                0.20
            }
    };
    let legend_height = if horizontal {
        line_height + 8.0
    } else {
        line_height * entries.len() as f32 + 8.0
    };
    let legend = constrained.as_ref().map(|(bounds, _, _)| *bounds).unwrap_or_else(|| chart.legend_bounds.map_or(
        chart.automatic_cartesian_layout(bounds).and_then(|(_, legend)| legend).unwrap_or(Rect {
            x: match chart.legend_position {
                ChartLegendPosition::Left => bounds.x + 4.0,
                ChartLegendPosition::Top | ChartLegendPosition::Bottom => {
                    bounds.x + (bounds.width - legend_width) / 2.0
                }
                ChartLegendPosition::Right | ChartLegendPosition::TopRight => {
                    bounds.x + bounds.width - legend_width - 4.0
                }
            },
            y: match chart.legend_position {
                ChartLegendPosition::Top => bounds.y + if chart.show_title { 34.0 } else { 6.0 },
                ChartLegendPosition::Bottom => bounds.y + bounds.height - legend_height - 6.0,
                ChartLegendPosition::TopRight => bounds.y + 6.0,
                ChartLegendPosition::Left | ChartLegendPosition::Right => {
                    bounds.y + (bounds.height - legend_height) / 2.0
                }
            },
            width: legend_width,
            height: legend_height,
        }),
        |legend| Rect {
            x: bounds.x + bounds.width * legend.x,
            y: bounds.y + bounds.height * legend.y,
            width: bounds.width * legend.width,
            height: bounds.height * legend.height,
        },
    ));
    if let Some((_, step, _)) = constrained.as_ref() { line_height = *step; }
    let content_scale = if chart.legend_bounds.is_some() {
        (legend.height / legend_height).min(1.0)
    } else {
        1.0
    };
    style.font_size *= content_scale;
    line_height *= content_scale;
    let inset_x = 4.0 * content_scale;
    let inset_y = 6.0 * content_scale;
    let swatch = style.font_size * 0.65;
    if chart.legend_fill.is_some() || chart.legend_stroke.is_some() {
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            legend,
            Geometry::Rectangle,
            chart
                .legend_fill
                .as_ref()
                .map_or(Paint::None, |fill| fill.paint(legend)),
            chart
                .legend_stroke
                .as_ref()
                .map_or(Paint::None, |stroke| stroke.paint(legend)),
            chart.legend_stroke_width,
            source,
        )?;
    }
    for (offset, (index, label, color)) in entries.into_iter().enumerate() {
        let x = legend.x
            + inset_x
            + if horizontal {
                offset as f32 * item_width
            } else {
                0.0
            };
        let y = legend.y
            + inset_y
            + if horizontal {
                0.0
            } else {
                offset as f32 * line_height
            };
        let legend_key_is_line = chart.legend_key_is_line(index);
        let key_width = chart.legend_key_width(index, swatch);
        let key_bounds = Rect {
            x,
            y: y + if legend_key_is_line {
                swatch / 2.0
            } else {
                0.0
            },
            width: key_width,
            height: if legend_key_is_line { 0.01 } else { swatch },
        };
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            key_bounds,
            if legend_key_is_line {
                Geometry::Line
            } else {
                Geometry::Rectangle
            },
            if legend_key_is_line {
                Paint::None
            } else {
                chart
                    .legend_fill(index)
                    .map_or(Paint::Solid(color), |fill| fill.paint(key_bounds))
            },
            if legend_key_is_line {
                Paint::Solid(color)
            } else {
                Paint::None
            },
            if legend_key_is_line { 2.0 } else { 0.0 },
            source,
        )?;
        if let Some((symbol, size)) = chart.legend_marker(index) {
            let size = (size * content_scale).clamp(2.0, swatch * 1.25);
            if let Some((geometry, filled)) = chart_marker_geometry(symbol, size) {
                push_spreadsheet_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    Rect {
                        x: x + (key_width - size) / 2.0,
                        y: y + (swatch - size) / 2.0,
                        width: size,
                        height: size,
                    },
                    geometry,
                    if filled {
                        Paint::Solid(color)
                    } else {
                        Paint::None
                    },
                    Paint::Solid(color),
                    1.0,
                    source,
                )?;
            }
        }
        push_spreadsheet_chart_text(
            state,
            chart_id,
            unit_index,
            Rect {
                x: x + key_width + 4.0,
                y: y - 3.0 * content_scale,
                width: if horizontal { item_width } else { legend.width }
                    - key_width
                    - 10.0 * content_scale,
                height: line_height,
            },
            label,
            style,
            0.0,
            TextOrientation::Horizontal,
            None,
            source,
            index,
        )?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_bar_chart(
    chart: &Chart,
    chart_id: u32,
    plot: Rect,
    source: &SourceRef,
    unit_index: u32,
    state: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    let view_3d = chart.view_3d.unwrap_or_default();
    if chart
        .series
        .iter()
        .any(|series| series.kind == ChartKind::Bar && series.three_d)
    {
        for (wall_bounds, wall_geometry, wall_fill) in chart_bar_3d_walls(plot, view_3d) {
            push_spreadsheet_chart_shape(
                state,
                chart_id,
                unit_index,
                wall_bounds,
                wall_geometry,
                wall_fill,
                Paint::None,
                0.0,
                source,
            )?;
        }
        let (floor_bounds, floor_geometry, floor_fill) = chart.bar_floor(plot);
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            floor_bounds,
            floor_geometry,
            floor_fill,
            Paint::Solid(0x2222_22ff),
            0.75,
            source,
        )?;
    }
    let mut bar_series = chart
        .series
        .iter()
        .enumerate()
        .filter(|(_, series)| series.kind == ChartKind::Bar)
        .collect::<Vec<_>>();
    if bar_series.iter().any(|(_, series)| series.bar_depth)
        && f32::from(chart.view_3d.unwrap_or_default().rot_y).to_radians().cos() >= 0.0 {
        bar_series.reverse();
    }
    for (series_index, series) in bar_series {
        let (minimum, maximum, _) = chart.value_axis_for_series(series);
        if maximum <= minimum {
            continue;
        }
        let color = series
            .color
            .unwrap_or_else(|| super::office_chart_palette_color(series_index));
        for (category_index, value) in series.values.iter().copied().enumerate() {
            let Some(mut bar_bounds) = chart_bar_segment_bounds(
                &chart,
                series_index,
                category_index,
                plot,
                (minimum, maximum),
                chart.bar_gap_width_percent,
            ) else {
                continue;
            };
            if series.three_d && !series.projected_column() {
                bar_bounds =
                    chart_bar_3d_plane_bounds(bar_bounds, plot, view_3d, series.bar_horizontal);
            }
            let color = series.point_colors.get(category_index).copied().unwrap_or(color);
            let bar_fill = series.bar_paint(category_index, bar_bounds, color);
            push_spreadsheet_chart_shape(
                state,
                chart_id,
                unit_index,
                bar_bounds,
                if let Some([front, _, _]) = chart.depth_box_faces(series_index, category_index, plot, bar_bounds) {
                    front
                } else if series.bar_cone_to_max {
                    chart_bar_cone_to_max_geometry(
                        bar_bounds,
                        series.bar_horizontal,
                        value >= 0.0,
                        value.abs()
                            / if value >= 0.0 { maximum } else { minimum.abs() }
                                .max(f32::MIN_POSITIVE),
                    )
                } else {
                    chart_bar_geometry(
                        bar_bounds,
                        series.bar_horizontal,
                        series.bar_cone,
                        series.bar_cylinder,
                    )
                },
                bar_fill,
                series.point_stroke(category_index),
                series.point_stroke_width(category_index),
                source,
            )?;
            series.apply_effects(state.last_mut());
            if series.bar_cone_to_max
                && let Some((cap_bounds, cap_geometry, cap_fill)) = chart_bar_cone_cap(
                    bar_bounds,
                    color,
                    series.bar_horizontal,
                    value >= 0.0,
                    value.abs()
                        / if value >= 0.0 { maximum } else { minimum.abs() }.max(f32::MIN_POSITIVE),
                )
            {
                push_spreadsheet_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    cap_bounds,
                    cap_geometry,
                    cap_fill,
                    Paint::None,
                    0.0,
                    source,
                )?;
            } else if series.bar_cylinder {
                let (cap_bounds, cap_geometry, cap_fill) =
                    chart_bar_cylinder_cap(bar_bounds, color, series.bar_horizontal, value >= 0.0);
                push_spreadsheet_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    cap_bounds,
                    cap_geometry,
                    cap_fill,
                    Paint::None,
                    0.0,
                    source,
                )?;
            } else if series.three_d && !series.bar_cone {
                for (face_bounds, face_geometry, face_fill) in
                    chart.bar_faces(series_index, category_index, plot, bar_bounds, color)
                {
                    push_spreadsheet_chart_shape(
                        state,
                        chart_id,
                        unit_index,
                        face_bounds,
                        face_geometry,
                        face_fill,
                        Paint::Solid(0x4444_44ff),
                        0.5,
                        source,
                    )?;
                }
            }
            if let Some((bounds, text, style, border)) =
                chart_bar_data_label(&chart, series, category_index, bar_bounds, value)
            {
                if let Some((color, width)) = border {
                    push_spreadsheet_chart_shape(
                        state,
                        chart_id,
                        unit_index,
                        bounds,
                        Geometry::Rectangle,
                        Paint::None,
                        Paint::Solid(color),
                        width,
                        source,
                    )?;
                }
                push_spreadsheet_chart_text(
                    state,
                    chart_id,
                    unit_index,
                    bounds,
                    text,
                    style,
                    series.data_label_rotation_degrees.unwrap_or(0.0),
                    TextOrientation::Horizontal,
                    None,
                    source,
                    category_index,
                )?;
            }
        }
    }
    let (minimum, maximum, _) = chart.value_axis();
    for (line_bounds, line_geometry) in
        chart_series_line_geometries(chart, plot, (minimum, maximum))
    {
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            line_bounds,
            line_geometry,
            Paint::None,
            Paint::Solid(0x7f7f_7fff),
            0.75,
            source,
        )?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_line_chart(
    chart: &Chart,
    chart_id: u32,
    plot: Rect,
    source: &SourceRef,
    unit_index: u32,
    state: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    let category_count = chart
        .series
        .iter()
        .filter(|series| series.kind == ChartKind::Line)
        .map(|series| series.values.len())
        .max()
        .unwrap_or(0);
    if category_count < 2 {
        return Ok(());
    }
    for (series_index, series) in chart.series.iter().enumerate() {
        if series.kind != ChartKind::Line {
            continue;
        }
        let (minimum, maximum, _) = chart.value_axis_for_series(series);
        if maximum <= minimum {
            continue;
        }
        let color = series
            .color
            .unwrap_or_else(|| super::office_chart_palette_color(series_index));
        let points = chart.line_points(
            series_index,
            plot,
            (minimum, maximum),
            chart.up_down_bars.is_some(),
        );
        if let Some((line_bounds, geometry)) = chart_line_geometry(&points) {
            push_spreadsheet_chart_shape(state, chart_id, unit_index, line_bounds, geometry,
                Paint::None, Paint::Solid(color), series.stroke_width.unwrap_or(2.0), source)?;
            series.apply_effects(state.last_mut());
        }
        if let Some(symbol) = series.marker_symbol.as_deref() {
            let size = series.marker_size.unwrap_or(8.0);
            for &(x, y, _) in &points {
                let Some((geometry, filled)) = chart_marker_geometry(symbol, size) else {
                    continue;
                };
                push_spreadsheet_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    Rect {
                        x: x - size / 2.0,
                        y: y - size / 2.0,
                        width: size,
                        height: size,
                    },
                    geometry,
                    if filled {
                        Paint::Solid(color)
                    } else {
                        Paint::None
                    },
                    Paint::Solid(color),
                    1.0,
                    source,
                )?;
            }
        }
        for &(x, y, index) in &points {
            let Some(value) = series.values.get(index).copied() else {
                continue;
            };
            if let Some((bounds, text, style, border)) = chart_bar_data_label(
                chart,
                series,
                index,
                Rect {
                    x,
                    y,
                    width: 0.0,
                    height: 0.0,
                },
                value,
            ) {
                if let Some((color, width)) = border {
                    push_spreadsheet_chart_shape(
                        state,
                        chart_id,
                        unit_index,
                        bounds,
                        Geometry::Rectangle,
                        Paint::None,
                        Paint::Solid(color),
                        width,
                        source,
                    )?;
                }
                push_spreadsheet_chart_text(
                    state,
                    chart_id,
                    unit_index,
                    bounds,
                    text,
                    style,
                    series.data_label_rotation_degrees.unwrap_or(0.0),
                    TextOrientation::Horizontal,
                    None,
                    source,
                    index,
                )?;
            }
        }
    }
    if let Some(options) = chart.up_down_bars.as_ref() {
        let (minimum, maximum, _) = chart.value_axis();
        for (_, bar_bounds, is_up) in chart.up_down_bar_bounds(plot, (minimum, maximum)) {
            let fill = if is_up {
                options.up_fill.as_ref()
            } else {
                options.down_fill.as_ref()
            };
            let stroke = if is_up {
                options.up_stroke.as_ref()
            } else {
                options.down_stroke.as_ref()
            };
            push_spreadsheet_chart_shape(
                state,
                chart_id,
                unit_index,
                bar_bounds,
                Geometry::Rectangle,
                fill.map_or(
                    Paint::Solid(if is_up { 0xffff_ffff } else { 0x0000_00ff }),
                    |fill| fill.paint(bar_bounds),
                ),
                stroke.map_or(Paint::Solid(0x0000_00ff), |stroke| stroke.paint(bar_bounds)),
                1.0,
                source,
            )?;
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_area_chart(
    chart: &Chart,
    chart_id: u32,
    plot: Rect,
    source: &SourceRef,
    unit_index: u32,
    state: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    for (series_index, series) in chart.series.iter().enumerate() {
        if series.kind != ChartKind::Area {
            continue;
        }
        let (minimum, maximum, _) = chart.value_axis_for_series(series);
        let color = series
            .color
            .unwrap_or_else(|| super::office_chart_palette_color(series_index));
        if series.three_d {
            for (face_bounds, face_geometry, face_fill) in chart_area_3d_faces(
                &chart,
                series_index,
                plot,
                (minimum, maximum),
                chart.view_3d.unwrap_or_default(),
                color,
            ) {
                push_spreadsheet_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    face_bounds,
                    face_geometry,
                    face_fill,
                    Paint::Solid(0x0000_00ff),
                    series.stroke_width.unwrap_or(0.5),
                    source,
                )?;
            }
            continue;
        }
        let area_plot = chart.area_plot_bounds(series, plot).unwrap_or(plot);
        let Some((geometry, _, _)) =
            chart_area_geometry(&chart, series_index, area_plot, (minimum, maximum))
        else {
            continue;
        };
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            area_plot,
            geometry,
            series
                .fill
                .as_ref()
                .map_or(Paint::Solid(color), |fill| fill.paint(area_plot)),
            series.point_stroke(0),
            series.point_stroke_width(0),
            source,
        )?;
        series.apply_effects(state.last_mut());
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_pie_chart(
    chart: &Chart,
    chart_id: u32,
    bounds: Rect,
    plot: Rect,
    source: &SourceRef,
    unit_index: u32,
    state: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    let Some(series) = chart.series.first() else {
        return Ok(());
    };
    let total = series
        .values
        .iter()
        .copied()
        .filter(|value| *value > 0.0)
        .sum::<f32>();
    if total <= 0.0 {
        return Ok(());
    }
    let split = match series.kind {
        ChartKind::BarOfPie(split) => Some(usize::from(split).min(series.values.len())),
        _ => None,
    };
    let three_d = series.three_d && series.kind != ChartKind::Doughnut;
    let size = if split.is_some() {
        plot.height.min(plot.width * 0.58)
    } else {
        plot.width.min(plot.height)
    };
    let pie = if three_d && chart.plot_bounds.is_none() && split.is_none() {
        plot
    } else {
        chart.pie_bounds(
            bounds,
            Rect {
                x: if split.is_some() {
                    plot.x
                } else {
                    plot.x + (plot.width - size) / 2.0
                },
                y: plot.y + (plot.height - size) / 2.0,
                width: size,
                height: size,
            },
            three_d,
        )
    };
    let radius = if three_d {
        pie.width * 0.52
    } else {
        pie.width.min(pie.height) * 0.47
    };
    let vertical_radius = if three_d {
        radius * chart.view_3d.unwrap_or_default().pie_vertical_ratio()
    } else {
        radius
    };
    let depth = if three_d {
        radius * chart.view_3d.unwrap_or_default().pie_depth_ratio()
    } else {
        0.0
    };
    let (_, slices) = chart_pie_slices(
        series,
        (pie.x + pie.width / 2.0, pie.y + (pie.height - depth) / 2.0),
        (radius, vertical_radius),
        depth,
        chart.view_3d.unwrap_or_default().pie_depth_perspective(),
        split,
    );
    // Keep each slice's faces together in the shared back-to-front order.
    for slice in &slices {
        for (side_bounds, side_geometry) in slice.side.iter().chain(&slice.cut_side) {
            push_spreadsheet_chart_shape(
                state,
                chart_id,
                unit_index,
                *side_bounds,
                side_geometry.clone(),
                chart_pie_side_paint(slice.color, *side_bounds),
                Paint::None,
                0.0,
                source,
            )?;
        }
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            slice.bounds,
            slice.geometry.clone(),
            series
                .point_fills
                .get(slice.index)
                .and_then(Option::as_ref)
                .map_or(Paint::Solid(slice.color), |fill| fill.paint(slice.bounds)),
            series.point_border_colors.get(slice.index).copied().flatten().map_or(Paint::None, Paint::Solid),
            series.point_border_widths.get(slice.index).copied().unwrap_or(0.0),
            source,
        )?;
        series.apply_effects(state.last_mut());
    }
    for label in chart_pie_label_layout(chart, series, &slices, bounds, radius, vertical_radius) {
        if let Some((leader_bounds, leader_geometry)) = label.leader {
            push_spreadsheet_chart_shape(
                state,
                chart_id,
                unit_index,
                leader_bounds,
                leader_geometry,
                Paint::None,
                Paint::Solid(label.style.color),
                0.75,
                source,
            )?;
        }
        push_spreadsheet_chart_text(
            state,
            chart_id,
            unit_index,
            label.bounds,
            label.text,
            label.style,
            0.0,
            TextOrientation::Horizontal,
            None,
            source,
            label.index,
        )?;
    }

    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_chart_axes(
    chart: &Chart,
    chart_id: u32,
    bounds: Rect,
    plot: Rect,
    source: &SourceRef,
    unit_index: u32,
    state: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    let horizontal_bars = chart
        .series
        .iter()
        .any(|series| series.kind == ChartKind::Bar && series.bar_horizontal);
    let area_3d = chart
        .series
        .iter()
        .any(|series| series.kind == ChartKind::Area && series.three_d);
    let bar_3d = chart
        .series
        .iter()
        .any(|series| series.kind == ChartKind::Bar && series.three_d);
    let view_3d = chart.view_3d.unwrap_or_default();
    let value_axis_on_right = area_3d && chart.view_3d.is_some_and(|view| view.rot_y > 90);
    let axis_style = ChartTextStyle {
        color: 0x5959_59ff,
        font_size: 9.0 * 96.0 / 72.0,
        bold: false,
        align: TextAlign::Center,
        shadow: None,
    };
    let axis_style = ChartTextStyle {
        font_size: chart.axis_label_font_size(&chart.value_axis_options, axis_style.font_size),
        bold: chart
            .value_axis_options
            .label_bold
            .unwrap_or(axis_style.bold),
        ..axis_style
    };
    let axis_title_style = |axis: &ChartValueAxis| chart.axis_title_text_style(axis);
    let (minimum, maximum, major) = chart.value_axis();
    if maximum > minimum && major > 0.0 {
        for (value, ratio) in chart.value_axis_ticks() {
            let x = plot.x + ratio * plot.width;
            let y = chart.depth_axis_point(plot, 0.0, ratio, 0.0).map_or(plot.y + plot.height - ratio * plot.height, |p| p.1);
            if chart.value_axis_grid_lines_visible() {
                let (grid_bounds, grid_geometry) =
                    chart.bar_grid_line(plot, ratio, bar_3d, horizontal_bars);
                push_spreadsheet_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    grid_bounds,
                    grid_geometry,
                    Paint::None,
                    Paint::Solid(chart.value_axis_grid_style(0xd9d9_d9ff, 0.75).0),
                    chart.value_axis_grid_style(0xd9d9_d9ff, 0.75).1,
                    source,
                )?;
            }
            if chart.axis_labels_visible(horizontal_bars)
                && let Some(tick_bounds) = if value_axis_on_right {
                    Some(Rect {
                        x: plot.x + plot.width,
                        y,
                        width: 4.0,
                        height: 0.01,
                    })
                } else {
                    chart.value_axis_tick_bounds(plot, ratio, horizontal_bars)
                }
            {
                push_spreadsheet_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    tick_bounds,
                    Geometry::Line,
                    Paint::None,
                    Paint::Solid(0x5959_59ff),
                    0.75,
                    source,
                )?;
            }
            if !chart.axis_labels_visible(horizontal_bars) {
                continue;
            }
            push_spreadsheet_chart_text(
                state,
                chart_id,
                unit_index,
                if horizontal_bars {
                    Rect {
                        x: x - 24.0,
                        y: plot.y + plot.height + 6.0,
                        width: 48.0,
                        height: axis_style.font_size * 1.25,
                    }
                } else if value_axis_on_right {
                    Rect {
                        x: plot.x + plot.width + 6.0,
                        y: y - axis_style.font_size * 0.625,
                        width: (bounds.x + bounds.width - plot.x - plot.width - 10.0).max(36.0),
                        height: axis_style.font_size * 1.25,
                    }
                } else {
                    let x = bounds.x
                        + if chart.value_axis_options.title.trim().is_empty() {
                            4.0
                        } else {
                            24.0
                        };
                    Rect {
                        x,
                        y: y - axis_style.font_size * 0.625,
                        width: (chart.depth_axis_point(plot, 0.0, ratio, 0.0).map_or(plot.x, |p| p.0) - x - 8.0).max(18.0),
                        height: axis_style.font_size * 1.25,
                    }
                },
                chart.value_axis_label(value),
                ChartTextStyle {
                    align: if value_axis_on_right {
                        TextAlign::Start
                    } else {
                        TextAlign::End
                    },
                    ..axis_style
                },
                0.0,
                TextOrientation::Horizontal,
                None,
                source,
                usize::MAX,
            )?;
        }
    }
    if !horizontal_bars && let Some((secondary, ticks)) = chart.secondary_value_axis_ticks() {
        let secondary_style = ChartTextStyle {
            font_size: secondary.label_font_size.unwrap_or(axis_style.font_size),
            bold: secondary.label_bold.unwrap_or(axis_style.bold),
            align: TextAlign::Start,
            ..axis_style
        };
        for (value, ratio) in ticks {
            let y = chart.depth_axis_point(plot, 0.0, ratio, 0.0).map_or(plot.y + plot.height * (1.0 - ratio), |p| p.1);
            if let Some(tick) = chart.secondary_value_axis_tick_bounds(plot, ratio) {
                push_spreadsheet_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    tick,
                    Geometry::Line,
                    Paint::None,
                    Paint::Solid(0x5959_59ff),
                    0.75,
                    source,
                )?;
            }
            push_spreadsheet_chart_text(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: plot.x + plot.width + 6.0,
                    y: y - secondary_style.font_size * 0.625,
                    width: (bounds.x + bounds.width - plot.x - plot.width - 10.0).max(18.0),
                    height: secondary_style.font_size * 1.25,
                },
                secondary.format_label(value),
                secondary_style,
                0.0,
                TextOrientation::Horizontal,
                None,
                source,
                usize::MAX,
            )?;
        }
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            Rect {
                x: plot.x + plot.width,
                y: plot.y,
                width: 0.0,
                height: plot.height,
            },
            Geometry::Line,
            Paint::None,
            Paint::Solid(0x6b72_80ff),
            1.0,
            source,
        )?;
    }
    let xy_chart = chart
        .series
        .iter()
        .all(|series| matches!(series.kind, ChartKind::Scatter | ChartKind::Bubble));
    if xy_chart {
        for (value, ratio) in chart.scatter_axis_ticks(true) {
            let x = plot.x + ratio * plot.width;
            if chart.horizontal_axis_options.major_gridlines {
                push_spreadsheet_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    Rect {
                        x,
                        y: plot.y,
                        width: 0.0,
                        height: plot.height,
                    },
                    Geometry::Line,
                    Paint::None,
                    Paint::Solid(0xd9d9_d9ff),
                    0.75,
                    source,
                )?;
            }
            if chart.axis_labels_visible(true) {
                push_spreadsheet_chart_text(
                    state,
                    chart_id,
                    unit_index,
                    Rect {
                        x: x - 28.0,
                        y: plot.y + plot.height + 6.0,
                        width: 56.0,
                        height: axis_style.font_size * 1.25,
                    },
                    chart.horizontal_axis_options.format_label(value),
                    ChartTextStyle {
                        align: TextAlign::Center,
                        ..axis_style
                    },
                    0.0,
                    TextOrientation::Horizontal,
                    None,
                    source,
                    usize::MAX,
                )?;
            }
        }
    }
    if !area_3d
        && !xy_chart
        && chart.axis_labels_visible(!horizontal_bars)
        && let Some(series) = chart.series.iter().find(|series| !series.values.is_empty())
    {
        let axis_style = ChartTextStyle { font_size: chart.depth_category_font_size(series, plot, axis_style.font_size), ..axis_style };
        let count = series.values.len();
        let visible_categories = series
            .categories
            .iter()
            .take(count)
            .enumerate()
            .filter_map(|(index, category)| {
                if horizontal_bars {
                    let slot = plot.height / count.max(1) as f32;
                    Some((
                        index,
                        category,
                        plot.y + (count - chart.category_index(index, count)) as f32 * slot - slot / 2.0,
                    ))
                } else if chart.category_has_label(series, index, plot, 18.0) {
                    chart
                        .category_x(series, index, plot, false)
                        .map(|x| (index, category, x))
                } else {
                    None
                }
            })
            .collect::<Vec<_>>();
        let width = plot.width / visible_categories.len().max(1) as f32;
        let dense = !horizontal_bars && chart.dense_category_labels(series, plot);
        let rotation = chart
            .horizontal_axis_options
            .label_rotation_degrees
            .unwrap_or(if dense { 0.0 } else if visible_categories.len() > 6 {
                -45.0
            } else {
                0.0
            });
        if !horizontal_bars {
            for index in 0..=count {
                let Some(tick_bounds) = chart.category_axis_tick_bounds(series, index, plot) else {
                    continue;
                };
                push_spreadsheet_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    tick_bounds,
                    Geometry::Line,
                    Paint::None,
                    Paint::Solid(0x5959_59ff),
                    0.75,
                    source,
                )?;
            }
        }
        for (index, category, position) in visible_categories {
            push_spreadsheet_chart_text(
                state,
                chart_id,
                unit_index,
                if horizontal_bars {
                    Rect {
                        x: bounds.x + 4.0,
                        y: position - axis_style.font_size * 0.75,
                        width: (plot.x - bounds.x - 8.0).max(18.0),
                        height: axis_style.font_size * 1.5,
                    }
                } else {
                    Rect {
                        x: position - width / 2.0,
                        y: chart.category_baseline(series, index, plot) + 6.0,
                        width,
                        height: (bounds.y + bounds.height - if dense { 10.0 } else { 28.0 } - (plot.y + plot.height))
                            .max(18.0),
                    }
                },
                category.clone(),
                ChartTextStyle { color: chart.category_axis_options().label_color.unwrap_or(axis_style.color), ..axis_style },
                rotation,
                if dense { TextOrientation::Rotated270 } else { TextOrientation::Horizontal },
                None,
                source,
                index,
            )?;
        }
    }
    for (series_index, label_bounds, label) in
        chart_3d_series_axis_labels(chart, plot, axis_style.font_size)
    {
        push_spreadsheet_chart_text(
            state,
            chart_id,
            unit_index,
            label_bounds,
            label,
            ChartTextStyle {
                align: TextAlign::End,
                ..axis_style
            },
            0.0,
            TextOrientation::Horizontal,
            None,
            source,
            series_index,
        )?;
    }
    let (series_labels, category_labels) =
        chart_area_3d_axis_labels(chart, plot, axis_style.font_size);
    for (series_index, label_bounds, label, rotation) in series_labels {
        push_spreadsheet_chart_text(
            state,
            chart_id,
            unit_index,
            label_bounds,
            label,
            ChartTextStyle {
                align: TextAlign::Center,
                ..axis_style
            },
            rotation,
            TextOrientation::Horizontal,
            None,
            source,
            series_index,
        )?;
    }
    for (category_index, label_bounds, label, rotation) in category_labels {
        push_spreadsheet_chart_text(
            state,
            chart_id,
            unit_index,
            label_bounds,
            label,
            ChartTextStyle {
                align: TextAlign::Start,
                ..axis_style
            },
            rotation,
            TextOrientation::Horizontal,
            None,
            source,
            category_index,
        )?;
    }
    for (bounds, geometry, color, width) in chart.minor_axis_lines(plot) {
        push_spreadsheet_chart_shape(state, chart_id, unit_index, bounds, geometry, Paint::None, Paint::Solid(color), width, source)?;
    }
    for (label_bounds, text) in chart.supplemental_axis_labels(plot, axis_style.font_size) {
        push_spreadsheet_chart_text(state, chart_id, unit_index, label_bounds, text,
            axis_style, 0.0, TextOrientation::Horizontal, None, source, usize::MAX)?;
    }
    if !chart.value_axis_options.title.trim().is_empty() {
        let style = axis_title_style(&chart.value_axis_options);
        let stacked = matches!(
            chart.value_axis_options.title_orientation,
            Some(TextOrientation::StackedRl | TextOrientation::StackedLr)
        );
        let stacked_width = style.font_size
            * 1.4
            * chart
                .value_axis_options
                .title
                .split_whitespace()
                .count()
                .max(1) as f32;
        let horizontal_box = chart.value_axis_options.title_orientation
            == Some(TextOrientation::Horizontal)
            && chart.value_axis_options.title_rotation_degrees == Some(0.0);
        let title_bounds = Rect {
            x: if stacked {
                plot.x - stacked_width - bounds.width * 0.02
            } else if horizontal_box {
                bounds.x + 8.0
            } else {
                bounds.x + 20.0
            },
            y: plot.y,
            width: if stacked {
                stacked_width
            } else if horizontal_box {
                bounds.width * 0.22
            } else {
                24.0
            },
            height: plot.height,
        };
        let title_bounds = chart.projected_value_title_bounds(plot, style.font_size).unwrap_or(title_bounds);
        let (rotation, orientation, box_paint) =
            spreadsheet_axis_title_presentation(&chart.value_axis_options, title_bounds, -90.0);
        push_spreadsheet_chart_text(
            state,
            chart_id,
            unit_index,
            title_bounds,
            chart.value_axis_options.title.clone(),
            style,
            rotation,
            orientation,
            box_paint,
            source,
            usize::MAX,
        )?;
    }
    if !chart.horizontal_axis_options.title.trim().is_empty() {
        let style = axis_title_style(&chart.horizontal_axis_options);
        let vertical = matches!(
            chart.horizontal_axis_options.title_orientation,
            Some(TextOrientation::StackedRl | TextOrientation::StackedLr)
        );
        let y = if vertical {
            plot.y + plot.height + 4.0
        } else {
            bounds.y + bounds.height - style.font_size * 1.75
        };
        let title_bounds = Rect {
            x: if vertical {
                plot.x + plot.width / 2.0 - 12.0
            } else {
                plot.x
            },
            y,
            width: if vertical { 24.0 } else { plot.width },
            height: if vertical {
                (bounds.y + bounds.height - y - 4.0).max(24.0)
            } else {
                style.font_size * 1.25
            },
        };
        let (rotation, orientation, box_paint) =
            spreadsheet_axis_title_presentation(&chart.horizontal_axis_options, title_bounds, 0.0);
        push_spreadsheet_chart_text(
            state,
            chart_id,
            unit_index,
            title_bounds,
            chart.horizontal_axis_options.title.clone(),
            style,
            rotation,
            orientation,
            box_paint,
            source,
            usize::MAX,
        )?;
    }
    if let Some(series_axis) = chart.series_axis_options.as_ref() {
        if !series_axis.title.trim().is_empty() {
            let style = axis_title_style(series_axis);
            let (depth_x, depth_y) = chart_3d_offset(plot, view_3d);
            let stacked = matches!(
                series_axis.title_orientation,
                Some(TextOrientation::StackedRl | TextOrientation::StackedLr)
            );
            let horizontal_box = series_axis.title_orientation == Some(TextOrientation::Horizontal)
                && series_axis.title_rotation_degrees == Some(0.0);
            let stacked_width =
                style.font_size * 1.4 * series_axis.title.split_whitespace().count().max(1) as f32;
            let title_bounds = Rect {
                x: if stacked {
                    plot.x + plot.width + depth_x + 4.0
                } else if horizontal_box {
                    bounds.x + bounds.width - bounds.width * 0.25 - 20.0
                } else {
                    plot.x + plot.width + depth_x + 4.0
                },
                y: if stacked {
                    plot.y
                } else {
                    plot.y + plot.height - depth_y - style.font_size * 0.625
                },
                width: if stacked {
                    stacked_width
                } else if horizontal_box {
                    bounds.width * 0.25
                } else {
                    plot.width * 0.22
                },
                height: if stacked {
                    plot.height
                } else {
                    style.font_size * 1.25
                },
            };
            let (rotation, orientation, box_paint) =
                spreadsheet_axis_title_presentation(series_axis, title_bounds, 0.0);
            push_spreadsheet_chart_text(
                state,
                chart_id,
                unit_index,
                title_bounds,
                series_axis.title.clone(),
                style,
                rotation,
                orientation,
                box_paint,
                source,
                usize::MAX,
            )?;
        }
    }
    for (horizontal, bounds) in [
        (
            false,
            Rect {
                x: plot.x,
                y: plot.y,
                width: 0.0,
                height: plot.height,
            },
        ),
        (
            true,
            Rect {
                x: plot.x,
                y: plot.y + plot.height,
                width: plot.width,
                height: 0.0,
            },
        ),
    ] {
        if !chart.axis_line_visible(horizontal) {
            continue;
        }
        let (bounds, geometry) = chart.projected_axis_line(plot, horizontal).unwrap_or((bounds, Geometry::Line));
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            bounds,
            geometry,
            Paint::None,
            Paint::Solid(0x6b72_80ff),
            1.0,
            source,
        )?;
    }
    if bar_3d && !horizontal_bars && chart.series_axis_line_visible() {
        let (depth_x, depth_y) = chart_3d_offset(plot, view_3d);
        push_spreadsheet_chart_shape(
            state,
            chart_id,
            unit_index,
            Rect {
                x: plot.x + plot.width,
                y: plot.y + plot.height - depth_y,
                width: depth_x,
                height: depth_y,
            },
            Geometry::Line,
            Paint::None,
            Paint::Solid(0x6b72_80ff),
            1.0,
            source,
        )?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_chart_shape(
    state: &mut Vec<Object>,
    parent_numeric_id: u32,
    unit_index: u32,
    bounds: Rect,
    geometry: Geometry,
    fill: Paint,
    stroke: Paint,
    stroke_width: f32,
    source: &SourceRef,
) -> Result<(), Diagnostic> {
    let numeric_id = u32::try_from(state.len())
        .map_err(|_| format_error(&source.part, "object count exceeds supported range"))?;
    state.push(Object {
        numeric_id,
        parent_numeric_id: Some(parent_numeric_id),
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: Some(format!("object:{parent_numeric_id}")),
        kind: ObjectKind::Shape,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: source.clone(),
        visual: Visual::PaintedShape {
            geometry,
            fill,
            stroke,
            stroke_width,
        },
    });
    Ok(())
}

type ChartTextBoxPaint = (Paint, Paint, f32, StrokeStyle, DrawingMlPictureEffects);

#[allow(clippy::too_many_arguments)]
fn push_spreadsheet_chart_text(
    state: &mut Vec<Object>,
    parent_numeric_id: u32,
    unit_index: u32,
    bounds: Rect,
    text: String,
    style: ChartTextStyle,
    rotation_degrees: f32,
    orientation: TextOrientation,
    box_paint: Option<ChartTextBoxPaint>,
    source: &SourceRef,
    entry_index: usize,
) -> Result<(), Diagnostic> {
    let numeric_id = u32::try_from(state.len())
        .map_err(|_| format_error(&source.part, "object count exceeds supported range"))?;
    let wrap = box_paint.is_some()
        && orientation == TextOrientation::Horizontal
        && rotation_degrees == 0.0
        && text
            .chars()
            .map(|character| drawingml_fallback_character_width(character, style.font_size))
            .sum::<f32>()
            > (bounds.width - DEFAULT_TEXT_HORIZONTAL_INSET * 2.0).max(1.0);
    let bounds = chart_text_tight_bounds(
        bounds,
        &text,
        style.font_size,
        orientation,
        box_paint.is_some(),
        wrap,
    );
    let visual = spreadsheet_chart_text_visual(
        bounds,
        text.clone(),
        &style,
        rotation_degrees,
        orientation,
        box_paint,
        wrap,
    );
    state.push(Object {
        numeric_id,
        parent_numeric_id: Some(parent_numeric_id),
        stable_id: format!("object:{numeric_id}:legend:{entry_index}"),
        parent_stable_id: Some(format!("object:{parent_numeric_id}")),
        kind: ObjectKind::Shape,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: Some(text),
        source: source.clone(),
        visual,
    });
    Ok(())
}

/// Tighten a chart label's bounds when it carries its own fill/border.
///
/// Excel fills only the text's own bounding box, not the whole reserved slot.
/// We do not have font metrics in the parser, so we estimate with the same
/// heuristic used elsewhere for Latin/Asian text and keep the slot centre.
pub(super) fn chart_text_tight_bounds(
    slot: Rect,
    text: &str,
    font_size: f32,
    orientation: TextOrientation,
    has_box_paint: bool,
    wrap_horizontal: bool,
) -> Rect {
    if !has_box_paint {
        return slot;
    }
    let lines = text.lines().count().max(1) as f32;
    let chars = text.chars().count().max(1) as f32;
    let advance = font_size * 0.65;
    let line_height = font_size * 1.4;
    let horizontal_width = text
        .lines()
        .map(|line| {
            line.chars()
                .map(|character| drawingml_fallback_character_width(character, font_size))
                .sum::<f32>()
        })
        .fold(font_size * 0.52, f32::max);
    let (width, height) = match orientation {
        TextOrientation::VerticalRl | TextOrientation::VerticalLr => {
            // One character per line, stacked vertically.
            (font_size * 1.2, advance * chars.max(lines))
        }
        TextOrientation::Rotated90 | TextOrientation::Rotated270 => {
            (line_height * lines, horizontal_width)
        }
        TextOrientation::StackedRl | TextOrientation::StackedLr => {
            let words = text.split_whitespace().count().max(1) as f32;
            (line_height * words, slot.height)
        }
        TextOrientation::Horizontal => {
            // Horizontal lines, possibly multiple lines separated by '\n'.
            let wrapped_lines = if wrap_horizontal {
                let available_width = (slot.width - DEFAULT_TEXT_HORIZONTAL_INSET * 2.0).max(1.0);
                text.lines()
                    .map(|line| {
                        let width = line
                            .chars()
                            .map(|character| {
                                drawingml_fallback_character_width(character, font_size)
                            })
                            .sum::<f32>();
                        (width / available_width).ceil().max(1.0)
                    })
                    .sum::<f32>()
            } else {
                lines
            };
            (
                if wrap_horizontal {
                    slot.width
                } else {
                    horizontal_width
                },
                line_height * wrapped_lines,
            )
        }
    };
    let cx = slot.x + slot.width / 2.0;
    let cy = slot.y + slot.height / 2.0;
    Rect {
        x: cx - width / 2.0,
        y: cy - height / 2.0,
        width,
        height,
    }
}

/// Rich-text visual for one chart label.
///
/// `box_paint` fills and outlines the label's own rectangle. Because the box
/// geometry is drawn outside `TextLayout`, it would otherwise stay axis-aligned
/// while the glyphs rotated; when a box is painted the rotation therefore moves
/// onto an outer layer so both rotate about the same centre.
fn spreadsheet_chart_text_visual(
    bounds: Rect,
    text: String,
    style: &ChartTextStyle,
    rotation_degrees: f32,
    orientation: TextOrientation,
    box_paint: Option<ChartTextBoxPaint>,
    wrap: bool,
) -> Visual {
    let (fill, stroke, stroke_width, stroke_style, effects, painted) = match box_paint {
        Some((fill, stroke, width, style, effects)) => (fill, stroke, width, style, effects, true),
        None => (
            Paint::None,
            Paint::None,
            0.0,
            StrokeStyle::default(),
            DrawingMlPictureEffects::default(),
            false,
        ),
    };
    let text_layout = Visual::TextLayout {
        layout: TextLayout {
            orientation,
            vertical_align: TextVerticalAlign::Center,
            rotation_degrees: if painted { 0.0 } else { rotation_degrees },
            wrap,
            ..TextLayout::default()
        },
        visual: Box::new(Visual::RichText {
            geometry: Geometry::Rectangle,
            fill,
            stroke,
            stroke_width,
            align: style.align,
            line_height: 0.0,
            runs: vec![TextRun {
                paint: None,
                east_asian_line_breaks: true,
                text,
                font_family: "Calibri".to_owned(),
                font_size: style.font_size,
                color: style.color,
                bold: style.bold,
                italic: false,
                underline: false,
                strikethrough: false,
                highlight: 0,
                baseline_shift: 0.0,
                letter_spacing: 0.0,
                horizontal_scale: 1.0,
            }],
        }),
    };
    let text_layout = effects.wrap(text_layout, None);
    let text_layout = if stroke_style == StrokeStyle::default() {
        text_layout
    } else {
        Visual::StrokeStyle {
            style: stroke_style,
            visual: Box::new(text_layout),
        }
    };
    if painted {
        Visual::Layer {
            transform: AffineTransform::rotation_about(
                bounds.x + bounds.width / 2.0,
                bounds.y + bounds.height / 2.0,
                rotation_degrees,
            ),
            opacity: 1.0,
            blend_mode: BlendMode::Normal,
            visual: Box::new(text_layout),
        }
    } else {
        text_layout
    }
}

/// Rotation, text direction, and box paint authored on an axis title.
///
/// `fallback_rotation_degrees` applies only when the title carries no
/// `a:bodyPr`, matching how Excel orients value-axis titles by default.
pub(super) fn spreadsheet_axis_title_presentation(
    axis: &ChartValueAxis,
    bounds: Rect,
    fallback_rotation_degrees: f32,
) -> (f32, TextOrientation, Option<ChartTextBoxPaint>) {
    let orientation = axis
        .title_orientation
        .unwrap_or(TextOrientation::Horizontal);
    let box_paint = (axis.title_fill.is_some()
        || axis.title_stroke.is_some()
        || !axis.title_effects.is_empty())
    .then(|| {
        (
            axis.title_fill
                .as_ref()
                .map_or(Paint::None, |fill| fill.paint(bounds)),
            axis.title_stroke
                .as_ref()
                .map_or(Paint::None, |stroke| stroke.paint(bounds)),
            axis.title_stroke_width,
            axis.title_stroke_style.clone(),
            axis.title_effects.clone(),
        )
    });
    (
        axis.title_rotation_degrees
            .unwrap_or(fallback_rotation_degrees),
        orientation,
        box_paint,
    )
}
