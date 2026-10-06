// These chart families use one scene builder in all OOXML hosts. Coordinates
// are in the supplied chart frame; host adapters only add source mappings.
struct ChartBinning {
    size: Option<f32>,
    count: Option<usize>,
    underflow: Option<f32>,
    overflow: Option<f32>,
    right_closed: bool,
    aggregate: bool,
}

impl Default for ChartBinning {
    fn default() -> Self {
        Self {
            size: None,
            count: None,
            underflow: None,
            overflow: None,
            right_closed: true,
            aggregate: false,
        }
    }
}

fn chart_histogram(
    values: &[f32],
    labels: &[String],
    options: &ChartBinning,
    limit: usize,
) -> Result<(Vec<f32>, Vec<String>), Diagnostic> {
    if options.aggregate {
        let mut categories = Vec::<String>::new();
        let mut totals = Vec::<f32>::new();
        let mut indices = HashMap::new();
        for (i, &value) in values
            .iter()
            .enumerate()
            .filter(|(_, value)| value.is_finite())
        {
            let name = labels
                .get(i)
                .cloned()
                .unwrap_or_else(|| (i + 1).to_string());
            let index = *indices.entry(name.clone()).or_insert_with(|| {
                let i = categories.len();
                categories.push(name);
                totals.push(0.0);
                i
            });
            totals[index] += value;
        }
        return Ok((totals, categories));
    }
    let values = values
        .iter()
        .copied()
        .filter(|v| v.is_finite())
        .collect::<Vec<_>>();
    if values.is_empty() {
        return Ok((Vec::new(), Vec::new()));
    }
    let minimum = values.iter().copied().fold(f32::INFINITY, f32::min);
    let maximum = values.iter().copied().fold(f32::NEG_INFINITY, f32::max);
    let lower = options.underflow.unwrap_or(minimum);
    let upper = options.overflow.unwrap_or(maximum).max(lower);
    let mean = values.iter().map(|v| *v as f64).sum::<f64>() / values.len() as f64;
    let deviation = (values
        .iter()
        .map(|v| (*v as f64 - mean).powi(2))
        .sum::<f64>()
        / values.len().saturating_sub(1).max(1) as f64)
        .sqrt();
    let size = options.size.unwrap_or_else(|| {
        options.count.map_or_else(
            || (3.5 * deviation / (values.len() as f64).cbrt()) as f32,
            |count| (upper - lower) / count as f32,
        )
    });
    let size = if size.is_finite() && size > 0.0 {
        size
    } else {
        1.0
    };
    let count = ((upper as f64 - lower as f64) / size as f64)
        .ceil()
        .max(1.0);
    if count > limit.saturating_sub(2) as f64 {
        return Err(Diagnostic::fatal(
            crate::diagnostic::DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "histogram bin count exceeds the object limit",
        ));
    }
    let count = count as usize;
    let under = usize::from(options.underflow.is_some());
    let over = usize::from(options.overflow.is_some());
    let mut counts = vec![0.0; under + count + over];
    let mut categories = Vec::with_capacity(counts.len());
    if under > 0 {
        categories.push(format!("≤ {}", format_chart_value(lower, None)));
    }
    for i in 0..count {
        let a = lower + i as f32 * size;
        let b = if options.overflow.is_some() {
            (a + size).min(upper)
        } else {
            a + size
        };
        categories.push(format!(
            "{}{}, {}{}",
            if i == 0 && under == 0 || !options.right_closed {
                "["
            } else {
                "("
            },
            format_chart_value(a, None),
            format_chart_value(b, None),
            if options.right_closed { "]" } else { ")" }
        ));
    }
    if over > 0 {
        categories.push(format!("> {}", format_chart_value(upper, None)));
    }
    for value in values {
        let index = if under > 0 && value <= lower {
            0
        } else if over > 0 && value > upper {
            counts.len() - 1
        } else {
            let position = (value - lower) / size;
            let bin = if options.right_closed {
                position.ceil() - 1.0
            } else {
                position.floor()
            };
            under + (bin.max(0.0) as usize).min(count - 1)
        };
        counts[index] += 1.0;
    }
    Ok((counts, categories))
}

fn extended_chart_text(bounds: Rect, text: String, size: f32, color: u32) -> DrawingMlElement {
    DrawingMlElement {
        bounds,
        text: Some(text.clone()),
        visual: Visual::TextLayout {
            layout: TextLayout {
                vertical_align: TextVerticalAlign::Center,
                auto_fit: crate::model::TextAutoFit::Shrink,
                ..TextLayout::default()
            },
            visual: Box::new(Visual::RichText {
                geometry: Geometry::Rectangle,
                fill: Paint::None,
                stroke: Paint::None,
                stroke_width: 0.0,
                align: TextAlign::Center,
                line_height: 0.0,
                runs: vec![TextRun {
                    text,
                    font_family: "Calibri".to_owned(),
                    font_size: size,
                    color,
                    bold: false,
                    italic: false,
                    underline: false,
                    strikethrough: false,
                    highlight: 0,
                    baseline_shift: 0.0,
                    letter_spacing: 0.0,
                    horizontal_scale: 1.0,
                    east_asian_line_breaks: true,
                    paint: None,
                }],
            }),
        },
    }
}

fn extended_chart_shape(
    bounds: Rect,
    geometry: Geometry,
    fill: Paint,
    stroke: Paint,
) -> DrawingMlElement {
    DrawingMlElement {
        bounds,
        text: None,
        visual: Visual::PaintedShape {
            geometry,
            fill,
            stroke,
            stroke_width: 1.0,
        },
    }
}

fn extended_chart_path(points: &[(f32, f32)], close: bool) -> Geometry {
    let mut commands = Vec::with_capacity(points.len() + usize::from(close));
    for (i, &(x, y)) in points.iter().enumerate() {
        commands.push(if i == 0 {
            PathCommand::MoveTo { x, y }
        } else {
            PathCommand::LineTo { x, y }
        });
    }
    if close {
        commands.push(PathCommand::ClosePath);
    }
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    }
}

// ponytail: automatic surface axes target three bands; size-aware tick density
// is needed to match every Office auto-layout variation. Authored units win.
fn chart_surface_axis(chart: &Chart) -> (f32, f32, f32) {
    let (min, max, _) = chart.value_axis();
    let major = chart
        .value_axis_options
        .major_unit
        .unwrap_or_else(|| nice_chart_step((max - min) / 3.0));
    (
        min,
        chart
            .value_axis_options
            .maximum
            .unwrap_or_else(|| (max / major).ceil() * major),
        major,
    )
}

pub(super) fn chart_extended_elements(
    chart: &Chart,
    bounds: Rect,
    object_limit: usize,
) -> Result<Option<Vec<DrawingMlElement>>, Diagnostic> {
    let Some(kind) = chart.series.iter().map(|s| s.kind).find(|kind| {
        if *kind == ChartKind::Radar && chart.series.iter().any(|s| s.kind != ChartKind::Radar) { return false; }
        matches!(
            kind,
            ChartKind::Stock
                | ChartKind::Surface
                | ChartKind::SurfaceWireframe
                | ChartKind::Line3D
                | ChartKind::PieOfPie(_)
                | ChartKind::BarOfPie(_)
                | ChartKind::Radar
                | ChartKind::Funnel
                | ChartKind::Sunburst
                | ChartKind::Histogram
                | ChartKind::Pareto
        )
    }) else {
        return Ok(None);
    };
    let points = chart.series.iter().map(|s| s.values.len()).sum::<usize>();
    let factor = if kind == ChartKind::Sunburst {
        chart
            .series
            .iter()
            .map(|s| s.category_levels.len().max(1))
            .max()
            .unwrap_or(1)
            .saturating_mul(2)
    } else if kind == ChartKind::Surface {
        let (min, max, major) = chart.value_axis();
        ((max - min) / major.max(f32::EPSILON))
            .ceil()
            .clamp(1.0, 100.0) as usize
            * 2
    } else {
        5
    };
    if points
        .saturating_mul(factor)
        .saturating_add(chart.series.len().saturating_mul(2))
        .saturating_add(64)
        > object_limit
    {
        return Err(Diagnostic::fatal(
            crate::diagnostic::DiagnosticCode::ObjectLimit,
            Phase::Layout,
            None,
            "extended chart geometry exceeds the configured object limit",
        )
        .in_part(&chart.source_part));
    }
    let mut elements = Vec::new();
    let ellipsize = |text: &mut String, font: f32, width: f32| {
        if chart_text_width(text, font) > width {
            while !text.is_empty() && chart_text_width(&format!("{text}…"), font) > width { text.pop(); }
            text.push('…');
        }
    };
    let compact = matches!(kind, ChartKind::BarOfPie(_) | ChartKind::Radar);
    let font = chart.font_size.unwrap_or(14.0).max(1.0);
    if chart.show_title {
        let style = chart.title_text_style();
        let mut title = extended_chart_text(chart.positioned_title_bounds(bounds, Rect {
            x: bounds.x, y: bounds.y + 4.0, width: bounds.width, height: style.font_size * 1.5,
        }), chart.title_text().to_owned(), style.font_size, style.color);
        if let Visual::TextLayout { layout, visual } = &mut title.visual {
            layout.auto_fit = crate::model::TextAutoFit::None;
            layout.wrap = false;
            layout.inset_left = 0.0; layout.inset_right = 0.0; layout.inset_top = 0.0; layout.inset_bottom = 0.0;
            if let Visual::RichText { runs, .. } = visual.as_mut() {
                for run in runs {
                    run.bold = style.bold;
                    ellipsize(&mut run.text, style.font_size, bounds.width - style.font_size);
                }
            }
        }
        elements.push(title);
    }
    // ponytail: automatic margins are fixed; authored ChartML plot bounds win.
    // Full ChartEx layout constraints are required for exact automatic placement.
    let plot = chart.plot_area_bounds(
        bounds,
        Rect {
            x: bounds.x + bounds.width * if matches!(kind, ChartKind::BarOfPie(_) | ChartKind::PieOfPie(_)) { 0.08 } else { 0.14 },
            y: bounds.y + bounds.height * if chart.show_title { 0.17 } else { 0.06 },
            width: bounds.width
                * if chart.show_legend && chart.legend_position == ChartLegendPosition::Right {
                    if matches!(kind, ChartKind::BarOfPie(_) | ChartKind::PieOfPie(_)) { 0.70 } else { 0.64 }
                } else {
                    0.78
                },
            height: bounds.height * 0.70,
        },
    );
    if chart.show_legend {
        let font = if compact { chart.legend_text_style().font_size } else { font };
        let constrained = chart.constrained_legend(bounds);
        let entries = constrained.as_ref().map_or_else(|| chart.legend_entries(), |(_, _, entries)| entries.clone());
        for (i, (series_index, label, color)) in entries.iter().enumerate() {
            let vertical = matches!(
                chart.legend_position,
                ChartLegendPosition::Right | ChartLegendPosition::Left
            );
            let mut rect = if vertical {
                Rect {
                    x: if chart.legend_position == ChartLegendPosition::Right {
                        bounds.x + bounds.width * if compact { 0.79 } else { 0.87 }
                    } else {
                        bounds.x
                    },
                    y: bounds.y
                        + (bounds.height - entries.len() as f32 * font * 1.6) / 2.0
                        + i as f32 * font * 1.6,
                    width: bounds.width * if compact { 0.21 } else { 0.13 },
                    height: font * 1.6,
                }
            } else {
                Rect {
                    x: bounds.x + bounds.width * i as f32 / entries.len().max(1) as f32,
                    y: if chart.legend_position == ChartLegendPosition::Top {
                        bounds.y + bounds.height * if chart.show_title { 0.12 } else { 0.0 }
                    } else {
                        bounds.y + bounds.height - font * 1.8
                    },
                    width: bounds.width / entries.len().max(1) as f32,
                    height: font * 1.6,
                }
            };
            if let Some((area, step, _)) = &constrained {
                rect = Rect { y: area.y + i as f32 * step, height: *step, ..*area };
            }
            if !vertical {
                let width =
                    (label.chars().count() as f32 * font * 0.55 + font * 1.5).min(rect.width);
                rect.x += (rect.width - width) / 2.0;
                rect.width = width;
            }
            elements.push(extended_chart_shape(
                Rect {
                    x: rect.x,
                    y: rect.y + font * 0.45,
                    width: font * 0.65,
                    height: font * 0.65,
                },
                Geometry::Rectangle,
                chart
                    .legend_fill(i)
                    .map_or(Paint::Solid(*color), |fill| fill.paint(rect)),
                Paint::None,
            ));
            if let Some(series) = chart.series.get(*series_index)
                && series.kind == ChartKind::Radar && series.fill.is_some()
                && let Some(key) = elements.last_mut()
            {
                key.visual = series.effects.clone().wrap(key.visual.clone(), None);
            }
            let mut label_element = extended_chart_text(
                Rect {
                    x: rect.x + font,
                    width: (rect.width - font).max(1.0),
                    ..rect
                },
                label.clone(),
                font,
                0x4040_40ff,
            );
            if compact && let Visual::TextLayout { layout, visual } = &mut label_element.visual {
                layout.auto_fit = crate::model::TextAutoFit::None;
                layout.inset_left = 0.0; layout.inset_right = 0.0; layout.inset_top = 0.0; layout.inset_bottom = 0.0;
                layout.vertical_overflow = crate::model::TextVerticalOverflow::Ellipsis;
                if let Visual::RichText { align, .. } = visual.as_mut() { *align = TextAlign::Start; }
            }
            elements.push(label_element);
        }
    }
    let series = chart
        .series
        .iter()
        .find(|s| s.kind == kind)
        .expect("selected chart family is present");
    let color = |s: &ChartSeries, i: usize| {
        s.point_colors
            .get(i)
            .copied()
            .or(s.color)
            .unwrap_or_else(|| super::office_chart_palette_color(i))
    };
    let paint = |s: &ChartSeries, i: usize, rect| {
        s.point_fills
            .get(i)
            .and_then(Option::as_ref)
            .or(s.fill.as_ref())
            .map_or_else(|| Paint::Solid(color(s, i)), |fill| fill.paint(rect))
    };
    if kind == ChartKind::Funnel {
        let maximum = series
            .values
            .iter()
            .copied()
            .filter(|v| v.is_finite())
            .fold(0.0_f32, f32::max);
        let step = plot.height / series.values.len().max(1) as f32;
        for (i, &value) in series
            .values
            .iter()
            .enumerate()
            .filter(|(_, v)| v.is_finite())
        {
            let width = plot.width
                * if maximum > 0.0 {
                    value.max(0.0) / maximum
                } else {
                    0.0
                };
            let height = step / (1.0 + chart.category_gap_width.unwrap_or(0.5));
            let rect = Rect {
                x: plot.x + (plot.width - width) / 2.0,
                y: plot.y + i as f32 * step + (step - height) / 2.0,
                width,
                height,
            };
            let mut bar = extended_chart_shape(
                rect,
                Geometry::Rectangle,
                paint(series, i, rect),
                series
                    .point_border_colors
                    .get(i)
                    .copied()
                    .flatten()
                    .map_or(Paint::None, Paint::Solid),
            );
            if let Visual::PaintedShape { stroke_width, .. } = &mut bar.visual {
                *stroke_width = series.point_border_widths.get(i).copied().unwrap_or(1.0);
            }
            elements.push(bar);
            elements.push(extended_chart_text(
                Rect {
                    x: bounds.x,
                    y: rect.y,
                    width: plot.x - bounds.x - 4.0,
                    height: rect.height,
                },
                series.categories.get(i).cloned().unwrap_or_default(),
                font,
                0x4040_40ff,
            ));
            if series.show_values {
                elements.push(extended_chart_text(
                    rect,
                    format_chart_value(value, series.number_format.as_deref()),
                    font,
                    0xffff_ffff,
                ));
            }
        }
        return Ok(Some(elements));
    }
    if matches!(kind, ChartKind::PieOfPie(_) | ChartKind::BarOfPie(_)) {
        let split = match kind { ChartKind::PieOfPie(n) | ChartKind::BarOfPie(n) => usize::from(n).min(series.values.len()), _ => unreachable!() };
        let explosion = series.point_explosions.iter().copied().fold(0.0_f32, f32::max);
        let extent = 1.0 + explosion;
        let radius = (plot.width / (2.0 * extent + 2.0)).min(plot.height / (2.0 * extent));
        let main_center = (plot.x + radius * extent, plot.y + plot.height / 2.0);
        let secondary_size = (plot.width * 0.28).min(plot.height / 2.5);
        let second_center = (plot.x + plot.width - secondary_size * 0.65, main_center.1);
        let (total, slices) = chart_pie_slices(series, main_center, (radius, radius), 0.0, 0.0, Some(split));
        for slice in &slices {
            elements.push(extended_chart_shape(slice.bounds, slice.geometry.clone(), Paint::Solid(slice.color), Paint::None));
        }
        for label in chart_pie_label_layout(chart, series, &slices, bounds, radius, radius) {
            elements.push(extended_chart_text(label.bounds, label.text, label.style.font_size, label.style.color));
        }
        let start = series.values.len() - split;
        let mut second = series.clone();
        second.kind = ChartKind::Pie;
        second.values = series.values[start..].to_vec();
        second.point_colors = (start..series.values.len()).map(|i| color(series, i)).collect();
        second.point_explosions.clear();
        let secondary_radius = secondary_size * 0.75;
        let secondary_left;
        if matches!(kind, ChartKind::BarOfPie(_)) {
            secondary_left = second_center.0 - secondary_size * 0.35;
            let total = second.values.iter().copied().filter(|v| v.is_finite() && *v > 0.0).sum::<f32>();
            let mut y = second_center.1 - secondary_radius;
            for (index, value) in second.values.iter().copied().enumerate() {
                if !value.is_finite() || value <= 0.0 || total <= 0.0 { continue; }
                let height = secondary_radius * 2.0 * value / total;
                let rect = Rect { x: secondary_left, y, width: secondary_size * 0.7, height };
                elements.push(extended_chart_shape(rect, Geometry::Rectangle, Paint::Solid(color(series, start + index)), Paint::None));
                if series.show_values {
                    elements.push(extended_chart_text(Rect { x: rect.x + rect.width + 2.0, width: font * 2.0, ..rect }, series.value_label(start + index, value, series.number_format.as_deref()), font, 0x0000_00ff));
                }
                y += height;
            }
        } else {
            secondary_left = second_center.0 - secondary_radius;
            let (_, secondary) = chart_pie_slices(&second, second_center, (secondary_radius, secondary_radius), 0.0, 0.0, None);
            for slice in secondary {
                elements.push(extended_chart_shape(slice.bounds, slice.geometry, Paint::Solid(slice.color), Paint::Solid(0xffff_ffff)));
            }
        }
        if let Some(aggregate) = slices.iter().find(|slice| slice.index == series.values.len()) {
            for sign in [-1.0, 1.0] {
                let angle = aggregate.label_angle + sign * std::f32::consts::PI * aggregate.value / total;
                elements.push(extended_chart_shape(bounds, extended_chart_path(&[
                    (aggregate.center.0 + radius * angle.cos() - bounds.x, aggregate.center.1 + radius * angle.sin() - bounds.y),
                    (secondary_left - bounds.x, second_center.1 + sign * secondary_radius - bounds.y),
                ], false), Paint::None, Paint::Solid(0x0000_00ff)));
            }
        }
        return Ok(Some(elements));
    }
    if kind == ChartKind::Radar {
        let (_, maximum, step) = chart.value_axis();
        let side = plot.width.min(plot.height);
        let radar = Rect { x: plot.x + (plot.width - side) / 2.0, y: plot.y + (plot.height - side) / 2.0, width: side, height: side };
        let count = series.categories.len().max(series.values.len());
        let center = (radar.x + side / 2.0, radar.y + side / 2.0);
        let radius = side * 0.42;
        for tick in 1..=((maximum / step).round() as usize).min(100) {
            let value = tick as f32 * step;
            if let Some(geometry) = radar_geometry(&vec![value; count], maximum, radar) {
                elements.push(extended_chart_shape(radar, geometry, Paint::None, Paint::Solid(0x9f9f_9fff)));
            }
        }
        for candidate in chart.series.iter().filter(|s| s.kind == ChartKind::Radar) {
            if let Some(geometry) = radar_geometry(&candidate.values, maximum, radar) {
                let mut element = extended_chart_shape(radar, geometry, candidate.fill.as_ref().map_or(Paint::None, |fill| fill.paint(radar)), Paint::Solid(candidate.color.unwrap_or(0x4f81_bdff)));
                element.visual = candidate.effects.clone().wrap(element.visual, None);
                elements.push(element);
            }
        }
        for (index, label) in series.categories.iter().enumerate() {
            let angle = -std::f32::consts::FRAC_PI_2 + std::f32::consts::TAU * index as f32 / count.max(1) as f32;
            let label_width = (side * 0.65).max(font * 1.5).min(font * 4.0);
            let mut element = extended_chart_text(Rect { x: center.0 + radius * 1.2 * angle.cos() - label_width / 2.0, y: center.1 + radius * 1.2 * angle.sin() - font * 0.6, width: label_width, height: font * 1.25 }, label.clone(), font, 0x0000_00ff);
            if let Visual::TextLayout { layout, visual } = &mut element.visual {
                layout.auto_fit = crate::model::TextAutoFit::None;
                layout.wrap = false;
                if let Visual::RichText { runs, .. } = visual.as_mut() {
                    for run in runs {
                        ellipsize(&mut run.text, font, label_width);
                    }
                }
                layout.inset_left = 0.0; layout.inset_right = 0.0; layout.inset_top = 0.0; layout.inset_bottom = 0.0;
            }
            elements.push(element);
        }
        for tick in 0..=((maximum / step).round() as usize).min(100) {
            let value = tick as f32 * step;
            elements.push(extended_chart_text(Rect { x: center.0 - font * 2.0, y: center.1 - radius * value / maximum - font * 0.6, width: font * 1.8, height: font * 1.2 }, format_chart_value(value, chart.value_axis_options.number_format.as_deref()), font, 0x0000_00ff));
        }
        return Ok(Some(elements));
    }
    if kind == ChartKind::Sunburst {
        let mut levels = if series.category_levels.is_empty() {
            vec![series.categories.clone()]
        } else {
            series.category_levels.clone()
        };
        // Empty ancestor cells continue the preceding branch; an empty leaf is
        // a shorter path and must not inherit the previous sibling's label.
        for row in 0..series.values.len() {
            let mut changed = false;
            for depth in 0..levels.len().saturating_sub(1) {
                levels[depth].resize(series.values.len(), String::new());
                if row > 0 && levels[depth][row].is_empty() && !changed {
                    levels[depth][row] = levels[depth][row - 1].clone();
                } else if row > 0 && levels[depth][row] != levels[depth][row - 1] {
                    changed = true;
                }
            }
        }
        let total = series
            .values
            .iter()
            .copied()
            .filter(|v| v.is_finite() && *v > 0.0)
            .sum::<f32>();
        if total <= 0.0 {
            return Ok(Some(elements));
        }
        let radius = plot.width.min(plot.height) / 2.0;
        let center = (plot.x + plot.width / 2.0, plot.y + plot.height / 2.0);
        let ring = radius / levels.len().max(1) as f32;
        let mut root_color_indices = Vec::with_capacity(series.values.len());
        let mut root_index = 0;
        for i in 0..series.values.len() {
            if i > 0 && levels[0].get(i) != levels[0].get(i - 1) {
                root_index += 1;
            }
            root_color_indices.push(root_index);
        }
        for depth in 0..levels.len() {
            let mut start = 0;
            let mut angle = -std::f32::consts::FRAC_PI_2;
            while start < series.values.len() {
                let label = levels[depth].get(start).cloned().unwrap_or_default();
                let mut end = start + 1;
                while end < series.values.len()
                    && (0..=depth).all(|d| levels[d].get(end) == levels[d].get(start))
                {
                    end += 1;
                }
                let sum = series.values[start..end]
                    .iter()
                    .copied()
                    .filter(|v| v.is_finite() && *v > 0.0)
                    .sum::<f32>();
                let sweep = std::f32::consts::TAU * sum / total;
                if sum > 0.0 && !label.is_empty() {
                    let leaf = (depth + 1..levels.len()).all(|d| {
                        (start..end).all(|i| levels[d].get(i).is_none_or(String::is_empty))
                    });
                    let outer = if leaf {
                        radius
                    } else {
                        ring * (depth + 1) as f32
                    };
                    let inner = ring * depth as f32;
                    let steps = ((sweep * outer / 3.0).ceil() as usize).clamp(2, 256);
                    let mut points = Vec::with_capacity(2 * (steps + 1));
                    for i in 0..=steps {
                        let a = angle + sweep * i as f32 / steps as f32;
                        points.push((radius + outer * a.cos(), radius + outer * a.sin()));
                    }
                    for i in (0..=steps).rev() {
                        let a = angle + sweep * i as f32 / steps as f32;
                        points.push((radius + inner * a.cos(), radius + inner * a.sin()));
                    }
                    let rect = Rect {
                        x: center.0 - radius,
                        y: center.1 - radius,
                        width: radius * 2.0,
                        height: radius * 2.0,
                    };
                    elements.push(extended_chart_shape(
                        rect,
                        extended_chart_path(&points, true),
                        {
                            let root = root_color_indices[start];
                            let mut fill = color(series, root);
                            if depth > 0 {
                                apply_color_transform(
                                    &mut fill,
                                    "tint",
                                    (1.0 - depth as f32 * 0.12).max(0.4),
                                );
                            }
                            Paint::Solid(fill)
                        },
                        Paint::Solid(0xffff_ffff),
                    ));
                    let a = angle + sweep / 2.0;
                    let r = (inner + outer) / 2.0;
                    if sweep * r > 4.0 {
                        let width = (outer - inner).max(ring);
                        let mut text = extended_chart_text(
                            Rect {
                                x: center.0 + r * a.cos() - width / 2.0,
                                y: center.1 + r * a.sin() - font * 0.6,
                                width,
                                height: (sweep * r * 0.8).min(font * 1.2),
                            },
                            label,
                            font,
                            0xffff_ffff,
                        );
                        if let Visual::TextLayout { layout, .. } = &mut text.visual {
                            layout.wrap = false;
                        }
                        elements.push(text);
                    }
                }
                angle += sweep;
                start = end;
            }
        }
        return Ok(Some(elements));
    }
    let surface = matches!(kind, ChartKind::Surface | ChartKind::SurfaceWireframe);
    let (minimum, maximum, major) = if surface {
        chart_surface_axis(chart)
    } else {
        chart.value_axis()
    };
    let range = (maximum - minimum).max(f32::EPSILON);
    let y = |v: f32| plot.height * (1.0 - (v - minimum) / range);
    let ticks = if surface {
        (0..=((maximum - minimum) / major).floor().clamp(0.0, 100.0) as usize)
            .map(|i| (minimum + i as f32 * major, i as f32 * major / range))
            .collect()
    } else {
        chart.value_axis_ticks()
    };
    let projected = series.three_d
        && matches!(
            kind,
            ChartKind::Surface | ChartKind::SurfaceWireframe | ChartKind::Line3D
        );
    for (value, ratio) in ticks.iter().copied().filter(|_| !projected && !surface) {
        let yy = plot.height * (1.0 - ratio);
        elements.push(extended_chart_shape(
            plot,
            extended_chart_path(&[(0.0, yy), (plot.width, yy)], false),
            Paint::None,
            Paint::Solid(0xd9d9_d9ff),
        ));
        elements.push(extended_chart_text(
            Rect {
                x: bounds.x,
                y: plot.y + yy - font,
                width: plot.x - bounds.x - 4.0,
                height: font * 2.0,
            },
            chart.value_axis_options.format_label(value),
            font,
            0x4040_40ff,
        ));
    }
    let count = chart
        .series
        .iter()
        .map(|s| s.values.len())
        .max()
        .unwrap_or(0);
    let step = plot.width / count.max(1) as f32;
    for i in (0..count).filter(|_| !projected && !chart.horizontal_axis_options.deleted) {
        elements.push(extended_chart_text(
            Rect {
                x: plot.x + i as f32 * step,
                y: plot.y + plot.height,
                width: step,
                height: font * 2.5,
            },
            series
                .categories
                .get(i)
                .cloned()
                .unwrap_or_else(|| (i + 1).to_string()),
            font,
            0x4040_40ff,
        ));
    }
    match kind {
        ChartKind::Histogram | ChartKind::Pareto => {
            let total = series.values.iter().sum::<f32>();
            let mut cumulative = 0.0;
            let mut points = Vec::new();
            for (i, &value) in series.values.iter().enumerate() {
                let yy = y(value);
                let rect = Rect {
                    x: plot.x + i as f32 * step,
                    y: plot.y + yy,
                    width: step,
                    height: (plot.height - yy).max(0.0),
                };
                elements.push(extended_chart_shape(
                    rect,
                    Geometry::Rectangle,
                    series
                        .fill
                        .as_ref()
                        .map_or(Paint::Solid(series.color.unwrap_or(0x4472_c4ff)), |f| {
                            f.paint(rect)
                        }),
                    Paint::Solid(0xffff_ffff),
                ));
                cumulative += value;
                points.push((
                    step * (i as f32 + 0.5),
                    plot.height * (1.0 - cumulative / total.max(1.0)),
                ));
            }
            if kind == ChartKind::Pareto {
                elements.push(extended_chart_shape(
                    plot,
                    extended_chart_path(&points, false),
                    Paint::None,
                    Paint::Solid(0xed7d_31ff),
                ));
                for i in 0..=5 {
                    elements.push(extended_chart_text(
                        Rect {
                            x: plot.x + plot.width,
                            y: plot.y + plot.height * (1.0 - i as f32 / 5.0) - font,
                            width: bounds.width * 0.08,
                            height: font * 2.0,
                        },
                        format!("{}%", i * 20),
                        font,
                        0x4040_40ff,
                    ));
                }
            }
        }
        ChartKind::Stock => {
            let stocks = chart
                .series
                .iter()
                .filter(|s| s.kind == ChartKind::Stock)
                .collect::<Vec<_>>();
            for volume in chart.series.iter().filter(|s| s.kind == ChartKind::Bar) {
                let (min, max, _) = chart.value_axis_for_series(volume);
                let yy = |v: f32| plot.height * (1.0 - (v - min) / (max - min).max(f32::EPSILON));
                for (i, &value) in volume
                    .values
                    .iter()
                    .enumerate()
                    .filter(|(_, v)| v.is_finite())
                {
                    let top = yy(value).min(yy(0.0));
                    let rect = Rect {
                        x: plot.x + step * (i as f32 + 0.3),
                        y: plot.y + top,
                        width: step * 0.4,
                        height: (yy(value) - yy(0.0)).abs(),
                    };
                    elements.push(extended_chart_shape(
                        rect,
                        Geometry::Rectangle,
                        paint(volume, i, rect),
                        Paint::None,
                    ));
                }
            }
            if let Some((axis, ticks)) = chart.secondary_value_axis_ticks() {
                for (value, ratio) in ticks {
                    elements.push(extended_chart_text(
                        Rect {
                            x: plot.x + plot.width,
                            y: plot.y + plot.height * (1.0 - ratio) - font,
                            width: bounds.width * 0.08,
                            height: font * 2.0,
                        },
                        axis.format_label(value),
                        font,
                        0x4040_40ff,
                    ));
                }
            }
            if stocks.len() >= 3 {
                let (min, max, _) = chart.value_axis_for_series(stocks[0]);
                let y = |v: f32| plot.height * (1.0 - (v - min) / (max - min).max(f32::EPSILON));
                let high = stocks[stocks.len() - 3];
                let low = stocks[stocks.len() - 2];
                let close = stocks[stocks.len() - 1];
                let open = (stocks.len() >= 4).then(|| stocks[stocks.len() - 4]);
                for i in 0..count {
                    let (Some(&h), Some(&l), Some(&c)) =
                        (high.values.get(i), low.values.get(i), close.values.get(i))
                    else {
                        continue;
                    };
                    if !h.is_finite() || !l.is_finite() || !c.is_finite() {
                        continue;
                    }
                    let x = step * (i as f32 + 0.5);
                    elements.push(extended_chart_shape(
                        plot,
                        extended_chart_path(&[(x, y(l)), (x, y(h))], false),
                        Paint::None,
                        Paint::Solid(0x4040_40ff),
                    ));
                    if let Some(o) = open.and_then(|s| s.values.get(i)).copied() {
                        let rect = Rect {
                            x: plot.x + x - step * 0.2,
                            y: plot.y + y(o.max(c)),
                            width: step * 0.4,
                            height: (y(o.min(c)) - y(o.max(c))).max(1.0),
                        };
                        elements.push(extended_chart_shape(
                            rect,
                            Geometry::Rectangle,
                            Paint::Solid(if c >= o { 0xffff_ffff } else { 0x0000_00ff }),
                            Paint::Solid(0x4040_40ff),
                        ));
                    } else {
                        elements.push(extended_chart_shape(
                            plot,
                            extended_chart_path(&[(x, y(c)), (x + step * 0.2, y(c))], false),
                            Paint::None,
                            Paint::Solid(0x4040_40ff),
                        ));
                    }
                }
            }
        }
        ChartKind::Surface | ChartKind::SurfaceWireframe | ChartKind::Line3D => {
            let rows = chart
                .series
                .iter()
                .filter(|s| s.kind == kind)
                .collect::<Vec<_>>();
            let view = chart.view_3d.unwrap_or_default();
            let rot_x = (view.rot_x as f32).to_radians();
            let rot_y = (view.rot_y as f32).to_radians();
            let depth_scale = view.depth_percent.unwrap_or(100).clamp(20, 2000) as f32 / 100.0;
            let raw_project = |x: f32, z: f32, height: f32| {
                let x = x - 0.5;
                let z = (z - 0.5) * depth_scale;
                let horizontal = x * rot_y.cos() + z * rot_y.sin();
                let depth = z * rot_y.cos() - x * rot_y.sin();
                let vertical = depth * rot_x.sin() - height * rot_x.cos();
                let perspective = if view.right_angle_axes {
                    1.0
                } else {
                    1.0 / (1.0
                        + (depth * rot_x.cos() + height * rot_x.sin())
                            * view.perspective.min(240) as f32
                            / 720.0)
                        .max(0.2)
                };
                (horizontal * perspective, vertical * perspective)
            };
            let mut extent = (
                f32::INFINITY,
                f32::INFINITY,
                f32::NEG_INFINITY,
                f32::NEG_INFINITY,
            );
            for x in [0.0, 1.0] {
                for z in [0.0, 1.0] {
                    for h in [0.0, 1.0] {
                        let p = raw_project(x, z, h);
                        extent = (
                            extent.0.min(p.0),
                            extent.1.min(p.1),
                            extent.2.max(p.0),
                            extent.3.max(p.1),
                        );
                    }
                }
            }
            let project = |row: f32, col: f32, value: f32| {
                let z = row / rows.len().saturating_sub(1).max(1) as f32;
                let x = col / count.saturating_sub(1).max(1) as f32;
                if !series.three_d {
                    (x * plot.width, (1.0 - z) * plot.height)
                } else {
                    let p = raw_project(x, z, (value - minimum) / range);
                    (
                        (p.0 - extent.0) / (extent.2 - extent.0).max(f32::EPSILON) * plot.width,
                        (p.1 - extent.1) / (extent.3 - extent.1).max(f32::EPSILON) * plot.height,
                    )
                }
            };
            if projected {
                let last_row = rows.len().saturating_sub(1) as f32;
                let last_col = count.saturating_sub(1) as f32;
                let far_row = if rot_y.cos() < 0.0 { last_row } else { 0.0 };
                let right = chart.value_axis_options.position.as_deref() == Some("r");
                let axis_col = if (project(far_row, 0.0, minimum).0
                    > project(far_row, last_col, minimum).0)
                    == right
                {
                    0.0
                } else {
                    last_col
                };
                for &(value, _) in &ticks {
                    let a = project(far_row, 0.0, value);
                    let b = project(far_row, last_col, value);
                    elements.push(extended_chart_shape(
                        plot,
                        extended_chart_path(&[a, b], false),
                        Paint::None,
                        Paint::Solid(0xd9d9_d9ff),
                    ));
                    let p = project(far_row, axis_col, value);
                    elements.push(extended_chart_text(
                        Rect {
                            x: plot.x + p.0 + if right { 2.0 } else { -font * 2.0 },
                            y: plot.y + p.1 - font,
                            width: font * 2.0,
                            height: font * 2.0,
                        },
                        chart.value_axis_options.format_label(value),
                        font,
                        0x4040_40ff,
                    ));
                }
                elements.push(extended_chart_shape(
                    plot,
                    extended_chart_path(
                        &[
                            project(0.0, 0.0, minimum),
                            project(0.0, last_col, minimum),
                            project(last_row, last_col, minimum),
                            project(last_row, 0.0, minimum),
                        ],
                        true,
                    ),
                    Paint::None,
                    Paint::Solid(0x8080_80ff),
                ));
                if !chart.horizontal_axis_options.deleted {
                    for c in 0..count {
                        let p = project(
                            if far_row == 0.0 { last_row } else { 0.0 },
                            c as f32,
                            minimum,
                        );
                        elements.push(extended_chart_text(
                            Rect {
                                x: plot.x + p.0 - step / 2.0,
                                y: plot.y + p.1,
                                width: step,
                                height: font * 2.0,
                            },
                            series.categories.get(c).cloned().unwrap_or_default(),
                            font,
                            0x4040_40ff,
                        ));
                    }
                }
                if chart
                    .series_axis_options
                    .as_ref()
                    .is_some_and(|axis| !axis.deleted)
                {
                    let label_col =
                        if project(0.0, 0.0, minimum).0 < project(0.0, last_col, minimum).0 {
                            0.0
                        } else {
                            last_col
                        };
                    for (r, row) in rows.iter().enumerate() {
                        let p = project(r as f32, label_col, minimum);
                        elements.push(extended_chart_text(
                            Rect {
                                x: plot.x + p.0 - font * 4.0,
                                y: plot.y + p.1 - font,
                                width: font * 3.8,
                                height: font * 2.0,
                            },
                            row.name.clone(),
                            font,
                            0x4040_40ff,
                        ));
                    }
                }
            }
            if surface && !series.three_d {
                for (r, row) in rows.iter().enumerate() {
                    let p = project(r as f32, 0.0, minimum);
                    elements.push(extended_chart_shape(
                        plot,
                        extended_chart_path(&[(0.0, p.1), (plot.width, p.1)], false),
                        Paint::None,
                        Paint::Solid(0xd9d9_d9ff),
                    ));
                    elements.push(extended_chart_text(
                        Rect {
                            x: bounds.x,
                            y: plot.y + p.1 - font,
                            width: plot.x - bounds.x - 4.0,
                            height: font * 2.0,
                        },
                        row.name.clone(),
                        font,
                        0x4040_40ff,
                    ));
                }
            }
            let mut row_order = (0..rows.len()).collect::<Vec<_>>();
            if rot_y.cos() < 0.0 {
                row_order.reverse();
            }
            for r in row_order {
                let mut column_order =
                    (0..rows[r].values.len().saturating_sub(1)).collect::<Vec<_>>();
                if rot_y.sin() > 0.0 {
                    column_order.reverse();
                }
                for c in column_order {
                    let (a, b) = (rows[r].values[c], rows[r].values[c + 1]);
                    if !a.is_finite() || !b.is_finite() {
                        continue;
                    }
                    if kind == ChartKind::Line3D {
                        elements.push(extended_chart_shape(
                            plot,
                            extended_chart_path(
                                &[
                                    project(r as f32, c as f32, a),
                                    project(r as f32, c as f32 + 1.0, b),
                                ],
                                false,
                            ),
                            Paint::None,
                            Paint::Solid(color(rows[r], r)),
                        ));
                    } else if r + 1 < rows.len() && c + 1 < rows[r + 1].values.len() {
                        let (d, e) = (rows[r + 1].values[c], rows[r + 1].values[c + 1]);
                        if !d.is_finite() || !e.is_finite() {
                            continue;
                        }
                        if kind == ChartKind::SurfaceWireframe {
                            elements.push(extended_chart_shape(
                                plot,
                                extended_chart_path(
                                    &[
                                        project(r as f32, c as f32, a),
                                        project(r as f32, c as f32 + 1.0, b),
                                        project(r as f32 + 1.0, c as f32 + 1.0, e),
                                        project(r as f32 + 1.0, c as f32, d),
                                    ],
                                    true,
                                ),
                                Paint::None,
                                Paint::Solid(color(rows[r], r)),
                            ));
                            continue;
                        }
                        // Split each cell into triangles and clip at actual value-axis
                        // bands; an average cell color loses peaks and valleys.
                        for triangle in [
                            vec![
                                (c as f32, r as f32, a),
                                (c as f32 + 1.0, r as f32, b),
                                (c as f32 + 1.0, r as f32 + 1.0, e),
                            ],
                            vec![
                                (c as f32, r as f32, a),
                                (c as f32 + 1.0, r as f32 + 1.0, e),
                                (c as f32, r as f32 + 1.0, d),
                            ],
                        ] {
                            let bands =
                                (range / major.max(f32::EPSILON)).ceil().clamp(1.0, 100.0) as usize;
                            for band in 0..bands {
                                let lower = minimum + band as f32 * major;
                                let upper = lower + major;
                                let mut polygon = triangle.clone();
                                for (edge, above) in [(lower, true), (upper, false)] {
                                    let input = std::mem::take(&mut polygon);
                                    for index in 0..input.len() {
                                        let a = input[index];
                                        let b = input[(index + 1) % input.len()];
                                        let inside =
                                            |v: f32| if above { v >= edge } else { v <= edge };
                                        if inside(a.2) {
                                            polygon.push(a);
                                        }
                                        if inside(a.2) != inside(b.2) {
                                            let t = (edge - a.2) / (b.2 - a.2);
                                            polygon.push((
                                                a.0 + t * (b.0 - a.0),
                                                a.1 + t * (b.1 - a.1),
                                                edge,
                                            ));
                                        }
                                    }
                                }
                                if polygon.len() < 3 {
                                    continue;
                                }
                                let points = polygon
                                    .into_iter()
                                    .map(|(col, row, value)| project(row, col, value))
                                    .collect::<Vec<_>>();
                                elements.push(extended_chart_shape(
                                    plot,
                                    extended_chart_path(&points, true),
                                    chart
                                        .surface_band_fills
                                        .get(band)
                                        .and_then(Option::as_ref)
                                        .map_or_else(
                                            || Paint::Solid(color(rows[band % rows.len()], band)),
                                            |fill| fill.paint(plot),
                                        ),
                                    Paint::None,
                                ));
                            }
                        }
                    }
                }
            }
        }
        _ => {}
    }
    Ok(Some(elements))
}
