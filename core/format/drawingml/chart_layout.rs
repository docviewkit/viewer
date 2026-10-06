use super::nice_chart_step;

pub(super) fn chart_marker_geometry(symbol: &str, size: f32) -> Option<(Geometry, bool)> {
    Some(match symbol {
        "square" => (Geometry::Rectangle, true),
        "diamond" => (
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo {
                        x: size / 2.0,
                        y: 0.0,
                    },
                    PathCommand::LineTo {
                        x: size,
                        y: size / 2.0,
                    },
                    PathCommand::LineTo {
                        x: size / 2.0,
                        y: size,
                    },
                    PathCommand::LineTo {
                        x: 0.0,
                        y: size / 2.0,
                    },
                    PathCommand::ClosePath,
                ],
            },
            true,
        ),
        "circle" => (Geometry::Ellipse, true),
        "triangle" => (
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo {
                        x: size / 2.0,
                        y: 0.0,
                    },
                    PathCommand::LineTo { x: size, y: size },
                    PathCommand::LineTo { x: 0.0, y: size },
                    PathCommand::ClosePath,
                ],
            },
            true,
        ),
        "x" => (
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo { x: 0.0, y: 0.0 },
                    PathCommand::LineTo { x: size, y: size },
                    PathCommand::MoveTo { x: size, y: 0.0 },
                    PathCommand::LineTo { x: 0.0, y: size },
                ],
            },
            false,
        ),
        _ => return None,
    })
}

impl ChartErrorBars {
    fn resolve(&mut self, values: &[f32]) {
        if self.standard_error {
            // Excel standard error: sample standard deviation / sqrt(n).
            let count = values.iter().filter(|value| value.is_finite()).count() as f64;
            if count > 1.0 {
                let mean = values
                    .iter()
                    .filter(|value| value.is_finite())
                    .map(|value| f64::from(*value))
                    .sum::<f64>()
                    / count;
                let error = (values
                    .iter()
                    .filter(|value| value.is_finite())
                    .map(|value| (f64::from(*value) - mean).powi(2))
                    .sum::<f64>()
                    / (count * (count - 1.0)))
                    .sqrt() as f32;
                self.plus = vec![error; values.len()];
                self.minus = vec![error; values.len()];
            }
        }
        if !self.show_plus {
            self.plus.clear();
        }
        if !self.show_minus {
            self.minus.clear();
        }
    }

    fn amounts(&self, index: usize) -> (f32, f32) {
        let amount = |values: &[f32]| {
            values
                .get(index)
                .copied()
                .filter(|value| value.is_finite() && *value > 0.0)
                .unwrap_or(0.0)
        };
        (amount(&self.minus), amount(&self.plus))
    }
}

impl ChartSeries {
    pub(super) fn projected_column(&self) -> bool {
        self.kind == ChartKind::Bar && !self.bar_horizontal
            && (self.bar_depth || (self.three_d && self.grouping == ChartGrouping::Standard
                && !self.bar_cone && !self.bar_cylinder))
    }

    pub(super) fn value_label(&self, index: usize, value: f32, format: Option<&str>) -> String {
        if format.is_some_and(|format| format.eq_ignore_ascii_case("General"))
            && let Some(raw) = self.raw_values.get(index).and_then(|raw| raw.parse::<f64>().ok())
            && raw.is_finite() && raw as f32 == value {
            return super::format_general_number(raw, 11);
        }
        format_chart_value(value, format)
    }

    pub(super) fn pie_start_angle_radians(&self) -> f32 {
        (self.first_slice_angle - 90.0).to_radians()
    }

    pub(super) fn pie_percentages(&self) -> Vec<u32> {
        let total = self
            .values
            .iter()
            .copied()
            .filter(|value| value.is_finite() && *value > 0.0)
            .sum::<f32>();
        if total <= 0.0 {
            return vec![0; self.values.len()];
        }
        let mut percentages = self
            .values
            .iter()
            .map(|value| ((*value).max(0.0) / total * 100.0).floor() as u32)
            .collect::<Vec<_>>();
        let mut remainders = self
            .values
            .iter()
            .enumerate()
            .map(|(index, value)| {
                let exact = (*value).max(0.0) / total * 100.0;
                (index, exact - exact.floor())
            })
            .collect::<Vec<_>>();
        remainders.sort_by(|left, right| {
            right
                .1
                .total_cmp(&left.1)
                .then_with(|| left.0.cmp(&right.0))
        });
        for (index, _) in remainders
            .into_iter()
            .take(100_usize.saturating_sub(percentages.iter().sum::<u32>() as usize))
        {
            percentages[index] += 1;
        }
        percentages
    }

    pub(super) fn value_spans(&self) -> Vec<(f32, f32)> {
        if self.kind != ChartKind::Waterfall {
            return self.values.iter().map(|value| (0.0, *value)).collect();
        }
        let mut total = 0.0_f32;
        self.values
            .iter()
            .enumerate()
            .map(|(index, value)| {
                if self.subtotals.binary_search(&index).is_ok() {
                    total = *value;
                    (0.0, *value)
                } else {
                    let start = total;
                    total += *value;
                    (start, total)
                }
            })
            .collect()
    }
}

#[derive(Clone, Debug)]
pub(super) struct ChartPieSlice {
    pub(super) index: usize,
    pub(super) bounds: Rect,
    pub(super) geometry: Geometry,
    pub(super) color: u32,
    pub(super) label_angle: f32,
    pub(super) value: f32,
    pub(super) center: (f32, f32),
    pub(super) side: Option<(Rect, Geometry)>,
    pub(super) cut_side: Option<(Rect, Geometry)>,
}

pub(super) struct ChartPieLabelLayout {
    pub(super) index: usize,
    pub(super) text: String,
    pub(super) style: ChartTextStyle,
    pub(super) bounds: Rect,
    pub(super) leader: Option<(Rect, Geometry)>,
}

fn chart_pie_label_anchor(slice: &ChartPieSlice, radius: f32, vertical_radius: f32) -> (f32, f32) {
    let depth = slice
        .side
        .as_ref()
        .filter(|_| slice.label_angle.sin() > 0.0)
        .map_or(0.0, |(side, _)| {
            (side.y + side.height - slice.bounds.y - slice.bounds.height).max(0.0)
        });
    (
        slice.center.0 + radius * slice.label_angle.cos(),
        slice.center.1 + vertical_radius * slice.label_angle.sin() + depth,
    )
}

fn chart_pie_label_leader(
    slice: &ChartPieSlice,
    label: &ChartPieLabelLayout,
    radius: f32,
    vertical_radius: f32,
) -> Option<(Rect, Geometry)> {
    let edge = chart_pie_label_anchor(slice, radius, vertical_radius);
    let label_edge = (
        edge.0
            .clamp(label.bounds.x, label.bounds.x + label.bounds.width),
        edge.1
            .clamp(label.bounds.y, label.bounds.y + label.bounds.height),
    );
    // Adjacent labels identify their slices without a line across the face.
    if (label_edge.0 - edge.0).hypot(label_edge.1 - edge.1) <= label.style.font_size * 0.8 {
        return None;
    }
    let line_bounds = Rect {
        x: edge.0.min(label_edge.0),
        y: edge.1.min(label_edge.1),
        width: (label_edge.0 - edge.0).abs().max(0.01),
        height: (label_edge.1 - edge.1).abs().max(0.01),
    };
    Some((
        line_bounds,
        Geometry::Path {
            fill_rule: FillRule::NonZero,
            commands: vec![
                PathCommand::MoveTo {
                    x: edge.0 - line_bounds.x,
                    y: edge.1 - line_bounds.y,
                },
                PathCommand::LineTo {
                    x: label_edge.0 - line_bounds.x,
                    y: label_edge.1 - line_bounds.y,
                },
            ],
        },
    ))
}

pub(super) fn chart_pie_label_layout(
    chart: &Chart,
    series: &ChartSeries,
    slices: &[ChartPieSlice],
    bounds: Rect,
    radius: f32,
    vertical_radius: f32,
) -> Vec<ChartPieLabelLayout> {
    let total = slices.iter().map(|slice| slice.value).sum::<f32>();
    let percentages = series.pie_percentages();
    let mut outside_indices = Vec::new();
    let mut labels = slices
        .iter()
        .filter_map(|slice| {
            let point_label = series
                .data_labels
                .iter()
                .find(|label| label.index == slice.index);
            let show_values = point_label
                .and_then(|label| label.show_values)
                .unwrap_or(series.show_values);
            let show_category_name = point_label
                .and_then(|label| label.show_category_name)
                .unwrap_or(series.show_category_name);
            let show_percent = point_label
                .and_then(|label| label.show_percent)
                .unwrap_or(series.show_percent);
            if !(show_values || show_category_name || show_percent)
                || series.hidden_labels.contains(&slice.index)
            {
                return None;
            }
            let category = series
                .categories
                .get(slice.index)
                .cloned()
                .unwrap_or_default();
            let mut lines = Vec::new();
            if show_category_name && !category.is_empty() {
                lines.push(category.clone());
            }
            if show_values {
                lines.push(
                    series.value_label(
                        slice.index,
                        slice.value,
                        point_label
                            .and_then(|label| label.number_format.as_deref())
                            .filter(|format| !(show_percent && format.contains('%')))
                            .or(series.number_format.as_deref()),
                    ),
                );
            }
            if show_percent {
                let format = point_label
                    .and_then(|label| label.number_format.as_deref())
                    .or(series.number_format.as_deref());
                lines.push(if format.is_some_and(|f| f.contains('%')) {
                    format_chart_value(slice.value / total, format)
                } else {
                    format!(
                        "{}%",
                        percentages.get(slice.index).copied().unwrap_or_default()
                    )
                });
            }
            let text = point_label
                .filter(|label| !label.text.is_empty())
                .map_or_else(|| lines.join("\n"), |label| label.text.clone());
            let style = chart.data_label_text_style(series, point_label);
            let estimated_label_width = text
                .lines()
                .map(|line| {
                    line.chars()
                        .map(|character| {
                            drawingml_fallback_character_width(character, style.font_size)
                        })
                        .sum::<f32>()
                })
                .fold(0.0, f32::max)
                + style.font_size;
            let label_width = point_label
                .and_then(|label| label.manual_size.map(|(width, _)| bounds.width * width))
                .unwrap_or(estimated_label_width.clamp(32.0, (bounds.width - 12.0).max(32.0)));
            let label_height = point_label
                .and_then(|label| label.manual_size.map(|(_, height)| bounds.height * height))
                .unwrap_or(style.font_size * (text.lines().count().max(1) as f32 * 1.2 + 0.2));
            let position = point_label
                .and_then(|label| label.position.as_deref())
                .or(series.data_label_position.as_deref());
            let sweep = std::f32::consts::TAU * slice.value / total.max(f32::MIN_POSITIVE);
            let inside_width = 2.0 * radius * 0.56 * (sweep / 2.0).sin();
            let outside = position == Some("outEnd")
                || (position == Some("bestFit") && inside_width < estimated_label_width);
            if outside {
                outside_indices.push(slice.index);
            }
            let distance = match position {
                Some("inBase") => 0.28,
                Some("ctr") => 0.56,
                _ => 0.82,
            };
            let label_center = (
                slice.center.0 + radius * distance * slice.label_angle.cos(),
                slice.center.1 + vertical_radius * distance * slice.label_angle.sin(),
            );
            let align = if !outside || slice.label_angle.cos().abs() < 0.3 {
                TextAlign::Center
            } else if slice.label_angle.cos() < 0.0 {
                TextAlign::End
            } else {
                TextAlign::Start
            };
            let mut label_bounds = if outside {
                let edge = chart_pie_label_anchor(slice, radius, vertical_radius);
                let gap = style.font_size * 0.35;
                Rect {
                    x: match align {
                        TextAlign::Center => edge.0 - label_width / 2.0,
                        TextAlign::End => edge.0 - gap - label_width,
                        _ => edge.0 + gap,
                    },
                    y: if slice.label_angle.sin() < -0.5 {
                        edge.1 - gap - label_height
                    } else if slice.label_angle.sin() > 0.5 {
                        edge.1 + gap
                    } else {
                        edge.1 - label_height / 2.0
                    },
                    width: label_width,
                    height: label_height,
                }
            } else {
                Rect {
                    x: label_center.0 - label_width / 2.0,
                    y: label_center.1 - label_height / 2.0,
                    width: label_width,
                    height: label_height,
                }
            };
            if position != Some("bestFit")
                && let Some((x, y)) = point_label.and_then(|label| label.manual_offset)
            {
                label_bounds.x += bounds.width * x;
                label_bounds.y += bounds.height * y;
            }
            label_bounds.x = label_bounds.x.clamp(
                bounds.x + 6.0,
                bounds.x + bounds.width - label_bounds.width - 6.0,
            );
            label_bounds.y = label_bounds.y.clamp(
                bounds.y + bounds.height * if chart.show_title { 0.18 } else { 0.02 },
                bounds.y + bounds.height - label_bounds.height - 6.0,
            );
            Some(ChartPieLabelLayout {
                index: slice.index,
                text,
                style: ChartTextStyle { align, ..style },
                bounds: label_bounds,
                leader: None,
            })
        })
        .collect::<Vec<_>>();
    let minimum_y = bounds.y + bounds.height * if chart.show_title { 0.18 } else { 0.02 };
    let maximum_y = bounds.y + bounds.height - 6.0;
    for left_side in [true, false] {
        let mut lane = labels
            .iter()
            .enumerate()
            .filter(|(_, label)| {
                outside_indices.contains(&label.index)
                    && (label.bounds.x + label.bounds.width / 2.0 < bounds.x + bounds.width / 2.0)
                        == left_side
            })
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        lane.sort_by(|left, right| labels[*left].bounds.y.total_cmp(&labels[*right].bounds.y));
        let mut cursor = minimum_y;
        for index in &lane {
            labels[*index].bounds.y = labels[*index].bounds.y.max(cursor);
            cursor = labels[*index].bounds.y + labels[*index].bounds.height + 2.0;
        }
        let mut cursor = maximum_y;
        for index in lane.iter().rev() {
            labels[*index].bounds.y = labels[*index]
                .bounds
                .y
                .min(cursor - labels[*index].bounds.height)
                .max(minimum_y);
            cursor = labels[*index].bounds.y - 2.0;
        }
    }
    for label in &mut labels {
        if outside_indices.contains(&label.index)
            && let Some(slice) = slices.iter().find(|slice| slice.index == label.index)
        {
            label.leader = chart_pie_label_leader(slice, label, radius, vertical_radius);
        }
    }
    labels
}

#[derive(Clone, Debug)]
pub(super) struct ChartBoxWhisker {
    pub(super) bounds: Rect,
    pub(super) median_y: f32,
    pub(super) lower_whisker_y: f32,
    pub(super) upper_whisker_y: f32,
    pub(super) mean_y: f32,
    pub(super) outlier_y: Vec<f32>,
}

pub(super) fn chart_box_whisker(
    values: &[f32],
    index: usize,
    count: usize,
    plot: Rect,
    axis: (f32, f32),
    gap_width: f32,
) -> Option<ChartBoxWhisker> {
    let mut sorted = values
        .iter()
        .copied()
        .filter(|value| value.is_finite())
        .collect::<Vec<_>>();
    sorted.sort_by(f32::total_cmp);
    let (minimum, maximum) = axis;
    if sorted.is_empty() || maximum <= minimum || index >= count || count == 0 {
        return None;
    }
    let percentile = |ratio: f32| {
        let rank = ratio * (sorted.len() as f32 + 1.0) - 1.0;
        let lower = rank.floor().max(0.0) as usize;
        let upper = rank.ceil().max(0.0) as usize;
        let lower = lower.min(sorted.len() - 1);
        let upper = upper.min(sorted.len() - 1);
        sorted[lower] + (sorted[upper] - sorted[lower]) * rank.fract().max(0.0)
    };
    let lower_quartile = percentile(0.25);
    let median = percentile(0.5);
    let upper_quartile = percentile(0.75);
    let interquartile_range = upper_quartile - lower_quartile;
    let lower_fence = lower_quartile - interquartile_range * 1.5;
    let upper_fence = upper_quartile + interquartile_range * 1.5;
    let lower_whisker = sorted
        .iter()
        .copied()
        .find(|value| *value >= lower_fence)
        .unwrap_or(sorted[0]);
    let upper_whisker = sorted
        .iter()
        .copied()
        .rev()
        .find(|value| *value <= upper_fence)
        .unwrap_or(*sorted.last()?);
    let y = |value: f32| plot.y + plot.height * (1.0 - (value - minimum) / (maximum - minimum));
    let slot = plot.width / count as f32;
    let width = (slot / (1.0 + gap_width.max(0.0))).clamp(slot * 0.12, slot * 0.9);
    let center_x = plot.x + slot * (index as f32 + 0.5);
    let upper_y = y(upper_quartile);
    let lower_y = y(lower_quartile);
    Some(ChartBoxWhisker {
        bounds: Rect {
            x: center_x - width / 2.0,
            y: upper_y,
            width,
            height: (lower_y - upper_y).max(0.01),
        },
        median_y: y(median),
        lower_whisker_y: y(lower_whisker),
        upper_whisker_y: y(upper_whisker),
        mean_y: y(sorted.iter().sum::<f32>() / sorted.len() as f32),
        outlier_y: sorted
            .iter()
            .copied()
            .filter(|value| *value < lower_fence || *value > upper_fence)
            .map(y)
            .collect(),
    })
}

pub(super) fn chart_pie_slices(
    series: &ChartSeries,
    center: (f32, f32),
    radii: (f32, f32),
    depth: f32,
    depth_perspective: f32,
    bar_of_pie_split: Option<usize>,
) -> (f32, Vec<ChartPieSlice>) {
    let total = series
        .values
        .iter()
        .copied()
        .filter(|value| value.is_finite() && *value > 0.0)
        .sum::<f32>();
    if total <= 0.0 || radii.0 <= 0.0 || radii.1 <= 0.0 {
        return (total, Vec::new());
    }
    let doughnut = series.kind == ChartKind::Doughnut;
    let mut angle = series.pie_start_angle_radians();
    let sweep_direction = 1.0;
    let mut rendered_values = series
        .values
        .iter()
        .copied()
        .enumerate()
        .collect::<Vec<_>>();
    if let Some(split) = bar_of_pie_split.filter(|split| *split > 0) {
        let primary_count = series.values.len() - split.min(series.values.len());
        let secondary_total = series.values[primary_count..]
            .iter()
            .copied()
            .filter(|value| value.is_finite() && *value > 0.0)
            .sum();
        angle = std::f32::consts::PI * secondary_total / total;
        rendered_values.truncate(primary_count);
        rendered_values.push((series.values.len(), secondary_total));
    }
    let mut slices = Vec::new();
    for (index, value) in rendered_values {
        if !value.is_finite() || value <= 0.0 {
            continue;
        }
        let sweep = sweep_direction * std::f32::consts::TAU * value / total;
        let label_angle = angle + sweep / 2.0;
        let explosion = series.point_explosions.get(index).copied().unwrap_or_else(|| {
            if index == series.values.len() { series.point_explosions.first().copied().unwrap_or(0.0) } else { 0.0 }
        });
        let offset_x = explosion * radii.0 * label_angle.cos();
        let offset_y = explosion * radii.1 * label_angle.sin();
        let slice_center = (center.0 + offset_x, center.1 + offset_y);
        let segment_count = (sweep.abs() / (std::f32::consts::PI / 36.0))
            .ceil()
            .max(2.0) as usize;
        let outer_start = (
            slice_center.0 + radii.0 * angle.cos(),
            slice_center.1 + radii.1 * angle.sin(),
        );
        let outer_arcs = (0..segment_count)
            .map(|segment| {
                let from = angle + sweep * segment as f32 / segment_count as f32;
                let to = angle + sweep * (segment + 1) as f32 / segment_count as f32;
                ellipse_arc_bezier_points(slice_center, radii, from, to)
            })
            .collect::<Vec<_>>();
        let inner_radii = (radii.0 * 0.55, radii.1 * 0.55);
        let inner_start_angle = angle + sweep;
        let inner_start = (
            slice_center.0 + inner_radii.0 * inner_start_angle.cos(),
            slice_center.1 + inner_radii.1 * inner_start_angle.sin(),
        );
        let inner_arcs = doughnut
            .then(|| {
                (0..segment_count)
                    .rev()
                    .map(|segment| {
                        let from = angle + sweep * (segment + 1) as f32 / segment_count as f32;
                        let to = angle + sweep * segment as f32 / segment_count as f32;
                        ellipse_arc_bezier_points(slice_center, inner_radii, from, to)
                    })
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        let mut points = vec![outer_start];
        points.extend(outer_arcs.iter().flatten().copied());
        if doughnut {
            points.push(inner_start);
            points.extend(inner_arcs.iter().flatten().copied());
        } else {
            points.push(slice_center);
        }
        let min_x = points.iter().map(|(x, _)| *x).fold(f32::INFINITY, f32::min);
        let max_x = points
            .iter()
            .map(|(x, _)| *x)
            .fold(f32::NEG_INFINITY, f32::max);
        let min_y = points.iter().map(|(_, y)| *y).fold(f32::INFINITY, f32::min);
        let max_y = points
            .iter()
            .map(|(_, y)| *y)
            .fold(f32::NEG_INFINITY, f32::max);
        let bounds = Rect {
            x: min_x,
            y: min_y,
            width: max_x - min_x,
            height: max_y - min_y,
        };
        let local = |point: (f32, f32)| (point.0 - min_x, point.1 - min_y);
        let mut commands = if doughnut {
            let point = local(outer_start);
            vec![PathCommand::MoveTo {
                x: point.0,
                y: point.1,
            }]
        } else {
            let local_center = local(slice_center);
            let start = local(outer_start);
            vec![
                PathCommand::MoveTo {
                    x: local_center.0,
                    y: local_center.1,
                },
                PathCommand::LineTo {
                    x: start.0,
                    y: start.1,
                },
            ]
        };
        commands.extend(outer_arcs.into_iter().map(|[control_1, control_2, end]| {
            let control_1 = local(control_1);
            let control_2 = local(control_2);
            let end = local(end);
            PathCommand::BezierCurveTo {
                cp1x: control_1.0,
                cp1y: control_1.1,
                cp2x: control_2.0,
                cp2y: control_2.1,
                x: end.0,
                y: end.1,
            }
        }));
        if doughnut {
            let point = local(inner_start);
            commands.push(PathCommand::LineTo {
                x: point.0,
                y: point.1,
            });
            commands.extend(inner_arcs.into_iter().map(|[control_1, control_2, end]| {
                let control_1 = local(control_1);
                let control_2 = local(control_2);
                let end = local(end);
                PathCommand::BezierCurveTo {
                    cp1x: control_1.0,
                    cp1y: control_1.1,
                    cp2x: control_2.0,
                    cp2y: control_2.1,
                    x: end.0,
                    y: end.1,
                }
            }));
        }
        commands.push(PathCommand::ClosePath);
        let face_geometry = |quads: Vec<[(f32, f32); 4]>| {
            if quads.is_empty() {
                return None;
            }
            let min_x = quads
                .iter()
                .flatten()
                .map(|(x, _)| *x)
                .fold(f32::INFINITY, f32::min);
            let max_x = quads
                .iter()
                .flatten()
                .map(|(x, _)| *x)
                .fold(f32::NEG_INFINITY, f32::max);
            let min_y = quads
                .iter()
                .flatten()
                .map(|(_, y)| *y)
                .fold(f32::INFINITY, f32::min);
            let max_y = quads
                .iter()
                .flatten()
                .map(|(_, y)| *y)
                .fold(f32::NEG_INFINITY, f32::max);
            Some((
                Rect {
                    x: min_x,
                    y: min_y,
                    width: max_x - min_x,
                    height: max_y - min_y,
                },
                Geometry::Path {
                    fill_rule: FillRule::NonZero,
                    commands: quads
                        .into_iter()
                        .flat_map(|quad| {
                            [
                                PathCommand::MoveTo {
                                    x: quad[0].0 - min_x,
                                    y: quad[0].1 - min_y,
                                },
                                PathCommand::LineTo {
                                    x: quad[1].0 - min_x,
                                    y: quad[1].1 - min_y,
                                },
                                PathCommand::LineTo {
                                    x: quad[2].0 - min_x,
                                    y: quad[2].1 - min_y,
                                },
                                PathCommand::LineTo {
                                    x: quad[3].0 - min_x,
                                    y: quad[3].1 - min_y,
                                },
                                PathCommand::ClosePath,
                            ]
                        })
                        .collect(),
                },
            ))
        };
        let project_depth = |point: (f32, f32)| {
            (
                center.0 + (point.0 - center.0) * (1.0 + depth_perspective),
                point.1 + depth,
            )
        };
        let side = (depth > 0.0)
            .then(|| {
                let mut quads = Vec::new();
                for segment in 0..segment_count {
                    let from = angle + sweep * segment as f32 / segment_count as f32;
                    let to = angle + sweep * (segment + 1) as f32 / segment_count as f32;
                    if ((from + to) / 2.0).sin() > 0.0 {
                        let point = |value: f32| {
                            (
                                slice_center.0 + radii.0 * value.cos(),
                                slice_center.1 + radii.1 * value.sin(),
                            )
                        };
                        let from = point(from);
                        let to = point(to);
                        quads.push([from, to, (to.0, to.1 + depth), (from.0, from.1 + depth)]);
                    }
                }
                face_geometry(quads)
            })
            .flatten();
        let cut_side = (depth > 0.0 && explosion > 0.0)
            .then(|| {
                face_geometry(
                    [(angle, true), (angle + sweep, false)]
                        .into_iter()
                        .filter(|(edge, start)| {
                            let horizontal = edge.cos();
                            if *start {
                                horizontal <= 0.001
                            } else {
                                horizontal >= -0.001
                            }
                        })
                        .map(|(edge, _)| {
                            let outer = (
                                slice_center.0 + radii.0 * edge.cos(),
                                slice_center.1 + radii.1 * edge.sin(),
                            );
                            [
                                slice_center,
                                outer,
                                (outer.0, outer.1 + depth),
                                project_depth(slice_center),
                            ]
                        })
                        .collect(),
                )
            })
            .flatten();
        slices.push(ChartPieSlice {
            index,
            bounds,
            geometry: Geometry::Path {
                fill_rule: if doughnut {
                    FillRule::EvenOdd
                } else {
                    FillRule::NonZero
                },
                commands,
            },
            color: series
                .point_colors
                .get(index)
                .copied()
                .unwrap_or_else(|| super::office_chart_palette_color(index)),
            label_angle,
            value,
            center: slice_center,
            side,
            cut_side,
        });
        angle += sweep;
    }
    if depth > 0.0 {
        slices.sort_by(|left, right| {
            left.center
                .1
                .total_cmp(&right.center.1)
                .then_with(|| left.index.cmp(&right.index))
        });
    }
    (total, slices)
}

pub(super) fn chart_pie_side_paint(color: u32, bounds: Rect) -> Paint {
    let mut light = color;
    apply_color_transform(&mut light, "shade", 0.82);
    let mut dark = color;
    apply_color_transform(&mut dark, "shade", 0.38);
    Paint::LinearGradient {
        x0: 0.0,
        y0: bounds.height / 2.0,
        x1: bounds.width,
        y1: bounds.height / 2.0,
        stops: vec![
            GradientStop {
                offset: 0.0,
                color: light,
            },
            GradientStop {
                offset: 1.0,
                color: dark,
            },
        ],
    }
}

impl Chart {
    pub(super) fn area_border(&self, bounds: Rect, fallback: (Paint, f32)) -> (Paint, f32) {
        self.chart_area_border
            .as_ref()
            .map_or(fallback, |(fill, width)| (fill.paint(bounds), *width))
    }

    pub(super) fn category_axis_options(&self) -> &ChartValueAxis {
        if self
            .series
            .iter()
            .any(|series| series.kind == ChartKind::Bar && series.bar_horizontal)
        {
            &self.value_axis_options
        } else {
            &self.horizontal_axis_options
        }
    }

    pub(super) fn numerical_axis_options(&self) -> &ChartValueAxis {
        if self
            .series
            .iter()
            .any(|series| series.kind == ChartKind::Bar && series.bar_horizontal)
        {
            &self.horizontal_axis_options
        } else {
            &self.value_axis_options
        }
    }

    pub(super) fn value_axis_at_top(&self) -> bool {
        self.numerical_axis_options().position.as_deref() == Some("t")
    }

    fn secondary_value_axis_shares_primary_scale(&self) -> bool {
        let primary = self.numerical_axis_options();
        self.secondary_value_axis_options
            .as_ref()
            .is_some_and(|secondary| secondary.deleted && secondary.position == primary.position)
    }

    fn series_uses_primary_value_axis(&self, series: &ChartSeries) -> bool {
        let primary = self.numerical_axis_options();
        primary.id.is_none()
            || series.axis_id.as_ref() == primary.id.as_ref()
            || (self.secondary_value_axis_shares_primary_scale()
                && self
                    .secondary_value_axis_options
                    .as_ref()
                    .is_some_and(|secondary| series.axis_id.as_ref() == secondary.id.as_ref()))
    }

    fn has_vertical_3d_bars(&self) -> bool {
        self.series
            .iter()
            .any(|series| series.kind == ChartKind::Bar && series.three_d && !series.bar_horizontal)
    }

    pub(super) fn axis_line_visible(&self, horizontal: bool) -> bool {
        let axis = if horizontal {
            &self.horizontal_axis_options
        } else {
            &self.value_axis_options
        };
        let floor_supplies_axis = horizontal && self.has_vertical_3d_bars();
        !floor_supplies_axis && !axis.deleted && !axis.axis_line_hidden
    }

    pub(super) fn series_axis_line_visible(&self) -> bool {
        !self.has_vertical_3d_bars()
            && self
                .series_axis_options
                .as_ref()
                .is_some_and(|axis| !axis.deleted && !axis.axis_line_hidden)
    }

    pub(super) fn axis_labels_visible(&self, horizontal: bool) -> bool {
        !(if horizontal {
            self.horizontal_axis_options.deleted
        } else {
            self.value_axis_options.deleted
        })
    }

    pub(super) fn value_axis_grid_style(&self, color: u32, width: f32) -> (u32, f32) {
        let axis = self.numerical_axis_options();
        (
            axis.major_gridline_color.unwrap_or(color),
            axis.major_gridline_width.unwrap_or(width),
        )
    }

    pub(super) fn value_axis_grid_lines_visible(&self) -> bool {
        self.numerical_axis_options().major_gridlines
    }

    pub(super) fn value_axis_for_series(&self, series: &ChartSeries) -> (f32, f32, f32) {
        if self.series_uses_primary_value_axis(series) {
            self.value_axis()
        } else {
            self.secondary_value_axis_options
                .as_ref()
                .filter(|axis| axis.id.as_ref() == series.axis_id.as_ref())
                .map_or_else(|| self.value_axis(), |axis| self.axis_range(axis))
        }
    }

    pub(super) fn secondary_value_axis_ticks(&self) -> Option<(&ChartValueAxis, Vec<(f32, f32)>)> {
        let axis = self.secondary_value_axis_options.as_ref()?;
        if axis.deleted || axis.position.as_deref() != Some("r") {
            return None;
        }
        let (minimum, maximum, major) = self.axis_range(axis);
        let mut ticks = Vec::new();
        let mut value = minimum;
        while value <= maximum + major * 0.01 && ticks.len() < 1_000 {
            ticks.push((value, (value - minimum) / (maximum - minimum)));
            value += major;
        }
        Some((axis, ticks))
    }

    pub(super) fn secondary_value_axis_tick_bounds(&self, plot: Rect, ratio: f32) -> Option<Rect> {
        let axis = self.secondary_value_axis_options.as_ref()?;
        let (x, width) = match axis.major_tick_mark {
            ChartAxisTickMark::None => return None,
            ChartAxisTickMark::In => (plot.x + plot.width - 4.0, 4.0),
            ChartAxisTickMark::Out => (plot.x + plot.width, 4.0),
            ChartAxisTickMark::Cross => (plot.x + plot.width - 2.0, 4.0),
        };
        Some(Rect {
            x,
            y: plot.y + plot.height * (1.0 - ratio.clamp(0.0, 1.0)),
            width,
            height: 0.01,
        })
    }

    fn log_axis_range(
        axis: &ChartValueAxis,
        values: impl Iterator<Item = f32>,
    ) -> Option<(f32, f32, f32)> {
        let base = axis.log_base?;
        let mut values = values.filter(|value| value.is_finite() && *value > 0.0);
        let first = values.next()?;
        let (raw_minimum, raw_maximum) = values
            .fold((first, first), |(minimum, maximum), value| {
                (minimum.min(value), maximum.max(value))
            });
        let minimum = axis
            .minimum
            .filter(|value| *value > 0.0)
            .unwrap_or_else(|| base.powf(raw_minimum.log(base).floor()));
        let maximum = axis
            .maximum
            .filter(|value| *value > minimum)
            .unwrap_or_else(|| base.powf(raw_maximum.log(base).ceil()).max(minimum * base));
        Some((minimum, maximum, base))
    }

    fn axis_range(&self, axis: &ChartValueAxis) -> (f32, f32, f32) {
        if let Some(range) = Self::log_axis_range(
            axis,
            self.series
                .iter()
                .filter(|series| series.axis_id.as_ref() == axis.id.as_ref())
                .flat_map(ChartSeries::value_spans)
                .flat_map(|(start, end)| [start, end]),
        ) {
            return range;
        }
        if let Some(base) = axis.log_base {
            return (1.0, base, base);
        }
        let mut values = self
            .series
            .iter()
            .filter(|series| series.axis_id.as_ref() == axis.id.as_ref())
            .flat_map(ChartSeries::value_spans)
            .flat_map(|(start, end)| [start, end])
            .filter(|value| value.is_finite());
        let Some(first) = values.next() else {
            return (0.0, 1.0, 0.2);
        };
        let (raw_minimum, raw_maximum) = values
            .fold((first, first), |(minimum, maximum), value| {
                (minimum.min(value), maximum.max(value))
            });
        let raw_minimum = raw_minimum.min(0.0);
        let raw_maximum = raw_maximum.max(0.0);
        let span = (axis.maximum.unwrap_or(raw_maximum) - axis.minimum.unwrap_or(raw_minimum))
            .abs()
            .max(f32::MIN_POSITIVE);
        if axis.minimum.is_none() && axis.maximum.is_none() && axis.major_unit.is_none() {
            return automatic_chart_axis_range(
                raw_minimum,
                raw_maximum,
                chart_axis_needs_auto_headroom(
                    self.series
                        .iter()
                        .filter(|series| series.axis_id.as_ref() == axis.id.as_ref()),
                ),
            );
        }
        let major = axis
            .major_unit
            .unwrap_or_else(|| nice_chart_step(span / 7.0));
        let minimum = axis
            .minimum
            .unwrap_or_else(|| (raw_minimum / major).floor() * major);
        let mut maximum = axis
            .maximum
            .unwrap_or_else(|| (raw_maximum / major).ceil() * major);
        if axis.maximum.is_none() && maximum - raw_maximum < major * 0.5 {
            maximum += major;
        }
        (minimum, maximum.max(minimum + major), major)
    }

    // With an authored base font, these ordinary 2-D axes have identical
    // automatic layout semantics in all OOXML hosts. Manual layouts, legacy
    // native canvases, 3-D, axis titles and rotated/wrapped labels retain their
    // existing layout paths rather than borrowing this narrower contract.
    pub(super) fn automatic_cartesian_layout(&self, bounds: Rect) -> Option<(Rect, Option<Rect>)> {
        let font_size = self
            .font_size
            .filter(|size| size.is_finite() && *size > 0.0)?;
        let series = self.series.first()?;
        let axis = self.category_axis_options();
        if self.native_size.is_some()
            || self.style.is_some()
            || self.plot_bounds.is_some()
            || self.legend_bounds.is_some()
            || self.show_title
            || self.data_table.is_some()
            || self.secondary_value_axis_options.is_some()
            || !axis.title.is_empty()
            || !self.numerical_axis_options().title.is_empty()
            || axis
                .label_rotation_degrees
                .is_some_and(|angle| angle != 0.0)
            || axis.is_date
            || series.categories.len() > 6
            || !self.axis_labels_visible(true)
            || !self.axis_labels_visible(false)
            || self.show_legend && self.legend_position != ChartLegendPosition::Right
            || self.series.iter().any(|s| {
                (s.three_d && !s.bar_depth)
                    || s.bar_horizontal
                    || !s.category_levels.is_empty()
                    || !matches!(s.kind, ChartKind::Bar | ChartKind::Line | ChartKind::Area)
            })
        {
            return None;
        }
        let value_font = self.axis_label_font_size(self.numerical_axis_options(), font_size);
        let category_font = self.axis_label_font_size(axis, font_size);
        let pad = font_size * 0.4;
        // ponytail: use the existing chart text estimate until chart layout receives font metrics.
        let value_width = self
            .value_axis_ticks()
            .iter()
            .map(|(value, _)| chart_text_width(&self.value_axis_label(*value), value_font))
            .fold(0.0_f32, f32::max);
        let legend = if self.show_legend {
            let style = self.legend_text_style();
            let entries = self.legend_entries();
            let width = entries
                .iter()
                .map(|(index, label, _)| {
                    chart_text_width(label, style.font_size)
                        + self.legend_key_width(*index, style.font_size * 0.65)
                        + 12.0
                })
                .fold(0.0_f32, f32::max);
            let height = entries.len() as f32 * style.font_size * 1.45 + 8.0;
            Some(Rect {
                x: bounds.x + bounds.width - width - 4.0,
                y: bounds.y + (bounds.height - height) / 2.0,
                width,
                height,
            })
        } else {
            None
        };
        let x = bounds.x + pad + value_width + value_font * 0.9;
        let y = bounds.y + font_size * 0.9;
        let right = legend.map_or(bounds.x + bounds.width - pad, |r| r.x - pad * 2.0)
            - if series.bar_depth { font_size * 3.0 } else { 0.0 };
        let bottom = bounds.y + bounds.height - pad - category_font * 1.8;
        let plot = Rect {
            x,
            y,
            width: right - x,
            height: bottom - y,
        };
        let slot = plot.width / series.categories.len().max(1) as f32;
        if plot.width <= 0.0
            || plot.height <= 0.0
            || series.categories.iter().any(|label| {
                label
                    .chars()
                    .map(|c| drawingml_fallback_character_width(c, category_font))
                    .sum::<f32>()
                    > slot * 0.9
            })
        {
            return None;
        }
        Some((plot, legend))
    }

    pub(super) fn plot_area_bounds(&self, bounds: Rect, mut fallback: Rect) -> Rect {
        if let Some((plot, _)) = self.automatic_cartesian_layout(bounds) {
            return plot;
        }
        if self.plot_bounds.is_none()
            && self.axis_labels_visible(false)
            && self.series.iter().any(|s| {
                matches!(s.kind, ChartKind::Line | ChartKind::Area | ChartKind::Bar)
                    && !s.bar_horizontal
            })
        {
            let axis = self.numerical_axis_options();
            let font_size = self.axis_label_font_size(axis, 12.0);
            let label_width = self
                .value_axis_ticks()
                .iter()
                .map(|(value, _)| {
                    self.value_axis_label(*value)
                        .chars()
                        .map(|c| drawingml_fallback_character_width(c, font_size))
                        .sum::<f32>()
                })
                .fold(0.0_f32, f32::max);
            // ponytail: fallback metrics plus breathing room; use font metrics when charts receive them.
            let left = bounds.x
                + label_width * 1.12
                + font_size * 0.8
                + if axis.title.is_empty() { 0.0 } else { 24.0 };
            let shift = (left - fallback.x).max(0.0).min(fallback.width * 0.5);
            fallback.x += shift;
            fallback.width -= shift;
        }
        if self.plot_bounds.is_none() && self.series.iter().any(|series|
            matches!(series.kind, ChartKind::Bar | ChartKind::Line | ChartKind::Area)) {
            let levels = self.series.first().map_or(0, |s| s.category_levels.len().saturating_sub(1));
            if levels > 0 && !self.category_axis_options().no_multi_level_labels {
                let space = (levels as f32 * self.axis_label_font_size(self.category_axis_options(), 12.0) * 1.5).min(fallback.height * 0.4);
                if self.series.iter().any(|s| s.bar_horizontal) {
                    fallback.x += space;
                    fallback.width = (fallback.width - space).max(1.0);
                } else {
                    fallback.height -= space;
                }
            }
        }
        if self.plot_bounds.is_none() {
            if let Some((legend, _, _)) = self.constrained_legend(bounds) {
                if self.legend_position == ChartLegendPosition::Left {
                    let right = fallback.x + fallback.width;
                    fallback.x = legend.x + legend.width + 12.0;
                    fallback.width = (right - fallback.x).max(1.0);
                } else { fallback.width = (legend.x - 12.0 - fallback.x).max(1.0); }
            }
        }
        if self.plot_bounds.is_none()
            && let Some(series) = self.series.first()
            && matches!(series.kind, ChartKind::Bar | ChartKind::Line | ChartKind::Area)
            && !series.bar_horizontal && self.dense_category_labels(series, fallback)
        {
            let font = self.axis_label_font_size(self.category_axis_options(), 12.0);
            let label_height = series.categories.iter().map(|text| chart_text_width(text, font))
                .fold(0.0_f32, f32::max) + 12.0;
            fallback.height = fallback.height.min((bounds.y + bounds.height - label_height - fallback.y).max(bounds.height * 0.3));
        }
        if self.classic_defaults && self.plot_bounds.is_none() && bounds.height < 240.0 {
            if self.series.iter().all(|s| matches!(s.kind, ChartKind::Scatter | ChartKind::Bubble)) {
                fallback.x = fallback.x.max(bounds.x + self.axis_label_font_size(&self.value_axis_options, 12.0) * 2.0);
                let reserve = if self.show_legend && self.legend_position == ChartLegendPosition::Right {
                    let font = self.legend_text_style().font_size;
                    font * 4.0 + self.legend_key_width(0, 8.0) + 20.0
                } else { 12.0 };
                fallback.width = (bounds.x + bounds.width - reserve - fallback.x).max(1.0);
            }
            let bottom = (fallback.y + fallback.height).min(bounds.y + bounds.height - 32.0);
            fallback.y = fallback.y.max(bounds.y + if self.show_title { self.title_text_style().font_size * 2.25 } else { 14.0 });
            fallback.height = (bottom - fallback.y).max(1.0);
        }
        self.plot_bounds.map_or(fallback, |layout| Rect {
            x: bounds.x + bounds.width * layout.x,
            y: bounds.y + bounds.height * layout.y,
            width: bounds.width * layout.width,
            height: bounds.height * layout.height,
        })
    }

    pub(super) fn value_axis_label(&self, value: f32) -> String {
        if self.value_axis_is_percent() {
            format!("{:.0}%", value * 100.0)
        } else {
            self.numerical_axis_options().format_label(value)
        }
    }

    fn date_axis_range(&self, series: &ChartSeries) -> Option<(f32, f32)> {
        let axis = self.category_axis_options();
        if !axis.is_date || series.x_values.len() != series.values.len() {
            return None;
        }
        if let (Some(minimum), Some(maximum)) = (axis.minimum, axis.maximum) {
            return (maximum > minimum).then_some((minimum, maximum));
        }
        let (data_minimum, data_maximum) = series
            .x_values
            .iter()
            .copied()
            .filter(|value| value.is_finite())
            .fold(
                (f32::INFINITY, f32::NEG_INFINITY),
                |(minimum, maximum), value| (minimum.min(value), maximum.max(value)),
            );
        let minimum = axis.minimum.unwrap_or(data_minimum);
        let maximum = axis.maximum.unwrap_or(data_maximum);
        (maximum > minimum).then_some((minimum, maximum))
    }

    #[cfg(feature = "native-formats")]
    fn date_axis_marks(&self, series: &ChartSeries, plot: Rect) -> Option<Vec<(f32, String, bool)>> {
        let axis = self.category_axis_options();
        let (minimum, maximum) = self.date_axis_range(series)?;
        let major = axis.major_unit?.round() as i64;
        let major_unit = axis.major_time_unit.or(axis.base_time_unit)?;
        if major <= 0 || !(-2_000_000.0..=2_000_000.0).contains(&minimum)
            || !(-2_000_000.0..=2_000_000.0).contains(&maximum) { return None; }
        let mut marks = Vec::new();
        for (unit, step, is_major) in [
            (major_unit, major, true),
            (axis.minor_time_unit.or(axis.base_time_unit).unwrap_or(ChartTimeUnit::Days),
             axis.minor_unit.unwrap_or(0.0).round() as i64, false),
        ] {
            if step <= 0 || (!is_major && !axis.minor_gridlines && axis.minor_tick_mark == ChartAxisTickMark::None) {
                continue;
            }
            let mut serial = minimum.floor() as i64;
            let (year, month, _) = excel_serial_date(serial, self.date_1904);
            let mut calendar_index = match unit {
                ChartTimeUnit::Days => serial,
                ChartTimeUnit::Months => year * 12 + i64::from(month) - 1,
                ChartTimeUnit::Years => year,
            };
            if unit != ChartTimeUnit::Days {
                calendar_index = calendar_index.div_euclid(step) * step;
            }
            for _ in 0..1_000 {
                serial = match unit {
                    ChartTimeUnit::Days => calendar_index,
                    ChartTimeUnit::Months => excel_serial_from_date(
                        calendar_index.div_euclid(12), calendar_index.rem_euclid(12) as u32 + 1,
                        1, self.date_1904),
                    ChartTimeUnit::Years => excel_serial_from_date(calendar_index, 1, 1, self.date_1904),
                };
                let Some(next_index) = calendar_index.checked_add(step) else { break; };
                calendar_index = next_index;
                if serial as f32 > maximum { break; }
                if (serial as f32) < minimum { continue; }
                let ratio = (serial as f32 - minimum) / (maximum - minimum);
                let ratio = if axis.reversed { 1.0 - ratio } else { ratio };
                let label = if is_major {
                    format_chart_category(
                        &(serial + i64::from(self.date_1904)).to_string(),
                        axis.number_format.as_deref(), self.date_1904,
                    ).unwrap_or_else(|| {
                        let (year, month, day) = excel_serial_date(serial, self.date_1904);
                        format!("{month}/{day}/{year}")
                    })
                } else { String::new() };
                marks.push((plot.x + ratio * plot.width, label, is_major));
            }
        }
        Some(marks)
    }

    fn date_axis_has_calendar_marks(&self, series: &ChartSeries) -> bool {
        let axis = self.category_axis_options();
        axis.is_date && axis.major_unit.is_some_and(|unit| unit >= 0.5)
            && axis.major_time_unit.or(axis.base_time_unit).is_some()
            && self.date_axis_range(series).is_some()
    }

    #[cfg(not(feature = "native-formats"))]
    fn date_axis_marks(&self, _series: &ChartSeries, _plot: Rect) -> Option<Vec<(f32, String, bool)>> {
        None
    }

    pub(super) fn category_index(&self, index: usize, count: usize) -> usize {
        if self.category_axis_options().reversed { count.saturating_sub(index + 1) } else { index }
    }

    pub(super) fn depth_category_font_size(
        &self,
        series: &ChartSeries,
        plot: Rect,
        font_size: f32,
    ) -> f32 {
        if !series.projected_column() {
            return font_size;
        }
        let count = series.values.len().max(1);
        let available = (0..count)
            .map(|index| {
                let left = self
                    .depth_axis_point(plot, index as f32 / count as f32, 0.0, 0.0)
                    .unwrap()
                    .0;
                let right = self
                    .depth_axis_point(plot, (index + 1) as f32 / count as f32, 0.0, 0.0)
                    .unwrap()
                    .0;
                (right - left).abs()
            })
            .fold(f32::INFINITY, f32::min);
        let width = series
            .categories
            .iter()
            .map(|label| chart_text_width(label, font_size))
            .fold(0.0_f32, f32::max);
        font_size * (available * 0.9 / width.max(1.0)).min(1.0)
    }

    pub(super) fn category_baseline(&self, series: &ChartSeries, index: usize, plot: Rect) -> f32 {
        let count = series.values.len().max(1);
        self.depth_axis_point(
            plot,
            (self.category_index(index, count) as f32 + 0.5) / count as f32,
            0.0,
            0.0,
        )
        .map_or(plot.y + plot.height, |p| p.1)
    }

    pub(super) fn category_x(
        &self,
        series: &ChartSeries,
        index: usize,
        plot: Rect,
        centered: bool,
    ) -> Option<f32> {
        if series.projected_column() {
            let x = (self.category_index(index, series.values.len()) as f32+0.5)/series.values.len().max(1) as f32;
            return Some(chart_project_3d(plot, self.column_view_3d(), x, 0.0, 0.0).0);
        }
        if let Some((minimum, maximum)) = self.date_axis_range(series) {
            let value = *series.x_values.get(index)?;
            if !value.is_finite() || !(minimum..=maximum).contains(&value) {
                return None;
            }
            let ratio = (value - minimum) / (maximum - minimum);
            return Some(plot.x + if self.category_axis_options().reversed { 1.0 - ratio } else { ratio } * plot.width);
        }
        let count = series.values.len().max(series.categories.len());
        if index >= count {
            return None;
        }
        let index = self.category_index(index, count);
        Some(if count == 1 {
            plot.x + plot.width / 2.0
        } else if centered || series.kind == ChartKind::Bar || self.value_axis_options.cross_between_categories {
            plot.x + (index as f32 + 0.5) / count as f32 * plot.width
        } else {
            plot.x + index as f32 / (count - 1) as f32 * plot.width
        })
    }

    pub(super) fn area_plot_bounds(&self, series: &ChartSeries, plot: Rect) -> Option<Rect> {
        let last = series.values.len().checked_sub(1)?;
        let first_x = self.category_x(series, 0, plot, false)?;
        let last_x = self.category_x(series, last, plot, false)?;
        Some(Rect {
            x: first_x.min(last_x),
            y: plot.y,
            width: (last_x - first_x).abs(),
            height: plot.height,
        })
    }

    pub(super) fn dense_category_labels(&self, series: &ChartSeries, plot: Rect) -> bool {
        let axis = self.category_axis_options();
        axis.label_rotation_degrees.is_none() && axis.label_skip.is_none()
            && series.categories.len().max(series.values.len()) as f32
                * series.categories.iter().map(|label| chart_text_width(label, self.axis_label_font_size(axis, 12.0)) + 4.0).fold(18.0_f32, f32::max)
                > plot.width
    }

    pub(super) fn category_has_label(
        &self,
        series: &ChartSeries,
        index: usize,
        plot: Rect,
        minimum_spacing: f32,
    ) -> bool {
        let axis = self.category_axis_options();
        if self.date_axis_has_calendar_marks(series) { return false; }
        let skip = axis.label_skip.unwrap_or_else(|| {
            if axis.is_date { 1 } else {
                (series.categories.len().max(series.values.len()) as f32 * minimum_spacing / plot.width.max(1.0))
                    .ceil().max(1.0) as usize
            }
        });
        if !index.is_multiple_of(skip) {
            return false;
        }
        if axis.is_date
            && axis.major_unit.is_none()
            && series.categories.len() as f32 * minimum_spacing <= plot.width
        {
            // Calendar categories fit: do not drop months or leap years because
            // their elapsed days are not divisible by a fixed tick interval.
            return true;
        }
        self.category_is_major_interval(series, index, plot, minimum_spacing)
    }

    fn category_is_major_interval(
        &self,
        series: &ChartSeries,
        index: usize,
        plot: Rect,
        minimum_spacing: f32,
    ) -> bool {
        let axis = self.category_axis_options();
        let Some((minimum, maximum)) = self.date_axis_range(series) else {
            return true;
        };
        let Some(value) = series.x_values.get(index).copied() else {
            return false;
        };
        let major = axis.major_unit.unwrap_or_else(|| {
            let capacity = (plot.width / minimum_spacing.max(1.0)).floor().max(1.0);
            let requested = (maximum - minimum) / capacity;
            [
                1.0, 2.0, 3.0, 5.0, 7.0, 10.0, 14.0, 30.0, 60.0, 90.0, 180.0, 365.0,
            ]
            .into_iter()
            .find(|unit| *unit >= requested)
            .unwrap_or_else(|| requested.ceil())
        });
        let tick = (value - minimum) / major;
        (tick - tick.round()).abs() < 0.001
    }

    pub(super) fn category_axis_tick_bounds(
        &self,
        series: &ChartSeries,
        index: usize,
        plot: Rect,
    ) -> Option<Rect> {
        let axis = self.category_axis_options();
        if self.date_axis_has_calendar_marks(series) { return None; }
        if !index.is_multiple_of(axis.tick_skip.unwrap_or(1)) {
            return None;
        }
        let category_count = series.values.len().max(series.categories.len());
        let x = if self.value_axis_options.cross_between_categories
            && !self.horizontal_axis_options.is_date
        {
            (index <= category_count).then(|| {
                    let index = if axis.reversed { category_count - index } else { index };
                    plot.x + index as f32 / category_count.max(1) as f32 * plot.width
                })?
        } else {
            self.category_is_major_interval(series, index, plot, 18.0)
                .then(|| self.category_x(series, index, plot, true))??
        };
        let (y, height) = match axis.major_tick_mark {
            ChartAxisTickMark::None => return None,
            ChartAxisTickMark::In => (plot.y + plot.height - 4.0, 4.0),
            ChartAxisTickMark::Out => (plot.y + plot.height, 4.0),
            ChartAxisTickMark::Cross => (plot.y + plot.height - 2.0, 4.0),
        };
        let bounds = Rect {
            x,
            y,
            width: 0.01,
            height,
        };
        Some(
            if series.kind == ChartKind::Bar && series.three_d && !series.bar_horizontal {
                chart_bar_3d_plane_bounds(bounds, plot, self.view_3d.unwrap_or_default(), false)
            } else {
                bounds
            },
        )
    }

    fn line_value(&self, target_index: usize, category_index: usize) -> Option<f32> {
        let target = self
            .series
            .get(target_index)
            .filter(|series| series.kind == ChartKind::Line)?;
        let value = target.values.get(category_index).copied()?;
        if !value.is_finite() || target.grouping == ChartGrouping::Standard {
            return value.is_finite().then_some(value);
        }
        let matching = self.series.iter().enumerate().filter(|(_, series)| {
            series.kind == ChartKind::Line
                && series.axis_id == target.axis_id
                && series.grouping == target.grouping
        });
        let same_sign =
            |candidate: f32| candidate.is_finite() && (candidate < 0.0) == (value < 0.0);
        let mut stacked = 0.0;
        let mut denominator = 0.0;
        for (index, series) in matching {
            let candidate = series.values.get(category_index).copied().unwrap_or(0.0);
            if same_sign(candidate) {
                denominator += candidate.abs();
                if index <= target_index {
                    stacked += candidate;
                }
            }
        }
        if target.grouping == ChartGrouping::PercentStacked {
            (denominator > f32::EPSILON).then_some(stacked / denominator)
        } else {
            Some(stacked)
        }
    }

    pub(super) fn line_points(
        &self,
        target_index: usize,
        plot: Rect,
        (minimum, maximum): (f32, f32),
        centered: bool,
    ) -> Vec<ChartPoint> {
        let Some(series) = self.series.get(target_index) else {
            return Vec::new();
        };
        series
            .values
            .iter()
            .enumerate()
            .filter_map(|(index, _)| {
                let value = self.line_value(target_index, index)?;
                let ratio = self.value_axis_ratio(value, minimum, maximum)?;
                Some((
                    self.category_x(series, index, plot, centered)?,
                    plot.y + plot.height - plot.height * ratio,
                    index,
                ))
            })
            .collect()
    }

    pub(super) fn up_down_bar_bounds(
        &self,
        plot: Rect,
        value_axis: (f32, f32),
    ) -> Vec<(usize, Rect, bool)> {
        let Some(options) = self.up_down_bars.as_ref() else {
            return Vec::new();
        };
        let line_indices = (options.series_start..options.series_end.min(self.series.len()))
            .filter(|index| self.series[*index].kind == ChartKind::Line)
            .collect::<Vec<_>>();
        let (Some(first), Some(last)) = (line_indices.first(), line_indices.last()) else {
            return Vec::new();
        };
        let category_count = self.series[*first]
            .values
            .len()
            .min(self.series[*last].values.len());
        let category_width = plot.width / category_count.max(1) as f32;
        let width = category_width / (1.0 + f32::from(options.gap_width) / 100.0);
        (0..category_count)
            .filter_map(|index| {
                let first_value = self.line_value(*first, index)?;
                let last_value = self.line_value(*last, index)?;
                let x = self.category_x(&self.series[*first], index, plot, true)?;
                let y = |value: f32| {
                    let ratio =
                        ((value - value_axis.0) / (value_axis.1 - value_axis.0)).clamp(0.0, 1.0);
                    plot.y + plot.height - plot.height * ratio
                };
                let first_y = y(first_value);
                let last_y = y(last_value);
                Some((
                    index,
                    Rect {
                        x: x - width / 2.0,
                        y: first_y.min(last_y),
                        width,
                        height: (first_y - last_y).abs().max(0.01),
                    },
                    last_value >= first_value,
                ))
            })
            .collect()
    }

    pub(super) fn value_axis_tick_bounds(
        &self,
        plot: Rect,
        ratio: f32,
        horizontal_bars: bool,
    ) -> Option<Rect> {
        let ratio = ratio.clamp(0.0, 1.0);
        let mark = self.numerical_axis_options().major_tick_mark;
        if horizontal_bars {
            let (y, height) = match mark {
                ChartAxisTickMark::None => return None,
                ChartAxisTickMark::In => (plot.y + plot.height - 4.0, 4.0),
                ChartAxisTickMark::Out => (plot.y + plot.height, 4.0),
                ChartAxisTickMark::Cross => (plot.y + plot.height - 2.0, 4.0),
            };
            return Some(Rect {
                x: plot.x + ratio * plot.width,
                y,
                width: 0.01,
                height,
            });
        }
        let (x, width) = match mark {
            ChartAxisTickMark::None => return None,
            ChartAxisTickMark::In => (plot.x, 4.0),
            ChartAxisTickMark::Out => (plot.x - 4.0, 4.0),
            ChartAxisTickMark::Cross => (plot.x - 2.0, 4.0),
        };
        Some(Rect {
            x,
            y: plot.y + plot.height * (1.0 - ratio),
            width,
            height: 0.01,
        })
    }

    pub(super) fn scatter_axis(&self, horizontal: bool) -> (f32, f32, f32) {
        let options = if horizontal {
            &self.horizontal_axis_options
        } else {
            &self.value_axis_options
        };
        let (mut minimum, mut maximum, has_error_bars) = self
            .series
            .iter()
            .filter(|series| matches!(series.kind, ChartKind::Scatter | ChartKind::Bubble))
            .fold(
                (f32::INFINITY, f32::NEG_INFINITY, false),
                |(minimum, maximum, has_error_bars), series| {
                    let values = if horizontal {
                        &series.x_values
                    } else {
                        &series.values
                    };
                    values
                        .iter()
                        .copied()
                        .enumerate()
                        .filter(|(_, value)| value.is_finite())
                        .fold(
                            (minimum, maximum, has_error_bars),
                            |(minimum, maximum, has_error_bars), (index, value)| {
                                let errors = if horizontal {
                                    &series.x_error_bars
                                } else {
                                    &series.y_error_bars
                                };
                                let (lower, upper) = errors
                                    .as_ref()
                                    .map_or((0.0, 0.0), |errors| errors.amounts(index));
                                (
                                    minimum.min(value - lower),
                                    maximum.max(value + upper),
                                    has_error_bars || lower > 0.0 || upper > 0.0,
                                )
                            },
                        )
                },
            );
        if !minimum.is_finite() || !maximum.is_finite() {
            minimum = 0.0;
            maximum = 1.0;
        }
        // Tight scatter data needs a local scale; broad positive/negative ranges include zero.
        if minimum > 0.0 && minimum < maximum * 0.8 {
            minimum = 0.0;
        }
        if maximum < 0.0 && maximum > minimum * 0.8 {
            maximum = 0.0;
        }
        let raw_minimum = minimum;
        if (maximum - minimum).abs() <= f32::EPSILON {
            maximum = maximum.max(1.0);
            minimum = minimum.min(0.0);
        }
        let raw_maximum = maximum;
        let bubble = self
            .series
            .iter()
            .any(|series| series.kind == ChartKind::Bubble);
        let span = (options.maximum.unwrap_or(maximum) - options.minimum.unwrap_or(minimum)).abs();
        let step = options
            .major_unit
            .unwrap_or_else(|| nice_chart_step(span / if bubble { 3.0 } else { 5.0 }));
        minimum = options
            .minimum
            .unwrap_or_else(|| (minimum / step).floor() * step);
        maximum = options
            .maximum
            .unwrap_or_else(|| (maximum / step).ceil() * step);
        if options.maximum.is_none()
            && raw_maximum > 0.0
            && maximum - raw_maximum
                < step
                    * if bubble || has_error_bars {
                        0.5
                    } else {
                        0.000_1
                    }
        {
            maximum += step;
        }
        if has_error_bars
            && options.minimum.is_none()
            && minimum != 0.0
            && raw_minimum - minimum < step * 0.5
        {
            minimum -= step;
        }
        (minimum, maximum.max(minimum + step), step)
    }

    pub(super) fn scatter_points(
        &self,
        series: &ChartSeries,
        plot: Rect,
    ) -> Vec<(f32, f32, usize)> {
        let (minimum_x, maximum_x, _) = self.scatter_axis(true);
        let (minimum_y, maximum_y, _) = self.scatter_axis(false);
        if maximum_x <= minimum_x || maximum_y <= minimum_y {
            return Vec::new();
        }
        series
            .x_values
            .iter()
            .copied()
            .zip(series.values.iter().copied())
            .enumerate()
            .filter_map(|(index, (x, y))| {
                (x.is_finite() && y.is_finite()).then_some((
                    plot.x + (x - minimum_x) / (maximum_x - minimum_x) * plot.width,
                    plot.y + plot.height - (y - minimum_y) / (maximum_y - minimum_y) * plot.height,
                    index,
                ))
            })
            .collect()
    }

    pub(super) fn scatter_error_bars(
        &self,
        series: &ChartSeries,
        plot: Rect,
    ) -> Vec<(usize, DrawingMlElement)> {
        let mut elements = Vec::new();
        let points = self.scatter_points(series, plot);
        for (horizontal, errors) in [(true, &series.x_error_bars), (false, &series.y_error_bars)] {
            let Some(errors) = errors.as_ref().filter(|errors| errors.stroke_width > 0.0) else {
                continue;
            };
            let (minimum, maximum, _) = self.scatter_axis(horizontal);
            let scale = if horizontal { plot.width } else { plot.height } / (maximum - minimum);
            for &(x, y, index) in &points {
                if x < plot.x || x > plot.x + plot.width || y < plot.y || y > plot.y + plot.height {
                    continue;
                }
                let (minus, plus) = errors.amounts(index);
                if minus == 0.0 && plus == 0.0 {
                    continue;
                }
                let (start, end, low, high) = if horizontal {
                    (
                        x - minus * scale,
                        x + plus * scale,
                        plot.x,
                        plot.x + plot.width,
                    )
                } else {
                    (
                        y - plus * scale,
                        y + minus * scale,
                        plot.y,
                        plot.y + plot.height,
                    )
                };
                let first = start.clamp(low, high);
                let last = end.clamp(low, high);
                let bounds = if horizontal {
                    Rect {
                        x: first,
                        y: y - 3.0,
                        width: (last - first).max(0.01),
                        height: 6.0,
                    }
                } else {
                    Rect {
                        x: x - 3.0,
                        y: first,
                        width: 6.0,
                        height: (last - first).max(0.01),
                    }
                };
                let point = |along: f32, across: f32| {
                    if horizontal {
                        (along, across)
                    } else {
                        (across, along)
                    }
                };
                let mut commands = Vec::new();
                let mut line = |a: (f32, f32), b: (f32, f32)| {
                    commands.push(PathCommand::MoveTo { x: a.0, y: a.1 });
                    commands.push(PathCommand::LineTo { x: b.0, y: b.1 });
                };
                line(point(0.0, 3.0), point(last - first, 3.0));
                if errors.end_caps {
                    if start >= low && (if horizontal { minus } else { plus }) > 0.0 {
                        line(point(0.0, 0.0), point(0.0, 6.0));
                    }
                    if end <= high && (if horizontal { plus } else { minus }) > 0.0 {
                        line(point(last - first, 0.0), point(last - first, 6.0));
                    }
                }
                elements.push((
                    index,
                    DrawingMlElement {
                        bounds,
                        text: None,
                        visual: Visual::PaintedShape {
                            geometry: Geometry::Path {
                                fill_rule: FillRule::NonZero,
                                commands,
                            },
                            fill: Paint::None,
                            stroke: errors
                                .stroke
                                .as_ref()
                                .map_or(Paint::Solid(0x0000_00ff), |fill| fill.paint(bounds)),
                            stroke_width: errors.stroke_width,
                        },
                    },
                ));
            }
        }
        elements
    }

    pub(super) fn pie_bounds(&self, bounds: Rect, fallback: Rect, three_d: bool) -> Rect {
        let area = self.plot_bounds.map_or(fallback, |layout| Rect {
            x: bounds.x + bounds.width * layout.x,
            y: bounds.y + bounds.height * layout.y,
            width: bounds.width * layout.width,
            height: bounds.height * layout.height,
        });
        if three_d {
            let explosion = self.series.first().map_or(0.0, |series| series.point_explosions.iter().copied().fold(0.0, f32::max));
            let width = area.width / (1.0 + explosion);
            return Rect { x: area.x + (area.width-width)/2.0, width, ..area };
        }
        let size = area.width.min(area.height).max(0.0);
        Rect {
            x: area.x + (area.width - size) / 2.0,
            y: area.y + (area.height - size) / 2.0,
            width: size,
            height: size,
        }
    }

    pub(super) fn value_axis_is_percent(&self) -> bool {
        let mut series = self
            .series
            .iter()
            .filter(|series| self.series_uses_primary_value_axis(series));
        // A shared axis stays absolute when an ordinary series is mixed with
        // a normalized series; only an entirely normalized axis uses percent.
        series.next().is_some_and(|first| {
            first.grouping == ChartGrouping::PercentStacked
                && series.all(|series| series.grouping == ChartGrouping::PercentStacked)
        })
    }

    pub(super) fn legend_entries(&self) -> Vec<(usize, String, u32)> {
        if self
            .series
            .iter()
            .any(|s| matches!(s.kind, ChartKind::Surface | ChartKind::SurfaceWireframe))
        {
            let (min, max, major) = chart_surface_axis(self);
            return (0..((max - min) / major).ceil().clamp(1.0, 100.0) as usize)
                .filter(|i| !self.deleted_legend_entries.contains(i))
                .map(|i| {
                    (
                        i,
                        format!(
                            "{}–{}",
                            format_chart_value(min + i as f32 * major, None),
                            format_chart_value((min + (i + 1) as f32 * major).min(max), None)
                        ),
                        self.series
                            .get(i)
                            .and_then(|s| s.color)
                            .unwrap_or_else(|| super::office_chart_palette_color(i)),
                    )
                })
                .collect();
        }
        let pie = self.legend_point_series();
        let mut entries: Vec<(usize, String, u32)> = match pie {
            Some(series) => series
                .categories
                .iter()
                .enumerate()
                .map(|(index, label)| {
                    (
                        index,
                        label.clone(),
                        series
                            .point_colors
                            .get(index)
                            .copied()
                            .unwrap_or_else(|| super::office_chart_palette_color(index)),
                    )
                })
                .collect(),
            None => self
                .series
                .iter()
                .enumerate()
                .map(|(index, series)| {
                    (
                        index,
                        if series.name.trim().is_empty() {
                            format!("Series {}", index + 1)
                        } else {
                            series.name.clone()
                        },
                        series
                            .color
                            .unwrap_or_else(|| super::office_chart_palette_color(index)),
                    )
                })
                .collect(),
        };
        if pie.is_none()
            && !self.legend_position.horizontal()
            && (self
                .series
                .iter()
                .any(|series| series.kind == ChartKind::Bar && series.bar_horizontal)
                || self.series.iter().all(|series| {
                    matches!(series.kind, ChartKind::Area | ChartKind::Line)
                        && matches!(
                            series.grouping,
                            ChartGrouping::Stacked | ChartGrouping::PercentStacked
                        )
                }))
        {
            entries.reverse();
        }
        entries.retain(|(index, _, _)| !self.deleted_legend_entries.contains(index));
        entries
    }

    fn legend_point_series(&self) -> Option<&ChartSeries> {
        self.series
            .iter()
            .all(|series| {
                matches!(
                    series.kind,
                    ChartKind::Pie
                        | ChartKind::BarOfPie(_)
                        | ChartKind::PieOfPie(_)
                        | ChartKind::Doughnut
                )
            })
            .then(|| self.series.first())
            .flatten()
    }

    pub(super) fn legend_fill(&self, index: usize) -> Option<&ChartFill> {
        if self
            .series
            .iter()
            .any(|s| matches!(s.kind, ChartKind::Surface | ChartKind::SurfaceWireframe))
        {
            return self.surface_band_fills.get(index)?.as_ref();
        }
        self.legend_point_series().map_or_else(
            || self.series.get(index)?.fill.as_ref(),
            |series| series.point_fills.get(index)?.as_ref(),
        )
    }

    pub(super) fn uses_compact_horizontal_3d_layout(&self) -> bool {
        self.show_legend
            && self.legend_bounds.is_none()
            && self.legend_position == ChartLegendPosition::Right
            && self.series.iter().any(|series| {
                series.kind == ChartKind::Bar && series.three_d && series.bar_horizontal
            })
    }

    pub(super) fn legend_key_is_line(&self, index: usize) -> bool {
        self.series.get(index).is_some_and(|series| {
            matches!(
                series.kind,
                ChartKind::Line | ChartKind::Scatter | ChartKind::Radar
            )
        })
    }

    pub(super) fn legend_key_width(&self, index: usize, swatch: f32) -> f32 {
        if self.legend_key_is_line(index) {
            swatch * 3.0
        } else {
            swatch
        }
    }

    pub(super) fn legend_marker(&self, index: usize) -> Option<(&str, f32)> {
        self.legend_key_is_line(index).then_some(())?;
        let series = self.series.get(index)?;
        Some((
            series.marker_symbol.as_deref()?,
            series.marker_size.unwrap_or(8.0),
        ))
    }

    /// Office keeps automatic lateral legends inside the chart and omits entries
    /// that do not fit. Authored and horizontal layouts retain their own rules.
    pub(super) fn constrained_legend(&self, bounds: Rect) -> Option<(Rect, f32, Vec<(usize, String, u32)>)> {
        if self.legend_bounds.is_some() || self.legend_position.horizontal() || !self.show_legend {
            return None;
        }
        let font = self.legend_text_style().font_size;
        let mut entries = self.legend_entries();
        let available = (bounds.height - if self.show_title { 46.0 } else { 16.0 }).max(0.0);
        let width = bounds.width * 0.28;
        let mut lines = 1;
        for (index, label, _) in &mut entries {
            let text_width = (width - self.legend_key_width(*index, font * 0.65) - 18.0).max(font);
            let mut wrapped = String::new();
            let mut line = String::new();
            for word in label.split_whitespace() {
                let candidate = if line.is_empty() { word.to_owned() } else { format!("{line} {word}") };
                if !line.is_empty() && chart_text_width(&candidate, font) > text_width {
                    wrapped.push_str(&line);
                    wrapped.push('\n');
                    line.clear();
                }
                if !line.is_empty() { line.push(' '); }
                line.push_str(word);
            }
            wrapped.push_str(&line);
            lines = lines.max(wrapped.lines().count());
            *label = wrapped;
        }
        let step = font * 1.15 * (lines + 1) as f32;
        if entries.len() as f32 * step + 8.0 <= available { return None; }
        entries.truncate(((available - 8.0) / step).floor().max(0.0) as usize);
        let height = entries.len() as f32 * step + 8.0;
        Some((Rect {
            x: if self.legend_position == ChartLegendPosition::Left { bounds.x + 4.0 }
                else { bounds.x + bounds.width - width - 4.0 },
            y: bounds.y + if self.show_title { 38.0 } else { 0.0 } + (available - height).max(0.0) / 2.0 + 8.0,
            width, height,
        }, step, entries))
    }

    pub(super) fn legend_text_style(&self) -> ChartTextStyle {
        self.style.as_ref().map_or(
            ChartTextStyle {
                color: 0x0000_00ff,
                font_size: self.font_size.unwrap_or(8.5 * 96.0 / 72.0),
                bold: self.font_bold,
                align: TextAlign::Start,
                shadow: None,
            },
            |style| style.legend,
        )
    }

    pub(super) fn title_text(&self) -> &str {
        if self.title.trim().is_empty() {
            self.series
                .first()
                .filter(|series| self.series.len() == 1 && !series.name.trim().is_empty())
                .map_or("Chart Title", |series| series.name.as_str())
        } else {
            &self.title
        }
    }

    pub(super) fn positioned_title_bounds(&self, bounds: Rect, mut title: Rect) -> Rect {
        if let Some([(x, x_edge), (y, y_edge)]) = self.title_position {
            title.x = if x_edge { bounds.x } else { title.x } + bounds.width * x;
            title.y = if y_edge { bounds.y } else { title.y } + bounds.height * y;
        }
        title
    }

    pub(super) fn title_text_style(&self) -> ChartTextStyle {
        let mut style = self.style.as_ref().map_or(
            ChartTextStyle {
                color: if self.classic_defaults { 0x0000_00ff } else { 0x5959_59ff },
                font_size: if self.classic_defaults { 24.0 } else { 14.0 * 96.0 / 72.0 },
                bold: self.classic_defaults,
                align: TextAlign::Center,
                shadow: None,
            },
            |style| style.title,
        );
        style.color = self.title_text_color.unwrap_or(style.color);
        style.font_size = self.title_font_size.unwrap_or(style.font_size);
        style.bold = self.title_font_bold.unwrap_or(style.bold);
        style
    }

    pub(super) fn axis_label_font_size(&self, axis: &ChartValueAxis, fallback: f32) -> f32 {
        axis.label_font_size.or(self.font_size).unwrap_or(fallback)
    }

    pub(super) fn axis_title_text_style(&self, axis: &ChartValueAxis) -> ChartTextStyle {
        ChartTextStyle {
            color: 0x0000_00ff,
            font_size: axis
                .title_font_size
                .or(self.font_size)
                .unwrap_or(10.0 * 96.0 / 72.0),
            bold: axis.title_bold.unwrap_or(self.font_bold || self.classic_defaults),
            align: TextAlign::Center,
            shadow: None,
        }
    }

    pub(super) fn data_label_text_style(
        &self,
        series: &ChartSeries,
        label: Option<&ChartDataLabel>,
    ) -> ChartTextStyle {
        let mut style = self.style.as_ref().map_or(
            ChartTextStyle {
                color: 0x0000_00ff,
                font_size: 10.0 * 96.0 / 72.0,
                bold: false,
                align: TextAlign::Center,
                shadow: None,
            },
            |style| style.data_label,
        );
        style.color = label
            .and_then(|label| label.text_color)
            .or(series.data_label_text_color)
            .unwrap_or(style.color);
        style.font_size = label
            .and_then(|label| label.font_size)
            .or(series.data_label_font_size)
            .unwrap_or(style.font_size);
        style.bold = label
            .and_then(|label| label.font_bold)
            .or(series.data_label_font_bold)
            .unwrap_or(style.bold);
        style
    }

    pub(super) fn column_view_3d(&self) -> ChartView3D {
        let mut view = self.view_3d.unwrap_or_default();
        if view.depth_percent.is_none()
            && let Some(series) = self.series.iter().find(|s| s.projected_column() && !s.bar_depth)
        {
            // Clustered columns occupy one row in depth: its default depth is one
            // column width plus the depth gap, rather than a full category axis.
            let count = self.series.iter().filter(|s| s.kind == ChartKind::Bar && s.axis_id == series.axis_id).count();
            let columns = series.values.len().max(1) as f32 * (count as f32 + self.bar_gap_width_percent.clamp(0.0, 500.0) / 100.0);
            view.depth_percent = Some(((1.0 + series.bar_gap_depth_percent / 100.0) / columns * 100.0).round().max(1.0) as u16);
        }
        view
    }

    pub(super) fn projected_value_title_bounds(&self, plot: Rect, font_size: f32) -> Option<Rect> {
        let bottom = self.depth_axis_point(plot, 0.0, 0.0, 0.0)?;
        let top = self.depth_axis_point(plot, 0.0, 1.0, 0.0)?;
        Some(Rect { x: (bottom.0 + top.0) / 2.0 - font_size * 1.8,
            y: top.1.min(bottom.1), width: font_size * 1.5, height: (bottom.1 - top.1).abs() })
    }

    pub(super) fn resolve_value_axis_layout(&mut self, bounds: Rect) {
        let projected = self.series.iter().any(ChartSeries::projected_column);
        if !projected && !(self.classic_defaults && bounds.height < 240.0) { return; }
        if !self.series.iter().all(|s| matches!(s.kind, ChartKind::Bar | ChartKind::Line | ChartKind::Area | ChartKind::Scatter | ChartKind::Bubble | ChartKind::Radar)) { return; }
        let scatter = self.series.iter().all(|s| matches!(s.kind, ChartKind::Scatter | ChartKind::Bubble));
        let plot = self.plot_area_bounds(bounds, Rect {
            x: bounds.x + bounds.width * 0.1, y: bounds.y + bounds.height * 0.16,
            width: bounds.width * if scatter && self.show_legend { 0.62 } else { 0.84 }, height: bounds.height * 0.66,
        });
        for horizontal in [false, true] {
            if horizontal && !scatter { continue; }
            let axis = if horizontal { &self.horizontal_axis_options } else { &self.value_axis_options };
            if axis.minimum.is_some() || axis.maximum.is_some() || axis.major_unit.is_some()
                || axis.log_base.is_some() { continue; }
            let bottom = self.depth_axis_point(plot, 0.0, 0.0, 0.0).map_or(plot.y + plot.height, |p| p.1);
            let top = self.depth_axis_point(plot, 0.0, 1.0, 0.0).map_or(plot.y, |p| p.1);
            let span = if horizontal { plot.width } else if self.series.iter().all(|s| s.kind == ChartKind::Radar) {
                plot.width.min(plot.height) / 2.0
            } else { (bottom - top).abs() };
            let intervals = (span / (self.axis_label_font_size(axis, 12.0)
                * if horizontal { 4.0 } else if projected { 2.5 } else { 1.5 })).floor().clamp(1.0, 7.0);
            let (min, max, old_step) = if horizontal { self.scatter_axis(true) } else { self.value_axis() };
            let mut step = nice_chart_step((max - min) / intervals);
            if step < (max - min) / intervals { step = nice_chart_step(step * 2.0); }
            if step > old_step {
                let axis = if horizontal { &mut self.horizontal_axis_options } else { &mut self.value_axis_options };
                axis.minimum = Some((min / step).floor() * step);
                axis.maximum = Some((max / step).ceil() * step);
                axis.major_unit = Some(step);
            }
        }
    }

    pub(super) fn value_axis(&self) -> (f32, f32, f32) {
        if !self.series.is_empty()
            && self
                .series
                .iter()
                .all(|series| matches!(series.kind, ChartKind::Scatter | ChartKind::Bubble))
        {
            return self.scatter_axis(false);
        }
        if self.value_axis_is_percent() {
            let has_negative = self
                .series
                .iter()
                .filter(|series| self.series_uses_primary_value_axis(series))
                .flat_map(|series| &series.values)
                .any(|value| value.is_finite() && *value < 0.0);
            return (if has_negative { -1.0 } else { 0.0 }, 1.0, 0.1);
        }
        let axis = self.numerical_axis_options();
        if let Some(base) = axis.log_base {
            return Self::log_axis_range(
                axis,
                self.series
                    .iter()
                    .filter(|series| self.series_uses_primary_value_axis(series))
                    .flat_map(ChartSeries::value_spans)
                    .flat_map(|(start, end)| [start, end]),
            )
            .unwrap_or((1.0, base, base));
        }
        let stacked = self
            .series
            .iter()
            .filter(|series| self.series_uses_primary_value_axis(series))
            .any(|series| series.grouping == ChartGrouping::Stacked);
        let (mut minimum, mut maximum) = if stacked {
            let category_count = self
                .series
                .iter()
                .filter(|series| self.series_uses_primary_value_axis(series))
                .map(|series| series.values.len())
                .max()
                .unwrap_or(0);
            let stacked_bounds =
                (0..category_count).fold((0.0_f32, 0.0_f32), |(minimum, maximum), index| {
                    let (negative, positive) = self
                        .series
                        .iter()
                        .filter(|series| self.series_uses_primary_value_axis(series))
                        .filter(|series| series.grouping == ChartGrouping::Stacked)
                        .filter_map(|series| series.values.get(index).copied())
                        .filter(|value| value.is_finite())
                        .fold((0.0_f32, 0.0_f32), |(negative, positive), value| {
                            if value < 0.0 {
                                (negative + value, positive)
                            } else {
                                (negative, positive + value)
                            }
                        });
                    (minimum.min(negative), maximum.max(positive))
                });
            self.series
                .iter()
                .filter(|series| self.series_uses_primary_value_axis(series))
                .filter(|series| series.grouping != ChartGrouping::Stacked)
                .flat_map(ChartSeries::value_spans)
                .flat_map(|(start, end)| [start, end])
                .filter(|value| value.is_finite())
                .fold(stacked_bounds, |(minimum, maximum), value| {
                    (minimum.min(value), maximum.max(value))
                })
        } else {
            (
                self.series
                    .iter()
                    .filter(|series| self.series_uses_primary_value_axis(series))
                    .flat_map(ChartSeries::value_spans)
                    .flat_map(|(start, end)| [start, end])
                    .filter(|value| value.is_finite())
                    .fold(0.0_f32, f32::min),
                self.series
                    .iter()
                    .filter(|series| self.series_uses_primary_value_axis(series))
                    .flat_map(ChartSeries::value_spans)
                    .flat_map(|(start, end)| [start, end])
                    .filter(|value| value.is_finite())
                    .fold(0.0_f32, f32::max),
            )
        };
        if (maximum - minimum).abs() <= f32::EPSILON {
            maximum = maximum.max(1.0);
            minimum = minimum.min(0.0);
        }
        if axis.minimum.is_some() || axis.maximum.is_some() || axis.major_unit.is_some() {
            let raw_minimum = minimum;
            let raw_maximum = maximum;
            let span = match (axis.minimum, axis.maximum) {
                (Some(explicit_minimum), Some(explicit_maximum)) => {
                    (explicit_maximum - explicit_minimum).abs()
                }
                (Some(explicit_minimum), None) => (raw_maximum - explicit_minimum).abs(),
                (None, Some(explicit_maximum)) => (explicit_maximum - raw_minimum).abs(),
                (None, None) => (raw_maximum - raw_minimum).abs(),
            };
            let step = axis
                .major_unit
                .unwrap_or_else(|| nice_chart_step(span / 7.0));
            let (minimum, maximum) = match (axis.minimum, axis.maximum) {
                (Some(minimum), Some(maximum)) => (minimum, maximum),
                (Some(minimum), None) => {
                    let mut maximum = minimum + step * 7.0;
                    while maximum < raw_maximum {
                        maximum += step;
                    }
                    (minimum, maximum)
                }
                (None, Some(maximum)) => {
                    let mut minimum = maximum - step * 7.0;
                    while minimum > raw_minimum {
                        minimum -= step;
                    }
                    (minimum, maximum)
                }
                (None, None) => (
                    (raw_minimum / step).floor() * step,
                    (raw_maximum / step).ceil() * step,
                ),
            };
            return (minimum, maximum.max(minimum + step), step);
        }
        automatic_chart_axis_range(
            minimum,
            maximum,
            chart_axis_needs_auto_headroom(
                self.series
                    .iter()
                    .filter(|series| self.series_uses_primary_value_axis(series)),
            ),
        )
    }

    pub(super) fn value_axis_ratio(&self, value: f32, minimum: f32, maximum: f32) -> Option<f32> {
        if self.numerical_axis_options().log_base.is_some() {
            (value > 0.0 && minimum > 0.0 && maximum > minimum).then(|| {
                ((value.ln() - minimum.ln()) / (maximum.ln() - minimum.ln())).clamp(0.0, 1.0)
            })
        } else {
            (maximum > minimum).then(|| ((value - minimum) / (maximum - minimum)).clamp(0.0, 1.0))
        }
    }

    pub(super) fn value_axis_ticks(&self) -> Vec<(f32, f32)> {
        let (minimum, maximum, major) = self.value_axis();
        if let Some(base) = self.numerical_axis_options().log_base {
            let first = minimum.log(base).ceil() as i32;
            let last = maximum.log(base).floor() as i32;
            return (first..=last)
                .filter_map(|exponent| {
                    let value = base.powi(exponent);
                    self.value_axis_ratio(value, minimum, maximum)
                        .map(|ratio| (value, ratio))
                })
                .collect();
        }
        chart_linear_axis_ticks((minimum, maximum, major))
    }

    pub(super) fn scatter_axis_ticks(&self, horizontal: bool) -> Vec<(f32, f32)> {
        chart_linear_axis_ticks(self.scatter_axis(horizontal))
    }
}

fn chart_linear_axis_ticks((minimum, maximum, major): (f32, f32, f32)) -> Vec<(f32, f32)> {
    let mut ticks = Vec::new();
    let mut value = minimum;
    while value <= maximum + major * 0.01 && ticks.len() < 1_000 {
        ticks.push((value, (value - minimum) / (maximum - minimum)));
        value += major;
    }
    ticks
}


fn chart_axis_needs_auto_headroom<'a>(mut series: impl Iterator<Item = &'a ChartSeries>) -> bool {
    let Some(first) = series.next() else {
        return true;
    };
    !std::iter::once(first).chain(series).all(|series| {
        series.kind == ChartKind::Bar
            && series.three_d
            && (series.bar_horizontal || (!series.bar_cone && !series.bar_cylinder))
            && series.grouping == ChartGrouping::Standard
    })
}

fn automatic_chart_axis_range(
    mut minimum: f32,
    mut maximum: f32,
    add_headroom: bool,
) -> (f32, f32, f32) {
    let raw_maximum = maximum;
    let mut step = nice_chart_step((maximum - minimum) / 7.0);
    minimum = (minimum / step).floor() * step;
    maximum = (maximum / step).ceil() * step;
    if (maximum - minimum) / step >= 10.0 - f32::EPSILON {
        step = nice_chart_step(step * 2.0);
        minimum = (minimum / step).floor() * step;
        maximum = (raw_maximum / step).ceil() * step;
    }
    if add_headroom
        && raw_maximum > 0.0
        && maximum - raw_maximum < step * 0.5
        && (maximum - minimum) / step < 10.0
    {
        maximum += step;
    }
    (minimum, maximum.max(minimum + step), step)
}

pub(super) type ChartPoint = (f32, f32, usize);

pub(super) fn format_axis_value(value: f32) -> String {
    if value.abs() < 0.000_000_1 {
        "0".to_owned()
    } else if value.abs() >= 1e10 {
        let scientific = format!("{value:.5e}");
        let (mantissa, exponent) = scientific.split_once('e').unwrap();
        format!("{}E{:+03}", mantissa.trim_end_matches('0').trim_end_matches('.'), exponent.parse::<i32>().unwrap_or_default())
    } else {
        format!("{value:.2}")
            .trim_end_matches('0')
            .trim_end_matches('.')
            .to_owned()
    }
}

pub(super) fn format_chart_value(value: f32, number_format: Option<&str>) -> String {
    if number_format.is_some_and(|format| format.contains('%')) {
        let decimals = number_format
            .and_then(|f| f.split_once('.'))
            .map_or(0, |(_, f)| {
                f.chars().take_while(|c| matches!(c, '0' | '#')).count()
            })
            .min(12);
        format!("{:.*}%", decimals, value * 100.0)
    } else if number_format.is_some_and(|format| format.contains('$')) {
        let digits = value.abs().round().min(i64::MAX as f32) as i64;
        let grouped = super::group_decimal_digits(&digits.to_string());
        if value < 0.0 && number_format.is_some_and(|format| format.contains('(')) {
            format!("(${grouped})")
        } else if value < 0.0 {
            format!("-${grouped}")
        } else {
            format!("${grouped}")
        }
    } else if let Some(format) = number_format.filter(|format| format.contains(",")) {
        let decimals = format
            .split_once('.')
            .map_or(0, |(_, fraction)| {
                fraction
                    .chars()
                    .take_while(|c| matches!(c, '0' | '#'))
                    .count()
            })
            .min(12);
        super::group_decimal_digits(&format!("{value:.decimals$}"))
    } else {
        format_axis_value(value)
    }
}

pub(super) fn chart_data_table_layout(
    chart: &Chart,
    chart_bounds: Rect,
    plot: Rect,
) -> Option<ChartDataTableLayout> {
    let table = chart.data_table?;
    let category_count = chart
        .series
        .iter()
        .map(|series| series.categories.len().max(series.values.len()))
        .max()
        .unwrap_or(0);
    if category_count == 0 || chart.series.is_empty() {
        return None;
    }
    let name_width = chart.series.iter().map(|series| chart_text_width(&series.name, 13.333_333))
        .fold(0.0_f32, f32::max) + if table.show_keys { 36.0 } else { 8.0 };
    let left = (chart_bounds.x + chart_bounds.width * 0.03 + name_width)
        .max(plot.x).min(plot.x + plot.width * 0.4);
    let plot = Rect { x: left, width: plot.x + plot.width - left, ..plot };
    let font_size = chart.font_size.unwrap_or(13.333_333);
    let projected = chart.series.iter().any(ChartSeries::projected_column);
    let table_left = if projected { plot.x - (name_width - 24.0) } else { chart_bounds.x + chart_bounds.width * 0.03 };
    let table_top = plot.y + plot.height + if projected { font_size } else { 0.0 };
    let table_bounds = Rect {
        x: table_left,
        y: table_top,
        width: plot.x + plot.width - table_left,
        height: (chart_bounds.y + chart_bounds.height * 0.98 - table_top)
            .max(0.0)
            .min((chart.series.len() + 1) as f32 * font_size * 1.6),
    };
    if table_bounds.height <= 0.0 {
        return None;
    }
    let row_count = chart.series.len() + 1;
    let row_height = table_bounds.height / row_count as f32;
    let category_width = plot.width / category_count as f32;
    let mut lines = Vec::new();
    if table.show_horizontal_borders {
        lines.extend((1..row_count).map(|row| ChartDataTableLine {
            bounds: Rect {
                x: table_bounds.x,
                y: table_bounds.y + row as f32 * row_height,
                width: table_bounds.width,
                height: 0.01,
            },
            geometry: Geometry::Line,
            color: 0x7f7f_7fff,
            width: 0.75,
            series_index: None,
        }));
    }
    if table.show_vertical_borders {
        lines.extend((usize::from(!table.show_horizontal_borders)..category_count).map(|column| ChartDataTableLine {
            bounds: Rect {
                x: plot.x + column as f32 * category_width,
                y: table_bounds.y,
                width: 0.01,
                height: table_bounds.height,
            },
            geometry: Geometry::Line,
            color: 0x7f7f_7fff,
            width: 0.75,
            series_index: None,
        }));
    }
    if table.show_outline {
        lines.push(ChartDataTableLine {
            bounds: table_bounds,
            geometry: Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo {
                        x: 0.0,
                        y: row_height,
                    },
                    PathCommand::LineTo {
                        x: 0.0,
                        y: table_bounds.height,
                    },
                    PathCommand::LineTo {
                        x: table_bounds.width,
                        y: table_bounds.height,
                    },
                    PathCommand::LineTo {
                        x: table_bounds.width,
                        y: 0.0,
                    },
                ],
            },
            color: 0x7f7f_7fff,
            width: 0.75,
            series_index: None,
        });
    }
    let mut texts = Vec::new();
    if let Some(header) = chart
        .series
        .iter()
        .max_by_key(|series| series.categories.len())
        .map(|series| series.categories.as_slice())
    {
        texts.extend(
            header
                .iter()
                .take(category_count)
                .enumerate()
                .map(|(index, category)| ChartDataTableText {
                    bounds: Rect {
                        x: plot.x + index as f32 * category_width,
                        y: table_bounds.y,
                        width: category_width,
                        height: row_height,
                    },
                    text: category.clone(),
                    align: TextAlign::Center,
                    series_index: None,
                }),
        );
    }
    for (series_index, series) in chart.series.iter().enumerate() {
        let y = table_bounds.y + (series_index + 1) as f32 * row_height;
        let key_width = if table.show_keys { if projected { 12.0 } else { 28.0 } } else { 4.0 };
        if table.show_keys {
            let square = matches!(series.kind, ChartKind::Bar | ChartKind::Area | ChartKind::Pie | ChartKind::BarOfPie(_) | ChartKind::Doughnut);
            lines.push(ChartDataTableLine {
                bounds: Rect {
                    x: table_bounds.x + if projected { 0.0 } else { 4.0 },
                    y: y + (row_height - if square { 7.0 } else { 0.0 }) / 2.0,
                    width: if square { 7.0 } else { 20.0 },
                    height: if square { 7.0 } else { 0.01 },
                },
                geometry: if square { Geometry::Rectangle } else { Geometry::Line },
                color: series
                    .color
                    .unwrap_or_else(|| super::office_chart_palette_color(series_index)),
                width: if square { 0.0 } else { 3.0 },
                series_index: Some(series_index),
            });
        }
        texts.push(ChartDataTableText {
            bounds: Rect {
                x: table_bounds.x + key_width,
                y,
                width: (plot.x - table_bounds.x - key_width).max(0.0),
                height: row_height,
            },
            text: series.name.clone(),
            align: TextAlign::Start,
            series_index: Some(series_index),
        });
        texts.extend(
            series
                .values
                .iter()
                .copied()
                .enumerate()
                .map(|(category_index, value)| ChartDataTableText {
                    bounds: Rect {
                        x: plot.x + category_index as f32 * category_width,
                        y,
                        width: category_width,
                        height: row_height,
                    },
                    text: series.value_label(category_index, value, series.number_format.as_deref()),
                    align: TextAlign::Center,
                    series_index: Some(series_index),
                }),
        );
    }
    Some(ChartDataTableLayout { lines, texts })
}

pub(super) fn chart_area_geometry(
    chart: &Chart,
    target_index: usize,
    plot: Rect,
    (minimum, maximum): (f32, f32),
) -> Option<(Geometry, Vec<ChartPoint>, Vec<ChartPoint>)> {
    let all_series = &chart.series;
    let target = all_series.get(target_index)?;
    if target.kind != ChartKind::Area || target.values.len() < 2 {
        return None;
    }
    let matching = all_series
        .iter()
        .enumerate()
        .filter(|(_, series)| {
            series.kind == ChartKind::Area
                && series.axis_id == target.axis_id
                && series.grouping == target.grouping
        })
        .collect::<Vec<_>>();
    let target_order = matching
        .iter()
        .position(|(index, _)| *index == target_index)?;
    let value_at = |series: &ChartSeries, category: usize| {
        series
            .values
            .get(category)
            .copied()
            .filter(|value| value.is_finite())
            .unwrap_or(0.0)
    };
    let mut lower = Vec::with_capacity(target.values.len());
    let mut upper = Vec::with_capacity(target.values.len());
    for category in 0..target.values.len() {
        let value = value_at(target, category);
        let same_sign = |candidate: f32| {
            if value < 0.0 {
                candidate < 0.0
            } else {
                candidate >= 0.0
            }
        };
        let stacked = matching
            .iter()
            .take(target_order)
            .map(|(_, series)| value_at(series, category))
            .filter(|candidate| same_sign(*candidate))
            .sum::<f32>();
        let denominator = if target.grouping == ChartGrouping::PercentStacked {
            matching
                .iter()
                .map(|(_, series)| value_at(series, category))
                .filter(|candidate| same_sign(*candidate))
                .map(f32::abs)
                .sum::<f32>()
                .max(f32::EPSILON)
        } else {
            1.0
        };
        let base = if target.grouping == ChartGrouping::Standard {
            0.0
        } else {
            stacked / denominator
        };
        lower.push(base);
        upper.push(base + value / denominator);
    }
    if maximum - minimum <= f32::EPSILON {
        return None;
    }
    let y = |value: f32| {
        plot.height - plot.height * ((value - minimum) / (maximum - minimum)).clamp(0.0, 1.0)
    };
    let x = |index: usize| chart.category_index(index, target.values.len()) as f32 / (target.values.len() - 1) as f32 * plot.width;
    let upper_points = upper
        .iter()
        .enumerate()
        .map(|(index, value)| (plot.x + x(index), plot.y + y(*value), index))
        .collect::<Vec<_>>();
    let label_points = upper
        .iter()
        .zip(&lower)
        .enumerate()
        .map(|(index, (upper, lower))| {
            (plot.x + x(index), plot.y + y((upper + lower) / 2.0), index)
        })
        .collect::<Vec<_>>();
    let mut commands = Vec::with_capacity(upper.len() + lower.len() + 2);
    commands.push(PathCommand::MoveTo {
        x: x(0),
        y: y(lower[0]),
    });
    commands.extend(
        upper
            .iter()
            .enumerate()
            .map(|(index, value)| PathCommand::LineTo {
                x: x(index),
                y: y(*value),
            }),
    );
    commands.extend(
        lower
            .iter()
            .enumerate()
            .rev()
            .map(|(index, value)| PathCommand::LineTo {
                x: x(index),
                y: y(*value),
            }),
    );
    commands.push(PathCommand::ClosePath);
    Some((
        Geometry::Path {
            fill_rule: FillRule::NonZero,
            commands,
        },
        upper_points,
        label_points,
    ))
}

pub(super) fn chart_area_3d_faces(
    chart: &Chart,
    target_index: usize,
    plot: Rect,
    (minimum, maximum): (f32, f32),
    view: ChartView3D,
    color: u32,
) -> Vec<(Rect, Geometry, Paint)> {
    let all_series = &chart.series;
    let Some(target) = all_series
        .get(target_index)
        .filter(|series| series.kind == ChartKind::Area && series.three_d)
    else {
        return Vec::new();
    };
    if target.values.len() < 2 || maximum - minimum <= f32::EPSILON {
        return Vec::new();
    }
    let matching = all_series
        .iter()
        .enumerate()
        .filter(|(_, series)| {
            series.kind == ChartKind::Area && series.three_d && series.axis_id == target.axis_id
        })
        .collect::<Vec<_>>();
    let Some(series_slot) = matching
        .iter()
        .position(|(index, _)| *index == target_index)
    else {
        return Vec::new();
    };
    let stacked = target.grouping != ChartGrouping::Standard;
    let series_slot = if stacked { 0 } else { series_slot };
    let depth_count = if stacked { 1 } else { matching.len().max(1) };
    if !view.right_angle_axes {
        let unit = Rect {
            x: 0.0,
            y: 0.0,
            width: 1.0,
            height: 1.0,
        };
        let Some((Geometry::Path { commands, .. }, upper, _)) =
            chart_area_geometry(chart, target_index, unit, (minimum, maximum))
        else {
            return Vec::new();
        };
        let count = depth_count as f32;
        let z0 = series_slot as f32 / count;
        let z1 = z0 + 1.0 / count / (1.0 + target.bar_gap_depth_percent / 100.0);
        let project = |x, y, z| {
            let (x, y) = chart_project_3d(plot, view, x, 1.0 - y, z);
            (x - plot.x, y - plot.y)
        };
        let outline = commands
            .iter()
            .filter_map(|c| match *c {
                PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => Some((x, y)),
                _ => None,
            })
            .collect::<Vec<_>>();
        let near_z = if f32::from(view.rot_y).to_radians().cos() >= 0.0 {
            z0
        } else {
            z1
        };
        let face = super::polygon_geometry(
            &outline
                .iter()
                .map(|&(x, y)| project(x, y, near_z))
                .collect::<Vec<_>>(),
        );
        let mut top = upper
            .iter()
            .map(|&(x, y, _)| project(x, y, z0))
            .collect::<Vec<_>>();
        top.extend(upper.iter().rev().map(|&(x, y, _)| project(x, y, z1)));
        let index = if f32::from(view.rot_y).to_radians().sin() >= 0.0 {
            target.values.len() - 1
        } else {
            0
        };
        let upper_point = outline[index + 1];
        let lower_point = outline[outline.len() - 1 - index];
        let side = super::polygon_geometry(&[
            project(upper_point.0, upper_point.1, z0),
            project(upper_point.0, upper_point.1, z1),
            project(lower_point.0, lower_point.1, z1),
            project(lower_point.0, lower_point.1, z0),
        ]);
        let mut top_color = color;
        apply_color_transform(&mut top_color, "shade", 0.75);
        let mut side_color = color;
        apply_color_transform(&mut side_color, "shade", 0.50);
        return vec![
            (plot, face, Paint::Solid(color)),
            (plot, super::polygon_geometry(&top), Paint::Solid(top_color)),
            (plot, side, Paint::Solid(side_color)),
        ];
    }
    let (depth_x, depth_y) = chart_area_3d_offset(plot, view);
    let front_width = (plot.width - depth_x).max(1.0);
    let slot_count = depth_count as f32;
    let depth = series_slot as f32 / slot_count;
    let projected_plot = Rect {
        x: plot.x + depth_x * depth,
        y: plot.y - depth_y * depth,
        width: front_width,
        height: plot.height,
    };
    let Some((area, upper, _)) =
        chart_area_geometry(chart, target_index, projected_plot, (minimum, maximum))
    else {
        return Vec::new();
    };
    let extrusion = 0.70 / slot_count;
    let extrusion_x = depth_x * extrusion;
    let extrusion_y = -depth_y * extrusion;
    let mut top_color = color;
    apply_color_transform(&mut top_color, "tint", 0.22);
    let mut front_color = color;
    apply_color_transform(&mut front_color, "shade", 0.72);
    let polygon = |points: &[(f32, f32)]| {
        let min_x = points
            .iter()
            .map(|point| point.0)
            .fold(f32::INFINITY, f32::min);
        let min_y = points
            .iter()
            .map(|point| point.1)
            .fold(f32::INFINITY, f32::min);
        let max_x = points
            .iter()
            .map(|point| point.0)
            .fold(f32::NEG_INFINITY, f32::max);
        let max_y = points
            .iter()
            .map(|point| point.1)
            .fold(f32::NEG_INFINITY, f32::max);
        let bounds = Rect {
            x: min_x,
            y: min_y,
            width: (max_x - min_x).max(0.01),
            height: (max_y - min_y).max(0.01),
        };
        let local = points
            .iter()
            .map(|&(x, y)| (x - bounds.x, y - bounds.y))
            .collect::<Vec<_>>();
        (bounds, super::polygon_geometry(&local))
    };
    let mut faces = vec![(projected_plot, area, Paint::Solid(color))];
    let mut top = upper.iter().map(|(x, y, _)| (*x, *y)).collect::<Vec<_>>();
    top.extend(
        upper
            .iter()
            .rev()
            .map(|(x, y, _)| (*x + extrusion_x, *y + extrusion_y)),
    );
    let (top_bounds, top_geometry) = polygon(&top);
    faces.push((top_bounds, top_geometry, Paint::Solid(top_color)));
    let Geometry::Path { commands, .. } = &faces[0].1 else {
        return faces;
    };
    let Some(PathCommand::LineTo {
        x: lower_x,
        y: lower_y,
    }) = commands.get(target.values.len() + 1)
    else {
        return faces;
    };
    let Some((upper_x, upper_y, _)) = upper.last() else {
        return faces;
    };
    let side = [
        (*upper_x, *upper_y),
        (*upper_x + extrusion_x, *upper_y + extrusion_y),
        (
            projected_plot.x + lower_x + extrusion_x,
            projected_plot.y + lower_y + extrusion_y,
        ),
        (projected_plot.x + lower_x, projected_plot.y + lower_y),
    ];
    let (side_bounds, side_geometry) = polygon(&side);
    faces.push((side_bounds, side_geometry, Paint::Solid(front_color)));
    faces
}

pub(super) fn chart_area_3d_axis_labels(
    chart: &Chart,
    plot: Rect,
    font_size: f32,
) -> (
    Vec<(usize, Rect, String, f32)>,
    Vec<(usize, Rect, String, f32)>,
) {
    let areas = chart
        .series
        .iter()
        .enumerate()
        .filter(|(_, series)| series.kind == ChartKind::Area && series.three_d)
        .collect::<Vec<_>>();
    let Some((_, first)) = areas.first() else {
        return (Vec::new(), Vec::new());
    };
    let view = chart.view_3d.unwrap_or_default();
    if !view.right_angle_axes {
        let count = areas.len() as f32;
        let near_x = if f32::from(view.rot_y).to_radians().sin() >= 0.0 {
            1.0
        } else {
            0.0
        };
        let series_labels = if chart
            .series_axis_options
            .as_ref()
            .is_some_and(|axis| axis.deleted)
        {
            Vec::new()
        } else {
            areas
                .iter()
                .enumerate()
                .map(|(slot, (index, series))| {
                    let (x, y) =
                        chart_project_3d(plot, view, near_x, 0.0, (slot as f32 + 0.5) / count);
                    (
                        *index,
                        Rect {
                            x: x - font_size * 2.0,
                            y: y + 4.0,
                            width: font_size * 4.0,
                            height: font_size * 1.25,
                        },
                        series.name.clone(),
                        -90.0,
                    )
                })
                .collect()
        };
        let categories = first
            .categories
            .iter()
            .enumerate()
            .map(|(index, name)| {
                let x = chart.category_index(index, first.categories.len()) as f32
                    / (first.categories.len() - 1).max(1) as f32;
                let (x, y) = chart_project_3d(plot, view, x, 0.0, 1.0);
                let rotation = chart.category_axis_options().label_rotation_degrees
                    .unwrap_or(if chart.dense_category_labels(first, plot) { -90.0 } else { 0.0 });
                let width = chart_text_width(name, font_size) + 4.0;
                (
                    index,
                    Rect {
                        x: if rotation == -90.0 { x - width / 2.0 } else { x + 4.0 },
                        y: if rotation == -90.0 { y + width / 2.0 } else { y - font_size * 0.625 },
                        width,
                        height: font_size * 1.25,
                    },
                    name.clone(),
                    rotation,
                )
            })
            .collect();
        return (series_labels, categories);
    }
    let (depth_x, depth_y) = chart_area_3d_offset(plot, view);
    let front_width = (plot.width - depth_x).max(1.0);
    let series_labels = chart
        .series_axis_options
        .as_ref()
        .filter(|options| !options.deleted)
        .map_or_else(Vec::new, |options| {
            let count = areas.len().max(1) as f32;
            areas
                .iter()
                .enumerate()
                .map(|(slot, (series_index, series))| {
                    let depth = slot as f32 / count;
                    (
                        *series_index,
                        Rect {
                            x: plot.x + front_width + depth_x * depth - font_size * 1.75,
                            y: plot.y + plot.height - depth_y * depth + 4.0,
                            width: (font_size * 3.5).max(40.0),
                            height: font_size * 1.25,
                        },
                        if series.name.trim().is_empty() {
                            (slot + 1).to_string()
                        } else {
                            series.name.clone()
                        },
                        options.label_rotation_degrees.unwrap_or(-45.0),
                    )
                })
                .collect()
        });
    let category_width = front_width / first.categories.len().max(1) as f32;
    let label_width = (font_size * 6.0).max(category_width);
    let category_plot = Rect {
        width: front_width,
        ..plot
    };
    let category_labels = first
        .categories
        .iter()
        .enumerate()
        .map(|(index, category)| {
            let center_x = chart
                .category_x(first, index, category_plot, false)
                .unwrap_or(plot.x + category_width * (index as f32 + 0.5));
            let rotation = chart.category_axis_options().label_rotation_degrees
                .unwrap_or(if chart.dense_category_labels(first, plot) { -90.0 } else { -45.0 });
            let label_width = if rotation == -90.0 { chart_text_width(category, font_size) + 4.0 } else { label_width };
            (
                index,
                Rect {
                    x: center_x - label_width / 2.0,
                    y: plot.y + plot.height + 2.0 + if rotation == -90.0 { (label_width - font_size * 1.25) / 2.0 } else { 0.0 },
                    width: label_width,
                    height: (plot.height * 0.10).max(font_size * 1.25),
                },
                category.clone(),
                rotation,
            )
        })
        .collect();
    (series_labels, category_labels)
}

pub(super) fn radar_geometry(values: &[f32], maximum: f32, bounds: Rect) -> Option<Geometry> {
    if values.len() < 3 || !maximum.is_finite() || maximum <= 0.0 {
        return None;
    }
    let center_x = bounds.width / 2.0;
    let center_y = bounds.height / 2.0;
    let radius = bounds.width.min(bounds.height) * 0.42;
    let mut commands = Vec::with_capacity(values.len() + 1);
    for (index, value) in values.iter().copied().enumerate() {
        let angle = -std::f32::consts::FRAC_PI_2
            + std::f32::consts::TAU * index as f32 / values.len() as f32;
        let point_radius = radius * (value / maximum).clamp(0.0, 1.0);
        let point = (
            center_x + point_radius * angle.cos(),
            center_y + point_radius * angle.sin(),
        );
        commands.push(if index == 0 {
            PathCommand::MoveTo {
                x: point.0,
                y: point.1,
            }
        } else {
            PathCommand::LineTo {
                x: point.0,
                y: point.1,
            }
        });
    }
    commands.push(PathCommand::ClosePath);
    Some(Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    })
}

pub(super) fn chart_bubble_bounds(
    chart: &Chart,
    target_index: usize,
    plot: Rect,
) -> Vec<(Rect, usize)> {
    let Some(target) = chart
        .series
        .get(target_index)
        .filter(|series| series.kind == ChartKind::Bubble)
    else {
        return Vec::new();
    };
    let bubble_series = chart
        .series
        .iter()
        .filter(|series| series.kind == ChartKind::Bubble);
    let (x_min, x_max, _) = chart.scatter_axis(true);
    let (y_min, y_max, _) = chart.scatter_axis(false);
    let size_max = bubble_series
        .flat_map(|series| &series.bubble_sizes)
        .copied()
        .filter(|value| value.is_finite() && *value > 0.0)
        .fold(0.0_f32, f32::max);
    if x_max <= x_min || y_max <= y_min || size_max <= 0.0 {
        return Vec::new();
    }
    let max_radius = plot.width.min(plot.height) * 0.09;
    target
        .x_values
        .iter()
        .copied()
        .zip(target.values.iter().copied())
        .zip(target.bubble_sizes.iter().copied())
        .enumerate()
        .filter(|(_, ((x, y), size))| {
            x.is_finite() && y.is_finite() && size.is_finite() && *size > 0.0
        })
        .map(|(index, ((x, y), size))| {
            let radius = (max_radius * (size / size_max).sqrt()).max(2.0);
            let center_x = plot.x + (x - x_min) / (x_max - x_min) * plot.width;
            let center_y = plot.y + (1.0 - (y - y_min) / (y_max - y_min)) * plot.height;
            (
                Rect {
                    x: center_x - radius,
                    y: center_y - radius,
                    width: radius * 2.0,
                    height: radius * 2.0,
                },
                index,
            )
        })
        .collect()
}

pub(super) fn chart_bubble_paint(bounds: Rect, color: u32, three_d: bool) -> Paint {
    if !three_d {
        return Paint::Solid((color & 0xffff_ff00) | 0xb3);
    }
    Paint::CircleGradient {
        x0: bounds.width * 0.28,
        y0: bounds.height * 0.25,
        r0: 0.0,
        x1: bounds.width / 2.0,
        y1: bounds.height / 2.0,
        r1: bounds.width.max(bounds.height) * 0.65,
        stops: vec![
            GradientStop {
                offset: 0.0,
                color: transform_luminance(color, 0.35, 0.65),
            },
            GradientStop {
                offset: 0.35,
                color,
            },
            GradientStop {
                offset: 1.0,
                color: transform_luminance(color, 0.38, 0.0),
            },
        ],
    }
}

impl ChartValueAxis {
    fn gridline_properties_mut(&mut self, minor: bool) -> (&mut bool, &mut Option<u32>, &mut Option<f32>) {
        if minor { (&mut self.minor_gridlines, &mut self.minor_gridline_color, &mut self.minor_gridline_width) }
        else { (&mut self.major_gridlines, &mut self.major_gridline_color, &mut self.major_gridline_width) }
    }

    fn minor_ticks(&self, range: (f32, f32, f32)) -> Vec<(f32, f32)> {
        let (minimum, maximum, major) = range;
        if maximum <= minimum || (!self.minor_gridlines && self.minor_tick_mark == ChartAxisTickMark::None) { return Vec::new(); }
        if let Some(base) = self.log_base {
            if minimum <= 0.0 { return Vec::new(); }
            let mut ticks = Vec::new();
            for exponent in (minimum.log(base).floor() as i32)..=(maximum.log(base).ceil() as i32) {
                for multiplier in 2..(base.ceil() as u32).min(1000) {
                    let value = base.powi(exponent) * multiplier as f32;
                    if value > minimum && value < maximum {
                        ticks.push((value, (value.ln() - minimum.ln()) / (maximum.ln() - minimum.ln())));
                        if ticks.len() == 1000 { return ticks; }
                    }
                }
            }
            return ticks;
        }
        let minor = self.minor_unit.unwrap_or(major / 5.0);
        if !minor.is_finite() || minor <= 0.0 || major <= 0.0 { return Vec::new(); }
        chart_linear_axis_ticks((minimum, maximum, minor)).into_iter()
            .filter(|(value, ratio)| *ratio > 0.0 && *ratio < 1.0 && {
                let major_index = (*value - minimum) / major;
                (major_index - major_index.round()).abs() > 0.0001
            }).collect()
    }
}

impl Chart {
    /// Minor-axis geometry is shared by all OOXML chart hosts.
    pub(super) fn minor_axis_lines(&self, plot: Rect) -> Vec<(Rect, Geometry, u32, f32)> {
        let horizontal_bars = self.series.iter().any(|series| series.bar_horizontal);
        let scatter = self.series.iter().any(|series| matches!(series.kind, ChartKind::Scatter | ChartKind::Bubble));
        let three_d = self.series.iter().any(|series| series.kind == ChartKind::Bar && series.three_d);
        let mut lines = Vec::new();
        for (horizontal, axis, secondary) in [(false, &self.value_axis_options, false), (true, &self.horizontal_axis_options, false)]
            .into_iter().chain(self.secondary_value_axis_options.iter().map(|axis| (false, axis, true))) {
            if axis.deleted || (!scatter && horizontal != horizontal_bars) { continue; }
            let range = if secondary { self.axis_range(axis) } else if scatter { self.scatter_axis(horizontal) } else { self.value_axis() };
            for (_, ratio) in axis.minor_ticks(range) {
                if axis.minor_gridlines {
                    let (bounds, geometry) = self.bar_grid_line(plot, ratio, three_d, horizontal);
                    lines.push((bounds, geometry, axis.minor_gridline_color.unwrap_or(0xe6e6_e6ff), axis.minor_gridline_width.unwrap_or(0.5)));
                }
                let size = 3.0;
                let far = matches!(axis.position.as_deref(), Some("r" | "t"));
                let offset = match axis.minor_tick_mark {
                    ChartAxisTickMark::None => continue,
                    ChartAxisTickMark::In => if far { -size } else { 0.0 },
                    ChartAxisTickMark::Out => if far { 0.0 } else { -size },
                    ChartAxisTickMark::Cross => -size / 2.0,
                };
                let bounds = if horizontal {
                    Rect { x: plot.x + ratio * plot.width, y: if far { plot.y - offset - size } else { plot.y + plot.height - offset - size }, width: 0.01, height: size }
                } else {
                    Rect { x: if far { plot.x + plot.width + offset } else { plot.x + offset }, y: plot.y + (1.0 - ratio) * plot.height, width: size, height: 0.01 }
                };
                lines.push((bounds, Geometry::Line, 0x5959_59ff, 0.75));
            }
        }
        if let Some(series) = self.series.first()
            && let Some(marks) = self.date_axis_marks(series, plot) {
            let axis = self.category_axis_options();
            if !axis.deleted {
                for (x, _, major) in marks {
                    let (gridlines, color, width, tick) = if major {
                        (axis.major_gridlines, axis.major_gridline_color.unwrap_or(0xd9d9_d9ff),
                         axis.major_gridline_width.unwrap_or(0.75), axis.major_tick_mark)
                    } else {
                        (axis.minor_gridlines, axis.minor_gridline_color.unwrap_or(0xe6e6_e6ff),
                         axis.minor_gridline_width.unwrap_or(0.5), axis.minor_tick_mark)
                    };
                    if gridlines {
                        lines.push((Rect { x, y: plot.y, width: 0.01, height: plot.height },
                            Geometry::Line, color, width));
                    }
                    let (y, height) = match tick {
                        ChartAxisTickMark::None => continue,
                        ChartAxisTickMark::In => (plot.y + plot.height - 4.0, 4.0),
                        ChartAxisTickMark::Out => (plot.y + plot.height, 4.0),
                        ChartAxisTickMark::Cross => (plot.y + plot.height - 2.0, 4.0),
                    };
                    lines.push((Rect { x, y, width: 0.01, height }, Geometry::Line, 0x5959_59ff, 0.75));
                }
            }
        }
        lines
    }
}

impl ChartValueAxis {
    pub(super) fn format_label(&self, value: f32) -> String {
        let scaled = value / self.display_unit.unwrap_or(1.0);
        if self.display_unit.is_some() && scaled != 0.0 && scaled.abs() < 0.01
            && self.number_format.as_deref().is_none_or(|format| format.eq_ignore_ascii_case("General")) {
            format!("{scaled:E}")
        } else {
            format_chart_value(scaled, self.number_format.as_deref())
        }
    }
}

impl Chart {
    /// Extra labels use identical geometry in DOCX, PPTX and XLSX adapters.
    pub(super) fn supplemental_axis_labels(&self, plot: Rect, font_size: f32) -> Vec<(Rect, String)> {
        let mut labels = Vec::new();
        let horizontal = self.series.iter().any(|s| s.bar_horizontal);
        for axis in [&self.value_axis_options, &self.horizontal_axis_options].into_iter()
            .chain(self.secondary_value_axis_options.iter()) {
            if axis.deleted { continue; }
            if let (Some(unit), Some(label)) = (axis.display_unit, &axis.display_unit_label) {
                let text = if !label.is_empty() { label.clone() } else {
                    match unit {
                        100.0 => "Hundreds".to_owned(), 1e3 => "Thousands".to_owned(),
                        1e4 => "Ten thousands".to_owned(), 1e5 => "Hundred thousands".to_owned(),
                        1e6 => "Millions".to_owned(), 1e7 => "Ten millions".to_owned(),
                        1e8 => "Hundred millions".to_owned(), 1e9 => "Billions".to_owned(),
                        1e12 => "Trillions".to_owned(), _ => format_chart_value(unit, None),
                    }
                };
                labels.push((Rect { x: plot.x, y: plot.y - font_size * 1.7,
                    width: plot.width, height: font_size * 1.4 }, text));
            }
        }
        let axis = self.category_axis_options();
        if !axis.deleted
            && let Some(series) = self.series.first()
            && let Some(marks) = self.date_axis_marks(series, plot) {
            let major_count = marks.iter().filter(|(_, _, major)| *major).count().max(1);
            let width = (plot.width / major_count as f32).max(font_size * 3.0);
            labels.extend(marks.into_iter().filter(|(_, _, major)| *major).map(|(x, text, _)| {
                (Rect { x: x - width / 2.0, y: plot.y + plot.height + 4.0,
                    width, height: font_size * 1.4 }, text)
            }));
        }
        if axis.deleted || axis.no_multi_level_labels { return labels; }
        let Some(series) = self.series.first() else { return labels; };
        let count = series.categories.len().max(series.values.len());
        if count == 0 { return labels; }
        let centered = matches!(series.kind, ChartKind::Bar);
        for (level_index, level) in series.category_levels.iter().enumerate().skip(1) {
            for (start, text) in level.iter().enumerate().filter(|(i, text)| *i < count && !text.is_empty()) {
                let end = level.iter().enumerate().skip(start + 1)
                    .find(|(_, text)| !text.is_empty()).map_or(count, |(i, _)| i.min(count));
                let bounds = if horizontal {
                    let slot = plot.height / count as f32;
                    Rect { x: plot.x - font_size * (5.0 + level_index as f32 * 4.0),
                        y: plot.y + (if axis.reversed { start } else { count - end }) as f32 * slot, width: font_size * 4.0,
                        height: (end - start) as f32 * slot }
                } else {
                    let first = self.category_x(series, start, plot, centered).unwrap_or(plot.x);
                    let last = self.category_x(series, end.saturating_sub(1), plot, centered).unwrap_or(first);
                    let width = (last - first).abs() + plot.width / count as f32;
                    Rect { x: (first + last - width) / 2.0,
                        y: plot.y + plot.height + 4.0 + font_size * 1.5 * level_index as f32,
                        width, height: font_size * 1.4 }
                };
                labels.push((bounds, text.clone()));
            }
        }
        labels
    }
}
