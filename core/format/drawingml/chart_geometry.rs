/// A line series is one stroked path: joins and effects belong to the series,
/// not to every adjacent pair of cached points.
pub(super) fn chart_line_geometry(points: &[ChartPoint]) -> Option<(Rect, Geometry)> {
    if points.len() < 2 { return None; }
    let (mut left, mut top, _) = points[0];
    let (mut right, mut bottom) = (left, top);
    for &(x, y, _) in &points[1..] {
        left = left.min(x); top = top.min(y);
        right = right.max(x); bottom = bottom.max(y);
    }
    let bounds = Rect { x: left, y: top, width: (right - left).max(0.01), height: (bottom - top).max(0.01) };
    let local = points.iter().map(|&(x, y, _)| (x-left, y-top)).collect::<Vec<_>>();
    Some((bounds, extended_chart_path(&local, false)))
}

impl Chart {
    pub(super) fn scatter_line_geometry(&self, series: &ChartSeries, plot: Rect) -> Option<(Rect, Geometry)> {
        if !self.scatter_has_lines || !series.line_visible { return None; }
        let points = self.scatter_points(series, plot);
        if !series.smooth || points.len() < 3 { return chart_line_geometry(&points); }
        let local = points.iter().map(|&(x, y, _)| (x - plot.x, y - plot.y)).collect::<Vec<_>>();
        let mut commands = vec![PathCommand::MoveTo { x: local[0].0, y: local[0].1 }];
        // ponytail: Catmull-Rom interpolation preserves cached points; native Office
        // spline tension may differ and needs a separate measured curve oracle.
        for index in 1..local.len() {
            let previous = local[index.saturating_sub(2)];
            let start = local[index - 1];
            let end = local[index];
            let next = local[(index + 1).min(local.len() - 1)];
            commands.push(PathCommand::BezierCurveTo {
                cp1x: start.0 + (end.0 - previous.0) / 6.0,
                cp1y: start.1 + (end.1 - previous.1) / 6.0,
                cp2x: end.0 - (next.0 - start.0) / 6.0,
                cp2y: end.1 - (next.1 - start.1) / 6.0,
                x: end.0, y: end.1,
            });
        }
        Some((plot, Geometry::Path { fill_rule: FillRule::NonZero, commands }))
    }
}

pub(super) fn chart_bar_bounds(
    plot: Rect,
    value: f32,
    value_axis: (f32, f32),
    category: (usize, usize),
    series: (usize, usize),
    horizontal: bool,
    group_ratio: f32,
) -> Rect {
    let (minimum, maximum) = value_axis;
    let ratio = |value: f32| ((value - minimum) / (maximum - minimum)).clamp(0.0, 1.0);
    if horizontal {
        let category_height = plot.height / category.1.max(1) as f32;
        let bar_height = category_height * group_ratio / series.1.max(1) as f32;
        let baseline = plot.x + plot.width * ratio(0.0);
        let value_x = plot.x + plot.width * ratio(value);
        Rect {
            x: baseline.min(value_x),
            y: plot.y
                + category.1.saturating_sub(category.0 + 1) as f32 * category_height
                + category_height * (1.0 - group_ratio) / 2.0
                + series.1.saturating_sub(series.0 + 1) as f32 * bar_height,
            width: (baseline - value_x).abs().max(0.01),
            height: bar_height,
        }
    } else {
        let category_width = plot.width / category.1.max(1) as f32;
        let bar_width = category_width * group_ratio / series.1.max(1) as f32;
        let baseline = plot.y + plot.height - plot.height * ratio(0.0);
        let value_y = plot.y + plot.height - plot.height * ratio(value);
        Rect {
            x: plot.x
                + category.0 as f32 * category_width
                + category_width * (1.0 - group_ratio) / 2.0
                + series.0 as f32 * bar_width,
            y: baseline.min(value_y),
            width: bar_width,
            height: (baseline - value_y).abs().max(0.01),
        }
    }
}

pub(super) fn chart_3d_series_axis_labels(
    chart: &Chart,
    plot: Rect,
    font_size: f32,
) -> Vec<(usize, Rect, String)> {
    if chart
        .series_axis_options
        .as_ref()
        .is_none_or(|options| options.deleted)
    {
        return Vec::new();
    }
    chart
        .series
        .iter()
        .enumerate()
        .filter(|(_, series)| {
            series.kind == ChartKind::Bar && series.three_d && !series.bar_horizontal
        })
        .filter_map(|(series_index, series)| {
            let (category_index, value) = series
                .values
                .iter()
                .copied()
                .enumerate()
                .find(|(_, value)| value.is_finite() && *value != 0.0)?;
            let (minimum, maximum, _) = chart.value_axis_for_series(series);
            let bar = chart_bar_segment_bounds(
                &chart,
                series_index,
                category_index,
                plot,
                (minimum, maximum),
                chart.bar_gap_width_percent,
            )?;
            let baseline = if value >= 0.0 {
                bar.y + bar.height
            } else {
                bar.y
            };
            Some((
                series_index,
                Rect {
                    x: if series.bar_depth {
                        chart_project_3d(plot, chart.view_3d.unwrap_or_default(), 1.0, 0.0, (series_index as f32+0.5)/chart.series.len() as f32).0 + 8.0
                    } else { plot.x + plot.width + 8.0 },
                    y: if series.bar_depth {
                        chart_project_3d(plot, chart.view_3d.unwrap_or_default(), 1.0, 0.0, (series_index as f32+0.5)/chart.series.len() as f32).1 - font_size * 0.625
                    } else { baseline - font_size * 0.625 },
                    width: plot.width * 0.22,
                    height: font_size * 1.25,
                },
                if series.name.trim().is_empty() {
                    (series_index + 1).to_string()
                } else {
                    series.name.clone()
                },
            ))
        })
        .collect()
}

// Project normalized category, value and series coordinates through the authored
// view. Fit the complete plot volume, preserving its height/depth proportions.
pub(super) fn chart_project_3d(
    plot: Rect,
    view: ChartView3D,
    x: f32,
    y: f32,
    z: f32,
) -> (f32, f32) {
    let (sx, cx) = f32::from(view.rot_x).to_radians().sin_cos();
    let (sy, cy) = f32::from(view.rot_y).to_radians().sin_cos();
    let height = view.height_percent.map_or(plot.height / plot.width.max(1.0), |v| f32::from(v).max(5.0) / 100.0);
    let depth = f32::from(view.depth_percent.unwrap_or(100)).max(20.0) / 100.0;
    let project = |x: f32, y: f32, z: f32| {
        let (x, y, z) = (x - 0.5, (y - 0.5) * height, (z - 0.5) * depth);
        let distance = (z * cy - x * sy) * cx + y * sx;
        let perspective = if view.right_angle_axes {
            0.0
        } else {
            f32::from(view.perspective) / 240.0
        };
        let scale = 1.0 / (1.0 + distance * perspective).max(0.1);
        (
            (x * cy + z * sy) * scale,
            (-y * cx + (x * sy - z * cy) * sx) * scale,
        )
    };
    let (mut left, mut top, mut right, mut bottom) = (
        f32::INFINITY,
        f32::INFINITY,
        f32::NEG_INFINITY,
        f32::NEG_INFINITY,
    );
    for x in [0.0, 1.0] {
        for y in [0.0, 1.0] {
            for z in [0.0, 1.0] {
                let (u, v) = project(x, y, z);
                left = left.min(u);
                right = right.max(u);
                top = top.min(v);
                bottom = bottom.max(v);
            }
        }
    }
    let scale = (plot.width / (right - left).max(0.01)).min(plot.height / (bottom - top).max(0.01));
    let (u, v) = project(x, y, z);
    (
        plot.x + plot.width / 2.0 + (u - (left + right) / 2.0) * scale,
        plot.y + plot.height / 2.0 + (v - (top + bottom) / 2.0) * scale,
    )
}

// All faces use the same camera as the floor and the axes.  A screen-space
// rectangle plus a fixed extrusion cannot represent a perspective cuboid.
impl Chart {
    pub(super) fn depth_box_faces(
        &self,
        series_index: usize,
        category: usize,
        plot: Rect,
        origin: Rect,
    ) -> Option<[Geometry; 3]> {
        let target = self.series.get(series_index)?;
        if !target.projected_column() || target.bar_cone || target.bar_cylinder {
            return None;
        }
        let matching = self
            .series
            .iter()
            .enumerate()
            .filter(|(_, s)| s.kind == ChartKind::Bar && s.axis_id == target.axis_id)
            .collect::<Vec<_>>();
        let slot = matching.iter().position(|(i, _)| *i == series_index)?;
        let count = matching.iter().map(|(_, s)| s.values.len()).max()?.max(1) as f32;
        let mut x = (self.category_index(category, count as usize) as f32 + 0.5) / count;
        let mut dx = 0.5 / count / (1.0 + self.bar_gap_width_percent.clamp(0.0, 500.0) / 100.0);
        if !target.bar_depth {
            dx = 0.5 / count / (matching.len() as f32 + self.bar_gap_width_percent.clamp(0.0, 500.0) / 100.0);
            x += dx * (slot as f32 * 2.0 + 1.0 - matching.len() as f32);
        }
        let z = if target.bar_depth { (slot as f32 + 0.5) / matching.len() as f32 } else { 0.5 };
        let dz = 0.5
            / if target.bar_depth { matching.len() as f32 } else { 1.0 }
            / (1.0 + target.bar_gap_depth_percent.clamp(0.0, 500.0) / 100.0);
        let (min, max, _) = self.value_axis_for_series(target);
        if max <= min {
            return None;
        }
        let ratio = |v: f32| ((v - min) / (max - min)).clamp(0.0, 1.0);
        let (bottom, top) = (
            ratio(0.0).min(ratio(target.values[category])),
            ratio(0.0).max(ratio(target.values[category])),
        );
        let view = self.column_view_3d();
        let front = if f32::from(view.rot_y).to_radians().cos() >= 0.0 {
            z - dz
        } else {
            z + dz
        };
        let side = if f32::from(view.rot_y).to_radians().sin() >= 0.0 {
            x + dx
        } else {
            x - dx
        };
        let cap = if view.rot_x >= 0 { top } else { bottom };
        let face = |points: [(f32, f32, f32); 4]| {
            super::polygon_geometry(&points.map(|(x, y, z)| {
                let (x, y) = chart_project_3d(plot, view, x, y, z);
                (x - origin.x, y - origin.y)
            }))
        };
        Some([
            face([
                (x - dx, bottom, front),
                (x + dx, bottom, front),
                (x + dx, top, front),
                (x - dx, top, front),
            ]),
            face([
                (x - dx, cap, z - dz),
                (x + dx, cap, z - dz),
                (x + dx, cap, z + dz),
                (x - dx, cap, z + dz),
            ]),
            face([
                (side, bottom, z - dz),
                (side, bottom, z + dz),
                (side, top, z + dz),
                (side, top, z - dz),
            ]),
        ])
    }

    pub(super) fn bar_faces(
        &self,
        series: usize,
        category: usize,
        plot: Rect,
        bounds: Rect,
        color: u32,
    ) -> [(Rect, Geometry, Paint); 2] {
        if let Some([_, top, side]) = self.depth_box_faces(series, category, plot, bounds) {
            let mut top_color = color;
            let mut side_color = color;
            apply_color_transform(&mut top_color, "shade", 0.75);
            apply_color_transform(&mut side_color, "shade", 0.60);
            [
                (bounds, top, Paint::Solid(top_color)),
                (bounds, side, Paint::Solid(side_color)),
            ]
        } else {
            chart_bar_3d_faces(
                bounds,
                color,
                self.view_3d.unwrap_or_default(),
                self.series[series].bar_horizontal,
            )
        }
    }
}

pub(super) fn chart_bar_segment_bounds(
    chart: &Chart,
    target_index: usize,
    category_index: usize,
    plot: Rect,
    value_axis: (f32, f32),
    gap_width_percent: f32,
) -> Option<Rect> {
    let all_series = &chart.series;
    let target = all_series
        .get(target_index)
        .filter(|series| series.kind == ChartKind::Bar)?;
    let value = *target.values.get(category_index)?;
    if !value.is_finite() || value == 0.0 {
        return None;
    }
    let bar_series = all_series
        .iter()
        .enumerate()
        .filter(|(_, series)| series.kind == ChartKind::Bar && series.axis_id == target.axis_id)
        .collect::<Vec<_>>();
    let category_count = bar_series
        .iter()
        .map(|(_, series)| series.values.len())
        .max()
        .unwrap_or(0);
    let category_slot = chart.category_index(category_index, category_count);
    let series_slot = bar_series
        .iter()
        .position(|(index, _)| *index == target_index)?;
    let gap_width = gap_width_percent.clamp(0.0, 500.0) / 100.0;
    if target.projected_column() {
        if let Some(faces) = chart.depth_box_faces(target_index, category_index, plot, plot) {
            let mut left = f32::INFINITY;
            let mut top = f32::INFINITY;
            let mut right = f32::NEG_INFINITY;
            let mut bottom = f32::NEG_INFINITY;
            for face in faces {
                if let Geometry::Path { commands, .. } = face {
                    for command in commands {
                        if let PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } = command {
                            left = left.min(x); right = right.max(x);
                            top = top.min(y); bottom = bottom.max(y);
                        }
                    }
                }
            }
            return Some(Rect { x: plot.x+left, y: plot.y+top, width: (right-left).max(0.01), height: (bottom-top).max(0.01) });
        }
        let view = chart.view_3d.unwrap_or_default();
        let x = (category_slot as f32 + 0.5) / category_count.max(1) as f32;
        let z = (series_slot as f32 + 0.5) / bar_series.len().max(1) as f32;
        let ratio = |v: f32| ((v - value_axis.0) / (value_axis.1 - value_axis.0)).clamp(0.0, 1.0);
        let (base_x, base_y) = chart_project_3d(plot, view, x, ratio(0.0), z);
        let (_, top_y) = chart_project_3d(plot, view, x, ratio(value), z);
        let half = 0.5 / category_count.max(1) as f32 / (1.0 + gap_width);
        let (left, _) = chart_project_3d(plot, view, x - half, ratio(0.0), z);
        let (right, _) = chart_project_3d(plot, view, x + half, ratio(0.0), z);
        let width = (right - left).abs().max(0.01);
        return Some(Rect {
            x: base_x - width / 2.0,
            y: base_y.min(top_y),
            width,
            height: (base_y - top_y).abs().max(0.01),
        });
    }
    let group_slots = if target.grouping == ChartGrouping::Standard {
        bar_series.len() as f32
    } else {
        1.0
    };
    let group_ratio = group_slots / (group_slots + gap_width);
    if target.grouping == ChartGrouping::Standard {
        return Some(chart_bar_bounds(
            plot,
            value,
            value_axis,
            (category_slot, category_count),
            (series_slot, bar_series.len()),
            target.bar_horizontal,
            group_ratio,
        ));
    }
    let same_sign = |candidate: f32| candidate.is_finite() && candidate * value > 0.0;
    let start = bar_series[..series_slot]
        .iter()
        .filter_map(|(_, series)| series.values.get(category_index).copied())
        .filter(|candidate| same_sign(*candidate))
        .sum::<f32>();
    let scale = if target.grouping == ChartGrouping::PercentStacked {
        bar_series
            .iter()
            .filter_map(|(_, series)| series.values.get(category_index).copied())
            .filter(|candidate| same_sign(*candidate))
            .map(f32::abs)
            .sum::<f32>()
    } else {
        1.0
    };
    if scale <= f32::EPSILON {
        return None;
    }
    let start = start / scale;
    let end = start + value / scale;
    let ratio = |candidate: f32| {
        ((candidate - value_axis.0) / (value_axis.1 - value_axis.0)).clamp(0.0, 1.0)
    };
    if target.bar_horizontal {
        let category_height = plot.height / category_count.max(1) as f32;
        let start_x = plot.x + plot.width * ratio(start);
        let end_x = plot.x + plot.width * ratio(end);
        Some(Rect {
            x: start_x.min(end_x),
            y: plot.y
                + category_count.saturating_sub(category_slot + 1) as f32 * category_height
                + category_height * (1.0 - group_ratio) / 2.0,
            width: (start_x - end_x).abs().max(0.01),
            height: category_height * group_ratio,
        })
    } else {
        let category_width = plot.width / category_count.max(1) as f32;
        let start_y = plot.y + plot.height - plot.height * ratio(start);
        let end_y = plot.y + plot.height - plot.height * ratio(end);
        Some(Rect {
            x: plot.x
                + category_slot as f32 * category_width
                + category_width * (1.0 - group_ratio) / 2.0,
            y: start_y.min(end_y),
            width: category_width * group_ratio,
            height: (start_y - end_y).abs().max(0.01),
        })
    }
}

pub(super) fn chart_series_line_geometries(
    chart: &Chart,
    plot: Rect,
    value_axis: (f32, f32),
) -> Vec<(Rect, Geometry)> {
    if !chart.series_lines {
        return Vec::new();
    }
    let mut lines = Vec::new();
    for (series_index, series) in chart.series.iter().enumerate().filter(|(_, series)| {
        series.kind == ChartKind::Bar && series.grouping != ChartGrouping::Standard
    }) {
        for category_index in 0..series.values.len().saturating_sub(1) {
            let (Some(left), Some(right)) = (
                chart_bar_segment_bounds(
                    &chart,
                    series_index,
                    category_index,
                    plot,
                    value_axis,
                    chart.bar_gap_width_percent,
                ),
                chart_bar_segment_bounds(
                    &chart,
                    series_index,
                    category_index + 1,
                    plot,
                    value_axis,
                    chart.bar_gap_width_percent,
                ),
            ) else {
                continue;
            };
            let (start, end) = if series.bar_horizontal {
                (
                    (left.x + left.width, if left.y < right.y { left.y + left.height } else { left.y }),
                    (right.x + right.width, if left.y < right.y { right.y } else { right.y + right.height }),
                )
            } else {
                ((if left.x > right.x { left.x } else { left.x + left.width }, left.y),
                 (if left.x > right.x { right.x + right.width } else { right.x }, right.y))
            };
            let bounds = Rect {
                x: start.0.min(end.0),
                y: start.1.min(end.1),
                width: (end.0 - start.0).abs().max(0.01),
                height: (end.1 - start.1).abs().max(0.01),
            };
            lines.push((
                bounds,
                Geometry::Path {
                    fill_rule: FillRule::NonZero,
                    commands: vec![
                        PathCommand::MoveTo {
                            x: start.0 - bounds.x,
                            y: start.1 - bounds.y,
                        },
                        PathCommand::LineTo {
                            x: end.0 - bounds.x,
                            y: end.1 - bounds.y,
                        },
                    ],
                },
            ));
        }
    }
    lines
}

pub(super) fn chart_bar_label_bounds(series: &ChartSeries, bar: Rect, value: f32) -> Rect {
    chart_bar_label_bounds_at(series, bar, value, series.data_label_position.as_deref())
}

fn chart_bar_label_bounds_at(
    series: &ChartSeries,
    bar: Rect,
    value: f32,
    position: Option<&str>,
) -> Rect {
    let position = position.unwrap_or(if series.grouping == ChartGrouping::Standard {
        "outEnd"
    } else {
        "ctr"
    });
    let outside = position == "outEnd";
    if series.bar_horizontal {
        let width = if outside {
            48.0
        } else if series.data_label_position.is_some() {
            (bar.width * 0.75).clamp(18.0, 48.0).min(bar.width)
        } else {
            48.0_f32.min(bar.width)
        };
        let height = 18.0;
        let center = match position {
            "inBase" if value < 0.0 => bar.x + bar.width - width / 2.0,
            "inBase" => bar.x + width / 2.0,
            "inEnd" if value < 0.0 => bar.x + width / 2.0,
            "inEnd" => bar.x + bar.width - width / 2.0,
            "outEnd" if value < 0.0 => bar.x - width / 2.0 - 2.0,
            "outEnd" => bar.x + bar.width + width / 2.0 + 2.0,
            _ => bar.x + bar.width / 2.0,
        };
        Rect {
            x: center - width / 2.0,
            y: bar.y + (bar.height - height) / 2.0,
            width,
            height,
        }
    } else {
        let width = 48.0;
        let height = if outside {
            18.0
        } else if series.data_label_position.is_some() {
            (bar.height * 0.75).clamp(18.0, 48.0).min(bar.height)
        } else {
            18.0_f32.min(bar.height)
        };
        let center = match position {
            "inBase" if value < 0.0 => bar.y + height / 2.0,
            "inBase" => bar.y + bar.height - height / 2.0,
            "inEnd" if value < 0.0 => bar.y + bar.height - height / 2.0,
            "inEnd" => bar.y + height / 2.0,
            "outEnd" if value < 0.0 => bar.y + bar.height + height / 2.0 + 2.0,
            "outEnd" => bar.y - height / 2.0 - 2.0,
            _ => bar.y + bar.height / 2.0,
        };
        Rect {
            x: bar.x + (bar.width - width) / 2.0,
            y: center - height / 2.0,
            width,
            height,
        }
    }
}

pub(super) fn chart_bar_custom_label_bounds(
    series: &ChartSeries,
    bar: Rect,
    label: &ChartDataLabel,
    font_size: f32,
    value: f32,
) -> Rect {
    let width = (label.text.chars().count() as f32 * font_size * 0.58 + 8.0).max(48.0);
    let height = font_size * 1.25 + 4.0;
    if series.bar_horizontal {
        Rect {
            x: if value >= 0.0 {
                bar.x + bar.width + 2.0
            } else {
                bar.x - width - 2.0
            },
            y: bar.y + (bar.height - height) / 2.0,
            width,
            height,
        }
    } else {
        Rect {
            x: bar.x + (bar.width - width) / 2.0,
            y: if value >= 0.0 {
                bar.y - height - 2.0
            } else {
                bar.y + bar.height + 2.0
            },
            width,
            height,
        }
    }
}

pub(super) fn chart_bar_geometry(
    bounds: Rect,
    horizontal: bool,
    cone: bool,
    cylinder: bool,
) -> Geometry {
    if !cone && !cylinder {
        return Geometry::Rectangle;
    }
    if cylinder {
        let radius = if horizontal {
            (bounds.height * 0.18).min(bounds.width / 4.0)
        } else {
            (bounds.width * 0.18).min(bounds.height / 4.0)
        };
        return Geometry::Path {
            fill_rule: FillRule::NonZero,
            commands: if horizontal {
                vec![
                    PathCommand::MoveTo { x: radius, y: 0.0 },
                    PathCommand::BezierCurveTo {
                        cp1x: 0.0,
                        cp1y: 0.0,
                        cp2x: 0.0,
                        cp2y: bounds.height,
                        x: radius,
                        y: bounds.height,
                    },
                    PathCommand::LineTo {
                        x: bounds.width - radius,
                        y: bounds.height,
                    },
                    PathCommand::BezierCurveTo {
                        cp1x: bounds.width,
                        cp1y: bounds.height,
                        cp2x: bounds.width,
                        cp2y: 0.0,
                        x: bounds.width - radius,
                        y: 0.0,
                    },
                    PathCommand::ClosePath,
                ]
            } else {
                vec![
                    PathCommand::MoveTo { x: 0.0, y: radius },
                    PathCommand::BezierCurveTo {
                        cp1x: 0.0,
                        cp1y: 0.0,
                        cp2x: bounds.width,
                        cp2y: 0.0,
                        x: bounds.width,
                        y: radius,
                    },
                    PathCommand::LineTo {
                        x: bounds.width,
                        y: bounds.height - radius,
                    },
                    PathCommand::BezierCurveTo {
                        cp1x: bounds.width,
                        cp1y: bounds.height,
                        cp2x: 0.0,
                        cp2y: bounds.height,
                        x: 0.0,
                        y: bounds.height - radius,
                    },
                    PathCommand::ClosePath,
                ]
            },
        };
    }
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands: if horizontal {
            vec![
                PathCommand::MoveTo { x: 0.0, y: 0.0 },
                PathCommand::LineTo {
                    x: bounds.width,
                    y: bounds.height / 2.0,
                },
                PathCommand::LineTo {
                    x: 0.0,
                    y: bounds.height,
                },
                PathCommand::ClosePath,
            ]
        } else {
            vec![
                PathCommand::MoveTo {
                    x: bounds.width / 2.0,
                    y: 0.0,
                },
                PathCommand::LineTo {
                    x: bounds.width,
                    y: bounds.height,
                },
                PathCommand::LineTo {
                    x: 0.0,
                    y: bounds.height,
                },
                PathCommand::ClosePath,
            ]
        },
    }
}

pub(super) fn chart_bar_cone_to_max_geometry(
    bounds: Rect,
    horizontal: bool,
    positive: bool,
    value_ratio: f32,
) -> Geometry {
    let inset = if horizontal {
        bounds.height
    } else {
        bounds.width
    } * value_ratio.clamp(0.0, 1.0)
        / 2.0;
    let commands = if horizontal {
        let (base_x, tip_x) = if positive {
            (0.0, bounds.width)
        } else {
            (bounds.width, 0.0)
        };
        vec![
            PathCommand::MoveTo { x: base_x, y: 0.0 },
            PathCommand::LineTo { x: tip_x, y: inset },
            PathCommand::LineTo {
                x: tip_x,
                y: bounds.height - inset,
            },
            PathCommand::LineTo {
                x: base_x,
                y: bounds.height,
            },
            PathCommand::ClosePath,
        ]
    } else {
        let (base_y, tip_y) = if positive {
            (bounds.height, 0.0)
        } else {
            (0.0, bounds.height)
        };
        vec![
            PathCommand::MoveTo { x: 0.0, y: base_y },
            PathCommand::LineTo { x: inset, y: tip_y },
            PathCommand::LineTo {
                x: bounds.width - inset,
                y: tip_y,
            },
            PathCommand::LineTo {
                x: bounds.width,
                y: base_y,
            },
            PathCommand::ClosePath,
        ]
    };
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    }
}

impl ChartSeries {
    pub(super) fn bar_paint(&self, index: usize, bounds: Rect, fallback_color: u32) -> Paint {
        let color = self
            .point_colors
            .get(index)
            .copied()
            .unwrap_or(fallback_color);
        let fill = self
            .point_fills
            .get(index)
            .and_then(Option::as_ref)
            .or(self.fill.as_ref());
        if self.bar_cone || self.bar_cylinder {
            match fill {
                Some(ChartFill::Solid(color)) => chart_bar_paint(
                    bounds,
                    *color,
                    self.bar_horizontal,
                    self.bar_cone,
                    self.bar_cylinder,
                ),
                Some(fill) => fill.paint(bounds),
                None => chart_bar_paint(
                    bounds,
                    color,
                    self.bar_horizontal,
                    self.bar_cone,
                    self.bar_cylinder,
                ),
            }
        } else {
            fill.map_or(Paint::Solid(color), |fill| fill.paint(bounds))
        }
    }
}

pub(super) fn chart_bar_paint(
    bounds: Rect,
    color: u32,
    horizontal: bool,
    cone: bool,
    cylinder: bool,
) -> Paint {
    if !cone && !cylinder {
        return Paint::Solid(color);
    }
    let mut edge = color;
    apply_color_transform(&mut edge, "shade", 0.68);
    let mut center = color;
    apply_color_transform(&mut center, "tint", 0.88);
    Paint::LinearGradient {
        x0: 0.0,
        y0: 0.0,
        x1: if horizontal { 0.0 } else { bounds.width },
        y1: if horizontal { bounds.height } else { 0.0 },
        stops: vec![
            GradientStop {
                offset: 0.0,
                color: edge,
            },
            GradientStop {
                offset: 0.5,
                color: center,
            },
            GradientStop {
                offset: 1.0,
                color: edge,
            },
        ],
    }
}

pub(super) fn chart_bar_cylinder_cap(
    bounds: Rect,
    color: u32,
    horizontal: bool,
    positive: bool,
) -> (Rect, Geometry, Paint) {
    let mut cap_color = color;
    apply_color_transform(&mut cap_color, "shade", 0.78);
    let cap = if horizontal {
        let width = (bounds.height * 0.36).min(bounds.width);
        Rect {
            x: if positive {
                bounds.x + bounds.width - width
            } else {
                bounds.x
            },
            y: bounds.y,
            width,
            height: bounds.height,
        }
    } else {
        let height = (bounds.width * 0.36).min(bounds.height);
        Rect {
            x: bounds.x,
            y: if positive {
                bounds.y
            } else {
                bounds.y + bounds.height - height
            },
            width: bounds.width,
            height,
        }
    };
    (cap, Geometry::Ellipse, Paint::Solid(cap_color))
}

pub(super) fn chart_bar_cone_cap(
    bounds: Rect,
    color: u32,
    horizontal: bool,
    positive: bool,
    value_ratio: f32,
) -> Option<(Rect, Geometry, Paint)> {
    let remaining = 1.0 - value_ratio.clamp(0.0, 1.0);
    if remaining <= 0.02 {
        return None;
    }
    let mut cap_color = color;
    apply_color_transform(&mut cap_color, "shade", 0.78);
    let cap = if horizontal {
        let height = bounds.height * remaining;
        let width = (height * 0.32).min(bounds.width);
        Rect {
            x: if positive {
                bounds.x + bounds.width - width / 2.0
            } else {
                bounds.x - width / 2.0
            },
            y: bounds.y + (bounds.height - height) / 2.0,
            width,
            height,
        }
    } else {
        let width = bounds.width * remaining;
        let height = (width * 0.32).min(bounds.height);
        Rect {
            x: bounds.x + (bounds.width - width) / 2.0,
            y: if positive {
                bounds.y - height / 2.0
            } else {
                bounds.y + bounds.height - height / 2.0
            },
            width,
            height,
        }
    };
    Some((cap, Geometry::Ellipse, Paint::Solid(cap_color)))
}

pub(super) fn chart_3d_offset(plot: Rect, view: ChartView3D) -> (f32, f32) {
    let projection = if view.right_angle_axes { 1.0 } else { 0.85 };
    let default_depth = if view.right_angle_axes {
        (0.12, 0.15)
    } else {
        (0.50, 0.40)
    };
    let (x_factor, y_factor) = view.depth_percent.map_or(default_depth, |depth| {
        let scale = f32::from(depth.clamp(20, 2_000)) / 100.0;
        (0.20 * scale, 0.30 * scale)
    });
    let x = plot.width * f32::from(view.rot_y.min(90)).to_radians().sin() * x_factor * projection;
    let y = plot.height
        * f32::from(view.rot_x.clamp(-90, 90))
            .to_radians()
            .sin()
            .abs()
        * y_factor
        * projection;
    (
        x.max(1.0).min(plot.width * 0.45),
        y.max(1.0).min(plot.height * 0.45),
    )
}

fn chart_3d_view_plane_skew(plot: Rect, view: ChartView3D) -> f32 {
    if view.right_angle_axes {
        return 0.0;
    }
    let rot_y = if view.rot_y > 180 {
        f32::from(view.rot_y) - 360.0
    } else {
        f32::from(view.rot_y)
    };
    let perspective = 1.0 + f32::from(view.perspective.min(240)) / 480.0;
    (plot.height * rot_y.to_radians().sin() * 0.55 * perspective)
        .clamp(-plot.height * 0.35, plot.height * 0.35)
}

pub(super) fn chart_bar_3d_plane_bounds(
    bounds: Rect,
    plot: Rect,
    view: ChartView3D,
    horizontal: bool,
) -> Rect {
    if !horizontal {
        let ratio =
            ((bounds.x + bounds.width / 2.0 - plot.x) / plot.width.max(1.0)).clamp(0.0, 1.0);
        return Rect {
            y: bounds.y + chart_3d_view_plane_skew(plot, view) * ratio,
            ..bounds
        };
    }
    let (depth_x, depth_y) = chart_3d_offset(plot, view);
    // ponytail: fixed intermediate plane; use parsed c:gapDepth when series-depth layout is modeled.
    Rect {
        x: bounds.x + depth_x * 0.35,
        y: bounds.y - depth_y * 0.35,
        ..bounds
    }
}

fn chart_area_3d_offset(plot: Rect, view: ChartView3D) -> (f32, f32) {
    let (x, y) = chart_3d_offset(plot, view);
    (x, (y * 2.4).min(plot.height * 0.30))
}

/// Major gridline of a cartesian value axis.
///
/// `three_d_bars` projects the line onto the chart floor like Excel does for
/// `bar3DChart`; otherwise the line spans the plot rect on the value axis only.
pub(super) fn chart_value_axis_grid_line(
    plot: Rect,
    ratio: f32,
    view: ChartView3D,
    three_d_bars: bool,
    horizontal_bars: bool,
) -> (Rect, Geometry) {
    let y = plot.y + plot.height * (1.0 - ratio);
    if three_d_bars && !horizontal_bars {
        let (depth_x, depth_y) = chart_3d_offset(plot, view);
        let skew = chart_3d_view_plane_skew(plot, view);
        let skew_min = skew.min(0.0);
        return (
            Rect {
                x: plot.x,
                y: y - depth_y + skew_min,
                width: plot.width + depth_x,
                height: depth_y + skew.abs(),
            },
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo {
                        x: 0.0,
                        y: depth_y - skew_min,
                    },
                    PathCommand::LineTo {
                        x: depth_x,
                        y: -skew_min,
                    },
                    PathCommand::LineTo {
                        x: plot.width + depth_x,
                        y: skew - skew_min,
                    },
                ],
            },
        );
    }
    let x = plot.x + plot.width * ratio;
    if three_d_bars && horizontal_bars {
        let (depth_x, depth_y) = chart_3d_offset(plot, view);
        return (
            Rect {
                x,
                y: plot.y,
                width: depth_x,
                height: plot.height,
            },
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo { x: depth_x, y: 0.0 },
                    PathCommand::LineTo {
                        x: depth_x,
                        y: plot.height - depth_y,
                    },
                    PathCommand::LineTo {
                        x: 0.0,
                        y: plot.height,
                    },
                ],
            },
        );
    }
    if horizontal_bars {
        (
            Rect {
                x,
                y: plot.y,
                width: 0.01,
                height: plot.height,
            },
            Geometry::Line,
        )
    } else {
        (
            Rect {
                x: plot.x,
                y,
                width: plot.width,
                height: 0.01,
            },
            Geometry::Line,
        )
    }
}

impl Chart {
    pub(super) fn depth_axis_point(&self, plot: Rect, x: f32, y: f32, z: f32) -> Option<(f32, f32)> {
        self.series.iter().any(ChartSeries::projected_column)
            .then(|| chart_project_3d(plot, self.column_view_3d(), x, y, z))
    }

    pub(super) fn projected_axis_line(&self, plot: Rect, horizontal: bool) -> Option<(Rect, Geometry)> {
        let start = self.depth_axis_point(plot, 0.0, 0.0, 0.0)?;
        let end = self.depth_axis_point(plot, if horizontal { 1.0 } else { 0.0 }, if horizontal { 0.0 } else { 1.0 }, 0.0)?;
        Some((plot, Geometry::Path { fill_rule: FillRule::NonZero, commands: vec![
            PathCommand::MoveTo { x: start.0 - plot.x, y: start.1 - plot.y },
            PathCommand::LineTo { x: end.0 - plot.x, y: end.1 - plot.y },
        ] }))
    }

    pub(super) fn bar_grid_line(&self, plot: Rect, ratio: f32, three_d: bool, horizontal: bool) -> (Rect, Geometry) {
        if self.depth_axis_point(plot, 0.0, ratio, 0.0).is_some() {
            let points = [(0.0,0.0), (0.0,1.0), (1.0,1.0)].map(|(x,z)| {
                let (x,y) = chart_project_3d(plot, self.column_view_3d(), x,ratio,z);
                (x-plot.x,y-plot.y)
            });
            return (plot, Geometry::Path { fill_rule: FillRule::NonZero, commands: vec![
                PathCommand::MoveTo {x:points[0].0,y:points[0].1},
                PathCommand::LineTo {x:points[1].0,y:points[1].1},
                PathCommand::LineTo {x:points[2].0,y:points[2].1},
            ] });
        }
        chart_value_axis_grid_line(plot, ratio, self.view_3d.unwrap_or_default(), three_d, horizontal)
    }

    pub(super) fn bar_floor(&self, plot: Rect) -> (Rect, Geometry, Paint) {
        let view = self.column_view_3d();
        if self.series.iter().any(ChartSeries::projected_column) {
            let points = [(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)].map(|(x, z)| {
                let (x, y) = chart_project_3d(plot, view, x, 0.0, z);
                (x - plot.x, y - plot.y)
            });
            (
                plot,
                super::polygon_geometry(&points),
                Paint::Solid(0xffff_ffff),
            )
        } else {
            chart_bar_3d_floor(plot, view)
        }
    }
}

pub(super) fn chart_bar_3d_floor(plot: Rect, view: ChartView3D) -> (Rect, Geometry, Paint) {
    let (depth_x, depth_y) = chart_3d_offset(plot, view);
    let skew = chart_3d_view_plane_skew(plot, view);
    let skew_min = skew.min(0.0);
    (
        Rect {
            x: plot.x,
            y: plot.y + plot.height - depth_y + skew_min,
            width: plot.width + depth_x,
            height: depth_y + skew.abs(),
        },
        Geometry::Path {
            fill_rule: FillRule::NonZero,
            commands: vec![
                PathCommand::MoveTo {
                    x: 0.0,
                    y: depth_y - skew_min,
                },
                PathCommand::LineTo {
                    x: plot.width,
                    y: depth_y + skew - skew_min,
                },
                PathCommand::LineTo {
                    x: plot.width + depth_x,
                    y: skew - skew_min,
                },
                PathCommand::LineTo {
                    x: depth_x,
                    y: -skew_min,
                },
                PathCommand::ClosePath,
            ],
        },
        // ponytail: default wireframe floor; parse c:floor/c:spPr when authored floor styles matter.
        Paint::None,
    )
}

/// Back and left-side walls of a 3D bar chart.
///
/// Returns wall geometry in painter's order (back wall first, then side wall).
/// Unstyled walls stay transparent so they do not cover authored gridlines.
pub(super) fn chart_bar_3d_walls(plot: Rect, view: ChartView3D) -> [(Rect, Geometry, Paint); 2] {
    let (depth_x, depth_y) = chart_3d_offset(plot, view);
    let skew = chart_3d_view_plane_skew(plot, view);
    let skew_min = skew.min(0.0);
    let paint = || Paint::None;
    [
        // Back wall.
        (
            Rect {
                x: plot.x + depth_x,
                y: plot.y - depth_y + skew_min,
                width: plot.width,
                height: plot.height + skew.abs(),
            },
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo {
                        x: 0.0,
                        y: plot.height - skew_min,
                    },
                    PathCommand::LineTo {
                        x: plot.width,
                        y: plot.height + skew - skew_min,
                    },
                    PathCommand::LineTo {
                        x: plot.width,
                        y: skew - skew_min,
                    },
                    PathCommand::LineTo {
                        x: 0.0,
                        y: -skew_min,
                    },
                    PathCommand::ClosePath,
                ],
            },
            paint(),
        ),
        // Side (left) wall.
        (
            Rect {
                x: plot.x,
                y: plot.y - depth_y,
                width: depth_x,
                height: plot.height + depth_y,
            },
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo { x: 0.0, y: depth_y },
                    PathCommand::LineTo { x: depth_x, y: 0.0 },
                    PathCommand::LineTo {
                        x: depth_x,
                        y: plot.height,
                    },
                    PathCommand::LineTo {
                        x: 0.0,
                        y: plot.height + depth_y,
                    },
                    PathCommand::ClosePath,
                ],
            },
            paint(),
        ),
    ]
}

pub(super) fn chart_bar_3d_faces(
    bounds: Rect,
    color: u32,
    view: ChartView3D,
    horizontal: bool,
) -> [(Rect, Geometry, Paint); 2] {
    let thickness = if horizontal {
        bounds.height
    } else {
        bounds.width
    };
    let thickness = thickness * f32::from(view.depth_percent.unwrap_or(100).clamp(20, 2000)) / 100.0;
    let projection = if view.right_angle_axes { 1.0 } else { 0.85 };
    let depth_x =
        (thickness * (view.rot_y.min(90) as f32).to_radians().sin().abs() * projection).max(2.0);
    let depth_y = (thickness * (view.rot_x as f32).to_radians().sin().abs() * projection).max(1.0);
    let mut top_color = color;
    apply_color_transform(&mut top_color, "shade", 0.55);
    let mut side_color = color;
    apply_color_transform(&mut side_color, "shade", 0.72);
    [
        (
            Rect {
                x: bounds.x,
                y: bounds.y - depth_y,
                width: bounds.width + depth_x,
                height: depth_y,
            },
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo { x: 0.0, y: depth_y },
                    PathCommand::LineTo { x: depth_x, y: 0.0 },
                    PathCommand::LineTo {
                        x: bounds.width + depth_x,
                        y: 0.0,
                    },
                    PathCommand::LineTo {
                        x: bounds.width,
                        y: depth_y,
                    },
                    PathCommand::ClosePath,
                ],
            },
            Paint::Solid(top_color),
        ),
        (
            Rect {
                x: bounds.x + bounds.width,
                y: bounds.y - depth_y,
                width: depth_x,
                height: bounds.height + depth_y,
            },
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo { x: 0.0, y: depth_y },
                    PathCommand::LineTo { x: depth_x, y: 0.0 },
                    PathCommand::LineTo {
                        x: depth_x,
                        y: bounds.height,
                    },
                    PathCommand::LineTo {
                        x: 0.0,
                        y: bounds.height + depth_y,
                    },
                    PathCommand::ClosePath,
                ],
            },
            Paint::Solid(side_color),
        ),
    ]
}

/// Resolve point overrides and series defaults once for all OOXML hosts.
pub(super) fn chart_bar_data_label(
    chart: &Chart,
    series: &ChartSeries,
    index: usize,
    bar: Rect,
    value: f32,
) -> Option<(Rect, String, ChartTextStyle, Option<(u32, f32)>)> {
    if series.hidden_labels.contains(&index) {
        return None;
    }
    let point = series.data_labels.iter().find(|label| label.index == index);
    let custom = point.is_some_and(|label| !label.text.is_empty());
    let show_value = point
        .and_then(|label| label.show_values)
        .unwrap_or(series.show_values);
    let show_category = point
        .and_then(|label| label.show_category_name)
        .unwrap_or(series.show_category_name);
    if !custom && !show_value && !show_category {
        return None;
    }
    let text = if let Some(label) = point.filter(|label| !label.text.is_empty()) {
        label.text.clone()
    } else {
        let mut parts = Vec::new();
        if show_category {
            parts.push(series.categories.get(index).cloned().unwrap_or_default());
        }
        if show_value {
            parts.push(
                series.value_label(
                    index,
                    value,
                    point
                        .and_then(|label| label.number_format.as_deref())
                        .or(series.number_format.as_deref()),
                ),
            );
        }
        parts.join("\n")
    };
    let mut style = chart.data_label_text_style(series, point);
    let border = point
        .and_then(|label| label.border_color.map(|color| (color, label.border_width)))
        .or(series.data_label_border);
    let position = point
        .and_then(|label| label.position.as_deref())
        .or(series.data_label_position.as_deref());
    if matches!(
        series.kind,
        ChartKind::Line | ChartKind::Area | ChartKind::Scatter
    ) {
        let width = text
            .lines()
            .map(|line| {
                line.chars()
                    .map(|c| drawingml_fallback_character_width(c, style.font_size))
                    .sum::<f32>()
            })
            .fold(0.0_f32, f32::max)
            * 1.12
            + 8.0;
        let height = style.font_size * 1.25 * text.lines().count().max(1) as f32;
        let gap = series.marker_size.unwrap_or(8.0) / 2.0 + 4.0;
        let mut bounds = Rect {
            x: bar.x - width / 2.0,
            y: bar.y - height / 2.0,
            width,
            height,
        };
        match position.unwrap_or("t") {
            "t" => bounds.y = bar.y - gap - height,
            "b" => bounds.y = bar.y + gap,
            "l" => bounds.x = bar.x - gap - width,
            "r" => bounds.x = bar.x + gap,
            _ => {}
        }
        return Some((bounds, text, style, border));
    }
    let mut bounds = if custom && position.is_none() {
        chart_bar_custom_label_bounds(series, bar, point.unwrap(), style.font_size, value)
    } else {
        if position == series.data_label_position.as_deref() {
            chart_bar_label_bounds(series, bar, value)
        } else {
            chart_bar_label_bounds_at(series, bar, value, position)
        }
    };
    if !custom && border.is_none() && point.is_none_or(|label| label.manual_size.is_none()) {
        let rotation = series
            .data_label_rotation_degrees
            .unwrap_or(0.0)
            .to_radians();
        let rotated = rotation.sin().abs() > 0.99;
        let outside = position.map_or(series.grouping == ChartGrouping::Standard, |p| {
            p == "outEnd"
        });
        if outside || rotated {
            let measured = text
                .chars()
                .map(|c| drawingml_fallback_character_width(c, style.font_size))
                .sum::<f32>()
                * 1.12
                + 4.0;
            let available = if series.bar_horizontal {
                bar.height
            } else {
                bar.width
            };
            if rotated && !outside && measured > available {
                style.font_size *= (available / measured).min(1.0);
            }
            let width = if rotated && !outside {
                measured.min(available)
            } else {
                measured
            };
            let height = style.font_size * 1.25;
            bounds.x += (bounds.width - width) / 2.0;
            bounds.y += (bounds.height - height) / 2.0;
            bounds.width = width;
            bounds.height = height;
            if outside && series.bar_horizontal && !rotated {
                bounds.x = if value >= 0.0 {
                    bar.x + bar.width + 2.0
                } else {
                    bar.x - width - 2.0
                };
            }
        }
    }
    if border.is_some() || point.is_some_and(|label| label.font_size.is_some()) {
        let width = text
            .lines()
            .map(|line| line.chars().count())
            .max()
            .unwrap_or(0) as f32
            * style.font_size
            * 0.58
            + 8.0;
        let height = style.font_size * 1.25 * text.lines().count().max(1) as f32 + 4.0;
        bounds.x += (bounds.width - width) / 2.0;
        bounds.y += (bounds.height - height) / 2.0;
        bounds.width = width;
        bounds.height = height;
    }
    Some((bounds, text, style, border))
}

impl ChartSeries {
    pub(super) fn point_stroke(&self, index: usize) -> Paint {
        self.point_border_colors
            .get(index)
            .copied()
            .flatten()
            .map_or(Paint::None, Paint::Solid)
    }
    pub(super) fn point_stroke_width(&self, index: usize) -> f32 {
        self.point_border_widths.get(index).copied().unwrap_or(0.0)
    }
    pub(super) fn apply_effects(&self, object: Option<&mut Object>) {
        if !self.effects.is_empty()
            && let Some(object) = object
        {
            object.visual = self.effects.clone().wrap(object.visual.clone(), None);
        }
    }
}
