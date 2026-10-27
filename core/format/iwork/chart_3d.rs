//! TSCH scene construction. Camera and scene properties stay in this adapter;
//! projection and polygon primitives are shared with the other chart formats.
use super::*;

type Point = [f32; 3];

#[derive(Clone, Debug)]
pub(super) struct View {
    kind: KeynoteChartKind,
    rotation: [f32; 2],
    depth: f32,
    series_gap: f32,
    lights: Option<Vec<Light>>,
}

#[derive(Clone, Debug)]
struct Light {
    direction: Point,
    space: u64,
    ambient: Point,
    diffuse: Point,
    intensity: f32,
}

// Newell's normal retains valid planar faces whose leading vertices are collinear.
fn polygon_normal(points: &[Point]) -> Option<Point> {
    let mut normal = [0.0_f64; 3];
    for (a, b) in points
        .iter()
        .zip(points.iter().cycle().skip(1))
        .take(points.len())
    {
        for (axis, value) in normal.iter_mut().enumerate() {
            let i = (axis + 1) % 3;
            let j = (axis + 2) % 3;
            *value += (f64::from(a[i]) - f64::from(b[i])) * (f64::from(a[j]) + f64::from(b[j]));
        }
    }
    let length = normal.iter().map(|v| v * v).sum::<f64>().sqrt();
    (length.is_finite() && length > f64::from(f32::EPSILON))
        .then(|| normal.map(|v| (v / length) as f32))
}

fn vector(bytes: Option<&[u8]>, part: &str, default: Point) -> Result<Point, Diagnostic> {
    let Some(bytes) = bytes else {
        return Ok(default);
    };
    let mut result = default;
    for (index, value) in result.iter_mut().enumerate() {
        *value = keynote_fixed32_field(bytes, index as u64 + 1, part)?
            .filter(|v| v.is_finite())
            .unwrap_or(*value);
    }
    Ok(result)
}

impl View {
    pub(super) fn parse(
        properties: &[u8],
        styles: &[u8],
        kind: KeynoteChartKind,
        part: &str,
    ) -> Result<Self, Diagnostic> {
        let rotation = numbers_first_bytes(properties, 4, part)?;
        let component = |bytes: Option<&[u8]>, field, default| -> Result<f32, Diagnostic> {
            Ok(bytes
                .map(|v| keynote_fixed32_field(v, field, part))
                .transpose()?
                .flatten()
                .filter(|v| v.is_finite())
                .unwrap_or(default))
        };
        let scale = numbers_first_bytes(
            properties,
            match kind {
                KeynoteChartKind::Area => 5,
                KeynoteChartKind::Bar => 6,
                KeynoteChartKind::Column => 7,
                KeynoteChartKind::Line => 9,
                KeynoteChartKind::Pie => 10,
                _ => 8,
            },
            part,
        )?;
        let relative_depth = component(scale, 3, 1.0)? / component(scale, 4, 1.0)?.max(0.01);
        let lighting_field = match kind {
            KeynoteChartKind::Area => 2,
            KeynoteChartKind::Bar => 3,
            KeynoteChartKind::Column => 4,
            KeynoteChartKind::Line => 6,
            KeynoteChartKind::Pie => 7,
            _ => 5,
        };
        let package = numbers_first_bytes(styles, lighting_field, part)?
            .or(numbers_first_bytes(styles, 5, part)?);
        let normalize = |direction: Point| {
            let length = direction
                .iter()
                .map(|v| v * v)
                .sum::<f32>()
                .sqrt()
                .max(f32::EPSILON);
            direction.map(|v| v / length)
        };
        let native_default = matches!(
            kind,
            KeynoteChartKind::Bar
                | KeynoteChartKind::Column
                | KeynoteChartKind::Line
                | KeynoteChartKind::Area
                | KeynoteChartKind::Pie
        );
        let mut lights = Vec::new();
        if let Some(package) = package {
            for light in numbers_bytes(package, 2, part)?.into_iter().take(64) {
                if numbers_varint(light, 8, part)? == Some(0) {
                    continue;
                }
                let point = numbers_first_bytes(light, 9, part)?;
                let directional = numbers_first_bytes(light, 10, part)?;
                let direction = if let Some(point) = point {
                    vector(numbers_first_bytes(point, 1, part)?, part, [0.0, 0.0, 1.0])?
                } else if let Some(directional) = directional {
                    vector(
                        numbers_first_bytes(directional, 1, part)?,
                        part,
                        [0.0, 0.0, -1.0],
                    )?
                    .map(|v| -v)
                } else {
                    [0.0, 0.0, 1.0]
                };
                lights.push(Light {
                    direction: normalize(direction),
                    space: numbers_varint(light, 7, part)?.unwrap_or(0),
                    ambient: vector(numbers_first_bytes(light, 2, part)?, part, [0.0; 3])?,
                    diffuse: vector(numbers_first_bytes(light, 3, part)?, part, [1.0; 3])?,
                    intensity: keynote_fixed32_field(light, 5, part)?
                        .filter(|v| v.is_finite())
                        .unwrap_or(1.0)
                        .clamp(0.0, 100.0),
                });
            }
        } else if native_default {
            // Independent native controls for all eight static 3D types:
            // omitted packages and serialized Default produce byte-identical
            // Apple PNGs. Authored packages (including Glossy) stay authoritative.
            for (direction, intensity) in [
                ([10.0, 15.0, 30.0], 0.4),
                ([0.251_457_27, 0.569_231_5, 0.782_780_1], 0.75),
                ([100.0, 0.0, 0.0], 0.2),
                ([-100.0, 0.0, 0.0], 0.1),
            ] {
                lights.push(Light {
                    direction: normalize(direction),
                    space: 1,
                    ambient: [0.0; 3],
                    diffuse: [1.0; 3],
                    intensity,
                });
            }
        }
        Ok(Self {
            kind,
            // An explicit empty/disabled rig has no enabled light. Retain the
            // legacy fallback only for kinds outside the native 3D controls.
            lights: (package.is_some() || native_default).then_some(lights),
            series_gap: keynote_fixed32_field(properties, 3, part)?
                .filter(|v| v.is_finite())
                .unwrap_or(1.0)
                .max(0.0),
            rotation: [
                component(rotation, 1, 18.5)?,
                -component(rotation, 2, -18.125)?,
            ],
            depth: (relative_depth
                * match kind {
                    KeynoteChartKind::Bar => 0.495_170_2,
                    KeynoteChartKind::Column => 0.110_204_68,
                    KeynoteChartKind::Pie => 0.060_881_1,
                    KeynoteChartKind::Area => 0.100_766,
                    KeynoteChartKind::Line => 0.245_881_33,
                    _ => 1.0,
                })
            .clamp(0.01, 20.0),
        })
    }

    fn series_depth(&self, index: usize, count: usize, stacked: bool) -> (f32, f32) {
        // TSCH stores the gap as a ratio of one series' depth. Stacked area
        // shares one extrusion, rather than staggering the cumulative layers.
        if stacked {
            return (0.0, 1.0);
        }
        let count = count.max(1) as f64;
        let gap = f64::from(self.series_gap);
        let span = count + (count - 1.0) * gap;
        let low = index as f64 * (1.0 + gap) / span;
        (low as f32, (low + 1.0 / span) as f32)
    }

    fn aspect(&self, plot: Rect) -> [f32; 2] {
        [plot.height / plot.width.max(1.0), self.depth]
    }

    fn project(&self, plot: Rect, point: Point) -> (f32, f32) {
        crate::format::chart_project_volume(
            plot,
            self.rotation,
            self.aspect(plot),
            0.5,
            point,
            -1.0,
        )
    }

    fn camera(&self, plot: Rect, point: Point) -> Point {
        crate::format::chart_camera_point(self.rotation, self.aspect(plot), point, -1.0)
    }
}

struct Face {
    points: Vec<Point>,
    color: u32,
    depth: f32,
}

struct Scene<'a> {
    view: &'a View,
    plot: Rect,
    faces: Vec<Face>,
    face_limit: usize,
    part: &'a str,
}

impl Scene<'_> {
    fn face(&mut self, points: Vec<Point>, color: u32) -> Result<(), Diagnostic> {
        if points.len() < 3 {
            return Ok(());
        }
        let screen = points
            .iter()
            .map(|p| self.view.project(self.plot, *p))
            .collect::<Vec<_>>();
        let area: f32 = screen
            .iter()
            .zip(screen.iter().cycle().skip(1))
            .take(screen.len())
            .map(|(a, b)| a.0 * b.1 - b.0 * a.1)
            .sum();
        if area <= 0.0 {
            return Ok(());
        }
        if self.faces.len() >= self.face_limit {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ObjectLimit,
                Phase::Parse,
                None,
                "iWork chart mesh exceeds the configured object limit",
            )
            .in_part(self.part));
        }
        let camera = points
            .iter()
            .map(|p| self.view.camera(self.plot, *p))
            .collect::<Vec<_>>();
        let Some(normal) = polygon_normal(&camera) else {
            return Ok(());
        };
        let bar_lighting = self.view.kind == KeynoteChartKind::Bar && self.view.lights.is_some();
        // TSCH bars light the pre-projection chart model, whose category/value
        // axes differ from the displayed horizontal data axes. Native Keynote
        // 09 export independently swaps the two bar scale components.
        let bar_normal = if bar_lighting {
            polygon_normal(&points).map(|n| [-n[1], n[0], -n[2]])
        } else {
            None
        };
        // The screen-space Newell normal carries the projection's horizontal
        // reflection. Native world-space lights use a right-handed X axis;
        // six signed directional controls isolate this from point/specular terms.
        let world_normal = if bar_lighting {
            [-normal[0], normal[1], normal[2]]
        } else {
            normal
        };
        let mut illumination = [0.0_f32; 3];
        if self.view.lights.is_none() {
            illumination = [1.0; 3];
        }
        for light in self.view.lights.iter().flatten() {
            let light_normal = match light.space {
                1 => bar_normal.unwrap_or(normal),
                0 => world_normal,
                _ => normal,
            };
            let diffuse = light_normal
                .iter()
                .zip(light.direction)
                .map(|(a, b)| a * b)
                .sum::<f32>()
                .max(0.0);
            for (index, value) in illumination.iter_mut().enumerate() {
                *value += light.intensity * (light.ambient[index] + diffuse * light.diffuse[index]);
            }
        }
        let color =
            [24, 16, 8]
                .into_iter()
                .enumerate()
                .fold(color & 255, |rgba, (index, shift)| {
                    let component = ((color >> shift) & 255) as f32 / 255.0;
                    let encoded = if bar_lighting {
                        // Fixed-function ambient defaults: scene 0.2 times
                        // material 0.2. The native all-disabled rig yields raw
                        // RGB 10; a single left fill adds its encoded diffuse.
                        // Native fourfold single-point-light control keeps the
                        // diffuse end-cap at its material color, rather than
                        // bleaching toward white. Specular remains separate
                        // and unvalidated; do not clamp the final surface RGB
                        // as a substitute for its missing contribution.
                        (0.2 * 0.2 + component * illumination[index].clamp(0.0, 1.0))
                            .clamp(0.0, 1.0)
                    } else {
                        let linear = crate::format::srgb_to_linear(component) * illumination[index];
                        crate::format::linear_to_srgb(linear.clamp(0.0, 1.0))
                    };
                    rgba | (((encoded * 255.0).round() as u32) << shift)
                });
        let depth = camera.iter().map(|p| p[2]).sum::<f32>() / camera.len() as f32;
        self.faces.push(Face {
            points,
            color,
            depth,
        });
        Ok(())
    }

    fn box_faces(&mut self, lo: Point, hi: Point, color: u32) -> Result<(), Diagnostic> {
        for (axis, fixed, reverse) in [
            (0, lo[0], true),
            (0, hi[0], false),
            (1, lo[1], true),
            (1, hi[1], false),
            (2, lo[2], true),
            (2, hi[2], false),
        ] {
            let i = (axis + 1) % 3;
            let j = (axis + 2) % 3;
            let mut points = Vec::with_capacity(4);
            for (u, v) in [
                (lo[i], lo[j]),
                (hi[i], lo[j]),
                (hi[i], hi[j]),
                (lo[i], hi[j]),
            ] {
                let mut p = [0.0; 3];
                p[axis] = fixed;
                p[i] = u;
                p[j] = v;
                points.push(p);
            }
            if reverse {
                points.reverse();
            }
            self.face(points, color)?;
        }
        Ok(())
    }

    fn prism(
        &mut self,
        profile: &[(f32, f32)],
        low: f32,
        high: f32,
        color: u32,
        show_side: impl Fn(usize) -> bool,
    ) -> Result<(), Diagnostic> {
        if profile.len() < 3 {
            return Ok(());
        }
        let clockwise = profile
            .iter()
            .zip(profile.iter().cycle().skip(1))
            .take(profile.len())
            .map(|(a, b)| f64::from(a.0) * f64::from(b.1) - f64::from(a.1) * f64::from(b.0))
            .sum::<f64>()
            < 0.0;
        let mut near = profile
            .iter()
            .map(|&(x, y)| [x, y, low])
            .collect::<Vec<_>>();
        let mut far = profile
            .iter()
            .map(|&(x, y)| [x, y, high])
            .collect::<Vec<_>>();
        if clockwise {
            far.reverse();
        } else {
            near.reverse();
        }
        self.face(near, color)?;
        self.face(far, color)?;
        for i in 0..profile.len() {
            if !show_side(i) {
                continue;
            }
            let a = profile[i];
            let b = profile[(i + 1) % profile.len()];
            let mut side = vec![
                [a.0, a.1, low],
                [b.0, b.1, low],
                [b.0, b.1, high],
                [a.0, a.1, high],
            ];
            // The low face must face outwards toward -Z; side winding depends
            // on the original 2D profile, independently of that face reversal.
            if clockwise {
                side.reverse();
            }
            self.face(side, color)?;
        }
        Ok(())
    }
}

// A self-crossing signed band has valid lobes even when its signed area is
// zero. Split at upper/lower intersections before culling and extrusion; 2D
// nonzero fills can keep the original path and do not need this mesh step.
fn split_crossing_area_profiles(
    upper: &[(f32, f32)],
    lower: &[(f32, f32)],
) -> Vec<Vec<(f32, f32)>> {
    debug_assert_eq!(upper.len(), lower.len());
    if upper.is_empty() {
        return Vec::new();
    }
    let finish = |top: &[(f32, f32)], bottom: &[(f32, f32)]| {
        if top.iter().zip(bottom).all(|(a, b)| a.1 == b.1) {
            return None;
        }
        let mut profile = top.to_vec();
        profile.extend(bottom.iter().rev().copied());
        profile.dedup();
        if profile.first() == profile.last() {
            profile.pop();
        }
        (profile.len() >= 3).then_some(profile)
    };
    let (mut top, mut bottom) = (vec![upper[0]], vec![lower[0]]);
    let mut profiles = Vec::new();
    for index in 1..upper.len() {
        let before = upper[index - 1].1 - lower[index - 1].1;
        let after = upper[index].1 - lower[index].1;
        if (before > 0.0 && after < 0.0) || (before < 0.0 && after > 0.0) {
            let t = f64::from(before) / (f64::from(before) - f64::from(after));
            let a = upper[index - 1];
            let b = upper[index];
            let intersection = (
                (f64::from(a.0) + (f64::from(b.0) - f64::from(a.0)) * t) as f32,
                (f64::from(a.1) + (f64::from(b.1) - f64::from(a.1)) * t) as f32,
            );
            top.push(intersection);
            bottom.push(intersection);
            if let Some(profile) = finish(&top, &bottom) {
                profiles.push(profile);
            }
            top.clear();
            bottom.clear();
            top.push(intersection);
            bottom.push(intersection);
        }
        top.push(upper[index]);
        bottom.push(lower[index]);
        if after == 0.0 {
            if let Some(profile) = finish(&top, &bottom) {
                profiles.push(profile);
            }
            top.clear();
            bottom.clear();
            top.push(upper[index]);
            bottom.push(lower[index]);
        }
    }
    if let Some(profile) = finish(&top, &bottom) {
        profiles.push(profile);
    }
    profiles
}

#[allow(clippy::too_many_arguments)]
pub(super) fn push(
    chart: &KeynoteChart,
    identifier: u64,
    unit_index: u32,
    part: &str,
    limits: Limits,
    objects: &mut Vec<Object>,
    plot: Rect,
    z: i32,
    parent_numeric_id: Option<u32>,
    parent_stable_id: Option<&str>,
    serial: &mut u32,
) -> Result<(), Diagnostic> {
    let Some(view) = &chart.view_3d else {
        return Ok(());
    };
    let mut scene = Scene {
        view,
        plot,
        faces: Vec::new(),
        face_limit: limits.max_document_objects.saturating_sub(objects.len()),
        part,
    };
    let series = iwork_chart_plot_series(chart);
    let categories = series.iter().map(Vec::len).max().unwrap_or(0);
    if categories == 0 {
        return Ok(());
    }
    let (minimum, maximum) = chart.value_axis.data_range(&series, chart.stacked);
    let ratio = |v: f32| {
        chart
            .value_axis
            .ratio(v.clamp(minimum, maximum), minimum, maximum)
            .unwrap_or(0.0)
    };
    let baseline = ratio(0.0);
    match chart.kind {
        KeynoteChartKind::Column | KeynoteChartKind::Bar => {
            let horizontal = chart.kind == KeynoteChartKind::Bar;
            let count = if chart.stacked {
                1
            } else {
                series.len().max(1)
            };
            let (bar_gap, set_gap) = chart.bar_gaps.unwrap_or((0.1, 0.4));
            let width =
                1.0 / categories as f32 / (count as f32 + (count - 1) as f32 * bar_gap + set_gap);
            let mut positive = vec![0.0; categories];
            let mut negative = positive.clone();
            for (index, values) in series.iter().enumerate() {
                for (category, &value) in values
                    .iter()
                    .enumerate()
                    .filter(|(_, v)| v.is_finite() && **v != 0.0)
                {
                    let start = if chart.stacked {
                        if value >= 0.0 {
                            positive[category]
                        } else {
                            negative[category]
                        }
                    } else {
                        0.0
                    };
                    let end = start + value;
                    if chart.stacked {
                        if value >= 0.0 {
                            positive[category] = end;
                        } else {
                            negative[category] = end;
                        }
                    }
                    let slot = category as f32 / categories as f32
                        + width
                            * (set_gap / 2.0
                                + if chart.stacked {
                                    0.0
                                } else {
                                    index as f32 * (1.0 + bar_gap)
                                });
                    let (low, high) = (ratio(start).min(ratio(end)), ratio(start).max(ratio(end)));
                    // Clipping can collapse a nonzero datum onto a range edge.
                    // A zero-height box still has a visible cap in perspective;
                    // native TSCH omits that bar rather than drawing a plate.
                    if high <= low {
                        continue;
                    }
                    let (lo, hi) = if horizontal {
                        ([low, 1.0 - slot - width, 0.25], [high, 1.0 - slot, 0.75])
                    } else {
                        ([slot, low, 0.25], [slot + width, high, 0.75])
                    };
                    scene.box_faces(lo, hi, keynote_chart_color(chart, index))?;
                }
            }
        }
        KeynoteChartKind::Line | KeynoteChartKind::Area => {
            let mut cumulative = vec![0.0; categories];
            let mut previous_lower = if chart.stacked && chart.kind == KeynoteChartKind::Area {
                vec![baseline; categories]
            } else {
                Vec::new()
            };
            // Opaque positive co-depth bands have matching internal interfaces.
            // Signed overlapping bands do not imply a solid union, so retain
            // their sides until that separate native behavior is established.
            let shared_interfaces = chart.stacked
                && series
                    .iter()
                    .flatten()
                    .all(|value| !value.is_finite() || *value >= 0.0)
                && (0..series.len()).all(|index| keynote_chart_color(chart, index) & 0xff == 0xff);
            for (index, values) in series.iter().enumerate() {
                let (low, high) = view.series_depth(index, series.len(), chart.stacked);
                let mut prior = None;
                let mut upper = Vec::new();
                let mut lower = Vec::new();
                for category in 0..categories {
                    let value = if chart.kind == KeynoteChartKind::Area {
                        iwork_chart_area_value(values, category)
                    } else {
                        values.get(category).copied().unwrap_or(f32::NAN)
                    };
                    // Native TSCH 3D differs from the 2D adapter: line strips
                    // bridge missing cells; area profiles retain them as zero,
                    // including both boundaries of each cumulative layer.
                    let value = if value.is_finite() {
                        value
                    } else if chart.kind == KeynoteChartKind::Line {
                        continue;
                    } else {
                        0.0
                    };
                    let x = if categories == 1 {
                        0.5
                    } else {
                        category as f32 / (categories - 1) as f32
                    };
                    let bottom = if chart.stacked {
                        cumulative[category]
                    } else {
                        0.0
                    };
                    let top = bottom + value;
                    cumulative[category] = top;
                    let point = (x, ratio(top));
                    if chart.kind == KeynoteChartKind::Line {
                        if let Some((px, py)) = prior {
                            let thickness = 0.004;
                            scene.prism(
                                &[
                                    (px, py - thickness),
                                    (x, point.1 - thickness),
                                    (x, point.1 + thickness),
                                    (px, py + thickness),
                                ],
                                low,
                                high,
                                keynote_chart_color(chart, index),
                                |_| true,
                            )?;
                        }
                        prior = Some(point);
                    } else {
                        upper.push(point);
                        lower.push((x, ratio(bottom)));
                    }
                }
                if chart.kind == KeynoteChartKind::Area {
                    let count = upper.len();
                    // A collapsed area is not an extruded floor plate. Lines
                    // keep their own finite-width strip, including at zero.
                    let has_area = upper.iter().zip(&lower).any(|(a, b)| a.1 != b.1);
                    let crossing = upper.iter().zip(&lower).any(|(a, b)| a.1 > b.1)
                        && upper.iter().zip(&lower).any(|(a, b)| a.1 < b.1);
                    if crossing {
                        for profile in split_crossing_area_profiles(&upper, &lower) {
                            scene.prism(
                                &profile,
                                low,
                                high,
                                keynote_chart_color(chart, index),
                                |_| true,
                            )?;
                        }
                    } else if has_area {
                        upper.extend(lower.iter().rev().copied());
                        scene.prism(
                            &upper,
                            low,
                            high,
                            keynote_chart_color(chart, index),
                            |edge| {
                                let (category, top) = if edge + 1 < count {
                                    (edge, true)
                                } else if edge >= count && edge + 1 < 2 * count {
                                    (2 * count - 2 - edge, false)
                                } else {
                                    return true;
                                };
                                let ends = [category, category + 1];
                                if ends.iter().all(|&c| upper[c].1 == lower[c].1) {
                                    return false;
                                }
                                if !shared_interfaces {
                                    return true;
                                }
                                if top {
                                    // The next band may itself be zero or clipped;
                                    // only a band with visible height hides this cap.
                                    series.get(index + 1).is_none_or(|next| {
                                        ends.into_iter().all(|c| {
                                            let value = next
                                                .get(c)
                                                .copied()
                                                .filter(|v| v.is_finite())
                                                .unwrap_or(0.0);
                                            ratio(cumulative[c] + value) == upper[c].1
                                        })
                                    })
                                } else {
                                    index == 0
                                        || ends.into_iter().all(|c| previous_lower[c] == lower[c].1)
                                }
                            },
                        )?;
                    }
                    for (previous, point) in previous_lower.iter_mut().zip(&lower) {
                        *previous = point.1;
                    }
                }
            }
        }
        KeynoteChartKind::Pie => {
            let Some(values) = series.first() else {
                return Ok(());
            };
            let values = iwork_chart_pie_ratios(values);
            let mut angle = -std::f32::consts::FRAC_PI_2 - chart.start_rotation.to_radians();
            for (index, &fraction) in values.iter().enumerate().filter(|(_, v)| **v > 0.0) {
                let end = angle + std::f32::consts::TAU * fraction;
                let count = ((end - angle).abs() / 0.05).ceil().max(1.0) as usize;
                let mut profile = vec![(0.5, 0.5)];
                for step in 0..=count {
                    let a = angle + (end - angle) * step as f32 / count as f32;
                    profile.push((0.5 + 0.48 * a.cos(), 0.5 - 0.48 * a.sin()));
                }
                scene.prism(
                    &profile,
                    0.0,
                    1.0,
                    keynote_chart_color(chart, index),
                    |_| true,
                )?;
                angle = end;
            }
        }
        _ => {}
    }
    let horizontal = chart.kind == KeynoteChartKind::Bar;
    if chart.kind != KeynoteChartKind::Pie {
        for tick in 0..=chart.value_axis.major_intervals {
            let t = tick as f32 / chart.value_axis.major_intervals as f32;
            let ends = if horizontal {
                [[t, 0.0, 1.0], [t, 1.0, 1.0]]
            } else {
                [[0.0, t, 1.0], [1.0, t, 1.0]]
            };
            if let Some((paint, width, style)) = &chart.value_axis.grid {
                let points = ends.map(|p| view.project(plot, p));
                keynote_push_chart_child(
                    identifier,
                    unit_index,
                    part,
                    limits,
                    objects,
                    plot,
                    z,
                    parent_numeric_id,
                    parent_stable_id,
                    serial,
                    Visual::StrokeStyle {
                        style: style.clone(),
                        visual: Box::new(Visual::PaintedShape {
                            geometry: Geometry::Path {
                                fill_rule: FillRule::NonZero,
                                commands: vec![
                                    PathCommand::MoveTo {
                                        x: points[0].0 - plot.x,
                                        y: points[0].1 - plot.y,
                                    },
                                    PathCommand::LineTo {
                                        x: points[1].0 - plot.x,
                                        y: points[1].1 - plot.y,
                                    },
                                ],
                            },
                            fill: Paint::None,
                            stroke: paint.clone(),
                            stroke_width: *width,
                        }),
                    },
                )?;
            }
            if chart.value_axis.tick_label_visible(tick) {
                let point = if horizontal {
                    [t, 1.0, 1.0]
                } else {
                    [if view.rotation[1] > 0.0 { 1.0 } else { 0.0 }, t, 1.0]
                };
                let (x, y) = view.project(plot, point);
                let text = chart
                    .value_axis
                    .label(chart.value_axis.tick_value(minimum, maximum, t));
                let bounds = Rect {
                    x: if horizontal { x - 30.0 } else { x + 8.0 },
                    y: if horizontal { y - 30.0 } else { y - 8.0 },
                    width: 60.0,
                    height: chart.value_axis.text.font_size * 1.2,
                };
                keynote_push_chart_child(
                    identifier,
                    unit_index,
                    part,
                    limits,
                    objects,
                    bounds,
                    z,
                    parent_numeric_id,
                    parent_stable_id,
                    serial,
                    keynote_chart_label_visual(
                        text,
                        if horizontal {
                            TextAlign::Center
                        } else {
                            TextAlign::Start
                        },
                        &chart.value_axis.text,
                    ),
                )?;
            }
        }
    }
    // Painter order is based on camera depth, not series order or face type.
    scene.faces.sort_by(|a, b| b.depth.total_cmp(&a.depth));
    for (index, face) in scene.faces.into_iter().enumerate() {
        let points = face
            .points
            .iter()
            .map(|p| {
                let (x, y) = view.project(plot, *p);
                (x - plot.x, y - plot.y)
            })
            .collect::<Vec<_>>();
        keynote_push_chart_child(
            identifier,
            unit_index,
            part,
            limits,
            objects,
            plot,
            z,
            parent_numeric_id,
            parent_stable_id,
            serial,
            Visual::PaintedShape {
                geometry: crate::format::polygon_geometry(&points),
                fill: Paint::Solid(face.color),
                stroke: Paint::None,
                stroke_width: 0.0,
            },
        )?;
        objects
            .last_mut()
            .unwrap()
            .stable_id
            .push_str(&format!(":3d-face-{index}"));
    }
    if chart.kind == KeynoteChartKind::Pie {
        let values = series
            .first()
            .map(|values| iwork_chart_pie_ratios(values))
            .unwrap_or_default();
        let mut angle = -std::f32::consts::FRAC_PI_2 - chart.start_rotation.to_radians();
        for (index, fraction) in values.into_iter().enumerate().filter(|(_, v)| *v > 0.0) {
            let end = angle + std::f32::consts::TAU * fraction;
            let midpoint = (angle + end) / 2.0;
            angle = end;
            let Some(label) = iwork_chart_pie_label(chart, index, fraction) else {
                continue;
            };
            let radius = 0.48 * chart.pie_label_radius;
            let (x, y) = view.project(
                plot,
                [
                    0.5 + radius * midpoint.cos(),
                    0.5 - radius * midpoint.sin(),
                    0.0,
                ],
            );
            let style = &chart.pie_label_style;
            let width = plot.width * 0.38;
            let height = style.font_size
                * if chart.pie_show_category && chart.show_data_labels.first() == Some(&true) {
                    2.4
                } else {
                    1.2
                };
            keynote_push_chart_child(
                identifier,
                unit_index,
                part,
                limits,
                objects,
                Rect {
                    x: x - width / 2.0,
                    y: y - height / 2.0,
                    width,
                    height,
                },
                z,
                parent_numeric_id,
                parent_stable_id,
                serial,
                keynote_chart_label_visual(label, TextAlign::Center, style),
            )?;
            objects
                .last_mut()
                .unwrap()
                .stable_id
                .push_str(&format!(":3d-pie-label-{index}"));
        }
        return Ok(());
    }
    for (category, label) in chart
        .categories
        .iter()
        .enumerate()
        .take(categories)
        .filter(|_| chart.category_axis.show_labels)
    {
        let t = if matches!(chart.kind, KeynoteChartKind::Line | KeynoteChartKind::Area) {
            if categories == 1 {
                0.5
            } else {
                category as f32 / (categories - 1) as f32
            }
        } else {
            (category as f32 + 0.5) / categories as f32
        };
        let point = if horizontal {
            [baseline, 1.0 - t, 0.0]
        } else {
            [t, baseline, 0.0]
        };
        let (x, y) = view.project(plot, point);
        keynote_push_chart_child(
            identifier,
            unit_index,
            part,
            limits,
            objects,
            Rect {
                x: if horizontal { x - 110.0 } else { x - 60.0 },
                y: if horizontal { y - 8.0 } else { y + 5.0 },
                width: 120.0,
                height: chart.category_axis.text.font_size * 1.2,
            },
            z,
            parent_numeric_id,
            parent_stable_id,
            serial,
            keynote_chart_label_visual(
                label.clone(),
                if horizontal {
                    TextAlign::End
                } else {
                    TextAlign::Center
                },
                &chart.category_axis.text,
            ),
        )?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn crossing_area_profiles_preserve_absolute_area_and_zero_intervals() {
        for (values, count, expected_area) in [
            ([1.0, -1.0, 1.0, -1.0], 4, 1.5),
            ([1.0, 0.0, 0.0, -1.0], 2, 1.0),
        ] {
            let upper = values
                .into_iter()
                .enumerate()
                .map(|(i, v)| (i as f32, v))
                .collect::<Vec<_>>();
            let lower = (0..4).map(|i| (i as f32, 0.0)).collect::<Vec<_>>();
            let profiles = split_crossing_area_profiles(&upper, &lower);
            assert_eq!(profiles.len(), count);
            let area = profiles
                .iter()
                .map(|profile| {
                    profile
                        .iter()
                        .zip(profile.iter().cycle().skip(1))
                        .take(profile.len())
                        .map(|(a, b)| a.0 * b.1 - a.1 * b.0)
                        .sum::<f32>()
                        .abs()
                        / 2.0
                })
                .sum::<f32>();
            assert_eq!(area, expected_area);
            assert!(profiles.iter().all(|p| p.len() >= 3));
        }
    }

    #[test]
    fn extrusion_winding_preserves_the_outward_box_faces() {
        let view = View {
            kind: KeynoteChartKind::Column,
            rotation: [18.5, 18.125],
            depth: 0.2,
            series_gap: 1.0,
            lights: None,
        };
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 800.0,
            height: 600.0,
        };
        let scene = || Scene {
            view: &view,
            plot,
            faces: Vec::new(),
            face_limit: 128,
            part: "native-flat-area",
        };
        let keys = |faces: &[Face]| {
            let mut result = faces
                .iter()
                .map(|face| {
                    let mut points = face
                        .points
                        .iter()
                        .map(|point| point.map(f32::to_bits))
                        .collect::<Vec<_>>();
                    points.sort_unstable();
                    points
                })
                .collect::<Vec<_>>();
            result.sort_unstable();
            result
        };
        let mut expected = scene();
        expected.box_faces([0.0; 3], [1.0; 3], 0x00b0_f0ff).unwrap();
        let profile = [(0.0, 0.0), (1.0, 0.0), (1.0, 1.0), (0.0, 1.0)];
        for profile in [profile, [profile[3], profile[2], profile[1], profile[0]]] {
            let mut actual = scene();
            actual
                .prism(&profile, 0.0, 1.0, 0x00b0_f0ff, |_| true)
                .unwrap();
            assert_eq!(
                keys(&actual.faces),
                keys(&expected.faces),
                "constant native area must expose the outward cuboid faces for either profile winding"
            );
        }
    }
}
