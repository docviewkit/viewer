//! Shared ODF chart semantics for ODT/ODS/ODP hosts.
//!
//! Host adapters keep embedding and source locators; chart model, style
//! properties, and pie geometry stay here so data labels, explosion, and
//! legend placement cannot drift per format.

use crate::model::{PathCommand, Rect};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum OdfDataLabelNumber {
    Value,
    Percentage,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum OdfLabelPosition {
    Center,
    Inside,
    Outside,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum OdfLegendPosition {
    Top,
    Bottom,
    Start,
    End,
    Center,
}

#[derive(Clone, Debug)]
pub(super) struct OdfChartDataLabels {
    pub number: Option<OdfDataLabelNumber>,
    pub show_text: Option<bool>,
    pub show_symbol: Option<bool>,
    pub position: Option<OdfLabelPosition>,
    pub font_size: f32,
    pub font_family: String,
    pub color: u32,
    pub bold: bool,
}

impl Default for OdfChartDataLabels {
    fn default() -> Self {
        Self {
            number: None,
            show_text: None,
            show_symbol: None,
            position: None,
            font_size: 11.0 * 96.0 / 72.0,
            font_family: "Arial".to_owned(),
            color: 0x0000_00ff,
            bold: false,
        }
    }
}

impl OdfChartDataLabels {
    pub fn is_enabled(&self) -> bool {
        self.number.is_some() || self.show_text == Some(true) || self.show_symbol == Some(true)
    }

    pub fn merge(&mut self, other: &Self) {
        if other.number.is_some() {
            self.number = other.number;
        }
        if other.show_text.is_some() {
            self.show_text = other.show_text;
        }
        if other.show_symbol.is_some() {
            self.show_symbol = other.show_symbol;
        }
        if other.position.is_some() {
            self.position = other.position;
        }
        if other.font_size > 0.0 {
            self.font_size = other.font_size;
        }
        if !other.font_family.is_empty() {
            self.font_family.clone_from(&other.font_family);
        }
        self.color = other.color;
        self.bold = other.bold;
    }

    pub fn format_label(&self, category: &str, value: f32, total: f32) -> Option<String> {
        let mut parts = Vec::new();
        if self.show_symbol.unwrap_or(false) {
            parts.push(String::from("●"));
        }
        if self.show_text.unwrap_or(false) && !category.is_empty() {
            parts.push(category.to_owned());
        }
        match self.number {
            Some(OdfDataLabelNumber::Value) => parts.push(format_chart_number(value)),
            Some(OdfDataLabelNumber::Percentage) => {
                parts.push(format_chart_percent(value, total));
            }
            None => {}
        }
        if parts.is_empty() {
            None
        } else {
            Some(parts.join(" "))
        }
    }
}

#[derive(Clone, Debug, Default)]
pub(super) struct OdfChartLegend {
    pub x: Option<f32>,
    pub y: Option<f32>,
    pub width: Option<f32>,
    pub height: Option<f32>,
    pub position: Option<OdfLegendPosition>,
    pub expansion_wide: bool,
    pub font_size: f32,
    pub font_family: String,
    pub color: u32,
}

impl OdfChartLegend {
    pub fn custom_bounds(&self) -> Option<Rect> {
        match (self.x, self.y, self.width, self.height) {
            (Some(x), Some(y), Some(width), Some(height)) if width > 0.0 && height > 0.0 => {
                Some(Rect {
                    x,
                    y,
                    width,
                    height,
                })
            }
            _ => None,
        }
    }

    pub fn is_horizontal(&self) -> bool {
        matches!(
            self.position,
            Some(OdfLegendPosition::Bottom | OdfLegendPosition::Top)
        ) || (self.expansion_wide && self.height.is_none())
    }
}

/// `style:chart-properties` fields shared by chart families.
#[derive(Clone, Debug, Default)]
pub(super) struct OdfChartStyleProps {
    pub data_label_number: Option<OdfDataLabelNumber>,
    pub data_label_text: Option<bool>,
    pub data_label_symbol: Option<bool>,
    pub label_position: Option<OdfLabelPosition>,
    pub pie_offset: Option<f32>,
    pub solid_type_cuboid: bool,
}

impl OdfChartStyleProps {
    pub fn apply_data_labels(&self, base: &mut OdfChartDataLabels) {
        if self.data_label_number.is_some() {
            base.number = self.data_label_number;
        }
        if self.data_label_text.is_some() {
            base.show_text = self.data_label_text;
        }
        if self.data_label_symbol.is_some() {
            base.show_symbol = self.data_label_symbol;
        }
        if self.label_position.is_some() {
            base.position = self.label_position;
        }
    }

    pub fn data_labels(
        &self,
        font_size: f32,
        font_family: &str,
        color: u32,
        bold: bool,
    ) -> OdfChartDataLabels {
        let mut labels = OdfChartDataLabels {
            font_size,
            font_family: font_family.to_owned(),
            color,
            bold,
            ..OdfChartDataLabels::default()
        };
        self.apply_data_labels(&mut labels);
        labels
    }
}

pub(super) fn format_chart_number(value: f32) -> String {
    if (value - value.round()).abs() < 0.05 {
        format!("{}", value.round() as i64)
    } else {
        format!("{value:.1}")
    }
}

pub(super) fn format_chart_percent(value: f32, total: f32) -> String {
    if total <= 0.0 {
        return "0%".to_owned();
    }
    let percent = value / total * 100.0;
    if (percent - percent.round()).abs() < 0.05 {
        format!("{}%", percent.round() as i64)
    } else {
        format!("{percent:.1}%")
    }
}

/// Parse ODF percent (`29` or `29%`) into a radius fraction (0.29).
pub(super) fn parse_percent_fraction(value: &str) -> Option<f32> {
    let digits = value.trim().trim_end_matches('%').trim();
    let number = digits.parse::<f32>().ok().filter(|n| n.is_finite())?;
    Some((number / 100.0).clamp(0.0, 4.0))
}

pub(super) fn parse_data_label_number(value: &str) -> Option<OdfDataLabelNumber> {
    match value {
        "value" => Some(OdfDataLabelNumber::Value),
        "percentage" => Some(OdfDataLabelNumber::Percentage),
        _ => None,
    }
}

pub(super) fn parse_label_position(value: &str) -> Option<OdfLabelPosition> {
    match value {
        "center" => Some(OdfLabelPosition::Center),
        "inside" => Some(OdfLabelPosition::Inside),
        "outside" => Some(OdfLabelPosition::Outside),
        _ => None,
    }
}

pub(super) fn parse_legend_position(value: &str) -> Option<OdfLegendPosition> {
    match value {
        "top" => Some(OdfLegendPosition::Top),
        "bottom" => Some(OdfLegendPosition::Bottom),
        "start" | "left" => Some(OdfLegendPosition::Start),
        "end" | "right" => Some(OdfLegendPosition::End),
        "center" => Some(OdfLegendPosition::Center),
        _ => None,
    }
}

/// Mid-angle explosion offset used by every ODF pie host.
pub(super) fn pie_explosion_offset(explosion: f32, radius: f32, mid_angle: f32) -> (f32, f32) {
    (
        explosion * radius * mid_angle.cos(),
        explosion * radius * mid_angle.sin(),
    )
}

/// Label anchor for an exploded or joined pie slice.
pub(super) fn pie_label_anchor(
    center_x: f32,
    center_y: f32,
    offset_x: f32,
    offset_y: f32,
    radius: f32,
    mid_angle: f32,
    position: OdfLabelPosition,
) -> (f32, f32) {
    let distance = match position {
        OdfLabelPosition::Outside => radius * 1.08,
        OdfLabelPosition::Inside | OdfLabelPosition::Center => radius * 0.55,
    };
    (
        center_x + offset_x + distance * mid_angle.cos(),
        center_y + offset_y + distance * mid_angle.sin(),
    )
}

/// Legend box when the file only partially positions the legend.
pub(super) fn resolve_legend_bounds(
    legend: &OdfChartLegend,
    frame: Rect,
    categories: &[String],
) -> Rect {
    if let Some(bounds) = legend.custom_bounds() {
        return bounds;
    }
    let row_height = (legend.font_size * 1.35).max(16.0);
    let item_width = |label: &str| {
        legend.font_size * 0.6 + 4.0 + label.chars().count() as f32 * legend.font_size * 0.55 + 8.0
    };
    if legend.is_horizontal() {
        let total_width: f32 = categories.iter().map(|label| item_width(label)).sum();
        let width = legend.width.unwrap_or(total_width).min(frame.width);
        let height = legend.height.unwrap_or(row_height);
        let x = legend.x.unwrap_or(((frame.width - width) / 2.0).max(0.0));
        let y = legend.y.unwrap_or(match legend.position {
            Some(OdfLegendPosition::Top) => 2.0,
            _ => (frame.height - height - 2.0).max(0.0),
        });
        return Rect {
            x,
            y,
            width,
            height,
        };
    }
    let width = legend.width.unwrap_or(frame.width * 0.22).min(frame.width);
    let height = legend
        .height
        .unwrap_or(frame.height * 0.8)
        .min(frame.height);
    let x = legend.x.unwrap_or(frame.width * 0.76);
    let y = legend.y.unwrap_or(frame.height * 0.1);
    Rect {
        x,
        y,
        width,
        height,
    }
}

/// ODF 1.3 section 20.27: step interpolation is defined between consecutive points.
/// Host-specific axis mapping, clipping and stroke styles remain in the adapters.
pub(super) fn line_segment_commands(
    (x0, y0): (f32, f32),
    (x1, y1): (f32, f32),
    interpolation: Option<&str>,
) -> Vec<PathCommand> {
    let mut commands = vec![PathCommand::MoveTo { x: x0, y: y0 }];
    match interpolation {
        Some("step-start") => commands.push(PathCommand::LineTo { x: x0, y: y1 }),
        Some("step-end") => commands.push(PathCommand::LineTo { x: x1, y: y0 }),
        Some("step-center-x") => {
            let x = (x0 + x1) / 2.0;
            commands.push(PathCommand::LineTo { x, y: y0 });
            commands.push(PathCommand::LineTo { x, y: y1 });
        }
        Some("step-center-y") => {
            let y = (y0 + y1) / 2.0;
            commands.push(PathCommand::LineTo { x: x0, y });
            commands.push(PathCommand::LineTo { x: x1, y });
        }
        _ => {}
    }
    commands.push(PathCommand::LineTo { x: x1, y: y1 });
    commands
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn line_segments_preserve_linear_and_step_boundaries() {
        for (mode, expected) in [
            (None, vec![(2., 8.), (6., 0.)]),
            (Some("none"), vec![(2., 8.), (6., 0.)]),
            (Some("step-start"), vec![(2., 8.), (2., 0.), (6., 0.)]),
            (Some("step-end"), vec![(2., 8.), (6., 8.), (6., 0.)]),
            (
                Some("step-center-x"),
                vec![(2., 8.), (4., 8.), (4., 0.), (6., 0.)],
            ),
            (
                Some("step-center-y"),
                vec![(2., 8.), (2., 4.), (6., 4.), (6., 0.)],
            ),
        ] {
            let points = line_segment_commands((2., 8.), (6., 0.), mode)
                .into_iter()
                .map(|c| match c {
                    PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => (x, y),
                    _ => unreachable!(),
                })
                .collect::<Vec<_>>();
            assert_eq!(points, expected);
        }
    }

    #[test]
    fn formats_percent_and_value_labels_like_office() {
        assert_eq!(format_chart_percent(5.0, 20.0), "25%");
        assert_eq!(format_chart_percent(8.0, 20.0), "40%");
        assert_eq!(format_chart_number(5.0), "5");
        let mut labels = OdfChartDataLabels {
            number: Some(OdfDataLabelNumber::Percentage),
            show_text: Some(false),
            show_symbol: Some(false),
            position: Some(OdfLabelPosition::Outside),
            ..OdfChartDataLabels::default()
        };
        assert_eq!(labels.format_label("C", 8.0, 20.0).as_deref(), Some("40%"));
        labels.show_text = Some(true);
        assert_eq!(
            labels.format_label("C", 8.0, 20.0).as_deref(),
            Some("C 40%")
        );
    }

    #[test]
    fn pie_offset_and_legend_layout_match_odf_hosts() {
        assert_eq!(parse_percent_fraction("29"), Some(0.29));
        assert_eq!(parse_percent_fraction("29%"), Some(0.29));
        let (dx, dy) = pie_explosion_offset(0.29, 100.0, 0.0);
        assert!((dx - 29.0).abs() < 0.01 && dy.abs() < 0.01);
        let (x, y) = pie_label_anchor(0.0, 0.0, 29.0, 0.0, 100.0, 0.0, OdfLabelPosition::Outside);
        assert!(x > 120.0 && y.abs() < 0.01);

        let legend = OdfChartLegend {
            x: Some(20.0),
            y: Some(180.0),
            width: None,
            height: None,
            position: Some(OdfLegendPosition::Bottom),
            expansion_wide: true,
            font_size: 12.0,
            font_family: "Arial".to_owned(),
            color: 0,
        };
        let bounds = resolve_legend_bounds(
            &legend,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 400.0,
                height: 200.0,
            },
            &["A".into(), "B".into()],
        );
        assert!((bounds.y - 180.0).abs() < 0.01);
        assert!(bounds.height > 10.0 && bounds.width > 40.0);
    }
}

/// ODF's parallel camera uses a scene matrix and view vectors. This is not the
/// categorical/Euler-angle projection used by DrawingML line/surface charts.
#[derive(Clone, Debug)]
pub(super) struct OdfChartProjection {
    axes: [[f32; 3]; 3],
    extent: [f32; 4],
}

impl OdfChartProjection {
    pub fn parallel(matrix: &str, vpn: &str, vup: &str) -> Option<Self> {
        fn numbers(value: &str) -> Option<Vec<f32>> {
            value
                .trim()
                .trim_start_matches("matrix")
                .trim()
                .trim_matches(['(', ')'])
                .split_whitespace()
                .map(|v| {
                    v.trim_end_matches("cm")
                        .parse::<f32>()
                        .ok()
                        .filter(|v| v.is_finite())
                })
                .collect()
        }
        fn normalize(v: [f32; 3]) -> Option<[f32; 3]> {
            let length = v.iter().map(|v| v * v).sum::<f32>().sqrt();
            (length > 1e-6 && length.is_finite()).then(|| v.map(|v| v / length))
        }
        let cross = |a: [f32; 3], b: [f32; 3]| {
            [
                a[1] * b[2] - a[2] * b[1],
                a[2] * b[0] - a[0] * b[2],
                a[0] * b[1] - a[1] * b[0],
            ]
        };
        let m = numbers(matrix)?;
        if m.len() != 12 {
            return None;
        }
        let normal = normalize(numbers(vpn)?.try_into().ok()?)?;
        let right = normalize(cross(numbers(vup)?.try_into().ok()?, normal))?;
        let up = cross(normal, right);
        let axes = [right, up.map(|v| -v), normal]
            .map(|basis| std::array::from_fn(|i| (0..3).map(|j| m[i * 3 + j] * basis[j]).sum()));
        let mut result = Self {
            axes,
            extent: [
                f32::INFINITY,
                f32::INFINITY,
                f32::NEG_INFINITY,
                f32::NEG_INFINITY,
            ],
        };
        for x in [0.0, 1.0] {
            for y in [0.0, 1.0] {
                for z in [0.0, 1.0] {
                    let p = result.raw([x, y, z]);
                    for i in 0..2 {
                        result.extent[i] = result.extent[i].min(p[i]);
                        result.extent[i + 2] = result.extent[i + 2].max(p[i]);
                    }
                }
            }
        }
        (result.extent.iter().all(|v| v.is_finite())
            && (0..2).all(|i| result.extent[i + 2] - result.extent[i] > 1e-6))
        .then_some(result)
    }

    fn raw(&self, point: [f32; 3]) -> [f32; 3] {
        self.axes
            .map(|axis| (0..3).map(|i| axis[i] * point[i]).sum())
    }

    pub fn project(&self, point: [f32; 3], plot: Rect) -> (f32, f32) {
        let p = self.raw(point);
        (
            (p[0] - self.extent[0]) / (self.extent[2] - self.extent[0]) * plot.width,
            (p[1] - self.extent[1]) / (self.extent[3] - self.extent[1]) * plot.height,
        )
    }

    pub fn depth(&self, point: [f32; 3]) -> f32 {
        self.raw(point)[2]
    }
}

#[cfg(test)]
mod projection_tests {
    use super::*;

    #[test]
    fn parallel_camera_preserves_axes_and_rejects_degenerate_input() {
        let matrix = "matrix (1 0 0 0 1 0 0 0 1 0cm 0cm 0cm)";
        let view = OdfChartProjection::parallel(matrix, "(0 0 1)", "(0 1 0)").unwrap();
        let plot = Rect {
            x: 30.0,
            y: 40.0,
            width: 200.0,
            height: 100.0,
        };
        assert_eq!(view.project([0.25, 0.75, 0.0], plot), (50.0, 25.0));
        assert_eq!(view.project([0.25, 0.75, 1.0], plot), (50.0, 25.0));
        assert!(view.depth([0.0, 0.0, 1.0]) > view.depth([0.0, 0.0, 0.0]));
        for (matrix, vpn, vup) in [
            (matrix, "(0 0 0)", "(0 1 0)"),
            (matrix, "(0 0 1)", "(0 0 1)"),
            (matrix, "(NaN 0 1)", "(0 1 0)"),
            ("matrix (1 0)", "(0 0 1)", "(0 1 0)"),
        ] {
            assert!(OdfChartProjection::parallel(matrix, vpn, vup).is_none());
        }
    }
}

/// ODF XY charts pad both ends of the numerical domain and reuse the ODF
/// value-axis scale. DrawingML uses different domain/headroom rules.
pub(super) fn scatter_axis_range(
    data: (f32, f32),
    authored: (Option<f32>, Option<f32>, Option<f32>),
    horizontal: bool,
) -> (f32, f32, f32) {
    let (mut minimum, mut maximum) = data;
    if !minimum.is_finite() || !maximum.is_finite() {
        minimum = 0.0;
        maximum = 1.0;
    }
    if maximum <= minimum {
        maximum = minimum + minimum.abs().max(1.0);
    }
    let lower = authored.0.filter(|v| v.is_finite());
    let upper = authored.1.filter(|v| v.is_finite());
    let value_axis = value_axis_range(minimum, maximum);
    let step = authored
        .2
        .filter(|v| v.is_finite() && *v > 0.0)
        .unwrap_or_else(|| {
            if horizontal || lower.is_some() || upper.is_some() {
                super::nice_chart_step(
                    (upper.unwrap_or(maximum) - lower.unwrap_or(minimum)).abs() / 7.0,
                )
            } else {
                value_axis.2
            }
        });
    let (auto_min, auto_max) = if horizontal {
        (
            (minimum / step).ceil() * step - step,
            (maximum / step).floor() * step + step,
        )
    } else {
        (
            (value_axis.0 / step).floor() * step,
            (value_axis.1 / step).ceil() * step,
        )
    };
    let minimum = lower.unwrap_or(auto_min);
    let maximum = upper.unwrap_or(auto_max);
    let maximum = if maximum > minimum {
        maximum
    } else {
        minimum + step
    };
    (minimum, maximum, step)
}

#[cfg(test)]
mod axis_tests {
    use super::scatter_axis_range;

    #[test]
    fn odf_xy_auto_bounds_pad_domain_and_preserve_authored_scale() {
        let auto = (None, None, None);
        assert_eq!(scatter_axis_range((2.0, 9.0), auto, true), (1.0, 10.0, 1.0));
        assert_eq!(scatter_axis_range((3.0, 8.0), auto, false), (0.0, 9.0, 1.0));
        assert_eq!(
            scatter_axis_range((-8.0, -3.0), auto, false),
            super::value_axis_range(-8.0, -3.0)
        );
        assert_eq!(
            scatter_axis_range((3.0, 8.0), (Some(2.0), Some(10.0), Some(2.0)), false),
            (2.0, 10.0, 2.0)
        );
        assert_eq!(
            scatter_axis_range((0.1, 0.4), (Some(0.0), Some(0.5), Some(1.0)), false),
            (0.0, 0.5, 1.0)
        );
        let tight = scatter_axis_range((100.0, 101.0), auto, false);
        assert!(tight.0 > 99.0 && tight.1 > 101.0);
        let constant = scatter_axis_range((4.0, 4.0), auto, true);
        assert!(constant.0 < 4.0 && constant.1 > 4.0 && constant.2 > 0.0);
    }
}

pub(super) fn value_axis_range(minimum: f32, maximum: f32) -> (f32, f32, f32) {
    let mut axis_minimum = minimum;
    let mut axis_maximum = maximum;
    if axis_minimum > 0.0 {
        if axis_minimum == axis_maximum || axis_minimum / axis_maximum < 5.0 / 6.0 {
            axis_minimum = 0.0;
        } else {
            axis_minimum -= (axis_maximum - axis_minimum) / 2.0;
        }
    }
    if (axis_maximum - axis_minimum).abs() <= f32::EPSILON {
        axis_maximum = if axis_maximum == 0.0 {
            1.0
        } else {
            axis_maximum * 2.0
        };
    }
    let raw_step = ((axis_maximum - axis_minimum) / 10.0).max(f32::MIN_POSITIVE);
    let magnitude = 10.0_f32.powf(raw_step.log10().floor());
    let normalized = raw_step / magnitude;
    let tick_step = if normalized <= 1.0 {
        magnitude
    } else if normalized <= 2.0 {
        2.0 * magnitude
    } else if normalized <= 5.0 {
        5.0 * magnitude
    } else {
        10.0 * magnitude
    };
    axis_minimum = (axis_minimum / tick_step).floor() * tick_step;
    axis_maximum = (axis_maximum / tick_step).ceil() * tick_step;
    if axis_minimum != 0.0 && (axis_maximum - minimum) / (axis_maximum - axis_minimum) > 20.0 / 21.0
    {
        axis_minimum -= tick_step;
    }
    if axis_maximum != 0.0 && (maximum - axis_minimum) / (axis_maximum - axis_minimum) > 20.0 / 21.0
    {
        axis_maximum += tick_step;
    }
    (axis_minimum, axis_maximum, tick_step)
}
