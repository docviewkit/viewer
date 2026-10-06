//! Shared DrawingML/ChartML data used by DOCX, PPTX, and XLSX host adapters.

pub(super) use super::ellipse_arc_bezier_points;

use std::collections::HashMap;
#[cfg(feature = "native-formats")]
use std::collections::HashSet;

use super::optional_xml_attribute as string_attribute;
use crate::diagnostic::{Diagnostic, DiagnosticCode, Phase};
use crate::model::{
    AffineTransform, Backdrop3D, Bevel3D, FillRule, Geometry, Glow, GradientStop, ImageCrop,
    LineAlignment, LineCap, LineCompound, OuterShadow, Paint, PathCommand, Rect, Reflection,
    Shadow, StrokeStyle, TextAlign, TextLayout, TextOrientation, TextRun, TextVerticalAlign,
    ThreeDStyle, Visual,
};
use crate::package::Package;
#[cfg(feature = "native-formats")]
use crate::package::Relationship;
use crate::xml::XmlAttribute;
#[cfg(feature = "native-formats")]
use crate::xml::{XmlEvent, decode_xml_text, parse_ooxml as parse_xml};

use super::presentation_image::{office_image_media_type, office_image_media_type_from_signature};
use super::{color_from_hsl, color_to_hsl, transform_luminance};
#[cfg(feature = "native-formats")]
use super::{excel_serial_date, excel_serial_from_date, local_name};

const EMU_PER_CSS_PIXEL: f32 = 9_525.0;

// A ChartML/ChartEx choice owns its data rendering, including an explicit
// diagnostic placeholder when data is unavailable. Its fallback bitmap is
// never an alternative final chart in either presentation or worksheet hosts.
#[derive(Default)]
pub(super) struct DrawingMlChartChoice {
    frames: Vec<(usize, bool)>,
    skipped_depth: Option<usize>,
}

impl DrawingMlChartChoice {
    pub(super) fn start(&mut self, local: &str, depth: usize, empty: bool) -> bool {
        if self.skipped_depth.is_some() {
            return true;
        }
        if local == "AlternateContent" && !empty {
            self.frames.push((depth, false));
        }
        if local == "Fallback" && self.frames.last().is_some_and(|(_, chart)| *chart) {
            self.skipped_depth = (!empty).then_some(depth);
            return true;
        }
        false
    }
    pub(super) fn chart(&mut self) {
        for (_, chart) in &mut self.frames {
            *chart = true;
        }
    }
    pub(super) fn end(&mut self, local: &str, depth: usize) -> bool {
        if let Some(start) = self.skipped_depth {
            if start == depth {
                self.skipped_depth = None;
            }
            return true;
        }
        if local == "AlternateContent"
            && self.frames.last().is_some_and(|(start, _)| *start == depth)
        {
            self.frames.pop();
        }
        false
    }
    pub(super) fn skipping(&self) -> bool {
        self.skipped_depth.is_some()
    }
}
pub(super) const DEFAULT_TEXT_HORIZONTAL_INSET: f32 = 91_440.0 / EMU_PER_CSS_PIXEL;
const CHART_ACCENTS: [&str; 6] = [
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
];
const CLASSIC_DARK_CHART_PLOT_COLOR: u32 = 0x3f3f_3fff;

#[derive(Clone, Debug)]
pub(super) struct DrawingMlThemeLineStyles {
    widths: Vec<f32>,
    list_depth: Option<usize>,
}

impl Default for DrawingMlThemeLineStyles {
    fn default() -> Self {
        Self {
            widths: [9_525, 25_400, 38_100]
                .map(|width| width as f32 / EMU_PER_CSS_PIXEL)
                .to_vec(),
            list_depth: None,
        }
    }
}

impl DrawingMlThemeLineStyles {
    pub(super) fn start(
        &mut self,
        local: &str,
        attributes: &[XmlAttribute<'_>],
        depth: usize,
        empty: bool,
        part: &str,
    ) -> Result<(), Diagnostic> {
        if local == "lnStyleLst" {
            self.widths.clear();
            self.list_depth = (!empty).then_some(depth);
        } else if local == "ln"
            && self
                .list_depth
                .is_some_and(|start| depth == start.saturating_add(1))
        {
            self.widths.push(
                numeric_attribute(attributes, "w", part)?.unwrap_or(12_700) as f32
                    / EMU_PER_CSS_PIXEL,
            );
        }
        Ok(())
    }

    pub(super) fn end(&mut self, local: &str, depth: usize) {
        if local == "lnStyleLst" && self.list_depth == Some(depth) {
            self.list_depth = None;
        }
    }

    pub(super) fn width(&self, index: u64) -> Option<f32> {
        usize::try_from(index)
            .ok()?
            .checked_sub(1)
            .and_then(|index| self.widths.get(index).copied())
    }

    pub(super) fn referenced_width(
        &self,
        attributes: &[XmlAttribute<'_>],
        part: &str,
    ) -> Result<Option<f32>, Diagnostic> {
        Ok(numeric_attribute(attributes, "idx", part)?.and_then(|index| self.width(index)))
    }
}

include!("drawingml/effects.rs");
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum ChartKind {
    Bar,
    BoxWhisker,
    Stock,
    Surface,
    SurfaceWireframe,
    Line3D,
    PieOfPie(u16),
    Funnel,
    Sunburst,
    Histogram,
    Pareto,
    Waterfall,
    Treemap,
    Line,
    Area,
    Scatter,
    Bubble,
    Radar,
    Pie,
    BarOfPie(u16),
    Doughnut,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub(super) enum ChartGrouping {
    #[default]
    Standard,
    Stacked,
    PercentStacked,
}

pub(super) fn drawingml_text_orientation(value: &str) -> TextOrientation {
    match value {
        "vert" => TextOrientation::Rotated90,
        "vert270" => TextOrientation::Rotated270,
        "wordArtVert" => TextOrientation::StackedLr,
        "eaVert" => TextOrientation::VerticalRl,
        "wordArtVertRtl" => TextOrientation::StackedRl,
        "mongolianVert" => TextOrientation::VerticalLr,
        _ => TextOrientation::Horizontal,
    }
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub(super) enum DrawingMlDashPattern {
    #[default]
    Solid,
    Dot,
    SystemDot,
    Dash,
    SystemDash,
    LongDash,
    DashDot,
    SystemDashDot,
    LongDashDot,
    LongDashDotDot,
    SystemDashDotDot,
}

impl DrawingMlDashPattern {
    pub(super) fn from_attribute(value: Option<&str>) -> Self {
        match value {
            Some("dot") => Self::Dot,
            Some("sysDot") => Self::SystemDot,
            Some("dash") => Self::Dash,
            Some("sysDash") => Self::SystemDash,
            Some("lgDash") => Self::LongDash,
            Some("dashDot") => Self::DashDot,
            Some("sysDashDot") => Self::SystemDashDot,
            Some("lgDashDot") => Self::LongDashDot,
            Some("lgDashDotDot") => Self::LongDashDotDot,
            Some("sysDashDotDot") => Self::SystemDashDotDot,
            _ => Self::Solid,
        }
    }

    pub(super) const fn lengths(self) -> &'static [f32] {
        match self {
            Self::Solid => &[],
            Self::Dot => &[1.0, 3.0],
            Self::SystemDot => &[1.0, 1.0],
            Self::Dash => &[4.0, 3.0],
            Self::SystemDash => &[3.0, 1.0],
            Self::LongDash => &[8.0, 3.0],
            Self::DashDot => &[4.0, 3.0, 1.0, 3.0],
            Self::SystemDashDot => &[3.0, 1.0, 1.0, 1.0],
            Self::LongDashDot => &[8.0, 3.0, 1.0, 3.0],
            Self::LongDashDotDot => &[8.0, 3.0, 1.0, 3.0, 1.0, 3.0],
            Self::SystemDashDotDot => &[3.0, 1.0, 1.0, 1.0, 1.0, 1.0],
        }
    }
}

pub(super) fn drawingml_dash_lengths(pattern: &[f32], stroke_width: f32) -> Vec<f32> {
    pattern
        .iter()
        .map(|length| length * stroke_width.max(1.0))
        .collect()
}

pub(super) fn drawingml_stroke_style(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<(f32, StrokeStyle), Diagnostic> {
    let width =
        numeric_attribute(attributes, "w", part)?.unwrap_or(12_700) as f32 / EMU_PER_CSS_PIXEL;
    let cap = match string_attribute(attributes, "cap", part)?.as_deref() {
        Some("rnd") => LineCap::Round,
        Some("sq") => LineCap::Square,
        _ => LineCap::Flat,
    };
    let compound = match string_attribute(attributes, "cmpd", part)?.as_deref() {
        Some("dbl") => LineCompound::Double,
        Some("thickThin") => LineCompound::ThickThin,
        Some("thinThick") => LineCompound::ThinThick,
        Some("tri") => LineCompound::Triple,
        _ => LineCompound::Single,
    };
    let alignment = match string_attribute(attributes, "algn", part)?.as_deref() {
        Some("in") => LineAlignment::Inset,
        _ => LineAlignment::Center,
    };
    Ok((
        width,
        StrokeStyle {
            cap,
            compound,
            alignment,
            ..StrokeStyle::default()
        },
    ))
}

#[derive(Clone, Debug, Default)]
pub(super) struct ChartValueAxis {
    id: Option<String>,
    pub(super) minimum: Option<f32>,
    pub(super) maximum: Option<f32>,
    pub(super) major_unit: Option<f32>,
    minor_unit: Option<f32>,
    pub(super) display_unit: Option<f32>,
    pub(super) display_unit_label: Option<String>,
    no_multi_level_labels: bool,
    reversed: bool,
    label_skip: Option<usize>,
    tick_skip: Option<usize>,
    log_base: Option<f32>,
    pub(super) number_format: Option<String>,
    pub(super) title: String,
    pub(super) title_font_size: Option<f32>,
    pub(super) title_bold: Option<bool>,
    /// `c:title/c:spPr` fill for the axis title box, e.g. `Solid Fill`.
    pub(super) title_fill: Option<ChartFill>,
    /// `c:title/c:spPr/a:ln` stroke for the axis title box.
    pub(super) title_stroke: Option<ChartFill>,
    pub(super) title_stroke_width: f32,
    pub(super) title_stroke_style: StrokeStyle,
    pub(super) title_effects: DrawingMlPictureEffects,
    /// Rotation authored on the axis title's own `a:bodyPr`, in degrees.
    pub(super) title_rotation_degrees: Option<f32>,
    /// Text direction authored on the axis title's own `a:bodyPr`.
    pub(super) title_orientation: Option<TextOrientation>,
    pub(super) label_font_size: Option<f32>,
    pub(super) label_bold: Option<bool>,
    pub(super) label_color: Option<u32>,
    pub(super) label_rotation_degrees: Option<f32>,
    pub(super) is_date: bool,
    base_time_unit: Option<ChartTimeUnit>,
    major_time_unit: Option<ChartTimeUnit>,
    minor_time_unit: Option<ChartTimeUnit>,
    major_tick_mark: ChartAxisTickMark,
    minor_tick_mark: ChartAxisTickMark,
    minor_gridlines: bool,
    minor_gridline_color: Option<u32>,
    minor_gridline_width: Option<f32>,
    major_gridlines: bool,
    major_gridline_color: Option<u32>,
    major_gridline_width: Option<f32>,
    axis_line_hidden: bool,
    deleted: bool,
    cross_between_categories: bool,
    pub(super) position: Option<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ChartTimeUnit {
    Days,
    Months,
    Years,
}

impl ChartTimeUnit {
    fn from_ooxml(value: &str) -> Option<Self> {
        match value {
            "days" => Some(Self::Days),
            "months" => Some(Self::Months),
            "years" => Some(Self::Years),
            _ => None,
        }
    }
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
enum ChartAxisTickMark {
    #[default]
    None,
    In,
    Out,
    Cross,
}

impl ChartAxisTickMark {
    fn from_ooxml(value: Option<&str>) -> Self {
        match value {
            Some("in") => Self::In,
            Some("out") => Self::Out,
            Some("cross") => Self::Cross,
            _ => Self::None,
        }
    }
}

#[derive(Clone, Copy, Debug, Default)]
pub(super) struct ChartDataTable {
    pub(super) show_horizontal_borders: bool,
    pub(super) show_vertical_borders: bool,
    pub(super) show_outline: bool,
    pub(super) show_keys: bool,
}

pub(super) struct ChartDataTableLine {
    pub(super) bounds: Rect,
    pub(super) geometry: Geometry,
    pub(super) color: u32,
    pub(super) width: f32,
    pub(super) series_index: Option<usize>,
}

pub(super) struct ChartDataTableText {
    pub(super) bounds: Rect,
    pub(super) text: String,
    pub(super) align: TextAlign,
    pub(super) series_index: Option<usize>,
}

pub(super) struct ChartDataTableLayout {
    pub(super) lines: Vec<ChartDataTableLine>,
    pub(super) texts: Vec<ChartDataTableText>,
}

#[derive(Clone, Debug)]
pub(super) struct ChartUpDownBars {
    pub(super) gap_width: u16,
    pub(super) up_fill: Option<ChartFill>,
    pub(super) down_fill: Option<ChartFill>,
    pub(super) up_stroke: Option<ChartFill>,
    pub(super) down_stroke: Option<ChartFill>,
    series_start: usize,
    series_end: usize,
}

impl Default for ChartUpDownBars {
    fn default() -> Self {
        Self {
            gap_width: 150,
            up_fill: None,
            down_fill: None,
            up_stroke: None,
            down_stroke: None,
            series_start: 0,
            series_end: 0,
        }
    }
}

#[derive(Clone, Debug)]
pub(super) struct Chart {
    pub(super) source_part: String,
    pub(super) date_1904: bool,
    pub(super) series: Vec<ChartSeries>,
    pub(super) show_title: bool,
    pub(super) title: String,
    pub(super) title_position: Option<[(f32, bool); 2]>,
    pub(super) title_fill: Option<ChartFill>,
    pub(super) title_font_size: Option<f32>,
    pub(super) title_font_bold: Option<bool>,
    pub(super) title_text_color: Option<u32>,
    pub(super) show_legend: bool,
    pub(super) legend_position: ChartLegendPosition,
    pub(super) deleted_legend_entries: Vec<usize>,
    pub(super) chart_area_no_fill: bool,
    pub(super) chart_area_fill: Option<ChartFill>,
    pub(super) chart_area_border: Option<(ChartFill, f32)>,
    pub(super) plot_area_fill: Option<ChartFill>,
    pub(super) legend_fill: Option<ChartFill>,
    pub(super) legend_stroke: Option<ChartFill>,
    pub(super) legend_stroke_width: f32,
    pub(super) plot_area_color: Option<u32>,
    pub(super) plot_area_border_color: Option<u32>,
    pub(super) plot_area_border_width: f32,
    pub(super) data_label_position: Option<String>,
    pub(super) style: Option<ChartStyle>,
    pub(super) classic_defaults: bool,
    pub(super) font_family: Option<String>,
    pub(super) font_size: Option<f32>,
    pub(super) font_bold: bool,
    pub(super) font_scale_basis: Option<(f32, f32)>,
    pub(super) native_size: Option<(f32, f32)>,
    pub(super) value_axis_options: ChartValueAxis,
    pub(super) secondary_value_axis_options: Option<ChartValueAxis>,
    pub(super) horizontal_axis_options: ChartValueAxis,
    pub(super) series_axis_options: Option<ChartValueAxis>,
    pub(super) scatter_has_lines: bool,
    pub(super) series_lines: bool,
    pub(super) data_table: Option<ChartDataTable>,
    pub(super) up_down_bars: Option<ChartUpDownBars>,
    pub(super) plot_bounds: Option<Rect>,
    pub(super) legend_bounds: Option<Rect>,
    pub(super) view_3d: Option<ChartView3D>,
    pub(super) surface_band_fills: Vec<Option<ChartFill>>,
    pub(super) bar_gap_width_percent: f32,
    pub(super) category_gap_width: Option<f32>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) enum ChartLegendPosition {
    Bottom,
    Left,
    #[default]
    Right,
    Top,
    TopRight,
}

impl ChartLegendPosition {
    fn from_ooxml(value: Option<&str>) -> Self {
        match value {
            Some("b") => Self::Bottom,
            Some("l") => Self::Left,
            Some("t") => Self::Top,
            Some("tr") => Self::TopRight,
            _ => Self::Right,
        }
    }

    pub(super) fn horizontal(self) -> bool {
        matches!(self, Self::Bottom | Self::Top)
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub(super) struct ChartView3D {
    pub(super) rot_x: i16,
    pub(super) rot_y: u16,
    pub(super) right_angle_axes: bool,
    pub(super) perspective: u16,
    pub(super) height_percent: Option<u16>,
    pub(super) depth_percent: Option<u16>,
}

impl Default for ChartView3D {
    fn default() -> Self {
        Self {
            rot_x: 15,
            rot_y: 20,
            right_angle_axes: true,
            perspective: 30,
            height_percent: None,
            depth_percent: None,
        }
    }
}

impl ChartView3D {
    pub(super) fn pie_vertical_ratio(self) -> f32 {
        let tilt = (f32::from(self.rot_x.clamp(-90, 90))
            .abs()
            .to_radians()
            .sin())
        .max(0.12);
        let perspective = f32::from(self.perspective.min(240)) / 240.0;
        if let Some(height_percent) = self.height_percent {
            let total_height = f32::from(height_percent.clamp(5, 500)) / 50.0;
            return (tilt * (1.0 - perspective * 0.20)).min(total_height * 0.475);
        }
        tilt * (1.0 - perspective * 0.20)
    }

    pub(super) fn pie_depth_ratio(self) -> f32 {
        if let Some(height_percent) = self.height_percent {
            let total_height = f32::from(height_percent.clamp(5, 500)) / 50.0;
            return (total_height - 2.0 * self.pie_vertical_ratio()).max(0.05);
        }
        0.22
    }

    pub(super) fn pie_depth_perspective(self) -> f32 {
        if self.right_angle_axes {
            0.0
        } else {
            f32::from(self.perspective.min(240)) / 240.0
        }
    }
}

#[derive(Clone, Copy, Debug)]
pub(super) struct ChartTextStyle {
    pub(super) color: u32,
    pub(super) font_size: f32,
    pub(super) bold: bool,
    pub(super) align: TextAlign,
    pub(super) shadow: Option<Shadow>,
}

#[derive(Clone, Debug)]
pub(super) struct ChartDataLabel {
    pub(super) index: usize,
    pub(super) text: String,
    pub(super) text_color: Option<u32>,
    pub(super) font_size: Option<f32>,
    pub(super) font_bold: Option<bool>,
    pub(super) manual_offset: Option<(f32, f32)>,
    pub(super) manual_size: Option<(f32, f32)>,
    pub(super) border_color: Option<u32>,
    pub(super) border_width: f32,
    pub(super) show_values: Option<bool>,
    pub(super) show_category_name: Option<bool>,
    pub(super) show_percent: Option<bool>,
    pub(super) number_format: Option<String>,
    pub(super) position: Option<String>,
}

#[derive(Debug, Default)]
struct ChartDataLabelCapture {
    text: String,
    text_color: Option<u32>,
    font_size: Option<f32>,
    font_bold: Option<bool>,
    manual_offset: [Option<f32>; 2],
    manual_offset_edge: [bool; 2],
    manual_size: [Option<f32>; 2],
    border_color: Option<u32>,
    border_width: f32,
    show_values: Option<bool>,
    show_category_name: Option<bool>,
    show_percent: Option<bool>,
    number_format: Option<String>,
    position: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub(super) enum ChartFill {
    MappedGradient {
        fill: Box<ChartFill>,
        mapping: DrawingMlGradientMapping,
    },
    Solid(u32),
    Pattern {
        preset: String,
        foreground: u32,
        background: u32,
    },
    LinearGradient {
        angle_degrees: f32,
        angle_scaled: bool,
        stops: Vec<GradientStop>,
    },
    PathGradient {
        circular: bool,
        fill_to_rectangle: Option<DrawingMlRelativeRectangle>,
        stops: Vec<GradientStop>,
    },
    ShapeGradient {
        focus: Option<DrawingMlRelativeRectangle>,
        stops: Vec<GradientStop>,
    },
    Image(Paint),
}

impl ChartFill {
    pub(super) fn paint(&self, bounds: Rect) -> Paint {
        match self {
            Self::MappedGradient { fill, mapping } => mapping.clone().paint(fill.paint(bounds)),
            Self::Solid(color) => Paint::Solid(*color),
            Self::Pattern {
                preset,
                foreground,
                background,
            } => Paint::Pattern {
                preset: preset.clone(),
                foreground: *foreground,
                background: *background,
            },
            Self::LinearGradient {
                angle_degrees,
                angle_scaled,
                stops,
            } => drawingml_linear_gradient(bounds, *angle_degrees, *angle_scaled, stops.clone()),
            Self::PathGradient {
                circular,
                fill_to_rectangle,
                stops,
            } => drawingml_path_gradient(bounds, *circular, *fill_to_rectangle, stops.clone()),
            Self::ShapeGradient { focus, stops } => drawingml_shape_gradient(*focus, stops.clone()),
            Self::Image(paint) => paint.clone(),
        }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub(super) struct DrawingMlRelativeRectangle {
    pub(super) left: f32,
    pub(super) top: f32,
    pub(super) right: f32,
    pub(super) bottom: f32,
}

#[derive(Clone, Debug, PartialEq)]
pub(super) struct DrawingMlGradientMapping {
    tile: ImageCrop,
    flip: crate::model::TileMode,
    rotate_with_shape: bool,
}
impl Default for DrawingMlGradientMapping {
    fn default() -> Self {
        Self {
            tile: ImageCrop::default(),
            flip: crate::model::TileMode::None,
            rotate_with_shape: true,
        }
    }
}
impl DrawingMlGradientMapping {
    pub(super) fn start(
        &mut self,
        local: &str,
        attributes: &[XmlAttribute<'_>],
        part: &str,
    ) -> Result<(), Diagnostic> {
        if local == "gradFill" {
            self.rotate_with_shape = string_attribute(attributes, "rotWithShape", part)?
                .is_none_or(|v| !matches!(v.as_str(), "0" | "false"));
            self.flip = match string_attribute(attributes, "flip", part)?.as_deref() {
                Some("x") => crate::model::TileMode::FlipX,
                Some("y") => crate::model::TileMode::FlipY,
                Some("xy") => crate::model::TileMode::FlipXY,
                _ => crate::model::TileMode::None,
            };
        } else if local == "tileRect" {
            let rect = drawingml_relative_rectangle(attributes, part)?;
            if rect.left + rect.right < 1.0 && rect.top + rect.bottom < 1.0 {
                self.tile = ImageCrop {
                    left: rect.left,
                    top: rect.top,
                    right: rect.right,
                    bottom: rect.bottom,
                };
            }
        }
        Ok(())
    }
    pub(super) fn paint(mut self, paint: Paint) -> Paint {
        // Preserve Office's path-gradient anchor behavior; tileRect maps linear ramps.
        if matches!(
            paint,
            Paint::CircleGradient { .. } | Paint::RectGradient { .. } | Paint::ShapeGradient { .. }
        ) {
            self.tile = ImageCrop::default();
        }
        if self.tile == ImageCrop::default() && self.rotate_with_shape {
            return paint;
        }
        Paint::MappedGradient {
            paint: Box::new(paint),
            tile: self.tile,
            flip: self.flip,
            rotate_with_shape: self.rotate_with_shape,
        }
    }
}

pub(super) fn drawingml_image_fill_mapping(
    mapping: &mut crate::model::ImageFillMapping,
    local: &str,
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<(), Diagnostic> {
    match local {
        "blipFill" => {
            mapping.rotate_with_shape = string_attribute(attributes, "rotWithShape", part)?
                .is_none_or(|v| !matches!(v.as_str(), "0" | "false"));
            mapping.dpi = float_attribute(attributes, "dpi", part)?
                .unwrap_or(0.0)
                .max(0.0);
        }
        "tile" => {
            mapping.scale_x = percentage_attribute(attributes, "sx", part)?
                .unwrap_or(1.0)
                .max(f32::EPSILON);
            mapping.scale_y = percentage_attribute(attributes, "sy", part)?
                .unwrap_or(1.0)
                .max(f32::EPSILON);
            mapping.offset_x =
                signed_numeric_attribute(attributes, "tx", part)?.unwrap_or(0) as f32 / 9525.0;
            mapping.offset_y =
                signed_numeric_attribute(attributes, "ty", part)?.unwrap_or(0) as f32 / 9525.0;
            let align =
                string_attribute(attributes, "algn", part)?.unwrap_or_else(|| "tl".to_owned());
            mapping.alignment_x = match align.as_str() {
                "t" | "ctr" | "b" => 0.5,
                "tr" | "r" | "br" => 1.0,
                _ => 0.0,
            };
            mapping.alignment_y = match align.as_str() {
                "l" | "ctr" | "r" => 0.5,
                "bl" | "b" | "br" => 1.0,
                _ => 0.0,
            };
            mapping.flip = match string_attribute(attributes, "flip", part)?.as_deref() {
                Some("x") => crate::model::TileMode::FlipX,
                Some("y") => crate::model::TileMode::FlipY,
                Some("xy") => crate::model::TileMode::FlipXY,
                _ => crate::model::TileMode::None,
            };
        }
        "fillRect" => {
            mapping.fill_rectangle = ImageCrop {
                left: percentage_attribute(attributes, "l", part)?.unwrap_or(0.0),
                top: percentage_attribute(attributes, "t", part)?.unwrap_or(0.0),
                right: percentage_attribute(attributes, "r", part)?.unwrap_or(0.0),
                bottom: percentage_attribute(attributes, "b", part)?.unwrap_or(0.0),
            }
        }
        _ => {}
    }
    Ok(())
}

pub(super) fn drawingml_relative_rectangle(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<DrawingMlRelativeRectangle, Diagnostic> {
    let percentage = |name| -> Result<f32, Diagnostic> {
        Ok(signed_numeric_attribute(attributes, name, part)?.unwrap_or(0) as f32 / 100_000.0)
    };
    Ok(DrawingMlRelativeRectangle {
        left: percentage("l")?,
        top: percentage("t")?,
        right: percentage("r")?,
        bottom: percentage("b")?,
    })
}

pub(super) fn drawingml_shape_gradient(
    focus: Option<DrawingMlRelativeRectangle>,
    stops: Vec<GradientStop>,
) -> Paint {
    let focus = focus.unwrap_or(DrawingMlRelativeRectangle {
        left: 0.5,
        top: 0.5,
        right: 0.5,
        bottom: 0.5,
    });
    Paint::ShapeGradient {
        focus: ImageCrop {
            left: focus.left,
            top: focus.top,
            right: focus.right,
            bottom: focus.bottom,
        },
        stops,
    }
}

pub(super) fn drawingml_path_gradient(
    bounds: Rect,
    circular: bool,
    fill_to_rectangle: Option<DrawingMlRelativeRectangle>,
    stops: Vec<GradientStop>,
) -> Paint {
    let outer_center_x = bounds.width / 2.0;
    let outer_center_y = bounds.height / 2.0;
    let outer_radius = outer_center_x.hypot(outer_center_y);
    let (anchor_left, anchor_top, anchor_right, anchor_bottom) = if circular {
        (
            outer_center_x - outer_radius,
            outer_center_y - outer_radius,
            outer_center_x + outer_radius,
            outer_center_y + outer_radius,
        )
    } else {
        (0.0, 0.0, bounds.width, bounds.height)
    };
    let anchor_width = anchor_right - anchor_left;
    let anchor_height = anchor_bottom - anchor_top;
    let (focus_left, focus_top, focus_right, focus_bottom) = fill_to_rectangle.map_or_else(
        || {
            (
                outer_center_x,
                outer_center_y,
                outer_center_x,
                outer_center_y,
            )
        },
        |focus| {
            (
                anchor_left + anchor_width * focus.left,
                anchor_top + anchor_height * focus.top,
                anchor_right - anchor_width * focus.right,
                anchor_bottom - anchor_height * focus.bottom,
            )
        },
    );
    let center_x = (focus_left + focus_right) / 2.0;
    let center_y = (focus_top + focus_bottom) / 2.0;
    if !circular {
        return Paint::RectGradient {
            center_x,
            center_y,
            stops,
        };
    }
    let inner_width = focus_right - focus_left;
    let inner_height = focus_bottom - focus_top;
    let origin_x = if inner_width > 0.0 && (anchor_width - inner_width).abs() > f32::EPSILON {
        focus_left + inner_width * (focus_left - anchor_left) / (anchor_width - inner_width)
    } else {
        focus_left
    };
    let origin_y = if inner_height > 0.0 && (anchor_height - inner_height).abs() > f32::EPSILON {
        focus_top + inner_height * (focus_top - anchor_top) / (anchor_height - inner_height)
    } else {
        focus_top
    };
    Paint::CircleGradient {
        x0: origin_x,
        y0: origin_y,
        r0: inner_width.abs().min(inner_height.abs()) / 2.0,
        x1: outer_center_x,
        y1: outer_center_y,
        r1: outer_radius,
        stops,
    }
}

pub(super) fn drawingml_linear_gradient(
    bounds: Rect,
    angle_degrees: f32,
    angle_scaled: bool,
    stops: Vec<GradientStop>,
) -> Paint {
    let radians = angle_degrees.to_radians();
    let cosine = radians.cos();
    let sine = radians.sin();
    let width = bounds.width.max(f32::EPSILON);
    let height = bounds.height.max(f32::EPSILON);
    let (gradient_x, gradient_y) = if angle_scaled {
        let span = cosine.abs() + sine.abs();
        (cosine / (width * span), sine / (height * span))
    } else {
        let span = cosine.abs() * width + sine.abs() * height;
        (cosine / span, sine / span)
    };
    let magnitude_squared = gradient_x * gradient_x + gradient_y * gradient_y;
    let dx = gradient_x / magnitude_squared;
    let dy = gradient_y / magnitude_squared;
    Paint::LinearGradient {
        x0: bounds.width / 2.0 - dx / 2.0,
        y0: bounds.height / 2.0 - dy / 2.0,
        x1: bounds.width / 2.0 + dx / 2.0,
        y1: bounds.height / 2.0 + dy / 2.0,
        stops,
    }
}

fn chart_text_width(text: &str, font_size: f32) -> f32 {
    text.chars()
        .map(|character| match character {
            ' ' => font_size * 0.28,
            character if character.is_ascii_punctuation() => font_size * 0.35,
            character if character.is_ascii_uppercase() => font_size * 0.55,
            character if character.is_ascii_lowercase() => font_size * 0.40,
            character if character.is_ascii() => font_size * 0.50,
            _ => font_size,
        })
        .sum::<f32>()
}

pub(super) fn chart_title_bounds(
    chart_bounds: Rect,
    y: f32,
    height: f32,
    text: &str,
    font_size: f32,
) -> Rect {
    let text_width = chart_text_width(text, font_size);
    let maximum_width = chart_bounds.width * 0.8;
    let width = (text_width + DEFAULT_TEXT_HORIZONTAL_INSET * 2.0 + font_size * 0.1)
        .max(height)
        .min(maximum_width);
    Rect {
        x: chart_bounds.x + (chart_bounds.width - width) / 2.0,
        y,
        width,
        height,
    }
}

pub(super) fn drawingml_fallback_character_width(character: char, font_size: f32) -> f32 {
    match character {
        ' ' => font_size * 0.28,
        character if character.is_ascii_punctuation() => font_size * 0.35,
        character if character.is_ascii() => font_size * 0.52,
        _ => font_size,
    }
}

#[derive(Debug)]
enum ChartFillCaptureKind {
    Solid,
    Gradient,
    Pattern,
    Image,
}

#[derive(Debug)]
pub(super) struct ChartFillCapture {
    gradient_mapping: DrawingMlGradientMapping,
    depth: usize,
    kind: ChartFillCaptureKind,
    color: Option<u32>,
    pattern_preset: Option<String>,
    pattern_foreground: Option<u32>,
    pattern_background: Option<u32>,
    pattern_foreground_active: bool,
    stops: Vec<GradientStop>,
    current_stop: Option<(f32, Option<u32>)>,
    angle_degrees: f32,
    angle_scaled: bool,
    gradient_path: Option<bool>,
    shape_gradient: bool,
    fill_to_rectangle: Option<DrawingMlRelativeRectangle>,
    image_relationship_id: Option<String>,
    image_crop: ImageCrop,
    image_tile: bool,
    image_mapping: crate::model::ImageFillMapping,
}

impl ChartFillCapture {
    pub(super) fn new(local: &str, depth: usize) -> Option<Self> {
        let kind = match local {
            "solidFill" => ChartFillCaptureKind::Solid,
            "gradFill" => ChartFillCaptureKind::Gradient,
            "pattFill" => ChartFillCaptureKind::Pattern,
            "blipFill" => ChartFillCaptureKind::Image,
            _ => return None,
        };
        Some(Self {
            gradient_mapping: DrawingMlGradientMapping::default(),
            depth,
            kind,
            color: None,
            pattern_preset: None,
            pattern_foreground: None,
            pattern_background: None,
            pattern_foreground_active: false,
            stops: Vec::new(),
            current_stop: None,
            angle_degrees: 0.0,
            angle_scaled: false,
            gradient_path: None,
            shape_gradient: false,
            fill_to_rectangle: None,
            image_relationship_id: None,
            image_crop: ImageCrop::default(),
            image_tile: false,
            image_mapping: crate::model::ImageFillMapping::default(),
        })
    }

    pub(super) fn closes_at(&self, depth: usize) -> bool {
        self.depth == depth
    }

    #[inline(never)]
    pub(super) fn start(
        &mut self,
        local: &str,
        attributes: &[XmlAttribute<'_>],
        part: &str,
        scheme_color: &dyn Fn(&str) -> Option<u32>,
    ) -> Result<(), Diagnostic> {
        if matches!(self.kind, ChartFillCaptureKind::Gradient) {
            self.gradient_mapping.start(local, attributes, part)?;
        }
        if matches!(self.kind, ChartFillCaptureKind::Image) {
            drawingml_image_fill_mapping(&mut self.image_mapping, local, attributes, part)?;
        }
        match local {
            "pattFill" => {
                self.pattern_preset = string_attribute(attributes, "prst", part)?;
            }
            "fgClr" => self.pattern_foreground_active = true,
            "bgClr" => self.pattern_foreground_active = false,
            "gs" => {
                let offset = percentage_attribute(attributes, "pos", part)?.unwrap_or(0.0);
                self.current_stop = Some((offset.clamp(0.0, 1.0), None));
            }
            "srgbClr" | "schemeClr" => {
                let value = string_attribute(attributes, "val", part)?;
                let color = match local {
                    "srgbClr" => value.as_deref().and_then(parse_rgb_color),
                    _ => value.as_deref().and_then(scheme_color),
                };
                if let Some(color) = color {
                    if let Some((_, stop_color)) = self.current_stop.as_mut() {
                        *stop_color = Some(color);
                    } else if matches!(self.kind, ChartFillCaptureKind::Pattern) {
                        if self.pattern_foreground_active {
                            self.pattern_foreground = Some(color);
                        } else {
                            self.pattern_background = Some(color);
                        }
                    } else {
                        self.color = Some(color);
                    }
                }
            }
            "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff" => {
                let ratio = percentage_attribute(attributes, "val", part)?.unwrap_or(1.0);
                if let Some((_, Some(color))) = self.current_stop.as_mut() {
                    apply_color_transform(color, local, ratio);
                } else if matches!(self.kind, ChartFillCaptureKind::Pattern) {
                    let color = if self.pattern_foreground_active {
                        self.pattern_foreground.as_mut()
                    } else {
                        self.pattern_background.as_mut()
                    };
                    if let Some(color) = color {
                        apply_color_transform(color, local, ratio);
                    }
                } else if let Some(color) = self.color.as_mut() {
                    apply_color_transform(color, local, ratio);
                }
            }
            "lin" => {
                self.angle_degrees = string_attribute(attributes, "ang", part)?
                    .map(|value| {
                        value
                            .parse::<i64>()
                            .map(|value| value as f32 / 60_000.0)
                            .map_err(|_| {
                                format_error(part, "chart title gradient angle is not an integer")
                            })
                    })
                    .transpose()?
                    .unwrap_or(0.0);
                self.angle_scaled = string_attribute(attributes, "scaled", part)?
                    .is_some_and(|value| matches!(value.as_str(), "1" | "true" | "on"));
            }
            "path" => {
                self.shape_gradient =
                    string_attribute(attributes, "path", part)?.as_deref() == Some("shape");
                self.gradient_path = match string_attribute(attributes, "path", part)?.as_deref() {
                    Some("circle") => Some(true),
                    Some("rect") => Some(false),
                    _ => None,
                };
            }
            "fillToRect" if self.gradient_path.is_some() || self.shape_gradient => {
                self.fill_to_rectangle = Some(drawingml_relative_rectangle(attributes, part)?);
            }
            "blip" => {
                self.image_relationship_id = string_attribute(attributes, "embed", part)?;
            }
            "srcRect" => {
                let crop = ImageCrop {
                    left: percentage_attribute(attributes, "l", part)?.unwrap_or(0.0),
                    top: percentage_attribute(attributes, "t", part)?.unwrap_or(0.0),
                    right: percentage_attribute(attributes, "r", part)?.unwrap_or(0.0),
                    bottom: percentage_attribute(attributes, "b", part)?.unwrap_or(0.0),
                };
                if !crop.is_valid() {
                    return Err(format_error(part, "chart title image-fill crop is invalid"));
                }
                self.image_crop = crop;
            }
            "tile" => self.image_tile = true,
            _ => {}
        }
        Ok(())
    }

    pub(super) fn end(&mut self, local: &str) {
        if local == "gs"
            && let Some((offset, Some(color))) = self.current_stop.take()
        {
            self.stops.push(GradientStop { offset, color });
        }
    }

    pub(super) fn finish(
        self,
        package: &Package<'_>,
        part: &str,
    ) -> Result<Option<ChartFill>, Diagnostic> {
        if !matches!(&self.kind, ChartFillCaptureKind::Image) {
            return Ok(self.finish_without_image());
        }
        let Some(relationship_id) = self.image_relationship_id else {
            return Ok(None);
        };
        let relationships = package.relationships(Some(part))?;
        let Some(relationship) = relationships
            .iter()
            .find(|relationship| relationship.id == relationship_id)
        else {
            return Err(format_error(
                part,
                format!("chart title image-fill relationship {relationship_id} does not exist"),
            ));
        };
        if relationship.external {
            return Ok(None);
        }
        if !relationship.type_uri.ends_with("/image") {
            return Err(format_error(
                part,
                format!("chart title relationship {relationship_id} is not an image"),
            ));
        }
        let bytes = package.required_part(&relationship.target)?;
        let Ok(media_type) = office_image_media_type(&relationship.target, &bytes)
            .or_else(|_| office_image_media_type_from_signature(&bytes))
        else {
            return Ok(None);
        };
        Ok(Some(ChartFill::Image(Paint::Image {
            mapping: (self.image_tile
                || self.image_mapping != crate::model::ImageFillMapping::default())
            .then(|| Box::new(self.image_mapping)),
            media_type: media_type.to_owned(),
            bytes: bytes.into_vec(),
            crop: self.image_crop,
            tile: self.image_tile,
            tile_width: None,
            tile_height: None,
        })))
    }

    pub(super) fn finish_without_image(mut self) -> Option<ChartFill> {
        match self.kind {
            ChartFillCaptureKind::Solid => self.color.map(ChartFill::Solid),
            ChartFillCaptureKind::Gradient => {
                self.stops
                    .sort_by(|left, right| left.offset.total_cmp(&right.offset));
                let fill = (!self.stops.is_empty()).then_some(if self.shape_gradient {
                    ChartFill::ShapeGradient {
                        focus: self.fill_to_rectangle,
                        stops: self.stops,
                    }
                } else if let Some(circular) = self.gradient_path {
                    ChartFill::PathGradient {
                        circular,
                        fill_to_rectangle: self.fill_to_rectangle,
                        stops: self.stops,
                    }
                } else {
                    ChartFill::LinearGradient {
                        angle_degrees: self.angle_degrees,
                        angle_scaled: self.angle_scaled,
                        stops: self.stops,
                    }
                });
                fill.map(|fill| {
                    if self.gradient_mapping == DrawingMlGradientMapping::default() {
                        fill
                    } else {
                        ChartFill::MappedGradient {
                            fill: Box::new(fill),
                            mapping: self.gradient_mapping,
                        }
                    }
                })
            }
            ChartFillCaptureKind::Pattern => self.pattern_preset.map(|preset| ChartFill::Pattern {
                preset,
                foreground: self.pattern_foreground.unwrap_or(0x0000_00ff),
                background: self.pattern_background.unwrap_or(0xffff_ffff),
            }),
            ChartFillCaptureKind::Image => None,
        }
    }
}

#[derive(Clone, Debug)]
pub(super) struct ChartStyle {
    chart_area_stops: Vec<GradientStop>,
    waterfall_stops: [Vec<GradientStop>; 3],
    pub(super) waterfall_colors: [u32; 3],
    pub(super) chart_border_color: Option<u32>,
    pub(super) data_point_shadow: Option<Shadow>,
    pub(super) axis_color: u32,
    pub(super) grid_color: u32,
    pub(super) connector_color: u32,
    pub(super) title: ChartTextStyle,
    pub(super) legend: ChartTextStyle,
    pub(super) axis_text: ChartTextStyle,
    pub(super) data_label: ChartTextStyle,
}

impl ChartStyle {
    fn office_372(colors: [u32; 3], scheme_color: &impl Fn(&str) -> Option<u32>) -> Self {
        let transformed = |mut color: u32, transforms: &[(&str, f32)]| {
            for (kind, ratio) in transforms {
                apply_color_transform(&mut color, kind, *ratio);
            }
            color
        };
        let dk1 = scheme_color("dk1").unwrap_or(0x0000_00ff);
        let lt1 = scheme_color("lt1").unwrap_or(0xffff_ffff);
        let chart_center = transformed(dk1, &[("lumMod", 0.65), ("lumOff", 0.35)]);
        let chart_edge = transformed(dk1, &[("lumMod", 0.85), ("lumOff", 0.15)]);
        let interpolate = |ratio: f32| {
            let channel = |shift| {
                let start = ((chart_center >> shift) & 0xff_u32) as f32;
                let end = ((chart_edge >> shift) & 0xff_u32) as f32;
                (start + (end - start) * ratio).round() as u32
            };
            (channel(24) << 24) | (channel(16) << 16) | (channel(8) << 8) | channel(0)
        };
        let waterfall_stops = colors.map(|color| {
            vec![
                GradientStop {
                    offset: 0.0,
                    color: transformed(
                        color,
                        &[("satMod", 1.03), ("lumMod", 1.02), ("tint", 0.94)],
                    ),
                },
                GradientStop {
                    offset: 0.5,
                    color: transformed(color, &[("satMod", 1.10), ("shade", 1.0)]),
                },
                GradientStop {
                    offset: 1.0,
                    color: transformed(
                        color,
                        &[("lumMod", 0.99), ("satMod", 1.20), ("shade", 0.78)],
                    ),
                },
            ]
        });
        let axis_text_color = transformed(lt1, &[("lumMod", 0.85)]);
        Self {
            // Office's circular path gradient advances non-linearly across the filled area;
            // sample that curve into ordinary scene stops so every host renderer shares it.
            chart_area_stops: (0..=8)
                .map(|index| {
                    let offset = index as f32 / 8.0;
                    GradientStop {
                        offset,
                        color: interpolate(offset.powf(1.6)),
                    }
                })
                .collect(),
            waterfall_stops,
            waterfall_colors: colors,
            chart_border_color: None,
            data_point_shadow: Some(Shadow {
                color: 0x0000_00a1,
                blur: 6.0,
                offset_x: 0.0,
                offset_y: 2.0,
            }),
            axis_color: transformed(lt1, &[("lumMod", 0.95), ("alpha", 0.54)]),
            grid_color: transformed(lt1, &[("lumMod", 0.95), ("alpha", 0.10)]),
            connector_color: colors[2],
            title: ChartTextStyle {
                color: transformed(lt1, &[("lumMod", 0.95)]),
                font_size: 16.0 * 96.0 / 72.0,
                bold: true,
                align: TextAlign::Center,
                shadow: Some(Shadow {
                    color: 0x0000_0066,
                    blur: 5.33,
                    offset_x: 0.0,
                    offset_y: 4.0,
                }),
            },
            legend: ChartTextStyle {
                color: axis_text_color,
                font_size: 12.0,
                bold: false,
                align: TextAlign::Start,
                shadow: None,
            },
            axis_text: ChartTextStyle {
                color: axis_text_color,
                font_size: 12.0,
                bold: false,
                align: TextAlign::Center,
                shadow: None,
            },
            data_label: ChartTextStyle {
                color: 0x4040_40ff,
                font_size: 12.0,
                bold: false,
                align: TextAlign::Center,
                shadow: None,
            },
        }
    }

    fn office_395(colors: [u32; 3], scheme_color: &impl Fn(&str) -> Option<u32>) -> Self {
        let tx1 = scheme_color("tx1").unwrap_or(0x0000_00ff);
        let bg1 = scheme_color("bg1").unwrap_or(0xffff_ffff);
        let border = transform_luminance(tx1, 0.15, 0.85);
        let text = transform_luminance(tx1, 0.65, 0.35);
        let solid_stops = colors.map(|color| {
            vec![
                GradientStop { offset: 0.0, color },
                GradientStop { offset: 1.0, color },
            ]
        });
        let text_style = ChartTextStyle {
            color: text,
            font_size: 12.0,
            bold: false,
            align: TextAlign::Center,
            shadow: None,
        };
        Self {
            chart_area_stops: vec![
                GradientStop {
                    offset: 0.0,
                    color: bg1,
                },
                GradientStop {
                    offset: 1.0,
                    color: bg1,
                },
            ],
            waterfall_stops: solid_stops,
            waterfall_colors: colors,
            chart_border_color: Some(border),
            data_point_shadow: None,
            axis_color: border,
            grid_color: border,
            connector_color: 0xd9d9_d9ff,
            title: ChartTextStyle {
                color: text,
                font_size: 14.0 * 96.0 / 72.0,
                bold: false,
                align: TextAlign::Center,
                shadow: None,
            },
            legend: ChartTextStyle {
                align: TextAlign::Start,
                ..text_style
            },
            axis_text: text_style,
            data_label: ChartTextStyle {
                color: transform_luminance(tx1, 0.75, 0.25),
                ..text_style
            },
        }
    }

    pub(super) fn chart_area_paint(&self, width: f32, height: f32) -> Paint {
        let radius = (width / 2.0).hypot(height / 2.0);
        Paint::CircleGradient {
            x0: width / 2.0,
            y0: height / 2.0,
            r0: 0.0,
            x1: width / 2.0,
            y1: height / 2.0,
            r1: radius,
            stops: self.chart_area_stops.clone(),
        }
    }

    pub(super) fn waterfall_paint(&self, role: usize, width: f32, height: f32) -> Paint {
        Paint::LinearGradient {
            x0: width / 2.0,
            y0: 0.0,
            x1: width / 2.0,
            y1: height,
            stops: self.waterfall_stops[role.min(2)].clone(),
        }
    }
}

#[derive(Clone, Debug)]
pub(super) struct ChartErrorBars {
    pub(super) plus: Vec<f32>,
    pub(super) minus: Vec<f32>,
    pub(super) standard_error: bool,
    pub(super) show_plus: bool,
    pub(super) show_minus: bool,
    pub(super) end_caps: bool,
    pub(super) stroke: Option<ChartFill>,
    pub(super) stroke_width: f32,
}

impl Default for ChartErrorBars {
    fn default() -> Self {
        Self {
            plus: Vec::new(),
            minus: Vec::new(),
            standard_error: false,
            show_plus: true,
            show_minus: true,
            end_caps: true,
            stroke: None,
            stroke_width: 0.75,
        }
    }
}

#[derive(Clone, Debug, Default)]
pub(super) struct ChartSourceData {
    pub(super) values: Vec<f32>,
    pub(super) category_levels: Vec<Vec<String>>,
}

#[derive(Clone, Debug)]
pub(super) struct ChartSeries {
    pub(super) kind: ChartKind,
    pub(super) first_slice_angle: f32,
    pub(super) grouping: ChartGrouping,
    pub(super) bar_horizontal: bool,
    pub(super) bar_depth: bool,
    pub(super) bar_cone: bool,
    pub(super) bar_cone_to_max: bool,
    pub(super) bar_cylinder: bool,
    pub(super) bar_gap_depth_percent: f32,
    pub(super) three_d: bool,
    pub(super) axis_id: Option<String>,
    pub(super) name: String,
    pub(super) categories: Vec<String>,
    pub(super) category_levels: Vec<Vec<String>>,
    pub(super) x_values: Vec<f32>,
    pub(super) values: Vec<f32>,
    pub(super) raw_values: Vec<String>,
    pub(super) bubble_sizes: Vec<f32>,
    pub(super) color: Option<u32>,
    pub(super) fill: Option<ChartFill>,
    pub(super) point_colors: Vec<u32>,
    pub(super) point_fills: Vec<Option<ChartFill>>,
    pub(super) point_explosions: Vec<f32>,
    pub(super) effects: DrawingMlPictureEffects,
    pub(super) point_border_colors: Vec<Option<u32>>,
    pub(super) point_border_widths: Vec<f32>,
    pub(super) stroke_width: Option<f32>,
    pub(super) line_visible: bool,
    pub(super) smooth: bool,
    pub(super) marker_symbol: Option<String>,
    pub(super) marker_size: Option<f32>,
    pub(super) subtotals: Vec<usize>,
    pub(super) show_values: bool,
    pub(super) show_category_name: bool,
    pub(super) show_percent: bool,
    pub(super) number_format: Option<String>,
    pub(super) data_label_position: Option<String>,
    pub(super) data_label_text_color: Option<u32>,
    pub(super) data_label_font_size: Option<f32>,
    pub(super) data_label_font_bold: Option<bool>,
    pub(super) data_label_rotation_degrees: Option<f32>,
    pub(super) hidden_labels: Vec<usize>,
    pub(super) data_labels: Vec<ChartDataLabel>,
    pub(super) data_label_border: Option<(u32, f32)>,
    pub(super) linear_trendline: bool,
    pub(super) show_trendline_equation: bool,
    pub(super) show_trendline_r_squared: bool,
    pub(super) trendline_label_offset: Option<(f32, f32)>,
    pub(super) trendline_label_font_size: Option<f32>,
    pub(super) x_error_bars: Option<ChartErrorBars>,
    pub(super) y_error_bars: Option<ChartErrorBars>,
}

pub(super) struct DrawingMlElement {
    pub(super) bounds: Rect,
    pub(super) text: Option<String>,
    pub(super) visual: Visual,
}

include!("drawingml/chart_layout.rs");
#[cfg(feature = "native-formats")]
include!("drawingml/diagram.rs");
#[cfg(feature = "native-formats")]
include!("drawingml/chart_parse.rs");
#[cfg(feature = "native-formats")]
include!("drawingml/diagram_parse.rs");
pub(super) fn apply_color_transform(color: &mut u32, local: &str, ratio: f32) {
    match local {
        "alpha" => {
            *color = (*color & 0xffff_ff00) | (ratio.clamp(0.0, 1.0) * 255.0).round() as u32;
        }
        "tint" | "shade" => {
            let ratio = ratio.clamp(0.0, 1.0);
            let alpha = *color & 0xff;
            let channel = |shift: u32| {
                let component = ((*color >> shift) & 0xff) as f32 / 255.0;
                let linear = srgb_to_linear(component);
                let transformed = if local == "tint" {
                    linear * ratio + 1.0 - ratio
                } else {
                    linear * ratio
                };
                (linear_to_srgb(transformed.clamp(0.0, 1.0)) * 255.0).round() as u32
            };
            *color = (channel(24) << 24) | (channel(16) << 16) | (channel(8) << 8) | alpha;
        }
        "lumMod" => *color = transform_luminance(*color, ratio, 0.0),
        "lumOff" => *color = transform_luminance(*color, 1.0, ratio),
        "satMod" => *color = transform_saturation(*color, ratio, 0.0),
        "satOff" => *color = transform_saturation(*color, 1.0, ratio),
        _ => {}
    }
}

include!("drawingml/chart_geometry.rs");
include!("drawingml/chart_extended.rs");
include!("drawingml/chart_render.rs");
#[allow(clippy::too_many_arguments)]
fn begin_chart(
    chart_kind: ChartKind,
    empty: bool,
    depth: usize,
    completed_len: usize,
    kind: &mut Option<ChartKind>,
    chart_depth: &mut Option<usize>,
    chart_series_start: &mut usize,
    chart_axis_ids: &mut Vec<String>,
) {
    *kind = Some(chart_kind);
    *chart_depth = (!empty).then_some(depth);
    *chart_series_start = completed_len;
    chart_axis_ids.clear();
}

fn is_chart_element(local: &str) -> bool {
    chart_kind(local).is_some()
}

fn chart_kind(local: &str) -> Option<ChartKind> {
    match local {
        "barChart" | "bar3DChart" => Some(ChartKind::Bar),
        "stockChart" => Some(ChartKind::Stock),
        "surfaceChart" | "surface3DChart" => Some(ChartKind::Surface),
        "line3DChart" => Some(ChartKind::Line3D),
        "lineChart" => Some(ChartKind::Line),
        "areaChart" | "area3DChart" => Some(ChartKind::Area),
        "scatterChart" => Some(ChartKind::Scatter),
        "bubbleChart" => Some(ChartKind::Bubble),
        "radarChart" => Some(ChartKind::Radar),
        "pieChart" | "pie3DChart" => Some(ChartKind::Pie),
        "ofPieChart" => Some(ChartKind::BarOfPie(0)),
        "doughnutChart" => Some(ChartKind::Doughnut),
        _ => None,
    }
}

fn set_parsed_number_at(values: &mut Vec<f32>, index: Option<usize>, value: &str) {
    let value = value
        .parse::<f32>()
        .ok()
        .filter(|value| value.is_finite())
        .unwrap_or(f32::NAN);
    set_number_at(values, index, value);
}

fn set_number_at(values: &mut Vec<f32>, index: Option<usize>, value: f32) {
    if let Some(index) = index {
        values.resize(index.saturating_add(1), f32::NAN);
        values[index] = value;
    } else {
        values.push(value);
    }
}

fn set_string_at(values: &mut Vec<String>, index: Option<usize>, value: String) {
    if let Some(index) = index {
        values.resize(index.saturating_add(1), String::new());
        values[index] = value;
    } else {
        values.push(value);
    }
}

fn transform_saturation(color: u32, multiplier: f32, offset: f32) -> u32 {
    let (hue, saturation, luminance) = color_to_hsl(color);
    color_from_hsl(
        color,
        hue,
        (saturation * multiplier + offset).clamp(0.0, 1.0),
        luminance,
    )
}

fn srgb_to_linear(component: f32) -> f32 {
    if component <= 0.040_45 {
        component / 12.92
    } else {
        ((component + 0.055) / 1.055).powf(2.4)
    }
}

pub(super) fn linear_to_srgb(component: f32) -> f32 {
    if component <= 0.003_130_8 {
        component * 12.92
    } else {
        1.055 * component.powf(1.0 / 2.4) - 0.055
    }
}

pub(super) fn numeric_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<u64>, Diagnostic> {
    string_attribute(attributes, name, part)?
        .map(|value| {
            value.parse::<u64>().map_err(|_| {
                format_error(
                    part,
                    format!("attribute {name} is not a non-negative integer"),
                )
            })
        })
        .transpose()
}

pub(super) fn signed_numeric_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<i64>, Diagnostic> {
    string_attribute(attributes, name, part)?
        .map(|value| {
            value.parse::<i64>().map_err(|_| {
                format_error(part, format!("attribute {name} is not a signed integer"))
            })
        })
        .transpose()
}

fn float_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<f32>, Diagnostic> {
    string_attribute(attributes, name, part)?
        .map(|value| {
            value
                .parse::<f32>()
                .ok()
                .filter(|value| value.is_finite())
                .ok_or_else(|| format_error(part, format!("attribute {name} is not finite")))
        })
        .transpose()
}

fn percentage_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<f32>, Diagnostic> {
    string_attribute(attributes, name, part)?
        .map(|value| {
            if let Some(value) = value.strip_suffix('%') {
                value
                    .parse::<f32>()
                    .map(|value| value / 100.0)
                    .map_err(|_| {
                        format_error(part, format!("attribute {name} is not a percentage"))
                    })
            } else {
                value
                    .parse::<i64>()
                    .map(|value| value as f32 / 100_000.0)
                    .map_err(|_| {
                        format_error(part, format!("attribute {name} is not a percentage"))
                    })
            }
        })
        .transpose()
}

pub(super) fn parse_rgb_color(value: &str) -> Option<u32> {
    let value = value.trim();
    (value.len() == 6)
        .then(|| u32::from_str_radix(value, 16).ok())
        .flatten()
        .map(|rgb| (rgb << 8) | 0xff)
}

fn default_scheme_color(value: &str) -> Option<u32> {
    let rgb = match value {
        "dk1" | "tx1" => 0x000000,
        "lt1" | "bg1" => 0xffffff,
        "dk2" | "tx2" => 0x44546a,
        "lt2" | "bg2" => 0xe7e6e6,
        "accent1" => 0x4472c4,
        "accent2" => 0xed7d31,
        "accent3" => 0xa5a5a5,
        "accent4" => 0xffc000,
        "accent5" => 0x5b9bd5,
        "accent6" => 0x70ad47,
        "hlink" => 0x0563c1,
        "folHlink" => 0x954f72,
        _ => return None,
    };
    Some((rgb << 8) | 0xff)
}

fn format_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message).in_part(part)
}

fn with_part(error: Diagnostic, part: &str) -> Diagnostic {
    if error.location.part.is_some() {
        error
    } else {
        error.in_part(part)
    }
}

#[cfg(test)]
#[cfg(feature = "native-formats")]
include!("drawingml/tests.rs");
