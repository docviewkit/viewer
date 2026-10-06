//! Native XLSX workbook, shared-string, and worksheet parsing.

use super::optional_xml_attribute as optional_attribute;
use std::collections::{HashMap, HashSet};

use crate::calculation::{
    CalculationResults, CellAddress as CalculatedCellAddress, FormulaValue, formula_calls_any,
};
use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
use crate::limits::Limits;
use crate::model::{
    AffineTransform, BlendMode, Document, DocumentFormat, DocumentKind, FillRule, Geometry,
    GradientStop, ImageAdjustment, ImageCrop, MappingQuality, Object, ObjectKind, OuterShadow,
    Paint, PathCommand, Rect, Shadow, SheetAxis, SheetAxisSpan, SheetOrientation,
    SheetPrintSettings, SheetViewMode, SourceLocator, SourceRef, StretchMode, StrokeStyle,
    TextAlign, TextAutoFit, TextDirection, TextHorizontalOverflow, TextLayout, TextRun,
    TextVerticalAlign, ThreeDStyle, TileMode, Unit, UnitKind, Visual, VisualBrushChild,
};
use crate::package::{Package, Relationship};
use crate::xml::{XmlAttribute, XmlEvent, decode_xml_text, parse_ooxml as parse_xml};

use super::drawingml::{
    Chart as XlsxChart, ChartFill, ChartFillCapture, Diagram as XlsxDiagram, DrawingMlDashPattern,
    DrawingMlPictureEffects, DrawingMlPictureEffectsCapture, DrawingMlThemeLineStyles,
    apply_color_transform, drawingml_camera_point, drawingml_dash_lengths,
    drawingml_fallback_character_width, drawingml_fill_reference_has_paint, drawingml_outer_shadow,
    drawingml_outer_shadow_is_identity, parse_basic_chart, parse_chart, parse_chart_ex_with_data,
    parse_diagram, parse_three_d_bevel, parse_three_d_camera, parse_three_d_light,
    parse_three_d_rotation,
};
use super::presentation_image::{
    OfficeImageError, office_image_media_type, office_image_media_type_from_mime,
    reserve_materialized_image_bytes,
};
use super::{
    ContentTypes, DEFAULT_SHEET_COLUMN_WIDTH as COLUMN_WIDTH,
    DEFAULT_SHEET_ROW_HEIGHT as ROW_HEIGHT, clone_materialized_text, default_excel_column_width,
    excel_serial_date, format_general_number, local_name, reserve_materialized_text_bytes,
    transform_luminance,
};

const MAX_COLUMNS: u32 = 16_384;
const MAX_ROWS: u32 = 1_048_576;
const EMU_PER_CSS_PIXEL: f32 = 9_525.0;
const THEME_COLOR_COUNT: usize = 12;

type ThemeColors = [u32; THEME_COLOR_COUNT];

fn contains_bytes(bytes: &[u8], needle: &[u8]) -> bool {
    !needle.is_empty() && bytes.windows(needle.len()).any(|window| window == needle)
}

fn find_bytes(bytes: &[u8], needle: &[u8], start: usize) -> Option<usize> {
    bytes
        .get(start..)?
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|offset| start + offset)
}

#[derive(Debug)]
struct Sheet {
    name: String,
    relationship_id: String,
    hidden: bool,
    print_area: Option<String>,
    print_titles: Option<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct CellAddress {
    column: u32,
    row: u32,
    canonical: String,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct Dimension {
    start_column: u32,
    start_row: u32,
    end_column: u32,
    end_row: u32,
}

#[derive(Clone, Debug)]
struct TableStyle {
    kind: TableStyleKind,
    range: Dimension,
    name: String,
    show_row_stripes: bool,
}

#[derive(Clone, Debug)]
enum TableStyleKind {
    Table,
    Pivot {
        first_data_row: u32,
        row_field_count: u32,
        rows: Vec<PivotRow>,
    },
}

#[derive(Clone, Debug)]
struct PivotRow {
    level: u32,
    grand_total: bool,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum CellValueType {
    Number,
    SharedString,
    InlineString,
    Boolean,
    String,
    Error,
    Date,
}

#[derive(Debug)]
struct CellState {
    depth: usize,
    address: CellAddress,
    style_index: usize,
    value_type: CellValueType,
    value_depth: Option<usize>,
    value_seen: bool,
    value: String,
    inline_string_depth: Option<usize>,
    inline_text_depth: Option<usize>,
    inline_string_seen: bool,
    inline_text: String,
    inline_run_depth: Option<usize>,
    inline_properties_depth: Option<usize>,
    inline_phonetic_depth: Option<usize>,
    inline_run: Option<SharedStringRun>,
    inline_runs: Vec<SharedStringRun>,
    has_formula: bool,
    shared_formula: Option<u32>,
    formula_depth: Option<usize>,
    formula: String,
}

#[derive(Clone, Debug, Default, PartialEq)]
struct SharedString {
    text: String,
    runs: Vec<SharedStringRun>,
}

#[derive(Clone, Debug, Default, PartialEq)]
struct SharedStringRun {
    text: String,
    font_family: Option<String>,
    font_size: Option<f32>,
    color: Option<u32>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strikethrough: Option<bool>,
    baseline_shift: Option<f32>,
}

#[derive(Clone, Copy, Debug)]
struct ColumnSpan {
    start: u32,
    end: u32,
    width: Option<f32>,
    style_index: Option<usize>,
}

#[derive(Clone, Copy, Debug)]
struct RowState {
    depth: usize,
    index: u32,
    style_index: Option<usize>,
}

#[derive(Debug)]
struct AxisMetrics {
    default_size: f32,
    overrides: Vec<(u32, f32)>,
    prefix_delta: Vec<f32>,
}

#[derive(Clone, Debug)]
struct CellStyle {
    font_family: String,
    font_size: f32,
    color: u32,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    baseline_shift: f32,
    fill: u32,
    fill_paint: Option<ChartFill>,
    has_fill: bool,
    borders: [BorderSide; 4],
    diagonal_border: BorderSide,
    diagonal_up: bool,
    diagonal_down: bool,
    has_border: bool,
    align: TextAlign,
    general_alignment: bool,
    vertical_align: TextVerticalAlign,
    wrap: bool,
    shrink_to_fit: bool,
    indent: u32,
    relative_indent: i32,
    direction: TextDirection,
    rotation_degrees: f32,
    number_format: Option<String>,
}

impl Default for CellStyle {
    fn default() -> Self {
        Self {
            font_family: "Arial".to_owned(),
            font_size: 11.0 * 96.0 / 72.0,
            color: 0x0000_00ff,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            baseline_shift: 0.0,
            fill: 0xffff_ffff,
            fill_paint: None,
            has_fill: false,
            borders: [BorderSide::default(); 4],
            diagonal_border: BorderSide::default(),
            diagonal_up: false,
            diagonal_down: false,
            has_border: false,
            align: TextAlign::Start,
            general_alignment: true,
            vertical_align: TextVerticalAlign::Bottom,
            wrap: false,
            shrink_to_fit: false,
            indent: 0,
            relative_indent: 0,
            direction: TextDirection::Auto,
            rotation_degrees: 0.0,
            number_format: Some("General".to_owned()),
        }
    }
}

#[derive(Clone, Debug)]
struct FontRecord {
    family: String,
    size: f32,
    color: u32,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    baseline_ratio: f32,
}

impl Default for FontRecord {
    fn default() -> Self {
        let style = CellStyle::default();
        Self {
            family: style.font_family,
            size: style.font_size,
            color: style.color,
            bold: style.bold,
            italic: style.italic,
            underline: style.underline,
            strikethrough: style.strikethrough,
            baseline_ratio: 0.0,
        }
    }
}

#[derive(Clone, Debug)]
struct FillRecord {
    color: u32,
    solid: bool,
    paint: Option<ChartFill>,
}

impl Default for FillRecord {
    fn default() -> Self {
        Self {
            color: 0xffff_ffff,
            solid: false,
            paint: None,
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct BorderSide {
    color: u32,
    width: f32,
}

impl Default for BorderSide {
    fn default() -> Self {
        Self {
            color: 0x0000_00ff,
            width: 0.0,
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct BorderRecord {
    sides: [BorderSide; 4],
    diagonal: BorderSide,
    diagonal_up: bool,
    diagonal_down: bool,
    has_border: bool,
}

impl Default for BorderRecord {
    fn default() -> Self {
        Self {
            sides: [BorderSide::default(); 4],
            diagonal: BorderSide::default(),
            diagonal_up: false,
            diagonal_down: false,
            has_border: false,
        }
    }
}

#[derive(Clone, Debug, Default)]
struct RawCellFormat {
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    number_format_id: u32,
    align: Option<TextAlign>,
    vertical_align: Option<TextVerticalAlign>,
    wrap: bool,
    shrink_to_fit: bool,
    indent: u32,
    relative_indent: i32,
    direction: Option<TextDirection>,
    rotation_degrees: f32,
}

#[derive(Clone, Debug)]
struct Styles {
    cells: Vec<CellStyle>,
    normal_font: Option<FontRecord>,
    differentials: Vec<DifferentialStyle>,
    table_styles: HashMap<String, DifferentialStyle>,
    theme_colors: ThemeColors,
    chart_theme_colors: Option<ThemeColors>,
    theme_line_styles: DrawingMlThemeLineStyles,
    diagnostics: Vec<Diagnostic>,
}

impl Default for Styles {
    fn default() -> Self {
        Self {
            cells: vec![CellStyle::default()],
            normal_font: None,
            differentials: Vec::new(),
            table_styles: HashMap::new(),
            theme_colors: default_theme_colors(),
            chart_theme_colors: None,
            theme_line_styles: DrawingMlThemeLineStyles::default(),
            diagnostics: Vec::new(),
        }
    }
}

#[derive(Clone, Debug, Default)]
struct DifferentialStyle {
    color: Option<u32>,
    bold: Option<bool>,
    italic: Option<bool>,
    fill: Option<u32>,
    fill_paint: Option<ChartFill>,
}

#[derive(Clone, Copy, Debug)]
enum CellOperator {
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
    GreaterThanOrEqual,
    GreaterThan,
    Between,
    NotBetween,
}

#[derive(Debug)]
struct ConditionalRule {
    priority: u32,
    stop_if_true: bool,
    kind: ConditionalRuleKind,
}

#[derive(Debug)]
enum ConditionalRuleKind {
    Calculated {
        dxf_id: usize,
        matches: HashSet<(u32, u32)>,
    },
    CellIs {
        dxf_id: usize,
        operator: CellOperator,
        formulas: Vec<ConditionalLiteral>,
    },
    ColorScale {
        thresholds: Vec<ConditionalValue>,
        colors: Vec<u32>,
    },
    DataBar {
        thresholds: Vec<ConditionalValue>,
        color: u32,
        show_value: bool,
    },
    IconSet {
        thresholds: Vec<ConditionalValue>,
        inclusive: Vec<bool>,
        icon_set: String,
        reverse: bool,
        show_value: bool,
    },
}

#[derive(Debug)]
enum ConditionalLiteral {
    Number(f64),
    Text(String),
}

#[derive(Clone, Copy, Debug)]
enum ConditionalValue {
    Minimum,
    Maximum,
    Number(f64),
    Percent(f64),
    Percentile(f64),
}

#[derive(Debug)]
struct ConditionalFormatting {
    ranges: Vec<Dimension>,
    rules: Vec<ConditionalRule>,
}

#[derive(Debug)]
struct ConditionalCell {
    address: CellAddress,
    bounds: Rect,
    value: f64,
}

#[derive(Debug)]
struct ConditionalStats {
    minimum: f64,
    maximum: f64,
    sorted: Vec<f64>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum DataValidationKind {
    List,
    Whole,
    Decimal,
    Date,
    Time,
    TextLength,
    Custom,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum SparklineKind {
    Line,
    Column,
    WinLoss,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum SparklineAxisMode {
    Individual,
    Group,
    Custom,
}

#[derive(Debug)]
struct Sparkline {
    source_formula: String,
    target: CellAddress,
}

#[derive(Debug)]
struct SparklineGroup {
    kind: SparklineKind,
    min_axis: SparklineAxisMode,
    max_axis: SparklineAxisMode,
    manual_min: Option<f64>,
    manual_max: Option<f64>,
    series_color: u32,
    negative_color: u32,
    axis_color: u32,
    display_axis: bool,
    sparklines: Vec<Sparkline>,
}

#[derive(Clone, Copy, Debug, Default)]
struct DrawingMarker {
    column: u32,
    row: u32,
    column_offset: i64,
    row_offset: i64,
}

#[derive(Clone, Copy, Debug)]
enum DrawingAnchorKind {
    TwoCell,
    OneCell,
    Absolute,
}

#[derive(Clone, Copy, Debug)]
enum DrawingMarkerField {
    FromColumn,
    FromColumnOffset,
    FromRow,
    FromRowOffset,
    ToColumn,
    ToColumnOffset,
    ToRow,
    ToRowOffset,
}

#[derive(Debug)]
struct DrawingAnchorState {
    depth: usize,
    kind: DrawingAnchorKind,
    from: DrawingMarker,
    to: DrawingMarker,
    marker_depth: Option<(usize, bool)>,
    field: Option<(usize, DrawingMarkerField, String)>,
    x: Option<i64>,
    y: Option<i64>,
    width: Option<i64>,
    height: Option<i64>,
    drawing_id: Option<u32>,
    image_relationship_id: Option<String>,
    image_crop: ImageCrop,
    image_adjustment: ImageAdjustment,
    image_geometry: super::drawingml::DrawingMlPictureGeometry,
    image_effects: DrawingMlPictureEffectsCapture,
    image_rotation_degrees: f32,
    image_flip_h: bool,
    image_flip_v: bool,
    image_blip_depth: Option<usize>,
    duotone_depth: Option<usize>,
    duotone_colors: Vec<u32>,
    duotone_color: Option<(usize, u32)>,
    chart_relationship_id: Option<String>,
    diagram_relationship_id: Option<String>,
}

#[derive(Clone, Copy, Debug, Default)]
struct DrawingTransform {
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    child_x: i64,
    child_y: i64,
    child_width: i64,
    child_height: i64,
    rotation_degrees: f32,
    flip_h: bool,
    flip_v: bool,
}

#[derive(Debug)]
struct DrawingGroupState {
    depth: usize,
    properties_depth: Option<usize>,
    transform_depth: Option<usize>,
    transform: DrawingTransform,
    image_fill: Option<WorksheetShapeImageFill>,
}

#[derive(Clone, Copy, Debug)]
enum DrawingColorTarget {
    Fill,
    Stroke,
    TextStroke,
    Text,
    Shadow,
}

#[derive(Debug)]
struct DrawingColorState {
    depth: usize,
    target: DrawingColorTarget,
    color: u32,
    luminance_modulation: f32,
    luminance_offset: f32,
    alpha: f32,
    explicit: bool,
}

#[derive(Debug)]
struct WorksheetShapeState {
    depth: usize,
    is_connector: bool,
    drawing_id: Option<u32>,
    transform_depth: Option<usize>,
    transform: DrawingTransform,
    properties_depth: Option<usize>,
    line_depth: Option<usize>,
    effect_depth: Option<usize>,
    shadow_depth: Option<usize>,
    style_line_depth: Option<usize>,
    style_fill_depth: Option<usize>,
    hidden_style_depth: Option<usize>,
    text_body_depth: Option<usize>,
    text_line_depth: Option<usize>,
    scene_3d_depth: Option<usize>,
    camera_3d_depth: Option<usize>,
    light_3d_depth: Option<usize>,
    text_depth: Option<usize>,
    paragraph_depth: Option<usize>,
    preset_geometry: Option<String>,
    geometry: Geometry,
    fill: Paint,
    stroke: Paint,
    stroke_width: f32,
    stroke_width_explicit: bool,
    stroke_explicit: bool,
    shape_stroke_authored: bool,
    fill_explicit: bool,
    image_fill: Option<WorksheetShapeImageFill>,
    dash_pattern: Vec<f32>,
    head_arrow: bool,
    tail_arrow: bool,
    connector_preset: Option<String>,
    shadow: Option<Shadow>,
    outer_shadow: Option<OuterShadow>,
    three_d: Option<ThreeDStyle>,
    text: String,
    font_family: String,
    font_size: f32,
    font_color: u32,
    text_stroke_color: u32,
    text_stroke_fill_capture: Option<ChartFillCapture>,
    text_stroke_fill: Option<ChartFill>,
    text_stroke_width: f32,
    bold: bool,
    italic: bool,
    align: TextAlign,
    vertical_align: TextVerticalAlign,
    wrap: bool,
}

impl WorksheetShapeState {
    fn new(depth: usize, is_connector: bool, theme_colors: &ThemeColors) -> Self {
        Self {
            depth,
            is_connector,
            drawing_id: None,
            transform_depth: None,
            transform: DrawingTransform::default(),
            properties_depth: None,
            line_depth: None,
            effect_depth: None,
            shadow_depth: None,
            style_line_depth: None,
            style_fill_depth: None,
            hidden_style_depth: None,
            text_body_depth: None,
            text_line_depth: None,
            scene_3d_depth: None,
            camera_3d_depth: None,
            light_3d_depth: None,
            text_depth: None,
            paragraph_depth: None,
            preset_geometry: None,
            geometry: if is_connector {
                Geometry::Line
            } else {
                Geometry::Rectangle
            },
            fill: if is_connector {
                Paint::None
            } else {
                Paint::Solid(theme_colors[4])
            },
            stroke: Paint::Solid(theme_colors[4]),
            stroke_width: 1.0,
            stroke_width_explicit: false,
            stroke_explicit: false,
            shape_stroke_authored: false,
            fill_explicit: false,
            image_fill: None,
            dash_pattern: Vec::new(),
            head_arrow: false,
            tail_arrow: false,
            connector_preset: None,
            shadow: None,
            outer_shadow: None,
            three_d: None,
            text: String::new(),
            font_family: "Arial".to_owned(),
            font_size: 11.0 * 96.0 / 72.0,
            font_color: 0x0000_00ff,
            text_stroke_color: 0,
            text_stroke_fill_capture: None,
            text_stroke_fill: None,
            text_stroke_width: 0.0,
            bold: false,
            italic: false,
            align: TextAlign::Center,
            vertical_align: TextVerticalAlign::Center,
            wrap: true,
        }
    }
}

#[derive(Debug)]
struct WorksheetShape {
    geometry: Geometry,
    fill: Paint,
    image_fill: Option<WorksheetShapeImageFill>,
    stroke: Paint,
    stroke_width: f32,
    dash: Vec<f32>,
    head_arrow: bool,
    tail_arrow: bool,
    connector_preset: Option<String>,
    flip_h: bool,
    flip_v: bool,
    rotation_degrees: f32,
    shadow: Option<Shadow>,
    outer_shadow: Option<OuterShadow>,
    three_d: Option<ThreeDStyle>,
    text: String,
    font_family: String,
    font_size: f32,
    font_color: u32,
    text_stroke_color: u32,
    text_stroke_fill: Option<ChartFill>,
    text_stroke_width: f32,
    bold: bool,
    italic: bool,
    align: TextAlign,
    vertical_align: TextVerticalAlign,
    wrap: bool,
}

#[derive(Clone, Debug)]
struct WorksheetShapeImageFill {
    mapping: crate::model::ImageFillMapping,
    parse_depth: Option<usize>,
    relationship_id: Option<String>,
    crop: ImageCrop,
    tile: bool,
}

impl Default for WorksheetShapeImageFill {
    fn default() -> Self {
        Self {
            mapping: crate::model::ImageFillMapping::default(),
            parse_depth: None,
            relationship_id: None,
            crop: ImageCrop::default(),
            tile: false,
        }
    }
}

#[derive(Debug)]
enum DrawingPayload {
    Image(Box<WorksheetImage>),
    Chart(String),
    Diagram(String),
    Shape(Box<WorksheetShape>),
}

#[derive(Debug)]
struct WorksheetImage {
    relationship_id: String,
    crop: ImageCrop,
    adjustment: ImageAdjustment,
    geometry: Option<Geometry>,
    effects: DrawingMlPictureEffects,
    rotation_degrees: f32,
    flip_h: bool,
    flip_v: bool,
}

#[derive(Debug)]
struct WorksheetDrawing {
    bounds: Rect,
    drawing_id: u32,
    payload: DrawingPayload,
}

#[derive(Debug)]
struct WorksheetVmlPreview {
    part: String,
    relationship: Relationship,
}

impl AxisMetrics {
    fn new(default_size: f32, mut overrides: Vec<(u32, f32)>) -> Self {
        overrides.sort_unstable_by_key(|(index, _)| *index);
        let mut prefix_delta = Vec::with_capacity(overrides.len() + 1);
        prefix_delta.push(0.0);
        for (_, size) in &overrides {
            let next = prefix_delta.last().copied().unwrap_or(0.0) + *size - default_size;
            prefix_delta.push(next);
        }
        Self {
            default_size,
            overrides,
            prefix_delta,
        }
    }

    fn offset(&self, index: u32) -> f32 {
        let override_count = self
            .overrides
            .partition_point(|(override_index, _)| *override_index < index);
        index as f32 * self.default_size + self.prefix_delta[override_count]
    }

    fn size(&self, index: u32) -> f32 {
        self.overrides
            .binary_search_by_key(&index, |(override_index, _)| *override_index)
            .ok()
            .map(|position| self.overrides[position].1)
            .unwrap_or(self.default_size)
    }

    fn span(&self, start: u32, end: u32) -> f32 {
        self.offset(end.saturating_add(1)) - self.offset(start)
    }

    fn count_for_extent(&self, extent: f32, limit: u32) -> u32 {
        crate::model::sheet_axis_count_for_extent(extent, limit, |index| self.offset(index))
    }

    fn descriptor(&self, count: u32) -> SheetAxis {
        let mut spans: Vec<SheetAxisSpan> = Vec::new();
        for &(index, size) in self.overrides.iter().filter(|(index, _)| *index < count) {
            if let Some(last) = spans.last_mut()
                && last.end.saturating_add(1) == index
                && last.size.to_bits() == size.to_bits()
            {
                last.end = index;
            } else {
                spans.push(SheetAxisSpan {
                    start: index,
                    end: index,
                    size,
                });
            }
        }
        SheetAxis {
            default_size: self.default_size,
            spans,
        }
    }
}

#[derive(Debug)]
struct XlsxParseState {
    objects: Vec<Object>,
    diagnostics: Vec<Diagnostic>,
    materialized_text_bytes: usize,
    materialized_text_limit: usize,
    materialized_image_bytes: usize,
    materialized_image_limit: usize,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum XlsxParseMode {
    Complete,
    Metadata,
    Unit(u32),
    UnitRegion {
        unit_index: u32,
        max_row: i32,
        max_column: i32,
    },
}

impl XlsxParseMode {
    const fn unit_index(self) -> Option<u32> {
        match self {
            Self::Complete | Self::Metadata => None,
            Self::Unit(unit_index) | Self::UnitRegion { unit_index, .. } => Some(unit_index),
        }
    }
}

#[derive(Clone, Debug)]
struct WorksheetMetadata {
    width: f32,
    height: f32,
    rows: u32,
    columns: u32,
    frozen_rows: u32,
    frozen_columns: u32,
    frozen_width: f32,
    frozen_height: f32,
    row_axis: SheetAxis,
    column_axis: SheetAxis,
    show_grid_lines: bool,
    tab_color: Option<u32>,
    diagnostics: Vec<Diagnostic>,
    print_settings: Option<SheetPrintSettings>,
    requires_formula_evaluation: bool,
}

impl WorksheetMetadata {
    fn unit(&self, index: u32, name: String) -> Unit {
        Unit {
            kind: UnitKind::Sheet,
            index,
            id: format!("unit:{index}"),
            name,
            width: self.width,
            height: self.height,
            rows: self.rows,
            columns: self.columns,
            frozen_rows: self.frozen_rows,
            frozen_columns: self.frozen_columns,
            frozen_width: self.frozen_width,
            frozen_height: self.frozen_height,
            row_axis: self.row_axis.clone(),
            column_axis: self.column_axis.clone(),
            show_grid_lines: self.show_grid_lines,
            tab_color: self.tab_color,
            sheet: self.print_settings.clone(),
            slide: None,
        }
    }
}

pub(super) fn parse(
    package: &Package<'_>,
    workbook_part: &str,
    content_types: &ContentTypes,
    calculation: Option<&CalculationResults>,
) -> Result<Document, Diagnostic> {
    parse_with_mode(
        package,
        workbook_part,
        content_types,
        XlsxParseMode::Complete,
        calculation,
    )
}

#[cfg(test)]
pub(super) fn parse_preview(
    package: &Package<'_>,
    workbook_part: &str,
    content_types: &ContentTypes,
) -> Result<Document, Diagnostic> {
    parse_with_mode(
        package,
        workbook_part,
        content_types,
        XlsxParseMode::Unit(0),
        None,
    )
}

pub(super) fn parse_metadata(
    package: &Package<'_>,
    workbook_part: &str,
    content_types: &ContentTypes,
) -> Result<Document, Diagnostic> {
    parse_with_mode(
        package,
        workbook_part,
        content_types,
        XlsxParseMode::Metadata,
        None,
    )
}

pub(super) fn parse_unit(
    package: &Package<'_>,
    workbook_part: &str,
    content_types: &ContentTypes,
    unit_index: u32,
    calculation: Option<&CalculationResults>,
) -> Result<Document, Diagnostic> {
    parse_with_mode(
        package,
        workbook_part,
        content_types,
        XlsxParseMode::Unit(unit_index),
        calculation,
    )
}

pub(super) fn parse_unit_region(
    package: &Package<'_>,
    workbook_part: &str,
    content_types: &ContentTypes,
    unit_index: u32,
    max_row: i32,
    max_column: i32,
    calculation: Option<&CalculationResults>,
) -> Result<Document, Diagnostic> {
    parse_with_mode(
        package,
        workbook_part,
        content_types,
        XlsxParseMode::UnitRegion {
            unit_index,
            max_row,
            max_column,
        },
        calculation,
    )
}

fn parse_with_mode(
    package: &Package<'_>,
    workbook_part: &str,
    content_types: &ContentTypes,
    mode: XlsxParseMode,
    calculation: Option<&CalculationResults>,
) -> Result<Document, Diagnostic> {
    let (sheets, date_1904, ole_size) = parse_workbook(package, workbook_part)?;
    if sheets.is_empty() {
        return Err(format_error(workbook_part, "workbook contains no sheets"));
    }

    let relationships = package.relationships(Some(workbook_part))?;
    let shared_relationships: Vec<_> = relationships
        .iter()
        .filter(|relationship| relationship.type_uri.ends_with("/sharedStrings"))
        .collect();
    if shared_relationships.len() > 1 {
        return Err(format_error(
            workbook_part,
            "workbook contains multiple shared-string relationships",
        ));
    }
    let theme_relationships: Vec<_> = relationships
        .iter()
        .filter(|relationship| relationship.type_uri.ends_with("/theme"))
        .collect();
    if theme_relationships.len() > 1 {
        return Err(format_error(
            workbook_part,
            "workbook contains multiple theme relationships",
        ));
    }
    let (theme_colors, theme_line_styles) = match theme_relationships.first() {
        Some(relationship) => {
            ensure_internal(relationship, workbook_part, "theme")?;
            parse_theme(package, &relationship.target)?
        }
        None => (default_theme_colors(), DrawingMlThemeLineStyles::default()),
    };
    let shared_strings = match shared_relationships.first() {
        Some(relationship) => {
            ensure_internal(relationship, workbook_part, "shared-string")?;
            parse_shared_strings(package, &relationship.target, &theme_colors)?
        }
        None => Vec::new(),
    };
    let style_relationships: Vec<_> = relationships
        .iter()
        .filter(|relationship| relationship.type_uri.ends_with("/styles"))
        .collect();
    if style_relationships.len() > 1 {
        return Err(format_error(
            workbook_part,
            "workbook contains multiple style relationships",
        ));
    }
    let mut styles = match style_relationships.first() {
        Some(relationship) => {
            ensure_internal(relationship, workbook_part, "style")?;
            parse_styles(package, &relationship.target, theme_colors)?
        }
        None => Styles {
            theme_colors,
            ..Styles::default()
        },
    };
    styles.theme_line_styles = theme_line_styles;
    styles.chart_theme_colors = theme_relationships.first().map(|_| styles.theme_colors);
    let maximum_digit_width = excel_maximum_digit_width(&styles);

    let relationship_map: HashMap<&str, &Relationship> = relationships
        .iter()
        .map(|relationship| (relationship.id.as_str(), relationship))
        .collect();
    let chart_named_data = if matches!(mode, XlsxParseMode::Metadata) {
        HashMap::new()
    } else {
        parse_chart_named_data(
            package,
            workbook_part,
            &sheets,
            &relationship_map,
            &shared_strings,
        )?
    };
    let mut units = Vec::with_capacity(sheets.iter().filter(|sheet| !sheet.hidden).count());
    let mut state = XlsxParseState {
        objects: Vec::new(),
        diagnostics: styles.diagnostics.clone(),
        materialized_text_bytes: 0,
        materialized_text_limit: package.limits().max_total_uncompressed_bytes,
        materialized_image_bytes: 0,
        materialized_image_limit: package.limits().max_total_uncompressed_bytes,
    };
    let unit_metadata = if !matches!(mode, XlsxParseMode::Complete) {
        sheets
            .iter()
            .map(|sheet| {
                let relationship = relationship_map
                    .get(sheet.relationship_id.as_str())
                    .ok_or_else(|| {
                        format_error(
                            workbook_part,
                            format!(
                                "sheet {} references missing relationship {}",
                                sheet.name, sheet.relationship_id
                            ),
                        )
                    })?;
                if relationship.external || !relationship.type_uri.ends_with("/worksheet") {
                    return Ok(None);
                }
                parse_worksheet_metadata(
                    package,
                    &relationship.target,
                    sheet.print_area.as_deref(),
                    &styles.theme_colors,
                    maximum_digit_width,
                    &shared_strings,
                    &styles,
                )
                .map(|mut metadata| {
                    if let Some(titles) = &sheet.print_titles {
                        metadata
                            .print_settings
                            .get_or_insert_with(SheetPrintSettings::default)
                            .print_titles = Some(titles.clone());
                    }
                    Some(metadata)
                })
            })
            .collect::<Result<Vec<_>, Diagnostic>>()?
    } else {
        vec![None; sheets.len()]
    };
    let requires_formula_evaluation = match mode {
        XlsxParseMode::Metadata | XlsxParseMode::Unit(_) | XlsxParseMode::UnitRegion { .. } => {
            unit_metadata
                .iter()
                .flatten()
                .any(|metadata| metadata.requires_formula_evaluation)
        }
        XlsxParseMode::Complete => sheets.iter().try_fold(false, |required, sheet| {
            if required {
                return Ok(true);
            }
            let Some(relationship) = relationship_map.get(sheet.relationship_id.as_str()) else {
                return Ok(false);
            };
            if !relationship.type_uri.ends_with("/worksheet") || relationship.external {
                return Ok(false);
            }
            worksheet_requires_formula_evaluation(package, &relationship.target)
        })?,
    };
    let formula_evaluator = requires_formula_evaluation.then_some(calculation).flatten();
    if requires_formula_evaluation && formula_evaluator.is_none() {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Unsupported,
                crate::calculation::REQUIRED_MESSAGE,
            )
            .in_part(workbook_part),
        );
    }
    let mut parsed_worksheet_parts = HashSet::new();
    for (source_sheet_index, sheet) in sheets.into_iter().enumerate() {
        let relationship = relationship_map
            .get(sheet.relationship_id.as_str())
            .ok_or_else(|| {
                format_error(
                    workbook_part,
                    format!(
                        "sheet {} references missing relationship {}",
                        sheet.name, sheet.relationship_id
                    ),
                )
            })?;
        ensure_internal(relationship, workbook_part, "worksheet")?;
        if relationship.type_uri.ends_with("/chartsheet") {
            if sheet.hidden {
                continue;
            }
            let unit_index = u32::try_from(units.len())
                .map_err(|_| format_error(workbook_part, "sheet count exceeds supported range"))?;
            let (drawing_relationship_ids, tab_color) = chartsheet_drawing_relationship_ids(
                package,
                &relationship.target,
                &styles.theme_colors,
                &mut state.diagnostics,
            )?;
            let column_metrics = AxisMetrics::new(COLUMN_WIDTH, Vec::new());
            let row_metrics = AxisMetrics::new(ROW_HEIGHT, Vec::new());
            let extent = if matches!(mode, XlsxParseMode::Metadata)
                || mode
                    .unit_index()
                    .is_some_and(|target_unit| unit_index != target_unit)
            {
                worksheet_drawing_extent(
                    package,
                    &relationship.target,
                    &drawing_relationship_ids,
                    &column_metrics,
                    &row_metrics,
                    &styles.theme_colors,
                    &styles.theme_line_styles,
                )?
            } else {
                push_worksheet_drawings(
                    package,
                    &relationship.target,
                    &drawing_relationship_ids,
                    &[],
                    &column_metrics,
                    &row_metrics,
                    &sheet.name,
                    unit_index,
                    content_types,
                    &styles.theme_colors,
                    &styles.theme_line_styles,
                    styles.chart_theme_colors.as_ref(),
                    &chart_named_data,
                    &mut state,
                )?
            };
            let columns = column_metrics
                .count_for_extent(extent.0, MAX_COLUMNS)
                .max(1);
            let rows = row_metrics.count_for_extent(extent.1, MAX_ROWS).max(1);
            units.push(Unit {
                kind: UnitKind::Sheet,
                index: unit_index,
                id: format!("unit:{unit_index}"),
                name: sheet.name,
                width: extent.0.max(COLUMN_WIDTH),
                height: extent.1.max(ROW_HEIGHT),
                rows,
                columns,
                frozen_rows: 0,
                frozen_columns: 0,
                frozen_width: 0.0,
                frozen_height: 0.0,
                row_axis: row_metrics.descriptor(rows),
                column_axis: column_metrics.descriptor(columns),
                show_grid_lines: false,
                tab_color,
                sheet: None,
                slide: None,
            });
            continue;
        }
        if relationship.type_uri.ends_with("/dialogsheet") {
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Parse,
                    Fidelity::Unsupported,
                    format!(
                        "Dialog sheet {:?} was omitted; legacy dialog controls are unsupported",
                        sheet.name
                    ),
                )
                .in_part(&relationship.target),
            );
            continue;
        }
        if !relationship.type_uri.ends_with("/worksheet") {
            return Err(format_error(
                workbook_part,
                format!("relationship {} is not a worksheet", sheet.relationship_id),
            ));
        }
        if !parsed_worksheet_parts.insert(relationship.target.as_str()) {
            return Err(format_error(
                workbook_part,
                format!(
                    "multiple sheets reference the same worksheet part {}",
                    relationship.target
                ),
            ));
        }
        if sheet.hidden {
            continue;
        }
        let unit_index = u32::try_from(units.len())
            .map_err(|_| format_error(workbook_part, "sheet count exceeds supported range"))?;
        let source_sheet_index = u32::try_from(source_sheet_index)
            .map_err(|_| format_error(workbook_part, "sheet count exceeds supported range"))?;
        if matches!(mode, XlsxParseMode::Metadata)
            || mode
                .unit_index()
                .is_some_and(|target_unit| unit_index != target_unit)
        {
            let metadata = unit_metadata[source_sheet_index as usize]
                .as_ref()
                .ok_or_else(|| {
                    format_error(&relationship.target, "worksheet metadata is missing")
                })?;
            state
                .diagnostics
                .extend(metadata.diagnostics.iter().cloned());
            let mut unit = metadata.unit(unit_index, sheet.name);
            if source_sheet_index == 0 {
                apply_ole_size(&mut unit, ole_size);
            }
            units.push(unit);
            continue;
        }
        let mut unit = parse_worksheet(
            package,
            &relationship.target,
            &sheet.name,
            sheet.print_area.as_deref(),
            source_sheet_index,
            unit_index,
            &shared_strings,
            &styles,
            date_1904,
            formula_evaluator,
            content_types,
            &chart_named_data,
            &mut state,
        )?;
        if let Some(titles) = sheet.print_titles {
            unit.sheet
                .get_or_insert_with(SheetPrintSettings::default)
                .print_titles = Some(titles);
        }
        if source_sheet_index == 0 {
            apply_ole_size(&mut unit, ole_size);
        }
        units.push(unit);
    }

    Ok(Document {
        fatal: false,
        format: Some(DocumentFormat::Xlsx),
        kind: Some(DocumentKind::Spreadsheet),
        units,
        outline: Vec::new(),
        objects: state.objects,
        embedded_fonts: Vec::new(),
        font_alternate_names: Vec::new(),
        diagnostics: state.diagnostics,
    })
}

struct FastSheetDataMetadata {
    structural_xml: Vec<u8>,
    dimension: Option<Dimension>,
    row_sizes: Vec<(u32, f32)>,
    trailing_text: Vec<FastTrailingText>,
    requires_formula_evaluation: bool,
}

#[derive(Clone, Copy)]
struct FastTrailingText {
    column: u32,
    row: u32,
    align: TextAlign,
    width: f32,
}

fn fast_sheet_data_metadata(
    bytes: &[u8],
    shared_strings: &[SharedString],
    styles: &Styles,
) -> Option<FastSheetDataMetadata> {
    if !contains_bytes(
        bytes,
        b"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"",
    ) && !contains_bytes(
        bytes,
        b"<worksheet xmlns=\"http://purl.oclc.org/ooxml/spreadsheetml/main\"",
    ) {
        return None;
    }
    let sheet_data_start = find_bytes(bytes, b"<sheetData", 0)?;
    let delimiter = *bytes.get(sheet_data_start + 10)?;
    if !delimiter.is_ascii_whitespace() && delimiter != b'>' {
        return None;
    }
    let sheet_data_open_end = bytes
        .get(sheet_data_start..)?
        .iter()
        .position(|byte| *byte == b'>')?
        + sheet_data_start;
    let sheet_data_close = find_bytes(bytes, b"</sheetData>", sheet_data_open_end + 1)?;
    let mut structural_xml =
        Vec::with_capacity(sheet_data_start + bytes.len().saturating_sub(sheet_data_close + 12));
    structural_xml.extend_from_slice(bytes.get(..sheet_data_start)?);
    structural_xml.extend_from_slice(bytes.get(sheet_data_close + 12..)?);

    let sheet_data = bytes.get(sheet_data_open_end + 1..sheet_data_close)?;
    let mut position = 0_usize;
    let mut dimension = None::<Dimension>;
    let mut row_sizes = Vec::new();
    let mut rightmost_cell_column = None;
    let mut trailing_text = Vec::new();
    let mut requires_formula_evaluation = false;
    while let Some(start) = find_bytes(sheet_data, b"<row", position) {
        let delimiter = *sheet_data.get(start + 4)?;
        if !delimiter.is_ascii_whitespace() && delimiter != b'>' && delimiter != b'/' {
            position = start + 4;
            continue;
        }
        let end = sheet_data
            .get(start..)?
            .iter()
            .position(|byte| *byte == b'>')?
            + start;
        let opening = sheet_data.get(start..=end)?;
        if let Some(index) = fast_attribute(opening, b"r") {
            let index = fast_positive_u32(index, MAX_ROWS)?.saturating_sub(1);
            let hidden = match fast_attribute(opening, b"hidden") {
                None | Some(b"0" | b"false") => false,
                Some(b"1" | b"true") => true,
                Some(_) => return None,
            };
            let height = if hidden {
                Some(0.0)
            } else if let Some(height) = fast_attribute(opening, b"ht") {
                let height = std::str::from_utf8(height).ok()?.parse::<f32>().ok()?;
                Some((height.is_finite() && height <= 409.5).then_some(height * 96.0 / 72.0)?)
            } else {
                None
            };
            if let Some(height) = height {
                row_sizes.push((index, height));
            }
        }
        position = end + 1;
    }
    position = 0;
    while let Some(start) = find_bytes(sheet_data, b"<c", position) {
        let delimiter = *sheet_data.get(start + 2)?;
        if !delimiter.is_ascii_whitespace() && delimiter != b'>' && delimiter != b'/' {
            position = start + 2;
            continue;
        }
        let open_end = sheet_data
            .get(start..)?
            .iter()
            .position(|byte| *byte == b'>')?
            + start;
        let opening = sheet_data.get(start..=open_end)?;
        let address = fast_attribute(opening, b"r").and_then(fast_a1);
        if let Some((column, row)) = address {
            let candidate = Dimension {
                start_column: column,
                start_row: row,
                end_column: column,
                end_row: row,
            };
            dimension = Some(match dimension {
                Some(current) => Dimension {
                    start_column: current.start_column.min(candidate.start_column),
                    start_row: current.start_row.min(candidate.start_row),
                    end_column: current.end_column.max(candidate.end_column),
                    end_row: current.end_row.max(candidate.end_row),
                },
                None => candidate,
            });
            if rightmost_cell_column.is_none_or(|rightmost| column > rightmost) {
                rightmost_cell_column = Some(column);
                trailing_text.clear();
            }
        }
        if opening
            .get(..opening.len().saturating_sub(1))?
            .iter()
            .rev()
            .find(|byte| !byte.is_ascii_whitespace())
            .is_none_or(|byte| *byte != b'/')
        {
            let close = find_bytes(sheet_data, b"</c>", open_end + 1)?;
            let content = sheet_data.get(open_end + 1..close)?;
            if contains_element_start(content, b'f') {
                requires_formula_evaluation |= !contains_element_start(content, b'v')
                    || fast_element_text(content, b"f")
                        .is_some_and(|formula| formula_calls_any(&formula, &["NOW", "TODAY"]));
            }
            if let Some((column, row)) = address
                && rightmost_cell_column == Some(column)
            {
                let style_index = match fast_attribute(opening, b"s") {
                    Some(value) => std::str::from_utf8(value).ok()?.parse::<usize>().ok()?,
                    None => 0,
                };
                let style = styles
                    .cells
                    .get(style_index)
                    .unwrap_or_else(|| &styles.cells[0]);
                if !style.wrap
                    && !style.shrink_to_fit
                    && style.rotation_degrees == 0.0
                    && matches!(style.align, TextAlign::Start | TextAlign::Center)
                    && let Some(text) = fast_cell_text(opening, content, shared_strings)
                {
                    trailing_text.push(FastTrailingText {
                        column,
                        row,
                        align: style.align,
                        width: text
                            .chars()
                            .map(|character| {
                                drawingml_fallback_character_width(character, style.font_size)
                            })
                            .sum(),
                    });
                }
            }
            position = close + 4;
        } else {
            position = open_end + 1;
        }
    }
    Some(FastSheetDataMetadata {
        structural_xml,
        dimension,
        row_sizes,
        trailing_text,
        requires_formula_evaluation,
    })
}

fn fast_cell_text(
    opening: &[u8],
    content: &[u8],
    shared_strings: &[SharedString],
) -> Option<String> {
    match fast_attribute(opening, b"t") {
        Some(b"s") => fast_element_text(content, b"v")?
            .trim()
            .parse::<usize>()
            .ok()
            .and_then(|index| shared_strings.get(index).map(|value| value.text.clone())),
        Some(b"inlineStr") => fast_element_text(content, b"t"),
        Some(b"str") => fast_element_text(content, b"v"),
        _ => None,
    }
}

fn fast_element_text(bytes: &[u8], name: &[u8]) -> Option<String> {
    let mut opening = Vec::with_capacity(name.len() + 1);
    opening.push(b'<');
    opening.extend_from_slice(name);
    let mut closing = Vec::with_capacity(name.len() + 3);
    closing.extend_from_slice(b"</");
    closing.extend_from_slice(name);
    closing.push(b'>');
    let mut position = 0;
    let mut value = String::new();
    while let Some(start) = find_bytes(bytes, &opening, position) {
        let delimiter = *bytes.get(start + opening.len())?;
        if !delimiter.is_ascii_whitespace() && delimiter != b'>' {
            position = start + opening.len();
            continue;
        }
        let open_end = bytes.get(start..)?.iter().position(|byte| *byte == b'>')? + start;
        let close = find_bytes(bytes, &closing, open_end + 1)?;
        value.push_str(
            &decode_xml_text(std::str::from_utf8(bytes.get(open_end + 1..close)?).ok()?).ok()?,
        );
        position = close + closing.len();
    }
    (!value.is_empty()).then_some(value)
}

fn fast_attribute<'a>(tag: &'a [u8], name: &[u8]) -> Option<&'a [u8]> {
    let mut position = 1_usize;
    while position < tag.len() {
        while tag.get(position).is_some_and(u8::is_ascii_whitespace) {
            position += 1;
        }
        let start = position;
        while tag.get(position).is_some_and(|byte| {
            byte.is_ascii_alphanumeric() || matches!(byte, b':' | b'_' | b'-' | b'.')
        }) {
            position += 1;
        }
        if start == position {
            position += 1;
            continue;
        }
        let attribute_name = tag.get(start..position)?;
        while tag.get(position).is_some_and(u8::is_ascii_whitespace) {
            position += 1;
        }
        if tag.get(position) != Some(&b'=') {
            continue;
        }
        position += 1;
        while tag.get(position).is_some_and(u8::is_ascii_whitespace) {
            position += 1;
        }
        let quote = *tag.get(position)?;
        if !matches!(quote, b'\'' | b'"') {
            return None;
        }
        position += 1;
        let value_start = position;
        while tag.get(position).is_some_and(|byte| *byte != quote) {
            position += 1;
        }
        let value = tag.get(value_start..position)?;
        position += 1;
        if attribute_name == name {
            return Some(value);
        }
    }
    None
}

fn fast_a1(reference: &[u8]) -> Option<(u32, u32)> {
    let reference = reference.strip_prefix(b"$").unwrap_or(reference);
    let letters = reference
        .iter()
        .take_while(|byte| byte.is_ascii_uppercase())
        .count();
    if letters == 0 {
        return None;
    }
    let mut column = 0_u32;
    for byte in &reference[..letters] {
        column = column
            .checked_mul(26)?
            .checked_add(u32::from(*byte - b'A') + 1)?;
    }
    if column > MAX_COLUMNS {
        return None;
    }
    let digits = reference
        .get(letters..)?
        .strip_prefix(b"$")
        .unwrap_or(&reference[letters..]);
    if digits.is_empty() || digits[0] == b'0' || !digits.iter().all(u8::is_ascii_digit) {
        return None;
    }
    let row = std::str::from_utf8(digits).ok()?.parse::<u32>().ok()?;
    (row <= MAX_ROWS).then_some((column - 1, row - 1))
}

fn fast_positive_u32(value: &[u8], maximum: u32) -> Option<u32> {
    if value.is_empty() || value[0] == b'0' || !value.iter().all(u8::is_ascii_digit) {
        return None;
    }
    let value = std::str::from_utf8(value).ok()?.parse::<u32>().ok()?;
    (value <= maximum).then_some(value)
}

fn contains_element_start(bytes: &[u8], name: u8) -> bool {
    bytes.windows(3).any(|window| {
        window[0] == b'<'
            && window[1] == name
            && (window[2].is_ascii_whitespace() || matches!(window[2], b'>' | b'/'))
    })
}

fn worksheet_requires_formula_evaluation(
    package: &Package<'_>,
    part: &str,
) -> Result<bool, Diagnostic> {
    #[derive(Clone, Copy)]
    struct CellScan {
        depth: usize,
        has_formula: bool,
        value: bool,
    }

    let bytes = package.required_part(part)?;
    let mut depth = 0_usize;
    let mut cell = None::<CellScan>;
    let mut formula_depth = None;
    let mut formula = String::new();
    let mut required = false;
    let mut conditional_formula_depth = None;
    let mut conditional_formula = String::new();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "cfRule" && depth == 2 {
                    let kind = optional_attribute(&attributes, "type", part)?.unwrap_or_default();
                    required |= crate::calculation::calculated_condition(&kind) && kind != "cellIs";
                }
                if local == "formula" && cell.is_none() {
                    conditional_formula.clear();
                    conditional_formula_depth = (!empty).then_some(depth);
                }
                if local == "c" && cell.is_none() {
                    cell = Some(CellScan {
                        depth,
                        has_formula: false,
                        value: false,
                    });
                } else if let Some(cell) = cell.as_mut() {
                    if local == "f" {
                        cell.has_formula = true;
                        formula.clear();
                        formula_depth = (!empty).then_some(depth);
                    } else if local == "v" {
                        cell.value = true;
                    }
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if local == "formula" && conditional_formula_depth == Some(depth) {
                    conditional_formula_depth = None;
                    required |= parse_conditional_literal(&conditional_formula).is_none();
                }
                if local == "f" && formula_depth == Some(depth) {
                    formula_depth = None;
                } else if local == "c" && cell.is_some_and(|cell| cell.depth == depth) {
                    let completed = cell.take().expect("cell scan is present");
                    required |= completed.has_formula
                        && (!completed.value || formula_calls_any(&formula, &["NOW", "TODAY"]));
                }
            }
            XmlEvent::Text(text) if conditional_formula_depth.is_some() => {
                conditional_formula
                    .push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
            }
            XmlEvent::Cdata(text) if conditional_formula_depth.is_some() => {
                conditional_formula.push_str(text)
            }
            XmlEvent::Text(text) if formula_depth.is_some() => {
                formula.push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
            }
            XmlEvent::Cdata(text) if formula_depth.is_some() => formula.push_str(text),
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(required)
}

fn parse_workbook(
    package: &Package<'_>,
    part: &str,
) -> Result<(Vec<Sheet>, bool, Option<Dimension>), Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut sheets = Vec::new();
    let mut names = HashSet::new();
    let mut date_1904 = false;
    let mut ole_size = None;
    parse_xml(&bytes, package.limits(), |event| {
        let XmlEvent::StartElement {
            name, attributes, ..
        } = event
        else {
            return Ok(());
        };
        let local = local_name(name);
        if local == "workbookPr" {
            date_1904 = parse_bool_attribute(&attributes, "date1904", part)?.unwrap_or(false);
            return Ok(());
        }
        if local == "oleSize" {
            ole_size = optional_attribute(&attributes, "ref", part)?
                .map(|value| parse_dimension(&value, part))
                .transpose()?;
            return Ok(());
        }
        if local != "sheet" {
            return Ok(());
        }
        if sheets.len() >= package.limits().max_relationship_edges {
            return Err(Diagnostic::fatal(
                DiagnosticCode::RelationshipLimit,
                Phase::Parse,
                None,
                "workbook exceeds the configured sheet limit",
            )
            .in_part(part));
        }
        let name = required_attribute(&attributes, "name", part)?;
        if name.is_empty() || !names.insert(name.clone()) {
            return Err(format_error(
                part,
                "sheet names must be non-empty and unique",
            ));
        }
        let relationship_id = relationship_id(&attributes, part)?;
        let hidden = match optional_attribute(&attributes, "state", part)?.as_deref() {
            None | Some("visible") => false,
            Some("hidden" | "veryHidden") => true,
            Some(_) => return Err(format_error(part, "sheet has an invalid visibility state")),
        };
        sheets.push(Sheet {
            name,
            relationship_id,
            hidden,
            print_area: None,
            print_titles: None,
        });
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    for (sheet_index, print_area) in
        parse_workbook_print_areas(&bytes, package.limits(), part, "_xlnm.Print_Area")?
    {
        let Some(sheet) = sheets.get_mut(sheet_index) else {
            return Err(format_error(part, "print area refers to a missing sheet"));
        };
        if sheet.print_area.replace(print_area).is_some() {
            return Err(format_error(part, "sheet has multiple print areas"));
        }
    }
    for (index, titles) in
        parse_workbook_print_areas(&bytes, package.limits(), part, "_xlnm.Print_Titles")?
    {
        if let Some(sheet) = sheets.get_mut(index) {
            sheet.print_titles = Some(titles);
        }
    }
    Ok((sheets, date_1904, ole_size))
}

fn apply_ole_size(unit: &mut Unit, ole_size: Option<Dimension>) {
    let Some(size) = ole_size else { return };
    unit.columns = unit.columns.max(size.end_column.saturating_add(1));
    unit.rows = unit.rows.max(size.end_row.saturating_add(1));
    unit.width = unit.width.max(unit.column_axis.offset(unit.columns));
    unit.height = unit.height.max(unit.row_axis.offset(unit.rows));
}

fn parse_workbook_print_areas(
    bytes: &[u8],
    limits: Limits,
    part: &str,
    builtin_name: &str,
) -> Result<Vec<(usize, String)>, Diagnostic> {
    struct CurrentPrintArea {
        depth: usize,
        sheet: usize,
        value: String,
    }

    let mut depth = 0_usize;
    let mut current = None::<CurrentPrintArea>;
    let mut areas = Vec::new();
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                if local_name(name) == "definedName"
                    && optional_attribute(&attributes, "name", part)?.as_deref()
                        == Some(builtin_name)
                {
                    let sheet = required_attribute(&attributes, "localSheetId", part)?
                        .parse::<usize>()
                        .map_err(|_| format_error(part, "print area has an invalid sheet scope"))?;
                    if !empty {
                        current = Some(CurrentPrintArea {
                            depth,
                            sheet,
                            value: String::new(),
                        });
                    }
                }
                depth = depth.saturating_add(usize::from(!empty));
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if local_name(name) == "definedName"
                    && current.as_ref().is_some_and(|area| area.depth == depth)
                {
                    let area = current.take().expect("print-area state is present");
                    let value = area.value.trim().to_owned();
                    if !value.is_empty() {
                        areas.push((area.sheet, value));
                    }
                }
            }
            XmlEvent::Text(text) => {
                if let Some(area) = current.as_mut() {
                    area.value
                        .push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(area) = current.as_mut() {
                    area.value.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(areas)
}

fn parse_chart_named_data(
    package: &Package<'_>,
    workbook_part: &str,
    sheets: &[Sheet],
    relationships: &HashMap<&str, &Relationship>,
    shared_strings: &[SharedString],
) -> Result<HashMap<String, super::drawingml::ChartSourceData>, Diagnostic> {
    struct Name {
        depth: usize,
        key: String,
        formula: String,
    }

    let bytes = package.required_part(workbook_part)?;
    let mut depth = 0_usize;
    let mut current = None::<Name>;
    let mut names = Vec::new();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                if local_name(name) == "definedName"
                    && let Some(key) = optional_attribute(&attributes, "name", workbook_part)?
                    && key.starts_with("_xlchart.")
                    && !empty
                {
                    current = Some(Name {
                        depth,
                        key,
                        formula: String::new(),
                    });
                }
                depth = depth.saturating_add(usize::from(!empty));
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if local_name(name) == "definedName"
                    && current.as_ref().is_some_and(|name| name.depth == depth)
                {
                    names.push(current.take().expect("chart-name state is present"));
                }
            }
            XmlEvent::Text(text) => {
                if let Some(name) = current.as_mut() {
                    name.formula.push_str(
                        &decode_xml_text(text).map_err(|error| with_part(error, workbook_part))?,
                    );
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(name) = current.as_mut() {
                    name.formula.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, workbook_part))?;

    let mut resolved = HashMap::new();
    for name in names {
        let Some((sheet_name, range)) = name.formula.trim().rsplit_once('!') else {
            continue;
        };
        let sheet_name = sheet_name.trim();
        let sheet_name = if sheet_name.starts_with('\'') && sheet_name.ends_with('\'') {
            sheet_name
                .get(1..sheet_name.len().saturating_sub(1))
                .unwrap_or_default()
                .replace("''", "'")
        } else {
            sheet_name.to_owned()
        };
        let Some(sheet) = sheets
            .iter()
            .find(|sheet| sheet.name.eq_ignore_ascii_case(&sheet_name))
        else {
            continue;
        };
        let Some(relationship) = relationships.get(sheet.relationship_id.as_str()) else {
            continue;
        };
        if relationship.external || !relationship.type_uri.ends_with("/worksheet") {
            continue;
        }
        let range = range.trim().replace('$', "").to_ascii_uppercase();
        let Ok(range) = parse_dimension(&range, workbook_part) else {
            continue;
        };
        let values =
            parse_chart_source_range(package, &relationship.target, range, shared_strings)?;
        if !values.values.is_empty() || !values.category_levels.is_empty() {
            resolved.insert(name.key, values);
        }
    }
    Ok(resolved)
}

fn parse_chart_source_range(
    package: &Package<'_>,
    part: &str,
    range: Dimension,
    shared_strings: &[SharedString],
) -> Result<super::drawingml::ChartSourceData, Diagnostic> {
    let cell_count = usize::try_from(range.end_row - range.start_row + 1)
        .ok()
        .and_then(|rows| {
            usize::try_from(range.end_column - range.start_column + 1)
                .ok()
                .and_then(|columns| rows.checked_mul(columns))
        })
        .filter(|count| *count <= package.limits().max_document_objects)
        .ok_or_else(|| object_limit_error(part, "chart source range exceeds the object limit"))?;
    let bytes = package.required_part(part)?;
    let mut depth = 0_usize;
    let mut cell = None::<(usize, CellAddress, Option<usize>, String)>;
    let mut cell_type = String::new();
    let mut cells = HashMap::with_capacity(cell_count);
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "c"
                    && let Some(address) = optional_attribute(&attributes, "r", part)?
                        .and_then(|address| parse_a1(&address.to_ascii_uppercase()))
                    && address.column >= range.start_column
                    && address.column <= range.end_column
                    && address.row >= range.start_row
                    && address.row <= range.end_row
                {
                    cell_type = optional_attribute(&attributes, "t", part)?.unwrap_or_default();
                    cell = Some((depth, address, None, String::new()));
                } else if matches!(local, "v" | "t")
                    && let Some((_, _, value_depth, value)) = cell.as_mut()
                    && !empty
                {
                    *value_depth = Some(depth);
                    value.clear();
                }
                depth = depth.saturating_add(usize::from(!empty));
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if matches!(local, "v" | "t")
                    && let Some((_, _, value_depth, _)) = cell.as_mut()
                    && *value_depth == Some(depth)
                {
                    *value_depth = None;
                } else if local == "c"
                    && cell
                        .as_ref()
                        .is_some_and(|(start, _, _, _)| *start == depth)
                {
                    let (_, address, _, value) = cell.take().expect("chart-cell state is present");
                    let value = if cell_type == "s" {
                        value
                            .parse::<usize>()
                            .ok()
                            .and_then(|index| shared_strings.get(index))
                            .map(|value| value.text.clone())
                            .unwrap_or_default()
                    } else if cell_type == "e" {
                        String::new()
                    } else {
                        value
                    };
                    cells.insert((address.row, address.column), value);
                }
            }
            XmlEvent::Text(text) => {
                if let Some((_, _, Some(_), value)) = cell.as_mut() {
                    value.push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some((_, _, Some(_), value)) = cell.as_mut() {
                    value.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    let mut result = super::drawingml::ChartSourceData::default();
    for row in range.start_row..=range.end_row {
        for column in range.start_column..=range.end_column {
            result.values.push(
                cells
                    .get(&(row, column))
                    .and_then(|s| s.parse::<f32>().ok())
                    .filter(|value| value.is_finite())
                    .unwrap_or(f32::NAN),
            );
        }
    }
    for column in range.start_column..=range.end_column {
        result.category_levels.push(
            (range.start_row..=range.end_row)
                .map(|row| cells.get(&(row, column)).cloned().unwrap_or_default())
                .collect(),
        );
    }
    Ok(result)
}

#[derive(Clone, Copy)]
enum HeaderFooterField {
    OddHeader,
    OddFooter,
    EvenHeader,
    EvenFooter,
    FirstHeader,
    FirstFooter,
}

struct PrintSettingsParser {
    settings: SheetPrintSettings,
    present: bool,
    depth: usize,
    header_footer: Option<(usize, HeaderFooterField)>,
    breaks: Option<(usize, bool)>,
    invalid_break_reported: bool,
    partial_break_reported: bool,
}

impl PrintSettingsParser {
    fn new(print_area: Option<String>) -> Self {
        let settings = SheetPrintSettings {
            print_area,
            ..SheetPrintSettings::default()
        };
        Self {
            present: settings.print_area.is_some(),
            settings,
            depth: 0,
            header_footer: None,
            breaks: None,
            invalid_break_reported: false,
            partial_break_reported: false,
        }
    }

    fn consume(
        &mut self,
        event: &XmlEvent<'_>,
        part: &str,
        diagnostics: &mut Vec<Diagnostic>,
    ) -> Result<(), Diagnostic> {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if matches!(local, "rowBreaks" | "colBreaks") && !*empty {
                    self.breaks = Some((self.depth, local == "rowBreaks"));
                } else if local == "brk"
                    && let Some((parent_depth, rows)) = self.breaks
                    && self.depth == parent_depth + 1
                    && optional_attribute(attributes, "man", part)?
                        .is_some_and(|value| matches!(value.as_str(), "1" | "true"))
                {
                    let number = |name, default| -> Result<u32, Diagnostic> {
                        Ok(optional_attribute(attributes, name, part)?
                            .map_or(default, |value| value.parse::<u32>().unwrap_or(u32::MAX)))
                    };
                    let id = number("id", 0)?;
                    let min = number("min", 0)?;
                    let max = number("max", if rows { 16_383 } else { 1_048_575 })?;
                    if id > 0
                        && id < if rows { 1_048_576 } else { 16_384 }
                        && min <= max
                        && max < if rows { 16_384 } else { 1_048_576 }
                    {
                        self.present = true;
                        let target = if rows {
                            &mut self.settings.row_breaks
                        } else {
                            &mut self.settings.column_breaks
                        };
                        target.push([id, min, max]);
                        if !rows && (min > 0 || max < 1_048_575) && !self.partial_break_reported {
                            diagnostics.push(Diagnostic::warning(DiagnosticCode::UnsupportedFeature, Phase::Render, Fidelity::Approximate,
                                "Partially scoped column page breaks apply to the printable row span").in_part(part));
                            self.partial_break_reported = true;
                        }
                    } else if !self.invalid_break_reported {
                        diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::UnsupportedFeature,
                                Phase::Render,
                                Fidelity::Approximate,
                                "Invalid worksheet manual page breaks were ignored",
                            )
                            .in_part(part),
                        );
                        self.invalid_break_reported = true;
                    }
                } else if local == "pageMargins" {
                    self.present = true;
                    self.settings.margins.left = optional_non_negative_f32_attribute(
                        attributes,
                        "left",
                        part,
                        "left page margin",
                    )?;
                    self.settings.margins.right = optional_non_negative_f32_attribute(
                        attributes,
                        "right",
                        part,
                        "right page margin",
                    )?;
                    self.settings.margins.top = optional_non_negative_f32_attribute(
                        attributes,
                        "top",
                        part,
                        "top page margin",
                    )?;
                    self.settings.margins.bottom = optional_non_negative_f32_attribute(
                        attributes,
                        "bottom",
                        part,
                        "bottom page margin",
                    )?;
                    self.settings.margins.header = optional_non_negative_f32_attribute(
                        attributes,
                        "header",
                        part,
                        "header page margin",
                    )?;
                    self.settings.margins.footer = optional_non_negative_f32_attribute(
                        attributes,
                        "footer",
                        part,
                        "footer page margin",
                    )?;
                } else if local == "sheetView" {
                    self.settings.view_mode =
                        match optional_attribute(attributes, "view", part)?.as_deref() {
                            Some("pageLayout") => Some(SheetViewMode::PageLayout),
                            Some("pageBreakPreview") => Some(SheetViewMode::PageBreakPreview),
                            _ => self.settings.view_mode,
                        };
                    self.present |= self.settings.view_mode.is_some();
                } else if local == "pageSetup" {
                    self.present = true;
                    self.settings.paper_size =
                        optional_u32_attribute(attributes, "paperSize", part)?;
                    self.settings.orientation =
                        match optional_attribute(attributes, "orientation", part)?.as_deref() {
                            None | Some("default") => None,
                            Some("portrait") => Some(SheetOrientation::Portrait),
                            Some("landscape") => Some(SheetOrientation::Landscape),
                            Some(_) => {
                                return Err(format_error(part, "page orientation is invalid"));
                            }
                        };
                    self.settings.scale = optional_u32_attribute(attributes, "scale", part)?;
                    self.settings.fit_to_width =
                        optional_u32_attribute(attributes, "fitToWidth", part)?;
                    self.settings.fit_to_height =
                        optional_u32_attribute(attributes, "fitToHeight", part)?;
                } else if local == "pageSetUpPr" {
                    self.present = true;
                    self.settings.fit_to_page =
                        parse_bool_attribute(attributes, "fitToPage", part)?.unwrap_or(false);
                } else if local == "headerFooter" {
                    self.present = true;
                    self.settings.different_odd_even =
                        parse_bool_attribute(attributes, "differentOddEven", part)?
                            .unwrap_or(false);
                    self.settings.different_first =
                        parse_bool_attribute(attributes, "differentFirst", part)?.unwrap_or(false);
                } else if let Some(field) = match local {
                    "oddHeader" => Some(HeaderFooterField::OddHeader),
                    "oddFooter" => Some(HeaderFooterField::OddFooter),
                    "evenHeader" => Some(HeaderFooterField::EvenHeader),
                    "evenFooter" => Some(HeaderFooterField::EvenFooter),
                    "firstHeader" => Some(HeaderFooterField::FirstHeader),
                    "firstFooter" => Some(HeaderFooterField::FirstFooter),
                    _ => None,
                } {
                    self.present = true;
                    if !*empty {
                        self.header_footer = Some((self.depth, field));
                    }
                }
                self.depth = self.depth.saturating_add(usize::from(!*empty));
            }
            XmlEvent::EndElement { name } => {
                self.depth = self.depth.saturating_sub(1);
                if self.breaks.is_some_and(|(depth, _)| depth == self.depth) {
                    self.breaks = None;
                }
                if matches!(
                    local_name(name),
                    "oddHeader"
                        | "oddFooter"
                        | "evenHeader"
                        | "evenFooter"
                        | "firstHeader"
                        | "firstFooter"
                ) && self
                    .header_footer
                    .is_some_and(|(field_depth, _)| field_depth == self.depth)
                {
                    self.header_footer = None;
                }
            }
            XmlEvent::Text(text) => {
                if let Some((_, field)) = self.header_footer {
                    append_header_footer_text(
                        &mut self.settings,
                        field,
                        &decode_xml_text(text).map_err(|error| with_part(error, part))?,
                    );
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some((_, field)) = self.header_footer {
                    append_header_footer_text(&mut self.settings, field, text);
                }
            }
        }
        Ok(())
    }

    fn finish(self) -> Option<SheetPrintSettings> {
        self.present.then_some(self.settings)
    }
}

#[cfg(test)]
fn parse_worksheet_print_settings(
    bytes: &[u8],
    limits: Limits,
    part: &str,
    print_area: Option<String>,
) -> Result<Option<SheetPrintSettings>, Diagnostic> {
    if ![
        b"rowBreaks".as_slice(),
        b"colBreaks".as_slice(),
        b"pageMargins".as_slice(),
        b"pageSetup".as_slice(),
        b"pageSetUpPr".as_slice(),
        b"headerFooter".as_slice(),
    ]
    .iter()
    .any(|name| contains_bytes(bytes, name))
    {
        return Ok(print_area.map(|print_area| SheetPrintSettings {
            print_area: Some(print_area),
            ..SheetPrintSettings::default()
        }));
    }
    let mut parser = PrintSettingsParser::new(print_area);
    parse_xml(bytes, limits, |event| {
        parser.consume(&event, part, &mut Vec::new())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(parser.finish())
}

fn append_header_footer_text(
    settings: &mut SheetPrintSettings,
    field: HeaderFooterField,
    text: &str,
) {
    let target = match field {
        HeaderFooterField::OddHeader => &mut settings.odd_header,
        HeaderFooterField::OddFooter => &mut settings.odd_footer,
        HeaderFooterField::EvenHeader => &mut settings.even_header,
        HeaderFooterField::EvenFooter => &mut settings.even_footer,
        HeaderFooterField::FirstHeader => &mut settings.first_header,
        HeaderFooterField::FirstFooter => &mut settings.first_footer,
    };
    target.get_or_insert_with(String::new).push_str(text);
}

fn optional_non_negative_f32_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
    label: &str,
) -> Result<Option<f32>, Diagnostic> {
    optional_attribute(attributes, name, part)?
        .map(|value| parse_finite_f32(&value, part, label))
        .transpose()
}

fn optional_u32_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<u32>, Diagnostic> {
    optional_attribute(attributes, name, part)?
        .map(|value| parse_u32(&value, part, name))
        .transpose()
}

fn parse_shared_strings(
    package: &Package<'_>,
    part: &str,
    theme_colors: &ThemeColors,
) -> Result<Vec<SharedString>, Diagnostic> {
    let bytes = package.required_part(part)?;
    parse_shared_strings_xml_with_theme(&bytes, package.limits(), part, theme_colors)
}

#[cfg(test)]
fn parse_shared_strings_xml(
    bytes: &[u8],
    limits: Limits,
    part: &str,
) -> Result<Vec<SharedString>, Diagnostic> {
    parse_shared_strings_xml_with_theme(bytes, limits, part, &default_theme_colors())
}

fn parse_shared_strings_xml_with_theme(
    bytes: &[u8],
    limits: Limits,
    part: &str,
    theme_colors: &ThemeColors,
) -> Result<Vec<SharedString>, Diagnostic> {
    #[derive(Debug)]
    struct CurrentSharedString {
        depth: usize,
        text_depth: Option<usize>,
        run_depth: Option<usize>,
        properties_depth: Option<usize>,
        phonetic_depth: Option<usize>,
        value: SharedString,
        run: Option<SharedStringRun>,
    }

    let mut depth = 0_usize;
    let mut current = None;
    let mut strings = Vec::new();
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "si" {
                    if current.is_some() {
                        return Err(format_error(part, "nested shared strings are invalid"));
                    }
                    if strings.len() >= limits.max_document_objects {
                        return Err(Diagnostic::fatal(
                            DiagnosticCode::ObjectLimit,
                            Phase::Parse,
                            None,
                            "shared-string table exceeds the configured item limit",
                        )
                        .in_part(part));
                    }
                    if empty {
                        strings.push(SharedString::default());
                    } else {
                        current = Some(CurrentSharedString {
                            depth,
                            text_depth: None,
                            run_depth: None,
                            properties_depth: None,
                            phonetic_depth: None,
                            value: SharedString::default(),
                            run: None,
                        });
                    }
                } else if let Some(shared) = current.as_mut() {
                    if local == "rPh" {
                        shared.phonetic_depth = (!empty).then_some(depth);
                    } else if shared.phonetic_depth.is_none() && local == "r" {
                        shared.run_depth = (!empty).then_some(depth);
                        shared.run = Some(SharedStringRun::default());
                    } else if shared.phonetic_depth.is_none() && local == "rPr" {
                        shared.properties_depth = (!empty).then_some(depth);
                    } else if shared.properties_depth.is_some() {
                        apply_spreadsheet_run_property(
                            local,
                            &attributes,
                            shared.run.as_mut(),
                            theme_colors,
                            part,
                        )?;
                    } else if shared.phonetic_depth.is_none() && local == "t" {
                        shared.text_depth = (!empty).then_some(depth);
                    }
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if local == "t" {
                    if let Some(shared) = current.as_mut()
                        && shared.text_depth == Some(depth)
                    {
                        shared.text_depth = None;
                    }
                } else if local == "rPr" {
                    if let Some(shared) = current.as_mut()
                        && shared.properties_depth == Some(depth)
                    {
                        shared.properties_depth = None;
                    }
                } else if local == "r" {
                    if let Some(shared) = current.as_mut()
                        && shared.run_depth == Some(depth)
                    {
                        shared.run_depth = None;
                        if let Some(run) = shared.run.take() {
                            shared.value.runs.push(run);
                        }
                    }
                } else if local == "rPh" {
                    if let Some(shared) = current.as_mut()
                        && shared.phonetic_depth == Some(depth)
                    {
                        shared.phonetic_depth = None;
                    }
                } else if local == "si"
                    && current.as_ref().is_some_and(|shared| shared.depth == depth)
                {
                    let shared = current.take().ok_or_else(|| {
                        format_error(part, "shared-string state was lost before closing")
                    })?;
                    strings.push(shared.value);
                }
            }
            XmlEvent::Text(text) => {
                if let Some(shared) = current
                    .as_mut()
                    .filter(|shared| shared.text_depth.is_some())
                {
                    let text = decode_xml_text(text).map_err(|error| with_part(error, part))?;
                    shared.value.text.push_str(&text);
                    if let Some(run) = shared.run.as_mut() {
                        run.text.push_str(&text);
                    }
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(shared) = current
                    .as_mut()
                    .filter(|shared| shared.text_depth.is_some())
                {
                    shared.value.text.push_str(text);
                    if let Some(run) = shared.run.as_mut() {
                        run.text.push_str(text);
                    }
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(strings)
}

fn apply_spreadsheet_run_property(
    local: &str,
    attributes: &[XmlAttribute<'_>],
    run: Option<&mut SharedStringRun>,
    theme_colors: &ThemeColors,
    part: &str,
) -> Result<(), Diagnostic> {
    let Some(run) = run else {
        return Ok(());
    };
    match local {
        "rFont" => run.font_family = optional_attribute(attributes, "val", part)?,
        "sz" => {
            if let Some(size) = optional_attribute(attributes, "val", part)? {
                run.font_size = Some(points_to_pixels(
                    parse_finite_f32(&size, part, "rich-text font size")?,
                    part,
                    "rich-text font size",
                )?);
            }
        }
        "color" => run.color = parse_color(attributes, theme_colors, part)?,
        "b" => run.bold = Some(boolean_value(attributes, part)?.unwrap_or(true)),
        "i" => run.italic = Some(boolean_value(attributes, part)?.unwrap_or(true)),
        "strike" => run.strikethrough = Some(boolean_value(attributes, part)?.unwrap_or(true)),
        "u" => {
            run.underline =
                Some(optional_attribute(attributes, "val", part)?.as_deref() != Some("none"));
        }
        "vertAlign" => {
            run.baseline_shift = match optional_attribute(attributes, "val", part)?.as_deref() {
                Some("superscript") => Some(0.35),
                Some("subscript") => Some(-0.2),
                Some("baseline") | None => Some(0.0),
                Some(_) => None,
            };
        }
        _ => {}
    }
    Ok(())
}

fn default_theme_colors() -> ThemeColors {
    [
        0xffff_ffff,
        0x0000_00ff,
        0xe7e6_e6ff,
        0x4454_6aff,
        0x5b9b_d5ff,
        0xed7d_31ff,
        0xa5a5_a5ff,
        0xffc0_00ff,
        0x4472_c4ff,
        0x70ad_47ff,
        0x0563_c1ff,
        0x954f_72ff,
    ]
}

fn theme_color_index(name: &str) -> Option<usize> {
    // SpreadsheetML indexes the first four slots as light1, dark1, light2, dark2.
    match name {
        "lt1" => Some(0),
        "dk1" => Some(1),
        "lt2" => Some(2),
        "dk2" => Some(3),
        "accent1" => Some(4),
        "accent2" => Some(5),
        "accent3" => Some(6),
        "accent4" => Some(7),
        "accent5" => Some(8),
        "accent6" => Some(9),
        "hlink" => Some(10),
        "folHlink" => Some(11),
        _ => None,
    }
}

fn parse_theme_rgb(value: &str, part: &str) -> Result<u32, Diagnostic> {
    if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(format_error(
            part,
            "theme color must contain six hexadecimal digits",
        ));
    }
    let color =
        u32::from_str_radix(value, 16).map_err(|_| format_error(part, "theme color is invalid"))?;
    Ok((color << 8) | 0xff)
}

fn parse_theme(
    package: &Package<'_>,
    part: &str,
) -> Result<(ThemeColors, DrawingMlThemeLineStyles), Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut colors = default_theme_colors();
    let mut depth = 0_usize;
    let mut color_scheme_depth = None;
    let mut current_slot = None::<(usize, usize)>;
    let mut line_styles = DrawingMlThemeLineStyles::default();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                line_styles.start(local, &attributes, depth, empty, part)?;
                if local == "clrScheme" && color_scheme_depth.is_none() {
                    color_scheme_depth = (!empty).then_some(depth);
                } else if color_scheme_depth.is_some()
                    && current_slot.is_none()
                    && let Some(index) = theme_color_index(local)
                {
                    current_slot = (!empty).then_some((depth, index));
                } else if let Some((_, index)) = current_slot {
                    let value = if local == "srgbClr" {
                        Some(required_attribute(&attributes, "val", part)?)
                    } else if local == "sysClr" {
                        optional_attribute(&attributes, "lastClr", part)?
                    } else {
                        None
                    };
                    if let Some(value) = value {
                        colors[index] = parse_theme_rgb(&value, part)?;
                    }
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                line_styles.end(local, depth);
                if current_slot.is_some_and(|(slot_depth, index)| {
                    slot_depth == depth && theme_color_index(local) == Some(index)
                }) {
                    current_slot = None;
                }
                if local == "clrScheme" && color_scheme_depth == Some(depth) {
                    color_scheme_depth = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok((colors, line_styles))
}

fn parse_styles(
    package: &Package<'_>,
    part: &str,
    theme_colors: ThemeColors,
) -> Result<Styles, Diagnostic> {
    #[derive(Debug)]
    struct Current<T> {
        depth: usize,
        value: T,
    }
    #[derive(Debug)]
    struct BorderState {
        sides: [BorderSide; 4],
        diagonal: BorderSide,
        diagonal_up: bool,
        diagonal_down: bool,
        active_side_depth: Option<(usize, usize)>,
    }

    let bytes = package.required_part(part)?;
    let limits = package.limits();
    let mut depth = 0_usize;
    let mut fonts = Vec::<FontRecord>::new();
    let mut fills = Vec::<FillRecord>::new();
    let mut borders = Vec::<BorderRecord>::new();
    let mut formats = HashMap::<u32, String>::new();
    let mut diagnostics = Vec::new();
    let mut cell_formats = Vec::<RawCellFormat>::new();
    let mut current_font = None::<Current<FontRecord>>;
    let mut current_fill = None::<Current<SpreadsheetFillCapture>>;
    let mut current_border = None::<Current<BorderState>>;
    let mut cell_xfs_depth = None::<usize>;
    let mut num_fmts_depth = None::<usize>;
    let mut current_format = None::<Current<RawCellFormat>>;

    parse_xml(&bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "numFmts" {
                    num_fmts_depth = (!empty).then_some(depth);
                } else if local == "cellXfs" {
                    if cell_xfs_depth.is_some() {
                        return Err(format_error(part, "nested cell style tables are invalid"));
                    }
                    cell_xfs_depth = (!empty).then_some(depth);
                } else if local == "numFmt"
                    && num_fmts_depth.is_some_and(|table_depth| depth == table_depth + 1)
                {
                    let id = parse_u32(
                        &required_attribute(&attributes, "numFmtId", part)?,
                        part,
                        "number format id",
                    )?;
                    let code = optional_attribute(&attributes, "formatCode16", part)?
                        .unwrap_or(required_attribute(&attributes, "formatCode", part)?);
                    if let Some(existing) = formats.get(&id) {
                        if existing != &code {
                            return Err(format_error(
                                part,
                                "number format id has conflicting definitions",
                            ));
                        }
                    } else {
                        formats.insert(id, code);
                    }
                } else if local == "font" && current_font.is_none() {
                    ensure_style_capacity(fonts.len(), limits, part)?;
                    let state = Current {
                        depth,
                        value: FontRecord::default(),
                    };
                    if empty {
                        fonts.push(state.value);
                    } else {
                        current_font = Some(state);
                    }
                } else if let Some(font) = current_font.as_mut() {
                    match local {
                        "b" => font.value.bold = boolean_value(&attributes, part)?.unwrap_or(true),
                        "i" => {
                            font.value.italic = boolean_value(&attributes, part)?.unwrap_or(true)
                        }
                        "u" => {
                            font.value.underline = optional_attribute(&attributes, "val", part)?
                                .as_deref()
                                != Some("none")
                        }
                        "strike" => {
                            font.value.strikethrough =
                                boolean_value(&attributes, part)?.unwrap_or(true)
                        }
                        "vertAlign" => {
                            font.value.baseline_ratio =
                                match optional_attribute(&attributes, "val", part)?.as_deref() {
                                    Some("superscript") => 0.35,
                                    Some("subscript") => -0.2,
                                    Some("baseline") | None => 0.0,
                                    Some(_) => {
                                        return Err(format_error(
                                            part,
                                            "XLSX font vertical alignment is invalid",
                                        ));
                                    }
                                };
                        }
                        "name" => {
                            if let Some(name) = optional_attribute(&attributes, "val", part)?
                                && !name.is_empty()
                            {
                                font.value.family = name;
                            }
                        }
                        "sz" => {
                            if let Some(size) = optional_attribute(&attributes, "val", part)? {
                                font.value.size = points_to_pixels(
                                    parse_finite_f32(&size, part, "font size")?,
                                    part,
                                    "font size",
                                )?;
                            }
                        }
                        "color" => {
                            if let Some(color) = parse_color(&attributes, &theme_colors, part)? {
                                font.value.color = color;
                            }
                        }
                        _ => {}
                    }
                } else if local == "fill" && current_fill.is_none() {
                    ensure_style_capacity(fills.len(), limits, part)?;
                    let state = Current {
                        depth,
                        value: SpreadsheetFillCapture::default(),
                    };
                    if empty {
                        fills.push(FillRecord::default());
                    } else {
                        current_fill = Some(state);
                    }
                } else if let Some(fill) = current_fill.as_mut() {
                    fill.value.start(local, &attributes, &theme_colors, part)?;
                } else if local == "border" && current_border.is_none() {
                    ensure_style_capacity(borders.len(), limits, part)?;
                    let state = Current {
                        depth,
                        value: BorderState {
                            sides: [BorderSide::default(); 4],
                            diagonal: BorderSide::default(),
                            diagonal_up: parse_bool_attribute(&attributes, "diagonalUp", part)?
                                .unwrap_or(false),
                            diagonal_down: parse_bool_attribute(&attributes, "diagonalDown", part)?
                                .unwrap_or(false),
                            active_side_depth: None,
                        },
                    };
                    if empty {
                        borders.push(BorderRecord::default());
                    } else {
                        current_border = Some(state);
                    }
                } else if let Some(border) = current_border.as_mut() {
                    if let Some(side) = match local {
                        "left" => Some(0),
                        "right" => Some(1),
                        "top" => Some(2),
                        "bottom" => Some(3),
                        _ => None,
                    } {
                        if let Some(style) = optional_attribute(&attributes, "style", part)? {
                            border.value.sides[side].width = border_width(&style);
                            border.value.active_side_depth = (!empty).then_some((depth, side));
                        }
                    } else if local == "diagonal" {
                        if let Some(style) = optional_attribute(&attributes, "style", part)? {
                            border.value.diagonal.width = border_width(&style);
                            border.value.active_side_depth = (!empty).then_some((depth, 4));
                        }
                    } else if local == "color"
                        && let Some((_, side)) = border.value.active_side_depth
                        && let Some(color) = parse_color(&attributes, &theme_colors, part)?
                    {
                        if side < border.value.sides.len() {
                            border.value.sides[side].color = color;
                        } else {
                            border.value.diagonal.color = color;
                        }
                    }
                } else if local == "xf" && cell_xfs_depth.is_some() && current_format.is_none() {
                    ensure_style_capacity(cell_formats.len(), limits, part)?;
                    let state = Current {
                        depth,
                        value: RawCellFormat {
                            font_id: parse_optional_usize(&attributes, "fontId", part)?
                                .unwrap_or(0),
                            fill_id: parse_optional_usize(&attributes, "fillId", part)?
                                .unwrap_or(0),
                            border_id: parse_optional_usize(&attributes, "borderId", part)?
                                .unwrap_or(0),
                            number_format_id: parse_optional_u32(&attributes, "numFmtId", part)?
                                .unwrap_or(0),
                            align: None,
                            vertical_align: None,
                            wrap: false,
                            shrink_to_fit: false,
                            indent: 0,
                            relative_indent: 0,
                            direction: None,
                            rotation_degrees: 0.0,
                        },
                    };
                    if empty {
                        cell_formats.push(state.value);
                    } else {
                        current_format = Some(state);
                    }
                } else if local == "alignment"
                    && let Some(format) = current_format.as_mut()
                {
                    if let Some(horizontal) = optional_attribute(&attributes, "horizontal", part)? {
                        format.value.align = parse_horizontal_alignment(&horizontal);
                    }
                    if let Some(vertical) = optional_attribute(&attributes, "vertical", part)? {
                        format.value.vertical_align = parse_vertical_alignment(&vertical);
                    }
                    format.value.wrap =
                        parse_bool_attribute(&attributes, "wrapText", part)?.unwrap_or(false);
                    format.value.shrink_to_fit =
                        parse_bool_attribute(&attributes, "shrinkToFit", part)?.unwrap_or(false);
                    format.value.indent = parse_optional_u32(&attributes, "indent", part)?
                        .unwrap_or(0)
                        .min(250);
                    if let Some(relative) = optional_attribute(&attributes, "relativeIndent", part)?
                    {
                        format.value.relative_indent =
                            parse_i64(&relative, part, "relative cell indentation")?
                                .clamp(-250, 250) as i32;
                    }
                    format.value.direction =
                        match parse_optional_u32(&attributes, "readingOrder", part)? {
                            Some(1) => Some(TextDirection::Ltr),
                            Some(2) => Some(TextDirection::Rtl),
                            Some(0) | None => None,
                            Some(_) => {
                                return Err(format_error(part, "XLSX reading order is invalid"));
                            }
                        };
                    if let Some(rotation) = optional_attribute(&attributes, "textRotation", part)? {
                        format.value.rotation_degrees = parse_excel_text_rotation(&rotation, part)?;
                    }
                }

                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if local == "font"
                    && current_font
                        .as_ref()
                        .is_some_and(|current| current.depth == depth)
                {
                    fonts.push(current_font.take().expect("checked font state").value);
                } else if local == "fill"
                    && current_fill
                        .as_ref()
                        .is_some_and(|current| current.depth == depth)
                {
                    let fill = current_fill
                        .take()
                        .expect("checked fill state")
                        .value
                        .finish();
                    fills.push(match fill {
                        Some(ChartFill::Solid(color)) => FillRecord {
                            color,
                            solid: true,
                            paint: None,
                        },
                        paint => FillRecord {
                            paint,
                            ..FillRecord::default()
                        },
                    });
                } else if matches!(local, "left" | "right" | "top" | "bottom" | "diagonal") {
                    if let Some(border) = current_border.as_mut()
                        && border
                            .value
                            .active_side_depth
                            .is_some_and(|(side_depth, _)| side_depth == depth)
                    {
                        border.value.active_side_depth = None;
                    }
                } else if local == "border"
                    && current_border
                        .as_ref()
                        .is_some_and(|current| current.depth == depth)
                {
                    let border = current_border.take().expect("checked border state").value;
                    borders.push(BorderRecord {
                        sides: border.sides,
                        diagonal: border.diagonal,
                        diagonal_up: border.diagonal_up,
                        diagonal_down: border.diagonal_down,
                        has_border: border.sides.iter().any(|side| side.width > 0.0)
                            || border.diagonal.width > 0.0
                                && (border.diagonal_up || border.diagonal_down),
                    });
                } else if local == "xf"
                    && current_format
                        .as_ref()
                        .is_some_and(|current| current.depth == depth)
                {
                    cell_formats.push(current_format.take().expect("checked format state").value);
                } else if local == "cellXfs" && cell_xfs_depth == Some(depth) {
                    cell_xfs_depth = None;
                } else if local == "numFmts" && num_fmts_depth == Some(depth) {
                    num_fmts_depth = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;

    if fonts.is_empty() {
        fonts.push(FontRecord::default());
    }
    if fills.is_empty() {
        fills.push(FillRecord::default());
    }
    if borders.is_empty() {
        borders.push(BorderRecord::default());
    }
    if cell_formats.is_empty() {
        cell_formats.push(RawCellFormat::default());
    }

    let cells = cell_formats
        .into_iter()
        .map(|format| {
            let font = match fonts.get(format.font_id) {
                Some(font) => font,
                None => {
                    if !diagnostics
                        .iter()
                        .any(|diagnostic: &Diagnostic| diagnostic.message.contains("missing font"))
                    {
                        diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::UnsupportedFeature,
                                Phase::Parse,
                                Fidelity::Approximate,
                                "XLSX cell style references a missing font; the default font was substituted",
                            )
                            .in_part(part),
                        );
                    }
                    fonts.first().expect("styles always contain a default font")
                }
            };
            let fill = fills
                .get(format.fill_id)
                .ok_or_else(|| format_error(part, "cell style references a missing fill"))?;
            let border = borders
                .get(format.border_id)
                .ok_or_else(|| format_error(part, "cell style references a missing border"))?;
            Ok(CellStyle {
                font_family: font.family.clone(),
                font_size: font.size,
                color: font.color,
                bold: font.bold,
                italic: font.italic,
                underline: font.underline,
                strikethrough: font.strikethrough,
                baseline_shift: font.baseline_ratio * font.size,
                fill: fill.color,
                fill_paint: fill.paint.clone(),
                has_fill: fill.solid || fill.paint.is_some(),
                borders: border.sides,
                diagonal_border: border.diagonal,
                diagonal_up: border.diagonal_up,
                diagonal_down: border.diagonal_down,
                has_border: border.has_border,
                align: format.align.unwrap_or(TextAlign::Start),
                general_alignment: format.align.is_none(),
                vertical_align: format.vertical_align.unwrap_or(TextVerticalAlign::Bottom),
                wrap: format.wrap,
                shrink_to_fit: format.shrink_to_fit,
                indent: format.indent,
                relative_indent: format.relative_indent,
                direction: format.direction.unwrap_or(TextDirection::Auto),
                rotation_degrees: format.rotation_degrees,
                number_format: formats
                    .get(&format.number_format_id)
                    .cloned()
                    .or_else(|| builtin_number_format(format.number_format_id).map(str::to_owned)),
            })
        })
        .collect::<Result<Vec<_>, Diagnostic>>()?;
    let normal_font = fonts.first().cloned();
    let differentials = parse_differential_styles_xml(&bytes, limits, &theme_colors, part)?;
    let table_styles = parse_custom_table_styles_xml(&bytes, limits, &differentials, part)?;
    Ok(Styles {
        cells,
        normal_font,
        differentials,
        table_styles,
        theme_colors,
        chart_theme_colors: None,
        theme_line_styles: DrawingMlThemeLineStyles::default(),
        diagnostics,
    })
}

fn parse_custom_table_styles_xml(
    bytes: &[u8],
    limits: Limits,
    differentials: &[DifferentialStyle],
    part: &str,
) -> Result<HashMap<String, DifferentialStyle>, Diagnostic> {
    let mut current_name = None;
    let mut table_styles = HashMap::new();
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name, attributes, ..
            } => match local_name(name) {
                "tableStyle" => current_name = optional_attribute(&attributes, "name", part)?,
                "tableStyleElement"
                    if optional_attribute(&attributes, "type", part)?.as_deref()
                        == Some("wholeTable") =>
                {
                    if let (Some(name), Some(differential)) = (
                        current_name.as_ref(),
                        parse_optional_usize(&attributes, "dxfId", part)?
                            .and_then(|index| differentials.get(index)),
                    ) {
                        table_styles
                            .entry(name.clone())
                            .or_insert_with(|| differential.clone());
                    }
                }
                _ => {}
            },
            XmlEvent::EndElement { name } if local_name(name) == "tableStyle" => {
                current_name = None;
            }
            XmlEvent::EndElement { .. } | XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(table_styles)
}

// SpreadsheetML keeps its own pattern names and density rules; only bitmap painting is shared.
#[derive(Debug, Default)]
struct SpreadsheetFillCapture {
    pattern: Option<String>,
    foreground: Option<u32>,
    background: Option<u32>,
    gradient: bool,
    path: bool,
    degree: f32,
    focus: [f32; 4],
    stop: f32,
    stops: Vec<GradientStop>,
}

impl SpreadsheetFillCapture {
    fn start(
        &mut self,
        local: &str,
        attributes: &[XmlAttribute<'_>],
        theme: &ThemeColors,
        part: &str,
    ) -> Result<(), Diagnostic> {
        let number = |key| -> Result<f32, Diagnostic> {
            optional_attribute(attributes, key, part)?
                .map(|value| parse_finite_f32(&value, part, key))
                .transpose()
                .map(|value| value.unwrap_or(0.0))
        };
        match local {
            "patternFill" => self.pattern = optional_attribute(attributes, "patternType", part)?,
            "fgColor" => self.foreground = parse_color(attributes, theme, part)?,
            "bgColor" => self.background = parse_color(attributes, theme, part)?,
            "gradientFill" => {
                self.gradient = true;
                self.path =
                    optional_attribute(attributes, "type", part)?.as_deref() == Some("path");
                self.degree = number("degree")?;
                self.focus = [
                    number("left")?,
                    number("top")?,
                    number("right")?,
                    number("bottom")?,
                ]
                .map(|value| value.clamp(0.0, 1.0));
            }
            "stop" if self.gradient => self.stop = number("position")?.clamp(0.0, 1.0),
            "color" if self.gradient => {
                if let Some(color) = parse_color(attributes, theme, part)? {
                    self.stops.push(GradientStop {
                        offset: self.stop,
                        color,
                    });
                }
            }
            _ => {}
        }
        Ok(())
    }

    fn finish(mut self) -> Option<ChartFill> {
        if self.gradient {
            if self.stops.is_empty() {
                return None;
            }
            self.stops.sort_by(|a, b| a.offset.total_cmp(&b.offset));
            return Some(if self.path {
                ChartFill::Image(spreadsheet_path_gradient(self.focus, self.stops))
            } else {
                ChartFill::LinearGradient {
                    angle_degrees: self.degree,
                    angle_scaled: false,
                    stops: self.stops,
                }
            });
        }
        match self.pattern.as_deref() {
            Some("solid") => Some(ChartFill::Solid(
                self.foreground.or(self.background).unwrap_or(0xffff_ffff),
            )),
            None => self.background.or(self.foreground).map(ChartFill::Solid),
            pattern => spreadsheet_pattern_fill(
                pattern,
                self.foreground.unwrap_or(0x0000_00ff),
                self.background.unwrap_or(0xffff_ffff),
            ),
        }
    }
}

// SpreadsheetML interpolates linearly from an inner rectangle to the cell edges.
// DrawingML's rectangular gradient uses an area-based ramp, so it is not equivalent.
fn spreadsheet_path_gradient(focus: [f32; 4], stops: Vec<GradientStop>) -> Paint {
    let [left, top, right, bottom] = focus;
    let bounds = Rect {
        x: 0.0,
        y: 0.0,
        width: 1.0,
        height: 1.0,
    };
    let mut children = Vec::new();
    for (points, end) in [
        (
            [(0.0, 0.0), (1.0, 0.0), (right, top), (left, top)],
            (left, top, left, 0.0),
        ),
        (
            [(1.0, 0.0), (1.0, 1.0), (right, bottom), (right, top)],
            (right, top, 1.0, top),
        ),
        (
            [(1.0, 1.0), (0.0, 1.0), (left, bottom), (right, bottom)],
            (left, bottom, left, 1.0),
        ),
        (
            [(0.0, 1.0), (0.0, 0.0), (left, top), (left, bottom)],
            (left, top, 0.0, top),
        ),
    ] {
        if end.0 == end.2 && end.1 == end.3 {
            continue;
        }
        let mut commands = vec![PathCommand::MoveTo {
            x: points[0].0,
            y: points[0].1,
        }];
        commands.extend(
            points[1..]
                .iter()
                .map(|&(x, y)| PathCommand::LineTo { x, y }),
        );
        commands.push(PathCommand::ClosePath);
        children.push(VisualBrushChild {
            bounds,
            visual: Visual::PaintedShape {
                geometry: Geometry::Path {
                    commands,
                    fill_rule: crate::model::FillRule::NonZero,
                },
                fill: Paint::LinearGradient {
                    x0: end.0,
                    y0: end.1,
                    x1: end.2,
                    y1: end.3,
                    stops: stops.clone(),
                },
                stroke: Paint::None,
                stroke_width: 0.0,
            },
        });
    }
    if right > left && bottom > top {
        children.push(VisualBrushChild {
            bounds: Rect {
                x: left,
                y: top,
                width: right - left,
                height: bottom - top,
            },
            visual: Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: Paint::Solid(stops[0].color),
                stroke: Paint::None,
                stroke_width: 0.0,
            },
        });
    }
    Paint::Visual {
        viewbox: bounds,
        viewport: bounds,
        viewbox_relative: false,
        viewport_relative: true,
        tile_mode: TileMode::None,
        stretch: StretchMode::Fill,
        alignment_x: 0.0,
        alignment_y: 0.0,
        transform: AffineTransform::IDENTITY,
        relative_transform: AffineTransform::IDENTITY,
        opacity: 1.0,
        children,
    }
}

fn spreadsheet_pattern_fill(
    pattern: Option<&str>,
    foreground: u32,
    background: u32,
) -> Option<ChartFill> {
    let pattern = pattern?;
    if !matches!(
        pattern,
        "mediumGray"
            | "darkGray"
            | "lightGray"
            | "gray125"
            | "gray0625"
            | "darkHorizontal"
            | "darkVertical"
            | "darkDown"
            | "darkUp"
            | "darkGrid"
            | "darkTrellis"
            | "lightHorizontal"
            | "lightVertical"
            | "lightDown"
            | "lightUp"
            | "lightGrid"
            | "lightTrellis"
    ) {
        return None;
    }
    Some(ChartFill::Pattern {
        preset: format!("xlsx:{pattern}"),
        foreground,
        background,
    })
}

fn parse_differential_styles_xml(
    bytes: &[u8],
    limits: Limits,
    theme_colors: &ThemeColors,
    part: &str,
) -> Result<Vec<DifferentialStyle>, Diagnostic> {
    #[derive(Debug)]
    struct DxfState {
        depth: usize,
        style: DifferentialStyle,
        font_depth: Option<usize>,
        fill_depth: Option<usize>,
        fill: SpreadsheetFillCapture,
    }

    let mut depth = 0_usize;
    let mut current = None::<DxfState>;
    let mut differentials = Vec::new();
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "dxf" && current.is_none() {
                    ensure_style_capacity(differentials.len(), limits, part)?;
                    let state = DxfState {
                        depth,
                        style: DifferentialStyle::default(),
                        font_depth: None,
                        fill_depth: None,
                        fill: SpreadsheetFillCapture::default(),
                    };
                    if empty {
                        differentials.push(state.style);
                    } else {
                        current = Some(state);
                    }
                } else if let Some(dxf) = current.as_mut() {
                    if local == "font" {
                        dxf.font_depth = (!empty).then_some(depth);
                    } else if local == "fill" {
                        dxf.fill_depth = (!empty).then_some(depth);
                    } else if dxf.font_depth.is_some() {
                        match local {
                            "b" => {
                                dxf.style.bold =
                                    Some(boolean_value(&attributes, part)?.unwrap_or(true));
                            }
                            "i" => {
                                dxf.style.italic =
                                    Some(boolean_value(&attributes, part)?.unwrap_or(true));
                            }
                            "color" => {
                                dxf.style.color = parse_color(&attributes, theme_colors, part)?;
                            }
                            _ => {}
                        }
                    } else if dxf.fill_depth.is_some() {
                        dxf.fill.start(local, &attributes, theme_colors, part)?;
                    }
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if let Some(dxf) = current.as_mut() {
                    if local == "font" && dxf.font_depth == Some(depth) {
                        dxf.font_depth = None;
                    } else if local == "fill" && dxf.fill_depth == Some(depth) {
                        dxf.fill_depth = None;
                    }
                }
                if local == "dxf" && current.as_ref().is_some_and(|dxf| dxf.depth == depth) {
                    let mut dxf = current
                        .take()
                        .ok_or_else(|| format_error(part, "differential style state was lost"))?;
                    match dxf.fill.finish() {
                        Some(ChartFill::Solid(color)) => dxf.style.fill = Some(color),
                        paint => dxf.style.fill_paint = paint,
                    }
                    differentials.push(dxf.style);
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(differentials)
}

fn parse_conditional_formatting_xml(
    bytes: &[u8],
    limits: Limits,
    theme_colors: &ThemeColors,
    part: &str,
) -> Result<(Vec<ConditionalFormatting>, bool), Diagnostic> {
    if !contains_bytes(bytes, b"conditionalFormatting") {
        return Ok((Vec::new(), false));
    }
    #[derive(Debug)]
    struct BlockState {
        depth: usize,
        ranges: Vec<Dimension>,
        rules: Vec<ConditionalRule>,
    }
    #[derive(Debug)]
    struct RuleState {
        depth: usize,
        rule_type: String,
        priority: u32,
        stop_if_true: bool,
        dxf_id: Option<usize>,
        operator: Option<CellOperator>,
        formula_depth: Option<usize>,
        formula: String,
        formulas: Vec<String>,
        thresholds: Vec<ConditionalValue>,
        inclusive: Vec<bool>,
        colors: Vec<u32>,
        icon_set: String,
        reverse: bool,
        show_value: bool,
    }

    let mut depth = 0_usize;
    let mut current_block = None::<BlockState>;
    let mut current_rule = None::<RuleState>;
    let mut blocks = Vec::new();
    let mut unsupported = false;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                // The x14 extension reuses this local name but carries sqref as a child element.
                if local == "conditionalFormatting" && depth == 1 {
                    if current_block.is_some() {
                        return Err(format_error(
                            part,
                            "nested conditional formatting is invalid",
                        ));
                    }
                    let ranges = required_attribute(&attributes, "sqref", part)?
                        .split_ascii_whitespace()
                        .map(|range| parse_dimension(range, part))
                        .collect::<Result<Vec<_>, _>>()?;
                    if ranges.is_empty() {
                        return Err(format_error(part, "conditional formatting has no ranges"));
                    }
                    let block = BlockState {
                        depth,
                        ranges,
                        rules: Vec::new(),
                    };
                    if empty {
                        blocks.push(ConditionalFormatting {
                            ranges: block.ranges,
                            rules: block.rules,
                        });
                    } else {
                        current_block = Some(block);
                    }
                } else if local == "conditionalFormatting" {
                    unsupported = true;
                } else if local == "cfRule" && current_block.is_some() {
                    if current_rule.is_some() {
                        return Err(format_error(part, "nested conditional rules are invalid"));
                    }
                    let rule = RuleState {
                        depth,
                        rule_type: required_attribute(&attributes, "type", part)?,
                        priority: parse_optional_u32(&attributes, "priority", part)?
                            .unwrap_or(u32::MAX),
                        stop_if_true: parse_bool_attribute(&attributes, "stopIfTrue", part)?
                            .unwrap_or(false),
                        dxf_id: parse_optional_usize(&attributes, "dxfId", part)?,
                        operator: optional_attribute(&attributes, "operator", part)?
                            .as_deref()
                            .and_then(parse_cell_operator),
                        formula_depth: None,
                        formula: String::new(),
                        formulas: Vec::new(),
                        thresholds: Vec::new(),
                        inclusive: Vec::new(),
                        colors: Vec::new(),
                        icon_set: "3TrafficLights1".to_owned(),
                        reverse: false,
                        show_value: true,
                    };
                    if empty {
                        if crate::calculation::calculated_condition(&rule.rule_type) {
                            if let Some(dxf_id) = rule.dxf_id {
                                let block = current_block.as_mut().unwrap();
                                if block.rules.len() >= limits.max_document_objects {
                                    return Err(object_limit_error(part, "conditional rule count exceeds the configured object limit"));
                                }
                                block.rules.push(ConditionalRule {
                                    priority: rule.priority, stop_if_true: rule.stop_if_true,
                                    kind: ConditionalRuleKind::Calculated { dxf_id, matches: HashSet::new() },
                                });
                            } else { unsupported = true; }
                        } else { unsupported = true; }
                    } else {
                        current_rule = Some(rule);
                    }
                } else if local == "formula"
                    && let Some(rule) = current_rule.as_mut()
                {
                    if rule.formula_depth.is_some() {
                        return Err(format_error(
                            part,
                            "nested conditional formulas are invalid",
                        ));
                    }
                    rule.formula.clear();
                    rule.formula_depth = (!empty).then_some(depth);
                    if empty {
                        rule.formulas.push(String::new());
                    }
                } else if local == "cfvo"
                    && let Some(rule) = current_rule.as_mut()
                {
                    match parse_conditional_value(&attributes, part)? {
                        Some(value) => {
                            rule.thresholds.push(value);
                            rule.inclusive.push(
                                parse_bool_attribute(&attributes, "gte", part)?.unwrap_or(true),
                            );
                        }
                        None => unsupported = true,
                    }
                } else if local == "color"
                    && let Some(rule) = current_rule.as_mut()
                    && matches!(rule.rule_type.as_str(), "colorScale" | "dataBar")
                    && let Some(color) = parse_color(&attributes, theme_colors, part)?
                {
                    rule.colors.push(color);
                } else if local == "iconSet"
                    && let Some(rule) = current_rule.as_mut()
                {
                    if let Some(icon_set) = optional_attribute(&attributes, "iconSet", part)? {
                        rule.icon_set = icon_set;
                    }
                    rule.reverse =
                        parse_bool_attribute(&attributes, "reverse", part)?.unwrap_or(false);
                    rule.show_value =
                        parse_bool_attribute(&attributes, "showValue", part)?.unwrap_or(true);
                } else if local == "dataBar"
                    && let Some(rule) = current_rule.as_mut()
                {
                    rule.show_value =
                        parse_bool_attribute(&attributes, "showValue", part)?.unwrap_or(true);
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if local == "formula"
                    && current_rule
                        .as_ref()
                        .is_some_and(|rule| rule.formula_depth == Some(depth))
                {
                    let rule = current_rule
                        .as_mut()
                        .ok_or_else(|| format_error(part, "conditional rule state was lost"))?;
                    rule.formula_depth = None;
                    rule.formulas.push(std::mem::take(&mut rule.formula));
                } else if local == "cfRule"
                    && current_rule
                        .as_ref()
                        .is_some_and(|rule| rule.depth == depth)
                {
                    let rule = current_rule
                        .take()
                        .ok_or_else(|| format_error(part, "conditional rule state was lost"))?;
                    let kind = match rule.rule_type.as_str() {
                        "cellIs" => {
                            rule.dxf_id
                                .zip(rule.operator)
                                .and_then(|(dxf_id, operator)| {
                                    rule.formulas
                                        .iter()
                                        .map(|formula| parse_conditional_literal(formula))
                                        .collect::<Option<Vec<_>>>()
                                        .filter(|formulas| !formulas.is_empty())
                                        .map(|formulas| ConditionalRuleKind::CellIs {
                                            dxf_id,
                                            operator,
                                            formulas,
                                        })
                                })
                        }
                        "colorScale"
                            if (2..=3).contains(&rule.colors.len())
                                && rule.thresholds.len() == rule.colors.len() =>
                        {
                            Some(ConditionalRuleKind::ColorScale {
                                thresholds: rule.thresholds,
                                colors: rule.colors,
                            })
                        }
                        "dataBar" if rule.thresholds.len() >= 2 && !rule.colors.is_empty() => {
                            Some(ConditionalRuleKind::DataBar {
                                thresholds: rule.thresholds,
                                color: rule.colors[0],
                                show_value: rule.show_value,
                            })
                        }
                        "iconSet" if (3..=5).contains(&rule.thresholds.len()) => {
                            Some(ConditionalRuleKind::IconSet {
                                thresholds: rule.thresholds,
                                inclusive: rule.inclusive,
                                icon_set: rule.icon_set,
                                reverse: rule.reverse,
                                show_value: rule.show_value,
                            })
                        }
                        _ => None,
                    };
                    let kind = kind.or_else(|| {
                        crate::calculation::calculated_condition(&rule.rule_type)
                            .then_some(rule.dxf_id).flatten().map(|dxf_id|
                                ConditionalRuleKind::Calculated { dxf_id, matches: HashSet::new() })
                    });
                    if let Some(kind) = kind {
                        let block = current_block.as_mut().ok_or_else(|| {
                            format_error(part, "conditional formatting state was lost")
                        })?;
                        if block.rules.len() >= limits.max_document_objects {
                            return Err(object_limit_error(
                                part,
                                "conditional rule count exceeds the configured object limit",
                            ));
                        }
                        block.rules.push(ConditionalRule {
                            priority: rule.priority,
                            stop_if_true: rule.stop_if_true,
                            kind,
                        });
                    } else {
                        unsupported = true;
                    }
                } else if local == "conditionalFormatting"
                    && current_block
                        .as_ref()
                        .is_some_and(|block| block.depth == depth)
                {
                    let mut block = current_block.take().ok_or_else(|| {
                        format_error(part, "conditional formatting state was lost")
                    })?;
                    block.rules.sort_by_key(|rule| rule.priority);
                    blocks.push(ConditionalFormatting {
                        ranges: block.ranges,
                        rules: block.rules,
                    });
                }
            }
            XmlEvent::Text(text) => {
                if let Some(rule) = current_rule
                    .as_mut()
                    .filter(|rule| rule.formula_depth.is_some())
                {
                    rule.formula
                        .push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(rule) = current_rule
                    .as_mut()
                    .filter(|rule| rule.formula_depth.is_some())
                {
                    rule.formula.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok((blocks, unsupported))
}

fn parse_conditional_literal(formula: &str) -> Option<ConditionalLiteral> {
    let formula = formula.trim();
    if let Ok(value) = formula.parse::<f64>()
        && value.is_finite()
    {
        return Some(ConditionalLiteral::Number(value));
    }
    formula
        .strip_prefix('"')
        .and_then(|value| value.strip_suffix('"'))
        .map(|value| ConditionalLiteral::Text(value.replace("\"\"", "\"")))
}

fn parse_cell_operator(value: &str) -> Option<CellOperator> {
    match value {
        "lessThan" => Some(CellOperator::LessThan),
        "lessThanOrEqual" => Some(CellOperator::LessThanOrEqual),
        "equal" => Some(CellOperator::Equal),
        "notEqual" => Some(CellOperator::NotEqual),
        "greaterThanOrEqual" => Some(CellOperator::GreaterThanOrEqual),
        "greaterThan" => Some(CellOperator::GreaterThan),
        "between" => Some(CellOperator::Between),
        "notBetween" => Some(CellOperator::NotBetween),
        _ => None,
    }
}

fn parse_conditional_value(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<Option<ConditionalValue>, Diagnostic> {
    let value_type = required_attribute(attributes, "type", part)?;
    let raw_value = optional_attribute(attributes, "val", part)?;
    let number = || {
        raw_value
            .as_deref()
            .and_then(|value| value.parse::<f64>().ok())
            .filter(|value| value.is_finite())
    };
    Ok(match value_type.as_str() {
        "min" | "autoMin" => Some(ConditionalValue::Minimum),
        "max" | "autoMax" => Some(ConditionalValue::Maximum),
        "num" => number().map(ConditionalValue::Number),
        "percent" => number()
            .filter(|value| (0.0..=100.0).contains(value))
            .map(ConditionalValue::Percent),
        "percentile" => number()
            .filter(|value| (0.0..=100.0).contains(value))
            .map(ConditionalValue::Percentile),
        "formula" => number().map(ConditionalValue::Number),
        _ => None,
    })
}

fn parse_data_validations_xml(
    bytes: &[u8],
    limits: Limits,
    part: &str,
) -> Result<(usize, bool), Diagnostic> {
    if !contains_bytes(bytes, b"dataValidation") {
        return Ok((0, false));
    }
    #[derive(Debug)]
    struct ValidationState {
        depth: usize,
        supported: bool,
        formula_depth: Option<(usize, u8)>,
        formula: String,
        formula_count: u8,
    }

    let mut depth = 0_usize;
    let mut current = None::<ValidationState>;
    let mut validation_count = 0_usize;
    let mut unsupported = false;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                // Standard validations are direct children of worksheet/dataValidations.
                if local == "dataValidation" && depth == 2 && current.is_none() {
                    let kind = optional_attribute(&attributes, "type", part)?
                        .as_deref()
                        .and_then(parse_data_validation_kind);
                    let ranges = required_attribute(&attributes, "sqref", part)?
                        .split_ascii_whitespace()
                        .map(|range| parse_dimension(range, part))
                        .collect::<Result<Vec<_>, _>>()?;
                    if ranges.is_empty() {
                        return Err(format_error(part, "data validation has no target ranges"));
                    }
                    let _ = parse_bool_attribute(&attributes, "allowBlank", part)?;
                    for name in ["promptTitle", "prompt", "errorTitle", "error"] {
                        let _ = optional_attribute(&attributes, name, part)?;
                    }
                    for name in ["showDropDown", "showInputMessage", "showErrorMessage"] {
                        let _ = parse_bool_attribute(&attributes, name, part)?;
                    }
                    if kind.is_none() {
                        unsupported |= optional_attribute(&attributes, "type", part)?.is_some();
                    }
                    let state = ValidationState {
                        depth,
                        supported: kind.is_some(),
                        formula_depth: None,
                        formula: String::new(),
                        formula_count: 0,
                    };
                    if empty {
                        if state.supported {
                            if validation_count >= limits.max_document_objects {
                                return Err(object_limit_error(
                                    part,
                                    "data-validation count exceeds the configured object limit",
                                ));
                            }
                            validation_count += 1;
                        }
                    } else {
                        current = Some(state);
                    }
                } else if local == "dataValidation" && current.is_none() {
                    unsupported = true;
                } else if matches!(local, "formula1" | "formula2")
                    && let Some(validation) = current.as_mut()
                {
                    if validation.formula_depth.is_some() {
                        return Err(format_error(
                            part,
                            "nested data-validation formulas are invalid",
                        ));
                    }
                    let index = if local == "formula1" { 1 } else { 2 };
                    validation.formula.clear();
                    validation.formula_depth = (!empty).then_some((depth, index));
                    validation.formula_count = validation.formula_count.max(index);
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if matches!(local, "formula1" | "formula2")
                    && current
                        .as_ref()
                        .and_then(|validation| validation.formula_depth)
                        .is_some_and(|(start, _)| start == depth)
                {
                    current
                        .as_mut()
                        .expect("validation state is present")
                        .formula_depth = None;
                } else if local == "dataValidation"
                    && current
                        .as_ref()
                        .is_some_and(|validation| validation.depth == depth)
                {
                    let state = current
                        .take()
                        .ok_or_else(|| format_error(part, "data-validation state was lost"))?;
                    if state.supported {
                        if validation_count >= limits.max_document_objects {
                            return Err(object_limit_error(
                                part,
                                "data-validation count exceeds the configured object limit",
                            ));
                        }
                        validation_count += 1;
                    }
                }
            }
            XmlEvent::Text(text) => {
                if let Some(validation) = current
                    .as_mut()
                    .filter(|validation| validation.formula_depth.is_some())
                {
                    validation
                        .formula
                        .push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(validation) = current
                    .as_mut()
                    .filter(|validation| validation.formula_depth.is_some())
                {
                    validation.formula.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok((validation_count, unsupported))
}

fn parse_data_validation_kind(value: &str) -> Option<DataValidationKind> {
    match value {
        "list" => Some(DataValidationKind::List),
        "whole" => Some(DataValidationKind::Whole),
        "decimal" => Some(DataValidationKind::Decimal),
        "date" => Some(DataValidationKind::Date),
        "time" => Some(DataValidationKind::Time),
        "textLength" => Some(DataValidationKind::TextLength),
        "custom" => Some(DataValidationKind::Custom),
        _ => None,
    }
}

fn parse_sparkline_groups_xml(
    bytes: &[u8],
    limits: Limits,
    theme_colors: &ThemeColors,
    part: &str,
) -> Result<(Vec<SparklineGroup>, bool), Diagnostic> {
    if !contains_bytes(bytes, b"sparkline") {
        return Ok((Vec::new(), false));
    }
    #[derive(Debug)]
    struct SparklineState {
        depth: usize,
        formula_depth: Option<usize>,
        target_depth: Option<usize>,
        formula: String,
        target: String,
    }

    #[derive(Debug)]
    struct GroupState {
        depth: usize,
        group: Option<SparklineGroup>,
        sparkline: Option<SparklineState>,
    }

    let mut depth = 0_usize;
    let mut current = None::<GroupState>;
    let mut groups = Vec::new();
    let mut unsupported = false;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "sparklineGroup" && current.is_none() {
                    let kind = match optional_attribute(&attributes, "type", part)?.as_deref() {
                        None | Some("line") => Some(SparklineKind::Line),
                        Some("column") => Some(SparklineKind::Column),
                        Some("stacked" | "winLoss" | "winloss") => Some(SparklineKind::WinLoss),
                        Some(_) => {
                            unsupported = true;
                            None
                        }
                    };
                    let min_axis = parse_sparkline_axis_mode(
                        optional_attribute(&attributes, "minAxisType", part)?.as_deref(),
                    );
                    let max_axis = parse_sparkline_axis_mode(
                        optional_attribute(&attributes, "maxAxisType", part)?.as_deref(),
                    );
                    unsupported |= min_axis.is_none() || max_axis.is_none();
                    let manual_min = parse_optional_finite_f64(
                        &attributes,
                        "manualMin",
                        part,
                        "sparkline manual minimum",
                    )?;
                    let manual_max = parse_optional_finite_f64(
                        &attributes,
                        "manualMax",
                        part,
                        "sparkline manual maximum",
                    )?;
                    let display_axis =
                        parse_bool_attribute(&attributes, "displayXAxis", part)?.unwrap_or(false);
                    let group = kind.map(|kind| SparklineGroup {
                        kind,
                        min_axis: min_axis.unwrap_or(SparklineAxisMode::Individual),
                        max_axis: max_axis.unwrap_or(SparklineAxisMode::Individual),
                        manual_min,
                        manual_max,
                        series_color: 0x4472_c4ff,
                        negative_color: 0xc000_00ff,
                        axis_color: 0x0000_00ff,
                        display_axis,
                        sparklines: Vec::new(),
                    });
                    let state = GroupState {
                        depth,
                        group,
                        sparkline: None,
                    };
                    if empty {
                        unsupported = true;
                    } else {
                        current = Some(state);
                    }
                } else if matches!(local, "colorSeries" | "colorNegative" | "colorAxis")
                    && let Some(group) = current.as_mut().and_then(|state| state.group.as_mut())
                    && let Some(color) = parse_color(&attributes, theme_colors, part)?
                {
                    match local {
                        "colorSeries" => group.series_color = color,
                        "colorNegative" => group.negative_color = color,
                        "colorAxis" => group.axis_color = color,
                        _ => unreachable!(),
                    }
                } else if local == "sparkline"
                    && let Some(group) = current.as_mut()
                {
                    if group.sparkline.is_some() {
                        return Err(format_error(part, "nested XLSX sparklines are invalid"));
                    }
                    if empty {
                        unsupported = true;
                    } else {
                        group.sparkline = Some(SparklineState {
                            depth,
                            formula_depth: None,
                            target_depth: None,
                            formula: String::new(),
                            target: String::new(),
                        });
                    }
                } else if matches!(local, "f" | "sqref")
                    && let Some(sparkline) =
                        current.as_mut().and_then(|group| group.sparkline.as_mut())
                {
                    let field = if local == "f" {
                        &mut sparkline.formula_depth
                    } else {
                        &mut sparkline.target_depth
                    };
                    if field.is_some() {
                        return Err(format_error(
                            part,
                            "nested XLSX sparkline fields are invalid",
                        ));
                    }
                    *field = (!empty).then_some(depth);
                    if empty {
                        unsupported = true;
                    }
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if matches!(local, "f" | "sqref")
                    && let Some(sparkline) =
                        current.as_mut().and_then(|group| group.sparkline.as_mut())
                {
                    let field = if local == "f" {
                        &mut sparkline.formula_depth
                    } else {
                        &mut sparkline.target_depth
                    };
                    if *field == Some(depth) {
                        *field = None;
                    }
                } else if local == "sparkline"
                    && current
                        .as_ref()
                        .and_then(|group| group.sparkline.as_ref())
                        .is_some_and(|sparkline| sparkline.depth == depth)
                {
                    let group = current.as_mut().expect("sparkline group state is present");
                    let sparkline = group
                        .sparkline
                        .take()
                        .ok_or_else(|| format_error(part, "sparkline state was lost"))?;
                    let target = parse_sparkline_target(&sparkline.target);
                    if let (Some(group), Some(target)) = (group.group.as_mut(), target)
                        && !sparkline.formula.trim().is_empty()
                    {
                        if group.sparklines.len() >= limits.max_document_objects {
                            return Err(object_limit_error(
                                part,
                                "sparkline count exceeds the configured object limit",
                            ));
                        }
                        group.sparklines.push(Sparkline {
                            source_formula: sparkline.formula,
                            target,
                        });
                    } else {
                        unsupported = true;
                    }
                } else if local == "sparklineGroup"
                    && current.as_ref().is_some_and(|group| group.depth == depth)
                {
                    let group = current
                        .take()
                        .ok_or_else(|| format_error(part, "sparkline-group state was lost"))?;
                    if let Some(group) = group.group {
                        if group.sparklines.is_empty() {
                            unsupported = true;
                        } else {
                            if groups.len() >= limits.max_document_objects {
                                return Err(object_limit_error(
                                    part,
                                    "sparkline-group count exceeds the configured object limit",
                                ));
                            }
                            groups.push(group);
                        }
                    }
                }
            }
            XmlEvent::Text(text) => {
                if let Some(sparkline) = current.as_mut().and_then(|group| group.sparkline.as_mut())
                {
                    let decoded = decode_xml_text(text).map_err(|error| with_part(error, part))?;
                    if sparkline.formula_depth.is_some() {
                        sparkline.formula.push_str(&decoded);
                    } else if sparkline.target_depth.is_some() {
                        sparkline.target.push_str(&decoded);
                    }
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(sparkline) = current.as_mut().and_then(|group| group.sparkline.as_mut())
                {
                    if sparkline.formula_depth.is_some() {
                        sparkline.formula.push_str(text);
                    } else if sparkline.target_depth.is_some() {
                        sparkline.target.push_str(text);
                    }
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok((groups, unsupported))
}

fn parse_sparkline_axis_mode(value: Option<&str>) -> Option<SparklineAxisMode> {
    match value {
        None | Some("individual") => Some(SparklineAxisMode::Individual),
        Some("group") => Some(SparklineAxisMode::Group),
        Some("custom") => Some(SparklineAxisMode::Custom),
        Some(_) => None,
    }
}

fn parse_optional_finite_f64(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
    label: &str,
) -> Result<Option<f64>, Diagnostic> {
    optional_attribute(attributes, name, part)?
        .map(|value| {
            value
                .parse::<f64>()
                .ok()
                .filter(|value| value.is_finite())
                .ok_or_else(|| format_error(part, format!("{label} must be finite")))
        })
        .transpose()
}

fn parse_sparkline_target(value: &str) -> Option<CellAddress> {
    let normalized = value.trim().replace('$', "").to_ascii_uppercase();
    if normalized.contains([':', '!', ' ', '\t', '\r', '\n']) {
        return None;
    }
    parse_a1(&normalized)
}

fn parse_worksheet_metadata(
    package: &Package<'_>,
    part: &str,
    print_area: Option<&str>,
    theme_colors: &ThemeColors,
    maximum_digit_width: f32,
    shared_strings: &[SharedString],
    styles: &Styles,
) -> Result<WorksheetMetadata, Diagnostic> {
    let bytes = package.required_part(part)?;
    let fast_metadata = fast_sheet_data_metadata(&bytes, shared_strings, styles);
    let metadata_bytes = fast_metadata.as_ref().map_or(bytes.as_ref(), |metadata| {
        metadata.structural_xml.as_slice()
    });
    let mut print_settings = PrintSettingsParser::new(print_area.map(str::to_owned));
    let mut depth = 0_usize;
    let mut dimension = fast_metadata
        .as_ref()
        .and_then(|metadata| metadata.dimension);
    let mut base_column_width = None;
    let mut default_column_width = None;
    let mut default_row_height = ROW_HEIGHT;
    let mut column_spans = Vec::new();
    let mut row_sizes = fast_metadata
        .as_ref()
        .map_or_else(Vec::new, |metadata| metadata.row_sizes.clone());
    let mut cols_depth = None::<usize>;
    let mut frozen_rows = 0_u32;
    let mut frozen_columns = 0_u32;
    let mut frozen_pane_seen = false;
    let mut show_grid_lines = true;
    let mut tab_color = None;
    let mut diagnostics = Vec::new();
    let mut formula_cell = None::<(usize, bool, bool)>;
    let mut drawing_relationship_ids = Vec::new();
    let mut requires_formula_evaluation = fast_metadata
        .as_ref()
        .is_some_and(|metadata| metadata.requires_formula_evaluation);
    parse_xml(metadata_bytes, package.limits(), |event| {
        print_settings.consume(&event, part, &mut diagnostics)?;
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "tabColor" && depth == 2 {
                    tab_color =
                        worksheet_tab_color(&attributes, theme_colors, part, &mut diagnostics);
                }
                if local == "drawing" {
                    let id = relationship_id(&attributes, part)?;
                    if !drawing_relationship_ids.contains(&id) {
                        drawing_relationship_ids.push(id);
                    }
                }
                if local == "c" && formula_cell.is_none() {
                    formula_cell = Some((depth, false, false));
                } else if let Some((_, formula, value)) = formula_cell.as_mut() {
                    if local == "f" {
                        *formula = true;
                    } else if local == "v" {
                        *value = true;
                    }
                }
                if local == "cols" && cols_depth.is_none() {
                    cols_depth = (!empty).then_some(depth);
                }
                if local == "pane"
                    && !frozen_pane_seen
                    && matches!(
                        optional_attribute(&attributes, "state", part)?.as_deref(),
                        Some("frozen" | "frozenSplit")
                    )
                {
                    frozen_columns = parse_frozen_split(
                        optional_attribute(&attributes, "xSplit", part)?.as_deref(),
                        MAX_COLUMNS,
                        part,
                        "frozen column count",
                    )?;
                    frozen_rows = parse_frozen_split(
                        optional_attribute(&attributes, "ySplit", part)?.as_deref(),
                        MAX_ROWS,
                        part,
                        "frozen row count",
                    )?;
                    frozen_pane_seen = true;
                } else if local == "sheetView" {
                    show_grid_lines =
                        parse_bool_attribute(&attributes, "showGridLines", part)?.unwrap_or(true);
                } else if local == "dimension" && dimension.is_none() {
                    let value = required_attribute(&attributes, "ref", part)?;
                    dimension = parse_dimension(&value, part).ok();
                } else if local == "sheetFormatPr" {
                    if let Some(width) = optional_attribute(&attributes, "baseColWidth", part)? {
                        let width = parse_finite_f32(&width, part, "base column width")?;
                        if width > 255.0 {
                            return Err(dimension_error(
                                part,
                                "base column width exceeds the XLSX limit",
                            ));
                        }
                        base_column_width = Some(width);
                    }
                    if let Some(width) = optional_attribute(&attributes, "defaultColWidth", part)? {
                        default_column_width = Some(excel_default_column_width(
                            parse_finite_f32(&width, part, "default column width")?,
                            maximum_digit_width,
                            part,
                        )?);
                    }
                    if let Some(height) = excel_default_row_height(&attributes, part)? {
                        default_row_height = height;
                    }
                } else if local == "col"
                    && cols_depth.is_some_and(|container| depth == container + 1)
                {
                    let start = parse_one_based_index(
                        &required_attribute(&attributes, "min", part)?,
                        MAX_COLUMNS,
                        part,
                        "column minimum",
                    )?;
                    let end = parse_one_based_index(
                        &required_attribute(&attributes, "max", part)?,
                        MAX_COLUMNS,
                        part,
                        "column maximum",
                    )?;
                    if start > end {
                        return Err(format_error(part, "column range is reversed"));
                    }
                    let hidden =
                        parse_bool_attribute(&attributes, "hidden", part)?.unwrap_or(false);
                    let width = if hidden {
                        Some(0.0)
                    } else {
                        optional_attribute(&attributes, "width", part)?
                            .map(|width| {
                                excel_column_width(
                                    parse_finite_f32(&width, part, "column width")?,
                                    maximum_digit_width,
                                    part,
                                )
                            })
                            .transpose()?
                    };
                    column_spans.push(ColumnSpan {
                        start,
                        end,
                        width,
                        style_index: None,
                    });
                } else if local == "row" {
                    let Some(index) = optional_attribute(&attributes, "r", part)? else {
                        if !empty {
                            depth = depth.saturating_add(1);
                        }
                        return Ok(());
                    };
                    let index = parse_one_based_index(&index, MAX_ROWS, part, "row index")?;
                    let hidden =
                        parse_bool_attribute(&attributes, "hidden", part)?.unwrap_or(false);
                    let height = if hidden {
                        Some(0.0)
                    } else {
                        optional_attribute(&attributes, "ht", part)?
                            .map(|height| {
                                points_to_pixels(
                                    parse_finite_f32(&height, part, "row height")?,
                                    part,
                                    "row height",
                                )
                            })
                            .transpose()?
                    };
                    if let Some(height) = height {
                        row_sizes.push((index, height));
                    }
                } else if matches!(local, "c" | "mergeCell") {
                    let reference = if local == "c" {
                        optional_attribute(&attributes, "r", part)?
                    } else {
                        Some(required_attribute(&attributes, "ref", part)?)
                    };
                    if let Some(reference) = reference {
                        let candidate = if local == "c" {
                            parse_a1(&reference).map(|address| Dimension {
                                start_column: address.column,
                                start_row: address.row,
                                end_column: address.column,
                                end_row: address.row,
                            })
                        } else {
                            parse_dimension(&reference, part).ok()
                        };
                        if let Some(candidate) = candidate {
                            dimension = Some(match dimension {
                                Some(current) => Dimension {
                                    start_column: current.start_column.min(candidate.start_column),
                                    start_row: current.start_row.min(candidate.start_row),
                                    end_column: current.end_column.max(candidate.end_column),
                                    end_row: current.end_row.max(candidate.end_row),
                                },
                                None => candidate,
                            });
                        }
                    }
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if local == "c"
                    && formula_cell.is_some_and(|(cell_depth, _, _)| cell_depth == depth)
                {
                    let (_, formula, value) = formula_cell.take().expect("cell scan is present");
                    requires_formula_evaluation |= formula && !value;
                }
                if local == "cols" && cols_depth == Some(depth) {
                    cols_depth = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    if !requires_formula_evaluation {
        requires_formula_evaluation = worksheet_requires_formula_evaluation(package, part)?;
    }
    let dimension = dimension.unwrap_or(Dimension {
        start_column: 0,
        start_row: 0,
        end_column: 0,
        end_row: 0,
    });
    let mut rows = dimension.end_row.saturating_add(1).max(frozen_rows).max(1);
    let mut columns = dimension
        .end_column
        .saturating_add(1)
        .max(frozen_columns)
        .max(1);
    let default_column_width = default_column_width.unwrap_or_else(|| {
        base_column_width
            .map(|width| (width * maximum_digit_width + 5.0).floor())
            .unwrap_or_else(|| default_excel_column_width(maximum_digit_width))
    });
    let column_metrics = column_metrics(default_column_width, columns, &column_spans, part)?;
    let row_metrics = AxisMetrics::new(default_row_height, row_sizes);
    if let Some(fast_metadata) = &fast_metadata {
        let mut trailing_text_extent = 0.0_f32;
        for text in &fast_metadata.trailing_text {
            if text.column.saturating_add(1) < columns {
                continue;
            }
            let left = column_metrics.offset(text.column);
            let width = column_metrics.size(text.column);
            if width == 0.0 || row_metrics.size(text.row) == 0.0 {
                continue;
            }
            let right = match text.align {
                TextAlign::Start => left + 2.0 + text.width,
                TextAlign::Center => left + (width + text.width) / 2.0,
                _ => left + width,
            };
            trailing_text_extent = trailing_text_extent.max(right);
        }
        columns = columns.max(column_metrics.count_for_extent(trailing_text_extent, MAX_COLUMNS));
    }
    let (drawing_width, drawing_height) = worksheet_drawing_extent(
        package,
        part,
        &drawing_relationship_ids,
        &column_metrics,
        &row_metrics,
        theme_colors,
        &styles.theme_line_styles,
    )?;
    columns = columns.max(column_metrics.count_for_extent(drawing_width, MAX_COLUMNS));
    rows = rows.max(row_metrics.count_for_extent(drawing_height, MAX_ROWS));
    Ok(WorksheetMetadata {
        width: column_metrics.offset(columns).max(drawing_width).max(1.0),
        height: row_metrics.offset(rows).max(drawing_height).max(1.0),
        rows,
        columns,
        frozen_rows,
        frozen_columns,
        frozen_width: column_metrics.offset(frozen_columns),
        frozen_height: row_metrics.offset(frozen_rows),
        row_axis: row_metrics.descriptor(rows),
        column_axis: column_metrics.descriptor(columns),
        show_grid_lines,
        tab_color,
        diagnostics,
        print_settings: print_settings.finish(),
        requires_formula_evaluation,
    })
}

fn chartsheet_drawing_relationship_ids(
    package: &Package<'_>,
    part: &str,
    theme_colors: &ThemeColors,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<(Vec<String>, Option<u32>), Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut relationship_ids = Vec::new();
    let mut tab_color = None;
    parse_xml(&bytes, package.limits(), |event| {
        if let XmlEvent::StartElement {
            name, attributes, ..
        } = event
        {
            match local_name(name) {
                "drawing" => {
                    let id = relationship_id(&attributes, part)?;
                    if !relationship_ids.contains(&id) {
                        relationship_ids.push(id);
                    }
                }
                "tabColor" => {
                    tab_color = worksheet_tab_color(&attributes, theme_colors, part, diagnostics)
                }
                _ => {}
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok((relationship_ids, tab_color))
}

// Use the same RGB/indexed/theme/tint semantics as cell colors. An automatic
// sheet tab is uncolored; a bad optional color must not discard the worksheet.
fn worksheet_tab_color(
    attributes: &[XmlAttribute<'_>],
    theme_colors: &ThemeColors,
    part: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Option<u32> {
    let result = parse_bool_attribute(attributes, "auto", part).and_then(|automatic| {
        if automatic == Some(true) {
            Ok(None)
        } else {
            parse_color(attributes, theme_colors, part)
        }
    });
    match result {
        Ok(color) => color,
        Err(_) => {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::FormatInvalid,
                    Phase::Parse,
                    Fidelity::Omitted,
                    "invalid XLSX worksheet tab color was ignored",
                )
                .in_part(part),
            );
            None
        }
    }
}

fn worksheet_drawing_extent(
    package: &Package<'_>,
    worksheet_part: &str,
    relationship_ids: &[String],
    column_metrics: &AxisMetrics,
    row_metrics: &AxisMetrics,
    theme_colors: &ThemeColors,
    theme_line_styles: &DrawingMlThemeLineStyles,
) -> Result<(f32, f32), Diagnostic> {
    if relationship_ids.is_empty() {
        return Ok((0.0, 0.0));
    }
    let worksheet_relationships = package.relationships(Some(worksheet_part))?;
    let worksheet_relationships = worksheet_relationships
        .iter()
        .map(|relationship| (relationship.id.as_str(), relationship))
        .collect::<HashMap<_, _>>();
    let mut extent = (0.0_f32, 0.0_f32);
    for relationship_id in relationship_ids {
        let Some(relationship) = worksheet_relationships.get(relationship_id.as_str()) else {
            continue;
        };
        if relationship.external || !relationship.type_uri.ends_with("/drawing") {
            continue;
        }
        for drawing in parse_worksheet_drawing(
            package,
            &relationship.target,
            column_metrics,
            row_metrics,
            theme_colors,
            theme_line_styles,
        )? {
            let drawing_extent = worksheet_drawing_visual_extent(&drawing);
            extent.0 = extent.0.max(drawing_extent.0);
            extent.1 = extent.1.max(drawing_extent.1);
        }
    }
    Ok(extent)
}

fn worksheet_drawing_visual_extent(drawing: &WorksheetDrawing) -> (f32, f32) {
    let bounds = drawing.bounds;
    let (outer_shadow, shadow, glow, reflection, soft_edge, three_d) = match &drawing.payload {
        DrawingPayload::Image(image) => (
            image.effects.outer_shadow.as_ref(),
            None,
            image.effects.glow.as_ref(),
            image.effects.reflection.as_ref(),
            image.effects.soft_edge,
            image.effects.three_d.as_ref(),
        ),
        DrawingPayload::Shape(shape) => (
            shape.outer_shadow.as_ref(),
            shape.shadow.as_ref(),
            None,
            None,
            None,
            shape.three_d.as_ref(),
        ),
        DrawingPayload::Chart(_) | DrawingPayload::Diagram(_) => {
            (None, None, None, None, None, None)
        }
    };
    let mut right = bounds.x + bounds.width;
    let mut bottom = bounds.y + bounds.height;
    let outset = glow
        .map_or(0.0, |glow| glow.radius)
        .max(soft_edge.unwrap_or(0.0));
    right += outset;
    bottom += outset;
    if let Some(style) = three_d {
        for (x, y) in [
            (bounds.x, bounds.y),
            (bounds.x + bounds.width, bounds.y),
            (bounds.x, bounds.y + bounds.height),
            (bounds.x + bounds.width, bounds.y + bounds.height),
        ] {
            if let Some((x, y)) = drawingml_camera_point(style, bounds, (x, y)) {
                right = right.max(x + outset);
                bottom = bottom.max(y + outset);
            }
        }
    }
    if let Some(shadow) = shadow {
        right = right.max(bounds.x + bounds.width + shadow.blur + shadow.offset_x.max(0.0));
        bottom = bottom.max(bounds.y + bounds.height + shadow.blur + shadow.offset_y.max(0.0));
    }
    if let Some(effect) = outer_shadow {
        let horizontal = effect.alignment % 3;
        let vertical = effect.alignment / 3;
        let anchor_x = bounds.x + bounds.width * f32::from(horizontal) / 2.0;
        let anchor_y = bounds.y + bounds.height * f32::from(vertical) / 2.0;
        let skew_x = effect.skew_x.to_radians().tan();
        let skew_y = effect.skew_y.to_radians().tan();
        for (x, y) in [
            (bounds.x, bounds.y),
            (bounds.x + bounds.width, bounds.y),
            (bounds.x, bounds.y + bounds.height),
            (bounds.x + bounds.width, bounds.y + bounds.height),
        ] {
            let local_x = x - anchor_x;
            let local_y = y - anchor_y;
            right = right.max(
                anchor_x
                    + effect.shadow.offset_x
                    + effect.scale_x * local_x
                    + skew_x * local_y
                    + effect.shadow.blur,
            );
            bottom = bottom.max(
                anchor_y
                    + effect.shadow.offset_y
                    + skew_y * local_x
                    + effect.scale_y * local_y
                    + effect.shadow.blur,
            );
        }
    }
    if let Some(reflection) = reflection {
        right =
            right.max(bounds.x + bounds.width * (1.0 + reflection.scale_x) / 2.0 + reflection.blur);
        bottom = bottom.max(
            bounds.y
                + bounds.height
                + reflection.distance
                + bounds.height * reflection.scale_y
                + reflection.blur,
        );
    }
    (right, bottom)
}

#[allow(clippy::too_many_arguments)]
fn parse_worksheet(
    package: &Package<'_>,
    part: &str,
    sheet_name: &str,
    print_area: Option<&str>,
    source_sheet_index: u32,
    unit_index: u32,
    shared_strings: &[SharedString],
    styles: &Styles,
    date_1904: bool,
    formula_evaluator: Option<&CalculationResults>,
    content_types: &ContentTypes,
    chart_named_data: &HashMap<String, super::drawingml::ChartSourceData>,
    state: &mut XlsxParseState,
) -> Result<Unit, Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut print_settings = PrintSettingsParser::new(print_area.map(str::to_owned));
    let mut depth = 0_usize;
    let mut dimension = None;
    let mut current_cell = None;
    let mut cells = Vec::new();
    let mut addresses = HashSet::new();
    let maximum_digit_width = excel_maximum_digit_width(styles);
    let mut base_column_width = None;
    let mut default_column_width = None;
    let mut default_row_height = ROW_HEIGHT;
    let mut column_spans = Vec::new();
    let mut row_sizes = HashMap::new();
    let mut row_styles = HashMap::<u32, usize>::new();
    let mut next_row = 0_u32;
    let mut next_cell_column = 0_u32;
    let mut cols_depth = None::<usize>;
    let mut sheet_data_depth = None::<usize>;
    let mut current_row = None::<RowState>;
    let mut merges = Vec::new();
    let mut background_relationship_id = None;
    let mut unsupported_pivot_reported = false;
    let mut unsupported_slicer_reported = false;
    let mut drawing_relationship_ids = Vec::new();
    let mut legacy_drawing_relationship_ids = Vec::new();
    let mut table_relationship_ids = Vec::new();
    let mut frozen_rows = 0_u32;
    let mut frozen_columns = 0_u32;
    let mut frozen_pane_seen = false;
    let mut show_grid_lines = true;
    let mut tab_color = None;

    parse_xml(&bytes, package.limits(), |event| {
        print_settings.consume(&event, part, &mut state.diagnostics)?;
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "tabColor" && depth == 2 {
                    tab_color = worksheet_tab_color(&attributes, &styles.theme_colors, part, &mut state.diagnostics);
                }
                if local == "cols" && cols_depth.is_none() {
                    cols_depth = (!empty).then_some(depth);
                } else if local == "sheetData" && sheet_data_depth.is_none() {
                    sheet_data_depth = (!empty).then_some(depth);
                }
                if matches!(local, "pivotTablePart" | "pivotTableDefinition")
                    && !unsupported_pivot_reported
                {
                    state.diagnostics.push(
                        Diagnostic::warning(
                            DiagnosticCode::UnsupportedFeature,
                            Phase::Render,
                            Fidelity::Omitted,
                            "interactive XLSX pivot-table features are not rendered",
                        )
                        .in_part(part),
                    );
                    unsupported_pivot_reported = true;
                } else if matches!(
                    local,
                    "slicer" | "slicerList" | "slicerRef" | "timelineRef" | "timelineRefs"
                ) && !unsupported_slicer_reported
                {
                    state.diagnostics.push(
                        Diagnostic::warning(
                            DiagnosticCode::UnsupportedFeature,
                            Phase::Render,
                            Fidelity::Omitted,
                            "XLSX slicers and timelines are not rendered",
                        )
                        .in_part(part),
                    );
                    unsupported_slicer_reported = true;
                }
                if local == "drawing" {
                    let id = relationship_id(&attributes, part)?;
                    if !drawing_relationship_ids.contains(&id) {
                        drawing_relationship_ids.push(id);
                    }
                } else if local == "legacyDrawing" {
                    let id = relationship_id(&attributes, part)?;
                    if !legacy_drawing_relationship_ids.contains(&id) {
                        legacy_drawing_relationship_ids.push(id);
                    }
                } else if matches!(local, "tablePart" | "pivotTablePart") {
                    let id = relationship_id(&attributes, part)?;
                    if !table_relationship_ids.contains(&id) {
                        table_relationship_ids.push(id);
                    }
                } else if local == "picture" {
                    background_relationship_id = Some(optional_attribute(&attributes, "id", part)?.unwrap_or_default());
                }
                if local == "pane"
                    && !frozen_pane_seen
                    && matches!(
                        optional_attribute(&attributes, "state", part)?.as_deref(),
                        Some("frozen" | "frozenSplit")
                    )
                {
                    frozen_columns = parse_frozen_split(
                        optional_attribute(&attributes, "xSplit", part)?.as_deref(),
                        MAX_COLUMNS,
                        part,
                        "frozen column count",
                    )?;
                    frozen_rows = parse_frozen_split(
                        optional_attribute(&attributes, "ySplit", part)?.as_deref(),
                        MAX_ROWS,
                        part,
                        "frozen row count",
                    )?;
                    frozen_pane_seen = true;
                }
                if local == "sheetView" {
                    show_grid_lines =
                        parse_bool_attribute(&attributes, "showGridLines", part)?.unwrap_or(true);
                }
                if local == "dimension" {
                    if dimension.is_some() {
                        return Err(format_error(part, "worksheet has multiple dimensions"));
                    }
                    let value = required_attribute(&attributes, "ref", part)?;
                    match parse_dimension(&value, part) {
                        Ok(parsed) => dimension = Some(parsed),
                        Err(_) => state.diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::UnsupportedFeature,
                                Phase::Parse,
                                Fidelity::Approximate,
                                "invalid XLSX worksheet dimension was ignored and derived from content",
                            )
                            .in_part(part),
                        ),
                    }
                } else if local == "sheetFormatPr" {
                    if let Some(width) = optional_attribute(&attributes, "baseColWidth", part)? {
                        let width = parse_finite_f32(&width, part, "base column width")?;
                        if width > 255.0 {
                            return Err(dimension_error(
                                part,
                                "base column width exceeds the XLSX limit",
                            ));
                        }
                        base_column_width = Some(width);
                    }
                    if let Some(width) = optional_attribute(&attributes, "defaultColWidth", part)? {
                        default_column_width = Some(excel_default_column_width(
                            parse_finite_f32(&width, part, "default column width")?,
                            maximum_digit_width,
                            part,
                        )?);
                    }
                    if let Some(height) = excel_default_row_height(&attributes, part)? {
                        default_row_height = height;
                    }
                } else if local == "col"
                    && cols_depth.is_some_and(|container_depth| depth == container_depth + 1)
                {
                    let start = parse_one_based_index(
                        &required_attribute(&attributes, "min", part)?,
                        MAX_COLUMNS,
                        part,
                        "column minimum",
                    )?;
                    let end = parse_one_based_index(
                        &required_attribute(&attributes, "max", part)?,
                        MAX_COLUMNS,
                        part,
                        "column maximum",
                    )?;
                    if start > end {
                        return Err(format_error(part, "column range is reversed"));
                    }
                    let hidden =
                        parse_bool_attribute(&attributes, "hidden", part)?.unwrap_or(false);
                    let width = if hidden {
                        Some(0.0)
                    } else {
                        optional_attribute(&attributes, "width", part)?
                            .map(|width| {
                                excel_column_width(
                                    parse_finite_f32(&width, part, "column width")?,
                                    maximum_digit_width,
                                    part,
                                )
                            })
                            .transpose()?
                    };
                    let style_index = optional_attribute(&attributes, "style", part)?
                        .map(|index| parse_zero_based_index(&index, part, "column style index"))
                        .transpose()?;
                    if style_index.is_some_and(|index| index >= styles.cells.len()) {
                        return Err(format_error(part, "column references a missing style"));
                    }
                    column_spans.push(ColumnSpan {
                        start,
                        end,
                        width,
                        style_index,
                    });
                } else if local == "row"
                    && sheet_data_depth
                        .is_some_and(|container_depth| depth == container_depth + 1)
                {
                    let row_index = match optional_attribute(&attributes, "r", part)? {
                        Some(index) => parse_one_based_index(&index, MAX_ROWS, part, "row index")?,
                        None => next_row,
                    };
                    next_row = row_index.saturating_add(1);
                    next_cell_column = 0;
                    let style_index = optional_attribute(&attributes, "s", part)?
                        .map(|index| parse_zero_based_index(&index, part, "row style index"))
                        .transpose()?;
                    if style_index.is_some_and(|index| index >= styles.cells.len()) {
                        return Err(format_error(part, "row references a missing style"));
                    }
                    if let Some(style_index) = style_index {
                        row_styles.insert(row_index, style_index);
                    }
                    current_row = (!empty).then_some(RowState {
                        depth,
                        index: row_index,
                        style_index,
                    });
                    let hidden =
                        parse_bool_attribute(&attributes, "hidden", part)?.unwrap_or(false);
                    let height = if hidden {
                        Some(0.0)
                    } else {
                        optional_attribute(&attributes, "ht", part)?
                            .map(|height| {
                                points_to_pixels(
                                    parse_finite_f32(&height, part, "row height")?,
                                    part,
                                    "row height",
                                )
                            })
                            .transpose()?
                    };
                    if let Some(height) = height
                        && row_sizes
                            .insert(row_index, height)
                            .is_some()
                    {
                        return Err(format_error(part, "worksheet contains a duplicate row"));
                    }
                } else if local == "mergeCell" {
                    if merges.len() >= package.limits().max_document_objects {
                        return Err(object_limit_error(
                            part,
                            "merged-cell count exceeds the configured object limit",
                        ));
                    }
                    merges.push(parse_dimension(
                        &required_attribute(&attributes, "ref", part)?,
                        part,
                    )?);
                } else if local == "c"
                    && (current_row.is_some_and(|row| depth == row.depth + 1)
                        || sheet_data_depth
                            .is_some_and(|container_depth| depth == container_depth + 1))
                {
                    if current_cell.is_some() {
                        return Err(format_error(part, "nested worksheet cells are invalid"));
                    }
                    let pending_objects = state
                        .objects
                        .len()
                        .checked_add(cells.len())
                        .ok_or_else(|| object_limit_error(part, "cell object count overflow"))?;
                    if pending_objects >= package.limits().max_document_objects {
                        return Err(Diagnostic::fatal(
                            DiagnosticCode::ObjectLimit,
                            Phase::Parse,
                            None,
                            "document exceeds the configured object limit",
                        )
                        .in_part(part));
                    }
                    let address = match optional_attribute(&attributes, "r", part)? {
                        Some(raw_address) => match parse_a1(&raw_address) {
                            Some(address) => address,
                            None => match parse_numeric_cell_address(&raw_address) {
                                Some(address) => {
                                    if !state.diagnostics.iter().any(|diagnostic| {
                                        diagnostic.message.contains("numeric cell address")
                                            && diagnostic.location.part.as_deref() == Some(part)
                                    }) {
                                        state.diagnostics.push(
                                            Diagnostic::warning(
                                                DiagnosticCode::UnsupportedFeature,
                                                Phase::Parse,
                                                Fidelity::Approximate,
                                                "non-standard numeric cell addresses were normalized",
                                            )
                                            .in_part(part),
                                        );
                                    }
                                    address
                                }
                                None => {
                                    return Err(format_error(
                                        part,
                                        format!("invalid cell address: {raw_address}"),
                                    ));
                                }
                            },
                        },
                        None => {
                            let row = current_row.map(|row| row.index).ok_or_else(|| {
                                format_error(part, "worksheet cell is outside a row")
                            })?;
                            let canonical = format_a1(next_cell_column, row);
                            CellAddress {
                                column: next_cell_column,
                                row,
                                canonical,
                            }
                        }
                    };
                    next_cell_column = address.column.saturating_add(1);
                    if !addresses.insert(address.canonical.clone()) {
                        return Err(format_error(
                            part,
                            format!("duplicate cell address: {}", address.canonical),
                        ));
                    }
                    let raw_type = optional_attribute(&attributes, "t", part)?;
                    let value_type =
                        parse_cell_value_type(raw_type.as_deref(), &address.canonical, part)?;
                    let explicit_style_index = optional_attribute(&attributes, "s", part)?
                        .map(|index| parse_zero_based_index(&index, part, "cell style index"))
                        .transpose()?;
                    let inherited_style_index = current_row
                        .and_then(|row| row.style_index)
                        .or_else(|| column_style_index(&column_spans, address.column));
                    let style_index = explicit_style_index.or(inherited_style_index).unwrap_or(0);
                    let style_index = if style_index < styles.cells.len() {
                        style_index
                    } else {
                        if !state.diagnostics.iter().any(|diagnostic| {
                            diagnostic.message.contains("missing style")
                                && diagnostic.location.part.as_deref() == Some(part)
                        }) {
                            state.diagnostics.push(
                                Diagnostic::warning(
                                    DiagnosticCode::UnsupportedFeature,
                                    Phase::Parse,
                                    Fidelity::Approximate,
                                    "XLSX cell references a missing style; the default style was substituted",
                                )
                                .in_part(part),
                            );
                        }
                        0
                    };
                    let cell = CellState {
                        depth,
                        address,
                        style_index,
                        value_type,
                        value_depth: None,
                        value_seen: false,
                        value: String::new(),
                        inline_string_depth: None,
                        inline_text_depth: None,
                        inline_string_seen: false,
                        inline_text: String::new(),
                        inline_run_depth: None,
                        inline_properties_depth: None,
                        inline_phonetic_depth: None,
                        inline_run: None,
                        inline_runs: Vec::new(),
                        has_formula: false,
                        shared_formula: None,
                        formula_depth: None,
                        formula: String::new(),
                    };
                    if empty {
                        cells.push(cell);
                    } else {
                        current_cell = Some(cell);
                    }
                } else if let Some(cell) = current_cell.as_mut() {
                    match local {
                        "v" => {
                            if cell.value_seen {
                                return Err(format_error(
                                    part,
                                    "cell contains multiple cached values",
                                ));
                            }
                            cell.value_seen = true;
                            cell.value_depth = (!empty).then_some(depth);
                        }
                        "is" => {
                            if cell.value_type != CellValueType::InlineString {
                                return Err(format_error(
                                    part,
                                    format!(
                                        "cell {} contains inline text without t=inlineStr",
                                        cell.address.canonical
                                    ),
                                ));
                            }
                            if cell.inline_string_seen {
                                return Err(format_error(
                                    part,
                                    "cell contains multiple inline strings",
                                ));
                            }
                            cell.inline_string_seen = true;
                            cell.inline_string_depth = (!empty).then_some(depth);
                        }
                        "rPh" if cell.inline_string_depth.is_some() => {
                            cell.inline_phonetic_depth = (!empty).then_some(depth);
                        }
                        "r" if cell.inline_string_depth.is_some()
                            && cell.inline_phonetic_depth.is_none() =>
                        {
                            cell.inline_run_depth = (!empty).then_some(depth);
                            cell.inline_run = Some(SharedStringRun::default());
                            if empty {
                                cell.inline_runs.push(
                                    cell.inline_run.take().expect("inline run was initialized"),
                                );
                            }
                        }
                        "rPr" if cell.inline_run_depth.is_some() => {
                            cell.inline_properties_depth = (!empty).then_some(depth);
                        }
                        _ if cell.inline_properties_depth.is_some() => {
                            apply_spreadsheet_run_property(
                                local,
                                &attributes,
                                cell.inline_run.as_mut(),
                                &styles.theme_colors,
                                part,
                            )?;
                        }
                        "t" if cell.inline_string_depth.is_some()
                            && cell.inline_phonetic_depth.is_none() =>
                        {
                            if cell.inline_text_depth.is_some() {
                                return Err(format_error(
                                    part,
                                    "nested inline-string text is invalid",
                                ));
                            }
                            cell.inline_text_depth = (!empty).then_some(depth);
                        }
                        "f" => {
                            if cell.has_formula {
                                return Err(format_error(part, "cell contains multiple formulas"));
                            }
                            cell.has_formula = true;
                            if optional_attribute(&attributes, "t", part)?.as_deref() == Some("shared") {
                                cell.shared_formula = optional_attribute(&attributes, "si", part)?
                                    .and_then(|index| index.parse().ok());
                            }
                            cell.formula_depth = (!empty).then_some(depth);
                        }
                        _ => {}
                    }
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if local == "v" {
                    if let Some(cell) = current_cell.as_mut()
                        && cell.value_depth == Some(depth)
                    {
                        cell.value_depth = None;
                    }
                } else if local == "t" {
                    if let Some(cell) = current_cell.as_mut()
                        && cell.inline_text_depth == Some(depth)
                    {
                        cell.inline_text_depth = None;
                    }
                } else if local == "rPr" {
                    if let Some(cell) = current_cell.as_mut()
                        && cell.inline_properties_depth == Some(depth)
                    {
                        cell.inline_properties_depth = None;
                    }
                } else if local == "r" {
                    if let Some(cell) = current_cell.as_mut()
                        && cell.inline_run_depth == Some(depth)
                    {
                        cell.inline_run_depth = None;
                        if let Some(run) = cell.inline_run.take() {
                            cell.inline_runs.push(run);
                        }
                    }
                } else if local == "rPh" {
                    if let Some(cell) = current_cell.as_mut()
                        && cell.inline_phonetic_depth == Some(depth)
                    {
                        cell.inline_phonetic_depth = None;
                    }
                } else if local == "is" {
                    if let Some(cell) = current_cell.as_mut()
                        && cell.inline_string_depth == Some(depth)
                    {
                        cell.inline_string_depth = None;
                    }
                } else if local == "f" {
                    if let Some(cell) = current_cell.as_mut()
                        && cell.formula_depth == Some(depth)
                    {
                        cell.formula_depth = None;
                    }
                } else if local == "c"
                    && current_cell
                        .as_ref()
                        .is_some_and(|cell| cell.depth == depth)
                {
                    let cell = current_cell
                        .take()
                        .ok_or_else(|| format_error(part, "cell state was lost before closing"))?;
                    cells.push(cell);
                } else if local == "row"
                    && current_row.is_some_and(|row| row.depth == depth)
                {
                    current_row = None;
                } else if local == "sheetData" && sheet_data_depth == Some(depth) {
                    sheet_data_depth = None;
                } else if local == "cols" && cols_depth == Some(depth) {
                    cols_depth = None;
                }
            }
            XmlEvent::Text(text) => {
                if let Some(cell) = current_cell
                    .as_mut()
                    .filter(|cell| {
                        cell.value_depth.is_some()
                            || cell.inline_text_depth.is_some()
                            || cell.formula_depth.is_some()
                    })
                {
                    let decoded = decode_xml_text(text).map_err(|error| with_part(error, part))?;
                    if cell.value_depth.is_some() {
                        cell.value.push_str(&decoded);
                    } else if cell.inline_text_depth.is_some() {
                        cell.inline_text.push_str(&decoded);
                        if let Some(run) = cell.inline_run.as_mut() {
                            run.text.push_str(&decoded);
                        }
                    } else {
                        cell.formula.push_str(&decoded);
                    }
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(cell) = current_cell
                    .as_mut()
                    .filter(|cell| {
                        cell.value_depth.is_some()
                            || cell.inline_text_depth.is_some()
                            || cell.formula_depth.is_some()
                    })
                {
                    if cell.value_depth.is_some() {
                        cell.value.push_str(text);
                    } else if cell.inline_text_depth.is_some() {
                        cell.inline_text.push_str(text);
                        if let Some(run) = cell.inline_run.as_mut() {
                            run.text.push_str(text);
                        }
                    } else {
                        cell.formula.push_str(text);
                    }
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;

    let table_styles = parse_table_styles(package, part, &table_relationship_ids, state)?;

    let (mut conditional_formatting, mut unsupported_conditional) =
        parse_conditional_formatting_xml(&bytes, package.limits(), &styles.theme_colors, part)?;
    for rule in conditional_formatting
        .iter_mut()
        .flat_map(|block| &mut block.rules)
    {
        if let ConditionalRuleKind::Calculated { matches, .. } = &mut rule.kind {
            if let Some(values) = formula_evaluator
                .and_then(|values| values.conditional_matches(source_sheet_index, rule.priority))
            {
                matches.clone_from(values);
            } else {
                unsupported_conditional = true;
            }
        }
    }
    let _ = parse_data_validations_xml(&bytes, package.limits(), part)?;
    let (sparkline_groups, mut unsupported_sparklines) =
        parse_sparkline_groups_xml(&bytes, package.limits(), &styles.theme_colors, part)?;
    if conditional_formatting
        .iter()
        .flat_map(|block| &block.rules)
        .filter_map(|rule| match &rule.kind {
            ConditionalRuleKind::CellIs { dxf_id, .. }
            | ConditionalRuleKind::Calculated { dxf_id, .. } => Some(*dxf_id),
            _ => None,
        })
        .any(|dxf_id| dxf_id >= styles.differentials.len())
    {
        return Err(format_error(
            part,
            "conditional formatting references a missing differential style",
        ));
    }
    if unsupported_conditional {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Omitted,
                "one or more XLSX conditional-format rules are not rendered",
            )
            .in_part(part),
        );
    }
    let mut dimension = dimension.unwrap_or_else(|| derive_dimension(&cells));
    for cell in &cells {
        dimension.start_column = dimension.start_column.min(cell.address.column);
        dimension.start_row = dimension.start_row.min(cell.address.row);
        dimension.end_column = dimension.end_column.max(cell.address.column);
        dimension.end_row = dimension.end_row.max(cell.address.row);
    }
    for merged in &merges {
        dimension.start_column = dimension.start_column.min(merged.start_column);
        dimension.start_row = dimension.start_row.min(merged.start_row);
        dimension.end_column = dimension.end_column.max(merged.end_column);
        dimension.end_row = dimension.end_row.max(merged.end_row);
    }
    for span in &column_spans {
        if span.end < MAX_COLUMNS - 1 {
            dimension.start_column = dimension.start_column.min(span.start);
            dimension.end_column = dimension.end_column.max(span.end);
        }
    }
    validate_merge_ranges(&mut merges, part)?;

    let base_style = styles
        .cells
        .first()
        .expect("XLSX styles always have a default cell");
    let mut inherited_cells = Vec::new();
    for row in dimension.start_row..=dimension.end_row {
        for column in dimension.start_column..=dimension.end_column {
            let style_index = row_styles
                .get(&row)
                .copied()
                .or_else(|| column_style_index(&column_spans, column))
                .unwrap_or(0);
            let canonical = format_a1(column, row);
            if addresses.contains(&canonical)
                || !empty_cell_style_is_visible(
                    &table_cell_style(
                        column,
                        row,
                        &styles.cells[style_index],
                        &table_styles,
                        styles,
                    ),
                    base_style,
                )
            {
                continue;
            }
            if state
                .objects
                .len()
                .checked_add(cells.len())
                .and_then(|count| count.checked_add(inherited_cells.len()))
                .is_none_or(|count| count >= package.limits().max_document_objects)
            {
                return Err(object_limit_error(
                    part,
                    "inherited row, column, and table styles exceed the configured object limit",
                ));
            }
            addresses.insert(canonical.clone());
            inherited_cells.push(CellState {
                depth: 0,
                address: CellAddress {
                    column,
                    row,
                    canonical,
                },
                style_index,
                value_type: CellValueType::Number,
                value_depth: None,
                value_seen: false,
                value: String::new(),
                inline_string_depth: None,
                inline_text_depth: None,
                inline_string_seen: false,
                inline_text: String::new(),
                inline_run_depth: None,
                inline_properties_depth: None,
                inline_phonetic_depth: None,
                inline_run: None,
                inline_runs: Vec::new(),
                has_formula: false,
                shared_formula: None,
                formula_depth: None,
                formula: String::new(),
            });
        }
    }
    inherited_cells.extend(cells);
    cells = inherited_cells;

    let mut rows = (dimension.end_row + 1).max(frozen_rows).max(1);
    let mut columns = (dimension.end_column + 1).max(frozen_columns).max(1);
    let default_column_width = default_column_width.unwrap_or_else(|| {
        base_column_width
            .map(|width| (width * maximum_digit_width + 5.0).floor())
            .unwrap_or_else(|| default_excel_column_width(maximum_digit_width))
    });
    let column_metrics = column_metrics(default_column_width, columns, &column_spans, part)?;
    let row_metrics = AxisMetrics::new(default_row_height, row_sizes.into_iter().collect());
    for merged in &merges {
        let canonical = format_a1(merged.start_column, merged.start_row);
        if addresses.insert(canonical.clone()) {
            if state
                .objects
                .len()
                .checked_add(cells.len())
                .is_none_or(|count| count >= package.limits().max_document_objects)
            {
                return Err(object_limit_error(
                    part,
                    "merged cells exceed the configured object limit",
                ));
            }
            cells.push(CellState {
                depth: 0,
                address: CellAddress {
                    column: merged.start_column,
                    row: merged.start_row,
                    canonical,
                },
                style_index: 0,
                value_type: CellValueType::Number,
                value_depth: None,
                value_seen: false,
                value: String::new(),
                inline_string_depth: None,
                inline_text_depth: None,
                inline_string_seen: false,
                inline_text: String::new(),
                inline_run_depth: None,
                inline_properties_depth: None,
                inline_phonetic_depth: None,
                inline_run: None,
                inline_runs: Vec::new(),
                has_formula: false,
                shared_formula: None,
                formula_depth: None,
                formula: String::new(),
            });
        }
    }
    if let Some(evaluator) = formula_evaluator {
        // Shared followers omit formula text, but inherit the master's refresh policy.
        let time_shared_formulas = cells
            .iter()
            .filter(|cell| formula_calls_any(&cell.formula, &["NOW", "TODAY"]))
            .filter_map(|cell| cell.shared_formula)
            .collect::<HashSet<_>>();
        for cell in &mut cells {
            let address = CalculatedCellAddress {
                sheet: source_sheet_index,
                row: i32::try_from(cell.address.row)
                    .unwrap_or(i32::MAX)
                    .saturating_add(1),
                column: i32::try_from(cell.address.column)
                    .unwrap_or(i32::MAX)
                    .saturating_add(1),
            };
            let array_cell = evaluator.is_array_cell(address);
            if array_cell {
                // Legacy CSE target cells can carry stale cached values even when
                // the anchor has no cache. The calculated array is authoritative.
                cell.value.clear();
                cell.value_seen = false;
            }
            if !(cell.has_formula || array_cell)
                || cell.value_seen
                    && !formula_calls_any(&cell.formula, &["NOW", "TODAY"])
                    && !cell
                        .shared_formula
                        .as_ref()
                        .is_some_and(|shared| time_shared_formulas.contains(shared))
            {
                continue;
            }
            match evaluator.value(address).cloned() {
                Ok(FormulaValue::Blank) => {
                    cell.value.clear();
                    cell.value_type = CellValueType::String;
                    cell.value_seen = true;
                }
                Ok(FormulaValue::Boolean(value)) => {
                    cell.value.clear();
                    cell.value_type = CellValueType::Boolean;
                    cell.value_seen = true;
                    cell.value.push(if value { '1' } else { '0' });
                }
                Ok(FormulaValue::Error(value)) => {
                    cell.value_type = CellValueType::Error;
                    cell.value_seen = true;
                    cell.value = value;
                }
                Ok(FormulaValue::Number(value)) if value.is_finite() => {
                    cell.value_type = CellValueType::Number;
                    cell.value_seen = true;
                    cell.value = value.to_string();
                }
                Ok(FormulaValue::Number(_)) | Err(_) => {}
                Ok(FormulaValue::Text(value)) => {
                    cell.value_type = CellValueType::String;
                    cell.value_seen = true;
                    cell.value = value;
                }
            }
        }
    }
    let formula_errors = cells
        .iter()
        .filter(|cell| {
            !cell.value_seen
                && (cell.has_formula
                    || formula_evaluator.is_some_and(|evaluator| {
                        evaluator.is_array_cell(CalculatedCellAddress {
                            sheet: source_sheet_index,
                            row: i32::try_from(cell.address.row)
                                .unwrap_or(i32::MAX)
                                .saturating_add(1),
                            column: i32::try_from(cell.address.column)
                                .unwrap_or(i32::MAX)
                                .saturating_add(1),
                        })
                    }))
        })
        .count();
    if formula_errors != 0 {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Approximate,
                format!(
                    "{formula_errors} XLSX formula cell(s) have neither a cached value nor a usable calculated result"
                ),
            )
            .in_part(part),
        );
    }
    let occupied_cells = cells
        .iter()
        .filter(|cell| cell.value_seen || cell.inline_string_seen || cell.has_formula)
        .map(|cell| (cell.address.column, cell.address.row))
        .collect::<HashSet<_>>();
    // Excel lets unwrapped text overflow through empty cells. Paint every empty
    // cell first so its background cannot erase text emitted by an earlier cell.
    // Non-empty cells remain in worksheet order and still stop that overflow.
    cells.sort_by_key(|cell| cell.value_seen || cell.inline_string_seen || cell.has_formula);
    let conditional_values = cells
        .iter()
        .filter_map(|cell| {
            cell_numeric_value(cell).map(|value| ((cell.address.column, cell.address.row), value))
        })
        .collect::<HashMap<_, _>>();
    let conditional_stats = conditional_formatting
        .iter()
        .map(|block| conditional_statistics(block, &conditional_values))
        .collect::<Vec<_>>();
    let cell_merges = assign_cell_merges(&cells, &merges);
    let mut conditional_cells = Vec::new();
    let mut trailing_text_extent = 0.0_f32;
    let cell_count = cells.len();
    for (cell_index, (cell, merge_index)) in cells.into_iter().zip(cell_merges).enumerate() {
        let merged = merge_index.map(|index| &merges[index]);
        if merged.is_some_and(|merged| {
            cell.address.column != merged.start_column || cell.address.row != merged.start_row
        }) {
            continue;
        }
        let bounds = match merged {
            Some(merged) => Rect {
                x: column_metrics.offset(merged.start_column),
                y: row_metrics.offset(merged.start_row),
                width: column_metrics.span(merged.start_column, merged.end_column),
                height: row_metrics.span(merged.start_row, merged.end_row),
            },
            None => Rect {
                x: column_metrics.offset(cell.address.column),
                y: row_metrics.offset(cell.address.row),
                width: column_metrics.size(cell.address.column),
                height: row_metrics.size(cell.address.row),
            },
        };
        if bounds.width == 0.0 || bounds.height == 0.0 {
            continue;
        }
        let style_index = cell.style_index;
        let table_style = table_cell_style(
            cell.address.column,
            cell.address.row,
            &styles.cells[style_index],
            &table_styles,
            styles,
        );
        let mut style = conditional_cell_style(
            &cell,
            &table_style,
            &conditional_formatting,
            &conditional_stats,
            &styles.differentials,
            shared_strings,
        );
        style.align = cell_text_align(&style, cell.value_type);
        if let Some(value) = cell_numeric_value(&cell) {
            conditional_cells.push(ConditionalCell {
                address: cell.address.clone(),
                bounds,
                value,
            });
        }
        let (hide_value, has_icon) = conditional_value_layout(&cell, &conditional_formatting);
        let icon_inset = if has_icon {
            conditional_icon_bounds(bounds).width + 2.0
        } else {
            0.0
        };
        let first_column = merged.map_or(cell.address.column, |range| range.start_column);
        let next_column = merged.map_or(cell.address.column, |range| range.end_column) + 1;
        let previous_occupied = first_column
            .checked_sub(1)
            .is_some_and(|column| occupied_cells.contains(&(column, cell.address.row)));
        let next_occupied = occupied_cells.contains(&(next_column, cell.address.row));
        let clip_horizontal_overflow = crate::model::spreadsheet_text_overflow_is_clipped(
            style.align,
            previous_occupied,
            next_occupied,
        );
        if !hide_value
            && !clip_horizontal_overflow
            && !style.wrap
            && !style.shrink_to_fit
            && style.rotation_degrees == 0.0
            && next_column >= columns
            && let Some(text) = cell_conditional_text(&cell, shared_strings)
        {
            let text_width = text
                .chars()
                .map(|character| drawingml_fallback_character_width(character, style.font_size))
                .sum::<f32>();
            let right = match style.align {
                TextAlign::Start => bounds.x + 2.0 + text_width,
                TextAlign::Center => bounds.x + (bounds.width + text_width) / 2.0,
                _ => bounds.x + bounds.width,
            };
            trailing_text_extent = trailing_text_extent.max(right);
        }
        push_cell(
            cell,
            bounds,
            (((bounds.width - icon_inset).max(0.0) / maximum_digit_width).floor() as usize)
                .clamp(1, 11),
            part,
            sheet_name,
            unit_index,
            shared_strings,
            &style,
            date_1904,
            hide_value,
            icon_inset,
            show_grid_lines,
            clip_horizontal_overflow,
            package.limits().max_document_objects,
            cell_count - cell_index,
            state,
        )?;
    }
    columns = columns.max(column_metrics.count_for_extent(trailing_text_extent, MAX_COLUMNS));
    push_conditional_overlays(
        &conditional_cells,
        &conditional_formatting,
        &conditional_stats,
        part,
        sheet_name,
        unit_index,
        package.limits().max_document_objects,
        state,
    )?;
    unsupported_sparklines |= push_sparklines(
        &sparkline_groups,
        &conditional_values,
        dimension,
        &column_metrics,
        &row_metrics,
        part,
        sheet_name,
        unit_index,
        package.limits().max_document_objects,
        state,
    )?;
    if unsupported_sparklines {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Omitted,
                "one or more XLSX sparkline groups or references are not rendered",
            )
            .in_part(part),
        );
    }
    let (drawing_width, drawing_height) = push_worksheet_drawings(
        package,
        part,
        &drawing_relationship_ids,
        &legacy_drawing_relationship_ids,
        &column_metrics,
        &row_metrics,
        sheet_name,
        unit_index,
        content_types,
        &styles.theme_colors,
        &styles.theme_line_styles,
        styles.chart_theme_colors.as_ref(),
        chart_named_data,
        state,
    )?;
    columns = columns.max(column_metrics.count_for_extent(drawing_width, MAX_COLUMNS));
    rows = rows.max(row_metrics.count_for_extent(drawing_height, MAX_ROWS));

    let width = column_metrics.offset(columns).max(drawing_width).max(1.0);
    let height = row_metrics.offset(rows).max(drawing_height).max(1.0);
    if let Some(id) = background_relationship_id {
        let relationships = package.relationships(Some(part))?;
        let image = match relationships
            .iter()
            .find(|relationship| relationship.id == id)
        {
            Some(relationship) => {
                materialize_worksheet_image(package, part, relationship, content_types, state)
            }
            None => Err(format_error(
                part,
                "worksheet background image relationship is missing",
            )),
        };
        match image {
            Ok(Some((media_type, bytes))) => {
                if state.objects.len() >= package.limits().max_document_objects {
                    return Err(Diagnostic::fatal(
                        DiagnosticCode::ObjectLimit,
                        Phase::Parse,
                        None,
                        "worksheet background exceeds the object limit",
                    )
                    .in_part(part));
                }
                let numeric_id = state.objects.len() as u32;
                state.objects.push(Object {
                    numeric_id,
                    parent_numeric_id: None,
                    stable_id: format!("object:{numeric_id}"),
                    parent_stable_id: None,
                    kind: ObjectKind::Shape,
                    unit_index,
                    bounds: Rect {
                        x: 0.0,
                        y: 0.0,
                        width,
                        height,
                    },
                    z: -1,
                    text: None,
                    source: SourceRef {
                        part: part.to_owned(),
                        mapping: MappingQuality::Exact,
                        locator: SourceLocator::Xlsx {
                            kind: "drawing",
                            sheet_name: sheet_name.to_owned(),
                            address: None,
                            formula: None,
                            drawing_id: None,
                        },
                    },
                    visual: Visual::PaintedShape {
                        geometry: Geometry::Rectangle,
                        fill: Paint::Image {
                            media_type,
                            bytes,
                            crop: ImageCrop::default(),
                            tile: true,
                            tile_width: None,
                            tile_height: None,
                            mapping: None,
                        },
                        stroke: Paint::None,
                        stroke_width: 0.0,
                    },
                });
            }
            Ok(None) => {}
            Err(error) if matches!(error.code, DiagnosticCode::FormatInvalid) => {
                state.diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Parse,
                        Fidelity::Omitted,
                        "worksheet background-picture drawing could not be loaded",
                    )
                    .in_part(part),
                );
            }
            Err(error) => return Err(error),
        }
    }

    Ok(Unit {
        kind: UnitKind::Sheet,
        index: unit_index,
        id: format!("unit:{unit_index}"),
        name: sheet_name.to_owned(),
        width,
        height,
        rows,
        columns,
        frozen_rows,
        frozen_columns,
        frozen_width: column_metrics.offset(frozen_columns),
        frozen_height: row_metrics.offset(frozen_rows),
        row_axis: row_metrics.descriptor(rows),
        column_axis: column_metrics.descriptor(columns),
        show_grid_lines,
        tab_color,
        sheet: print_settings.finish(),
        slide: None,
    })
}

fn parse_worksheet_drawing(
    package: &Package<'_>,
    part: &str,
    column_metrics: &AxisMetrics,
    row_metrics: &AxisMetrics,
    theme_colors: &ThemeColors,
    theme_line_styles: &DrawingMlThemeLineStyles,
) -> Result<Vec<WorksheetDrawing>, Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut depth = 0_usize;
    let mut anchor = None::<DrawingAnchorState>;
    let mut groups = Vec::<DrawingGroupState>::new();
    let mut shape = None::<WorksheetShapeState>;
    let mut color = None::<DrawingColorState>;
    let mut chart_choice = super::drawingml::DrawingMlChartChoice::default();
    let mut drawings = Vec::new();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if chart_choice.start(local, depth, empty) {
                    depth += usize::from(!empty);
                    return Ok(());
                }
                let anchor_kind = match local {
                    "twoCellAnchor" => Some(DrawingAnchorKind::TwoCell),
                    "oneCellAnchor" => Some(DrawingAnchorKind::OneCell),
                    "absoluteAnchor" => Some(DrawingAnchorKind::Absolute),
                    _ => None,
                };
                if let Some(kind) = anchor_kind {
                    if anchor.is_some() {
                        return Err(format_error(
                            part,
                            "nested worksheet drawing anchors are invalid",
                        ));
                    }
                    anchor = Some(DrawingAnchorState {
                        depth,
                        kind,
                        from: DrawingMarker::default(),
                        to: DrawingMarker::default(),
                        marker_depth: None,
                        field: None,
                        x: None,
                        y: None,
                        width: None,
                        height: None,
                        drawing_id: None,
                        image_relationship_id: None,
                        image_crop: ImageCrop::default(),
                        image_adjustment: ImageAdjustment::default(),
                        image_geometry: super::drawingml::DrawingMlPictureGeometry::default(),
                        image_effects: DrawingMlPictureEffectsCapture::default(),
                        image_rotation_degrees: 0.0,
                        image_flip_h: false,
                        image_flip_v: false,
                        image_blip_depth: None,
                        duotone_depth: None,
                        duotone_colors: Vec::new(),
                        duotone_color: None,
                        chart_relationship_id: None,
                        diagram_relationship_id: None,
                    });
                } else if let Some(current) = anchor.as_mut() {
                    if matches!(local, "from" | "to") {
                        current.marker_depth = (!empty).then_some((depth, local == "to"));
                    } else if let Some((_, to)) = current.marker_depth {
                        let field = match (to, local) {
                            (false, "col") => Some(DrawingMarkerField::FromColumn),
                            (false, "colOff") => Some(DrawingMarkerField::FromColumnOffset),
                            (false, "row") => Some(DrawingMarkerField::FromRow),
                            (false, "rowOff") => Some(DrawingMarkerField::FromRowOffset),
                            (true, "col") => Some(DrawingMarkerField::ToColumn),
                            (true, "colOff") => Some(DrawingMarkerField::ToColumnOffset),
                            (true, "row") => Some(DrawingMarkerField::ToRow),
                            (true, "rowOff") => Some(DrawingMarkerField::ToRowOffset),
                            _ => None,
                        };
                        if let Some(field) = field {
                            if current.field.is_some() {
                                return Err(format_error(
                                    part,
                                    "nested worksheet drawing marker fields are invalid",
                                ));
                            }
                            current.field = (!empty).then_some((depth, field, String::new()));
                        }
                    } else if local == "pos" && depth == current.depth + 1 {
                        current.x = Some(parse_i64(
                            &required_attribute(&attributes, "x", part)?,
                            part,
                            "drawing x coordinate",
                        )?);
                        current.y = Some(parse_i64(
                            &required_attribute(&attributes, "y", part)?,
                            part,
                            "drawing y coordinate",
                        )?);
                    } else if local == "ext" && depth == current.depth + 1 {
                        current.width = Some(parse_i64(
                            &required_attribute(&attributes, "cx", part)?,
                            part,
                            "drawing width",
                        )?);
                        current.height = Some(parse_i64(
                            &required_attribute(&attributes, "cy", part)?,
                            part,
                            "drawing height",
                        )?);
                    } else if current.image_blip_depth.is_some() {
                        if let Some((_, color)) = current.duotone_color.as_mut() {
                            if matches!(
                                local,
                                "tint"
                                    | "shade"
                                    | "lumMod"
                                    | "lumOff"
                                    | "satMod"
                                    | "satOff"
                                    | "alpha"
                            ) {
                                apply_color_transform(
                                    color,
                                    local,
                                    parse_drawing_percentage(&attributes, part)?,
                                );
                            }
                        } else if local == "grayscl" {
                            current.image_adjustment.grayscale = true;
                        } else if local == "biLevel" {
                            current.image_adjustment.bilevel_threshold = Some(
                                (parse_i64(
                                    &required_attribute(&attributes, "thresh", part)?,
                                    part,
                                    "image bilevel threshold",
                                )? as f32
                                    / 100_000.0)
                                    .clamp(0.0, 1.0),
                            );
                        } else if local == "lum" {
                            current.image_adjustment.brightness =
                                optional_attribute(&attributes, "bright", part)?
                                    .map(|value| {
                                        parse_i64(&value, part, "image brightness")
                                            .map(|value| value as f32 / 100_000.0)
                                    })
                                    .transpose()?
                                    .unwrap_or(0.0)
                                    .clamp(-1.0, 1.0);
                            current.image_adjustment.contrast =
                                optional_attribute(&attributes, "contrast", part)?
                                    .map(|value| {
                                        parse_i64(&value, part, "image contrast")
                                            .map(|value| value as f32 / 100_000.0)
                                    })
                                    .transpose()?
                                    .unwrap_or(0.0)
                                    .clamp(-1.0, 1.0);
                        } else if local == "duotone" {
                            current.duotone_depth = (!empty).then_some(depth);
                            current.duotone_colors.clear();
                        } else if current.duotone_depth.is_some()
                            && matches!(local, "srgbClr" | "schemeClr" | "prstClr")
                        {
                            let value = required_attribute(&attributes, "val", part)?;
                            let color = drawing_color_value(local, &value, theme_colors, part)?;
                            if empty {
                                current.duotone_colors.push(color);
                            } else {
                                current.duotone_color = Some((depth, color));
                            }
                        }
                    } else if current.image_relationship_id.is_some()
                        && shape.is_none()
                        && current.image_effects.start(
                            local,
                            &attributes,
                            empty,
                            depth,
                            part,
                            |kind, value| drawing_color_value(kind, value, theme_colors, part),
                        )?
                    {
                    } else if local == "xfrm" && current.image_relationship_id.is_some() {
                        current.image_flip_h =
                            parse_bool_attribute(&attributes, "flipH", part)?.unwrap_or(false);
                        current.image_flip_v =
                            parse_bool_attribute(&attributes, "flipV", part)?.unwrap_or(false);
                        current.image_rotation_degrees =
                            optional_attribute(&attributes, "rot", part)?
                                .map(|rotation| {
                                    parse_i64(&rotation, part, "picture rotation")
                                        .map(|value| value as f32 / 60_000.0)
                                })
                                .transpose()?
                                .unwrap_or(0.0);
                    } else if local == "srcRect" && current.image_relationship_id.is_some() {
                        current.image_crop =
                            super::drawingml::drawingml_picture_crop(&attributes, part)?;
                    } else if current.image_relationship_id.is_some()
                        && shape.is_none()
                        && matches!(
                            local,
                            "prstGeom"
                                | "custGeom"
                                | "avLst"
                                | "gdLst"
                                | "gd"
                                | "pathLst"
                                | "path"
                                | "moveTo"
                                | "lnTo"
                                | "cubicBezTo"
                                | "quadBezTo"
                                | "pt"
                                | "arcTo"
                                | "close"
                        )
                    {
                        let bounds = drawing_anchor_bounds(current, column_metrics, row_metrics)
                            .unwrap_or_default();
                        current.image_geometry.start(
                            local,
                            &attributes,
                            empty,
                            depth,
                            part,
                            bounds.width * EMU_PER_CSS_PIXEL,
                            bounds.height * EMU_PER_CSS_PIXEL,
                        )?;
                    } else if matches!(local, "sp" | "cxnSp") && shape.is_none() {
                        shape = Some(WorksheetShapeState::new(
                            depth,
                            local == "cxnSp",
                            theme_colors,
                        ));
                    } else if let Some(active) = shape.as_mut() {
                        if let Some(image_fill) = active
                            .image_fill
                            .as_mut()
                            .filter(|image_fill| image_fill.parse_depth.is_some())
                        {
                            super::drawingml::drawingml_image_fill_mapping(
                                &mut image_fill.mapping,
                                local,
                                &attributes,
                                part,
                            )?;
                            if local == "blip" {
                                image_fill.relationship_id =
                                    optional_attribute(&attributes, "embed", part)?;
                            } else if local == "srcRect" {
                                let crop = |name| -> Result<f32, Diagnostic> {
                                    optional_attribute(&attributes, name, part)?
                                        .map(|value| {
                                            parse_i64(&value, part, "shape image-fill crop")
                                                .map(|value| value as f32 / 100_000.0)
                                        })
                                        .transpose()
                                        .map(|value| value.unwrap_or(0.0))
                                };
                                image_fill.crop = ImageCrop {
                                    left: crop("l")?,
                                    top: crop("t")?,
                                    right: crop("r")?,
                                    bottom: crop("b")?,
                                };
                            } else if local == "tile" {
                                image_fill.tile = true;
                            }
                        } else if let Some(capture) = active.text_stroke_fill_capture.as_mut() {
                            capture.start(local, &attributes, part, &|value| {
                                drawing_color_value("schemeClr", value, theme_colors, part).ok()
                            })?;
                        } else if local == "gradFill" && !empty && active.text_line_depth.is_some()
                        {
                            active.text_stroke_fill_capture = ChartFillCapture::new(local, depth);
                        } else if let Some(active_color) = color.as_mut() {
                            match local {
                                "lumMod" => {
                                    active_color.luminance_modulation =
                                        parse_drawing_percentage(&attributes, part)?;
                                }
                                "lumOff" => {
                                    active_color.luminance_offset =
                                        parse_drawing_percentage(&attributes, part)?;
                                }
                                "alpha" => {
                                    active_color.alpha =
                                        parse_drawing_percentage(&attributes, part)?;
                                }
                                "tint" | "shade" | "satMod" | "satOff" | "hueMod" | "hueOff" => {
                                    apply_color_transform(
                                        &mut active_color.color,
                                        local,
                                        parse_drawing_percentage(&attributes, part)?,
                                    )
                                }
                                _ => {}
                            }
                        } else if local == "spPr" {
                            active.properties_depth = (!empty).then_some(depth);
                        } else if local == "blipFill"
                            && active
                                .properties_depth
                                .is_some_and(|properties_depth| depth == properties_depth + 1)
                        {
                            active.fill = Paint::None;
                            active.fill_explicit = true;
                            active.image_fill = (!empty).then(|| WorksheetShapeImageFill {
                                parse_depth: Some(depth),
                                ..WorksheetShapeImageFill::default()
                            });
                            if let Some(fill) = active.image_fill.as_mut() {
                                super::drawingml::drawingml_image_fill_mapping(
                                    &mut fill.mapping,
                                    local,
                                    &attributes,
                                    part,
                                )?;
                            }
                        } else if local == "grpFill"
                            && active
                                .properties_depth
                                .is_some_and(|properties_depth| depth == properties_depth + 1)
                            && let Some(group) =
                                groups.iter().rev().find(|group| group.image_fill.is_some())
                            && let Some(mut image_fill) = group.image_fill.clone()
                        {
                            image_fill.parse_depth = None;
                            if !image_fill.tile
                                && group.transform.child_width > 0
                                && group.transform.child_height > 0
                            {
                                let left = (active.transform.x - group.transform.child_x) as f32
                                    / group.transform.child_width as f32;
                                let top = (active.transform.y - group.transform.child_y) as f32
                                    / group.transform.child_height as f32;
                                let right = left
                                    + active.transform.width as f32
                                        / group.transform.child_width as f32;
                                let bottom = top
                                    + active.transform.height as f32
                                        / group.transform.child_height as f32;
                                let source_width =
                                    1.0 - image_fill.crop.left - image_fill.crop.right;
                                let source_height =
                                    1.0 - image_fill.crop.top - image_fill.crop.bottom;
                                image_fill.crop = ImageCrop {
                                    left: image_fill.crop.left
                                        + left.clamp(0.0, 1.0) * source_width,
                                    top: image_fill.crop.top + top.clamp(0.0, 1.0) * source_height,
                                    right: image_fill.crop.right
                                        + (1.0 - right.clamp(0.0, 1.0)) * source_width,
                                    bottom: image_fill.crop.bottom
                                        + (1.0 - bottom.clamp(0.0, 1.0)) * source_height,
                                };
                            }
                            active.fill = Paint::None;
                            active.fill_explicit = true;
                            active.image_fill = Some(image_fill);
                        } else if local == "xfrm" {
                            active.transform_depth = (!empty).then_some(depth);
                            active.transform.flip_h =
                                parse_bool_attribute(&attributes, "flipH", part)?.unwrap_or(false);
                            active.transform.flip_v =
                                parse_bool_attribute(&attributes, "flipV", part)?.unwrap_or(false);
                            active.transform.rotation_degrees =
                                optional_attribute(&attributes, "rot", part)?
                                    .map(|rotation| {
                                        parse_i64(&rotation, part, "shape rotation")
                                            .map(|value| value as f32 / 60_000.0)
                                    })
                                    .transpose()?
                                    .unwrap_or(0.0);
                        } else if active.transform_depth.is_some() && local == "off" {
                            active.transform.x = parse_i64(
                                &required_attribute(&attributes, "x", part)?,
                                part,
                                "shape x coordinate",
                            )?;
                            active.transform.y = parse_i64(
                                &required_attribute(&attributes, "y", part)?,
                                part,
                                "shape y coordinate",
                            )?;
                        } else if active.transform_depth.is_some() && local == "ext" {
                            active.transform.width = parse_i64(
                                &required_attribute(&attributes, "cx", part)?,
                                part,
                                "shape width",
                            )?;
                            active.transform.height = parse_i64(
                                &required_attribute(&attributes, "cy", part)?,
                                part,
                                "shape height",
                            )?;
                        } else if local == "cNvPr" {
                            active.drawing_id = Some(parse_u32(
                                &required_attribute(&attributes, "id", part)?,
                                part,
                                "shape id",
                            )?);
                        } else if local == "prstGeom" {
                            let preset = optional_attribute(&attributes, "prst", part)?;
                            if preset
                                .as_deref()
                                .is_some_and(|value| value.contains("Connector"))
                            {
                                active.is_connector = true;
                                active.fill = Paint::None;
                            }
                            active.geometry =
                                drawing_shape_geometry(preset.as_deref(), active.is_connector);
                            if active.is_connector {
                                active.connector_preset = preset;
                            } else {
                                active.preset_geometry = preset;
                            }
                        } else if local == "ln" {
                            let text_line = active.text_body_depth.is_some();
                            if text_line {
                                active.text_line_depth = (!empty).then_some(depth);
                            } else {
                                active.line_depth = (!empty).then_some(depth);
                                active.shape_stroke_authored = true;
                            }
                            if let Some(width) = optional_attribute(&attributes, "w", part)? {
                                let width = (parse_i64(&width, part, "shape line width")? as f32
                                    / EMU_PER_CSS_PIXEL)
                                    .max(0.0);
                                if text_line {
                                    active.text_stroke_width = width;
                                } else {
                                    active.stroke_width = width;
                                    active.stroke_width_explicit = true;
                                }
                            }
                        } else if local == "effectLst" {
                            active.effect_depth = (!empty).then_some(depth);
                        } else if local == "outerShdw" {
                            let blur = optional_attribute(&attributes, "blurRad", part)?
                                .map(|value| {
                                    parse_i64(&value, part, "shape shadow blur")
                                        .map(|value| value as f32 / EMU_PER_CSS_PIXEL)
                                })
                                .transpose()?
                                .unwrap_or(0.0);
                            let distance = optional_attribute(&attributes, "dist", part)?
                                .map(|value| {
                                    parse_i64(&value, part, "shape shadow distance")
                                        .map(|value| value as f32 / EMU_PER_CSS_PIXEL)
                                })
                                .transpose()?
                                .unwrap_or(0.0);
                            let direction = optional_attribute(&attributes, "dir", part)?
                                .map(|value| parse_i64(&value, part, "shape shadow direction"))
                                .transpose()?
                                .unwrap_or(0) as f32
                                / 60_000.0
                                * std::f32::consts::PI
                                / 180.0;
                            active.outer_shadow = Some(drawingml_outer_shadow(
                                Shadow {
                                    color: 0x0000_0040,
                                    blur,
                                    offset_x: direction.cos() * distance,
                                    offset_y: direction.sin() * distance,
                                },
                                &attributes,
                                part,
                            )?);
                            active.shadow_depth = (!empty).then_some(depth);
                        } else if local == "lnRef" {
                            active.style_line_depth = (!empty).then_some(depth);
                            active.shape_stroke_authored = true;
                            if !active.stroke_width_explicit
                                && let Some(width) =
                                    theme_line_styles.referenced_width(&attributes, part)?
                            {
                                active.stroke_width = width;
                            }
                        } else if local == "fillRef" {
                            active.style_fill_depth = (!empty
                                && drawingml_fill_reference_has_paint(&attributes, part)?)
                            .then_some(depth);
                        } else if matches!(local, "hiddenFill" | "hiddenLine") {
                            active.hidden_style_depth = (!empty).then_some(depth);
                        } else if local == "prstDash" && active.line_depth.is_some() {
                            active.dash_pattern = drawing_dash_pattern(
                                optional_attribute(&attributes, "val", part)?.as_deref(),
                            );
                        } else if local == "headEnd" && active.line_depth.is_some() {
                            active.head_arrow = optional_attribute(&attributes, "type", part)?
                                .as_deref()
                                .is_some_and(|kind| kind != "none");
                        } else if local == "tailEnd" && active.line_depth.is_some() {
                            active.tail_arrow = optional_attribute(&attributes, "type", part)?
                                .as_deref()
                                .is_some_and(|kind| kind != "none");
                        } else if local == "noFill" {
                            if active.text_line_depth.is_some() {
                                active.text_stroke_color = 0;
                                active.text_stroke_fill_capture = None;
                                active.text_stroke_fill = None;
                            } else if active.line_depth.is_some() {
                                active.stroke = Paint::None;
                                active.stroke_explicit = true;
                            } else if active.text_body_depth.is_some() {
                                active.font_color = 0x0000_0000;
                            } else if active.effect_depth.is_some() {
                            } else {
                                active.fill = Paint::None;
                                active.fill_explicit = true;
                            }
                        } else if matches!(local, "srgbClr" | "schemeClr" | "prstClr") {
                            let target = if active.text_line_depth.is_some() {
                                Some((DrawingColorTarget::TextStroke, true))
                            } else if active.line_depth.is_some() {
                                Some((DrawingColorTarget::Stroke, true))
                            } else if active.shadow_depth.is_some() {
                                Some((DrawingColorTarget::Shadow, true))
                            } else if active.text_body_depth.is_some() {
                                Some((DrawingColorTarget::Text, true))
                            } else if active.hidden_style_depth.is_some() {
                                None
                            } else if active.style_line_depth.is_some() && !active.stroke_explicit {
                                Some((DrawingColorTarget::Stroke, false))
                            } else if active.style_fill_depth.is_some()
                                && !active.fill_explicit
                                && !active.is_connector
                            {
                                Some((DrawingColorTarget::Fill, false))
                            } else if active.properties_depth.is_some() {
                                Some((DrawingColorTarget::Fill, true))
                            } else {
                                None
                            };
                            if let Some((target, explicit)) = target {
                                if matches!(target, DrawingColorTarget::TextStroke) {
                                    active.text_stroke_fill = None;
                                }
                                let base = drawing_color_value(
                                    local,
                                    &required_attribute(&attributes, "val", part)?,
                                    theme_colors,
                                    part,
                                )?;
                                let next = DrawingColorState {
                                    depth,
                                    target,
                                    color: base,
                                    luminance_modulation: 1.0,
                                    luminance_offset: 0.0,
                                    alpha: 1.0,
                                    explicit,
                                };
                                if empty {
                                    apply_drawing_color(active, next);
                                } else {
                                    color = Some(next);
                                }
                            }
                        } else if local == "txBody" {
                            active.text_body_depth = (!empty).then_some(depth);
                        } else if local == "bodyPr" {
                            active.vertical_align =
                                match optional_attribute(&attributes, "anchor", part)?.as_deref() {
                                    Some("ctr") => TextVerticalAlign::Center,
                                    Some("b") => TextVerticalAlign::Bottom,
                                    _ => TextVerticalAlign::Top,
                                };
                            active.wrap = optional_attribute(&attributes, "wrap", part)?.as_deref()
                                != Some("none");
                        } else if local == "scene3d"
                            && (active.properties_depth.is_some()
                                || active.text_body_depth.is_some())
                        {
                            active.three_d.get_or_insert_with(ThreeDStyle::default);
                            active.scene_3d_depth = (!empty).then_some(depth);
                        } else if local == "camera" && active.scene_3d_depth.is_some() {
                            parse_three_d_camera(
                                active.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                            )?;
                            active.camera_3d_depth = (!empty).then_some(depth);
                        } else if local == "lightRig" && active.scene_3d_depth.is_some() {
                            parse_three_d_light(
                                active.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                            )?;
                            active.light_3d_depth = (!empty).then_some(depth);
                        } else if local == "rot" && active.camera_3d_depth.is_some() {
                            parse_three_d_rotation(
                                active.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                                true,
                            )?;
                        } else if local == "rot" && active.light_3d_depth.is_some() {
                            parse_three_d_rotation(
                                active.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                                false,
                            )?;
                        } else if local == "bevelT" && active.properties_depth.is_some() {
                            active
                                .three_d
                                .get_or_insert_with(ThreeDStyle::default)
                                .bevel_top = Some(parse_three_d_bevel(&attributes, part)?);
                        } else if local == "bevelB" && active.properties_depth.is_some() {
                            active
                                .three_d
                                .get_or_insert_with(ThreeDStyle::default)
                                .bevel_bottom = Some(parse_three_d_bevel(&attributes, part)?);
                        } else if local == "flatTx" {
                            active
                                .three_d
                                .get_or_insert_with(ThreeDStyle::default)
                                .flat_text_z = Some(
                                optional_attribute(&attributes, "z", part)?
                                    .map(|value| parse_i64(&value, part, "flat text z"))
                                    .transpose()?
                                    .unwrap_or(0) as f32
                                    / EMU_PER_CSS_PIXEL,
                            );
                        } else if local == "p" {
                            active.paragraph_depth = (!empty).then_some(depth);
                        } else if local == "pPr" {
                            active.align =
                                match optional_attribute(&attributes, "algn", part)?.as_deref() {
                                    Some("ctr") => TextAlign::Center,
                                    Some("r") => TextAlign::End,
                                    Some("just") | Some("dist") => TextAlign::Justify,
                                    _ => TextAlign::Start,
                                };
                        } else if matches!(local, "rPr" | "defRPr" | "endParaRPr") {
                            if let Some(size) = optional_attribute(&attributes, "sz", part)? {
                                active.font_size =
                                    parse_u32(&size, part, "shape font size")? as f32 / 100.0
                                        * 96.0
                                        / 72.0;
                            }
                            active.bold = parse_bool_attribute(&attributes, "b", part)?
                                .unwrap_or(active.bold);
                            active.italic = parse_bool_attribute(&attributes, "i", part)?
                                .unwrap_or(active.italic);
                        } else if local == "latin" {
                            if let Some(typeface) =
                                optional_attribute(&attributes, "typeface", part)?
                                && !typeface.starts_with('+')
                            {
                                active.font_family = typeface;
                            }
                        } else if local == "t" {
                            active.text_depth = (!empty).then_some(depth);
                        } else if local == "br" && !active.text.ends_with('\n') {
                            active.text.push('\n');
                        }
                    } else if local == "grpSp" {
                        groups.push(DrawingGroupState {
                            depth,
                            properties_depth: None,
                            transform_depth: None,
                            transform: DrawingTransform::default(),
                            image_fill: None,
                        });
                    } else if let Some(image_fill) = groups
                        .last_mut()
                        .and_then(|group| group.image_fill.as_mut())
                        .filter(|image_fill| image_fill.parse_depth.is_some())
                    {
                        super::drawingml::drawingml_image_fill_mapping(
                            &mut image_fill.mapping,
                            local,
                            &attributes,
                            part,
                        )?;
                        if local == "blip" {
                            image_fill.relationship_id =
                                optional_attribute(&attributes, "embed", part)?;
                        } else if local == "srcRect" {
                            let crop = |name| -> Result<f32, Diagnostic> {
                                optional_attribute(&attributes, name, part)?
                                    .map(|value| {
                                        parse_i64(&value, part, "group image-fill crop")
                                            .map(|value| value as f32 / 100_000.0)
                                    })
                                    .transpose()
                                    .map(|value| value.unwrap_or(0.0))
                            };
                            image_fill.crop = ImageCrop {
                                left: crop("l")?,
                                top: crop("t")?,
                                right: crop("r")?,
                                bottom: crop("b")?,
                            };
                        } else if local == "tile" {
                            image_fill.tile = true;
                        }
                    } else if local == "grpSpPr"
                        && let Some(group) = groups.last_mut()
                    {
                        group.properties_depth = (!empty).then_some(depth);
                    } else if local == "blipFill"
                        && let Some(group) = groups.last_mut()
                        && group
                            .properties_depth
                            .is_some_and(|properties_depth| depth == properties_depth + 1)
                    {
                        group.image_fill = (!empty).then(|| WorksheetShapeImageFill {
                            parse_depth: Some(depth),
                            ..WorksheetShapeImageFill::default()
                        });
                        if let Some(fill) = group.image_fill.as_mut() {
                            super::drawingml::drawingml_image_fill_mapping(
                                &mut fill.mapping,
                                local,
                                &attributes,
                                part,
                            )?;
                        }
                    } else if groups.last().is_some_and(|group| {
                        local == "xfrm"
                            || (group.transform_depth.is_some()
                                && matches!(local, "off" | "ext" | "chOff" | "chExt"))
                    }) {
                        let group = groups.last_mut().expect("checked drawing group");
                        if local == "xfrm" {
                            group.transform_depth = (!empty).then_some(depth);
                        } else if group.transform_depth.is_some() {
                            match local {
                                "off" => {
                                    group.transform.x = parse_i64(
                                        &required_attribute(&attributes, "x", part)?,
                                        part,
                                        "group x coordinate",
                                    )?;
                                    group.transform.y = parse_i64(
                                        &required_attribute(&attributes, "y", part)?,
                                        part,
                                        "group y coordinate",
                                    )?;
                                }
                                "ext" => {
                                    group.transform.width = parse_i64(
                                        &required_attribute(&attributes, "cx", part)?,
                                        part,
                                        "group width",
                                    )?;
                                    group.transform.height = parse_i64(
                                        &required_attribute(&attributes, "cy", part)?,
                                        part,
                                        "group height",
                                    )?;
                                }
                                "chOff" => {
                                    group.transform.child_x = parse_i64(
                                        &required_attribute(&attributes, "x", part)?,
                                        part,
                                        "group child x coordinate",
                                    )?;
                                    group.transform.child_y = parse_i64(
                                        &required_attribute(&attributes, "y", part)?,
                                        part,
                                        "group child y coordinate",
                                    )?;
                                }
                                "chExt" => {
                                    group.transform.child_width = parse_i64(
                                        &required_attribute(&attributes, "cx", part)?,
                                        part,
                                        "group child width",
                                    )?;
                                    group.transform.child_height = parse_i64(
                                        &required_attribute(&attributes, "cy", part)?,
                                        part,
                                        "group child height",
                                    )?;
                                }
                                _ => {}
                            }
                        }
                    } else if local == "cNvPr" {
                        current.drawing_id = Some(parse_u32(
                            &required_attribute(&attributes, "id", part)?,
                            part,
                            "drawing id",
                        )?);
                    } else if local == "blip" {
                        current.image_relationship_id =
                            optional_attribute(&attributes, "embed", part)?;
                        current.image_blip_depth = (!empty).then_some(depth);
                    } else if local == "chart" {
                        chart_choice.chart();
                        current.chart_relationship_id = Some(relationship_id(&attributes, part)?);
                    } else if local == "relIds" {
                        current.diagram_relationship_id =
                            optional_attribute(&attributes, "dm", part)?;
                    }
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if chart_choice.end(local, depth) {
                    return Ok(());
                }
                if let Some(current) = anchor.as_mut() {
                    if current
                        .duotone_color
                        .is_some_and(|(color_depth, _)| color_depth == depth)
                        && let Some((_, color)) = current.duotone_color.take()
                    {
                        current.duotone_colors.push(color);
                    }
                    if current.duotone_depth == Some(depth) && local == "duotone" {
                        current.duotone_depth = None;
                        if let [first, second, ..] = current.duotone_colors.as_slice() {
                            current.image_adjustment.duotone = Some([*first, *second]);
                        }
                    }
                    if current.image_blip_depth == Some(depth) && local == "blip" {
                        current.image_blip_depth = None;
                    }
                }
                if color
                    .as_ref()
                    .is_some_and(|active_color| active_color.depth == depth)
                    && let Some(completed) = color.take()
                {
                    if let Some(active) = shape.as_mut() {
                        apply_drawing_color(active, completed);
                    }
                }
                if let Some(active) = shape.as_mut() {
                    if active
                        .image_fill
                        .as_ref()
                        .and_then(|image_fill| image_fill.parse_depth)
                        == Some(depth)
                        && local == "blipFill"
                    {
                        if let Some(image_fill) = active.image_fill.as_mut() {
                            image_fill.parse_depth = None;
                        }
                    }
                    if active.text_depth == Some(depth) && local == "t" {
                        active.text_depth = None;
                    }
                    if active.paragraph_depth == Some(depth) && local == "p" {
                        active.paragraph_depth = None;
                        if !active.text.is_empty() && !active.text.ends_with('\n') {
                            active.text.push('\n');
                        }
                    }
                    if active.text_line_depth == Some(depth) && local == "ln" {
                        active.text_line_depth = None;
                    } else if active.line_depth == Some(depth) && local == "ln" {
                        active.line_depth = None;
                    }
                    if active.camera_3d_depth == Some(depth) && local == "camera" {
                        active.camera_3d_depth = None;
                    } else if active.light_3d_depth == Some(depth) && local == "lightRig" {
                        active.light_3d_depth = None;
                    } else if active.scene_3d_depth == Some(depth) && local == "scene3d" {
                        active.scene_3d_depth = None;
                    }
                    if active.effect_depth == Some(depth) && local == "effectLst" {
                        active.effect_depth = None;
                    }
                    if active.shadow_depth == Some(depth) && local == "outerShdw" {
                        active.shadow_depth = None;
                    }
                    if active.style_line_depth == Some(depth) && local == "lnRef" {
                        active.style_line_depth = None;
                    }
                    if active.style_fill_depth == Some(depth) && local == "fillRef" {
                        active.style_fill_depth = None;
                    }
                    if active.hidden_style_depth == Some(depth)
                        && matches!(local, "hiddenFill" | "hiddenLine")
                    {
                        active.hidden_style_depth = None;
                    }
                    if active.transform_depth == Some(depth) && local == "xfrm" {
                        active.transform_depth = None;
                    }
                    if active.properties_depth == Some(depth) && local == "spPr" {
                        active.properties_depth = None;
                    }
                    if active.text_body_depth == Some(depth) && local == "txBody" {
                        active.text_body_depth = None;
                    }
                    if let Some(capture) = active.text_stroke_fill_capture.as_mut() {
                        capture.end(local);
                    }
                    if active
                        .text_stroke_fill_capture
                        .as_ref()
                        .is_some_and(|capture| capture.closes_at(depth) && local == "gradFill")
                    {
                        active.text_stroke_fill = active
                            .text_stroke_fill_capture
                            .take()
                            .expect("checked text stroke fill")
                            .finish(package, part)?;
                    }
                }
                if matches!(local, "sp" | "cxnSp")
                    && shape.as_ref().is_some_and(|active| active.depth == depth)
                {
                    let completed = shape.take().expect("checked worksheet shape");
                    let current = anchor
                        .as_ref()
                        .ok_or_else(|| format_error(part, "worksheet shape lost its anchor"))?;
                    if let Some(bounds) = worksheet_shape_bounds(
                        &completed,
                        &groups,
                        current,
                        column_metrics,
                        row_metrics,
                    ) {
                        let drawing_id = completed.drawing_id.unwrap_or_else(|| {
                            u32::try_from(drawings.len() + 1).unwrap_or(u32::MAX)
                        });
                        let geometry = completed
                            .preset_geometry
                            .as_deref()
                            .and_then(|preset| {
                                super::pptx::drawingml_preset_geometry(
                                    preset,
                                    bounds,
                                    &HashMap::new(),
                                )
                            })
                            .map_or_else(|| completed.geometry.clone(), |(geometry, _)| geometry);
                        let (shadow, outer_shadow) = match completed.outer_shadow {
                            Some(effect) if drawingml_outer_shadow_is_identity(&effect) => {
                                (Some(effect.shadow), None)
                            }
                            effect @ Some(_) => (None, effect),
                            None => (completed.shadow, None),
                        };
                        drawings.push(WorksheetDrawing {
                            bounds,
                            drawing_id,
                            payload: DrawingPayload::Shape(Box::new(WorksheetShape {
                                geometry,
                                fill: completed.fill,
                                image_fill: completed.image_fill,
                                stroke: if !completed.text.is_empty()
                                    && !completed.shape_stroke_authored
                                {
                                    Paint::None
                                } else {
                                    completed.stroke
                                },
                                stroke_width: completed.stroke_width,
                                dash: drawingml_dash_lengths(
                                    &completed.dash_pattern,
                                    completed.stroke_width,
                                ),
                                head_arrow: completed.head_arrow,
                                tail_arrow: completed.tail_arrow,
                                connector_preset: completed.connector_preset,
                                flip_h: completed.transform.flip_h,
                                flip_v: completed.transform.flip_v,
                                rotation_degrees: completed.transform.rotation_degrees,
                                shadow,
                                outer_shadow,
                                three_d: completed.three_d,
                                text: completed.text.trim_end_matches('\n').to_owned(),
                                font_family: completed.font_family,
                                font_size: completed.font_size,
                                font_color: completed.font_color,
                                text_stroke_color: completed.text_stroke_color,
                                text_stroke_fill: completed.text_stroke_fill,
                                text_stroke_width: completed.text_stroke_width,
                                bold: completed.bold,
                                italic: completed.italic,
                                align: completed.align,
                                vertical_align: completed.vertical_align,
                                wrap: completed.wrap,
                            })),
                        });
                    }
                }
                if let Some(group) = groups.last_mut()
                    && group.transform_depth == Some(depth)
                    && local == "xfrm"
                {
                    group.transform_depth = None;
                }
                if let Some(group) = groups.last_mut() {
                    if group
                        .image_fill
                        .as_ref()
                        .and_then(|image_fill| image_fill.parse_depth)
                        == Some(depth)
                        && local == "blipFill"
                    {
                        if let Some(image_fill) = group.image_fill.as_mut() {
                            image_fill.parse_depth = None;
                        }
                    }
                    if group.properties_depth == Some(depth) && local == "grpSpPr" {
                        group.properties_depth = None;
                    }
                }
                if local == "grpSp" && groups.last().is_some_and(|group| group.depth == depth) {
                    groups.pop();
                }
                if let Some(current) = anchor.as_mut() {
                    current.image_effects.end(depth);
                    current.image_geometry.end(local, depth);
                    if current
                        .field
                        .as_ref()
                        .is_some_and(|(field_depth, _, _)| *field_depth == depth)
                    {
                        let (_, field, value) = current.field.take().ok_or_else(|| {
                            format_error(part, "worksheet drawing marker state was lost")
                        })?;
                        let value = value.trim();
                        match field {
                            DrawingMarkerField::FromColumn => {
                                current.from.column = parse_u32(value, part, "anchor column")?
                            }
                            DrawingMarkerField::FromColumnOffset => {
                                current.from.column_offset =
                                    parse_i64(value, part, "anchor column offset")?
                            }
                            DrawingMarkerField::FromRow => {
                                current.from.row = parse_u32(value, part, "anchor row")?
                            }
                            DrawingMarkerField::FromRowOffset => {
                                current.from.row_offset =
                                    parse_i64(value, part, "anchor row offset")?
                            }
                            DrawingMarkerField::ToColumn => {
                                current.to.column = parse_u32(value, part, "anchor column")?
                            }
                            DrawingMarkerField::ToColumnOffset => {
                                current.to.column_offset =
                                    parse_i64(value, part, "anchor column offset")?
                            }
                            DrawingMarkerField::ToRow => {
                                current.to.row = parse_u32(value, part, "anchor row")?
                            }
                            DrawingMarkerField::ToRowOffset => {
                                current.to.row_offset = parse_i64(value, part, "anchor row offset")?
                            }
                        }
                    }
                    if matches!(local, "from" | "to")
                        && current
                            .marker_depth
                            .is_some_and(|(marker_depth, _)| marker_depth == depth)
                    {
                        current.marker_depth = None;
                    }
                }
                if matches!(local, "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor")
                    && anchor
                        .as_ref()
                        .is_some_and(|current| current.depth == depth)
                {
                    let mut current = anchor.take().ok_or_else(|| {
                        format_error(part, "worksheet drawing anchor state was lost")
                    })?;
                    if (current.image_relationship_id.is_some()
                        || current.chart_relationship_id.is_some()
                        || current.diagram_relationship_id.is_some())
                        && let Some(bounds) =
                            drawing_anchor_bounds(&current, column_metrics, row_metrics)
                    {
                        let effects = current.image_effects.finish();
                        let image_geometry = current.image_geometry.geometry(bounds);
                        let payload = current
                            .image_relationship_id
                            .map(|relationship_id| {
                                DrawingPayload::Image(Box::new(WorksheetImage {
                                    relationship_id,
                                    crop: current.image_crop,
                                    adjustment: current.image_adjustment,
                                    geometry: image_geometry,
                                    effects,
                                    rotation_degrees: current.image_rotation_degrees,
                                    flip_h: current.image_flip_h,
                                    flip_v: current.image_flip_v,
                                }))
                            })
                            .or_else(|| current.chart_relationship_id.map(DrawingPayload::Chart))
                            .or_else(|| {
                                current.diagram_relationship_id.map(DrawingPayload::Diagram)
                            })
                            .ok_or_else(|| {
                                format_error(part, "worksheet drawing payload state was lost")
                            })?;
                        if drawings.len() >= package.limits().max_document_objects {
                            return Err(object_limit_error(
                                part,
                                "worksheet drawing count exceeds the configured object limit",
                            ));
                        }
                        drawings.push(WorksheetDrawing {
                            bounds,
                            drawing_id: current.drawing_id.unwrap_or_else(|| {
                                u32::try_from(drawings.len() + 1).unwrap_or(u32::MAX)
                            }),
                            payload,
                        });
                    }
                }
            }
            XmlEvent::Text(text) => {
                if chart_choice.skipping() {
                    return Ok(());
                }
                if let Some((_, _, value)) =
                    anchor.as_mut().and_then(|anchor| anchor.field.as_mut())
                {
                    value.push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                } else if let Some(active) =
                    shape.as_mut().filter(|active| active.text_depth.is_some())
                {
                    active
                        .text
                        .push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                }
            }
            XmlEvent::Cdata(text) => {
                if chart_choice.skipping() {
                    return Ok(());
                }
                if let Some((_, _, value)) =
                    anchor.as_mut().and_then(|anchor| anchor.field.as_mut())
                {
                    value.push_str(text);
                } else if let Some(active) =
                    shape.as_mut().filter(|active| active.text_depth.is_some())
                {
                    active.text.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(drawings)
}

fn parse_drawing_percentage(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<f32, Diagnostic> {
    let value = parse_i64(
        &required_attribute(attributes, "val", part)?,
        part,
        "drawing percentage",
    )? as f32
        / 100_000.0;
    if !value.is_finite() {
        return Err(format_error(part, "drawing percentage is not finite"));
    }
    Ok(value)
}

fn drawing_color_value(
    kind: &str,
    value: &str,
    theme_colors: &ThemeColors,
    part: &str,
) -> Result<u32, Diagnostic> {
    match kind {
        "srgbClr" | "sysClr" => parse_theme_rgb(value, part),
        "schemeClr" => theme_color_index(match value {
            "bg1" => "lt1",
            "tx1" => "dk1",
            "bg2" => "lt2",
            "tx2" => "dk2",
            "phClr" => "accent1",
            other => other,
        })
        .and_then(|index| theme_colors.get(index).copied())
        .ok_or_else(|| format_error(part, format!("unknown drawing theme color `{value}`"))),
        "prstClr" => match value {
            "black" => Ok(0x0000_00ff),
            "white" => Ok(0xffff_ffff),
            "red" => Ok(0xff00_00ff),
            "green" => Ok(0x00ff_00ff),
            "blue" => Ok(0x0000_ffff),
            "yellow" => Ok(0xffff_00ff),
            "gray" | "grey" => Ok(0x8080_80ff),
            _ => Ok(0x0000_00ff),
        },
        _ => Err(format_error(part, "unsupported drawing color kind")),
    }
}

fn apply_drawing_color(shape: &mut WorksheetShapeState, color: DrawingColorState) {
    let transformed = transform_luminance(
        color.color,
        color.luminance_modulation.max(0.0),
        color.luminance_offset,
    );
    let alpha = (color.alpha.clamp(0.0, 1.0) * 255.0).round() as u32;
    let transformed = (transformed & 0xffff_ff00) | alpha;
    match color.target {
        DrawingColorTarget::Fill => {
            shape.fill = Paint::Solid(transformed);
            shape.fill_explicit |= color.explicit;
        }
        DrawingColorTarget::Stroke => {
            shape.stroke = Paint::Solid(transformed);
            shape.stroke_explicit |= color.explicit;
        }
        DrawingColorTarget::TextStroke => shape.text_stroke_color = transformed,
        DrawingColorTarget::Text => shape.font_color = transformed,
        DrawingColorTarget::Shadow => {
            if let Some(effect) = shape.outer_shadow.as_mut() {
                effect.shadow.color = transformed;
            } else if let Some(shadow) = shape.shadow.as_mut() {
                shadow.color = transformed;
            }
        }
    }
}

fn drawing_dash_pattern(preset: Option<&str>) -> Vec<f32> {
    DrawingMlDashPattern::from_attribute(preset)
        .lengths()
        .to_vec()
}

fn drawing_shape_geometry(preset: Option<&str>, connector: bool) -> Geometry {
    if connector {
        return Geometry::Line;
    }
    match preset {
        Some("ellipse") => Geometry::Ellipse,
        Some("roundRect")
        | Some("round1Rect")
        | Some("round2SameRect")
        | Some("round2DiagRect") => Geometry::RoundedRectangle {
            radius_x: 8.0,
            radius_y: 8.0,
        },
        Some("line") | Some("straightConnector1") => Geometry::Line,
        _ => Geometry::Rectangle,
    }
}

fn apply_group_transform(bounds: &mut [f64; 4], transform: DrawingTransform) {
    if transform.width <= 0
        || transform.height <= 0
        || transform.child_width <= 0
        || transform.child_height <= 0
    {
        return;
    }
    let scale_x = transform.width as f64 / transform.child_width as f64;
    let scale_y = transform.height as f64 / transform.child_height as f64;
    bounds[0] = transform.x as f64 + (bounds[0] - transform.child_x as f64) * scale_x;
    bounds[1] = transform.y as f64 + (bounds[1] - transform.child_y as f64) * scale_y;
    bounds[2] *= scale_x;
    bounds[3] *= scale_y;
}

fn worksheet_shape_bounds(
    shape: &WorksheetShapeState,
    groups: &[DrawingGroupState],
    anchor: &DrawingAnchorState,
    column_metrics: &AxisMetrics,
    row_metrics: &AxisMetrics,
) -> Option<Rect> {
    if shape.transform.width > 0 && shape.transform.height > 0 {
        let mut emu = [
            shape.transform.x as f64,
            shape.transform.y as f64,
            shape.transform.width as f64,
            shape.transform.height as f64,
        ];
        for group in groups.iter().rev() {
            apply_group_transform(&mut emu, group.transform);
        }
        let mut bounds = Rect {
            x: (emu[0] / EMU_PER_CSS_PIXEL as f64) as f32,
            y: (emu[1] / EMU_PER_CSS_PIXEL as f64) as f32,
            width: (emu[2] / EMU_PER_CSS_PIXEL as f64) as f32,
            height: (emu[3] / EMU_PER_CSS_PIXEL as f64) as f32,
        };
        if shape.is_connector && shape.transform.rotation_degrees != 0.0 {
            let radians = shape.transform.rotation_degrees.to_radians();
            let rotated_width =
                bounds.width * radians.cos().abs() + bounds.height * radians.sin().abs();
            let rotated_height =
                bounds.width * radians.sin().abs() + bounds.height * radians.cos().abs();
            bounds.x += (bounds.width - rotated_width) / 2.0;
            bounds.y += (bounds.height - rotated_height) / 2.0;
            bounds.width = rotated_width;
            bounds.height = rotated_height;
        }
        if bounds.is_valid() && bounds.width > 0.0 && bounds.height > 0.0 {
            return Some(bounds);
        }
    }
    drawing_anchor_bounds(anchor, column_metrics, row_metrics)
}

fn worksheet_connector_geometry(shape: &WorksheetShape, bounds: Rect) -> Geometry {
    let quarter_turn = (shape.rotation_degrees.abs() % 180.0 - 90.0).abs() < 0.01;
    let (source_width, source_height) = if quarter_turn {
        (bounds.height, bounds.width)
    } else {
        (bounds.width, bounds.height)
    };
    let mut points = if shape
        .connector_preset
        .as_deref()
        .is_some_and(|preset| preset.starts_with("bentConnector"))
    {
        vec![
            (0.0, 0.0),
            (source_width / 2.0, 0.0),
            (source_width / 2.0, source_height),
            (source_width, source_height),
        ]
    } else {
        vec![(0.0, 0.0), (source_width, source_height)]
    };
    for (x, y) in &mut points {
        if shape.flip_h {
            *x = source_width - *x;
        }
        if shape.flip_v {
            *y = source_height - *y;
        }
    }
    if shape.rotation_degrees != 0.0 {
        let radians = shape.rotation_degrees.to_radians();
        let cosine = radians.cos();
        let sine = radians.sin();
        let center_x = source_width / 2.0;
        let center_y = source_height / 2.0;
        for (x, y) in &mut points {
            let delta_x = *x - center_x;
            let delta_y = *y - center_y;
            *x = delta_x * cosine - delta_y * sine;
            *y = delta_x * sine + delta_y * cosine;
        }
        let minimum_x = points
            .iter()
            .map(|point| point.0)
            .fold(f32::INFINITY, f32::min);
        let minimum_y = points
            .iter()
            .map(|point| point.1)
            .fold(f32::INFINITY, f32::min);
        for (x, y) in &mut points {
            *x -= minimum_x;
            *y -= minimum_y;
        }
    }
    let mut commands = Vec::with_capacity(points.len() + 6);
    if let Some((x, y)) = points.first().copied() {
        commands.push(PathCommand::MoveTo { x, y });
        commands.extend(
            points
                .iter()
                .skip(1)
                .map(|(x, y)| PathCommand::LineTo { x: *x, y: *y }),
        );
    }
    if shape.head_arrow && points.len() >= 2 {
        append_worksheet_arrow(&mut commands, points[0], points[1], shape.stroke_width);
    }
    if shape.tail_arrow && points.len() >= 2 {
        append_worksheet_arrow(
            &mut commands,
            points[points.len() - 1],
            points[points.len() - 2],
            shape.stroke_width,
        );
    }
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    }
}

fn worksheet_drawing_transform(
    rotation_degrees: f32,
    flip_h: bool,
    flip_v: bool,
    bounds: Rect,
) -> AffineTransform {
    super::drawingml::drawingml_shape_transform(bounds, rotation_degrees, flip_h, flip_v)
}

fn append_worksheet_arrow(
    commands: &mut Vec<PathCommand>,
    tip: (f32, f32),
    adjacent: (f32, f32),
    stroke_width: f32,
) {
    let delta_x = adjacent.0 - tip.0;
    let delta_y = adjacent.1 - tip.1;
    let length = delta_x.hypot(delta_y);
    if length <= f32::EPSILON {
        return;
    }
    let direction_x = delta_x / length;
    let direction_y = delta_y / length;
    let arrow_length = (stroke_width.max(8.0 / 3.0) * 3.0).min(length);
    let half_width = stroke_width.max(8.0 / 3.0) * 1.5;
    let base_x = tip.0 + direction_x * arrow_length;
    let base_y = tip.1 + direction_y * arrow_length;
    let perpendicular_x = -direction_y * half_width;
    let perpendicular_y = direction_x * half_width;
    commands.extend([
        PathCommand::MoveTo {
            x: base_x + perpendicular_x,
            y: base_y + perpendicular_y,
        },
        PathCommand::LineTo { x: tip.0, y: tip.1 },
        PathCommand::LineTo {
            x: base_x - perpendicular_x,
            y: base_y - perpendicular_y,
        },
    ]);
}

fn parse_worksheet_vml_previews(
    package: &Package<'_>,
    worksheet_part: &str,
    relationship_ids: &[String],
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<HashMap<u32, WorksheetVmlPreview>, Diagnostic> {
    let worksheet_relationships = package.relationships(Some(worksheet_part))?;
    let worksheet_relationships = worksheet_relationships
        .iter()
        .map(|relationship| (relationship.id.as_str(), relationship))
        .collect::<HashMap<_, _>>();
    let mut previews = HashMap::new();
    for relationship_id in relationship_ids {
        let Some(vml_relationship) = worksheet_relationships.get(relationship_id.as_str()) else {
            continue;
        };
        if vml_relationship.external || !vml_relationship.type_uri.ends_with("/vmlDrawing") {
            continue;
        }
        let part = vml_relationship.target.as_str();
        let relationships = package.relationships(Some(part))?;
        let relationships = relationships
            .iter()
            .map(|relationship| (relationship.id.as_str(), relationship))
            .collect::<HashMap<_, _>>();
        let bytes = package.required_part(part)?;
        let mut shape_id = None;
        parse_xml(&bytes, package.limits(), |event| {
            match event {
                XmlEvent::StartElement {
                    name, attributes, ..
                } => match local_name(name) {
                    "shape" => {
                        shape_id = optional_attribute(&attributes, "spid", part)?
                            .or(optional_attribute(&attributes, "id", part)?)
                            .and_then(|id| id.rsplit_once("_s").map(|(_, id)| id.to_owned()))
                            .and_then(|id| id.parse::<u32>().ok());
                    }
                    "imagedata" => {
                        let Some(drawing_id) = shape_id else {
                            return Ok(());
                        };
                        let Some(relationship_id) = optional_attribute(&attributes, "relid", part)?
                        else {
                            return Ok(());
                        };
                        let Some(relationship) = relationships.get(relationship_id.as_str()) else {
                            return Ok(());
                        };
                        if relationship.external || !relationship.type_uri.ends_with("/image") {
                            return Ok(());
                        }
                        previews
                            .entry(drawing_id)
                            .or_insert_with(|| WorksheetVmlPreview {
                                part: part.to_owned(),
                                relationship: (*relationship).clone(),
                            });
                    }
                    _ => {}
                },
                XmlEvent::EndElement { name } if local_name(name) == "shape" => shape_id = None,
                _ => {}
            }
            Ok(())
        })
        .map_err(|error| with_part(error, part))?;
    }
    if !relationship_ids.is_empty() && previews.is_empty() {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Omitted,
                "XLSX legacy drawing contains no renderable image previews",
            )
            .in_part(worksheet_part),
        );
    }
    Ok(previews)
}

#[allow(clippy::too_many_arguments)]
fn push_worksheet_image(
    package: &Package<'_>,
    source_part: &str,
    relationship: &Relationship,
    bounds: Rect,
    crop: ImageCrop,
    adjustment: ImageAdjustment,
    geometry: Option<Geometry>,
    mut effects: DrawingMlPictureEffects,
    transform: Option<AffineTransform>,
    drawing_id: u32,
    sheet_name: &str,
    unit_index: u32,
    content_types: &ContentTypes,
    state: &mut XlsxParseState,
) -> Result<bool, Diagnostic> {
    let Some((media_type, bytes)) =
        materialize_worksheet_image(package, source_part, relationship, content_types, state)?
    else {
        return Ok(false);
    };
    let background_fill = effects
        .background_fill
        .take()
        .filter(|fill| !matches!(fill, Paint::None));
    if state
        .objects
        .len()
        .saturating_add(1 + usize::from(background_fill.is_some()))
        > package.limits().max_document_objects
    {
        return Err(object_limit_error(
            source_part,
            "worksheet images exceed the configured object limit",
        ));
    }
    let parent = if let Some(fill) = background_fill {
        let numeric_id = u32::try_from(state.objects.len())
            .map_err(|_| format_error(source_part, "object count exceeds supported range"))?;
        let stable_id = format!("object:{numeric_id}");
        let visual = Visual::PaintedShape {
            geometry: geometry.clone().unwrap_or(Geometry::Rectangle),
            fill,
            stroke: Paint::None,
            stroke_width: 0.0,
        };
        let visual = if let Some(transform) = transform {
            Visual::Layer {
                transform,
                opacity: 1.0,
                blend_mode: BlendMode::Normal,
                visual: Box::new(visual),
            }
        } else {
            visual
        };
        state.objects.push(Object {
            numeric_id,
            parent_numeric_id: None,
            stable_id: stable_id.clone(),
            parent_stable_id: None,
            kind: ObjectKind::Shape,
            unit_index,
            bounds,
            z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
            text: None,
            source: SourceRef {
                part: source_part.to_owned(),
                mapping: MappingQuality::Exact,
                locator: SourceLocator::Xlsx {
                    kind: "drawing",
                    sheet_name: sheet_name.to_owned(),
                    address: None,
                    formula: None,
                    drawing_id: Some(drawing_id),
                },
            },
            visual,
        });
        Some((numeric_id, stable_id))
    } else {
        None
    };
    let numeric_id = u32::try_from(state.objects.len())
        .map_err(|_| format_error(source_part, "object count exceeds supported range"))?;
    state.objects.push(Object {
        numeric_id,
        parent_numeric_id: parent.as_ref().map(|(numeric_id, _)| *numeric_id),
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: parent.map(|(_, stable_id)| stable_id),
        kind: ObjectKind::Image,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: source_part.to_owned(),
            mapping: MappingQuality::Exact,
            locator: SourceLocator::Xlsx {
                kind: "drawing",
                sheet_name: sheet_name.to_owned(),
                address: None,
                formula: None,
                drawing_id: Some(drawing_id),
            },
        },
        visual: {
            let image = Visual::Image {
                media_type,
                bytes,
                crop,
            };
            let image = if adjustment == ImageAdjustment::default() {
                image
            } else {
                Visual::ImageAdjustment {
                    adjustment,
                    visual: Box::new(image),
                }
            };
            let image = effects.wrap(image, geometry);
            if let Some(transform) = transform {
                Visual::Layer {
                    transform,
                    opacity: 1.0,
                    blend_mode: BlendMode::Normal,
                    visual: Box::new(image),
                }
            } else {
                image
            }
        },
    });
    Ok(true)
}

fn materialize_worksheet_image(
    package: &Package<'_>,
    source_part: &str,
    relationship: &Relationship,
    content_types: &ContentTypes,
    state: &mut XlsxParseState,
) -> Result<Option<(String, Vec<u8>)>, Diagnostic> {
    if relationship.external {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ExternalResourceBlocked,
                Phase::Security,
                Fidelity::Blocked,
                "external worksheet image relationship was blocked",
            )
            .in_part(source_part),
        );
        return Ok(None);
    }
    if !relationship.type_uri.ends_with("/image") {
        return Err(format_error(
            source_part,
            format!("relationship {} is not an image", relationship.id),
        ));
    }
    let bytes = package.required_part(&relationship.target)?;
    let declared_content_type = content_types.for_part(&relationship.target);
    let media_type = match declared_content_type.map_or_else(
        || office_image_media_type(&relationship.target, &bytes),
        |media_type| office_image_media_type_from_mime(media_type, &bytes),
    ) {
        Ok(media_type) => media_type,
        Err(error) => {
            state
                .diagnostics
                .push(unsupported_xlsx_image_diagnostic(source_part, error));
            return Ok(None);
        }
    };
    reserve_materialized_image_bytes(
        &mut state.materialized_image_bytes,
        bytes.len(),
        state.materialized_image_limit,
        source_part,
    )?;
    Ok(Some((media_type.to_owned(), bytes.into_vec())))
}

#[allow(clippy::too_many_arguments)]
fn push_worksheet_drawings(
    package: &Package<'_>,
    worksheet_part: &str,
    relationship_ids: &[String],
    legacy_relationship_ids: &[String],
    column_metrics: &AxisMetrics,
    row_metrics: &AxisMetrics,
    sheet_name: &str,
    unit_index: u32,
    content_types: &ContentTypes,
    theme_colors: &ThemeColors,
    theme_line_styles: &DrawingMlThemeLineStyles,
    chart_theme_colors: Option<&ThemeColors>,
    chart_named_data: &HashMap<String, super::drawingml::ChartSourceData>,
    state: &mut XlsxParseState,
) -> Result<(f32, f32), Diagnostic> {
    let mut extent = (0.0_f32, 0.0_f32);
    if relationship_ids.is_empty() && legacy_relationship_ids.is_empty() {
        return Ok(extent);
    }
    let previews = parse_worksheet_vml_previews(
        package,
        worksheet_part,
        legacy_relationship_ids,
        &mut state.diagnostics,
    )?;
    let worksheet_relationships = package.relationships(Some(worksheet_part))?;
    let worksheet_relationships = worksheet_relationships
        .iter()
        .map(|relationship| (relationship.id.as_str(), relationship))
        .collect::<HashMap<_, _>>();
    for relationship_id in relationship_ids {
        let Some(relationship) = worksheet_relationships.get(relationship_id.as_str()) else {
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Parse,
                    Fidelity::Omitted,
                    format!(
                        "worksheet drawing relationship {relationship_id} is missing and was omitted"
                    ),
                )
                .in_part(worksheet_part),
            );
            continue;
        };
        if relationship.external {
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::ExternalResourceBlocked,
                    Phase::Security,
                    Fidelity::Blocked,
                    "external worksheet drawing relationship was blocked",
                )
                .in_part(worksheet_part),
            );
            continue;
        }
        if !relationship.type_uri.ends_with("/drawing") {
            return Err(format_error(
                worksheet_part,
                format!("relationship {relationship_id} is not a worksheet drawing"),
            ));
        }
        let drawing_part = relationship.target.as_str();
        let drawings = parse_worksheet_drawing(
            package,
            drawing_part,
            column_metrics,
            row_metrics,
            theme_colors,
            theme_line_styles,
        )?;
        let drawing_relationships = package.relationships(Some(drawing_part))?;
        let drawing_relationships = drawing_relationships
            .iter()
            .map(|relationship| (relationship.id.as_str(), relationship))
            .collect::<HashMap<_, _>>();
        for drawing in drawings {
            let bounds = drawing.bounds;
            let drawing_extent = worksheet_drawing_visual_extent(&drawing);
            if let Some(preview) = previews.get(&drawing.drawing_id) {
                if push_worksheet_image(
                    package,
                    &preview.part,
                    &preview.relationship,
                    bounds,
                    ImageCrop::default(),
                    ImageAdjustment::default(),
                    None,
                    DrawingMlPictureEffects::default(),
                    None,
                    drawing.drawing_id,
                    sheet_name,
                    unit_index,
                    content_types,
                    state,
                )? {
                    extent.0 = extent.0.max(drawing_extent.0);
                    extent.1 = extent.1.max(drawing_extent.1);
                }
                continue;
            }
            match drawing.payload {
                DrawingPayload::Image(image) => {
                    let image = *image;
                    let image_relationship = drawing_relationships
                        .get(image.relationship_id.as_str())
                        .ok_or_else(|| {
                            format_error(
                                drawing_part,
                                format!(
                                    "drawing image relationship {} does not exist",
                                    image.relationship_id
                                ),
                            )
                        })?;
                    if push_worksheet_image(
                        package,
                        drawing_part,
                        image_relationship,
                        bounds,
                        image.crop,
                        image.adjustment,
                        image.geometry,
                        image.effects,
                        (image.rotation_degrees != 0.0 || image.flip_h || image.flip_v).then(
                            || {
                                worksheet_drawing_transform(
                                    image.rotation_degrees,
                                    image.flip_h,
                                    image.flip_v,
                                    bounds,
                                )
                            },
                        ),
                        drawing.drawing_id,
                        sheet_name,
                        unit_index,
                        content_types,
                        state,
                    )? {
                        extent.0 = extent.0.max(drawing_extent.0);
                        extent.1 = extent.1.max(drawing_extent.1);
                    }
                }
                DrawingPayload::Shape(shape) => {
                    if state.objects.len() >= package.limits().max_document_objects {
                        return Err(object_limit_error(
                            drawing_part,
                            "worksheet shapes exceed the configured object limit",
                        ));
                    }
                    let numeric_id = u32::try_from(state.objects.len()).map_err(|_| {
                        format_error(drawing_part, "object count exceeds supported range")
                    })?;
                    let mut shape = *shape;
                    if let Some(image_fill) = shape.image_fill.take()
                        && let Some(relationship_id) = image_fill.relationship_id
                    {
                        let relationship = drawing_relationships
                            .get(relationship_id.as_str())
                            .ok_or_else(|| {
                                format_error(
                                    drawing_part,
                                    format!(
                                        "shape image relationship {relationship_id} does not exist"
                                    ),
                                )
                            })?;
                        if let Some((media_type, bytes)) = materialize_worksheet_image(
                            package,
                            drawing_part,
                            relationship,
                            content_types,
                            state,
                        )? {
                            shape.fill = Paint::Image {
                                mapping: (image_fill.tile
                                    || image_fill.mapping
                                        != crate::model::ImageFillMapping::default())
                                .then(|| Box::new(image_fill.mapping)),
                                media_type,
                                bytes,
                                crop: image_fill.crop,
                                tile: image_fill.tile,
                                tile_width: None,
                                tile_height: None,
                            };
                        }
                    }
                    let has_text = !shape.text.is_empty();
                    if shape.connector_preset.is_some() {
                        shape.geometry = worksheet_connector_geometry(&shape, bounds);
                    }
                    let shape_transform = worksheet_drawing_transform(
                        shape.rotation_degrees,
                        shape.flip_h,
                        shape.flip_v,
                        bounds,
                    );
                    let visual = if has_text {
                        Visual::RichText {
                            geometry: shape.geometry.clone(),
                            fill: shape.fill,
                            stroke: shape.stroke,
                            stroke_width: shape.stroke_width,
                            align: shape.align,
                            line_height: 0.0,
                            runs: vec![TextRun {
                                paint: None,
                                east_asian_line_breaks: true,
                                text: shape.text.clone(),
                                font_family: shape.font_family,
                                font_size: shape.font_size,
                                color: shape.font_color,
                                bold: shape.bold,
                                italic: shape.italic,
                                underline: false,
                                strikethrough: false,
                                highlight: 0,
                                baseline_shift: 0.0,
                                letter_spacing: 0.0,
                                horizontal_scale: 1.0,
                            }],
                        }
                    } else {
                        Visual::PaintedShape {
                            geometry: shape.geometry.clone(),
                            fill: shape.fill,
                            stroke: shape.stroke,
                            stroke_width: shape.stroke_width,
                        }
                    };
                    let visual = if has_text {
                        Visual::TextLayout {
                            layout: TextLayout {
                                vertical_align: shape.vertical_align,
                                wrap: shape.wrap,
                                inset_left: 4.0,
                                inset_right: 4.0,
                                inset_top: 2.0,
                                inset_bottom: 2.0,
                                text_stroke_color: shape.text_stroke_color,
                                text_stroke_paint: shape
                                    .text_stroke_fill
                                    .map_or(Paint::None, |fill| fill.paint(bounds)),
                                text_stroke_width: shape.text_stroke_width,
                                ..TextLayout::default()
                            },
                            visual: Box::new(visual),
                        }
                    } else {
                        visual
                    };
                    let visual = if shape.outer_shadow.is_some() || shape.three_d.is_some() {
                        Visual::AdvancedEffect {
                            outer_shadow: shape.outer_shadow,
                            inner_shadow: None,
                            glow: None,
                            reflection: None,
                            soft_edge: None,
                            three_d: shape.three_d,
                            visual: Box::new(visual),
                        }
                    } else {
                        visual
                    };
                    let visual = if shape.dash.is_empty() {
                        visual
                    } else {
                        Visual::StrokeStyle {
                            style: StrokeStyle {
                                dash: shape.dash,
                                ..StrokeStyle::default()
                            },
                            visual: Box::new(visual),
                        }
                    };
                    let visual = if let Some(shadow) = shape.shadow {
                        Visual::Effect {
                            shadow: Some(shadow),
                            clip: None,
                            visual: Box::new(visual),
                        }
                    } else {
                        visual
                    };
                    let visual = if shape.connector_preset.is_none()
                        && (shape.rotation_degrees != 0.0 || shape.flip_h || shape.flip_v)
                    {
                        Visual::Layer {
                            transform: shape_transform,
                            opacity: 1.0,
                            blend_mode: BlendMode::Normal,
                            visual: Box::new(visual),
                        }
                    } else {
                        visual
                    };
                    state.objects.push(Object {
                        numeric_id,
                        parent_numeric_id: None,
                        stable_id: format!("object:{numeric_id}"),
                        parent_stable_id: None,
                        kind: ObjectKind::Shape,
                        unit_index,
                        bounds,
                        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
                        text: has_text.then_some(shape.text),
                        source: xlsx_drawing_source(drawing_part, sheet_name, drawing.drawing_id),
                        visual,
                    });
                    extent.0 = extent.0.max(drawing_extent.0);
                    extent.1 = extent.1.max(drawing_extent.1);
                }
                DrawingPayload::Chart(chart_relationship_id) => {
                    let chart_relationship = drawing_relationships
                        .get(chart_relationship_id.as_str())
                        .ok_or_else(|| {
                            format_error(
                                drawing_part,
                                format!(
                                    "drawing chart relationship {chart_relationship_id} does not exist"
                                ),
                            )
                        })?;
                    if chart_relationship.external {
                        state.diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::ExternalResourceBlocked,
                                Phase::Security,
                                Fidelity::Blocked,
                                "external worksheet chart relationship was blocked",
                            )
                            .in_part(drawing_part),
                        );
                        continue;
                    }
                    if !chart_relationship.type_uri.ends_with("/chart")
                        && !chart_relationship.type_uri.ends_with("/chartEx")
                    {
                        return Err(format_error(
                            drawing_part,
                            format!("relationship {chart_relationship_id} is not a chart"),
                        ));
                    }
                    let chart = if chart_relationship.type_uri.ends_with("/chartEx") {
                        parse_chart_ex_with_data(
                            package,
                            &chart_relationship.target,
                            |value| {
                                chart_theme_colors.and_then(|colors| {
                                    drawing_color_value("schemeClr", value, colors, drawing_part)
                                        .ok()
                                })
                            },
                            |name| chart_named_data.get(name).cloned(),
                        )?
                    } else {
                        match chart_theme_colors {
                            Some(colors) => {
                                parse_chart(package, &chart_relationship.target, |value| {
                                    drawing_color_value("schemeClr", value, colors, drawing_part)
                                        .ok()
                                })?
                            }
                            None => parse_basic_chart(package, &chart_relationship.target)?,
                        }
                    };
                    if let Some(chart) = chart {
                        push_xlsx_chart(
                            chart,
                            bounds,
                            drawing.drawing_id,
                            sheet_name,
                            unit_index,
                            package.limits().max_document_objects,
                            state,
                        )?;
                        extent.0 = extent.0.max(drawing_extent.0);
                        extent.1 = extent.1.max(drawing_extent.1);
                    } else {
                        state.diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::UnsupportedFeature,
                                Phase::Render,
                                Fidelity::Omitted,
                                "XLSX chart has no supported cached series data and was omitted",
                            )
                            .in_part(&chart_relationship.target),
                        );
                    }
                }
                DrawingPayload::Diagram(diagram_relationship_id) => {
                    let diagram_relationship = drawing_relationships
                        .get(diagram_relationship_id.as_str())
                        .ok_or_else(|| {
                            format_error(
                                drawing_part,
                                format!(
                                    "drawing diagram relationship {diagram_relationship_id} does not exist"
                                ),
                            )
                        })?;
                    if diagram_relationship.external {
                        state.diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::ExternalResourceBlocked,
                                Phase::Security,
                                Fidelity::Blocked,
                                "external worksheet SmartArt relationship was blocked",
                            )
                            .in_part(drawing_part),
                        );
                    } else if !diagram_relationship.type_uri.ends_with("/diagramData") {
                        return Err(format_error(
                            drawing_part,
                            format!(
                                "relationship {diagram_relationship_id} is not SmartArt diagram data"
                            ),
                        ));
                    } else {
                        let diagram = parse_diagram(
                            package,
                            &diagram_relationship.target,
                            &drawing_relationships,
                            None,
                            None,
                            |value| {
                                drawing_color_value("schemeClr", value, theme_colors, drawing_part)
                                    .ok()
                            },
                        )?;
                        let has_content =
                            !diagram.nodes.is_empty() || diagram.background_fill.is_some();
                        if has_content {
                            push_xlsx_diagram(
                                &diagram,
                                bounds,
                                drawing.drawing_id,
                                sheet_name,
                                unit_index,
                                &diagram_relationship.target,
                                theme_colors,
                                package.limits().max_document_objects,
                                state,
                            )?;
                            extent.0 = extent.0.max(drawing_extent.0);
                            extent.1 = extent.1.max(drawing_extent.1);
                        }
                        state.diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::UnsupportedFeature,
                                Phase::Render,
                                if has_content {
                                    Fidelity::Approximate
                                } else {
                                    Fidelity::Omitted
                                },
                                if has_content {
                                    "XLSX SmartArt used a deterministic data-model layout"
                                } else {
                                    "XLSX SmartArt has no visible data nodes and was omitted"
                                },
                            )
                            .in_part(&diagram_relationship.target),
                        );
                    }
                }
            }
        }
    }
    Ok(extent)
}

#[allow(clippy::too_many_arguments)]
fn push_xlsx_diagram(
    diagram: &XlsxDiagram,
    bounds: Rect,
    drawing_id: u32,
    sheet_name: &str,
    unit_index: u32,
    source_part: &str,
    theme_colors: &ThemeColors,
    object_limit: usize,
    state: &mut XlsxParseState,
) -> Result<(), Diagnostic> {
    state
        .diagnostics
        .extend(diagram.diagnostics.iter().cloned());
    let nodes = &diagram.nodes;
    let fallback_elements =
        super::drawingml::diagram_fallback_elements(diagram, bounds, theme_colors[4]);
    if state
        .objects
        .len()
        .saturating_add(fallback_elements.as_ref().map_or(nodes.len(), Vec::len))
        .saturating_add(usize::from(diagram.background_fill.is_some()))
        .saturating_add(1)
        > object_limit
    {
        return Err(object_limit_error(
            source_part,
            "SmartArt nodes exceed the configured object limit",
        ));
    }
    let group_id = u32::try_from(state.objects.len())
        .map_err(|_| format_error(source_part, "object count exceeds supported range"))?;
    state.objects.push(Object {
        numeric_id: group_id,
        parent_numeric_id: None,
        stable_id: format!("object:{group_id}"),
        parent_stable_id: None,
        kind: ObjectKind::Group,
        unit_index,
        bounds,
        z: i32::try_from(group_id).unwrap_or(i32::MAX),
        text: None,
        source: xlsx_drawing_source(source_part, sheet_name, drawing_id),
        visual: Visual::None,
    });
    if let Some(fill) = &diagram.background_fill {
        let numeric_id = u32::try_from(state.objects.len())
            .map_err(|_| format_error(source_part, "object count exceeds supported range"))?;
        let visual = Visual::PaintedShape {
            geometry: Geometry::Rectangle,
            fill: fill.paint(bounds),
            stroke: Paint::None,
            stroke_width: 0.0,
        };
        state.objects.push(Object {
            numeric_id,
            parent_numeric_id: Some(group_id),
            stable_id: format!("object:{numeric_id}"),
            parent_stable_id: Some(format!("object:{group_id}")),
            kind: ObjectKind::Shape,
            unit_index,
            bounds,
            z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
            text: None,
            source: xlsx_drawing_source(source_part, sheet_name, drawing_id),
            visual: if let Some(shadow) = diagram.background_shadow {
                Visual::Effect {
                    shadow: Some(shadow),
                    clip: None,
                    visual: Box::new(visual),
                }
            } else {
                visual
            },
        });
    }
    if let Some(elements) = fallback_elements {
        for element in elements {
            let numeric_id = u32::try_from(state.objects.len())
                .map_err(|_| format_error(source_part, "object count exceeds supported range"))?;
            state.objects.push(Object {
                numeric_id,
                parent_numeric_id: Some(group_id),
                stable_id: format!("object:{numeric_id}"),
                parent_stable_id: Some(format!("object:{group_id}")),
                kind: if element.text.is_some() {
                    ObjectKind::TextBox
                } else {
                    ObjectKind::Shape
                },
                unit_index,
                bounds: element.bounds,
                z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
                text: element.text,
                source: xlsx_drawing_source(source_part, sheet_name, drawing_id),
                visual: element.visual,
            });
        }
        return Ok(());
    }
    // ponytail: row layout preserves visible data; parse diagramDrawing when exact SmartArt geometry is required.
    let gap = (bounds.width / (nodes.len() as f32 * 8.0)).clamp(4.0, 12.0);
    let node_width =
        ((bounds.width - gap * (nodes.len() + 1) as f32) / nodes.len() as f32).max(1.0);
    let node_height = (bounds.height * 0.45).max(1.0);
    let y = bounds.y + (bounds.height - node_height) / 2.0;
    for (index, node) in nodes.iter().enumerate() {
        let numeric_id = u32::try_from(state.objects.len())
            .map_err(|_| format_error(source_part, "object count exceeds supported range"))?;
        let node_bounds = Rect {
            x: bounds.x + gap + index as f32 * (node_width + gap),
            y,
            width: node_width,
            height: node_height,
        };
        state.objects.push(Object {
            numeric_id,
            parent_numeric_id: Some(group_id),
            stable_id: format!("object:{numeric_id}"),
            parent_stable_id: Some(format!("object:{group_id}")),
            kind: ObjectKind::Shape,
            unit_index,
            bounds: node_bounds,
            z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
            text: Some(node.text.clone()),
            source: xlsx_drawing_source(source_part, sheet_name, drawing_id),
            visual: Visual::TextLayout {
                layout: TextLayout {
                    vertical_align: TextVerticalAlign::Center,
                    wrap: true,
                    inset_left: 4.0,
                    inset_right: 4.0,
                    inset_top: 2.0,
                    inset_bottom: 2.0,
                    ..TextLayout::default()
                },
                visual: Box::new(Visual::RichText {
                    geometry: drawing_shape_geometry(node.preset_geometry.as_deref(), false),
                    fill: node
                        .fill
                        .as_ref()
                        .map_or(Paint::Solid(theme_colors[4]), |fill| {
                            fill.paint(node_bounds)
                        }),
                    stroke: Paint::Solid(0xffff_ffff),
                    stroke_width: 2.0,
                    align: TextAlign::Center,
                    line_height: 0.0,
                    runs: vec![TextRun {
                        paint: None,
                        east_asian_line_breaks: true,
                        text: node.text.clone(),
                        font_family: "Arial".to_owned(),
                        font_size: (node_height * 0.30).clamp(10.0, 24.0),
                        color: 0xffff_ffff,
                        bold: false,
                        italic: false,
                        underline: false,
                        strikethrough: false,
                        highlight: 0,
                        baseline_shift: 0.0,
                        letter_spacing: 0.0,
                        horizontal_scale: 1.0,
                    }],
                }),
            },
        });
    }
    Ok(())
}

fn unsupported_xlsx_image_diagnostic(part: &str, error: OfficeImageError) -> Diagnostic {
    let message = match error {
        OfficeImageError::UnsupportedFormat => "XLSX image format is unsupported",
        OfficeImageError::SignatureMismatch => "XLSX image bytes do not match their extension",
        OfficeImageError::DisabledByOffice => "XLSX image format is disabled by Microsoft Office",
    };
    Diagnostic::warning(
        DiagnosticCode::UnsupportedFeature,
        Phase::Parse,
        Fidelity::Omitted,
        message,
    )
    .in_part(part)
}

fn push_xlsx_chart(
    chart: XlsxChart,
    bounds: Rect,
    drawing_id: u32,
    sheet_name: &str,
    unit_index: u32,
    object_limit: usize,
    state: &mut XlsxParseState,
) -> Result<(), Diagnostic> {
    let source = xlsx_drawing_source(&chart.source_part, sheet_name, drawing_id);
    super::drawingml::push_spreadsheet_chart(
        chart,
        bounds,
        &source,
        unit_index,
        object_limit,
        &mut state.objects,
        &mut state.diagnostics,
    )
}

#[cfg(test)]
use super::drawingml::{
    chart_text_tight_bounds, spreadsheet_axis_title_presentation as xlsx_axis_title_presentation,
};

fn xlsx_drawing_source(part: &str, sheet_name: &str, drawing_id: u32) -> SourceRef {
    SourceRef {
        part: part.to_owned(),
        mapping: MappingQuality::Derived,
        locator: SourceLocator::Xlsx {
            kind: "drawing",
            sheet_name: sheet_name.to_owned(),
            address: None,
            formula: None,
            drawing_id: Some(drawing_id),
        },
    }
}

fn drawing_anchor_bounds(
    anchor: &DrawingAnchorState,
    column_metrics: &AxisMetrics,
    row_metrics: &AxisMetrics,
) -> Option<Rect> {
    let point = |marker: DrawingMarker| {
        (
            column_metrics.offset(marker.column) + marker.column_offset as f32 / EMU_PER_CSS_PIXEL,
            row_metrics.offset(marker.row) + marker.row_offset as f32 / EMU_PER_CSS_PIXEL,
        )
    };
    let (x, y, width, height) = match anchor.kind {
        DrawingAnchorKind::TwoCell => {
            let (x, y) = point(anchor.from);
            let (end_x, end_y) = point(anchor.to);
            (x, y, end_x - x, end_y - y)
        }
        DrawingAnchorKind::OneCell => {
            let (x, y) = point(anchor.from);
            (
                x,
                y,
                anchor.width.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
                anchor.height.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
            )
        }
        DrawingAnchorKind::Absolute => (
            anchor.x.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
            anchor.y.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
            anchor.width.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
            anchor.height.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        ),
    };
    let bounds = Rect {
        x,
        y,
        width,
        height,
    };
    if !bounds.is_valid() || bounds.width == 0.0 || bounds.height == 0.0 {
        return None;
    }
    Some(bounds)
}

fn parse_table_styles(
    package: &Package<'_>,
    worksheet_part: &str,
    relationship_ids: &[String],
    state: &mut XlsxParseState,
) -> Result<Vec<TableStyle>, Diagnostic> {
    let relationships = package.relationships(Some(worksheet_part))?;
    let mut tables = Vec::new();
    for relationship in relationships.iter().filter(|relationship| {
        relationship_ids.contains(&relationship.id)
            || relationship.type_uri.ends_with("/pivotTable")
    }) {
        if relationship.external
            || !(relationship.type_uri.ends_with("/table")
                || relationship.type_uri.ends_with("/pivotTable"))
        {
            continue;
        }
        let pivot = relationship.type_uri.ends_with("/pivotTable");
        if pivot
            && !state
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("pivot-table"))
        {
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Render,
                    Fidelity::Omitted,
                    "interactive XLSX pivot-table features are not rendered",
                )
                .in_part(worksheet_part),
            );
        }
        let bytes = match package.required_part(&relationship.target) {
            Ok(bytes) => bytes,
            Err(error) => {
                state.diagnostics.push(error);
                continue;
            }
        };
        let mut range = None;
        let mut name = None;
        let mut show_row_stripes = false;
        let mut first_data_row = 1_u32;
        let mut row_field_count = 0_u32;
        let mut rows: Vec<PivotRow> = Vec::new();
        let mut in_row_fields = false;
        let mut in_row_items = false;
        let mut row_item_members = 0_u32;
        if let Err(error) = parse_xml(&bytes, package.limits(), |event| {
            if let XmlEvent::EndElement { name } = &event {
                match local_name(name) {
                    "rowFields" => in_row_fields = false,
                    "rowItems" => in_row_items = false,
                    _ => {}
                }
            }
            if let XmlEvent::StartElement {
                name: element,
                attributes,
                ..
            } = event
            {
                match local_name(element) {
                    "table" if !pivot && range.is_none() => {
                        range = Some(parse_dimension(
                            &required_attribute(&attributes, "ref", &relationship.target)?,
                            &relationship.target,
                        )?);
                    }
                    "rowFields" if pivot => in_row_fields = true,
                    "field" if in_row_fields => row_field_count = row_field_count.saturating_add(1),
                    "rowItems" if pivot => in_row_items = true,
                    "i" if in_row_items => {
                        if rows.len() >= package.limits().max_document_objects {
                            return Err(object_limit_error(
                                &relationship.target,
                                "pivot row items exceed the configured object limit",
                            ));
                        }
                        rows.push(PivotRow {
                            level: parse_optional_u32(&attributes, "r", &relationship.target)?
                                .unwrap_or(0),
                            grand_total: optional_attribute(
                                &attributes,
                                "t",
                                &relationship.target,
                            )?
                            .as_deref()
                                == Some("grand"),
                        });
                        row_item_members = 0;
                    }
                    "x" if in_row_items => {
                        // r repeats prior labels; each x supplies another field level.
                        if row_item_members > 0
                            && let Some(row) = rows.last_mut()
                        {
                            row.level = row.level.saturating_add(1);
                        }
                        row_item_members = row_item_members.saturating_add(1);
                    }
                    "location" if pivot && range.is_none() => {
                        range = Some(parse_dimension(
                            &required_attribute(&attributes, "ref", &relationship.target)?,
                            &relationship.target,
                        )?);
                        first_data_row =
                            parse_optional_u32(&attributes, "firstDataRow", &relationship.target)?
                                .unwrap_or(1);
                    }
                    "tableStyleInfo" | "pivotTableStyleInfo" => {
                        name = optional_attribute(&attributes, "name", &relationship.target)?;
                        show_row_stripes = parse_bool_attribute(
                            &attributes,
                            "showRowStripes",
                            &relationship.target,
                        )?
                        .unwrap_or(false);
                    }
                    _ => {}
                }
            }
            Ok(())
        }) {
            state
                .diagnostics
                .push(with_part(error, &relationship.target));
            continue;
        }
        if let (Some(range), Some(name)) = (range, name) {
            tables.push(TableStyle {
                kind: if pivot {
                    TableStyleKind::Pivot {
                        first_data_row,
                        row_field_count,
                        rows,
                    }
                } else {
                    TableStyleKind::Table
                },
                range,
                name,
                show_row_stripes,
            });
        }
    }
    Ok(tables)
}

fn table_cell_style(
    column: u32,
    row: u32,
    base: &CellStyle,
    tables: &[TableStyle],
    styles: &Styles,
) -> CellStyle {
    let mut style = base.clone();
    let Some(table) = tables.iter().find(|table| {
        column >= table.range.start_column
            && column <= table.range.end_column
            && row >= table.range.start_row
            && row <= table.range.end_row
    }) else {
        return style;
    };
    if let Some(differential) = styles.table_styles.get(&table.name) {
        apply_differential_style(&mut style, differential);
        return style;
    }
    match &table.kind {
        TableStyleKind::Table
            if matches!(
                table.name.as_str(),
                "TableStyleMedium2" | "TableStyleMedium9"
            ) =>
        {
            if row == table.range.start_row {
                style.fill = styles.theme_colors[4];
                style.fill_paint = None;
                style.has_fill = true;
                style.color = 0xffff_ffff;
                style.bold = true;
            } else if table.show_row_stripes && (row - table.range.start_row) % 2 == 1 {
                style.fill = transform_luminance(styles.theme_colors[4], 0.2, 0.8);
                style.fill_paint = None;
                style.has_fill = true;
            }
        }
        TableStyleKind::Pivot {
            first_data_row,
            row_field_count,
            rows,
        } if matches!(
            table.name.as_str(),
            "PivotStyleLight16" | "PivotStyleMedium2"
        ) =>
        {
            let relative_row = row - table.range.start_row;
            let header = relative_row < *first_data_row;
            let item = relative_row
                .checked_sub(*first_data_row)
                .and_then(|index| rows.get(index as usize));
            let grand_total = item.is_some_and(|item| item.grand_total);
            let subheading = item.filter(|item| {
                !item.grand_total && item.level.saturating_add(1) < *row_field_count
            });
            let medium = table.name == "PivotStyleMedium2";
            let accent = styles.theme_colors[4];
            let fill = if medium {
                if header {
                    Some(transform_luminance(accent, 0.75, 0.0))
                } else {
                    subheading.map(|item| {
                        transform_luminance(
                            accent,
                            if item.level == 0 { 0.6 } else { 0.2 },
                            if item.level == 0 { 0.4 } else { 0.8 },
                        )
                    })
                }
            } else if header || grand_total {
                Some(transform_luminance(accent, 0.2, 0.8))
            } else {
                None
            };
            if let Some(fill) = fill {
                style.fill = fill;
                style.fill_paint = None;
                style.has_fill = true;
            }
            if header || grand_total || subheading.is_some() {
                style.bold = true;
            }
            if medium && (header || subheading.is_some_and(|item| item.level == 0)) {
                style.color = 0xffff_ffff;
            }
        }
        _ => {}
    }
    style
}

fn apply_differential_style(style: &mut CellStyle, differential: &DifferentialStyle) {
    if let Some(color) = differential.color {
        style.color = color;
    }
    if let Some(bold) = differential.bold {
        style.bold = bold;
    }
    if let Some(italic) = differential.italic {
        style.italic = italic;
    }
    if let Some(fill) = differential.fill {
        style.fill = fill;
        style.fill_paint = None;
        style.has_fill = true;
    }
    if let Some(fill) = differential.fill_paint.as_ref() {
        style.fill_paint = Some(fill.clone());
        style.has_fill = true;
    }
}

fn conditional_rules_for_cell<'a>(
    blocks: &'a [ConditionalFormatting],
    address: &CellAddress,
) -> Vec<(usize, &'a ConditionalRule)> {
    let mut rules = blocks
        .iter()
        .enumerate()
        .filter(|(_, block)| conditional_block_contains(block, address.column, address.row))
        .flat_map(|(index, block)| block.rules.iter().map(move |rule| (index, rule)))
        .collect::<Vec<_>>();
    rules.sort_by_key(|(_, rule)| rule.priority);
    rules
}

fn conditional_cell_style(
    cell: &CellState,
    base: &CellStyle,
    blocks: &[ConditionalFormatting],
    stats: &[ConditionalStats],
    differentials: &[DifferentialStyle],
    shared_strings: &[SharedString],
) -> CellStyle {
    let numeric_value = cell_numeric_value(cell);
    let mut style = base.clone();
    let mut rules = conditional_rules_for_cell(blocks, &cell.address);
    let mut stop = None;
    for (index, (_, rule)) in rules.iter().enumerate() {
        let applies = match &rule.kind {
            ConditionalRuleKind::CellIs {
                operator, formulas, ..
            } => cell_rule_matches(cell, shared_strings, *operator, formulas),
            ConditionalRuleKind::Calculated { matches, .. } => {
                matches.contains(&(cell.address.column, cell.address.row))
            }
            _ => numeric_value.is_some(),
        };
        if applies && rule.stop_if_true {
            stop = Some(index);
            break;
        }
    }
    if let Some(index) = stop {
        rules.truncate(index + 1);
    }
    for (block_index, rule) in rules.into_iter().rev() {
        let dxf_id = match &rule.kind {
            ConditionalRuleKind::Calculated { dxf_id, matches }
                if matches.contains(&(cell.address.column, cell.address.row)) =>
            {
                Some(*dxf_id)
            }
            ConditionalRuleKind::CellIs {
                dxf_id,
                operator,
                formulas,
            } if cell_rule_matches(cell, shared_strings, *operator, formulas) => Some(*dxf_id),
            ConditionalRuleKind::ColorScale { thresholds, colors } => {
                if let (Some(value), Some(block_stats)) = (numeric_value, stats.get(block_index)) {
                    style.fill = color_scale_color(value, thresholds, colors, block_stats);
                    style.fill_paint = None;
                    style.has_fill = true;
                }
                None
            }
            _ => None,
        };
        if let Some(differential) = dxf_id.and_then(|index| differentials.get(index)) {
            apply_differential_style(&mut style, differential);
        }
    }
    style
}

fn conditional_block_contains(block: &ConditionalFormatting, column: u32, row: u32) -> bool {
    block.ranges.iter().any(|range| {
        column >= range.start_column
            && column <= range.end_column
            && row >= range.start_row
            && row <= range.end_row
    })
}

fn conditional_statistics(
    block: &ConditionalFormatting,
    values: &HashMap<(u32, u32), f64>,
) -> ConditionalStats {
    let mut sorted = values
        .iter()
        .filter(|((column, row), _)| conditional_block_contains(block, *column, *row))
        .map(|(_, value)| *value)
        .collect::<Vec<_>>();
    sorted.sort_by(f64::total_cmp);
    ConditionalStats {
        minimum: sorted.first().copied().unwrap_or(0.0),
        maximum: sorted.last().copied().unwrap_or(0.0),
        sorted,
    }
}

fn conditional_value(value: ConditionalValue, stats: &ConditionalStats) -> f64 {
    match value {
        ConditionalValue::Minimum => stats.minimum,
        ConditionalValue::Maximum => stats.maximum,
        ConditionalValue::Number(value) => value,
        ConditionalValue::Percent(percent) => {
            stats.minimum + (stats.maximum - stats.minimum) * percent / 100.0
        }
        ConditionalValue::Percentile(percent) => {
            let position = (stats.sorted.len().saturating_sub(1)) as f64 * percent / 100.0;
            let lower = position.floor() as usize;
            let upper = position.ceil() as usize;
            match (stats.sorted.get(lower), stats.sorted.get(upper)) {
                (Some(lower), Some(upper)) => lower + (upper - lower) * position.fract(),
                _ => stats.minimum,
            }
        }
    }
}

fn color_scale_color(
    value: f64,
    thresholds: &[ConditionalValue],
    colors: &[u32],
    stats: &ConditionalStats,
) -> u32 {
    let points = thresholds
        .iter()
        .zip(colors)
        .map(|(threshold, color)| (conditional_value(*threshold, stats), *color))
        .collect::<Vec<_>>();
    if let Some((_, color)) = points.first()
        && value <= points[0].0
    {
        return *color;
    }
    for segment in points.windows(2) {
        let (start, start_color) = segment[0];
        let (end, end_color) = segment[1];
        if value <= end {
            let progress = if end > start {
                ((value - start) / (end - start)).clamp(0.0, 1.0)
            } else {
                1.0
            };
            return interpolate_color(start_color, end_color, progress);
        }
    }
    points.last().map_or(0xffff_ffff, |(_, color)| *color)
}

fn interpolate_color(start: u32, end: u32, progress: f64) -> u32 {
    let channel = |shift: u32| {
        let start = f64::from((start >> shift) & 0xff);
        let end = f64::from((end >> shift) & 0xff);
        (start + (end - start) * progress).round() as u32
    };
    (channel(24) << 24) | (channel(16) << 16) | (channel(8) << 8) | channel(0)
}

#[allow(clippy::too_many_arguments)]
fn push_sparklines(
    groups: &[SparklineGroup],
    cached_values: &HashMap<(u32, u32), f64>,
    dimension: Dimension,
    column_metrics: &AxisMetrics,
    row_metrics: &AxisMetrics,
    part: &str,
    sheet_name: &str,
    unit_index: u32,
    object_limit: usize,
    state: &mut XlsxParseState,
) -> Result<bool, Diagnostic> {
    let mut unsupported = false;
    let mut resolved_points = 0_usize;
    for group in groups {
        let mut resolved = Vec::with_capacity(group.sparklines.len());
        for sparkline in &group.sparklines {
            if sparkline.target.column < dimension.start_column
                || sparkline.target.column > dimension.end_column
                || sparkline.target.row < dimension.start_row
                || sparkline.target.row > dimension.end_row
            {
                unsupported = true;
                continue;
            }
            let Some(source) =
                parse_sparkline_source_range(&sparkline.source_formula, sheet_name, part)
            else {
                unsupported = true;
                continue;
            };
            let row_count = u64::from(source.end_row - source.start_row + 1);
            let column_count = u64::from(source.end_column - source.start_column + 1);
            let Some(point_count) = row_count
                .checked_mul(column_count)
                .and_then(|count| usize::try_from(count).ok())
            else {
                unsupported = true;
                continue;
            };
            let Some(next_resolved_points) = resolved_points.checked_add(point_count) else {
                unsupported = true;
                continue;
            };
            if next_resolved_points > object_limit {
                unsupported = true;
                continue;
            }
            resolved_points = next_resolved_points;
            let mut values = Vec::with_capacity(point_count);
            for row in source.start_row..=source.end_row {
                for column in source.start_column..=source.end_column {
                    values.push(cached_values.get(&(column, row)).copied());
                }
            }
            if values.iter().all(Option::is_none) {
                unsupported = true;
                continue;
            }
            resolved.push((&sparkline.target, values));
        }
        let group_minimum = resolved
            .iter()
            .flat_map(|(_, values)| values.iter().flatten().copied())
            .fold(f64::INFINITY, f64::min);
        let group_maximum = resolved
            .iter()
            .flat_map(|(_, values)| values.iter().flatten().copied())
            .fold(f64::NEG_INFINITY, f64::max);
        for (target, values) in resolved {
            let width = column_metrics.size(target.column);
            let height = row_metrics.size(target.row);
            if width == 0.0 || height == 0.0 {
                continue;
            }
            let bounds = Rect {
                x: column_metrics.offset(target.column),
                y: row_metrics.offset(target.row),
                width,
                height,
            };
            let inset_x = (width * 0.06).clamp(1.0, 4.0);
            let inset_y = (height * 0.12).clamp(1.0, 3.0);
            let plot = Rect {
                x: bounds.x + inset_x,
                y: bounds.y + inset_y,
                width: (bounds.width - inset_x * 2.0).max(1.0),
                height: (bounds.height - inset_y * 2.0).max(1.0),
            };
            let local_minimum = values
                .iter()
                .flatten()
                .copied()
                .fold(f64::INFINITY, f64::min);
            let local_maximum = values
                .iter()
                .flatten()
                .copied()
                .fold(f64::NEG_INFINITY, f64::max);
            let automatic_minimum = if group.min_axis == SparklineAxisMode::Group {
                group_minimum
            } else {
                local_minimum
            };
            let automatic_maximum = if group.max_axis == SparklineAxisMode::Group {
                group_maximum
            } else {
                local_maximum
            };
            let mut minimum = match group.min_axis {
                SparklineAxisMode::Custom => group.manual_min.unwrap_or_else(|| {
                    unsupported = true;
                    automatic_minimum
                }),
                SparklineAxisMode::Individual | SparklineAxisMode::Group => automatic_minimum,
            };
            let mut maximum = match group.max_axis {
                SparklineAxisMode::Custom => group.manual_max.unwrap_or_else(|| {
                    unsupported = true;
                    automatic_maximum
                }),
                SparklineAxisMode::Individual | SparklineAxisMode::Group => automatic_maximum,
            };
            if group.kind != SparklineKind::Line {
                if group.min_axis != SparklineAxisMode::Custom {
                    minimum = minimum.min(0.0);
                }
                if group.max_axis != SparklineAxisMode::Custom {
                    maximum = maximum.max(0.0);
                }
            }
            if minimum >= maximum {
                unsupported |= minimum > maximum;
                let padding = minimum.abs().max(1.0) * 0.05;
                minimum -= padding;
                maximum += padding;
            }
            let value_y = |value: f64| {
                plot.y
                    + plot.height
                        * (1.0 - ((value - minimum) / (maximum - minimum)).clamp(0.0, 1.0)) as f32
            };
            let axis_y = value_y(0.0);
            if group.display_axis {
                push_sparkline_shape(
                    state,
                    object_limit,
                    bounds,
                    Geometry::Path {
                        fill_rule: FillRule::NonZero,
                        commands: vec![
                            PathCommand::MoveTo {
                                x: plot.x - bounds.x,
                                y: axis_y - bounds.y,
                            },
                            PathCommand::LineTo {
                                x: plot.x + plot.width - bounds.x,
                                y: axis_y - bounds.y,
                            },
                        ],
                    },
                    Paint::None,
                    Paint::Solid(group.axis_color),
                    1.0,
                    part,
                    sheet_name,
                    unit_index,
                    target,
                )?;
            }
            match group.kind {
                SparklineKind::Line => {
                    let denominator = values.len().saturating_sub(1).max(1) as f32;
                    let mut commands = Vec::with_capacity(values.len());
                    let mut connected = false;
                    let mut point_count = 0_usize;
                    for (index, value) in values.iter().enumerate() {
                        let Some(value) = value else {
                            connected = false;
                            continue;
                        };
                        let x = plot.x + plot.width * index as f32 / denominator - bounds.x;
                        let y = value_y(*value) - bounds.y;
                        if connected {
                            commands.push(PathCommand::LineTo { x, y });
                        } else {
                            commands.push(PathCommand::MoveTo { x, y });
                            connected = true;
                        }
                        point_count += 1;
                    }
                    if point_count >= 2 {
                        push_sparkline_shape(
                            state,
                            object_limit,
                            bounds,
                            Geometry::Path {
                                fill_rule: FillRule::NonZero,
                                commands,
                            },
                            Paint::None,
                            Paint::Solid(group.series_color),
                            1.5,
                            part,
                            sheet_name,
                            unit_index,
                            target,
                        )?;
                    }
                }
                SparklineKind::Column | SparklineKind::WinLoss => {
                    let slot_width = plot.width / values.len().max(1) as f32;
                    let bar_width = (slot_width * 0.68).max(1.0);
                    for (index, value) in values.iter().enumerate() {
                        let Some(value) = value.filter(|value| *value != 0.0) else {
                            continue;
                        };
                        let value_position = if group.kind == SparklineKind::WinLoss {
                            if value > 0.0 {
                                plot.y
                            } else {
                                plot.y + plot.height
                            }
                        } else {
                            value_y(value)
                        };
                        let top = axis_y.min(value_position);
                        let height = (axis_y - value_position).abs().max(1.0);
                        push_sparkline_shape(
                            state,
                            object_limit,
                            Rect {
                                x: plot.x
                                    + slot_width * index as f32
                                    + (slot_width - bar_width) / 2.0,
                                y: top,
                                width: bar_width,
                                height,
                            },
                            Geometry::Rectangle,
                            Paint::Solid(if value < 0.0 {
                                group.negative_color
                            } else {
                                group.series_color
                            }),
                            Paint::None,
                            0.0,
                            part,
                            sheet_name,
                            unit_index,
                            target,
                        )?;
                    }
                }
            }
        }
    }
    Ok(unsupported)
}

fn parse_sparkline_source_range(formula: &str, sheet_name: &str, part: &str) -> Option<Dimension> {
    let formula = formula.trim().strip_prefix('=').unwrap_or(formula.trim());
    if formula.contains(['[', ']']) {
        return None;
    }
    let (source_sheet, range) = match formula.rsplit_once('!') {
        Some((source_sheet, range)) => (Some(source_sheet.trim()), range.trim()),
        None => (None, formula),
    };
    if let Some(source_sheet) = source_sheet {
        let source_sheet = if source_sheet.starts_with('\'') && source_sheet.ends_with('\'') {
            source_sheet
                .get(1..source_sheet.len().saturating_sub(1))?
                .replace("''", "'")
        } else {
            source_sheet.to_owned()
        };
        if !source_sheet.eq_ignore_ascii_case(sheet_name) {
            return None;
        }
    }
    let normalized = range.replace('$', "").to_ascii_uppercase();
    if normalized
        .bytes()
        .any(|byte| byte.is_ascii_whitespace() || matches!(byte, b',' | b'(' | b')'))
    {
        return None;
    }
    parse_dimension(&normalized, part).ok()
}

#[allow(clippy::too_many_arguments)]
fn push_sparkline_shape(
    state: &mut XlsxParseState,
    object_limit: usize,
    bounds: Rect,
    geometry: Geometry,
    fill: Paint,
    stroke: Paint,
    stroke_width: f32,
    part: &str,
    sheet_name: &str,
    unit_index: u32,
    target: &CellAddress,
) -> Result<(), Diagnostic> {
    if state.objects.len() >= object_limit {
        return Err(object_limit_error(
            part,
            "sparklines exceed the configured object limit",
        ));
    }
    let numeric_id = u32::try_from(state.objects.len())
        .map_err(|_| format_error(part, "object count exceeds supported range"))?;
    state.objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: None,
        kind: ObjectKind::Shape,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: part.to_owned(),
            mapping: MappingQuality::Exact,
            locator: SourceLocator::Xlsx {
                kind: "cell",
                sheet_name: sheet_name.to_owned(),
                address: Some(target.canonical.clone()),
                formula: None,
                drawing_id: None,
            },
        },
        visual: Visual::PaintedShape {
            geometry,
            fill,
            stroke,
            stroke_width,
        },
    });
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_conditional_overlays(
    cells: &[ConditionalCell],
    blocks: &[ConditionalFormatting],
    stats: &[ConditionalStats],
    part: &str,
    sheet_name: &str,
    unit_index: u32,
    object_limit: usize,
    state: &mut XlsxParseState,
) -> Result<(), Diagnostic> {
    for cell in cells {
        for (block_index, rule) in conditional_rules_for_cell(blocks, &cell.address) {
            let Some(block_stats) = stats.get(block_index) else {
                continue;
            };
            let applied = conditional_rule_applies(rule, cell.value, &cell.address);
            match &rule.kind {
                ConditionalRuleKind::DataBar {
                    thresholds,
                    color,
                    show_value: _,
                } => {
                    let minimum = thresholds.first().map_or(block_stats.minimum, |value| {
                        conditional_value(*value, block_stats)
                    });
                    let maximum = thresholds.last().map_or(block_stats.maximum, |value| {
                        conditional_value(*value, block_stats)
                    });
                    let ratio = if maximum > minimum {
                        ((cell.value - minimum) / (maximum - minimum)).clamp(0.0, 1.0)
                    } else if cell.value >= maximum {
                        1.0
                    } else {
                        0.0
                    };
                    if ratio > 0.0 {
                        let inset = cell.bounds.height.min(cell.bounds.width) * 0.10;
                        push_conditional_object(
                            Rect {
                                x: cell.bounds.x + inset,
                                y: cell.bounds.y + cell.bounds.height * 0.22,
                                width: (cell.bounds.width - inset * 2.0).max(0.0) * ratio as f32,
                                height: cell.bounds.height * 0.56,
                            },
                            None,
                            Visual::Shape {
                                geometry: Geometry::Rectangle,
                                fill: (*color & 0xffff_ff00) | 0x99,
                                stroke: 0,
                                stroke_width: 0.0,
                            },
                            "cell",
                            &cell.address,
                            part,
                            sheet_name,
                            unit_index,
                            object_limit,
                            state,
                        )?;
                    }
                }
                ConditionalRuleKind::IconSet {
                    thresholds,
                    inclusive,
                    icon_set,
                    reverse,
                    show_value: _,
                } => {
                    // The first cfvo defines the lowest bucket, not a comparison.
                    let mut index = thresholds
                        .iter()
                        .zip(inclusive)
                        .enumerate()
                        .skip(1)
                        .filter(|(_, (value, inclusive))| {
                            let threshold = conditional_value(**value, block_stats);
                            cell.value > threshold || (**inclusive && cell.value == threshold)
                        })
                        .map(|(index, _)| index)
                        .last()
                        .unwrap_or(0);
                    if *reverse {
                        index = thresholds.len() - 1 - index;
                    }
                    let icon_bounds = conditional_icon_bounds(cell.bounds);
                    if thresholds.len() == 3
                        && matches!(icon_set.as_str(), "3Arrows" | "3ArrowsGray" | "3Symbols2")
                    {
                        let size = icon_bounds.width;
                        // Conditional-format icons have fixed pixel proportions, unlike DrawingML arrows.
                        let symbol_points: &[(f32, f32)] = match index {
                            0 => &[
                                (3.0, 2.0),
                                (8.0, 6.0),
                                (13.0, 2.0),
                                (15.0, 4.0),
                                (10.0, 8.0),
                                (15.0, 12.0),
                                (13.0, 14.0),
                                (8.0, 10.0),
                                (3.0, 14.0),
                                (1.0, 12.0),
                                (6.0, 8.0),
                                (1.0, 4.0),
                            ],
                            1 => &[(6.0, 1.0), (10.0, 1.0), (9.0, 10.0), (7.0, 10.0)],
                            _ => &[
                                (1.0, 8.0),
                                (4.0, 5.0),
                                (7.0, 8.0),
                                (12.0, 1.0),
                                (15.0, 3.0),
                                (7.0, 14.0),
                            ],
                        };
                        let arrow_points = [
                            (14.5, 8.0),
                            (8.0, 1.5),
                            (8.0, 5.0),
                            (1.5, 5.0),
                            (1.5, 11.0),
                            (8.0, 11.0),
                            (8.0, 14.5),
                        ];
                        let symbols = icon_set == "3Symbols2";
                        let points: Vec<_> = (if symbols {
                            symbol_points
                        } else {
                            &arrow_points
                        })
                        .iter()
                        .map(|&(x, y)| {
                            let (x, y) = if symbols {
                                (x, y)
                            } else {
                                match index {
                                    0 => (y, x),
                                    2 => (y, 16.0 - x),
                                    _ => (x, y),
                                }
                            };
                            (x * size / 16.0, y * size / 16.0)
                        })
                        .collect();
                        let mut geometry = super::polygon_geometry(&points);
                        if symbols && index == 1 {
                            let dot = [(6.5, 12.0), (9.5, 12.0), (9.5, 15.0), (6.5, 15.0)]
                                .map(|(x, y)| (x * size / 16.0, y * size / 16.0));
                            if let (
                                Geometry::Path { commands, .. },
                                Geometry::Path { commands: dot, .. },
                            ) = (&mut geometry, super::polygon_geometry(&dot))
                            {
                                commands.extend(dot);
                            }
                        }
                        let (fill, stroke) = if symbols {
                            [
                                (0xd36a_67ff, 0xa33f_3cff),
                                (0xf0ae_4cff, 0xba78_29ff),
                                (0x63a3_6aff, 0x3c74_42ff),
                            ][index]
                        } else if icon_set == "3ArrowsGray" {
                            (0xa6a6_a6ff, 0x5959_59ff)
                        } else {
                            [
                                (0xff50_50ff, 0xa824_24ff),
                                (0xffc0_00ff, 0xa67c_00ff),
                                (0x70ad_47ff, 0x507e_32ff),
                            ][index]
                        };
                        push_conditional_object(
                            icon_bounds,
                            None,
                            Visual::Shape {
                                geometry,
                                fill,
                                stroke,
                                stroke_width: size / 16.0,
                            },
                            "cell",
                            &cell.address,
                            part,
                            sheet_name,
                            unit_index,
                            object_limit,
                            state,
                        )?;
                        if applied && rule.stop_if_true {
                            break;
                        }
                        continue;
                    }
                    let (symbol, color) = conditional_icon(icon_set, thresholds.len(), index);
                    let traffic_light = icon_set == "3TrafficLights1";
                    let inset = if traffic_light {
                        icon_bounds.width / 16.0
                    } else {
                        0.0
                    };
                    push_conditional_object(
                        Rect {
                            x: icon_bounds.x + inset,
                            y: icon_bounds.y + inset,
                            width: icon_bounds.width - inset * 2.0,
                            height: icon_bounds.height - inset * 2.0,
                        },
                        (!traffic_light).then(|| symbol.to_owned()),
                        if traffic_light {
                            Visual::Shape {
                                geometry: Geometry::Ellipse,
                                fill: color,
                                stroke: 0,
                                stroke_width: 0.0,
                            }
                        } else {
                            Visual::TextLayout {
                                layout: TextLayout {
                                    vertical_align: TextVerticalAlign::Center,
                                    inset_left: 0.0,
                                    inset_right: 0.0,
                                    inset_top: 0.0,
                                    inset_bottom: 0.0,
                                    wrap: false,
                                    ..TextLayout::default()
                                },
                                visual: Box::new(Visual::Text {
                                    geometry: Geometry::Rectangle,
                                    fill: 0,
                                    stroke: 0,
                                    stroke_width: 0.0,
                                    font_family: "Arial".to_owned(),
                                    font_size: icon_bounds.height,
                                    color,
                                    bold: true,
                                    italic: false,
                                    align: TextAlign::Center,
                                }),
                            }
                        },
                        "cell",
                        &cell.address,
                        part,
                        sheet_name,
                        unit_index,
                        object_limit,
                        state,
                    )?;
                }
                ConditionalRuleKind::Calculated { .. }
                | ConditionalRuleKind::CellIs { .. }
                | ConditionalRuleKind::ColorScale { .. } => {}
            }
            if applied && rule.stop_if_true {
                break;
            }
        }
    }
    Ok(())
}

fn conditional_icon(icon_set: &str, count: usize, index: usize) -> (&'static str, u32) {
    let colors = [
        0xd930_25ff,
        0xf59e_0bff,
        0xfacc_15ff,
        0x84cc_16ff,
        0x16a3_4aff,
    ];
    let color_index = if count <= 1 {
        colors.len() - 1
    } else {
        index * (colors.len() - 1) / (count - 1)
    };
    let symbols = if icon_set.contains("Arrow") {
        ["↓", "↘", "→", "↗", "↑"]
    } else if icon_set.contains("Flag") {
        ["⚑", "⚑", "⚑", "⚑", "⚑"]
    } else if icon_set.contains("Rating") || icon_set.contains("Star") {
        ["☆", "☆", "★", "★", "★"]
    } else {
        ["●", "●", "●", "●", "●"]
    };
    let symbol_index = if count <= 1 {
        symbols.len() - 1
    } else {
        index * (symbols.len() - 1) / (count - 1)
    };
    (symbols[symbol_index], colors[color_index])
}

#[allow(clippy::too_many_arguments)]
fn push_conditional_object(
    bounds: Rect,
    text: Option<String>,
    visual: Visual,
    source_kind: &'static str,
    address: &CellAddress,
    part: &str,
    sheet_name: &str,
    unit_index: u32,
    object_limit: usize,
    state: &mut XlsxParseState,
) -> Result<(), Diagnostic> {
    if state.objects.len() >= object_limit {
        return Err(object_limit_error(
            part,
            "conditional formatting exceeds the configured object limit",
        ));
    }
    let numeric_id = u32::try_from(state.objects.len())
        .map_err(|_| format_error(part, "object count exceeds supported range"))?;
    state.objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: None,
        kind: ObjectKind::Shape,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text,
        source: SourceRef {
            part: part.to_owned(),
            mapping: MappingQuality::Derived,
            locator: SourceLocator::Xlsx {
                kind: source_kind,
                sheet_name: sheet_name.to_owned(),
                address: Some(address.canonical.clone()),
                formula: None,
                drawing_id: None,
            },
        },
        visual,
    });
    Ok(())
}

fn cell_numeric_value(cell: &CellState) -> Option<f64> {
    (cell.value_type == CellValueType::Number && cell.value_seen)
        .then(|| cell.value.trim().parse::<f64>().ok())
        .flatten()
        .filter(|value| value.is_finite())
}

fn conditional_icon_bounds(bounds: Rect) -> Rect {
    let size = 16.0_f32
        .min((bounds.height - 2.0).max(1.0))
        .min((bounds.width - 4.0).max(1.0));
    Rect {
        x: bounds.x + 2.0,
        y: bounds.y + (bounds.height - size) / 2.0,
        width: size,
        height: size,
    }
}

fn conditional_value_layout(cell: &CellState, blocks: &[ConditionalFormatting]) -> (bool, bool) {
    let Some(value) = cell_numeric_value(cell) else {
        return (false, false);
    };
    let mut hidden = false;
    let mut has_icon = false;
    for (_, rule) in conditional_rules_for_cell(blocks, &cell.address) {
        let applied = conditional_rule_applies(rule, value, &cell.address);
        if applied {
            match &rule.kind {
                ConditionalRuleKind::DataBar { show_value, .. } => hidden |= !show_value,
                ConditionalRuleKind::IconSet { show_value, .. } => {
                    hidden |= !show_value;
                    has_icon = true;
                }
                _ => {}
            }
        }
        if applied && rule.stop_if_true {
            break;
        }
    }
    (hidden, has_icon)
}

fn conditional_rule_applies(rule: &ConditionalRule, value: f64, address: &CellAddress) -> bool {
    match &rule.kind {
        ConditionalRuleKind::Calculated { matches, .. } => {
            matches.contains(&(address.column, address.row))
        }
        ConditionalRuleKind::CellIs {
            operator, formulas, ..
        } => numeric_rule_matches(value, *operator, formulas),
        ConditionalRuleKind::ColorScale { .. }
        | ConditionalRuleKind::DataBar { .. }
        | ConditionalRuleKind::IconSet { .. } => true,
    }
}

fn numeric_rule_matches(
    value: f64,
    operator: CellOperator,
    formulas: &[ConditionalLiteral],
) -> bool {
    let number = |index| match formulas.get(index) {
        Some(ConditionalLiteral::Number(value)) => Some(*value),
        Some(ConditionalLiteral::Text(_)) | None => None,
    };
    let first = number(0);
    let second = number(1);
    match (operator, first, second) {
        (CellOperator::LessThan, Some(first), _) => value < first,
        (CellOperator::LessThanOrEqual, Some(first), _) => value <= first,
        (CellOperator::Equal, Some(first), _) => value == first,
        (CellOperator::NotEqual, Some(first), _) => value != first,
        (CellOperator::GreaterThanOrEqual, Some(first), _) => value >= first,
        (CellOperator::GreaterThan, Some(first), _) => value > first,
        (CellOperator::Between, Some(first), Some(second)) => value >= first && value <= second,
        (CellOperator::NotBetween, Some(first), Some(second)) => value < first || value > second,
        _ => false,
    }
}

fn cell_rule_matches(
    cell: &CellState,
    shared_strings: &[SharedString],
    operator: CellOperator,
    formulas: &[ConditionalLiteral],
) -> bool {
    if let Some(value) = cell_numeric_value(cell) {
        return numeric_rule_matches(value, operator, formulas);
    }
    let Some(value) = cell_conditional_text(cell, shared_strings) else {
        return false;
    };
    let text = |index| match formulas.get(index) {
        Some(ConditionalLiteral::Text(value)) => Some(value.as_str()),
        Some(ConditionalLiteral::Number(_)) | None => None,
    };
    let first = text(0);
    let second = text(1);
    let compare = |left: &str, right: &str| left.to_lowercase().cmp(&right.to_lowercase());
    match (operator, first, second) {
        (CellOperator::LessThan, Some(first), _) => compare(value, first).is_lt(),
        (CellOperator::LessThanOrEqual, Some(first), _) => !compare(value, first).is_gt(),
        (CellOperator::Equal, Some(first), _) => value.eq_ignore_ascii_case(first),
        (CellOperator::NotEqual, Some(first), _) => !value.eq_ignore_ascii_case(first),
        (CellOperator::GreaterThanOrEqual, Some(first), _) => !compare(value, first).is_lt(),
        (CellOperator::GreaterThan, Some(first), _) => compare(value, first).is_gt(),
        (CellOperator::Between, Some(first), Some(second)) => {
            !compare(value, first).is_lt() && !compare(value, second).is_gt()
        }
        (CellOperator::NotBetween, Some(first), Some(second)) => {
            compare(value, first).is_lt() || compare(value, second).is_gt()
        }
        _ => false,
    }
}

fn cell_conditional_text<'a>(
    cell: &'a CellState,
    shared_strings: &'a [SharedString],
) -> Option<&'a str> {
    match cell.value_type {
        CellValueType::SharedString if cell.value_seen => cell
            .value
            .trim()
            .parse::<usize>()
            .ok()
            .and_then(|index| shared_strings.get(index))
            .map(|value| value.text.as_str()),
        CellValueType::InlineString if cell.inline_string_seen => Some(cell.inline_text.as_str()),
        CellValueType::String | CellValueType::Error | CellValueType::Date if cell.value_seen => {
            Some(cell.value.as_str())
        }
        CellValueType::Number
        | CellValueType::SharedString
        | CellValueType::InlineString
        | CellValueType::Boolean
        | CellValueType::String
        | CellValueType::Error
        | CellValueType::Date => None,
    }
}

fn cell_text_align(style: &CellStyle, value_type: CellValueType) -> TextAlign {
    if !style.general_alignment {
        return style.align;
    }
    match value_type {
        CellValueType::Number | CellValueType::Date => TextAlign::End,
        CellValueType::Boolean | CellValueType::Error => TextAlign::Center,
        CellValueType::SharedString | CellValueType::InlineString | CellValueType::String => {
            TextAlign::Start
        }
    }
}

fn cell_text_runs(
    cell: &CellState,
    shared_strings: &[SharedString],
    style: &CellStyle,
    displayed_text: Option<&str>,
) -> Option<Vec<TextRun>> {
    let rich_runs = match cell.value_type {
        CellValueType::SharedString if cell.value_seen => cell
            .value
            .trim()
            .parse::<usize>()
            .ok()
            .and_then(|index| shared_strings.get(index))
            .map(|shared| shared.runs.as_slice()),
        CellValueType::InlineString if cell.inline_string_seen => Some(cell.inline_runs.as_slice()),
        _ => None,
    };
    if let Some(runs) = rich_runs.filter(|runs| !runs.is_empty()) {
        return Some(
            runs.iter()
                .map(|run| {
                    let font_size = run.font_size.unwrap_or(style.font_size);
                    TextRun {
                        paint: None,
                        east_asian_line_breaks: true,
                        text: run.text.clone(),
                        font_family: run
                            .font_family
                            .clone()
                            .unwrap_or_else(|| style.font_family.clone()),
                        font_size,
                        color: run.color.unwrap_or(style.color),
                        bold: run.bold.unwrap_or(style.bold),
                        italic: run.italic.unwrap_or(style.italic),
                        underline: run.underline.unwrap_or(style.underline),
                        strikethrough: run.strikethrough.unwrap_or(style.strikethrough),
                        highlight: 0,
                        baseline_shift: run
                            .baseline_shift
                            .map_or(style.baseline_shift, |ratio| ratio * font_size),
                        letter_spacing: 0.0,
                        horizontal_scale: 1.0,
                    }
                })
                .collect(),
        );
    }
    (style.underline || style.strikethrough || style.baseline_shift != 0.0).then(|| {
        vec![TextRun {
            paint: None,
            east_asian_line_breaks: true,
            text: displayed_text.unwrap_or_default().to_owned(),
            font_family: style.font_family.clone(),
            font_size: style.font_size,
            color: style.color,
            bold: style.bold,
            italic: style.italic,
            underline: style.underline,
            strikethrough: style.strikethrough,
            highlight: 0,
            baseline_shift: style.baseline_shift,
            letter_spacing: 0.0,
            horizontal_scale: 1.0,
        }]
    })
}

#[allow(clippy::too_many_arguments)]
fn push_cell(
    cell: CellState,
    bounds: Rect,
    general_width: usize,
    part: &str,
    sheet_name: &str,
    unit_index: u32,
    shared_strings: &[SharedString],
    style: &CellStyle,
    date_1904: bool,
    hide_value: bool,
    icon_inset: f32,
    show_grid_lines: bool,
    clip_horizontal_overflow: bool,
    object_limit: usize,
    remaining_cells: usize,
    state: &mut XlsxParseState,
) -> Result<(), Diagnostic> {
    let text = if hide_value {
        None
    } else {
        resolve_cell_text(
            &cell,
            shared_strings,
            style,
            date_1904,
            general_width,
            part,
            &mut state.materialized_text_bytes,
            state.materialized_text_limit,
        )?
    };
    let (stroke, stroke_width) = if show_grid_lines && !style.has_fill && !style.has_border {
        (0xd0d0_d0ff, 1.0)
    } else {
        (0x0000_0000, 0.0)
    };
    let visual = match (!hide_value)
        .then(|| cell_text_runs(&cell, shared_strings, style, text.as_deref()))
        .flatten()
        .or_else(|| {
            style.fill_paint.as_ref().map(|_| {
                vec![TextRun {
                    paint: None,
                    east_asian_line_breaks: true,
                    text: text.clone().unwrap_or_default(),
                    font_family: style.font_family.clone(),
                    font_size: style.font_size,
                    color: style.color,
                    bold: style.bold,
                    italic: style.italic,
                    underline: style.underline,
                    strikethrough: style.strikethrough,
                    highlight: 0,
                    baseline_shift: style.baseline_shift,
                    letter_spacing: 0.0,
                    horizontal_scale: 1.0,
                }]
            })
        }) {
        Some(runs) => Visual::RichText {
            geometry: Geometry::Rectangle,
            fill: style
                .fill_paint
                .as_ref()
                .map_or(Paint::Solid(style.fill), |fill| fill.paint(bounds)),
            stroke: Paint::Solid(stroke),
            stroke_width,
            align: style.align,
            line_height: 0.0,
            runs,
        },
        None => Visual::Text {
            geometry: Geometry::Rectangle,
            fill: style.fill,
            stroke,
            stroke_width,
            font_family: style.font_family.clone(),
            font_size: style.font_size,
            color: style.color,
            bold: style.bold,
            italic: style.italic,
            align: style.align,
        },
    };
    let visual = Visual::TextLayout {
        layout: TextLayout {
            vertical_align: style.vertical_align,
            auto_fit: if style.shrink_to_fit {
                TextAutoFit::Shrink
            } else {
                TextAutoFit::None
            },
            wrap: style.wrap,
            rotation_degrees: style.rotation_degrees,
            horizontal_overflow: if clip_horizontal_overflow {
                TextHorizontalOverflow::Clip
            } else {
                TextHorizontalOverflow::Overflow
            },
            direction: style.direction,
            inset_left: 2.0
                + icon_inset
                + if style.direction == TextDirection::Rtl {
                    style.indent as f32
                } else {
                    style.indent as f32 + style.relative_indent as f32
                }
                .max(0.0)
                    * style.font_size
                    * 1.5,
            inset_right: 2.0
                + if style.direction == TextDirection::Rtl {
                    style.relative_indent as f32
                } else {
                    0.0
                }
                .max(0.0)
                    * style.font_size
                    * 1.5,
            inset_top: 1.0,
            inset_bottom: 1.0,
            ..TextLayout::default()
        },
        visual: Box::new(visual),
    };
    if style.has_border {
        let mut children = Vec::with_capacity(6);
        for (side_index, side) in style.borders.iter().enumerate() {
            if side.width <= 0.0 {
                continue;
            }
            let side_bounds = match side_index {
                0 => Rect {
                    width: side.width,
                    ..bounds
                },
                1 => Rect {
                    x: bounds.x + bounds.width - side.width,
                    width: side.width,
                    ..bounds
                },
                2 => Rect {
                    height: side.width,
                    ..bounds
                },
                3 => Rect {
                    y: bounds.y + bounds.height - side.width,
                    height: side.width,
                    ..bounds
                },
                _ => unreachable!(),
            };
            children.push(VisualBrushChild {
                bounds: side_bounds,
                visual: Visual::PaintedShape {
                    geometry: Geometry::Rectangle,
                    fill: Paint::Solid(side.color),
                    stroke: Paint::None,
                    stroke_width: 0.0,
                },
            });
        }
        for (enabled, (start_x, start_y, end_x, end_y)) in [
            (style.diagonal_down, (0.0, 0.0, bounds.width, bounds.height)),
            (style.diagonal_up, (0.0, bounds.height, bounds.width, 0.0)),
        ] {
            if !enabled || style.diagonal_border.width <= 0.0 {
                continue;
            }
            children.push(VisualBrushChild {
                bounds,
                visual: Visual::PaintedShape {
                    geometry: Geometry::Path {
                        fill_rule: FillRule::NonZero,
                        commands: vec![
                            PathCommand::MoveTo {
                                x: start_x,
                                y: start_y,
                            },
                            PathCommand::LineTo { x: end_x, y: end_y },
                        ],
                    },
                    fill: Paint::None,
                    stroke: Paint::Solid(style.diagonal_border.color),
                    stroke_width: style.diagonal_border.width,
                },
            });
        }
        if state
            .objects
            .len()
            .saturating_add(remaining_cells)
            .saturating_add(1)
            <= object_limit
        {
            let numeric_id = u32::try_from(state.objects.len())
                .map_err(|_| format_error(part, "object count exceeds supported range"))?;
            state.objects.push(Object {
                numeric_id,
                parent_numeric_id: None,
                stable_id: format!("object:{numeric_id}"),
                parent_stable_id: None,
                kind: ObjectKind::Shape,
                unit_index,
                bounds,
                z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
                text: None,
                source: SourceRef {
                    part: part.to_owned(),
                    mapping: MappingQuality::Exact,
                    locator: SourceLocator::Xlsx {
                        kind: "cell",
                        sheet_name: sheet_name.to_owned(),
                        address: Some(cell.address.canonical.clone()),
                        formula: None,
                        drawing_id: None,
                    },
                },
                visual: Visual::Group { children },
            });
        } else if !state
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message.contains("cell borders were omitted"))
        {
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::ObjectLimit,
                    Phase::Render,
                    Fidelity::Omitted,
                    "XLSX cell borders were omitted to preserve worksheet content within the object limit",
                )
                .in_part(part),
            );
        }
    }
    // ponytail: shared-formula followers have no authored text; omit formula-bar text until
    // FormulaEvaluator exposes the translated follower formula.
    let formula = if cell.formula.is_empty() {
        None
    } else {
        reserve_materialized_text_bytes(
            &mut state.materialized_text_bytes,
            cell.formula.len(),
            usize::from(!cell.formula.starts_with('=')),
            state.materialized_text_limit,
            part,
        )?;
        Some(if cell.formula.starts_with('=') {
            cell.formula.clone()
        } else {
            format!("={}", cell.formula)
        })
    };
    let numeric_id = u32::try_from(state.objects.len())
        .map_err(|_| format_error(part, "object count exceeds supported range"))?;
    let z = i32::try_from(numeric_id)
        .map_err(|_| object_limit_error(part, "object z-order exceeds supported range"))?;
    state.objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: None,
        kind: ObjectKind::Cell,
        unit_index,
        bounds,
        z,
        text,
        source: SourceRef {
            part: part.to_owned(),
            mapping: MappingQuality::Exact,
            locator: SourceLocator::Xlsx {
                kind: "cell",
                sheet_name: sheet_name.to_owned(),
                address: Some(cell.address.canonical),
                formula,
                drawing_id: None,
            },
        },
        visual,
    });
    Ok(())
}

fn resolve_cell_text(
    cell: &CellState,
    shared_strings: &[SharedString],
    style: &CellStyle,
    date_1904: bool,
    general_width: usize,
    part: &str,
    materialized_text_bytes: &mut usize,
    materialized_text_limit: usize,
) -> Result<Option<String>, Diagnostic> {
    if cell.value_type == CellValueType::InlineString {
        if cell.value_seen {
            return Err(format_error(
                part,
                format!(
                    "inline-string cell {} must not contain a cached value",
                    cell.address.canonical
                ),
            ));
        }
        if !cell.inline_string_seen {
            return Ok(None);
        }
        reserve_materialized_text_bytes(
            materialized_text_bytes,
            cell.inline_text.len(),
            1,
            materialized_text_limit,
            part,
        )?;
        return clone_materialized_text(&cell.inline_text, part).map(Some);
    }
    if !cell.value_seen {
        return Ok(None);
    }

    let value = cell.value.trim();
    match cell.value_type {
        CellValueType::SharedString => {
            let index = value.parse::<usize>().map_err(|_| {
                format_error(
                    part,
                    format!(
                        "cell {} has invalid shared-string index",
                        cell.address.canonical
                    ),
                )
            })?;
            let shared = shared_strings.get(index).ok_or_else(|| {
                format_error(
                    part,
                    format!(
                        "cell {} references a missing shared string",
                        cell.address.canonical
                    ),
                )
            })?;
            reserve_materialized_text_bytes(
                materialized_text_bytes,
                shared.text.len(),
                1,
                materialized_text_limit,
                part,
            )?;
            clone_materialized_text(&shared.text, part).map(Some)
        }
        CellValueType::Boolean => match value {
            "0" | "1" | "false" | "true" => {
                let normalized = if matches!(value, "0" | "false") {
                    "false"
                } else {
                    "true"
                };
                reserve_materialized_text_bytes(
                    materialized_text_bytes,
                    normalized.len(),
                    1,
                    materialized_text_limit,
                    part,
                )?;
                clone_materialized_text(normalized, part).map(Some)
            }
            _ => Err(format_error(
                part,
                format!(
                    "cell {} has an invalid boolean value",
                    cell.address.canonical
                ),
            )),
        },
        CellValueType::Number => {
            let formatted = style
                .number_format
                .as_deref()
                .and_then(|format| format_number(value, format, date_1904, general_width));
            let displayed = formatted.as_deref().unwrap_or(value);
            reserve_materialized_text_bytes(
                materialized_text_bytes,
                displayed.len(),
                1,
                materialized_text_limit,
                part,
            )?;
            clone_materialized_text(displayed, part).map(Some)
        }
        CellValueType::String | CellValueType::Error | CellValueType::Date => {
            reserve_materialized_text_bytes(
                materialized_text_bytes,
                value.len(),
                1,
                materialized_text_limit,
                part,
            )?;
            clone_materialized_text(value, part).map(Some)
        }
        CellValueType::InlineString => Err(format_error(
            part,
            "inline-string state was resolved inconsistently",
        )),
    }
}

fn derive_dimension(cells: &[CellState]) -> Dimension {
    let Some(first) = cells.first() else {
        return Dimension {
            start_column: 0,
            start_row: 0,
            end_column: 0,
            end_row: 0,
        };
    };
    cells.iter().skip(1).fold(
        Dimension {
            start_column: first.address.column,
            start_row: first.address.row,
            end_column: first.address.column,
            end_row: first.address.row,
        },
        |mut dimension, cell| {
            dimension.start_column = dimension.start_column.min(cell.address.column);
            dimension.start_row = dimension.start_row.min(cell.address.row);
            dimension.end_column = dimension.end_column.max(cell.address.column);
            dimension.end_row = dimension.end_row.max(cell.address.row);
            dimension
        },
    )
}

fn parse_cell_value_type(
    value: Option<&str>,
    address: &str,
    part: &str,
) -> Result<CellValueType, Diagnostic> {
    match value {
        None | Some("n") => Ok(CellValueType::Number),
        Some("s") => Ok(CellValueType::SharedString),
        Some("inlineStr") => Ok(CellValueType::InlineString),
        Some("b") => Ok(CellValueType::Boolean),
        Some("str") => Ok(CellValueType::String),
        Some("e") => Ok(CellValueType::Error),
        Some("d") => Ok(CellValueType::Date),
        Some(value) => Err(Diagnostic::fatal(
            DiagnosticCode::UnsupportedFeature,
            Phase::Parse,
            None,
            format!("cell {address} uses unsupported XLSX value type {value}"),
        )
        .in_part(part)),
    }
}

fn parse_finite_f32(value: &str, part: &str, label: &str) -> Result<f32, Diagnostic> {
    value
        .parse::<f32>()
        .ok()
        .filter(|number| number.is_finite() && *number >= 0.0)
        .ok_or_else(|| {
            format_error(
                part,
                format!("{label} must be a finite non-negative number"),
            )
        })
}

fn ensure_style_capacity(count: usize, limits: Limits, part: &str) -> Result<(), Diagnostic> {
    if count >= limits.max_document_objects {
        return Err(object_limit_error(
            part,
            "style table exceeds the configured object limit",
        ));
    }
    Ok(())
}

fn parse_u32(value: &str, part: &str, label: &str) -> Result<u32, Diagnostic> {
    if value.is_empty()
        || (value.len() > 1 && value.starts_with('0'))
        || !value.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(format_error(
            part,
            format!("{label} must be a canonical non-negative integer"),
        ));
    }
    value
        .parse::<u32>()
        .map_err(|_| dimension_error(part, format!("{label} exceeds its limit")))
}

fn parse_i64(value: &str, part: &str, label: &str) -> Result<i64, Diagnostic> {
    if value.is_empty()
        || value == "-0"
        || (value.starts_with('0') && value.len() > 1)
        || (value.starts_with("-0") && value.len() > 2)
        || !value
            .strip_prefix('-')
            .unwrap_or(value)
            .bytes()
            .all(|byte| byte.is_ascii_digit())
    {
        return Err(format_error(
            part,
            format!("{label} must be a canonical integer"),
        ));
    }
    value
        .parse::<i64>()
        .map_err(|_| dimension_error(part, format!("{label} exceeds its limit")))
}

fn parse_frozen_split(
    value: Option<&str>,
    maximum: u32,
    part: &str,
    label: &str,
) -> Result<u32, Diagnostic> {
    let Some(value) = value else {
        return Ok(0);
    };
    let value = value
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite() && value.fract() == 0.0 && *value >= 0.0)
        .ok_or_else(|| format_error(part, format!("{label} must be a non-negative integer")))?;
    if value > f64::from(maximum) {
        return Err(dimension_error(part, format!("{label} exceeds its limit")));
    }
    Ok(value as u32)
}

fn parse_zero_based_index(value: &str, part: &str, label: &str) -> Result<usize, Diagnostic> {
    usize::try_from(parse_u32(value, part, label)?)
        .map_err(|_| dimension_error(part, format!("{label} exceeds its limit")))
}

fn parse_optional_usize(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<usize>, Diagnostic> {
    optional_attribute(attributes, name, part)?
        .map(|value| parse_zero_based_index(&value, part, name))
        .transpose()
}

fn parse_optional_u32(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<u32>, Diagnostic> {
    optional_attribute(attributes, name, part)?
        .map(|value| parse_u32(&value, part, name))
        .transpose()
}

fn boolean_value(attributes: &[XmlAttribute<'_>], part: &str) -> Result<Option<bool>, Diagnostic> {
    parse_bool_attribute(attributes, "val", part)
}

fn parse_color(
    attributes: &[XmlAttribute<'_>],
    theme_colors: &ThemeColors,
    part: &str,
) -> Result<Option<u32>, Diagnostic> {
    let color = if let Some(rgb) = optional_attribute(attributes, "rgb", part)? {
        let digits = match rgb.len() {
            6 => rgb.as_str(),
            8 => &rgb[2..],
            _ => {
                return Err(format_error(
                    part,
                    "style color must use 6 or 8 hexadecimal digits",
                ));
            }
        };
        if !digits.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(format_error(
                part,
                "style color contains invalid hexadecimal digits",
            ));
        }
        let color = u32::from_str_radix(digits, 16)
            .map_err(|_| format_error(part, "style color is invalid"))?;
        Some((color << 8) | 0xff)
    } else if let Some(theme) = optional_attribute(attributes, "theme", part)? {
        let index = parse_zero_based_index(&theme, part, "theme color")?;
        theme_colors.get(index).copied()
    } else if let Some(indexed) = optional_attribute(attributes, "indexed", part)? {
        let indexed = parse_u32(&indexed, part, "indexed color")?;
        super::excel_indexed_color(indexed)
    } else if parse_bool_attribute(attributes, "auto", part)?.unwrap_or(false) {
        Some(0x0000_00ff)
    } else {
        None
    };
    let Some(color) = color else {
        return Ok(None);
    };
    let Some(tint) = optional_attribute(attributes, "tint", part)? else {
        return Ok(Some(color));
    };
    let tint = tint
        .parse::<f32>()
        .ok()
        .filter(|value| value.is_finite() && (-1.0..=1.0).contains(value))
        .ok_or_else(|| format_error(part, "style color tint must be between -1 and 1"))?;
    if tint < 0.0 {
        Ok(Some(transform_luminance(color, 1.0 + tint, 0.0)))
    } else {
        Ok(Some(transform_luminance(color, 1.0 - tint, tint)))
    }
}

fn border_width(style: &str) -> f32 {
    match style {
        "medium" | "mediumDashed" | "mediumDashDot" | "mediumDashDotDot" => 2.0,
        "thick" | "double" => 3.0,
        "none" => 0.0,
        _ => 1.0,
    }
}

fn parse_horizontal_alignment(value: &str) -> Option<TextAlign> {
    match value {
        "center" | "centerContinuous" | "distributed" | "fill" | "justify" => {
            Some(TextAlign::Center)
        }
        "right" => Some(TextAlign::End),
        "left" => Some(TextAlign::Start),
        _ => None,
    }
}

fn parse_vertical_alignment(value: &str) -> Option<TextVerticalAlign> {
    match value {
        "top" => Some(TextVerticalAlign::Top),
        "center" | "distributed" | "justify" => Some(TextVerticalAlign::Center),
        "bottom" => Some(TextVerticalAlign::Bottom),
        _ => None,
    }
}

fn parse_excel_text_rotation(value: &str, part: &str) -> Result<f32, Diagnostic> {
    let value = parse_u32(value, part, "cell text rotation")?;
    match value {
        0..=90 => Ok(-(value as f32)),
        91..=180 => Ok((180 - value) as f32),
        // Excel's stacked-text sentinel has no exact Canvas equivalent. Keep it
        // horizontal instead of rotating the entire cell to an unrelated angle.
        255 => Ok(0.0),
        _ => Err(format_error(
            part,
            "cell text rotation must be between 0 and 180 or equal to 255",
        )),
    }
}

fn builtin_number_format(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        14 => Some("yyyy/m/d"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        37 => Some("#,##0 ;(#,##0)"),
        38 => Some("#,##0 ;[Red](#,##0)"),
        39 => Some("#,##0.00;(#,##0.00)"),
        40 => Some("#,##0.00;[Red](#,##0.00)"),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

fn format_number(
    value: &str,
    format: &str,
    date_1904: bool,
    general_width: usize,
) -> Option<String> {
    let number = value
        .parse::<f64>()
        .ok()
        .filter(|number| number.is_finite())?;
    let number = if number == 0.0 { 0.0 } else { number };
    let section = format.split(';').next()?.trim();
    if section.eq_ignore_ascii_case("general") || section == "@" {
        return Some(format_general_number(number, general_width));
    }
    if let Some(rendered) = format_excel_datetime(number, section, date_1904) {
        return Some(rendered);
    }
    let placeholder_count = section
        .bytes()
        .filter(|byte| matches!(byte, b'0' | b'#' | b'?'))
        .count();
    if placeholder_count == 0 {
        return None;
    }
    let percent_count = section.bytes().filter(|byte| *byte == b'%').count();
    let scaled = number * 100_f64.powi(i32::try_from(percent_count).ok()?);
    let decimal_places = section
        .split_once('.')
        .map(|(_, fraction)| {
            fraction
                .bytes()
                .take_while(|byte| matches!(byte, b'0' | b'#' | b'?'))
                .count()
        })
        .unwrap_or(0);
    let mut rendered = format!("{scaled:.decimal_places$}");
    if section.contains(',') {
        rendered = super::group_decimal_digits(&rendered);
    }
    for _ in 0..percent_count {
        rendered.push('%');
    }
    if let Some((symbol, before_number)) = number_format_currency(section) {
        if before_number {
            rendered.insert_str(0, symbol);
        } else {
            rendered.push_str(symbol);
        }
    }
    Some(rendered)
}

fn number_format_currency(section: &str) -> Option<(&str, bool)> {
    let start = section.find("[$")?;
    let end = section[start + 2..].find(']')? + start + 2;
    let symbol = section[start + 2..end]
        .split_once('-')
        .map_or(&section[start + 2..end], |(symbol, _)| symbol);
    (!symbol.is_empty()).then(|| {
        let first_placeholder = section
            .find(|character| matches!(character, '0' | '#' | '?'))
            .unwrap_or(section.len());
        (symbol, start < first_placeholder)
    })
}

fn format_excel_datetime(number: f64, format: &str, date_1904: bool) -> Option<String> {
    let characters = format.chars().collect::<Vec<_>>();
    let lowercase_format = format.to_ascii_lowercase();
    let mut japanese = false;
    let mut gannen = false;
    let mut has_year = false;
    let mut has_era = false;
    let mut has_day = false;
    let mut has_hour = false;
    let mut has_second = false;
    let mut longest_month = 0_usize;
    let mut quoted = false;
    let mut index = 0_usize;
    while index < characters.len() {
        let character = characters[index];
        if character == '"' {
            quoted = !quoted;
            index += 1;
            continue;
        }
        if quoted || character == '\\' || matches!(character, '_' | '*') {
            index += if character == '\\' || matches!(character, '_' | '*') {
                2
            } else {
                1
            };
            continue;
        }
        if character == '[' {
            let end = characters[index + 1..]
                .iter()
                .position(|candidate| *candidate == ']')
                .map(|offset| index + offset + 1)?;
            let elapsed = characters[index + 1..end]
                .iter()
                .collect::<String>()
                .to_ascii_lowercase();
            if let Some(locale) = elapsed.strip_prefix("$-") {
                japanese |= locale == "ja-jp"
                    || locale.starts_with("ja-jp-")
                    || u32::from_str_radix(locale, 16).is_ok_and(|id| id & 0xffff == 0x411);
                gannen |= locale == "ja-jp-x-gannen";
            }
            has_hour |= elapsed.chars().all(|value| value == 'h');
            has_second |= elapsed.chars().all(|value| value == 's');
            index = end + 1;
            continue;
        }
        let lower = character.to_ascii_lowercase();
        if matches!(lower, 'y' | 'd' | 'h' | 's' | 'm' | 'g' | 'e') {
            let mut end = index + 1;
            while end < characters.len() && characters[end].to_ascii_lowercase() == lower {
                end += 1;
            }
            let count = end - index;
            match lower {
                'y' => has_year = true,
                'g' => has_era = true,
                'e' => has_era |= !matches!(characters.get(end), Some('+' | '-')),
                'd' => has_day = true,
                'h' => has_hour = true,
                's' => has_second = true,
                'm' => longest_month = longest_month.max(count),
                _ => {}
            }
            index = end;
        } else {
            index += 1;
        }
    }
    has_year |= japanese && has_era;
    let has_meridiem = lowercase_format.contains("am/pm");
    let date_like = has_year || longest_month >= 3 || has_day && longest_month > 0;
    let time_like =
        has_second || has_hour && (longest_month > 0 || format.contains(':') || has_meridiem);
    if !(date_like || time_like) {
        return None;
    }

    let mut serial_day = number.floor();
    if !(-2_000_000.0..=2_000_000.0).contains(&serial_day) {
        return None;
    }
    let mut seconds = ((number - serial_day) * 86_400.0).round() as i64;
    if seconds >= 86_400 {
        serial_day += 1.0;
        seconds = 0;
    }
    let serial_day = serial_day as i64;
    let (year, month, day) = excel_serial_date(serial_day, date_1904);
    if !(1..=9999).contains(&year) {
        return None;
    }
    // Era boundaries use civil dates, including the mid-year transitions.
    let era = [
        ((2019, 5, 1), ["R", "令", "令和"]),
        ((1989, 1, 8), ["H", "平", "平成"]),
        ((1926, 12, 25), ["S", "昭", "昭和"]),
        ((1912, 7, 30), ["T", "大", "大正"]),
        ((1868, 9, 8), ["M", "明", "明治"]),
    ]
    .into_iter()
    .find(|(start, _)| (year, month, day) >= *start);
    let hour = seconds / 3600;
    let minute = seconds % 3600 / 60;
    let second = seconds % 60;
    let has_date = has_year || has_day || longest_month >= 3;
    let has_time = has_hour || has_second;
    const MONTHS_SHORT: [&str; 12] = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ];
    const MONTHS_LONG: [&str; 12] = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];

    let mut rendered = String::with_capacity(format.len() + 8);
    let mut quoted = false;
    let mut index = 0_usize;
    while index < characters.len() {
        let character = characters[index];
        if character == '"' {
            quoted = !quoted;
            index += 1;
            continue;
        }
        if quoted {
            rendered.push(character);
            index += 1;
            continue;
        }
        if character == '\\' {
            if let Some(literal) = characters.get(index + 1) {
                rendered.push(*literal);
            }
            index += 2;
            continue;
        }
        if matches!(character, '_' | '*') {
            index += 2;
            continue;
        }
        if character == '[' {
            let end = characters[index + 1..]
                .iter()
                .position(|candidate| *candidate == ']')
                .map(|offset| index + offset + 1)?;
            let elapsed = characters[index + 1..end]
                .iter()
                .collect::<String>()
                .to_ascii_lowercase();
            let width = elapsed.len();
            let value = if elapsed.chars().all(|value| value == 'h') {
                Some((number * 24.0).floor() as i64)
            } else if elapsed.chars().all(|value| value == 'm') {
                Some((number * 1_440.0).floor() as i64)
            } else if elapsed.chars().all(|value| value == 's') {
                Some((number * 86_400.0).floor() as i64)
            } else {
                None
            };
            if let Some(value) = value {
                rendered.push_str(&format!("{value:0width$}"));
            }
            index = end + 1;
            continue;
        }
        if index + 5 <= characters.len()
            && characters[index..index + 5]
                .iter()
                .collect::<String>()
                .eq_ignore_ascii_case("am/pm")
        {
            rendered.push_str(if hour < 12 { "AM" } else { "PM" });
            index += 5;
            continue;
        }

        let lower = character.to_ascii_lowercase();
        if matches!(lower, 'y' | 'd' | 'h' | 's' | 'm') || japanese && matches!(lower, 'g' | 'e') {
            let mut end = index + 1;
            while end < characters.len() && characters[end].to_ascii_lowercase() == lower {
                end += 1;
            }
            let count = end - index;
            match lower {
                'g' => {
                    if let Some((_, names)) = era {
                        rendered.push_str(names[count.min(3) - 1]);
                    }
                }
                'e' => {
                    let era_year = era.map_or(year, |(start, _)| year - start.0 + 1);
                    if gannen && era_year == 1 {
                        rendered.push('元');
                    } else {
                        rendered.push_str(&format!("{era_year:0count$}"));
                    }
                }
                'y' if count == 2 => rendered.push_str(&format!("{:02}", year % 100)),
                'y' => rendered.push_str(&format!("{year:04}")),
                'd' if count == 1 => rendered.push_str(&day.to_string()),
                'd' => rendered.push_str(&format!("{day:02}")),
                'h' => {
                    let displayed = if has_meridiem {
                        let value = hour % 12;
                        if value == 0 { 12 } else { value }
                    } else {
                        hour
                    };
                    if count == 1 {
                        rendered.push_str(&displayed.to_string());
                    } else {
                        rendered.push_str(&format!("{displayed:02}"));
                    }
                }
                's' if count == 1 => rendered.push_str(&second.to_string()),
                's' => rendered.push_str(&format!("{second:02}")),
                'm' => {
                    let minute_token = has_time
                        && (!has_date
                            || index.checked_sub(1).is_some_and(|at| characters[at] == ':')
                            || characters.get(end).is_some_and(|value| *value == ':'));
                    let value = if minute_token { minute as u32 } else { month };
                    match (minute_token, count) {
                        (false, 3) => rendered.push_str(MONTHS_SHORT[(month - 1) as usize]),
                        (false, 4..) => rendered.push_str(MONTHS_LONG[(month - 1) as usize]),
                        (_, 1) => rendered.push_str(&value.to_string()),
                        _ => rendered.push_str(&format!("{value:02}")),
                    }
                }
                _ => {}
            }
            index = end;
        } else {
            rendered.push(character);
            index += 1;
        }
    }
    Some(rendered)
}

fn excel_column_width(width: f32, maximum_digit_width: f32, part: &str) -> Result<f32, Diagnostic> {
    if width > 255.0 {
        return Err(dimension_error(part, "column width exceeds the XLSX limit"));
    }
    Ok(excel_width_pixels(width, maximum_digit_width))
}

fn excel_default_column_width(
    width: f32,
    maximum_digit_width: f32,
    part: &str,
) -> Result<f32, Diagnostic> {
    if width >= 65_536.0 {
        return Err(dimension_error(
            part,
            "default column width exceeds the XLSX limit",
        ));
    }
    Ok(excel_width_pixels(width, maximum_digit_width))
}

fn excel_width_pixels(width: f32, maximum_digit_width: f32) -> f32 {
    (width * maximum_digit_width).floor()
}

fn excel_maximum_digit_width(styles: &Styles) -> f32 {
    let Some(font) = styles.normal_font.as_ref() else {
        return 7.0;
    };
    let family = font.family.to_ascii_lowercase();
    // ponytail: Wasm cannot rasterize host fonts; replace these Office defaults with
    // browser-supplied Normal-style digit metrics when sheet metadata carries font demand.
    let reference = if family.contains("calibri") {
        (11.0, 8.0)
    } else if family.contains("arial") {
        (10.0, 7.0)
    } else {
        (11.0, 8.0)
    };
    (font.size * 72.0 / 96.0 * reference.1 / reference.0)
        .round()
        .max(1.0)
}

fn excel_default_row_height(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<Option<f32>, Diagnostic> {
    // Saved heights include the author's font metrics. Re-estimating leading here
    // distorts both the grid and drawings anchored across rows.
    optional_attribute(attributes, "defaultRowHeight", part)?
        .map(|height| {
            points_to_pixels(
                parse_finite_f32(&height, part, "default row height")?,
                part,
                "default row height",
            )
        })
        .transpose()
}

fn points_to_pixels(points: f32, part: &str, label: &str) -> Result<f32, Diagnostic> {
    if points > 409.5 {
        return Err(dimension_error(
            part,
            format!("{label} exceeds the XLSX limit"),
        ));
    }
    Ok(points * 96.0 / 72.0)
}

fn parse_one_based_index(
    value: &str,
    maximum: u32,
    part: &str,
    label: &str,
) -> Result<u32, Diagnostic> {
    if value.starts_with('0') || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(format_error(
            part,
            format!("{label} must be a canonical positive integer"),
        ));
    }
    value
        .parse::<u32>()
        .ok()
        .filter(|index| *index != 0 && *index <= maximum)
        .map(|index| index - 1)
        .ok_or_else(|| dimension_error(part, format!("{label} exceeds its limit")))
}

fn parse_bool_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<bool>, Diagnostic> {
    optional_attribute(attributes, name, part)?
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(format_error(
                part,
                format!("attribute {name} must be an XML boolean"),
            )),
        })
        .transpose()
}

fn column_metrics(
    default_width: f32,
    _column_count: u32,
    spans: &[ColumnSpan],
    part: &str,
) -> Result<AxisMetrics, Diagnostic> {
    let mut claimed = vec![false; MAX_COLUMNS as usize];
    let mut overrides = Vec::new();
    for span in spans {
        for column in span.start..=span.end {
            let slot = claimed
                .get_mut(column as usize)
                .ok_or_else(|| dimension_error(part, "column range exceeds its limit"))?;
            if *slot {
                return Err(format_error(part, "worksheet column ranges overlap"));
            }
            *slot = true;
            if let Some(width) = span.width {
                overrides.push((column, width));
            }
        }
    }
    Ok(AxisMetrics::new(default_width, overrides))
}

fn column_style_index(spans: &[ColumnSpan], column: u32) -> Option<usize> {
    spans
        .iter()
        .find(|span| span.start <= column && column <= span.end)
        .and_then(|span| span.style_index)
}

fn empty_cell_style_is_visible(style: &CellStyle, base: &CellStyle) -> bool {
    style.has_fill != base.has_fill
        || style.fill != base.fill
        || style.fill_paint != base.fill_paint
        || style.has_border
}

fn validate_merge_ranges(merges: &mut [Dimension], part: &str) -> Result<(), Diagnostic> {
    merges.sort_unstable_by_key(|merged| (merged.start_row, merged.start_column));
    let mut active = Vec::<Dimension>::new();
    for merged in merges.iter().copied() {
        active.retain(|candidate| candidate.end_row >= merged.start_row);
        if active.iter().any(|candidate| {
            candidate.start_column <= merged.end_column
                && merged.start_column <= candidate.end_column
        }) {
            return Err(format_error(part, "merged-cell ranges overlap"));
        }
        active.push(merged);
    }
    Ok(())
}

fn assign_cell_merges(cells: &[CellState], merges: &[Dimension]) -> Vec<Option<usize>> {
    let mut order = (0..cells.len()).collect::<Vec<_>>();
    order.sort_unstable_by_key(|index| {
        let address = &cells[*index].address;
        (address.row, address.column)
    });
    let mut assignments = vec![None; cells.len()];
    let mut active = Vec::<usize>::new();
    let mut merge_cursor = 0_usize;
    let mut active_row = None;
    for cell_index in order {
        let address = &cells[cell_index].address;
        if active_row != Some(address.row) {
            active.retain(|index| merges[*index].end_row >= address.row);
            while merge_cursor < merges.len() && merges[merge_cursor].start_row <= address.row {
                if merges[merge_cursor].end_row >= address.row {
                    active.push(merge_cursor);
                }
                merge_cursor += 1;
            }
            active.sort_unstable_by_key(|index| merges[*index].start_column);
            active_row = Some(address.row);
        }
        let position =
            active.partition_point(|index| merges[*index].start_column <= address.column);
        if position != 0 {
            let index = active[position - 1];
            if address.column <= merges[index].end_column {
                assignments[cell_index] = Some(index);
            }
        }
    }
    assignments
}

fn parse_dimension(value: &str, part: &str) -> Result<Dimension, Diagnostic> {
    let (start, end) = match value.split_once(':') {
        Some((start, end)) if !end.contains(':') => (start, end),
        None => (value, value),
        _ => {
            return Err(format_error(
                part,
                format!("invalid worksheet dimension: {value}"),
            ));
        }
    };
    let start = parse_a1(start)
        .ok_or_else(|| format_error(part, format!("invalid worksheet dimension: {value}")))?;
    let end = parse_a1(end)
        .ok_or_else(|| format_error(part, format!("invalid worksheet dimension: {value}")))?;
    if start.column > end.column || start.row > end.row {
        return Err(format_error(
            part,
            format!("reversed worksheet dimension: {value}"),
        ));
    }
    Ok(Dimension {
        start_column: start.column,
        start_row: start.row,
        end_column: end.column,
        end_row: end.row,
    })
}

fn parse_a1(value: &str) -> Option<CellAddress> {
    let bytes = value.as_bytes();
    let letter_count = bytes
        .iter()
        .take_while(|byte| byte.is_ascii_uppercase())
        .count();
    if letter_count == 0 || letter_count == bytes.len() {
        return None;
    }
    let mut column = 0_u32;
    for byte in &bytes[..letter_count] {
        column = column
            .checked_mul(26)?
            .checked_add(u32::from(*byte - b'A') + 1)?;
        if column > MAX_COLUMNS {
            return None;
        }
    }
    let row_text = std::str::from_utf8(&bytes[letter_count..]).ok()?;
    if row_text.starts_with('0') || !row_text.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    let row = row_text.parse::<u32>().ok()?;
    if row == 0 || row > MAX_ROWS {
        return None;
    }
    Some(CellAddress {
        column: column - 1,
        row: row - 1,
        canonical: value.to_owned(),
    })
}

fn parse_numeric_cell_address(value: &str) -> Option<CellAddress> {
    let (column, row) = value.split_once('_')?;
    if column.is_empty()
        || row.is_empty()
        || !column.bytes().all(|byte| byte.is_ascii_digit())
        || !row.bytes().all(|byte| byte.is_ascii_digit())
    {
        return None;
    }
    let column = column.parse::<u32>().ok()?.checked_sub(1)?;
    let row = row.parse::<u32>().ok()?.checked_sub(1)?;
    if column >= MAX_COLUMNS || row >= MAX_ROWS {
        return None;
    }
    Some(CellAddress {
        column,
        row,
        canonical: format_a1(column, row),
    })
}

fn format_a1(mut column: u32, row: u32) -> String {
    let mut letters = Vec::with_capacity(3);
    loop {
        letters.push((b'A' + (column % 26) as u8) as char);
        if column < 26 {
            break;
        }
        column = column / 26 - 1;
    }
    letters.reverse();
    let mut address = letters.into_iter().collect::<String>();
    address.push_str(&(row + 1).to_string());
    address
}

fn ensure_internal(relationship: &Relationship, part: &str, kind: &str) -> Result<(), Diagnostic> {
    if relationship.external {
        Err(Diagnostic::fatal(
            DiagnosticCode::ExternalResourceBlocked,
            Phase::Security,
            None,
            format!("external {kind} relationship is forbidden"),
        )
        .in_part(part))
    } else {
        Ok(())
    }
}

fn relationship_id(attributes: &[XmlAttribute<'_>], part: &str) -> Result<String, Diagnostic> {
    attributes
        .iter()
        .find(|attribute| attribute.name.contains(':') && local_name(attribute.name) == "id")
        .map(|attribute| {
            decode_xml_text(attribute.value)
                .map(|value| value.into_owned())
                .map_err(|error| with_part(error, part))
        })
        .transpose()?
        .ok_or_else(|| format_error(part, "sheet is missing its relationship id"))
}

fn required_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<String, Diagnostic> {
    optional_attribute(attributes, name, part)?
        .ok_or_else(|| format_error(part, format!("element is missing {name}")))
}

fn format_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message).in_part(part)
}

fn dimension_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(
        DiagnosticCode::LayoutBudgetExceeded,
        Phase::Layout,
        None,
        message,
    )
    .in_part(part)
}

fn object_limit_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::ObjectLimit, Phase::Parse, None, message).in_part(part)
}

fn with_part(error: Diagnostic, part: &str) -> Diagnostic {
    if error.location.part.is_some() {
        error
    } else {
        error.in_part(part)
    }
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use super::{
        AxisMetrics, CellState, CellStyle, CellValueType, ContentTypes, Dimension, FontRecord,
        SharedString, SparklineAxisMode, SparklineKind, Styles, XlsxParseState,
        builtin_number_format, default_theme_colors, derive_dimension, excel_column_width,
        excel_maximum_digit_width, format_number, parse_a1, parse_basic_chart,
        parse_cell_value_type, parse_conditional_formatting_xml, parse_data_validations_xml,
        parse_metadata, parse_numeric_cell_address, parse_shared_strings_xml,
        parse_sparkline_groups_xml, parse_sparkline_source_range, parse_styles, parse_unit,
        parse_unit_region, parse_workbook, parse_worksheet, parse_worksheet_drawing,
        parse_worksheet_metadata, parse_worksheet_print_settings, push_xlsx_chart,
        resolve_cell_text,
    };
    use crate::diagnostic::DiagnosticCode;
    use crate::format::presentation_image::stored_zip;
    use crate::limits::Limits;
    use crate::model::{
        ObjectKind, Paint, Rect, SheetOrientation, SheetViewMode, SourceLocator, TextAlign, Visual,
    };
    use crate::package::Package;

    #[test]
    fn real_dialogsheet_does_not_block_worksheets() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/Dialogsheet.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let content_types = ContentTypes::default();
        let documents = [
            super::parse(&package, "xl/workbook.xml", &content_types, None),
            parse_metadata(&package, "xl/workbook.xml", &content_types),
            parse_unit(&package, "xl/workbook.xml", &content_types, 0, None),
            parse_unit_region(&package, "xl/workbook.xml", &content_types, 2, 10, 10, None),
        ];
        for result in documents {
            let document = result.unwrap();
            assert_eq!(
                document
                    .units
                    .iter()
                    .map(|unit| unit.name.as_str())
                    .collect::<Vec<_>>(),
                ["Sheet1", "Sheet2", "Sheet3"]
            );
            assert_eq!(
                document
                    .units
                    .iter()
                    .map(|unit| unit.index)
                    .collect::<Vec<_>>(),
                [0, 1, 2]
            );
            assert!(document.diagnostics.iter().any(|diagnostic| diagnostic.code
                == DiagnosticCode::UnsupportedFeature
                && diagnostic.location.part.as_deref() == Some("xl/dialogsheets/sheet1.xml")));
        }
    }

    #[test]
    fn real_pivot_fills_follow_cached_row_roles() {
        for (bytes, unit, expected) in [
            (
                include_bytes!("../../tests/fixtures/Pivot.xlsx").as_slice(),
                0,
                vec![
                    ("A3", 0xdce6_f1ff),
                    ("B3", 0xdce6_f1ff),
                    ("A4", 0xdce6_f1ff),
                    ("A5", 0xffff_ffff),
                    ("B5", 0xffff_ffff),
                ],
            ),
            (
                include_bytes!("../../tests/fixtures/Pivot2.xlsx").as_slice(),
                0,
                vec![
                    ("A1", 0xffff_ffff),
                    ("A3", 0x3660_92ff),
                    ("B4", 0x3660_92ff),
                    ("AB5", 0x3660_92ff),
                    ("A6", 0x95b3_d7ff),
                    ("B7", 0x95b3_d7ff),
                    ("A8", 0xdce6_f1ff),
                    ("T8", 0xdce6_f1ff),
                    ("B9", 0xffff_ffff),
                    ("B10", 0x95b3_d7ff),
                    ("B11", 0xdce6_f1ff),
                    ("B12", 0xffff_ffff),
                    ("A35", 0xffff_ffff),
                    ("AB35", 0xffff_ffff),
                ],
            ),
            (
                include_bytes!("../../tests/fixtures/Pivot2.xlsx").as_slice(),
                2,
                vec![
                    ("B2", 0xdce6_f1ff),
                    ("F3", 0xdce6_f1ff),
                    ("B4", 0xffff_ffff),
                    ("D5", 0xffff_ffff),
                    ("B6", 0xffff_ffff),
                    ("B7", 0xdce6_f1ff),
                ],
            ),
        ] {
            let package = Package::open(bytes, Limits::default()).unwrap();
            for document in [
                parse_unit(
                    &package,
                    "xl/workbook.xml",
                    &ContentTypes::default(),
                    unit,
                    None,
                )
                .unwrap(),
                parse_unit_region(
                    &package,
                    "xl/workbook.xml",
                    &ContentTypes::default(),
                    unit,
                    40,
                    30,
                    None,
                )
                .unwrap(),
            ] {
                for &(address, expected_fill) in &expected {
                    let cell = document.objects.iter().find(|o| o.kind == ObjectKind::Cell &&
                        matches!(&o.source.locator, SourceLocator::Xlsx { address: Some(a), .. } if a == address))
                        .unwrap_or_else(|| panic!("missing styled cell {address}"));
                    let Visual::TextLayout { visual, .. } = &cell.visual else {
                        panic!("cell text layout")
                    };
                    let Visual::Text { fill, .. } = **visual else {
                        panic!("cell fill")
                    };
                    assert!(
                        [8, 16, 24].iter().all(|shift| ((fill >> shift) & 255)
                            .abs_diff((expected_fill >> shift) & 255)
                            <= 1),
                        "unit {unit} {address}: {fill:08x} != {expected_fill:08x}"
                    );
                }
            }
        }
    }

    #[test]
    fn real_olap_traffic_lights_leave_room_for_values() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/OlapPivotA3.xlsx"),
            Limits::default(),
        )
        .unwrap();
        for document in [
            super::parse(&package, "xl/workbook.xml", &ContentTypes::default(), None).unwrap(),
            parse_unit(
                &package,
                "xl/workbook.xml",
                &ContentTypes::default(),
                0,
                None,
            )
            .unwrap(),
            parse_unit_region(
                &package,
                "xl/workbook.xml",
                &ContentTypes::default(),
                0,
                12,
                17,
                None,
            )
            .unwrap(),
        ] {
            for (address, color) in [
                ("P4", 0x16a3_4aff),
                ("P5", 0xfacc_15ff),
                ("P6", 0xfacc_15ff),
                ("P7", 0xfacc_15ff),
                ("P8", 0xd930_25ff),
                ("P9", 0xd930_25ff),
                ("P10", 0x16a3_4aff),
            ] {
                let objects: Vec<_> = document.objects.iter().filter(|o|
                    matches!(&o.source.locator, SourceLocator::Xlsx { address: Some(a), .. } if a == address)).collect();
                let cell = objects.iter().find(|o| o.kind == ObjectKind::Cell).unwrap();
                let icon = objects
                    .iter()
                    .find(|o| o.kind == ObjectKind::Shape)
                    .unwrap();
                assert!(
                    icon.bounds.x < cell.bounds.x + cell.bounds.width / 2.0,
                    "{address}: icon must precede the value"
                );
                assert!(
                    icon.text.is_none(),
                    "traffic lights must not depend on a font baseline"
                );
                assert!(
                    matches!(icon.visual, Visual::Shape { geometry: crate::model::Geometry::Ellipse, fill, .. } if fill == color)
                );
                assert_eq!(icon.bounds.width, icon.bounds.height);
                assert!(
                    (icon.bounds.y + icon.bounds.height / 2.0
                        - cell.bounds.y
                        - cell.bounds.height / 2.0)
                        .abs()
                        < 0.01
                );
                let Visual::TextLayout { layout, visual } = &cell.visual else {
                    panic!("cell text layout")
                };
                assert!(
                    cell.bounds.x + layout.inset_left >= icon.bounds.x + icon.bounds.width + 1.0
                );
                assert_eq!(layout.inset_right, 2.0);
                assert!(matches!(
                    **visual,
                    Visual::Text {
                        align: TextAlign::End,
                        ..
                    }
                ));
                if address == "P4" {
                    assert_eq!(cell.text.as_deref(), Some("41269602.97"));
                }
            }
        }
    }

    #[test]
    fn real_ooxml_text_conditions_are_retained() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/ooxml-conditional-priority.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let bytes = package.required_part("xl/worksheets/sheet1.xml").unwrap();
        let (blocks, _) = parse_conditional_formatting_xml(
            &bytes,
            Limits::default(),
            &default_theme_colors(),
            "xl/worksheets/sheet1.xml",
        )
        .unwrap();
        assert_eq!(
            blocks.iter().map(|block| block.rules.len()).sum::<usize>(),
            2
        );
    }

    #[test]
    fn real_filter_by_color_symbols() {
        use crate::model::Geometry;
        let package = Package::open(
            include_bytes!("../../tests/fixtures/FilterByColor.xlsx"),
            Limits::default(),
        )
        .unwrap();
        for document in [
            super::parse(&package, "xl/workbook.xml", &ContentTypes::default(), None).unwrap(),
            parse_unit(
                &package,
                "xl/workbook.xml",
                &ContentTypes::default(),
                0,
                None,
            )
            .unwrap(),
            parse_unit_region(
                &package,
                "xl/workbook.xml",
                &ContentTypes::default(),
                0,
                20,
                8,
                None,
            )
            .unwrap(),
        ] {
            let icons: Vec<_> = document
                .objects
                .iter()
                .filter(|o| o.kind == ObjectKind::Shape)
                .collect();
            assert_eq!(icons.len(), 7);
            for (icon, address) in icons
                .iter()
                .zip(["H11", "H13", "H14", "H15", "H17", "H19", "H20"])
            {
                let SourceLocator::Xlsx {
                    address: Some(actual),
                    ..
                } = &icon.source.locator
                else {
                    panic!("cell source")
                };
                assert_eq!(actual, address);
                let cell = document.objects.iter().find(|o| o.kind == ObjectKind::Cell && matches!(&o.source.locator, SourceLocator::Xlsx { address: Some(a), .. } if a == address)).unwrap();
                assert!((icon.bounds.x - cell.bounds.x - 2.0).abs() < 0.01);
                assert_eq!(icon.bounds.width, 16.0);
                assert!(icon.text.is_none());
                let Visual::Shape {
                    geometry: Geometry::Path { commands, .. },
                    fill,
                    ..
                } = &icon.visual
                else {
                    panic!("symbols must be vector paths, not fallback circles")
                };
                let expected = match address {
                    "H13" => (10, 0xf0ae_4cff),
                    "H14" => (13, 0xd36a_67ff),
                    _ => (7, 0x63a3_6aff),
                };
                assert_eq!((commands.len(), *fill), expected);
            }
        }
    }

    #[test]
    fn tdf162948_conditional_arrows() {
        use super::parse;
        use crate::model::{Geometry, PathCommand};
        let package = Package::open(
            include_bytes!("../../tests/fixtures/tdf162948.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let document = parse(&package, "xl/workbook.xml", &ContentTypes::default(), None).unwrap();
        let icons: Vec<_> = document
            .objects
            .iter()
            .filter(|o| o.kind == ObjectKind::Shape)
            .collect();
        assert_eq!(icons.len(), 8);
        for (i, icon) in icons.iter().enumerate() {
            let cell = document
                .objects
                .iter()
                .filter(|o| o.kind == ObjectKind::Cell)
                .nth(i)
                .unwrap();
            assert!(
                (icon.bounds.x - cell.bounds.x - 2.0).abs() < 0.1,
                "icon belongs at the left edge"
            );
            assert!(
                (icon.bounds.y + icon.bounds.height / 2.0
                    - cell.bounds.y
                    - cell.bounds.height / 2.0)
                    .abs()
                    < 0.1
            );
            let Visual::Shape {
                geometry: Geometry::Path { commands, .. },
                fill,
                stroke,
                ..
            } = &icon.visual
            else {
                panic!("arrows must be outlined vector icons")
            };
            assert_eq!(*fill, if i == 1 { 0x70ad_47ff } else { 0xffc0_00ff });
            assert_ne!(*stroke, 0);
            let PathCommand::MoveTo { x, y } = commands[0] else {
                panic!("arrow tip")
            };
            assert!(
                if i == 1 {
                    y < 2.0 && x > 5.0
                } else {
                    x > 12.0 && y > 5.0
                },
                "A1 is right, B1 is up"
            );
        }
    }

    #[test]
    fn supplied_extended_charts_resolve_workbook_data() {
        use crate::format::drawingml::{ChartKind, parse_chart_ex_with_data};
        for (bytes, kind) in [
            (
                include_bytes!("../../tests/fixtures/SimpleHistogram.xlsx").as_slice(),
                ChartKind::Histogram,
            ),
            (
                include_bytes!("../../tests/fixtures/paretoLine.xlsx").as_slice(),
                ChartKind::Pareto,
            ),
            (
                include_bytes!("../../tests/fixtures/tdf163727_histogram_underflow_overflow.xlsx")
                    .as_slice(),
                ChartKind::Histogram,
            ),
            (
                include_bytes!("../../tests/fixtures/sunburst.xlsx").as_slice(),
                ChartKind::Sunburst,
            ),
            (
                include_bytes!("../../tests/fixtures/color_funnel.xlsx").as_slice(),
                ChartKind::Funnel,
            ),
        ] {
            let package = Package::open(bytes, Limits::default()).unwrap();
            let (sheets, _, _) = parse_workbook(&package, "xl/workbook.xml").unwrap();
            let relationships = package.relationships(Some("xl/workbook.xml")).unwrap();
            let map = relationships.iter().map(|r| (r.id.as_str(), r)).collect();
            let strings = super::parse_shared_strings(
                &package,
                "xl/sharedStrings.xml",
                &default_theme_colors(),
            )
            .unwrap_or_default();
            let sources =
                super::parse_chart_named_data(&package, "xl/workbook.xml", &sheets, &map, &strings)
                    .unwrap();
            let chart = parse_chart_ex_with_data(
                &package,
                "xl/charts/chartEx1.xml",
                |name| super::theme_color_index(name).map(|index| default_theme_colors()[index]),
                |name| sources.get(name).cloned(),
            )
            .unwrap()
            .unwrap();
            assert_eq!(chart.series[0].kind, kind);
            if kind == ChartKind::Histogram {
                assert_eq!(chart.title, "Chart Title");
            }
            if kind == ChartKind::Pareto {
                assert_eq!(chart.title, "ParetoLine");
            }
            let series = &chart.series[0];
            assert!(!series.values.is_empty());
            assert!(series.values.iter().all(|v| v.is_finite()));
            assert_eq!(series.values.len(), series.categories.len());
            if kind == ChartKind::Sunburst {
                assert!(series.category_levels.len() >= 2);
                assert!(
                    series
                        .category_levels
                        .iter()
                        .flatten()
                        .any(|s| !s.is_empty())
                );
            } else if kind == ChartKind::Funnel {
                assert!(series.fill.is_some());
                assert!(
                    series
                        .point_border_colors
                        .iter()
                        .all(|c| *c == Some(0x00b0_50ff))
                );
                assert!(series.show_values);
            } else {
                let total = sources
                    .get("_xlchart.v1.0")
                    .unwrap()
                    .values
                    .iter()
                    .filter(|v| v.is_finite())
                    .count();
                assert_eq!(series.values.iter().sum::<f32>(), total as f32);
                if kind == ChartKind::Pareto {
                    assert!(series.values.windows(2).all(|p| p[0] >= p[1]));
                }
            }
            let elements = crate::format::drawingml::chart_extended_elements(
                &chart,
                Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 600.0,
                    height: 400.0,
                },
                Limits::default().max_document_objects,
            )
            .unwrap()
            .unwrap();
            assert!(elements.len() > series.values.len());
        }
    }

    #[test]
    fn word_art_vertical_axis_titles_stack_upright_glyphs() {
        let mut axis = super::super::drawingml::ChartValueAxis::default();
        axis.title_orientation = Some(super::super::drawingml::drawingml_text_orientation(
            "wordArtVert",
        ));
        let (_, orientation, _) =
            super::xlsx_axis_title_presentation(&axis, crate::model::Rect::default(), 0.0);

        assert_eq!(orientation, crate::model::TextOrientation::StackedLr);
    }

    #[test]
    fn painted_chart_title_uses_shared_glyph_widths() {
        let slot = Rect {
            x: 0.0,
            y: 0.0,
            width: 240.0,
            height: 40.0,
        };
        let bounds = super::chart_text_tight_bounds(
            slot,
            "Gradient Fill",
            10.0,
            crate::model::TextOrientation::Horizontal,
            true,
            false,
        );
        let expected = "Gradient Fill"
            .chars()
            .map(|character| {
                super::super::drawingml::drawingml_fallback_character_width(character, 10.0)
            })
            .sum::<f32>();

        assert!((bounds.width - expected).abs() < f32::EPSILON);
    }

    #[test]
    fn painted_axis_titles_wrap_and_contain_stacked_columns() {
        let horizontal = super::chart_text_tight_bounds(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 80.0,
                height: 160.0,
            },
            "Fill & Perspective Shadow",
            10.0,
            crate::model::TextOrientation::Horizontal,
            true,
            true,
        );
        assert!(horizontal.width <= 80.0);
        assert!(horizontal.height > 14.0);

        let stacked = super::chart_text_tight_bounds(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 80.0,
                height: 160.0,
            },
            "Border Color & Style",
            10.0,
            crate::model::TextOrientation::StackedLr,
            true,
            false,
        );
        assert!(stacked.width >= 40.0);
    }

    #[test]
    fn axis_title_boxes_keep_dashes_and_shadows() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:view3D/><c:plotArea><c:bar3DChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:bar3DChart>
              <c:catAx><c:axPos val="b"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Border</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="lgDashDotDot"/></a:ln></c:spPr></c:title></c:catAx>
              <c:valAx><c:axPos val="l"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Inner</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:solidFill><a:srgbClr val="C0504D"/></a:solidFill><a:effectLst><a:innerShdw blurRad="63500" dist="50800" dir="18900000"><a:prstClr val="black"/></a:innerShdw></a:effectLst></c:spPr></c:title></c:valAx>
              <c:serAx><c:axPos val="b"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Perspective</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:solidFill><a:srgbClr val="C0504D"/></a:solidFill><a:effectLst><a:outerShdw blurRad="76200" sy="23000" kx="-1200000" algn="bl"><a:prstClr val="black"/></a:outerShdw></a:effectLst></c:spPr></c:title></c:serAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_basic_chart(&package, PART).unwrap().unwrap();
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };
        push_xlsx_chart(
            chart,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 400.0,
                height: 300.0,
            },
            1,
            "Sheet1",
            0,
            100,
            &mut state,
        )
        .unwrap();
        let visual = |text| {
            &state
                .objects
                .iter()
                .find(|object| object.text.as_deref() == Some(text))
                .unwrap()
                .visual
        };
        fn stroke_style(visual: &Visual) -> Option<&crate::model::StrokeStyle> {
            match visual {
                Visual::StrokeStyle { style, .. } => Some(style),
                Visual::Layer { visual, .. }
                | Visual::TextLayout { visual, .. }
                | Visual::Effect { visual, .. }
                | Visual::AdvancedEffect { visual, .. } => stroke_style(visual),
                _ => None,
            }
        }
        fn shadows(visual: &Visual) -> (bool, bool) {
            match visual {
                Visual::AdvancedEffect {
                    outer_shadow,
                    inner_shadow,
                    ..
                } => (outer_shadow.is_some(), inner_shadow.is_some()),
                Visual::Layer { visual, .. }
                | Visual::TextLayout { visual, .. }
                | Visual::StrokeStyle { visual, .. }
                | Visual::Effect { visual, .. } => shadows(visual),
                _ => (false, false),
            }
        }

        assert_eq!(
            stroke_style(visual("Border")).unwrap().dash,
            [16.0, 6.0, 2.0, 6.0, 2.0, 6.0]
        );
        assert_eq!(shadows(visual("Inner")), (false, true));
        assert_eq!(shadows(visual("Perspective")), (true, false));
    }

    #[test]
    fn axis_title_boxes_keep_distinct_bevels() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:view3D/><c:plotArea><c:bar3DChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:bar3DChart>
              <c:catAx><c:axPos val="b"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Top</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d><a:bevelT prst="relaxedInset"/></a:sp3d></c:spPr></c:title></c:catAx>
              <c:valAx><c:axPos val="l"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Angle</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d><a:bevelT prst="angle"/></a:sp3d></c:spPr></c:title></c:valAx>
              <c:serAx><c:axPos val="b"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Bottom</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d prstMaterial="powder"><a:bevelT/><a:bevelB w="101600" prst="riblet"/></a:sp3d></c:spPr></c:title></c:serAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_basic_chart(&package, PART).unwrap().unwrap();

        let category = chart.horizontal_axis_options.title_effects.three_d.unwrap();
        assert_eq!(category.bevel_top.unwrap().preset, "relaxedInset");
        let value = chart.value_axis_options.title_effects.three_d.unwrap();
        assert_eq!(value.bevel_top.unwrap().preset, "angle");
        let series = chart
            .series_axis_options
            .unwrap()
            .title_effects
            .three_d
            .unwrap();
        assert_eq!(series.material, "powder");
        assert_eq!(series.bevel_bottom.unwrap().preset, "riblet");
    }

    #[test]
    fn drawing_payload_storage_stays_bounded() {
        assert!(
            std::mem::size_of::<super::DrawingPayload>()
                <= std::mem::size_of::<String>() + std::mem::size_of::<usize>()
        );
    }

    #[test]
    fn joins_shared_string_runs_in_document_order() {
        let strings = parse_shared_strings_xml(
            br#"<sst><si><r><t>Hello </t></r><r><t>sheet</t></r></si></sst>"#,
            Limits::default(),
            "xl/sharedStrings.xml",
        )
        .expect("valid shared-string table");
        assert_eq!(strings.len(), 1);
        assert_eq!(strings[0].text, "Hello sheet");
    }

    #[test]
    fn preserves_shared_string_run_formatting_in_cell_visuals() {
        const STYLES_PART: &str = "xl/styles.xml";
        const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
        let styles_xml = br#"<styleSheet>
          <fonts><font><name val="Arial"/><sz val="11"/></font><font><name val="Arial"/><sz val="12"/><u/><strike/><vertAlign val="subscript"/></font></fonts>
          <fills><fill><patternFill patternType="none"/></fill></fills>
          <borders><border/></borders>
          <cellXfs><xf fontId="0" fillId="0" borderId="0"/><xf fontId="1" fillId="0" borderId="0"><alignment readingOrder="2" relativeIndent="2"/></xf></cellXfs>
        </styleSheet>"#;
        let sheet_xml = br#"<worksheet><dimension ref="A1:C1"/><sheetData><row r="1">
          <c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><r><rPr><rFont val="Aptos"/><sz val="18"/><b/><i/><u val="double"/><strike/><vertAlign val="superscript"/><color rgb="FFFF0000"/></rPr><t>Rich</t></r><r><t> text</t></r></is></c><c r="C1" s="1" t="inlineStr"><is><t>RTL style</t></is></c>
        </row></sheetData></worksheet>"#;
        let bytes = stored_zip(&[(STYLES_PART, styles_xml), (SHEET_PART, sheet_xml)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let styles = parse_styles(&package, STYLES_PART, default_theme_colors()).unwrap();
        let shared_strings = parse_shared_strings_xml(
            br#"<sst><si><r><rPr><rFont val="Aptos"/><sz val="18"/><b/><i/><u val="double"/><strike/><vertAlign val="superscript"/><color rgb="FFFF0000"/></rPr><t>Rich</t></r><r><t> text</t></r></si></sst>"#,
            Limits::default(),
            "xl/sharedStrings.xml",
        )
        .unwrap();
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };

        parse_worksheet(
            &package,
            SHEET_PART,
            "Sheet1",
            None,
            0,
            0,
            &shared_strings,
            &styles,
            false,
            None,
            &ContentTypes::default(),
            &std::collections::HashMap::new(),
            &mut state,
        )
        .unwrap();

        let cells = state
            .objects
            .iter()
            .filter(|object| object.text.as_deref() == Some("Rich text"))
            .collect::<Vec<_>>();
        assert_eq!(
            cells.len(),
            2,
            "shared and inline rich text must both render"
        );
        for cell in cells {
            let Visual::TextLayout { visual, .. } = &cell.visual else {
                panic!("cell text must use a text layout");
            };
            let Visual::RichText { runs, .. } = visual.as_ref() else {
                panic!("formatted spreadsheet text must remain rich text");
            };
            assert_eq!(runs.len(), 2);
            assert_eq!(runs[0].text, "Rich");
            assert_eq!(runs[0].font_family, "Aptos");
            assert!((runs[0].font_size - 24.0).abs() < 0.001);
            assert_eq!(runs[0].color, 0xff00_00ff);
            assert!(runs[0].bold && runs[0].italic && runs[0].underline && runs[0].strikethrough);
            assert!(runs[0].baseline_shift > 0.0);
            assert_eq!(runs[1].text, " text");
        }
        let styled = state
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("RTL style"))
            .unwrap();
        let Visual::TextLayout { layout, visual } = &styled.visual else {
            panic!("styled cell text must use a text layout");
        };
        assert_eq!(layout.direction, crate::model::TextDirection::Rtl);
        assert!(layout.inset_right > layout.inset_left);
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("styled cell underline, strike, and baseline require rich text");
        };
        assert!(runs[0].underline && runs[0].strikethrough);
        assert!(runs[0].baseline_shift < 0.0);
    }

    #[test]
    fn validates_excel_a1_addresses_and_limits() {
        let first = parse_a1("A1").expect("first cell");
        assert_eq!((first.column, first.row), (0, 0));
        let last = parse_a1("XFD1048576").expect("last cell");
        assert_eq!((last.column, last.row), (16_383, 1_048_575));
        for invalid in ["A0", "A01", "1A", "a1", "XFE1", "A1048577", "A1:B2"] {
            assert!(parse_a1(invalid).is_none(), "{invalid} must be rejected");
        }
        let numeric = parse_numeric_cell_address("32_2").expect("numeric compatibility address");
        assert_eq!(
            (numeric.column, numeric.row, numeric.canonical.as_str()),
            (31, 1, "AF2")
        );
        for invalid in ["0_1", "1_0", "1_1_1", "A_1", "16385_1", "1_1048577"] {
            assert!(parse_numeric_cell_address(invalid).is_none());
        }
    }

    #[test]
    fn converts_authored_column_widths_with_the_normal_style_digit_width() {
        let styles = Styles {
            normal_font: Some(FontRecord {
                family: "宋体".to_owned(),
                ..FontRecord::default()
            }),
            ..Styles::default()
        };
        let maximum_digit_width = excel_maximum_digit_width(&styles);

        assert_eq!(maximum_digit_width, 8.0);
        assert_eq!(
            excel_column_width(9.164_062_5, maximum_digit_width, "sheet1.xml").unwrap(),
            73.0,
        );
        assert_eq!(
            excel_column_width(28.5, maximum_digit_width, "sheet1.xml").unwrap(),
            228.0,
        );
    }

    #[test]
    fn derives_missing_dimensions_and_defaults_an_empty_sheet() {
        assert_eq!(
            derive_dimension(&[]),
            Dimension {
                start_column: 0,
                start_row: 0,
                end_column: 0,
                end_row: 0,
            }
        );
        let cells = [
            test_cell("B2", CellValueType::Number),
            test_cell("D4", CellValueType::Number),
        ];
        assert_eq!(
            derive_dimension(&cells),
            Dimension {
                start_column: 1,
                start_row: 1,
                end_column: 3,
                end_row: 3,
            }
        );
    }

    #[test]
    fn materializes_only_the_requested_visible_sheet() {
        let bytes = stored_zip(&[
            (
                "xl/workbook.xml",
                br#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="First" r:id="rId1"/><sheet name="Second" r:id="rId2"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                br#"<worksheet><dimension ref="A1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>first</t></is></c></row></sheetData></worksheet>"#,
            ),
            (
                "xl/worksheets/sheet2.xml",
                br#"<worksheet><dimension ref="A1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>second</t></is></c></row></sheetData></worksheet>"#,
            ),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let document = parse_unit(
            &package,
            "xl/workbook.xml",
            &ContentTypes::default(),
            1,
            None,
        )
        .unwrap();

        assert_eq!(
            document
                .units
                .iter()
                .map(|unit| unit.name.as_str())
                .collect::<Vec<_>>(),
            ["First", "Second"]
        );
        assert!(!document.objects.is_empty());
        assert!(document.objects.iter().all(|object| object.unit_index == 1));
        assert!(document.objects.iter().any(|object| {
            object.kind == ObjectKind::Cell && object.text.as_deref() == Some("second")
        }));
        assert!(!document.objects.iter().any(|object| {
            object.kind == ObjectKind::Cell && object.text.as_deref() == Some("first")
        }));
    }

    #[test]
    fn metadata_preserves_hidden_sheet_and_drawing_extents() {
        const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
        let hidden_bytes = stored_zip(&[(
            SHEET_PART,
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><cols><col min="1" max="1" hidden="1"/></cols><sheetData><row r="1" hidden="1"><c r="A1"/></row></sheetData></worksheet>"#,
        )]);
        let hidden_package = Package::open(&hidden_bytes, Limits::default()).unwrap();
        let hidden = parse_worksheet_metadata(
            &hidden_package,
            SHEET_PART,
            None,
            &default_theme_colors(),
            8.0,
            &[],
            &Styles::default(),
        )
        .unwrap();
        assert_eq!((hidden.width, hidden.height), (1.0, 1.0));

        let drawing_bytes = stored_zip(&[
            (
                SHEET_PART,
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/><sheetData/><drawing r:id="rIdDrawing"/></worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
            ),
            (
                "xl/drawings/drawing1.xml",
                br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="7"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor></xdr:wsDr>"#,
            ),
        ]);
        let drawing_package = Package::open(&drawing_bytes, Limits::default()).unwrap();
        let drawing = parse_worksheet_metadata(
            &drawing_package,
            SHEET_PART,
            None,
            &default_theme_colors(),
            8.0,
            &[],
            &Styles::default(),
        )
        .unwrap();
        assert_eq!(
            (drawing.columns, drawing.rows, drawing.width, drawing.height),
            (2, 2, 138.0, 48.0)
        );
    }

    #[test]
    fn resolves_common_types_and_rejects_unknown_types() {
        let mut boolean = test_cell("A1", CellValueType::Boolean);
        boolean.value_seen = true;
        boolean.value.push('1');
        assert_eq!(
            resolve_test_cell_text(&boolean, &[]),
            Some("true".to_owned())
        );

        let mut inline = test_cell("B1", CellValueType::InlineString);
        inline.inline_string_seen = true;
        inline.inline_text = "rich text".to_owned();
        assert_eq!(
            resolve_test_cell_text(&inline, &[]),
            Some("rich text".to_owned())
        );

        for (value_type, value) in [
            (CellValueType::Number, "42"),
            (CellValueType::String, "cached text"),
            (CellValueType::Error, "#DIV/0!"),
            (CellValueType::Date, "2026-07-15T00:00:00Z"),
        ] {
            let mut cell = test_cell("C1", value_type);
            cell.value_seen = true;
            cell.value = value.to_owned();
            assert_eq!(resolve_test_cell_text(&cell, &[]), Some(value.to_owned()));
        }

        let error = parse_cell_value_type(Some("unknown"), "C1", "sheet1.xml")
            .expect_err("unknown cell types must not fall back to numbers");
        assert_eq!(error.code, DiagnosticCode::UnsupportedFeature);
    }

    #[test]
    #[ignore = "requires the user-supplied CBAM XLSX and Excel-calculated golden"]
    fn renders_uncached_cbam_formulas_from_the_supplied_workbook() {
        let path = std::env::var("OFFICEVIEWER_XLSX_SOURCE").unwrap();
        let bytes = std::fs::read(path).unwrap();
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let document =
            super::parse(&package, "xl/workbook.xml", &ContentTypes::default(), None).unwrap();
        let preview =
            super::parse_preview(&package, "xl/workbook.xml", &ContentTypes::default()).unwrap();
        assert_eq!(preview.units, document.units);
        assert_eq!(document.units.len(), 14);
        assert!(document.units.iter().all(|unit| !matches!(
            unit.name.as_str(),
            "InputOutput"
                | "Parameters_Constants"
                | "Parameters_CNCodes"
                | "Translations"
                | "VersionDocumentation"
        )));
        for unit in &document.units {
            let print = unit.sheet.as_ref().expect("worksheet print settings");
            assert_eq!(print.paper_size, Some(9));
            assert_eq!(print.orientation, Some(SheetOrientation::Portrait));
            assert_eq!(print.margins.left, Some(0.7));
            assert_eq!(print.margins.right, Some(0.7));
            assert_eq!(print.margins.header, Some(0.3));
            assert_eq!(print.margins.footer, Some(0.3));
        }
        let value = document.objects.iter().find_map(|object| {
            let crate::model::SourceLocator::Xlsx {
                sheet_name,
                address,
                ..
            } = &object.source.locator
            else {
                return None;
            };
            (sheet_name == "0_Versions" && address.as_deref() == Some("C6"))
                .then_some(object.text.as_deref())
                .flatten()
        });
        assert_eq!(value, Some("Sheet \"Version history\""));
        assert!(!document.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .message
                .contains("neither a cached value nor a usable calculated result")
        }));

        let golden_path = std::env::var("OFFICEVIEWER_XLSX_EXCEL_GOLDEN").unwrap();
        let golden_bytes = std::fs::read(golden_path).unwrap();
        let golden_package = Package::open(&golden_bytes, Limits::default()).unwrap();
        let golden = super::parse(
            &golden_package,
            "xl/workbook.xml",
            &ContentTypes::default(),
            None,
        )
        .unwrap();
        let cell_text = |document: &crate::model::Document| {
            document
                .objects
                .iter()
                .filter_map(|object| {
                    if object.kind != ObjectKind::Cell {
                        return None;
                    }
                    let SourceLocator::Xlsx {
                        sheet_name,
                        address: Some(address),
                        ..
                    } = &object.source.locator
                    else {
                        return None;
                    };
                    Some(((sheet_name.clone(), address.clone()), object.text.clone()?))
                })
                .collect::<BTreeMap<_, _>>()
        };
        let actual = cell_text(&document);
        let expected = cell_text(&golden);
        let differences = actual
            .iter()
            .filter(|(cell, value)| expected.get(*cell) != Some(*value))
            .map(|(cell, value)| {
                format!(
                    "{cell:?}: actual={value:?}, expected={:?}",
                    expected.get(cell)
                )
            })
            .chain(
                expected
                    .keys()
                    .filter(|cell| !actual.contains_key(*cell))
                    .map(|cell| format!("{cell:?}: missing from actual")),
            )
            .take(100)
            .collect::<Vec<_>>();
        assert_eq!(actual.len(), expected.len(), "cell object count differs");
        assert!(
            differences.is_empty(),
            "calculated cell display differs from Excel:\n{}",
            differences.join("\n")
        );
    }

    #[test]
    #[ignore = "requires the user-supplied CBAM XLSX and Excel-calculated golden"]
    fn previews_the_first_cbam_sheet_with_excel_exact_text() {
        let source = std::fs::read(std::env::var("OFFICEVIEWER_XLSX_SOURCE").unwrap()).unwrap();
        let source_package = Package::open(&source, Limits::default()).unwrap();
        for sheet in 1..=19 {
            let part = format!("xl/worksheets/sheet{sheet}.xml");
            let bytes = source_package.required_part(&part).unwrap();
            assert!(
                super::fast_sheet_data_metadata(&bytes, &[], &Styles::default()).is_some(),
                "golden performance path rejected {part}"
            );
        }
        let preview =
            super::parse_preview(&source_package, "xl/workbook.xml", &ContentTypes::default())
                .unwrap();
        let metadata =
            parse_metadata(&source_package, "xl/workbook.xml", &ContentTypes::default()).unwrap();
        let golden =
            std::fs::read(std::env::var("OFFICEVIEWER_XLSX_EXCEL_GOLDEN").unwrap()).unwrap();
        let golden_package = Package::open(&golden, Limits::default()).unwrap();
        let expected = super::parse(
            &golden_package,
            "xl/workbook.xml",
            &ContentTypes::default(),
            None,
        )
        .unwrap();
        let first_sheet_text = |document: &crate::model::Document| {
            document
                .objects
                .iter()
                .filter_map(|object| {
                    if object.kind != ObjectKind::Cell || object.unit_index != 0 {
                        return None;
                    }
                    let SourceLocator::Xlsx {
                        address: Some(address),
                        ..
                    } = &object.source.locator
                    else {
                        return None;
                    };
                    Some((address.clone(), object.text.clone()?))
                })
                .collect::<BTreeMap<_, _>>()
        };
        assert_eq!(
            preview
                .units
                .iter()
                .map(|unit| unit.name.as_str())
                .collect::<Vec<_>>(),
            expected
                .units
                .iter()
                .map(|unit| unit.name.as_str())
                .collect::<Vec<_>>()
        );
        assert_eq!(metadata.units, preview.units);
        assert!(metadata.objects.is_empty());
        assert_eq!(first_sheet_text(&preview), first_sheet_text(&expected));
        assert!(preview.objects.iter().all(|object| object.unit_index == 0));
    }

    #[test]
    #[ignore = "requires the supplied OFCAT XLSX"]
    fn resolves_ofcat_rectangle_theme_line_widths() {
        let source = std::fs::read(std::env::var("OFFICEVIEWER_XLSX_SOURCE").unwrap()).unwrap();
        let package = Package::open(&source, Limits::default()).unwrap();
        let document =
            super::parse(&package, "xl/workbook.xml", &ContentTypes::default(), None).unwrap();
        let widths = document
            .objects
            .iter()
            .filter_map(|object| {
                let SourceLocator::Xlsx {
                    drawing_id: Some(2 | 3),
                    ..
                } = object.source.locator
                else {
                    return None;
                };
                let Visual::PaintedShape { stroke_width, .. } = object.visual else {
                    return None;
                };
                Some(stroke_width)
            })
            .collect::<Vec<_>>();

        assert_eq!(widths, vec![25_400.0 / super::EMU_PER_CSS_PIXEL; 2]);
    }

    #[test]
    #[ignore = "requires the user-supplied CBAM XLSX and Excel-calculated golden"]
    fn materializes_cbam_summary_products_viewport_with_excel_exact_text() {
        const UNIT_INDEX: u32 = 12;
        const MAX_ROW: i32 = 64;
        const MAX_COLUMN: i32 = 24;
        let source = std::fs::read(std::env::var("OFFICEVIEWER_XLSX_SOURCE").unwrap()).unwrap();
        let source_package = Package::open(&source, Limits::default()).unwrap();
        let actual = parse_unit_region(
            &source_package,
            "xl/workbook.xml",
            &ContentTypes::default(),
            UNIT_INDEX,
            MAX_ROW,
            MAX_COLUMN,
            None,
        )
        .unwrap();
        let golden =
            std::fs::read(std::env::var("OFFICEVIEWER_XLSX_EXCEL_GOLDEN").unwrap()).unwrap();
        let golden_package = Package::open(&golden, Limits::default()).unwrap();
        let expected = super::parse(
            &golden_package,
            "xl/workbook.xml",
            &ContentTypes::default(),
            None,
        )
        .unwrap();
        let viewport_text = |document: &crate::model::Document| {
            document
                .objects
                .iter()
                .filter_map(|object| {
                    if object.kind != ObjectKind::Cell || object.unit_index != UNIT_INDEX {
                        return None;
                    }
                    let SourceLocator::Xlsx {
                        address: Some(address),
                        ..
                    } = &object.source.locator
                    else {
                        return None;
                    };
                    let cell = parse_a1(address)?;
                    let text = object.text.clone()?;
                    (cell.row < MAX_ROW as u32 && cell.column < MAX_COLUMN as u32)
                        .then(|| (address.clone(), text))
                })
                .collect::<BTreeMap<_, _>>()
        };
        assert_eq!(viewport_text(&actual), viewport_text(&expected));
        assert!(!actual.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .message
                .contains("neither a cached value nor a usable calculated result")
        }));
    }

    #[test]
    #[ignore = "requires the user-supplied CBAM XLSX and Excel-calculated golden"]
    fn materializes_every_cbam_sheet_with_excel_exact_text() {
        let source = std::fs::read(std::env::var("OFFICEVIEWER_XLSX_SOURCE").unwrap()).unwrap();
        let source_package = Package::open(&source, Limits::default()).unwrap();
        let golden =
            std::fs::read(std::env::var("OFFICEVIEWER_XLSX_EXCEL_GOLDEN").unwrap()).unwrap();
        let golden_package = Package::open(&golden, Limits::default()).unwrap();
        let expected = super::parse(
            &golden_package,
            "xl/workbook.xml",
            &ContentTypes::default(),
            None,
        )
        .unwrap();
        let cell_text = |document: &crate::model::Document, unit_index: u32| {
            document
                .objects
                .iter()
                .filter_map(|object| {
                    if object.kind != ObjectKind::Cell || object.unit_index != unit_index {
                        return None;
                    }
                    let SourceLocator::Xlsx {
                        address: Some(address),
                        ..
                    } = &object.source.locator
                    else {
                        return None;
                    };
                    Some((address.clone(), object.text.clone()?))
                })
                .collect::<BTreeMap<_, _>>()
        };

        let unit_indices = match std::env::var("OFFICEVIEWER_XLSX_UNIT") {
            Ok(value) => vec![value.parse::<u32>().expect("valid zero-based unit index")],
            Err(_) => (0..expected.units.len() as u32).collect(),
        };
        for unit_index in unit_indices {
            let actual = parse_unit(
                &source_package,
                "xl/workbook.xml",
                &ContentTypes::default(),
                unit_index,
                None,
            )
            .unwrap();
            assert_eq!(
                actual
                    .units
                    .iter()
                    .map(|unit| unit.name.as_str())
                    .collect::<Vec<_>>(),
                expected
                    .units
                    .iter()
                    .map(|unit| unit.name.as_str())
                    .collect::<Vec<_>>()
            );
            assert!(
                actual
                    .objects
                    .iter()
                    .all(|object| object.unit_index == unit_index)
            );
            assert_eq!(
                cell_text(&actual, unit_index),
                cell_text(&expected, unit_index),
                "sheet {} differs from Excel; diagnostics: {:?}",
                expected.units[unit_index as usize].name,
                actual.diagnostics
            );
        }
    }

    #[test]
    fn expands_a_stale_worksheet_dimension_to_material_cells() {
        const PART: &str = "xl/worksheets/sheet1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<worksheet><dimension ref="A1"/><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>value</t></is></c></row></sheetData></worksheet>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };

        let unit = parse_worksheet(
            &package,
            PART,
            "Sheet1",
            None,
            0,
            0,
            &[],
            &Styles::default(),
            false,
            None,
            &ContentTypes::default(),
            &std::collections::HashMap::new(),
            &mut state,
        )
        .expect("stale worksheet dimensions are expanded like Office");

        assert_eq!(unit.columns, 2);
        assert_eq!(state.objects[0].text.as_deref(), Some("value"));
    }

    #[test]
    fn derives_missing_cell_addresses_and_ignores_non_column_col_elements() {
        const PART: &str = "xl/worksheets/sheet1.xml";
        let sheet = br#"<worksheet xmlns:xdr="xdr">
          <dimension ref="1:5"/>
          <sheetData><row r="1">
            <c t="inlineStr"><is><t>first</t></is></c>
            <c r="C1" t="b"><v>true</v></c>
            <c t="inlineStr"><is><t>fourth</t></is></c>
          </row></sheetData>
          <controls><anchor><from><xdr:col>7</xdr:col></from></anchor></controls>
        </worksheet>"#;
        let bytes = stored_zip(&[(PART, sheet)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };

        let unit = parse_worksheet(
            &package,
            PART,
            "Sheet1",
            None,
            0,
            0,
            &[],
            &Styles::default(),
            false,
            None,
            &ContentTypes::default(),
            &std::collections::HashMap::new(),
            &mut state,
        )
        .expect("Office-style implicit cell addresses are derived in row order");

        assert_eq!(unit.columns, 4);
        assert_eq!(
            state
                .objects
                .iter()
                .filter_map(|object| object.text.as_deref())
                .collect::<Vec<_>>(),
            ["first", "true", "fourth"],
        );
        assert!(
            state
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("dimension"))
        );
    }

    #[test]
    fn scopes_number_formats_and_substitutes_a_missing_style_font() {
        const PART: &str = "xl/styles.xml";
        let styles = br#"<styleSheet>
          <numFmts><numFmt numFmtId="164" formatCode="0.00"/></numFmts>
          <fonts><font><name val="Calibri"/></font></fonts>
          <cellXfs><xf numFmtId="164" fontId="1"/></cellXfs>
          <dxfs><dxf><numFmt numFmtId="164" formatCode="0.00"/></dxf></dxfs>
        </styleSheet>"#;
        let bytes = stored_zip(&[(PART, styles)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let parsed = parse_styles(&package, PART, default_theme_colors())
            .expect("differential number formats do not redefine the global format table");

        assert_eq!(parsed.cells[0].font_family, "Calibri");
        assert_eq!(parsed.cells[0].number_format.as_deref(), Some("0.00"));
        assert!(
            parsed
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("missing font"))
        );
    }

    #[test]
    fn diagnoses_remaining_unsupported_worksheet_layers_once_per_part() {
        const PART: &str = "xl/worksheets/sheet1.xml";
        let sheet = br#"<worksheet xmlns:r="r">
          <dimension ref="A1:D1"/>
          <mergeCells><mergeCell ref="A1:B1"/><mergeCell ref="C1:D1"/></mergeCells>
          <conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting>
          <conditionalFormatting sqref="B1"><cfRule type="expression" priority="2"><formula>B1&gt;0</formula></cfRule></conditionalFormatting>
          <picture/><picture/>
          <pivotTableParts><pivotTablePart r:id="pivot1"/><pivotTablePart r:id="pivot2"/></pivotTableParts>
          <extLst><ext><slicerList/><slicerList/><timelineRefs/></ext></extLst>
          <sheetData/>
        </worksheet>"#;
        let bytes = stored_zip(&[(PART, sheet)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };

        parse_worksheet(
            &package,
            PART,
            "Sheet1",
            None,
            0,
            0,
            &[],
            &Styles::default(),
            false,
            None,
            &ContentTypes::default(),
            &std::collections::HashMap::new(),
            &mut state,
        )
        .expect("unsupported worksheet layers are diagnosed non-fatally");

        let unsupported = state
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.code == DiagnosticCode::UnsupportedFeature)
            .collect::<Vec<_>>();
        assert_eq!(unsupported.len(), 4);
        for needle in ["drawing", "conditional", "pivot", "slicer"] {
            assert_eq!(
                unsupported
                    .iter()
                    .filter(|diagnostic| diagnostic.message.contains(needle))
                    .count(),
                1,
                "{needle} diagnostic must be deduplicated"
            );
        }
        assert!(
            unsupported
                .iter()
                .all(|diagnostic| diagnostic.location.part.as_deref() == Some(PART))
        );
    }

    #[test]
    fn ignores_zero_sized_anchors_for_unsupported_drawing_shapes() {
        const PART: &str = "xl/drawings/drawing1.xml";
        let drawing = br#"<xdr:wsDr xmlns:xdr="xdr" xmlns:a="a" xmlns:r="r">
          <xdr:twoCellAnchor>
            <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>10</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>10</xdr:rowOff></xdr:to>
            <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1"/></xdr:nvSpPr></xdr:sp>
          </xdr:twoCellAnchor>
          <xdr:twoCellAnchor>
            <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
            <xdr:pic>
              <xdr:nvPicPr><xdr:cNvPr id="2"/></xdr:nvPicPr>
              <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
            </xdr:pic>
          </xdr:twoCellAnchor>
        </xdr:wsDr>"#;
        let bytes = stored_zip(&[(PART, drawing)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let drawings = parse_worksheet_drawing(
            &package,
            PART,
            &AxisMetrics::new(64.0, Vec::new()),
            &AxisMetrics::new(20.0, Vec::new()),
            &default_theme_colors(),
            &super::DrawingMlThemeLineStyles::default(),
        )
        .expect("unsupported zero-sized shapes must not prevent supported drawings from loading");

        assert_eq!(drawings.len(), 1);
        assert_eq!(drawings[0].drawing_id, 2);
        assert_eq!(drawings[0].bounds.width, 64.0);
        assert_eq!(drawings[0].bounds.height, 20.0);
    }

    #[test]
    fn supported_chart_choice_omits_its_compatibility_fallback_shape() {
        const PART: &str = "xl/drawings/drawing1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<xdr:wsDr xmlns:xdr="xdr" xmlns:a="a" xmlns:r="r" xmlns:mc="mc" xmlns:cx="cx"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:row>8</xdr:row></xdr:to><mc:AlternateContent><mc:Choice><xdr:graphicFrame><a:graphic><a:graphicData><cx:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="9"/></xdr:nvSpPr><xdr:txBody><a:p><a:r><a:t>unsupported chart</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent></xdr:twoCellAnchor></xdr:wsDr>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let drawings = parse_worksheet_drawing(
            &package,
            PART,
            &AxisMetrics::new(64.0, Vec::new()),
            &AxisMetrics::new(20.0, Vec::new()),
            &default_theme_colors(),
            &super::DrawingMlThemeLineStyles::default(),
        )
        .unwrap();

        assert_eq!(drawings.len(), 1);
        let super::DrawingPayload::Chart(relationship_id) = &drawings[0].payload else {
            panic!("chart choice was not retained");
        };
        assert_eq!(relationship_id, "rId1");
    }

    #[test]
    fn preserves_drawingml_picture_color_adjustments() {
        const PART: &str = "xl/drawings/drawing1.xml";
        let drawing = br#"<xdr:wsDr xmlns:xdr="xdr" xmlns:a="a" xmlns:r="r">
          <xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/>
            <xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2"/></xdr:nvPicPr><xdr:blipFill>
              <a:blip r:embed="rId1"><a:grayscl/><a:duotone><a:prstClr val="black"/><a:srgbClr val="D9C3A5"><a:tint val="50000"/><a:satMod val="180000"/></a:srgbClr></a:duotone><a:lum bright="70000" contrast="-70000"/><a:biLevel thresh="50000"/></a:blip>
              <a:srcRect l="10000" t="20000" r="30000" b="40000"/>
            </xdr:blipFill></xdr:pic><xdr:clientData/>
          </xdr:absoluteAnchor>
        </xdr:wsDr>"#;
        let bytes = stored_zip(&[(PART, drawing)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let drawings = parse_worksheet_drawing(
            &package,
            PART,
            &AxisMetrics::new(64.0, Vec::new()),
            &AxisMetrics::new(20.0, Vec::new()),
            &default_theme_colors(),
            &super::DrawingMlThemeLineStyles::default(),
        )
        .expect("picture adjustments remain attached to the image payload");

        let super::DrawingPayload::Image(image) = &drawings[0].payload else {
            panic!("picture payload was lost");
        };
        assert_eq!(image.relationship_id, "rId1");
        assert_eq!(
            image.crop,
            super::ImageCrop {
                left: 0.1,
                top: 0.2,
                right: 0.3,
                bottom: 0.4
            }
        );
        assert!(image.adjustment.grayscale);
        assert_eq!(image.adjustment.bilevel_threshold, Some(0.5));
        assert_eq!(image.adjustment.brightness, 0.7);
        assert_eq!(image.adjustment.contrast, -0.7);
        assert!(image.adjustment.duotone.is_some());
    }

    #[test]
    fn chart_legend_preserves_the_series_pattern_fill() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="134"/><c:style val="34"/><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:rPr cap="all"/><a:t>All Consuming</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:ser><c:tx><c:v>Series</c:v></c:tx><c:spPr><a:pattFill prst="narHorz"><a:fgClr><a:srgbClr val="4472C4"/></a:fgClr><a:bgClr><a:srgbClr val="D9E2F3"/></a:bgClr></a:pattFill></c:spPr><c:dLbls><c:delete val="1"/></c:dLbls><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser><c:dLbls><c:showVal val="1"/></c:dLbls></c:barChart></c:plotArea><c:legend><c:legendPos val="t"/></c:legend></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_basic_chart(&package, PART).unwrap().unwrap();
        assert_eq!(chart.title, "ALL CONSUMING");
        assert_eq!(chart.plot_area_color, None);
        assert!(!chart.series[0].show_values);
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };
        push_xlsx_chart(
            chart,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 400.0,
                height: 300.0,
            },
            1,
            "Sheet1",
            0,
            100,
            &mut state,
        )
        .unwrap();

        let patterned_shapes = state
            .objects
            .iter()
            .filter(|object| {
                matches!(
                    &object.visual,
                    Visual::PaintedShape {
                        fill: Paint::Pattern { preset, .. },
                        ..
                    } if preset == "narHorz"
                )
            })
            .count();
        assert_eq!(patterned_shapes, 2, "the bar and its legend key must match");
        let title = state
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("ALL CONSUMING"))
            .unwrap();
        let legend = state
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Series"))
            .unwrap();
        assert!(title.bounds.y + title.bounds.height <= legend.bounds.y);
    }

    #[test]
    fn preserves_drawingml_picture_parallel_camera_presets() {
        const PART: &str = "xl/drawings/drawing1.xml";
        let anchors = [
            "isometricLeftDown",
            "isometricOffAxis1Top",
            "isometricOffAxis1Right",
        ]
        .into_iter()
        .enumerate()
        .map(|(index, preset)| {
            format!(
                r#"<xdr:absoluteAnchor><xdr:pos x="0" y="{}"/><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{}"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr><a:scene3d><a:camera prst="{preset}"/><a:lightRig rig="threePt" dir="t"/></a:scene3d></xdr:spPr></xdr:pic></xdr:absoluteAnchor>"#,
                index * 952_500,
                index + 1,
            )
        })
        .collect::<String>();
        let drawing =
            format!(r#"<xdr:wsDr xmlns:xdr="xdr" xmlns:a="a" xmlns:r="r">{anchors}</xdr:wsDr>"#);
        let bytes = stored_zip(&[(PART, drawing.as_bytes())]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let drawings = parse_worksheet_drawing(
            &package,
            PART,
            &AxisMetrics::new(64.0, Vec::new()),
            &AxisMetrics::new(20.0, Vec::new()),
            &default_theme_colors(),
            &super::DrawingMlThemeLineStyles::default(),
        )
        .unwrap();
        let presets = drawings
            .iter()
            .map(|drawing| match &drawing.payload {
                super::DrawingPayload::Image(image) => image
                    .effects
                    .three_d
                    .as_ref()
                    .map(|style| style.camera_preset.as_str()),
                _ => None,
            })
            .collect::<Vec<_>>();

        assert_eq!(
            presets,
            [
                Some("isometricLeftDown"),
                Some("isometricOffAxis1Top"),
                Some("isometricOffAxis1Right"),
            ]
        );
    }

    #[test]
    fn preserves_drawingml_shape_top_bevel() {
        const PART: &str = "xl/drawings/drawing1.xml";
        let drawing = br#"<xdr:wsDr xmlns:xdr="xdr" xmlns:a="a"><xdr:absoluteAnchor>
          <xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="476250"/><xdr:sp>
            <xdr:nvSpPr><xdr:cNvPr id="1"/></xdr:nvSpPr>
            <xdr:spPr><a:prstGeom prst="rect"/><a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d><a:bevelT w="152400" h="50800" prst="softRound"/></a:sp3d></xdr:spPr>
          </xdr:sp></xdr:absoluteAnchor></xdr:wsDr>"#;
        let bytes = stored_zip(&[(PART, drawing)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let drawings = parse_worksheet_drawing(
            &package,
            PART,
            &AxisMetrics::new(64.0, Vec::new()),
            &AxisMetrics::new(20.0, Vec::new()),
            &default_theme_colors(),
            &super::DrawingMlThemeLineStyles::default(),
        )
        .unwrap();
        let super::DrawingPayload::Shape(shape) = &drawings[0].payload else {
            panic!("shape payload was lost");
        };
        let bevel = shape
            .three_d
            .as_ref()
            .and_then(|style| style.bevel_top.as_ref())
            .expect("top bevel");

        assert_eq!(bevel.preset, "softRound");
        assert!((bevel.width - 16.0).abs() < 0.001);
        assert!((bevel.height - 16.0 / 3.0).abs() < 0.001);
    }

    #[test]
    fn expands_worksheet_extent_for_drawingml_camera_and_transformed_shadow() {
        let bounds = crate::model::Rect {
            x: 100.0,
            y: 50.0,
            width: 600.0,
            height: 380.0,
        };
        let mut effects = super::DrawingMlPictureEffects::default();
        effects.three_d = Some(crate::model::ThreeDStyle {
            camera_preset: "isometricLeftDown".to_owned(),
            camera_latitude: 35.0,
            camera_longitude: 45.0,
            ..crate::model::ThreeDStyle::default()
        });
        effects.outer_shadow = Some(crate::model::OuterShadow {
            shadow: crate::model::Shadow {
                color: 0x0000_0040,
                blur: 4.0,
                offset_x: 10.0,
                offset_y: 0.0,
            },
            scale_x: 2.0,
            scale_y: 1.0,
            skew_x: 0.0,
            skew_y: 0.0,
            alignment: 4,
        });
        let drawing = super::WorksheetDrawing {
            bounds,
            drawing_id: 1,
            payload: super::DrawingPayload::Image(Box::new(super::WorksheetImage {
                relationship_id: "rId1".to_owned(),
                crop: crate::model::ImageCrop::default(),
                adjustment: crate::model::ImageAdjustment::default(),
                geometry: None,
                effects,
                rotation_degrees: 0.0,
                flip_h: false,
                flip_v: false,
            })),
        };

        let (right, bottom) = super::worksheet_drawing_visual_extent(&drawing);
        assert!(right > bounds.x + bounds.width + 250.0);
        assert!(bottom > bounds.y + bounds.height + 50.0);
    }

    #[test]
    fn preserves_drawingml_text_camera_and_gradient_outline_without_a_shape_line() {
        const PART: &str = "xl/drawings/drawing1.xml";
        let drawing = br#"<xdr:wsDr xmlns:xdr="xdr" xmlns:a="a"><xdr:absoluteAnchor>
          <xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="476250"/><xdr:sp>
            <xdr:nvSpPr><xdr:cNvPr id="1"/></xdr:nvSpPr>
            <xdr:spPr><a:prstGeom prst="rect"/><a:noFill/></xdr:spPr>
            <xdr:txBody><a:bodyPr><a:scene3d><a:camera prst="isometricTopUp"/><a:lightRig rig="threePt" dir="t"/></a:scene3d></a:bodyPr>
              <a:p><a:r><a:rPr sz="1200"><a:ln w="19050"><a:gradFill><a:gsLst><a:gs pos="70000"><a:schemeClr val="accent6"><a:shade val="50000"/></a:schemeClr></a:gs><a:gs pos="0"><a:schemeClr val="accent6"><a:tint val="77000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></a:ln><a:solidFill><a:srgbClr val="FBE5D6"/></a:solidFill></a:rPr><a:t>Text</a:t></a:r></a:p>
            </xdr:txBody>
          </xdr:sp></xdr:absoluteAnchor></xdr:wsDr>"#;
        let bytes = stored_zip(&[(PART, drawing)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let drawings = parse_worksheet_drawing(
            &package,
            PART,
            &AxisMetrics::new(64.0, Vec::new()),
            &AxisMetrics::new(20.0, Vec::new()),
            &default_theme_colors(),
            &super::DrawingMlThemeLineStyles::default(),
        )
        .unwrap();
        let super::DrawingPayload::Shape(shape) = &drawings[0].payload else {
            panic!("text shape payload was lost");
        };

        assert_eq!(shape.stroke, super::Paint::None);
        let Some(super::ChartFill::LinearGradient {
            angle_degrees,
            stops,
            ..
        }) = &shape.text_stroke_fill
        else {
            panic!("text gradient outline was flattened");
        };
        assert_eq!(*angle_degrees, 90.0);
        assert_eq!(stops.len(), 2);
        assert_eq!(stops[0].offset, 0.0);
        assert_eq!(stops[1].offset, 0.7);
        assert_ne!(stops[0].color, stops[1].color);
        assert_eq!(shape.text_stroke_width, 2.0);
        assert_eq!(
            shape
                .three_d
                .as_ref()
                .map(|style| style.camera_preset.as_str()),
            Some("isometricTopUp")
        );
    }

    #[test]
    fn preserves_supported_worksheet_shape_payloads() {
        const PART: &str = "xl/drawings/drawing1.xml";
        let drawing = br#"<xdr:wsDr xmlns:xdr="xdr" xmlns:a="a">
          <xdr:twoCellAnchor>
            <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
            <xdr:sp>
              <xdr:nvSpPr><xdr:cNvPr id="7" name="Project Manager"/></xdr:nvSpPr>
              <xdr:spPr>
                <a:prstGeom prst="roundRect"/>
                <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
                <a:effectLst><a:outerShdw sx="90000" sy="-19000"><a:prstClr val="black"><a:alpha val="10000"/></a:prstClr></a:outerShdw></a:effectLst>
              </xdr:spPr>
              <xdr:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1200" b="1"/><a:t>Project Manager</a:t></a:r></a:p></xdr:txBody>
            </xdr:sp>
          </xdr:twoCellAnchor>
          <xdr:twoCellAnchor>
            <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
            <xdr:sp>
              <xdr:nvSpPr><xdr:cNvPr id="8" name="Arrow"/></xdr:nvSpPr>
              <xdr:spPr>
                <a:xfrm flipV="1"><a:off x="0" y="0"/><a:ext cx="1219200" cy="190500"/></a:xfrm>
                <a:prstGeom prst="straightConnector1"/>
                <a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/><a:tailEnd type="arrow"/></a:ln>
              </xdr:spPr>
            </xdr:sp>
          </xdr:twoCellAnchor>
        </xdr:wsDr>"#;
        let bytes = stored_zip(&[(PART, drawing)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let drawings = parse_worksheet_drawing(
            &package,
            PART,
            &AxisMetrics::new(64.0, Vec::new()),
            &AxisMetrics::new(20.0, Vec::new()),
            &default_theme_colors(),
            &super::DrawingMlThemeLineStyles::default(),
        )
        .expect("supported worksheet shapes remain renderable");

        assert_eq!(drawings.len(), 2);
        assert_eq!(drawings[0].drawing_id, 7);
        let super::DrawingPayload::Shape(shape) = &drawings[0].payload else {
            panic!("worksheet shape payload was lost");
        };
        assert_eq!(shape.fill, super::Paint::Solid(0x4472_c4ff));
        assert_eq!(
            shape.outer_shadow.as_ref().map(|effect| effect.scale_y),
            Some(-0.19)
        );
        let super::DrawingPayload::Shape(connector) = &drawings[1].payload else {
            panic!("worksheet connector payload was lost");
        };
        assert_eq!(
            connector.connector_preset.as_deref(),
            Some("straightConnector1")
        );
        assert_eq!(connector.dash, vec![8.0, 6.0]);
        assert!(connector.tail_arrow);
        assert!(connector.flip_v);
        assert_eq!(connector.fill, super::Paint::None);
    }

    #[test]
    fn resolves_worksheet_shape_theme_line_widths_without_overriding_explicit_widths() {
        const THEME_PART: &str = "xl/theme/theme1.xml";
        const DRAWING_PART: &str = "xl/drawings/drawing1.xml";
        let bytes = stored_zip(&[
            (
                THEME_PART,
                br#"<a:theme xmlns:a="a"><a:themeElements><a:fmtScheme><a:lnStyleLst><a:ln w="9525"/><a:ln w="25400"/></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>"#,
            ),
            (
                DRAWING_PART,
                br#"<xdr:wsDr xmlns:xdr="xdr" xmlns:a="a">
                  <xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="476250"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr><xdr:style><a:lnRef idx="2"/></xdr:style></xdr:sp></xdr:absoluteAnchor>
                  <xdr:absoluteAnchor><xdr:pos x="952500" y="0"/><xdr:ext cx="952500" cy="476250"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/><a:ln w="19050"/></xdr:spPr><xdr:style><a:lnRef idx="2"/></xdr:style></xdr:sp></xdr:absoluteAnchor>
                </xdr:wsDr>"#,
            ),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let (colors, line_styles) = super::parse_theme(&package, THEME_PART).unwrap();
        let drawings = parse_worksheet_drawing(
            &package,
            DRAWING_PART,
            &AxisMetrics::new(64.0, Vec::new()),
            &AxisMetrics::new(20.0, Vec::new()),
            &colors,
            &line_styles,
        )
        .unwrap();
        let widths = drawings
            .iter()
            .filter_map(|drawing| match &drawing.payload {
                super::DrawingPayload::Shape(shape) => Some(shape.stroke_width),
                _ => None,
            })
            .collect::<Vec<_>>();

        assert_eq!(widths, [25_400.0 / super::EMU_PER_CSS_PIXEL, 2.0]);
    }

    #[test]
    fn preserves_worksheet_smartart_relationship_for_diagnostics() {
        const PART: &str = "xl/drawings/drawing1.xml";
        let drawing = br#"<xdr:wsDr xmlns:xdr="xdr" xmlns:a="a" xmlns:dgm="dgm" xmlns:r="r">
          <xdr:twoCellAnchor>
            <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
            <xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="7"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData><dgm:relIds r:dm="rIdDiagram"/></a:graphicData></a:graphic></xdr:graphicFrame>
          </xdr:twoCellAnchor>
        </xdr:wsDr>"#;
        let bytes = stored_zip(&[(PART, drawing)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let drawings = parse_worksheet_drawing(
            &package,
            PART,
            &AxisMetrics::new(64.0, Vec::new()),
            &AxisMetrics::new(20.0, Vec::new()),
            &default_theme_colors(),
            &super::DrawingMlThemeLineStyles::default(),
        )
        .expect("SmartArt relationship remains available for best-effort diagnostics");

        assert_eq!(drawings.len(), 1);
        assert!(matches!(
            &drawings[0].payload,
            super::DrawingPayload::Diagram(relationship) if relationship == "rIdDiagram"
        ));
    }

    #[test]
    fn real_org_chart_xlsx_adapter_preserves_shared_geometry_and_sources() {
        for (bytes, three_d) in [
            (
                &include_bytes!("../../tests/fixtures/smartart-orgchart.pptx")[..],
                false,
            ),
            (
                &include_bytes!("../../tests/fixtures/smartart-orgchart-3d.pptx")[..],
                true,
            ),
        ] {
            let package = Package::open(bytes, Limits::default()).unwrap();
            let relationships = package
                .relationships(Some("ppt/slides/slide1.xml"))
                .unwrap();
            let map = relationships.iter().map(|r| (r.id.as_str(), r)).collect();
            let diagram = crate::format::drawingml::parse_diagram(
                &package,
                "ppt/diagrams/data1.xml",
                &map,
                Some("ppt/diagrams/colors1.xml"),
                Some("ppt/theme/theme1.xml"),
                |name| match name {
                    "accent1" => Some(0x1234_56ff),
                    "lt1" => Some(0xffff_ffff),
                    _ => None,
                },
            )
            .unwrap();
            let bounds = crate::model::Rect {
                x: 40.0,
                y: 80.0,
                width: 600.0,
                height: 360.0,
            };
            let expected =
                crate::format::drawingml::diagram_org_chart_elements(&diagram, bounds, 0).unwrap();
            let mut state = XlsxParseState {
                objects: Vec::new(),
                diagnostics: Vec::new(),
                materialized_text_bytes: 0,
                materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
                materialized_image_bytes: 0,
                materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
            };
            super::push_xlsx_diagram(
                &diagram,
                bounds,
                7,
                "Sheet1",
                0,
                "xl/drawings/drawing1.xml",
                &default_theme_colors(),
                100,
                &mut state,
            )
            .unwrap();
            assert_eq!(state.objects.len(), expected.len() + 1);
            for (object, element) in state.objects.iter().skip(1).zip(expected) {
                assert_eq!(object.bounds, element.bounds);
                assert_eq!(object.parent_numeric_id, Some(0));
                assert_eq!(object.source.part, "xl/drawings/drawing1.xml");
                if object.text.is_some() {
                    if three_d {
                        assert!(
                            matches!(&object.visual, Visual::AdvancedEffect { three_d: Some(style), .. } if style.material == "plastic")
                        );
                        continue;
                    }
                    // Flat quick styles still carry an explicit front camera.
                    let visual = match &object.visual {
                        Visual::AdvancedEffect {
                            visual,
                            three_d: Some(style),
                            outer_shadow: None,
                            ..
                        } => {
                            assert!(style.bevel_top.is_none());
                            assert_eq!(style.extrusion_height, 0.0);
                            visual.as_ref()
                        }
                        visual => visual,
                    };
                    let Visual::TextLayout { visual, .. } = visual else {
                        panic!("text layout")
                    };
                    assert!(matches!(
                        visual.as_ref(),
                        Visual::RichText {
                            fill: Paint::Solid(0x1234_56ff),
                            ..
                        }
                    ));
                }
            }
        }
    }

    #[test]
    fn xlsx_smartart_adapter_materializes_supplied_target_list() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/complex2005_12rtm.docx"),
            Limits::default(),
        )
        .unwrap();
        let diagram = crate::format::drawingml::parse_diagram(
            &package,
            "word/diagrams/data1.xml",
            &Default::default(),
            None,
            None,
            |_| None,
        )
        .unwrap();
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };
        super::push_xlsx_diagram(
            &diagram,
            crate::model::Rect {
                x: 40.0,
                y: 80.0,
                width: 600.0,
                height: 360.0,
            },
            7,
            "Sheet1",
            0,
            "xl/drawings/drawing1.xml",
            &default_theme_colors(),
            100,
            &mut state,
        )
        .unwrap();
        assert_eq!(state.objects.len(), 12);
        assert_eq!(
            state
                .objects
                .iter()
                .filter(|o| matches!(
                    o.visual,
                    Visual::PaintedShape {
                        geometry: crate::model::Geometry::Ellipse,
                        ..
                    }
                ))
                .count(),
            3
        );
        assert!(
            state
                .objects
                .iter()
                .skip(1)
                .all(|o| o.parent_numeric_id == Some(0))
        );
        assert_eq!(
            state
                .objects
                .iter()
                .find(|o| o.text.as_deref() == Some("SMART ART!"))
                .unwrap()
                .bounds
                .x,
            40.0 + 180.0 + 10.8
        );
    }

    #[test]
    fn xlsx_smartart_adapter_materializes_shared_horizontal_list() {
        let diagram = super::XlsxDiagram {
            data_part: "xl/diagrams/data1.xml".to_owned(),
            drawing_parts: Vec::new(),
            drawing_text_colors: Default::default(),
            right_to_left: false,
            role_fills: Default::default(),
            diagnostics: Vec::new(),
            role_line_colors: Default::default(),
            layout_type: Some(
                "urn:microsoft.com/office/officeart/2005/8/layout/hList7#1".to_owned(),
            ),
            scene_three_d: false,
            background_fill: None,
            background_shadow: None,
            nodes: (0..2)
                .map(|index| crate::format::drawingml::DiagramNode {
                    model_id: index.to_string(),
                    sibling_order: index,
                    assistant: false,
                    hierarchy_branch: None,
                    parent_transition_id: None,
                    bold: false,
                    font_size: None,
                    parent_id: None,
                    text: String::new(),
                    placeholder: true,
                    preset_geometry: None,
                    custom_geometry: false,
                    fill: Some(crate::format::drawingml::ChartFill::Solid(0xff00_00ff)),
                    no_fill: false,
                    shadow: None,
                    three_d: None,
                    fill_scheme: None,
                    fill_transforms: Vec::new(),
                })
                .collect(),
        };
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };

        super::push_xlsx_diagram(
            &diagram,
            crate::model::Rect {
                x: 0.0,
                y: 0.0,
                width: 600.0,
                height: 400.0,
            },
            1,
            "Sheet1",
            0,
            "xl/drawings/drawing1.xml",
            &default_theme_colors(),
            100,
            &mut state,
        )
        .unwrap();

        assert_eq!(state.objects.len(), 8);
        assert_eq!(
            state
                .objects
                .iter()
                .filter(|object| object.text.as_deref() == Some("[Text]"))
                .count(),
            2,
        );
        assert_eq!(
            state
                .objects
                .iter()
                .filter(|object| matches!(
                    object.visual,
                    Visual::PaintedShape {
                        geometry: crate::model::Geometry::Ellipse,
                        ..
                    }
                ))
                .count(),
            2,
        );
    }

    #[test]
    fn xlsx_smartart_adapter_materializes_shared_vertical_list_roots() {
        let diagram = super::XlsxDiagram {
            data_part: "xl/diagrams/data1.xml".to_owned(),
            drawing_parts: Vec::new(),
            drawing_text_colors: Default::default(),
            right_to_left: false,
            role_fills: Default::default(),
            diagnostics: Vec::new(),
            role_line_colors: Default::default(),
            layout_type: Some(
                "urn:microsoft.com/office/officeart/2005/8/layout/vList6#1".to_owned(),
            ),
            scene_three_d: false,
            background_fill: None,
            background_shadow: None,
            nodes: (0..6)
                .map(|index| crate::format::drawingml::DiagramNode {
                    model_id: index.to_string(),
                    sibling_order: index,
                    assistant: false,
                    hierarchy_branch: None,
                    parent_transition_id: None,
                    bold: false,
                    font_size: None,
                    parent_id: match index {
                        1 | 2 => Some("0".to_owned()),
                        4 | 5 => Some("3".to_owned()),
                        _ => None,
                    },
                    text: String::new(),
                    placeholder: true,
                    preset_geometry: None,
                    custom_geometry: false,
                    fill: None,
                    no_fill: false,
                    shadow: None,
                    three_d: None,
                    fill_scheme: None,
                    fill_transforms: Vec::new(),
                })
                .collect(),
        };
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };

        super::push_xlsx_diagram(
            &diagram,
            crate::model::Rect {
                x: 0.0,
                y: 0.0,
                width: 600.0,
                height: 400.0,
            },
            1,
            "Sheet1",
            0,
            "xl/drawings/drawing1.xml",
            &default_theme_colors(),
            100,
            &mut state,
        )
        .unwrap();

        assert_eq!(state.objects.len(), 5);
        assert_eq!(
            state
                .objects
                .iter()
                .filter(|object| matches!(
                    object.visual,
                    Visual::PaintedShape {
                        geometry: crate::model::Geometry::RoundedRectangle { .. },
                        ..
                    }
                ))
                .count(),
            2,
        );
        assert!(state.objects.iter().all(|object| object.text.is_none()));
    }

    #[test]
    fn materializes_inherited_row_and_column_fills_without_hidden_gridlines() {
        const STYLES_PART: &str = "xl/styles.xml";
        const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
        let styles_xml = br#"<styleSheet>
          <fonts><font><name val="Arial"/></font></fonts>
          <fills>
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>
          </fills>
          <borders><border/></borders>
          <cellXfs>
            <xf fontId="0" fillId="0" borderId="0"/>
            <xf fontId="0" fillId="1" borderId="0"/>
          </cellXfs>
        </styleSheet>"#;
        let sheet_xml = br#"<worksheet>
          <dimension ref="A1:C2"/>
          <sheetViews><sheetView showGridLines="0"/></sheetViews>
          <cols><col min="3" max="3" style="1"/></cols>
          <sheetData><row r="1" s="1" customFormat="1"/></sheetData>
        </worksheet>"#;
        let bytes = stored_zip(&[(STYLES_PART, styles_xml), (SHEET_PART, sheet_xml)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let styles = parse_styles(&package, STYLES_PART, default_theme_colors()).unwrap();
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };

        parse_worksheet(
            &package,
            SHEET_PART,
            "Sheet1",
            None,
            0,
            0,
            &[],
            &styles,
            false,
            None,
            &ContentTypes::default(),
            &std::collections::HashMap::new(),
            &mut state,
        )
        .unwrap();

        assert_eq!(state.objects.len(), 4);
        for object in &state.objects {
            let Visual::TextLayout { visual, .. } = &object.visual else {
                panic!("inherited formatted cell must retain text layout");
            };
            let Visual::Text {
                fill, stroke_width, ..
            } = visual.as_ref()
            else {
                panic!("inherited formatted cell must remain a text cell");
            };
            assert_eq!(*fill, 0xff00_00ff);
            assert_eq!(*stroke_width, 0.0);
        }
    }

    #[test]
    fn paints_empty_cells_before_overflowing_text() {
        const STYLES_PART: &str = "xl/styles.xml";
        const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
        let styles_xml = br#"<styleSheet>
          <fonts><font><name val="Arial"/></font></fonts>
          <fills><fill><patternFill patternType="none"/></fill></fills>
          <borders><border/></borders>
          <cellXfs><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
        </styleSheet>"#;
        let sheet_xml = br#"<worksheet><dimension ref="A1:B1"/><sheetData><row r="1">
          <c r="A1" t="inlineStr"><is><t>Text extending into B1</t></is></c><c r="B1"/>
        </row></sheetData></worksheet>"#;
        let bytes = stored_zip(&[(STYLES_PART, styles_xml), (SHEET_PART, sheet_xml)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let styles = parse_styles(&package, STYLES_PART, default_theme_colors()).unwrap();
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };

        parse_worksheet(
            &package,
            SHEET_PART,
            "Sheet1",
            None,
            0,
            0,
            &[],
            &styles,
            false,
            None,
            &ContentTypes::default(),
            &std::collections::HashMap::new(),
            &mut state,
        )
        .unwrap();

        assert_eq!(state.objects.len(), 2);
        assert_eq!(state.objects[0].text, None);
        assert_eq!(
            state.objects[1].text.as_deref(),
            Some("Text extending into B1")
        );
    }

    #[test]
    fn keeps_sheet_space_for_text_overflowing_the_last_used_column() {
        const STYLES_PART: &str = "xl/styles.xml";
        const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
        let styles_xml = br#"<styleSheet>
          <fonts><font><name val="Arial"/><sz val="11"/></font></fonts>
          <fills><fill><patternFill patternType="none"/></fill></fills>
          <borders><border/></borders>
          <cellXfs><xf fontId="0" fillId="0" borderId="0"/></cellXfs>
        </styleSheet>"#;
        let sheet_xml = br#"<worksheet><dimension ref="A1"/><cols><col min="1" max="1" width="5"/></cols><sheetData><row r="1">
          <c r="A1" t="inlineStr"><is><t>image with a description</t></is></c>
        </row></sheetData></worksheet>"#;
        let bytes = stored_zip(&[(STYLES_PART, styles_xml), (SHEET_PART, sheet_xml)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let styles = parse_styles(&package, STYLES_PART, default_theme_colors()).unwrap();
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };

        let unit = parse_worksheet(
            &package,
            SHEET_PART,
            "Sheet1",
            None,
            0,
            0,
            &[],
            &styles,
            false,
            None,
            &ContentTypes::default(),
            &std::collections::HashMap::new(),
            &mut state,
        )
        .unwrap();

        let cell = state
            .objects
            .iter()
            .find(|object| object.text.is_some())
            .unwrap();
        assert!(unit.columns > 1);
        assert!(unit.width > cell.bounds.x + cell.bounds.width);
    }

    #[test]
    fn applies_xlsx_cell_text_layout_attributes() {
        const STYLES_PART: &str = "xl/styles.xml";
        const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
        let styles_xml = br#"<styleSheet>
          <fonts><font><name val="Arial"/><sz val="11"/></font></fonts>
          <fills><fill><patternFill patternType="none"/></fill></fills>
          <borders><border/></borders>
          <cellXfs><xf fontId="0" fillId="0" borderId="0"><alignment horizontal="left" vertical="center" wrapText="1" indent="1"/></xf></cellXfs>
        </styleSheet>"#;
        let sheet_xml = br#"<worksheet><dimension ref="A1"/><sheetData><row r="1" ht="30"><c r="A1" s="0" t="inlineStr"><is><t>Wrapped cell text</t></is></c></row></sheetData></worksheet>"#;
        let bytes = stored_zip(&[(STYLES_PART, styles_xml), (SHEET_PART, sheet_xml)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let styles = parse_styles(&package, STYLES_PART, default_theme_colors()).unwrap();
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };

        parse_worksheet(
            &package,
            SHEET_PART,
            "Sheet1",
            None,
            0,
            0,
            &[],
            &styles,
            false,
            None,
            &ContentTypes::default(),
            &std::collections::HashMap::new(),
            &mut state,
        )
        .unwrap();

        assert!(matches!(state.objects[0].visual, Visual::TextLayout { .. }));
    }

    #[test]
    fn applies_general_alignment_by_cell_value_type() {
        const STYLES_PART: &str = "xl/styles.xml";
        const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
        let styles_xml = br#"<styleSheet>
          <fonts><font><name val="Arial"/><sz val="11"/></font></fonts>
          <fills><fill><patternFill patternType="none"/></fill></fills>
          <borders><border/></borders>
          <cellXfs>
            <xf fontId="0" fillId="0" borderId="0"/>
            <xf fontId="0" fillId="0" borderId="0"><alignment horizontal="left"/></xf>
          </cellXfs>
        </styleSheet>"#;
        let sheet_xml = br#"<worksheet><dimension ref="A1:C1"/><sheetData><row r="1">
          <c r="A1"><v>1</v></c>
          <c r="B1" t="inlineStr"><is><t>Text</t></is></c>
          <c r="C1" s="1"><v>2</v></c>
        </row></sheetData></worksheet>"#;
        let bytes = stored_zip(&[(STYLES_PART, styles_xml), (SHEET_PART, sheet_xml)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let styles = parse_styles(&package, STYLES_PART, default_theme_colors()).unwrap();
        let mut state = XlsxParseState {
            objects: Vec::new(),
            diagnostics: Vec::new(),
            materialized_text_bytes: 0,
            materialized_text_limit: Limits::default().max_total_uncompressed_bytes,
            materialized_image_bytes: 0,
            materialized_image_limit: Limits::default().max_total_uncompressed_bytes,
        };

        parse_worksheet(
            &package,
            SHEET_PART,
            "Sheet1",
            None,
            0,
            0,
            &[],
            &styles,
            false,
            None,
            &ContentTypes::default(),
            &std::collections::HashMap::new(),
            &mut state,
        )
        .unwrap();

        let alignments = state
            .objects
            .iter()
            .map(|object| {
                let Visual::TextLayout { visual, .. } = &object.visual else {
                    panic!("cell must retain text layout");
                };
                let Visual::Text { align, .. } = visual.as_ref() else {
                    panic!("cell must retain a text visual");
                };
                *align
            })
            .collect::<Vec<_>>();
        assert_eq!(
            alignments,
            [TextAlign::End, TextAlign::Start, TextAlign::Start]
        );
    }

    #[test]
    fn ignores_extension_conditional_formatting_with_child_sqref() {
        const PART: &str = "xl/worksheets/sheet1.xml";
        let sheet = br#"<worksheet xmlns:x14="x14" xmlns:xm="xm">
          <conditionalFormatting sqref="A1">
            <cfRule type="dataBar" priority="1">
              <dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF4472C4"/></dataBar>
            </cfRule>
          </conditionalFormatting>
          <extLst><ext><x14:conditionalFormattings>
            <x14:conditionalFormatting>
              <x14:cfRule type="dataBar"><x14:dataBar/></x14:cfRule>
              <xm:sqref>A1</xm:sqref>
            </x14:conditionalFormatting>
          </x14:conditionalFormattings></ext></extLst>
        </worksheet>"#;

        let (blocks, unsupported) = parse_conditional_formatting_xml(
            sheet,
            Limits::default(),
            &default_theme_colors(),
            PART,
        )
        .expect("extension conditional formatting must not shadow the worksheet element");

        assert_eq!(blocks.len(), 1);
        assert_eq!(blocks[0].rules.len(), 1);
        assert!(unsupported);
    }

    #[test]
    fn ignores_extension_data_validation_with_child_sqref() {
        const PART: &str = "xl/worksheets/sheet1.xml";
        let sheet = br#"<worksheet xmlns:x14="x14" xmlns:xm="xm">
          <dataValidations>
            <dataValidation type="list" sqref="A1"><formula1>&quot;A,B&quot;</formula1></dataValidation>
          </dataValidations>
          <extLst><ext><x14:dataValidations>
            <x14:dataValidation type="list">
              <x14:formula1><xm:f>&quot;A,B&quot;</xm:f></x14:formula1>
              <xm:sqref>A1</xm:sqref>
            </x14:dataValidation>
          </x14:dataValidations></ext></extLst>
        </worksheet>"#;

        let (validation_count, unsupported) =
            parse_data_validations_xml(sheet, Limits::default(), PART)
                .expect("extension data validation must not shadow the worksheet element");

        assert_eq!(validation_count, 1);
        assert!(unsupported);
    }

    #[test]
    fn parses_data_validations_and_sparkline_groups_without_formula_evaluation() {
        const PART: &str = "xl/worksheets/sheet1.xml";
        let sheet = br#"<worksheet>
          <dataValidations>
            <dataValidation type="list" allowBlank="1" prompt="Choose" error="Invalid" sqref="A1"><formula1>&quot;A,B&quot;</formula1></dataValidation>
            <dataValidation type="whole" sqref="B1"><formula1>1</formula1></dataValidation>
            <dataValidation type="decimal" sqref="C1"/><dataValidation type="date" sqref="D1"/>
            <dataValidation type="time" sqref="E1"/><dataValidation type="textLength" sqref="F1"/>
            <dataValidation type="custom" sqref="G1"><formula1>EXECUTE_ARBITRARY_FORMULA(G1)</formula1></dataValidation>
          </dataValidations>
          <sparklineGroups>
            <sparklineGroup type="line" minAxisType="custom" manualMin="0" maxAxisType="custom" manualMax="30" displayXAxis="1">
              <colorSeries rgb="FF4472C4"/><sparklines><sparkline><f>'Data'!$A$1:$C$1</f><sqref>$D$1</sqref></sparkline></sparklines>
            </sparklineGroup>
            <sparklineGroup type="column"><sparklines><sparkline><f>Data!A2:C2</f><sqref>D2</sqref></sparkline></sparklines></sparklineGroup>
            <sparklineGroup type="stacked"><sparklines><sparkline><f>Data!A3:C3</f><sqref>D3</sqref></sparkline></sparklines></sparklineGroup>
          </sparklineGroups>
        </worksheet>"#;

        let (validation_count, unsupported) =
            parse_data_validations_xml(sheet, Limits::default(), PART).unwrap();
        assert!(!unsupported);
        assert_eq!(validation_count, 7);

        let (groups, unsupported) =
            parse_sparkline_groups_xml(sheet, Limits::default(), &default_theme_colors(), PART)
                .unwrap();
        assert!(!unsupported);
        assert_eq!(groups.len(), 3);
        assert_eq!(groups[0].kind, SparklineKind::Line);
        assert_eq!(groups[1].kind, SparklineKind::Column);
        assert_eq!(groups[2].kind, SparklineKind::WinLoss);
        assert_eq!(groups[0].min_axis, SparklineAxisMode::Custom);
        assert_eq!(groups[0].manual_min, Some(0.0));
        assert_eq!(groups[0].manual_max, Some(30.0));
        assert_eq!(groups[0].series_color, 0x4472_c4ff);
        assert_eq!(groups[0].sparklines[0].target.canonical, "D1");
        assert!(parse_sparkline_source_range("'Data'!$A$1:$C$1", "Data", PART).is_some());
        assert!(parse_sparkline_source_range("SUM(A1:C1)", "Data", PART).is_none());
        assert!(parse_sparkline_source_range("[book.xlsx]Data!A1:C1", "Data", PART).is_none());
        assert!(parse_sparkline_source_range("Other!A1:C1", "Data", PART).is_none());
    }

    fn test_cell(address: &str, value_type: CellValueType) -> CellState {
        CellState {
            depth: 0,
            address: parse_a1(address).expect("test address is valid"),
            style_index: 0,
            value_type,
            value_depth: None,
            value_seen: false,
            value: String::new(),
            inline_string_depth: None,
            inline_text_depth: None,
            inline_string_seen: false,
            inline_text: String::new(),
            inline_run_depth: None,
            inline_properties_depth: None,
            inline_phonetic_depth: None,
            inline_run: None,
            inline_runs: Vec::new(),
            has_formula: false,
            shared_formula: None,
            formula_depth: None,
            formula: String::new(),
        }
    }

    fn resolve_test_cell_text(cell: &CellState, shared_strings: &[SharedString]) -> Option<String> {
        let mut materialized_text_bytes = 0;
        resolve_cell_text(
            cell,
            shared_strings,
            &CellStyle::default(),
            false,
            8,
            "sheet1.xml",
            &mut materialized_text_bytes,
            usize::MAX,
        )
        .unwrap()
    }

    #[test]
    fn formats_japanese_era_dates_from_tdf161301() {
        assert_eq!(
            super::format_excel_datetime(45440.0, "[$-ja-JP]0.00E+00", false),
            None
        );
        let bytes = include_bytes!("../../tests/fixtures/tdf161301.xlsx");
        let package = Package::open(bytes, Limits::default()).unwrap();
        let styles = parse_styles(&package, "xl/styles.xml", default_theme_colors()).unwrap();
        let format = styles.cells[3].number_format.as_deref().unwrap();
        for (serial, expected) in [
            ("45440", "令和6年5月28日"),
            ("43586", "令和元年5月1日"),
            ("43585", "平成31年4月30日"),
            ("32516", "平成元年1月8日"),
            ("32515", "昭和64年1月7日"),
            ("9856", "昭和元年12月25日"),
            ("9855", "大正15年12月24日"),
            ("4595", "大正元年7月30日"),
            ("4594", "明治45年7月29日"),
        ] {
            assert_eq!(
                format_number(serial, format, false, 8).as_deref(),
                Some(expected)
            );
        }
        for (format, expected) in [
            ("[$-411]ggge", "令和6"),
            ("[$-ja-JP]g ee", "R 06"),
            ("[$-30411]gg e", "令 6"),
            ("[$-ja-JP]yyyy", "2024"),
            (r#"yyyy\g"ggge""#, "2024gggge"),
        ] {
            assert_eq!(
                format_number("45440", format, false, 8).as_deref(),
                Some(expected)
            );
        }
        assert_eq!(
            format_number("43586", "[$-ja-JP]ggge", false, 8).as_deref(),
            Some("令和1")
        );
        assert_eq!(
            format_number("43978", format, true, 8).as_deref(),
            Some("令和6年5月28日")
        );
    }

    #[test]
    fn formats_excel_date_serials() {
        let format = builtin_number_format(14).expect("built-in short date format");
        assert_eq!(
            format_number("44927", format, false, 8),
            Some("2023/1/1".to_owned())
        );
        assert_eq!(
            format_number("45291", format, false, 8),
            Some("2023/12/31".to_owned())
        );
        assert_eq!(
            format_number("0", format, true, 8),
            Some("1904/1/1".to_owned())
        );

        let bytes = stored_zip(&[(
            "xl/workbook.xml",
            br#"<workbook xmlns:r="r"><workbookPr date1904="1"/><sheets><sheet name="Sheet1" r:id="r1"/></sheets></workbook>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        assert!(parse_workbook(&package, "xl/workbook.xml").unwrap().1);
    }

    #[test]
    fn formats_general_numbers_like_excel() {
        assert_eq!(
            format_number("83.116531372070313", "General", false, 8),
            Some("83.11653".to_owned())
        );
        assert_eq!(
            format_number("1.4164407253265381", "General", false, 8),
            Some("1.416441".to_owned())
        );
        assert_eq!(
            format_number("0.60000002384185791", "General", false, 8),
            Some("0.6".to_owned())
        );
        assert_eq!(
            format_number("83.116531372070313", "General", false, 11),
            Some("83.11653137".to_owned())
        );
    }

    #[test]
    fn preserves_excel_sheet_visibility_state() {
        let bytes = stored_zip(&[(
            "xl/workbook.xml",
            br#"<workbook xmlns:r="r"><sheets>
              <sheet name="Visible" r:id="r1"/>
              <sheet name="Hidden" state="hidden" r:id="r2"/>
              <sheet name="VeryHidden" state="veryHidden" r:id="r3"/>
            </sheets><definedNames>
              <definedName name="_xlnm.Print_Area" localSheetId="0">Visible!$A$1:$D$20</definedName>
            </definedNames></workbook>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let (sheets, _, _) = parse_workbook(&package, "xl/workbook.xml").unwrap();
        assert!(!sheets[0].hidden);
        assert_eq!(sheets[0].print_area.as_deref(), Some("Visible!$A$1:$D$20"));
        assert!(sheets[1].hidden);
        assert!(sheets[2].hidden);
    }

    #[test]
    fn real_chart_minor_ticks_are_rendered() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/ooxml-minor-ticks.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let document =
            super::parse(&package, "xl/workbook.xml", &ContentTypes::default(), None).unwrap();
        let ticks = document
            .objects
            .iter()
            .filter(|object| {
                (object.bounds.width - 3.0).abs() < 0.01 && object.bounds.height <= 0.02
            })
            .count();
        assert!(ticks >= 12, "expected minor ticks, got {ticks}");
    }

    #[test]
    fn real_workbook_print_metadata_survives_complete_parse() {
        for (bytes, titles) in [
            (
                include_bytes!("../../tests/fixtures/ooxml-print-titles.xlsx").as_slice(),
                true,
            ),
            (
                include_bytes!("../../tests/fixtures/ooxml-manual-breaks.xlsx").as_slice(),
                false,
            ),
        ] {
            let package = Package::open(bytes, Limits::default()).unwrap();
            let document =
                super::parse(&package, "xl/workbook.xml", &ContentTypes::default(), None).unwrap();
            let settings = document.units[0].sheet.as_ref().unwrap();
            if titles {
                assert_eq!(settings.print_titles.as_deref(), Some("DIET!$5:$5"));
            } else {
                assert_eq!(settings.row_breaks, vec![[8, 0, 16383]]);
            }
        }
        let mut parser = super::PrintSettingsParser::new(None);
        let mut diagnostics = Vec::new();
        crate::xml::parse_ooxml(br#"<worksheet><rowBreaks><brk id="9" min="bad" man="1"/><brk id="8" max="16383" man="1"/><brk id="7" max="16383"/></rowBreaks><colBreaks><brk id="2" min="3" max="9" man="1"/></colBreaks></worksheet>"#, Limits::default(), |event| parser.consume(&event, "sheet.xml", &mut diagnostics)).unwrap();
        assert_eq!(parser.finish().unwrap().row_breaks, vec![[8, 0, 16383]]);
        assert_eq!(diagnostics.len(), 2);
    }

    #[test]
    fn preserves_worksheet_print_setup_margins_and_headers() {
        let settings = parse_worksheet_print_settings(
            br#"<worksheet><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView view="pageLayout"/></sheetViews>
              <pageMargins left="0.7" right="0.7" top="0.8" bottom="0.8" header="0.3" footer="0.3"/>
              <pageSetup paperSize="9" orientation="landscape" scale="85" fitToWidth="1" fitToHeight="2"/>
              <headerFooter differentOddEven="1"><oddHeader>&amp;CCBAM</oddHeader><evenFooter>&amp;P</evenFooter></headerFooter>
            </worksheet>"#,
            Limits::default(),
            "xl/worksheets/sheet1.xml",
            Some("Sheet1!$A$1:$D$20".to_owned()),
        )
        .unwrap()
        .expect("print settings");
        assert_eq!(settings.view_mode, Some(SheetViewMode::PageLayout));
        assert_eq!(settings.paper_size, Some(9));
        assert_eq!(settings.orientation, Some(SheetOrientation::Landscape));
        assert_eq!(settings.scale, Some(85));
        assert_eq!(settings.fit_to_width, Some(1));
        assert_eq!(settings.fit_to_height, Some(2));
        assert!(settings.fit_to_page);
        assert!(settings.different_odd_even);
        assert_eq!(settings.margins.left, Some(0.7));
        assert_eq!(settings.odd_header.as_deref(), Some("&CCBAM"));
        assert_eq!(settings.even_footer.as_deref(), Some("&P"));
        assert_eq!(settings.print_area.as_deref(), Some("Sheet1!$A$1:$D$20"));
    }
}
