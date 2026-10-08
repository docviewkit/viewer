//! Native ODS spreadsheet parsing.

use super::odf_chart::{
    OdfChartDataLabels, OdfChartLegend, OdfChartStyleProps, OdfDataLabelNumber, OdfLabelPosition,
    parse_data_label_number, parse_label_position, parse_legend_position, parse_percent_fraction,
    pie_explosion_offset, pie_label_anchor, resolve_legend_bounds,
};
use super::optional_xml_attribute as optional_attribute;
use std::collections::{HashMap, HashSet};

use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
use crate::limits::Limits;
use crate::model::{
    Document, DocumentFormat, DocumentKind, FillRule, Geometry, ImageCrop, MappingQuality, Object,
    ObjectKind, Paint, PathCommand, Rect, SheetAxis, SheetAxisSpan, SourceLocator, SourceRef,
    StrokeStyle, TextAlign, TextLayout, TextRun, Unit, UnitKind, Visual,
};
use crate::package::Package;
use crate::xml::{XmlAttribute, XmlEvent, decode_xml_text, parse_xml};

use super::presentation_image::{
    OfficeImageError, office_image_media_type, reserve_materialized_image_bytes,
};
use super::{
    DEFAULT_SHEET_COLUMN_WIDTH as COLUMN_WIDTH, DEFAULT_SHEET_ROW_HEIGHT as ROW_HEIGHT,
    clone_materialized_text, local_name, reserve_materialized_text_bytes,
};

const CONTENT_PART: &str = "content.xml";
const MAX_COLUMNS: u32 = 16_384;
const MAX_ROWS: u32 = 1_048_576;

#[derive(Debug)]
struct TableState {
    tab_color: Option<u32>,
    depth: usize,
    source_index: u32,
    unit_index: u32,
    name: String,
    next_row: u32,
    next_source_row: u32,
    used_rows: u32,
    used_columns: u32,
    next_defined_column: u32,
    column_sizes: HashMap<u32, f32>,
    column_cell_styles: Vec<(u32, u32, String)>,
    row_sizes: Vec<SizeSpan>,
    placements: Vec<CellPlacement>,
    merges: Vec<CellRange>,
}

#[derive(Debug)]
struct RowState {
    depth: usize,
    source_index: u32,
    first_row: u32,
    repeat: u32,
    next_column: u32,
    next_source_cell: u32,
    height: f32,
    automatic_height: bool,
    optimal_height: bool,
    default_cell_style: Option<String>,
    cells: Vec<CellTemplate>,
}

#[derive(Debug)]
struct CellState {
    depth: usize,
    source_index: u32,
    first_column: u32,
    repeat: u32,
    column_span: u32,
    row_span: u32,
    style_name: Option<String>,
    element_id: Option<String>,
    value_type: Option<String>,
    value: Option<String>,
    string_value: Option<String>,
    date_value: Option<String>,
    boolean_value: Option<String>,
    paragraph_depth: Option<usize>,
    paragraph_count: u32,
    paragraph_text: String,
}

#[derive(Debug)]
struct CellTemplate {
    source_index: u32,
    first_column: u32,
    repeat: u32,
    column_span: u32,
    row_span: u32,
    style_name: Option<String>,
    element_id: Option<String>,
    text: Option<String>,
    numeric_value: Option<f64>,
    default_align: TextAlign,
}

#[derive(Clone, Copy, Debug)]
struct CellPlacement {
    object_index: usize,
    row: u32,
    column: u32,
    row_span: u32,
    column_span: u32,
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct CellRange {
    start_row: u32,
    end_row: u32,
    start_column: u32,
    end_column: u32,
}

#[derive(Clone, Copy, Debug)]
struct SizeSpan {
    start: u32,
    end: u32,
    size: f32,
}

#[derive(Debug)]
struct AxisLayout {
    default_size: f32,
    spans: Vec<SizeSpan>,
    prefix_delta: Vec<f32>,
}

impl AxisLayout {
    fn new(default_size: f32, mut spans: Vec<SizeSpan>) -> Self {
        spans.sort_unstable_by_key(|span| span.start);
        let mut prefix_delta = Vec::with_capacity(spans.len() + 1);
        prefix_delta.push(0.0);
        for span in &spans {
            let count = span.end - span.start + 1;
            let delta = count as f32 * (span.size - default_size);
            prefix_delta.push(prefix_delta.last().copied().unwrap_or(0.0) + delta);
        }
        Self {
            default_size,
            spans,
            prefix_delta,
        }
    }

    fn offset(&self, index: u32) -> f32 {
        let count = self.spans.partition_point(|span| span.start < index);
        let mut delta = self.prefix_delta[count];
        if count != 0 {
            let span = self.spans[count - 1];
            if index <= span.end {
                delta -= (span.end - index + 1) as f32 * (span.size - self.default_size);
            }
        }
        index as f32 * self.default_size + delta
    }

    fn span(&self, start: u32, end: u32) -> f32 {
        self.offset(end.saturating_add(1)) - self.offset(start)
    }

    fn descriptor(&self, count: u32) -> SheetAxis {
        let mut spans: Vec<SheetAxisSpan> = Vec::new();
        for span in self
            .spans
            .iter()
            .filter(|span| span.start < count && span.size.to_bits() != self.default_size.to_bits())
        {
            let end = span.end.min(count.saturating_sub(1));
            if let Some(last) = spans.last_mut()
                && last.end.saturating_add(1) == span.start
                && last.size.to_bits() == span.size.to_bits()
            {
                last.end = end;
            } else {
                spans.push(SheetAxisSpan {
                    start: span.start,
                    end,
                    size: span.size,
                });
            }
        }
        SheetAxis {
            default_size: self.default_size,
            spans,
        }
    }
}

#[derive(Clone, Debug)]
struct OdsCellStyle {
    data_style: Option<String>,
    parent_style: Option<String>,
    font_family: String,
    font_size: f32,
    color: u32,
    bold: bool,
    italic: bool,
    fill: u32,
    stroke: u32,
    stroke_width: f32,
    stroke_style: StrokeStyle,
    stroke_none: bool,
    align: Option<TextAlign>,
    data_label_number: Option<OdfDataLabelNumber>,
    data_label_text: Option<bool>,
    data_label_symbol: Option<bool>,
    label_position: Option<OdfLabelPosition>,
    pie_offset: Option<f32>,
    three_dimensional: bool,
    deep: bool,
    axis_minimum: Option<f32>,
    axis_maximum: Option<f32>,
    interval_major: Option<f32>,
    symbol_name: Option<String>,
    interpolation: Option<String>,
    regression_type: Option<String>,
    regression_name: Option<String>,
    extrapolate_forward: f32,
    extrapolate_backward: f32,
}

impl Default for OdsCellStyle {
    fn default() -> Self {
        Self {
            data_style: None,
            parent_style: None,
            font_family: "Arial".to_owned(),
            font_size: 11.0 * 96.0 / 72.0,
            color: 0x0000_00ff,
            bold: false,
            italic: false,
            fill: 0xffff_ffff,
            stroke: 0xd0d0_d0ff,
            stroke_width: 1.0,
            stroke_style: StrokeStyle::default(),
            stroke_none: false,
            align: None,
            data_label_number: None,
            data_label_text: None,
            data_label_symbol: None,
            label_position: None,
            pie_offset: None,
            three_dimensional: false,
            deep: false,
            axis_minimum: None,
            axis_maximum: None,
            interval_major: None,
            symbol_name: None,
            interpolation: None,
            regression_type: None,
            regression_name: None,
            extrapolate_forward: 0.0,
            extrapolate_backward: 0.0,
        }
    }
}

#[derive(Clone, Debug, Default)]
struct OdsStyles {
    tables: HashMap<String, (Option<String>, Option<String>)>,
    columns: HashMap<String, f32>,
    rows: HashMap<String, (f32, bool)>, // height, use-optimal-row-height
    cells: HashMap<String, OdsCellStyle>,
    numbers: HashMap<String, OdsNumberStyle>,
}

#[derive(Clone, Debug, Default)]
struct OdsNumberStyle {
    decimal: Option<OdsDecimalFormat>,
    fill: Option<char>,
    number_before_fill: bool,
    prefix: String,
    suffix: String,
    maps: Vec<(OdsCellOperator, f64, String)>,
}

// ODF optional decimal places/replacement are not equivalent to XLSX's current
// fixed-placeholder formatter; only digit grouping is shared.
#[derive(Clone, Debug)]
struct OdsDecimalFormat {
    places: Option<usize>,
    minimum: usize,
    integer_digits: usize,
    grouping: bool,
    replacement: String,
    prefix: String,
    suffix: String,
}

impl OdsDecimalFormat {
    fn render(&self, value: f64) -> String {
        let value = if value == 0.0 { 0.0 } else { value };
        let mut text = self
            .places
            .map_or_else(|| value.to_string(), |places| format!("{value:.places$}"));
        if let Some(dot) = text.find('.') {
            let end = text.trim_end_matches('0').len().max(dot + 1 + self.minimum);
            let removed = text.len() - end;
            text.truncate(end);
            if text.ends_with('.') && self.replacement.is_empty() {
                text.pop();
            }
            text.push_str(&self.replacement.repeat(removed));
        }
        let start = usize::from(text.starts_with('-'));
        let integer_end = text.find('.').unwrap_or(text.len());
        if self.integer_digits == 0 && &text[start..integer_end] == "0" {
            text.remove(start);
        } else {
            text.insert_str(
                start,
                &"0".repeat(self.integer_digits.saturating_sub(integer_end - start)),
            );
        }
        if self.grouping {
            text = super::group_decimal_digits(&text);
        }
        format!("{}{}{}", self.prefix, text, self.suffix)
    }
}

#[derive(Debug)]
struct ParsedContent {
    units: Vec<Unit>,
    objects: Vec<Object>,
    diagnostics: Vec<Diagnostic>,
}

#[derive(Clone, Copy, Debug, Default)]
struct FreezePanes {
    rows: u32,
    columns: u32,
}

#[derive(Clone, Copy, Debug)]
enum OdsCellOperator {
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
    GreaterThanOrEqual,
    GreaterThan,
}

#[derive(Clone, Copy, Debug)]
enum OdsConditionalValue {
    Minimum,
    Maximum,
    Number(f64),
    Percent(f64),
    Percentile(f64),
}

#[derive(Debug)]
enum OdsConditionalRule {
    CellIs {
        operator: OdsCellOperator,
        threshold: f64,
        style_name: String,
    },
    ColorScale {
        entries: Vec<(OdsConditionalValue, u32)>,
    },
    DataBar {
        thresholds: Vec<OdsConditionalValue>,
        positive_color: u32,
        negative_color: u32,
        max_length: f64,
    },
    IconSet {
        icon_type: String,
    },
}

#[derive(Debug)]
struct OdsConditionalFormat {
    table_name: String,
    range: CellRange,
    rules: Vec<OdsConditionalRule>,
}

#[derive(Debug)]
struct OdsConditionalCell {
    object_index: usize,
    unit_index: u32,
    table_name: String,
    row: u32,
    column: u32,
    value: f64,
}

#[derive(Debug)]
struct OdsConditionalStats {
    minimum: f64,
    maximum: f64,
    sorted: Vec<f64>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum OdsChartKind {
    Bar,
    Line,
    Pie,
    Scatter,
    Area,
    Radar,
}

fn data_labels_from_style(style: &OdsCellStyle) -> OdfChartDataLabels {
    let mut labels = OdfChartDataLabels {
        font_size: style.font_size,
        font_family: style.font_family.clone(),
        color: style.color,
        bold: style.bold,
        ..OdfChartDataLabels::default()
    };
    OdfChartStyleProps {
        data_label_number: style.data_label_number,
        data_label_text: style.data_label_text,
        data_label_symbol: style.data_label_symbol,
        label_position: style.label_position,
        pie_offset: style.pie_offset,
        solid_type_cuboid: false,
    }
    .apply_data_labels(&mut labels);
    labels
}

#[derive(Debug)]
struct OdsBasicChart {
    three_dimensional: bool,
    deep: bool,
    projection: Option<super::odf_chart::OdfChartProjection>,
    axis_fonts: [f32; 3],
    series_colors: Vec<u32>,
    floor_color: u32,
    plot_border: Option<(u32, f32)>,
    kind: OdsChartKind,
    series: Vec<Vec<f32>>,
    domains: Vec<Vec<f32>>,
    series_labels: Vec<String>,
    source_part: String,
    categories: Vec<String>,
    colors: Vec<u32>,
    point_explosions: Vec<f32>,
    plot_area: Option<Rect>,
    legend: Option<OdfChartLegend>,
    data_labels: Option<OdfChartDataLabels>,
    title: Option<String>,
    x_axis: OdsChartAxis,
    y_axis: OdsChartAxis,
    grid_axes: [bool; 2],
    regressions: Vec<OdsRegressionCurve>,
    series_style: OdsSeriesStyle,
}

#[derive(Clone, Copy, Debug)]
struct OdsChartAxis {
    minimum: f32,
    maximum: f32,
    interval: f32,
}

impl OdsChartAxis {
    // ODF ticks align to interval multiples; DrawingML starts at its minimum.
    // Index in f64 so adding a tiny interval cannot stall at a large f32 value.
    fn ticks(self) -> impl ExactSizeIterator<Item = f32> {
        let step = if self.interval.is_finite() && self.interval > 0.0 {
            self.interval as f64
        } else {
            ((self.maximum as f64 - self.minimum as f64) / 5.0).max(f64::EPSILON)
        };
        let start = (self.minimum as f64 / step).ceil() * step;
        let count = (((self.maximum as f64 - start) / step + 1e-6).floor() + 1.0).max(0.0) as usize;
        (0..count).map(move |i| (start + i as f64 * step) as f32)
    }
}

impl Default for OdsChartAxis {
    fn default() -> Self {
        Self {
            minimum: 0.0,
            maximum: 1.0,
            interval: 0.0,
        }
    }
}

#[derive(Clone, Debug)]
struct OdsRegressionCurve {
    kind: String,
    name: String,
    series_index: usize,
    forward: f32,
    backward: f32,
    color: u32,
    stroke_width: f32,
}

#[derive(Clone, Debug)]
struct OdsSeriesStyle {
    fill: u32,
    stroke: Option<u32>,
    stroke_width: f32,
    stroke_none: bool,
    symbol_name: Option<String>,
    interpolation: Option<String>,
}

impl Default for OdsSeriesStyle {
    fn default() -> Self {
        Self {
            fill: 0x0045_86ff,
            stroke: None,
            stroke_width: 2.0,
            stroke_none: false,
            symbol_name: None,
            interpolation: None,
        }
    }
}

#[derive(Clone, Debug)]
struct OdsChartCell {
    value: Option<f32>,
    text: String,
    paragraph_depth: Option<usize>,
    source_range: String,
    source_range_depth: Option<usize>,
}

#[derive(Debug)]
struct OdsChartRow {
    depth: usize,
    repeat: u32,
    cells: Vec<OdsChartCell>,
}

#[derive(Debug)]
struct OdsChartParseState {
    three_dimensional: bool,
    deep: bool,
    projection: Option<super::odf_chart::OdfChartProjection>,
    axis_fonts: [f32; 3],
    series_colors: Vec<u32>,
    floor_color: u32,
    plot_border: Option<(u32, f32)>,
    depth: usize,
    kind: Option<OdsChartKind>,
    series_ranges: Vec<String>,
    label_ranges: Vec<Option<String>>,
    domain_ranges: Vec<String>,
    rows: Vec<Vec<OdsChartCell>>,
    row: Option<OdsChartRow>,
    cell: Option<(usize, u32, OdsChartCell)>,
    categories_range: Option<String>,
    colors: Vec<u32>,
    point_explosions: Vec<f32>,
    plot_area: Option<Rect>,
    legend: Option<OdfChartLegend>,
    plot_data_labels: Option<OdfChartDataLabels>,
    series_data_labels: Option<OdfChartDataLabels>,
    title: Option<String>,
    x_axis: Option<OdsChartAxis>,
    y_axis: Option<OdsChartAxis>,
    grid_axes: [bool; 2],
    regressions: Vec<OdsRegressionCurve>,
    series_style: OdsSeriesStyle,
    title_depth: Option<usize>,
    current_axis: Option<usize>,
}

#[derive(Debug)]
struct OdsChartFrame {
    unit_index: u32,
    table_name: String,
    frame_index: u32,
    element_id: Option<String>,
    bounds: Rect,
    inline_chart_index: Option<usize>,
    object_href: Option<String>,
}

fn parse_styles_xml(
    bytes: &[u8],
    limits: Limits,
    part: &str,
    styles: &mut OdsStyles,
) -> Result<(), Diagnostic> {
    #[derive(Debug)]
    struct StyleState {
        depth: usize,
        name: String,
        family: String,
        column_width: Option<f32>,
        row_height: Option<f32>,
        optimal_row_height: bool,
        tab_color: Option<String>,
        cell: OdsCellStyle,
    }

    fn store_style(
        state: StyleState,
        styles: &mut OdsStyles,
        limits: Limits,
        part: &str,
    ) -> Result<(), Diagnostic> {
        let count =
            styles.columns.len() + styles.rows.len() + styles.cells.len() + styles.tables.len();
        if count >= limits.max_document_objects {
            return Err(object_limit_error(
                part,
                "style table exceeds the configured object limit",
            ));
        }
        match state.family.as_str() {
            "table" => {
                styles
                    .tables
                    .insert(state.name, (state.cell.parent_style, state.tab_color));
            }
            "table-column" => {
                if let Some(width) = state.column_width {
                    styles.columns.insert(state.name, width);
                }
            }
            "table-row" => {
                if state.row_height.is_some() || state.optimal_row_height {
                    styles.rows.insert(
                        state.name,
                        (
                            state.row_height.unwrap_or(ROW_HEIGHT),
                            state.optimal_row_height,
                        ),
                    );
                }
            }
            "table-cell" | "chart" => {
                styles.cells.insert(state.name, state.cell);
            }
            _ => {}
        }
        Ok(())
    }

    let mut depth = 0_usize;
    let mut current = None::<StyleState>;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "style" && current.is_none() {
                    let Some(style_name) = optional_attribute(&attributes, "name", part)? else {
                        if !empty {
                            depth = depth.saturating_add(1);
                        }
                        return Ok(());
                    };
                    let family = required_attribute(&attributes, "family", part)?;
                    let state = StyleState {
                        depth,
                        name: style_name,
                        family,
                        column_width: None,
                        row_height: None,
                        optimal_row_height: false,
                        tab_color: None,
                        cell: OdsCellStyle {
                            data_style: optional_attribute(&attributes, "data-style-name", part)?,
                            parent_style: optional_attribute(
                                &attributes,
                                "parent-style-name",
                                part,
                            )?,
                            ..OdsCellStyle::default()
                        },
                    };
                    if empty {
                        store_style(state, styles, limits, part)?;
                    } else {
                        current = Some(state);
                    }
                } else if let Some(style) = current.as_mut() {
                    match local {
                        "table-properties" => {
                            style.tab_color = optional_attribute(&attributes, "tab-color", part)?;
                        }
                        "table-column-properties" => {
                            if let Some(width) =
                                optional_attribute(&attributes, "column-width", part)?
                            {
                                style.column_width =
                                    Some(parse_length(&width, part, "column width")?);
                            }
                        }
                        "table-row-properties" => {
                            style.optimal_row_height =
                                optional_attribute(&attributes, "use-optimal-row-height", part)?
                                    .is_some_and(|value| matches!(value.as_str(), "true" | "1"));
                            if let Some(height) =
                                optional_attribute(&attributes, "row-height", part)?
                            {
                                style.row_height = Some(parse_length(&height, part, "row height")?);
                            }
                        }
                        "table-cell-properties" | "graphic-properties" => {
                            if let Some(fill) = optional_attribute(&attributes, "fill-color", part)?
                                && let Some(color) = parse_odf_color(&fill, part)?
                            {
                                style.cell.fill = color;
                            }
                            if let Some(background) =
                                optional_attribute(&attributes, "background-color", part)?
                                && let Some(color) = parse_odf_color(&background, part)?
                            {
                                style.cell.fill = color;
                            }
                            if optional_attribute(&attributes, "stroke", part)?.as_deref()
                                == Some("none")
                            {
                                style.cell.stroke_none = true;
                            }
                            if let Some(stroke) =
                                optional_attribute(&attributes, "stroke-color", part)?
                                && let Some(color) = parse_odf_color(&stroke, part)?
                            {
                                style.cell.stroke = color;
                            }
                            if let Some(width) =
                                optional_attribute(&attributes, "stroke-width", part)?
                            {
                                style.cell.stroke_width =
                                    parse_length(&width, part, "stroke width")?;
                            }
                            if let Some(border) = optional_attribute(&attributes, "border", part)? {
                                apply_border(&border, &mut style.cell, part)?;
                            }
                        }
                        "text-properties" => {
                            if let Some(family) =
                                optional_attribute(&attributes, "font-family", part)?
                            {
                                style.cell.font_family =
                                    super::odf_primary_font_family(&family).to_owned();
                            }
                            if let Some(size) = optional_attribute(&attributes, "font-size", part)?
                            {
                                style.cell.font_size = parse_length(&size, part, "font size")?;
                            }
                            if optional_attribute(&attributes, "font-weight", part)?.as_deref()
                                == Some("bold")
                            {
                                style.cell.bold = true;
                            }
                            if optional_attribute(&attributes, "font-style", part)?.as_deref()
                                == Some("italic")
                            {
                                style.cell.italic = true;
                            }
                            if let Some(color) = optional_attribute(&attributes, "color", part)?
                                && let Some(color) = parse_odf_color(&color, part)?
                            {
                                style.cell.color = color;
                            }
                        }
                        "paragraph-properties" => {
                            if let Some(align) =
                                optional_attribute(&attributes, "text-align", part)?
                            {
                                style.cell.align = Some(match align.as_str() {
                                    "center" | "justify" => TextAlign::Center,
                                    "end" | "right" => TextAlign::End,
                                    _ => TextAlign::Start,
                                });
                            }
                        }
                        "chart-properties" => {
                            style.cell.interpolation =
                                optional_attribute(&attributes, "interpolation", part)?;
                            if let Some(number) =
                                optional_attribute(&attributes, "data-label-number", part)?
                            {
                                style.cell.data_label_number = parse_data_label_number(&number);
                            }
                            if let Some(text) =
                                optional_attribute(&attributes, "data-label-text", part)?
                            {
                                style.cell.data_label_text =
                                    Some(matches!(text.as_str(), "true" | "1"));
                            }
                            if let Some(symbol) =
                                optional_attribute(&attributes, "data-label-symbol", part)?
                            {
                                style.cell.data_label_symbol =
                                    Some(matches!(symbol.as_str(), "true" | "1"));
                            }
                            if let Some(position) =
                                optional_attribute(&attributes, "label-position", part)?
                            {
                                style.cell.label_position = parse_label_position(&position);
                            }
                            if let Some(offset) =
                                optional_attribute(&attributes, "pie-offset", part)?
                            {
                                style.cell.pie_offset = parse_percent_fraction(&offset);
                            }
                            style.cell.three_dimensional =
                                optional_attribute(&attributes, "three-dimensional", part)?
                                    .as_deref()
                                    .is_some_and(|v| matches!(v, "true" | "1"));
                            style.cell.deep = optional_attribute(&attributes, "deep", part)?
                                .as_deref()
                                .is_some_and(|v| matches!(v, "true" | "1"));
                            if let Some(minimum) = optional_attribute(&attributes, "minimum", part)?
                            {
                                style.cell.axis_minimum = minimum.parse::<f32>().ok();
                            }
                            if let Some(maximum) = optional_attribute(&attributes, "maximum", part)?
                            {
                                style.cell.axis_maximum = maximum.parse::<f32>().ok();
                            }
                            if let Some(interval) =
                                optional_attribute(&attributes, "interval-major", part)?
                            {
                                style.cell.interval_major = interval.parse::<f32>().ok();
                            }
                            if let Some(name) =
                                optional_attribute(&attributes, "symbol-name", part)?
                            {
                                style.cell.symbol_name = Some(name);
                            }
                            if let Some(regression) =
                                optional_attribute(&attributes, "regression-type", part)?
                            {
                                style.cell.regression_type = Some(regression);
                            }
                            style.cell.regression_name =
                                optional_attribute(&attributes, "regression-name", part)?;
                            for (name, target) in [
                                (
                                    "regression-extrapolate-forward",
                                    &mut style.cell.extrapolate_forward,
                                ),
                                (
                                    "regression-extrapolate-backward",
                                    &mut style.cell.extrapolate_backward,
                                ),
                            ] {
                                *target = optional_attribute(&attributes, name, part)?
                                    .and_then(|v| v.parse::<f32>().ok())
                                    .filter(|v| v.is_finite() && *v >= 0.0)
                                    .unwrap_or(0.0);
                            }
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
                if local_name(name) == "style"
                    && current.as_ref().is_some_and(|style| style.depth == depth)
                {
                    store_style(
                        current.take().expect("checked style state"),
                        styles,
                        limits,
                        part,
                    )?;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    parse_number_styles(bytes, limits, part, styles)?;
    Ok(())
}

fn parse_number_styles(
    bytes: &[u8],
    limits: Limits,
    part: &str,
    styles: &mut OdsStyles,
) -> Result<(), Diagnostic> {
    let mut current = None::<(String, OdsNumberStyle)>;
    let mut capture = String::new();
    let mut literal = String::new();
    let mut decimal_supported = false;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let tag = local_name(name);
                if matches!(tag, "number-style" | "currency-style" | "percentage-style") {
                    capture.clear();
                    // Text language (fo:language) does not set number locale.
                    // Preserve cached output for explicit locales until supported.
                    decimal_supported = tag == "number-style"
                        && [
                            "language",
                            "country",
                            "rfc-language-tag",
                            "transliteration-format",
                        ]
                        .iter()
                        .all(|key| !attributes.iter().any(|a| local_name(a.name) == *key));
                    if empty {
                        current = None;
                        return Ok(());
                    }
                    current = Some((
                        required_attribute(&attributes, "name", part)?,
                        OdsNumberStyle::default(),
                    ));
                } else if let Some((_, style)) = current.as_mut() {
                    match tag {
                        "number" | "scientific-number" | "fraction" => {
                            if style.fill.is_none() {
                                style.number_before_fill = true;
                            }
                            if tag == "number" && decimal_supported {
                                let bounded =
                                    |name, default| -> Result<Option<usize>, Diagnostic> {
                                        Ok(optional_attribute(&attributes, name, part)?
                                            .map_or(Some(default), |v| {
                                                v.parse::<usize>().ok().filter(|v| *v <= 20)
                                            }))
                                    };
                                let places =
                                    optional_attribute(&attributes, "decimal-places", part)?;
                                let maximum = bounded("decimal-places", 0)?;
                                let minimum = bounded("min-decimal-places", maximum.unwrap_or(0))?;
                                let integer = bounded("min-integer-digits", 1)?;
                                let replacement =
                                    optional_attribute(&attributes, "decimal-replacement", part)?
                                        .unwrap_or_default();
                                if let (Some(maximum), Some(minimum), Some(integer)) =
                                    (maximum, minimum, integer)
                                    && minimum <= maximum
                                    && matches!(replacement.as_str(), "" | " ")
                                    && optional_attribute(&attributes, "display-factor", part)?
                                        .is_none()
                                {
                                    style.decimal = Some(OdsDecimalFormat {
                                        places: places.map(|_| maximum),
                                        minimum,
                                        integer_digits: integer,
                                        grouping: optional_attribute(
                                            &attributes,
                                            "grouping",
                                            part,
                                        )?
                                        .is_some_and(|v| matches!(v.as_str(), "true" | "1")),
                                        replacement,
                                        prefix: style.prefix.clone(),
                                        suffix: String::new(),
                                    });
                                }
                            } else {
                                style.decimal = None;
                            }
                            decimal_supported = false;
                        }
                        "embedded-text" => style.decimal = None,
                        "text" | "currency-symbol" | "fill-character" if !empty => {
                            capture = tag.to_owned();
                            literal.clear();
                        }
                        "map" => {
                            if let Some(condition) =
                                optional_attribute(&attributes, "condition", part)?
                                && let Some(suffix) = condition.strip_prefix("value()")
                                && let Some((operator, threshold)) =
                                    parse_ods_value_comparison(&format!("cell-content(){suffix}"))
                            {
                                style.maps.push((
                                    operator,
                                    threshold,
                                    required_attribute(&attributes, "apply-style-name", part)?,
                                ));
                            }
                        }
                        _ => {}
                    }
                }
            }
            XmlEvent::Text(text) if !capture.is_empty() => {
                literal.push_str(&decode_xml_text(text)?)
            }
            XmlEvent::Cdata(text) if !capture.is_empty() => literal.push_str(text),
            XmlEvent::EndElement { name } => {
                let tag = local_name(name);
                if tag == capture {
                    if let Some((_, style)) = current.as_mut() {
                        if tag != "fill-character" {
                            if let Some(decimal) = style.decimal.as_mut() {
                                decimal.suffix.push_str(&literal);
                            }
                        }
                        if tag == "fill-character" {
                            style.fill = literal
                                .chars()
                                .next()
                                .filter(|character| !character.is_control());
                        } else if style.fill.is_some() {
                            style.suffix.push_str(&literal);
                        } else {
                            style.prefix.push_str(&literal);
                        }
                    }
                    capture.clear();
                }
                if matches!(tag, "number-style" | "currency-style" | "percentage-style")
                    && let Some((name, style)) = current.take()
                {
                    if styles.numbers.len() >= limits.max_document_objects {
                        return Err(object_limit_error(
                            part,
                            "number styles exceed the configured object limit",
                        ));
                    }
                    styles.numbers.insert(name, style);
                }
            }
            _ => {}
        }
        Ok(())
    })
    .map(|_| ())
    .map_err(|error| with_part(error, part))
}

fn number_style_name<'a>(style: &'a OdsCellStyle, styles: &'a OdsStyles) -> Option<&'a str> {
    let mut cell_style = style;
    for _ in 0..32 {
        if cell_style.data_style.is_some() {
            return cell_style.data_style.as_deref();
        }
        cell_style = styles.cells.get(cell_style.parent_style.as_deref()?)?;
    }
    None
}

fn resolve_number_style<'a>(
    name: &str,
    styles: &'a OdsStyles,
    value: f64,
) -> Option<&'a OdsNumberStyle> {
    let mut number = styles.numbers.get(name)?;
    // Bounded traversal also isolates cyclic style maps and inheritance.
    for _ in 0..32 {
        let Some((_, _, name)) = number
            .maps
            .iter()
            .find(|(operator, threshold, _)| ods_value_matches(value, *operator, *threshold))
        else {
            return Some(number);
        };
        number = styles.numbers.get(name)?;
    }
    None
}

fn number_fill_character(
    style: &OdsCellStyle,
    styles: &OdsStyles,
    text: &str,
    value: f64,
) -> Option<(u32, char)> {
    let number = resolve_number_style(number_style_name(style, styles)?, styles, value)?;
    let character = number.fill?;
    // Cached text preserves locale-specific formatting; locate the authored literal boundary.
    let offset = if number.number_before_fill {
        text.strip_suffix(&number.suffix)?.encode_utf16().count()
    } else {
        text.strip_prefix(&number.prefix)?;
        number.prefix.encode_utf16().count()
    };
    Some((u32::try_from(offset).ok()?, character))
}

fn parse_length(value: &str, part: &str, label: &str) -> Result<f32, Diagnostic> {
    let (number, scale) = [
        ("in", 96.0),
        ("cm", 96.0 / 2.54),
        ("mm", 96.0 / 25.4),
        ("pt", 96.0 / 72.0),
        ("pc", 16.0),
        ("px", 1.0),
    ]
    .into_iter()
    .find_map(|(suffix, scale)| value.strip_suffix(suffix).map(|number| (number, scale)))
    .ok_or_else(|| format_error(part, format!("{label} uses an unsupported length unit")))?;
    let number = number
        .parse::<f32>()
        .ok()
        .filter(|number| number.is_finite() && *number >= 0.0)
        .ok_or_else(|| {
            format_error(part, format!("{label} is not a finite non-negative length"))
        })?;
    let pixels = number * scale;
    if !pixels.is_finite() || pixels > 1_000_000.0 {
        return Err(dimension_error(part, format!("{label} exceeds its limit")));
    }
    Ok(pixels)
}

fn parse_odf_color(value: &str, part: &str) -> Result<Option<u32>, Diagnostic> {
    if value == "transparent" {
        return Ok(None);
    }
    let digits = value
        .strip_prefix('#')
        .filter(|digits| digits.len() == 6 && digits.bytes().all(|byte| byte.is_ascii_hexdigit()))
        .ok_or_else(|| format_error(part, "ODF color must use #RRGGBB"))?;
    let rgb =
        u32::from_str_radix(digits, 16).map_err(|_| format_error(part, "ODF color is invalid"))?;
    Ok(Some((rgb << 8) | 0xff))
}

fn apply_border(value: &str, style: &mut OdsCellStyle, part: &str) -> Result<(), Diagnostic> {
    if value == "none" {
        style.stroke_width = 0.0;
        style.stroke_style = StrokeStyle::default();
        return Ok(());
    }
    let mut tokens = value.split_whitespace();
    let width = tokens
        .next()
        .ok_or_else(|| format_error(part, "ODF border is missing its width"))?;
    style.stroke_width = match width {
        "thin" => 1.0,
        "medium" => 3.0,
        "thick" => 5.0,
        _ => parse_length(width, part, "border width")?,
    }
    .max(0.5);
    style.stroke_style =
        super::odf_border_stroke_style(tokens.next().unwrap_or("solid"), style.stroke_width);
    if let Some(color) = value
        .split_whitespace()
        .find(|token| token.starts_with('#'))
    {
        style.stroke = parse_odf_color(color, part)?.unwrap_or(style.stroke);
    }
    Ok(())
}

fn parse_freeze_panes(
    bytes: &[u8],
    limits: Limits,
    part: &str,
) -> Result<HashMap<String, FreezePanes>, Diagnostic> {
    #[derive(Debug)]
    struct Entry {
        depth: usize,
        name: String,
        horizontal_mode: Option<u32>,
        horizontal_position: Option<u32>,
        vertical_mode: Option<u32>,
        vertical_position: Option<u32>,
    }

    let mut depth = 0_usize;
    let mut entry = None::<Entry>;
    let mut item = None::<(usize, String, String)>;
    let mut panes = HashMap::new();
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "config-item-map-entry" && entry.is_none() {
                    if let Some(name) = optional_attribute(&attributes, "name", part)? {
                        entry = Some(Entry {
                            depth,
                            name,
                            horizontal_mode: None,
                            horizontal_position: None,
                            vertical_mode: None,
                            vertical_position: None,
                        });
                    }
                } else if local == "config-item"
                    && entry.is_some()
                    && item.is_none()
                    && let Some(name) = optional_attribute(&attributes, "name", part)?
                    && matches!(
                        name.as_str(),
                        "HorizontalSplitMode"
                            | "HorizontalSplitPosition"
                            | "VerticalSplitMode"
                            | "VerticalSplitPosition"
                    )
                {
                    item = (!empty).then_some((depth, name, String::new()));
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if local == "config-item"
                    && item
                        .as_ref()
                        .is_some_and(|(item_depth, _, _)| *item_depth == depth)
                {
                    let (_, name, value) = item
                        .take()
                        .ok_or_else(|| format_error(part, "view-setting item state was lost"))?;
                    let value = value.trim().parse::<u32>().map_err(|_| {
                        format_error(part, format!("view setting {name} must be an integer"))
                    })?;
                    let entry = entry
                        .as_mut()
                        .ok_or_else(|| format_error(part, "view-setting entry state was lost"))?;
                    match name.as_str() {
                        "HorizontalSplitMode" => entry.horizontal_mode = Some(value),
                        "HorizontalSplitPosition" => entry.horizontal_position = Some(value),
                        "VerticalSplitMode" => entry.vertical_mode = Some(value),
                        "VerticalSplitPosition" => entry.vertical_position = Some(value),
                        _ => {}
                    }
                } else if local == "config-item-map-entry"
                    && entry.as_ref().is_some_and(|entry| entry.depth == depth)
                {
                    let entry = entry.take().ok_or_else(|| {
                        format_error(part, "view-setting map-entry state was lost")
                    })?;
                    let rows = if entry.horizontal_mode == Some(2) {
                        entry.horizontal_position.unwrap_or(0)
                    } else {
                        0
                    };
                    let columns = if entry.vertical_mode == Some(2) {
                        entry.vertical_position.unwrap_or(0)
                    } else {
                        0
                    };
                    if rows > MAX_ROWS || columns > MAX_COLUMNS {
                        return Err(dimension_error(part, "frozen panes exceed sheet limits"));
                    }
                    if (rows != 0 || columns != 0)
                        && panes
                            .insert(entry.name, FreezePanes { rows, columns })
                            .is_some()
                    {
                        return Err(format_error(part, "duplicate sheet view settings"));
                    }
                }
            }
            XmlEvent::Text(text) => {
                if let Some((_, _, value)) = item.as_mut() {
                    value.push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some((_, _, value)) = item.as_mut() {
                    value.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(panes)
}

pub fn parse(package: &Package<'_>) -> Result<Document, Diagnostic> {
    let bytes = package.required_part(CONTENT_PART)?;
    let mut styles = OdsStyles::default();
    if let Some(style_bytes) = package.part("styles.xml")? {
        parse_styles_xml(&style_bytes, package.limits(), "styles.xml", &mut styles)?;
    }
    parse_styles_xml(&bytes, package.limits(), CONTENT_PART, &mut styles)?;
    let frozen_panes = match package.part("settings.xml")? {
        Some(settings) => parse_freeze_panes(&settings, package.limits(), "settings.xml")?,
        None => HashMap::new(),
    };
    let mut content =
        parse_content_xml_with_styles(&bytes, package.limits(), &styles, &frozen_panes)?;
    if content.units.is_empty() {
        return Err(format_error(CONTENT_PART, "spreadsheet contains no sheets"));
    }
    push_ods_images(
        package,
        &bytes,
        &mut content.units,
        &mut content.objects,
        &mut content.diagnostics,
    )?;
    push_ods_charts(
        Some(package),
        &bytes,
        package.limits(),
        &mut content.units,
        &mut content.objects,
        &mut content.diagnostics,
    )?;

    Ok(into_document(content))
}

pub(super) fn parse_flat(bytes: &[u8], limits: Limits) -> Result<Document, Diagnostic> {
    let mut styles = OdsStyles::default();
    parse_styles_xml(bytes, limits, CONTENT_PART, &mut styles)?;
    let mut content = parse_content_xml_with_styles(bytes, limits, &styles, &HashMap::new())?;
    if content.units.is_empty() {
        return Err(format_error(CONTENT_PART, "spreadsheet contains no sheets"));
    }
    push_ods_charts(
        None,
        bytes,
        limits,
        &mut content.units,
        &mut content.objects,
        &mut content.diagnostics,
    )?;
    let mut document = into_document(content);
    document
        .diagnostics
        .extend(super::security::inspect_flat_odf(bytes, limits)?);
    Ok(document)
}

fn into_document(mut content: ParsedContent) -> Document {
    // Drawing frames can extend beyond populated cells. Keep their viewport and
    // resize axes in agreement, using authored sizes even in otherwise empty rows/columns.
    for unit in &mut content.units {
        for (count, extent, axis, limit) in [
            (
                &mut unit.columns,
                &mut unit.width,
                &mut unit.column_axis,
                MAX_COLUMNS,
            ),
            (
                &mut unit.rows,
                &mut unit.height,
                &mut unit.row_axis,
                MAX_ROWS,
            ),
        ] {
            if *extent > axis.offset(*count).max(1.0) + 0.001 {
                *count = (*count).max(crate::model::sheet_axis_count_for_extent(
                    *extent,
                    limit,
                    |index| axis.offset(index),
                ));
                *extent = axis.offset(*count).max(1.0);
            }
            axis.spans.retain(|span| span.start < *count);
            for span in &mut axis.spans {
                span.end = span.end.min(count.saturating_sub(1));
            }
        }
    }
    Document {
        fatal: false,
        format: Some(DocumentFormat::Ods),
        kind: Some(DocumentKind::Spreadsheet),
        units: content.units,
        outline: Vec::new(),
        objects: content.objects,
        embedded_fonts: Vec::new(),
        font_alternate_names: Vec::new(),
        diagnostics: content.diagnostics,
    }
}

fn push_ods_images(
    package: &Package<'_>,
    content: &[u8],
    units: &mut [Unit],
    objects: &mut Vec<Object>,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<(), Diagnostic> {
    #[derive(Debug)]
    struct Frame {
        depth: usize,
        unit_index: u32,
        table_name: String,
        element_id: Option<String>,
        bounds: Rect,
        href: Option<String>,
    }

    let mut depth = 0_usize;
    let mut table = None::<(usize, u32, String)>;
    let mut frame = None::<Frame>;
    let mut frames = Vec::new();
    let mut next_unit = 0_u32;
    parse_xml(content, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if name == "table:table" && table.is_none() {
                    let table_name = units
                        .get(next_unit as usize)
                        .ok_or_else(|| format_error(CONTENT_PART, "sheet index is missing"))?
                        .name
                        .clone();
                    table = (!empty).then_some((depth, next_unit, table_name));
                    next_unit = next_unit.saturating_add(1);
                } else if name == "draw:frame" && frame.is_none() {
                    if let Some((_, unit_index, table_name)) = &table {
                        let x = optional_attribute(&attributes, "x", CONTENT_PART)?
                            .as_deref()
                            .map(|value| parse_length(value, CONTENT_PART, "frame x"))
                            .transpose()?
                            .unwrap_or(0.0);
                        let y = optional_attribute(&attributes, "y", CONTENT_PART)?
                            .as_deref()
                            .map(|value| parse_length(value, CONTENT_PART, "frame y"))
                            .transpose()?
                            .unwrap_or(0.0);
                        let width = required_attribute(&attributes, "width", CONTENT_PART)
                            .and_then(|value| parse_length(&value, CONTENT_PART, "frame width"))?;
                        let height = required_attribute(&attributes, "height", CONTENT_PART)
                            .and_then(|value| parse_length(&value, CONTENT_PART, "frame height"))?;
                        frame = (!empty).then_some(Frame {
                            depth,
                            unit_index: *unit_index,
                            table_name: table_name.clone(),
                            element_id: optional_attribute(&attributes, "id", CONTENT_PART)?,
                            bounds: Rect {
                                x,
                                y,
                                width,
                                height,
                            },
                            href: None,
                        });
                    }
                } else if local == "image"
                    && let Some(frame) = frame.as_mut()
                {
                    frame.href = optional_attribute(&attributes, "href", CONTENT_PART)?;
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if name == "draw:frame" && frame.as_ref().is_some_and(|frame| frame.depth == depth)
                {
                    let frame = frame
                        .take()
                        .ok_or_else(|| format_error(CONTENT_PART, "image frame state was lost"))?;
                    if frame.href.is_some() {
                        frames.push(frame);
                    }
                } else if name == "table:table"
                    && table
                        .as_ref()
                        .is_some_and(|(table_depth, _, _)| *table_depth == depth)
                {
                    table = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, CONTENT_PART))?;

    let mut materialized_image_bytes = 0_usize;
    for (frame_index, frame) in frames.into_iter().enumerate() {
        let href = frame.href.as_deref().unwrap_or_default();
        let Some(target) = internal_ods_image_target(href) else {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::ExternalResourceBlocked,
                    Phase::Security,
                    Fidelity::Blocked,
                    "external or unsafe ODS image reference was blocked",
                )
                .in_part(CONTENT_PART),
            );
            continue;
        };
        let Some(bytes) = package.part(target)? else {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Container,
                    Fidelity::Omitted,
                    "ODS image alternative is missing from the package",
                )
                .in_part(target),
            );
            continue;
        };
        let media_type = match office_image_media_type(target, &bytes) {
            Ok(media_type) => media_type,
            Err(error) => {
                diagnostics.push(unsupported_ods_image_diagnostic(target, error));
                continue;
            }
        };
        reserve_materialized_image_bytes(
            &mut materialized_image_bytes,
            bytes.len(),
            package.limits().max_total_uncompressed_bytes,
            target,
        )?;
        if objects.len() >= package.limits().max_document_objects {
            return Err(object_limit_error(
                CONTENT_PART,
                "ODS images exceed the configured object limit",
            ));
        }
        let numeric_id = u32::try_from(objects.len()).map_err(|_| {
            object_limit_error(CONTENT_PART, "object count exceeds supported range")
        })?;
        objects.push(Object {
            numeric_id,
            parent_numeric_id: None,
            stable_id: format!("object:{numeric_id}"),
            parent_stable_id: None,
            kind: ObjectKind::Image,
            unit_index: frame.unit_index,
            bounds: frame.bounds,
            z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
            text: None,
            source: SourceRef {
                part: CONTENT_PART.to_owned(),
                mapping: MappingQuality::Exact,
                locator: SourceLocator::Ods {
                    kind: "element",
                    table_name: frame.table_name,
                    row: None,
                    column: None,
                    element_id: frame.element_id,
                    path: format!(
                        "/office:document-content/office:body/office:spreadsheet/table:table/draw:frame[{}]",
                        frame_index + 1
                    ),
                },
            },
            visual: Visual::Image {
                media_type: media_type.to_owned(),
                bytes: bytes.into_vec(),
                crop: ImageCrop::default(),
            },
        });
        if let Some(unit) = units.get_mut(frame.unit_index as usize) {
            unit.width = unit.width.max(frame.bounds.x + frame.bounds.width);
            unit.height = unit.height.max(frame.bounds.y + frame.bounds.height);
        }
    }
    Ok(())
}

fn internal_ods_image_target(href: &str) -> Option<&str> {
    let target = href.strip_prefix("./").unwrap_or(href);
    (!target.is_empty()
        && !target.starts_with('/')
        && !target.starts_with("//")
        && !target.contains(':')
        && target
            .split('/')
            .all(|component| !matches!(component, "" | "." | "..")))
    .then_some(target)
}

fn unsupported_ods_image_diagnostic(part: &str, error: OfficeImageError) -> Diagnostic {
    let message = match error {
        OfficeImageError::UnsupportedFormat => "ODS image format is unsupported",
        OfficeImageError::SignatureMismatch => "ODS image bytes do not match their extension",
        OfficeImageError::DisabledByOffice => "ODS image format is disabled by Microsoft Office",
    };
    Diagnostic::warning(
        DiagnosticCode::UnsupportedFeature,
        Phase::Parse,
        Fidelity::Omitted,
        message,
    )
    .in_part(part)
}

fn parse_ods_charts_in_xml(
    bytes: &[u8],
    limits: Limits,
    part: &str,
) -> Result<Vec<Option<OdsBasicChart>>, Diagnostic> {
    let mut styles = OdsStyles::default();
    parse_styles_xml(bytes, limits, part, &mut styles)?;
    let mut depth = 0_usize;
    let mut chart = None::<OdsChartParseState>;
    let mut parsed = Vec::new();
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                match local {
                    "chart"
                        if chart.is_none()
                            && optional_attribute(&attributes, "class", part)?.is_some() =>
                    {
                        let kind =
                            optional_attribute(&attributes, "class", part)?.and_then(|value| {
                                match local_name(&value) {
                                    "bar" | "column" => Some(OdsChartKind::Bar),
                                    "line" => Some(OdsChartKind::Line),
                                    "circle" | "pie" => Some(OdsChartKind::Pie),
                                    "scatter" => Some(OdsChartKind::Scatter),
                                    "area" => Some(OdsChartKind::Area),
                                    "radar" | "filled-radar" | "net" => Some(OdsChartKind::Radar),
                                    _ => None,
                                }
                            });
                        if empty {
                            parsed.push(None);
                        } else {
                            chart = Some(OdsChartParseState {
                                three_dimensional: false,
                                deep: false,
                                projection: None,
                                axis_fonts: [10.0; 3],
                                series_colors: Vec::new(),
                                floor_color: 0xcccc_ccff,
                                plot_border: None,
                                depth,
                                kind,
                                series_ranges: Vec::new(),
                                label_ranges: Vec::new(),
                                domain_ranges: Vec::new(),
                                rows: Vec::new(),
                                row: None,
                                cell: None,
                                categories_range: None,
                                colors: Vec::new(),
                                point_explosions: Vec::new(),
                                plot_area: None,
                                legend: None,
                                plot_data_labels: None,
                                series_data_labels: None,
                                title: None,
                                x_axis: None,
                                y_axis: None,
                                grid_axes: [false; 2],
                                current_axis: None,
                                regressions: Vec::new(),
                                series_style: OdsSeriesStyle::default(),
                                title_depth: None,
                            });
                        }
                    }
                    "title" if chart.is_some() => {
                        if empty {
                            chart.as_mut().unwrap().title_depth = None;
                        } else {
                            chart.as_mut().unwrap().title_depth = Some(depth);
                            chart.as_mut().unwrap().title = Some(String::new());
                        }
                    }
                    "grid" if chart.is_some() => {
                        let state = chart.as_mut().unwrap();
                        if optional_attribute(&attributes, "class", part)?.as_deref()
                            == Some("major")
                        {
                            if let Some(axis) = state.current_axis {
                                state.grid_axes[axis] = true;
                            }
                        }
                    }
                    "axis" if chart.is_some() => {
                        let dimension =
                            optional_attribute(&attributes, "dimension", part)?.unwrap_or_default();
                        let style = optional_attribute(&attributes, "style-name", part)?
                            .and_then(|name| styles.cells.get(&name))
                            .cloned()
                            .unwrap_or_default();
                        let axis = OdsChartAxis {
                            minimum: style.axis_minimum.unwrap_or(f32::NAN),
                            maximum: style.axis_maximum.unwrap_or(f32::NAN),
                            interval: style.interval_major.unwrap_or(0.0),
                        };
                        let state = chart.as_mut().unwrap();
                        state.current_axis = if empty {
                            None
                        } else {
                            ["x", "y"].iter().position(|v| *v == dimension)
                        };
                        if let Some(index) = ["x", "y", "z"].iter().position(|v| *v == dimension) {
                            state.axis_fonts[index] = style.font_size;
                        }
                        match dimension.as_str() {
                            "x" => state.x_axis = Some(axis),
                            "y" => state.y_axis = Some(axis),
                            _ => {}
                        }
                    }
                    "regression-curve" if chart.is_some() => {
                        let style = optional_attribute(&attributes, "style-name", part)?
                            .and_then(|name| styles.cells.get(&name))
                            .cloned()
                            .unwrap_or_default();
                        let kind = style
                            .regression_type
                            .clone()
                            .unwrap_or_else(|| "linear".to_owned());
                        let state = chart.as_mut().unwrap();
                        state.regressions.push(OdsRegressionCurve {
                            name: style.regression_name.unwrap_or_else(|| kind.clone()),
                            series_index: state.series_ranges.len().saturating_sub(1),
                            forward: style.extrapolate_forward,
                            backward: style.extrapolate_backward,
                            kind,
                            color: style.stroke,
                            stroke_width: style.stroke_width.max(1.0),
                        });
                    }
                    "domain" if chart.is_some() => {
                        if let Some(range) =
                            optional_attribute(&attributes, "cell-range-address", part)?
                        {
                            chart.as_mut().unwrap().domain_ranges.push(range);
                        }
                    }
                    "categories" if chart.is_some() => {
                        chart.as_mut().unwrap().categories_range =
                            optional_attribute(&attributes, "cell-range-address", part)?;
                    }
                    "wall" if chart.is_some() => {
                        if let Some(style) = optional_attribute(&attributes, "style-name", part)?
                            .and_then(|name| styles.cells.get(&name))
                        {
                            chart.as_mut().unwrap().plot_border =
                                (!style.stroke_none).then_some((style.stroke, style.stroke_width));
                        }
                    }
                    "floor" if chart.is_some() => {
                        if let Some(style) = optional_attribute(&attributes, "style-name", part)?
                            .and_then(|name| styles.cells.get(&name))
                        {
                            chart.as_mut().unwrap().floor_color = style.fill;
                        }
                    }
                    "plot-area" | "coordinate-region" if chart.is_some() => {
                        if local == "plot-area" {
                            let state = chart.as_mut().unwrap();
                            if optional_attribute(&attributes, "projection", part)?.as_deref()
                                == Some("parallel")
                            {
                                state.projection = super::odf_chart::OdfChartProjection::parallel(
                                    &optional_attribute(&attributes, "transform", part)?
                                        .unwrap_or_default(),
                                    &optional_attribute(&attributes, "vpn", part)?
                                        .unwrap_or_default(),
                                    &optional_attribute(&attributes, "vup", part)?
                                        .unwrap_or_default(),
                                );
                            }
                        }
                        if let Some(bounds) = ods_chart_rect(&attributes, part)? {
                            chart.as_mut().unwrap().plot_area = Some(bounds);
                        }
                        if let Some(style) = optional_attribute(&attributes, "style-name", part)?
                            .and_then(|name| styles.cells.get(&name))
                        {
                            chart.as_mut().unwrap().three_dimensional = style.three_dimensional;
                            chart.as_mut().unwrap().deep = style.deep;
                            chart
                                .as_mut()
                                .unwrap()
                                .series_style
                                .interpolation
                                .clone_from(&style.interpolation);
                            chart.as_mut().unwrap().plot_data_labels =
                                Some(data_labels_from_style(style));
                        }
                    }
                    "legend" if chart.is_some() => {
                        let style = optional_attribute(&attributes, "style-name", part)?
                            .and_then(|name| styles.cells.get(&name))
                            .cloned()
                            .unwrap_or_default();
                        let position = optional_attribute(&attributes, "legend-position", part)?
                            .as_deref()
                            .and_then(parse_legend_position);
                        let expansion_wide =
                            optional_attribute(&attributes, "legend-expansion", part)?.as_deref()
                                == Some("wide");
                        chart.as_mut().unwrap().legend = Some(OdfChartLegend {
                            x: optional_attribute(&attributes, "x", part)?
                                .and_then(|value| parse_length(&value, part, "legend x").ok()),
                            y: optional_attribute(&attributes, "y", part)?
                                .and_then(|value| parse_length(&value, part, "legend y").ok()),
                            width: optional_attribute(&attributes, "width", part)?
                                .and_then(|value| parse_length(&value, part, "legend width").ok()),
                            height: optional_attribute(&attributes, "height", part)?
                                .and_then(|value| parse_length(&value, part, "legend height").ok()),
                            position,
                            expansion_wide,
                            font_size: style.font_size,
                            font_family: style.font_family,
                            color: style.color,
                        });
                    }
                    "data-point" if chart.is_some() => {
                        let chart = chart.as_mut().unwrap();
                        let repeat =
                            repeat_attribute(&attributes, "repeated", MAX_ROWS, part)? as usize;
                        if chart.colors.len().saturating_add(repeat) > limits.max_document_objects {
                            return Err(object_limit_error(
                                part,
                                "chart data points exceed the configured object limit",
                            ));
                        }
                        let style = optional_attribute(&attributes, "style-name", part)?
                            .and_then(|name| styles.cells.get(&name));
                        let color = style.map(|style| style.fill);
                        let explosion = style
                            .and_then(|style| style.pie_offset)
                            .unwrap_or(0.0)
                            .clamp(0.0, 4.0);
                        for _ in 0..repeat {
                            chart.colors.push(color.unwrap_or_else(|| {
                                super::office_chart_palette_color(chart.colors.len())
                            }));
                            chart.point_explosions.push(explosion);
                        }
                    }
                    "series" if chart.is_some() => {
                        if let Some(range) =
                            optional_attribute(&attributes, "values-cell-range-address", part)?
                            && let Some(chart) = chart.as_mut()
                        {
                            chart.series_ranges.push(range);
                            chart.series_colors.push(
                                optional_attribute(&attributes, "style-name", part)?
                                    .and_then(|name| styles.cells.get(&name))
                                    .map_or_else(
                                        || {
                                            super::office_chart_palette_color(
                                                chart.series_colors.len(),
                                            )
                                        },
                                        |style| style.fill,
                                    ),
                            );
                            chart.label_ranges.push(optional_attribute(
                                &attributes,
                                "label-cell-address",
                                part,
                            )?);
                        }
                        if let Some(style) = optional_attribute(&attributes, "style-name", part)?
                            .and_then(|name| styles.cells.get(&name))
                            .cloned()
                        {
                            let chart = chart.as_mut().expect("chart state is present");
                            chart.series_style = OdsSeriesStyle {
                                fill: style.fill,
                                stroke: Some(style.stroke),
                                stroke_width: style.stroke_width.max(1.0),
                                stroke_none: style.stroke_none,
                                symbol_name: style.symbol_name.clone(),
                                interpolation: style
                                    .interpolation
                                    .clone()
                                    .or_else(|| chart.series_style.interpolation.clone()),
                            };
                            let labels = chart
                                .series_data_labels
                                .get_or_insert_with(|| data_labels_from_style(&style));
                            labels.merge(&data_labels_from_style(&style));
                        }
                    }
                    "table-row" if chart.is_some() => {
                        let chart = chart.as_mut().expect("chart state is present");
                        if chart.row.is_some() {
                            return Err(format_error(part, "nested chart data rows are invalid"));
                        }
                        let repeat =
                            repeat_attribute(&attributes, "number-rows-repeated", MAX_ROWS, part)?;
                        chart.row = Some(OdsChartRow {
                            depth,
                            repeat,
                            cells: Vec::new(),
                        });
                    }
                    "table-cell" | "covered-table-cell"
                        if chart.as_ref().is_some_and(|chart| chart.row.is_some()) =>
                    {
                        let chart = chart.as_mut().expect("chart state is present");
                        if chart.cell.is_some() {
                            return Err(format_error(part, "nested chart data cells are invalid"));
                        }
                        let cell = OdsChartCell {
                            value: optional_attribute(&attributes, "value", part)?
                                .and_then(|value| value.parse::<f32>().ok())
                                .filter(|value| value.is_finite()),
                            text: optional_attribute(&attributes, "string-value", part)?
                                .unwrap_or_default(),
                            paragraph_depth: None,
                            source_range: String::new(),
                            source_range_depth: None,
                        };
                        let repeat = repeat_attribute(
                            &attributes,
                            "number-columns-repeated",
                            MAX_COLUMNS,
                            part,
                        )?;
                        if empty {
                            extend_ods_chart_cells(
                                &mut chart.row.as_mut().expect("chart row is present").cells,
                                cell,
                                repeat,
                                limits,
                                part,
                            )?;
                        } else {
                            chart.cell = Some((depth, repeat, cell));
                        }
                    }
                    "desc" => {
                        if let Some((_, _, cell)) =
                            chart.as_mut().and_then(|chart| chart.cell.as_mut())
                        {
                            cell.source_range_depth = (!empty).then_some(depth);
                        }
                    }
                    "p" => {
                        if let Some((_, _, cell)) =
                            chart.as_mut().and_then(|chart| chart.cell.as_mut())
                        {
                            cell.paragraph_depth = (!empty).then_some(depth);
                        }
                    }
                    _ => {}
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if local == "axis" {
                    if let Some(state) = chart.as_mut() {
                        state.current_axis = None;
                    }
                }
                if local == "title"
                    && chart
                        .as_ref()
                        .is_some_and(|chart| chart.title_depth == Some(depth))
                {
                    let chart = chart.as_mut().expect("chart state is present");
                    chart.title_depth = None;
                    if let Some(title) = chart.title.as_mut() {
                        *title = title.split_whitespace().collect::<Vec<_>>().join(" ");
                    }
                    return Ok(());
                }
                if local == "desc" {
                    if let Some((_, _, cell)) = chart.as_mut().and_then(|chart| chart.cell.as_mut())
                        && cell.source_range_depth == Some(depth)
                    {
                        cell.source_range_depth = None;
                    }
                }
                if local == "p" {
                    if let Some((_, _, cell)) = chart.as_mut().and_then(|chart| chart.cell.as_mut())
                        && cell.paragraph_depth == Some(depth)
                    {
                        cell.paragraph_depth = None;
                    }
                } else if matches!(local, "table-cell" | "covered-table-cell")
                    && chart
                        .as_ref()
                        .and_then(|chart| chart.cell.as_ref())
                        .is_some_and(|(start, _, _)| *start == depth)
                {
                    let chart = chart.as_mut().expect("chart state is present");
                    let (_, repeat, cell) = chart
                        .cell
                        .take()
                        .ok_or_else(|| format_error(part, "chart data cell state was lost"))?;
                    extend_ods_chart_cells(
                        &mut chart.row.as_mut().expect("chart row is present").cells,
                        cell,
                        repeat,
                        limits,
                        part,
                    )?;
                } else if local == "table-row"
                    && chart
                        .as_ref()
                        .and_then(|chart| chart.row.as_ref())
                        .is_some_and(|row| row.depth == depth)
                {
                    let chart = chart.as_mut().expect("chart state is present");
                    let row = chart
                        .row
                        .take()
                        .ok_or_else(|| format_error(part, "chart row state was lost"))?;
                    let repeat = usize::try_from(row.repeat)
                        .map_err(|_| object_limit_error(part, "chart row repeat is too large"))?;
                    if chart
                        .rows
                        .len()
                        .checked_add(repeat)
                        .is_none_or(|count| count > limits.max_document_objects)
                    {
                        return Err(object_limit_error(
                            part,
                            "chart rows exceed the configured object limit",
                        ));
                    }
                    chart.rows.extend(std::iter::repeat_n(row.cells, repeat));
                } else if local == "chart"
                    && chart.as_ref().is_some_and(|chart| chart.depth == depth)
                {
                    let chart = chart
                        .take()
                        .ok_or_else(|| format_error(part, "chart state was lost"))?;
                    parsed.push(finish_ods_chart_state(chart, part));
                }
            }
            XmlEvent::Text(text) => {
                if let Some((_, _, cell)) = chart.as_mut().and_then(|chart| chart.cell.as_mut())
                    && cell.source_range_depth.is_some()
                {
                    let decoded = decode_xml_text(text).map_err(|error| with_part(error, part))?;
                    append_ods_chart_text(&mut cell.source_range, &decoded, limits, part)?;
                    return Ok(());
                }
                if let Some(chart) = chart.as_mut()
                    && chart.title_depth.is_some()
                    && chart.cell.is_none()
                {
                    if let Some(title) = chart.title.as_mut() {
                        title.push_str(
                            &decode_xml_text(text).map_err(|error| with_part(error, part))?,
                        );
                    }
                    return Ok(());
                }
                if let Some((_, _, cell)) = chart
                    .as_mut()
                    .and_then(|chart| chart.cell.as_mut())
                    .filter(|(_, _, cell)| cell.paragraph_depth.is_some())
                {
                    let decoded = decode_xml_text(text).map_err(|error| with_part(error, part))?;
                    append_ods_chart_text(&mut cell.text, &decoded, limits, part)?;
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some((_, _, cell)) = chart.as_mut().and_then(|chart| chart.cell.as_mut())
                    && cell.source_range_depth.is_some()
                {
                    append_ods_chart_text(&mut cell.source_range, text, limits, part)?;
                    return Ok(());
                }
                if let Some(chart) = chart.as_mut()
                    && chart.title_depth.is_some()
                    && chart.cell.is_none()
                {
                    if let Some(title) = chart.title.as_mut() {
                        title.push_str(text);
                    }
                    return Ok(());
                }
                if let Some((_, _, cell)) = chart
                    .as_mut()
                    .and_then(|chart| chart.cell.as_mut())
                    .filter(|(_, _, cell)| cell.paragraph_depth.is_some())
                {
                    append_ods_chart_text(&mut cell.text, text, limits, part)?;
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(parsed)
}

fn ods_chart_rect(attributes: &[XmlAttribute<'_>], part: &str) -> Result<Option<Rect>, Diagnostic> {
    let mut values = Vec::with_capacity(4);
    for name in ["x", "y", "width", "height"] {
        let Some(value) = optional_attribute(attributes, name, part)? else {
            return Ok(None);
        };
        values.push(parse_length(&value, part, name)?);
    }
    Ok((values[2] > 0.0 && values[3] > 0.0).then_some(Rect {
        x: values[0],
        y: values[1],
        width: values[2],
        height: values[3],
    }))
}

fn extend_ods_chart_cells(
    cells: &mut Vec<OdsChartCell>,
    cell: OdsChartCell,
    repeat: u32,
    limits: Limits,
    part: &str,
) -> Result<(), Diagnostic> {
    let repeat = usize::try_from(repeat)
        .map_err(|_| object_limit_error(part, "chart cell repeat is too large"))?;
    if cells
        .len()
        .checked_add(repeat)
        .is_none_or(|count| count > limits.max_document_objects)
    {
        return Err(object_limit_error(
            part,
            "chart cells exceed the configured object limit",
        ));
    }
    cells.extend(std::iter::repeat_n(cell, repeat));
    Ok(())
}

fn append_ods_chart_text(
    target: &mut String,
    text: &str,
    limits: Limits,
    part: &str,
) -> Result<(), Diagnostic> {
    if target
        .len()
        .checked_add(text.len())
        .is_none_or(|length| length > limits.max_xml_bytes)
    {
        return Err(Diagnostic::fatal(
            DiagnosticCode::XmlSizeLimit,
            Phase::Xml,
            None,
            "ODS chart text exceeds the configured XML byte limit",
        )
        .in_part(part));
    }
    target.push_str(text);
    Ok(())
}

// Embedded Calc chart caches can insert label columns. Their svg:desc ranges
// identify the original worksheet range; worksheet coordinates are not cache coordinates.
fn ods_chart_cells<'a>(rows: &'a [Vec<OdsChartCell>], address: &str) -> Vec<&'a OdsChartCell> {
    let Some((table, mut range)) = parse_ods_cell_range(address) else {
        return Vec::new();
    };
    let anchor = rows.iter().enumerate().find_map(|(row_index, cells)| {
        cells.iter().enumerate().find_map(|(column_index, cell)| {
            let (source_table, source) = parse_ods_cell_range(&cell.source_range)?;
            (source_table == table && source == range).then_some((row_index, column_index))
        })
    });
    if let Some((row, column)) = anchor {
        range.end_row = row as u32 + range.end_row - range.start_row;
        range.end_column = column as u32 + range.end_column - range.start_column;
        range.start_row = row as u32;
        range.start_column = column as u32;
    }
    rows.iter()
        .skip(range.start_row as usize)
        .take((range.end_row - range.start_row + 1) as usize)
        .flat_map(|row| {
            row.iter()
                .skip(range.start_column as usize)
                .take((range.end_column - range.start_column + 1) as usize)
        })
        .collect()
}

fn finish_ods_chart_state(state: OdsChartParseState, part: &str) -> Option<OdsBasicChart> {
    let kind = state.kind?;
    let values = |address: &str| {
        ods_chart_cells(&state.rows, address)
            .into_iter()
            .filter_map(|cell| cell.value.or_else(|| cell.text.trim().parse::<f32>().ok()))
            .filter(|value| value.is_finite())
            .collect::<Vec<_>>()
    };
    let series = state
        .series_ranges
        .iter()
        .map(|address| values(address))
        .collect::<Vec<_>>();
    let domains = state
        .domain_ranges
        .iter()
        .map(|address| values(address))
        .collect::<Vec<_>>();
    let series_labels = state
        .label_ranges
        .iter()
        .map(|address| {
            address
                .as_deref()
                .and_then(|address| ods_chart_cells(&state.rows, address).first().copied())
                .map(|cell| cell.text.clone())
                .unwrap_or_default()
        })
        .collect();
    let categories = state
        .categories_range
        .as_deref()
        .map(|address| {
            ods_chart_cells(&state.rows, address)
                .iter()
                .map(|cell| cell.text.clone())
                .collect()
        })
        .unwrap_or_default();
    let y_values = series.iter().flatten().copied().collect::<Vec<_>>();
    let x_values = domains.iter().flatten().copied().collect::<Vec<_>>();
    let fallback_x = OdsChartAxis {
        minimum: x_values.iter().copied().fold(f32::INFINITY, f32::min),
        maximum: x_values.iter().copied().fold(f32::NEG_INFINITY, f32::max),
        interval: 0.0,
    };
    let fallback_y = OdsChartAxis {
        minimum: y_values.iter().copied().fold(
            if state.three_dimensional {
                0.0
            } else {
                f32::INFINITY
            },
            f32::min,
        ),
        maximum: y_values.iter().copied().fold(f32::NEG_INFINITY, f32::max),
        interval: 0.0,
    };
    let resolve_axis = |axis: Option<OdsChartAxis>, fallback: OdsChartAxis, horizontal: bool| {
        if kind == OdsChartKind::Scatter && !state.three_dimensional {
            let authored = axis.unwrap_or(OdsChartAxis {
                minimum: f32::NAN,
                maximum: f32::NAN,
                interval: 0.0,
            });
            let (minimum, maximum, interval) = super::odf_chart::scatter_axis_range(
                (fallback.minimum, fallback.maximum),
                (
                    Some(authored.minimum),
                    Some(authored.maximum),
                    Some(authored.interval),
                ),
                horizontal,
            );
            return OdsChartAxis {
                minimum,
                maximum,
                interval,
            };
        }
        let mut axis = axis.unwrap_or(fallback);
        if !axis.minimum.is_finite() {
            axis.minimum = fallback.minimum;
        }
        if !axis.maximum.is_finite() {
            axis.maximum = fallback.maximum;
        }
        if !axis.minimum.is_finite() {
            axis.minimum = 0.0;
        }
        if !axis.maximum.is_finite() || axis.maximum <= axis.minimum {
            axis.maximum = axis.minimum + 1.0;
        }
        axis
    };
    series
        .iter()
        .any(|values| !values.is_empty())
        .then_some(OdsBasicChart {
            three_dimensional: state.three_dimensional,
            deep: state.deep,
            projection: state.projection,
            axis_fonts: state.axis_fonts,
            series_colors: state.series_colors,
            floor_color: state.floor_color,
            plot_border: state.plot_border,
            kind,
            series,
            domains,
            series_labels,
            source_part: part.to_owned(),
            categories,
            colors: state.colors,
            point_explosions: state.point_explosions,
            plot_area: state.plot_area,
            legend: state.legend,
            data_labels: state.series_data_labels.or(state.plot_data_labels),
            title: state.title,
            x_axis: resolve_axis(state.x_axis, fallback_x, true),
            y_axis: resolve_axis(state.y_axis, fallback_y, false),
            grid_axes: state.grid_axes,
            regressions: state.regressions,
            series_style: state.series_style,
        })
}

fn parse_ods_chart_frames(
    bytes: &[u8],
    limits: Limits,
    units: &[Unit],
) -> Result<Vec<OdsChartFrame>, Diagnostic> {
    #[derive(Debug)]
    struct FrameState {
        depth: usize,
        unit_index: u32,
        table_name: String,
        frame_index: u32,
        element_id: Option<String>,
        bounds: Rect,
        inline_chart_index: Option<usize>,
        object_href: Option<String>,
    }

    let mut depth = 0_usize;
    let mut next_unit = 0_u32;
    let mut table = None::<(usize, u32, String, u32)>;
    let mut frame = None::<FrameState>;
    let mut chart_depth = None;
    let mut chart_index = 0_usize;
    let mut frames = Vec::new();
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                if chart_depth.is_some() {
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }
                let local = local_name(name);
                if local == "table" && table.is_none() && frame.is_none() {
                    let table_name = units
                        .get(next_unit as usize)
                        .ok_or_else(|| format_error(CONTENT_PART, "sheet index is missing"))?
                        .name
                        .clone();
                    table = (!empty).then_some((depth, next_unit, table_name, 0));
                    next_unit = next_unit.saturating_add(1);
                } else if local == "frame" && frame.is_none() {
                    if let Some((_, unit_index, table_name, next_frame)) = table.as_mut() {
                        let frame_index = *next_frame;
                        *next_frame = next_frame.saturating_add(1);
                        let bounds = Rect {
                            x: optional_attribute(&attributes, "x", CONTENT_PART)?
                                .as_deref()
                                .map(|value| parse_length(value, CONTENT_PART, "chart frame x"))
                                .transpose()?
                                .unwrap_or(0.0),
                            y: optional_attribute(&attributes, "y", CONTENT_PART)?
                                .as_deref()
                                .map(|value| parse_length(value, CONTENT_PART, "chart frame y"))
                                .transpose()?
                                .unwrap_or(0.0),
                            width: required_attribute(&attributes, "width", CONTENT_PART)
                                .and_then(|value| {
                                    parse_length(&value, CONTENT_PART, "chart frame width")
                                })?,
                            height: required_attribute(&attributes, "height", CONTENT_PART)
                                .and_then(|value| {
                                    parse_length(&value, CONTENT_PART, "chart frame height")
                                })?,
                        };
                        if !bounds.is_valid() || bounds.width == 0.0 || bounds.height == 0.0 {
                            return Err(format_error(
                                CONTENT_PART,
                                "ODS chart frame has invalid bounds",
                            ));
                        }
                        frame = (!empty).then_some(FrameState {
                            depth,
                            unit_index: *unit_index,
                            table_name: table_name.clone(),
                            frame_index,
                            element_id: optional_attribute(&attributes, "id", CONTENT_PART)?,
                            bounds,
                            inline_chart_index: None,
                            object_href: None,
                        });
                    }
                } else if local == "chart"
                    && frame.is_some()
                    && optional_attribute(&attributes, "class", CONTENT_PART)?.is_some()
                {
                    if let Some(frame) = frame.as_mut() {
                        frame.inline_chart_index = Some(chart_index);
                    }
                    chart_index = chart_index.checked_add(1).ok_or_else(|| {
                        object_limit_error(CONTENT_PART, "ODS chart count overflow")
                    })?;
                    chart_depth = (!empty).then_some(depth);
                } else if local == "object"
                    && let Some(frame) = frame.as_mut()
                {
                    frame.object_href = optional_attribute(&attributes, "href", CONTENT_PART)?;
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if let Some(start) = chart_depth {
                    if local_name(name) == "chart" && start == depth {
                        chart_depth = None;
                    }
                    return Ok(());
                }
                let local = local_name(name);
                if local == "frame" && frame.as_ref().is_some_and(|frame| frame.depth == depth) {
                    let frame = frame
                        .take()
                        .ok_or_else(|| format_error(CONTENT_PART, "chart frame state was lost"))?;
                    if frame.inline_chart_index.is_some() || frame.object_href.is_some() {
                        frames.push(OdsChartFrame {
                            unit_index: frame.unit_index,
                            table_name: frame.table_name,
                            frame_index: frame.frame_index,
                            element_id: frame.element_id,
                            bounds: frame.bounds,
                            inline_chart_index: frame.inline_chart_index,
                            object_href: frame.object_href,
                        });
                        if frames.len() > limits.max_relationship_edges {
                            return Err(Diagnostic::fatal(
                                DiagnosticCode::RelationshipLimit,
                                Phase::Parse,
                                None,
                                "ODS chart references exceed the configured relationship limit",
                            )
                            .in_part(CONTENT_PART));
                        }
                    }
                } else if local == "table"
                    && table
                        .as_ref()
                        .is_some_and(|(table_depth, _, _, _)| *table_depth == depth)
                {
                    table = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, CONTENT_PART))?;
    Ok(frames)
}

fn push_ods_charts(
    package: Option<&Package<'_>>,
    content: &[u8],
    limits: Limits,
    units: &mut [Unit],
    objects: &mut Vec<Object>,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<(), Diagnostic> {
    let mut inline_charts = parse_ods_charts_in_xml(content, limits, CONTENT_PART)?;
    let frames = parse_ods_chart_frames(content, limits, units)?;
    for frame in frames {
        let mut chart = frame
            .inline_chart_index
            .and_then(|index| inline_charts.get_mut(index))
            .and_then(Option::take);
        let mut diagnostic_part = CONTENT_PART.to_owned();
        if chart.is_none()
            && let Some(href) = frame.object_href.as_deref()
        {
            let Some(target) = internal_ods_object_target(href) else {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::ExternalResourceBlocked,
                        Phase::Security,
                        Fidelity::Blocked,
                        "external or unsafe ODS chart object reference was blocked",
                    )
                    .in_part(CONTENT_PART),
                );
                continue;
            };
            diagnostic_part = target.clone();
            if let Some(package) = package {
                if let Some(bytes) = package.part(&target)? {
                    chart = parse_ods_charts_in_xml(&bytes, limits, &target)?
                        .into_iter()
                        .next()
                        .flatten();
                }
            }
        }
        let Some(chart) = chart else {
            let message = if frame.inline_chart_index.is_some() {
                "ODS chart has no usable inline table series and was omitted"
            } else {
                "ODS embedded chart has no usable bar, line, or pie table series and was omitted"
            };
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Render,
                    Fidelity::Omitted,
                    message,
                )
                .in_part(&diagnostic_part),
            );
            continue;
        };
        if ods_chart_has_3d_ribbons(&chart) {
            diagnostics.push(Diagnostic::warning(DiagnosticCode::ApproximateLayout, Phase::Render, Fidelity::Approximate,
                "ODF 3D ribbons use a normalized parallel view with flat shading; scene lighting is approximated").in_part(&chart.source_part));
        }
        if chart.three_dimensional && !ods_chart_has_3d_ribbons(&chart) {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Render,
                    Fidelity::Approximate,
                    "ODF 3D chart camera or family is unsupported; rendering available data in 2D",
                )
                .in_part(&chart.source_part),
            );
        }
        push_ods_chart(&chart, &frame, objects, limits.max_document_objects)?;
        if let Some(unit) = units.get_mut(frame.unit_index as usize) {
            unit.width = unit.width.max(frame.bounds.x + frame.bounds.width);
            unit.height = unit.height.max(frame.bounds.y + frame.bounds.height);
        }
    }
    Ok(())
}

fn internal_ods_object_target(href: &str) -> Option<String> {
    let target = internal_ods_image_target(href)?;
    Some(if target.ends_with(".xml") {
        target.to_owned()
    } else {
        format!("{target}/content.xml")
    })
}

fn push_ods_chart(
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    objects: &mut Vec<Object>,
    object_limit: usize,
) -> Result<(), Diagnostic> {
    let point_count = chart.series.iter().map(Vec::len).sum::<usize>();
    let axis_objects = if matches!(
        chart.kind,
        OdsChartKind::Scatter | OdsChartKind::Line | OdsChartKind::Area
    ) && !ods_chart_has_3d_ribbons(chart)
    {
        chart
            .x_axis
            .ticks()
            .len()
            .saturating_add(chart.y_axis.ticks().len())
            .saturating_mul(3)
    } else {
        0
    };
    let legend_count = if chart.legend.is_some() {
        (if chart.kind == OdsChartKind::Pie {
            chart.categories.len()
        } else {
            chart.series.len() + chart.regressions.len()
        })
        .saturating_mul(2)
        .saturating_add(1)
    } else {
        0
    };
    let label_count = if chart.kind == OdsChartKind::Pie
        && chart
            .data_labels
            .as_ref()
            .is_some_and(OdfChartDataLabels::is_enabled)
    {
        chart.series.first().map_or(0, |series| series.len())
    } else {
        0
    };
    if objects
        .len()
        .checked_add(point_count.saturating_mul(2).saturating_add(axis_objects))
        .and_then(|count| count.checked_add(legend_count.saturating_add(label_count)))
        .and_then(|count| {
            count.checked_add(if ods_chart_has_3d_ribbons(chart) {
                160
            } else {
                64
            })
        })
        .is_none_or(|count| count > object_limit)
    {
        return Err(object_limit_error(
            &chart.source_part,
            "ODS chart objects exceed the configured object limit",
        ));
    }
    let chart_id = u32::try_from(objects.len())
        .map_err(|_| format_error(&chart.source_part, "object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id: chart_id,
        parent_numeric_id: None,
        stable_id: format!("object:{chart_id}"),
        parent_stable_id: None,
        kind: ObjectKind::Group,
        unit_index: frame.unit_index,
        bounds: frame.bounds,
        z: i32::try_from(chart_id).unwrap_or(i32::MAX),
        text: None,
        source: ods_chart_source(chart, frame, None, None, MappingQuality::Derived),
        visual: Visual::PaintedShape {
            geometry: Geometry::Rectangle,
            fill: Paint::Solid(0xffff_ffff),
            stroke: Paint::Solid(0xd1d5_dbff),
            stroke_width: 1.0,
        },
    });
    let plot = chart
        .plot_area
        .map(|bounds| Rect {
            x: frame.bounds.x + bounds.x,
            y: frame.bounds.y + bounds.y,
            ..bounds
        })
        .unwrap_or(Rect {
            x: frame.bounds.x + frame.bounds.width * 0.10,
            y: frame.bounds.y + frame.bounds.height * 0.08,
            width: frame.bounds.width * 0.84,
            height: frame.bounds.height * 0.82,
        });
    if matches!(chart.kind, OdsChartKind::Bar | OdsChartKind::Radar) {
        for bounds in [
            Rect {
                x: plot.x,
                y: plot.y,
                width: 0.01,
                height: plot.height,
            },
            Rect {
                x: plot.x,
                y: plot.y + plot.height,
                width: plot.width,
                height: 0.01,
            },
        ] {
            push_ods_chart_shape(
                objects,
                chart_id,
                chart,
                frame,
                bounds,
                Geometry::Line,
                Paint::None,
                Paint::Solid(0x6b72_80ff),
                1.0,
                None,
                None,
                MappingQuality::Derived,
            )?;
        }
    }
    match chart.kind {
        OdsChartKind::Bar => push_ods_bar_chart(objects, chart_id, chart, frame, plot),
        OdsChartKind::Line | OdsChartKind::Scatter | OdsChartKind::Area => {
            if ods_chart_has_3d_ribbons(chart) {
                push_ods_3d_ribbons(objects, chart_id, chart, frame, plot)
            } else {
                push_ods_xy_chart(objects, chart_id, chart, frame, plot)
            }
        }
        OdsChartKind::Radar => push_ods_radar_chart(objects, chart_id, chart, frame, plot),
        OdsChartKind::Pie => push_ods_pie_chart(objects, chart_id, chart, frame, plot),
    }?;
    if let Some(title) = chart
        .title
        .as_deref()
        .map(str::trim)
        .filter(|title| !title.is_empty())
    {
        push_ods_chart_text(
            objects,
            chart_id,
            chart,
            frame,
            title,
            Rect {
                x: frame.bounds.x,
                y: frame.bounds.y + 2.0,
                width: frame.bounds.width,
                height: 18.0,
            },
            10.0,
        )?;
    }
    if let Some(legend) = &chart.legend {
        push_ods_chart_legend(objects, chart_id, chart, frame, legend)?;
    }
    Ok(())
}

fn push_ods_bar_chart(
    objects: &mut Vec<Object>,
    chart_id: u32,
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    plot: Rect,
) -> Result<(), Diagnostic> {
    let categories = chart.series.iter().map(Vec::len).max().unwrap_or(0);
    let maximum = chart
        .series
        .iter()
        .flatten()
        .copied()
        .filter(|value| *value > 0.0)
        .fold(0.0_f32, f32::max);
    if categories == 0 || maximum <= 0.0 {
        return Ok(());
    }
    let category_width = plot.width / categories as f32;
    let bar_width = category_width * 0.70 / chart.series.len().max(1) as f32;
    for (series_index, series) in chart.series.iter().enumerate() {
        for (category_index, value) in series.iter().copied().enumerate() {
            if value <= 0.0 {
                continue;
            }
            let height = plot.height * (value / maximum).clamp(0.0, 1.0);
            push_ods_chart_shape(
                objects,
                chart_id,
                chart,
                frame,
                Rect {
                    x: plot.x
                        + category_index as f32 * category_width
                        + category_width * 0.15
                        + series_index as f32 * bar_width,
                    y: plot.y + plot.height - height,
                    width: bar_width,
                    height,
                },
                Geometry::Rectangle,
                Paint::Solid(super::office_chart_palette_color(series_index)),
                Paint::None,
                0.0,
                Some(series_index),
                Some(category_index),
                MappingQuality::Exact,
            )?;
        }
    }
    Ok(())
}

fn ods_chart_has_3d_ribbons(chart: &OdsBasicChart) -> bool {
    chart.three_dimensional && chart.projection.is_some() && chart.kind == OdsChartKind::Scatter
}

fn push_ods_3d_ribbons(
    objects: &mut Vec<Object>,
    chart_id: u32,
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    plot: Rect,
) -> Result<(), Diagnostic> {
    let view = chart.projection.as_ref().expect("checked 3D camera");
    let project = |x, y, z| view.project([x, y, z], plot);
    let far_z = if view.depth([0.0, 0.0, 0.0]) < view.depth([0.0, 0.0, 1.0]) {
        0.0
    } else {
        1.0
    };
    let far_x = if view.depth([0.0, 0.0, 0.0]) < view.depth([1.0, 0.0, 0.0]) {
        0.0
    } else {
        1.0
    };
    let shape = |objects: &mut Vec<Object>, points: &[(f32, f32)], fill, stroke, series, point| {
        push_ods_chart_shape(
            objects,
            chart_id,
            chart,
            frame,
            plot,
            super::polygon_geometry(points),
            fill,
            stroke,
            0.7,
            series,
            point,
            MappingQuality::Derived,
        )
    };
    let grid = Paint::Solid(0xb3b3_b3ff);
    shape(
        objects,
        &[
            project(0.0, 0.0, 0.0),
            project(1.0, 0.0, 0.0),
            project(1.0, 0.0, 1.0),
            project(0.0, 0.0, 1.0),
        ],
        Paint::Solid(chart.floor_color),
        grid.clone(),
        None,
        None,
    )?;
    for points in [
        [
            project(0.0, 0.0, far_z),
            project(1.0, 0.0, far_z),
            project(1.0, 1.0, far_z),
            project(0.0, 1.0, far_z),
        ],
        [
            project(far_x, 0.0, 0.0),
            project(far_x, 0.0, 1.0),
            project(far_x, 1.0, 1.0),
            project(far_x, 1.0, 0.0),
        ],
    ] {
        shape(objects, &points, Paint::None, grid.clone(), None, None)?;
    }
    let label = |objects: &mut Vec<Object>,
                 text: &str,
                 p: (f32, f32),
                 axis: usize,
                 left: bool|
     -> Result<(), Diagnostic> {
        let font = chart.axis_fonts[axis];
        push_ods_chart_text(
            objects,
            chart_id,
            chart,
            frame,
            text,
            Rect {
                x: plot.x + p.0 - if left { font * 2.5 + 4.0 } else { font },
                y: plot.y + p.1 + if left { -font * 0.6 } else { 3.0 },
                width: font * if left { 2.5 } else { 2.0 },
                height: font * 1.5,
            },
            font,
        )?;
        if left {
            if let Some(Object {
                visual: Visual::RichText { align, .. },
                ..
            }) = objects.last_mut()
            {
                *align = TextAlign::End;
            }
        }
        Ok(())
    };
    let y_span = (chart.y_axis.maximum - chart.y_axis.minimum).max(f32::EPSILON);
    let x_span = (chart.x_axis.maximum - chart.x_axis.minimum).max(f32::EPSILON);
    for (axis_index, axis, span) in [(0, chart.x_axis, x_span), (1, chart.y_axis, y_span)] {
        let step = if axis.interval.is_finite() && axis.interval > 0.0 {
            axis.interval
        } else {
            super::nice_chart_step(
                span / (plot.height / (chart.axis_fonts[axis_index] * 2.0)).clamp(2.0, 10.0),
            )
        };
        let start = (axis.minimum / step).ceil() * step;
        for i in 0..=24 {
            let value = start + i as f32 * step;
            if value > axis.maximum + step * 0.001 {
                break;
            }
            let ratio = (value - axis.minimum) / span;
            let p = if axis_index == 0 {
                project(ratio, 0.0, 1.0 - far_z)
            } else {
                project(far_x, ratio, 1.0 - far_z)
            };
            label(
                objects,
                &super::odf_chart::format_chart_number(value),
                p,
                axis_index,
                axis_index == 1,
            )?;
            if axis_index == 1 && chart.grid_axes[1] {
                for points in [
                    [project(0.0, ratio, far_z), project(1.0, ratio, far_z)],
                    [project(far_x, ratio, 0.0), project(far_x, ratio, 1.0)],
                ] {
                    shape(objects, &points, Paint::None, grid.clone(), None, None)?;
                }
            }
        }
    }
    let mut faces = Vec::new();
    for (series_index, values) in chart.series.iter().enumerate() {
        let lane = if chart.deep { chart.series.len() } else { 1 } as f32;
        let center = if chart.deep {
            1.0 - (series_index as f32 + 0.5) / lane
        } else {
            0.5
        };
        let z0 = center - 0.25 / lane;
        let z1 = center + 0.25 / lane;
        if let Some(name) = chart.series_labels.get(series_index) {
            label(objects, name, project(1.0 - far_x, 0.0, center), 2, false)?;
        }
        let domain = chart
            .domains
            .get(series_index)
            .or_else(|| chart.domains.first());
        for (index, pair) in values.windows(2).enumerate() {
            let x = |i| {
                domain
                    .and_then(|values| values.get(i))
                    .copied()
                    .unwrap_or(i as f32)
            };
            let x0 = ((x(index) - chart.x_axis.minimum) / x_span).clamp(0.0, 1.0);
            let x1 = ((x(index + 1) - chart.x_axis.minimum) / x_span).clamp(0.0, 1.0);
            let y0 = ((pair[0] - chart.y_axis.minimum) / y_span).clamp(0.0, 1.0);
            let y1 = ((pair[1] - chart.y_axis.minimum) / y_span).clamp(0.0, 1.0);
            let points = [
                project(x0, y0, z0),
                project(x1, y1, z0),
                project(x1, y1, z1),
                project(x0, y0, z1),
            ];
            faces.push((
                view.depth([(x0 + x1) * 0.5, (y0 + y1) * 0.5, center]),
                series_index,
                index,
                points,
            ));
        }
    }
    faces.sort_by(|a, b| a.0.total_cmp(&b.0));
    for (_, series, index, points) in faces {
        let color = chart
            .series_colors
            .get(series)
            .copied()
            .unwrap_or_else(|| super::office_chart_palette_color(series));
        shape(
            objects,
            &points,
            Paint::Solid(color),
            Paint::None,
            Some(series),
            Some(index),
        )?;
    }
    Ok(())
}

fn push_ods_xy_chart(
    objects: &mut Vec<Object>,
    chart_id: u32,
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    plot: Rect,
) -> Result<(), Diagnostic> {
    let point_count = chart.series.iter().map(Vec::len).sum::<usize>();
    if point_count == 0 {
        return Ok(());
    }
    let x_min = chart.x_axis.minimum;
    let x_max = chart.x_axis.maximum;
    let y_min = chart.y_axis.minimum;
    let y_max = chart.y_axis.maximum;
    let x_span = (x_max - x_min).max(f32::EPSILON);
    let y_span = (y_max - y_min).max(f32::EPSILON);
    let map_x = |value: f32| plot.x + ((value - x_min) / x_span).clamp(0.0, 1.0) * plot.width;
    let map_y = |value: f32| {
        plot.y + plot.height - ((value - y_min) / y_span).clamp(0.0, 1.0) * plot.height
    };

    push_ods_chart_axes(objects, chart_id, chart, frame, plot, &map_x, &map_y)?;

    let filled = chart.kind == OdsChartKind::Area;
    let markers_only = chart.series_style.stroke_none;
    let symbol_square = chart
        .series_style
        .symbol_name
        .as_deref()
        .is_some_and(|name| name.contains("square") || name.contains("rect"));
    let marker_size = 7.0_f32;

    for (series_index, series) in chart.series.iter().enumerate() {
        let points = series
            .iter()
            .enumerate()
            .filter(|(_, value)| value.is_finite())
            .map(|(index, value)| {
                let x_value = chart
                    .domains
                    .get(series_index)
                    .and_then(|domains| domains.get(index).copied())
                    .unwrap_or(index as f32);
                (map_x(x_value), map_y(*value), index)
            })
            .collect::<Vec<_>>();
        if points.is_empty() {
            continue;
        }
        let color = chart
            .series_colors
            .get(series_index)
            .copied()
            .unwrap_or(chart.series_style.fill);
        if filled && points.len() >= 2 {
            let mut commands = vec![PathCommand::MoveTo {
                x: 0.0,
                y: plot.height,
            }];
            for (x, y, _) in &points {
                commands.push(PathCommand::LineTo {
                    x: *x - plot.x,
                    y: *y - plot.y,
                });
            }
            commands.push(PathCommand::LineTo {
                x: points.last().unwrap().0 - plot.x,
                y: plot.height,
            });
            commands.push(PathCommand::LineTo {
                x: 0.0,
                y: plot.height,
            });
            push_ods_chart_shape(
                objects,
                chart_id,
                chart,
                frame,
                plot,
                Geometry::Path {
                    fill_rule: FillRule::NonZero,
                    commands,
                },
                Paint::Solid(color & 0xffff_ff80),
                Paint::None,
                0.0,
                Some(series_index),
                None,
                MappingQuality::Derived,
            )?;
        }
        if !markers_only {
            for segment in points.windows(2) {
                let (x0, y0, category_index) = segment[0];
                let (x1, y1, _) = segment[1];
                let bounds = Rect {
                    x: x0.min(x1),
                    y: y0.min(y1),
                    width: (x1 - x0).abs().max(0.5),
                    height: (y1 - y0).abs().max(0.5),
                };
                push_ods_chart_shape(
                    objects,
                    chart_id,
                    chart,
                    frame,
                    bounds,
                    Geometry::Path {
                        fill_rule: FillRule::NonZero,
                        commands: super::odf_chart::line_segment_commands(
                            (x0 - bounds.x, y0 - bounds.y),
                            (x1 - bounds.x, y1 - bounds.y),
                            chart.series_style.interpolation.as_deref(),
                        ),
                    },
                    Paint::None,
                    Paint::Solid(chart.series_style.stroke.unwrap_or(color)),
                    chart.series_style.stroke_width,
                    Some(series_index),
                    Some(category_index),
                    MappingQuality::Exact,
                )?;
            }
        }
        for (x, y, category_index) in &points {
            let geometry = if symbol_square {
                Geometry::Rectangle
            } else {
                Geometry::Ellipse
            };
            push_ods_chart_shape(
                objects,
                chart_id,
                chart,
                frame,
                Rect {
                    x: *x - marker_size / 2.0,
                    y: *y - marker_size / 2.0,
                    width: marker_size,
                    height: marker_size,
                },
                geometry,
                Paint::Solid(color),
                Paint::None,
                0.0,
                Some(series_index),
                Some(*category_index),
                MappingQuality::Exact,
            )?;
        }
    }

    for regression in &chart.regressions {
        let xs = chart
            .domains
            .get(regression.series_index)
            .map(Vec::as_slice)
            .unwrap_or(&[]);
        let ys = chart
            .series
            .get(regression.series_index)
            .map(Vec::as_slice)
            .unwrap_or(&[]);
        if xs.len() < 2 || ys.len() < 2 {
            continue;
        }
        let n = xs.len().min(ys.len());
        let samples = (0..n).map(|i| (xs[i], ys[i])).collect::<Vec<_>>();
        let x_min =
            (xs.iter().copied().fold(f32::INFINITY, f32::min) - regression.backward).max(x_min);
        let x_max =
            (xs.iter().copied().fold(f32::NEG_INFINITY, f32::max) + regression.forward).min(x_max);
        if x_max <= x_min {
            continue;
        }
        // Keep the fitted curve intact; the plot clips it rather than clamping its points.
        let map_y = |value: f32| plot.y + plot.height - (value - y_min) / y_span * plot.height;
        let path = match regression.kind.as_str() {
            "exponential" => {
                exponential_regression_path(&samples, x_min, x_max, &map_x, &map_y, plot)
            }
            "mean-value" | "average" => {
                let mean = samples.iter().map(|(_, y)| *y).sum::<f32>() / n as f32;
                Some(vec![
                    PathCommand::MoveTo {
                        x: map_x(x_min) - plot.x,
                        y: map_y(mean) - plot.y,
                    },
                    PathCommand::LineTo {
                        x: map_x(x_max) - plot.x,
                        y: map_y(mean) - plot.y,
                    },
                ])
            }
            _ => linear_regression_path(&samples, x_min, x_max, &map_x, &map_y, plot),
        };
        let Some(commands) = path else {
            continue;
        };
        push_ods_chart_shape(
            objects,
            chart_id,
            chart,
            frame,
            plot,
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands,
            },
            Paint::None,
            Paint::Solid(regression.color),
            regression.stroke_width,
            Some(regression.series_index),
            None,
            MappingQuality::Derived,
        )?;
        let object = objects.last_mut().expect("regression object was pushed");
        object.visual = Visual::Effect {
            shadow: None,
            clip: Some(Geometry::Rectangle),
            visual: Box::new(object.visual.clone()),
        };
    }
    Ok(())
}

fn push_ods_chart_axes(
    objects: &mut Vec<Object>,
    chart_id: u32,
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    plot: Rect,
    map_x: &dyn Fn(f32) -> f32,
    map_y: &dyn Fn(f32) -> f32,
) -> Result<(), Diagnostic> {
    if let Some((color, width)) = chart.plot_border {
        push_ods_chart_shape(
            objects,
            chart_id,
            chart,
            frame,
            plot,
            Geometry::Rectangle,
            Paint::None,
            Paint::Solid(color),
            width,
            None,
            None,
            MappingQuality::Derived,
        )?;
    }
    let color = Paint::Solid(0xb3b3_b3ff);
    let shape = |objects: &mut Vec<Object>, bounds| {
        push_ods_chart_shape(
            objects,
            chart_id,
            chart,
            frame,
            bounds,
            Geometry::Rectangle,
            color.clone(),
            Paint::None,
            0.0,
            None,
            None,
            MappingQuality::Derived,
        )
    };
    // Grid belongs to its declaring axis; a y-axis grid produces horizontal lines.
    for (axis, scale, horizontal) in [(chart.x_axis, map_x, true), (chart.y_axis, map_y, false)] {
        let index = usize::from(!horizontal);
        if chart.grid_axes[index] {
            for value in axis.ticks() {
                let bounds = if horizontal {
                    Rect {
                        x: scale(value),
                        y: plot.y,
                        width: 0.5,
                        height: plot.height,
                    }
                } else {
                    Rect {
                        x: plot.x,
                        y: scale(value),
                        width: plot.width,
                        height: 0.5,
                    }
                };
                shape(objects, bounds)?;
            }
        }
    }
    let x_axis_y = if chart.y_axis.minimum < 0.0 && chart.y_axis.maximum > 0.0 {
        map_y(0.0)
    } else {
        plot.y + plot.height
    };
    let y_axis_x = if chart.x_axis.minimum < 0.0 && chart.x_axis.maximum > 0.0 {
        map_x(0.0)
    } else {
        plot.x
    };
    for bounds in [
        Rect {
            x: plot.x,
            y: x_axis_y,
            width: plot.width,
            height: 0.8,
        },
        Rect {
            x: y_axis_x,
            y: plot.y,
            width: 0.8,
            height: plot.height,
        },
    ] {
        shape(objects, bounds)?;
    }
    for (axis, scale, horizontal) in [(chart.x_axis, map_x, true), (chart.y_axis, map_y, false)] {
        let font = chart.axis_fonts[usize::from(!horizontal)];
        for value in axis.ticks() {
            let position = scale(value);
            let (tick, label) = if horizontal {
                (
                    Rect {
                        x: position,
                        y: x_axis_y,
                        width: 0.7,
                        height: 4.0,
                    },
                    Rect {
                        x: position - font * 1.5,
                        y: x_axis_y + 5.0,
                        width: font * 3.0,
                        height: font * 1.5,
                    },
                )
            } else {
                (
                    Rect {
                        x: y_axis_x - 4.0,
                        y: position,
                        width: 4.0,
                        height: 0.7,
                    },
                    Rect {
                        x: y_axis_x - font * 3.0 - 6.0,
                        y: position - font * 0.6,
                        width: font * 3.0,
                        height: font * 1.5,
                    },
                )
            };
            shape(objects, tick)?;
            push_ods_chart_text(
                objects,
                chart_id,
                chart,
                frame,
                &format_axis_tick(value),
                label,
                font,
            )?;
            if !horizontal {
                if let Some(Object {
                    visual: Visual::RichText { align, .. },
                    ..
                }) = objects.last_mut()
                {
                    *align = TextAlign::End;
                }
            }
        }
    }
    Ok(())
}

fn format_axis_tick(value: f32) -> String {
    if (value - value.round()).abs() < 0.001 {
        format!("{}", value.round() as i32)
    } else {
        format!("{value:.1}")
    }
}

fn push_ods_chart_text(
    objects: &mut Vec<Object>,
    chart_id: u32,
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    text: &str,
    bounds: Rect,
    font_size: f32,
) -> Result<(), Diagnostic> {
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(&chart.source_part, "object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: Some(chart_id),
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: Some(format!("object:{chart_id}")),
        kind: ObjectKind::TextBox,
        unit_index: frame.unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: Some(text.to_owned()),
        source: ods_chart_source(chart, frame, None, None, MappingQuality::Derived),
        visual: Visual::RichText {
            geometry: Geometry::Rectangle,
            fill: Paint::None,
            stroke: Paint::None,
            stroke_width: 0.0,
            align: TextAlign::Center,
            line_height: font_size * 1.2,
            runs: vec![TextRun {
                paint: None,
                east_asian_line_breaks: true,
                text: text.to_owned(),
                font_family: "Arial".to_owned(),
                font_size,
                color: 0x3333_33ff,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                highlight: 0,
                baseline_shift: 0.0,
                letter_spacing: 0.0,
                horizontal_scale: 1.0,
            }],
        },
    });
    Ok(())
}

fn linear_regression_path(
    samples: &[(f32, f32)],
    x_min: f32,
    x_max: f32,
    map_x: &dyn Fn(f32) -> f32,
    map_y: &dyn Fn(f32) -> f32,
    plot: Rect,
) -> Option<Vec<PathCommand>> {
    let (xs, ys): (Vec<_>, Vec<_>) = samples.iter().copied().unzip();
    let (slope, intercept, _) = super::linear_regression(&xs, &ys)?;
    let y0 = slope * x_min + intercept;
    let y1 = slope * x_max + intercept;
    Some(vec![
        PathCommand::MoveTo {
            x: map_x(x_min) - plot.x,
            y: map_y(y0) - plot.y,
        },
        PathCommand::LineTo {
            x: map_x(x_max) - plot.x,
            y: map_y(y1) - plot.y,
        },
    ])
}

fn exponential_regression_path(
    samples: &[(f32, f32)],
    x_min: f32,
    x_max: f32,
    map_x: &dyn Fn(f32) -> f32,
    map_y: &dyn Fn(f32) -> f32,
    plot: Rect,
) -> Option<Vec<PathCommand>> {
    let positive = samples
        .iter()
        .copied()
        .filter(|(x, y)| y.is_finite() && *y > 0.0 && x.is_finite())
        .collect::<Vec<_>>();
    if positive.len() < 2 {
        return None;
    }
    let (xs, log_ys): (Vec<_>, Vec<_>) = positive.iter().map(|(x, y)| (*x, y.ln())).unzip();
    let (slope, intercept, _) = super::linear_regression(&xs, &log_ys)?;
    let steps = 24;
    let mut commands = Vec::with_capacity(steps + 1);
    for step in 0..=steps {
        let x = x_min + (x_max - x_min) * step as f32 / steps as f32;
        let y = (intercept + slope * x).exp();
        let command = if step == 0 {
            PathCommand::MoveTo {
                x: map_x(x) - plot.x,
                y: map_y(y) - plot.y,
            }
        } else {
            PathCommand::LineTo {
                x: map_x(x) - plot.x,
                y: map_y(y) - plot.y,
            }
        };
        commands.push(command);
    }
    Some(commands)
}

fn push_ods_radar_chart(
    objects: &mut Vec<Object>,
    chart_id: u32,
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    plot: Rect,
) -> Result<(), Diagnostic> {
    let spokes = chart.series.iter().map(Vec::len).max().unwrap_or(0);
    let maximum = chart
        .series
        .iter()
        .flatten()
        .copied()
        .filter(|value| *value >= 0.0)
        .fold(0.0_f32, f32::max);
    if spokes < 3 || maximum <= 0.0 {
        return Ok(());
    }
    let center = (plot.x + plot.width / 2.0, plot.y + plot.height / 2.0);
    let radius = plot.width.min(plot.height) * 0.45;
    for (series_index, series) in chart.series.iter().enumerate() {
        let points = (0..spokes)
            .map(|index| {
                let value = series.get(index).copied().unwrap_or(0.0).max(0.0);
                let angle = -std::f32::consts::FRAC_PI_2
                    + std::f32::consts::TAU * index as f32 / spokes as f32;
                let r = radius * (value / maximum).clamp(0.0, 1.0);
                (
                    center.0 + r * angle.cos(),
                    center.1 + r * angle.sin(),
                    index,
                )
            })
            .collect::<Vec<_>>();
        let mut commands = Vec::with_capacity(points.len() + 1);
        for (index, (x, y, _)) in points.iter().enumerate() {
            let command = if index == 0 {
                PathCommand::MoveTo {
                    x: *x - plot.x,
                    y: *y - plot.y,
                }
            } else {
                PathCommand::LineTo {
                    x: *x - plot.x,
                    y: *y - plot.y,
                }
            };
            commands.push(command);
        }
        commands.push(PathCommand::LineTo {
            x: points[0].0 - plot.x,
            y: points[0].1 - plot.y,
        });
        push_ods_chart_shape(
            objects,
            chart_id,
            chart,
            frame,
            plot,
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands,
            },
            Paint::Solid(super::office_chart_palette_color(series_index) & 0xffff_ff40),
            Paint::Solid(super::office_chart_palette_color(series_index)),
            2.0,
            Some(series_index),
            None,
            MappingQuality::Derived,
        )?;
    }
    Ok(())
}

fn push_ods_pie_chart(
    objects: &mut Vec<Object>,
    chart_id: u32,
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    plot: Rect,
) -> Result<(), Diagnostic> {
    let Some(series) = chart.series.first() else {
        return Ok(());
    };
    let total = series
        .iter()
        .copied()
        .filter(|value| *value > 0.0)
        .sum::<f32>();
    if total <= 0.0 {
        return Ok(());
    }
    let size = plot.width.min(plot.height);
    let pie = Rect {
        x: plot.x + (plot.width - size) / 2.0,
        y: plot.y + (plot.height - size) / 2.0,
        width: size,
        height: size,
    };
    let center = size / 2.0;
    let radius = center * 0.94;
    let mut angle = -std::f32::consts::FRAC_PI_2;
    let labels = chart
        .data_labels
        .as_ref()
        .filter(|labels| labels.is_enabled());
    for (category_index, value) in series.iter().copied().enumerate() {
        if value <= 0.0 {
            continue;
        }
        let sweep = std::f32::consts::TAU * value / total;
        let mid_angle = angle + sweep / 2.0;
        let explosion = chart
            .point_explosions
            .get(category_index)
            .copied()
            .unwrap_or(0.0);
        let (offset_x, offset_y) = pie_explosion_offset(explosion, radius, mid_angle);
        let slice_cx = center + offset_x;
        let slice_cy = center + offset_y;
        let segments = ((sweep.abs() / std::f32::consts::TAU * 48.0).ceil() as usize).max(2);
        let mut points = Vec::with_capacity(segments + 2);
        points.push((slice_cx, slice_cy));
        for step in 0..=segments {
            let current = angle + sweep * step as f32 / segments as f32;
            points.push((
                slice_cx + radius * current.cos(),
                slice_cy + radius * current.sin(),
            ));
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
            x: pie.x + min_x,
            y: pie.y + min_y,
            width: (max_x - min_x).max(0.01),
            height: (max_y - min_y).max(0.01),
        };
        let mut commands = Vec::with_capacity(points.len() + 1);
        for (index, (x, y)) in points.iter().enumerate() {
            let x = x - min_x;
            let y = y - min_y;
            commands.push(if index == 0 {
                PathCommand::MoveTo { x, y }
            } else {
                PathCommand::LineTo { x, y }
            });
        }
        commands.push(PathCommand::ClosePath);
        push_ods_chart_shape(
            objects,
            chart_id,
            chart,
            frame,
            bounds,
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands,
            },
            Paint::Solid(
                chart
                    .colors
                    .get(category_index)
                    .copied()
                    .unwrap_or_else(|| super::office_chart_palette_color(category_index)),
            ),
            Paint::Solid(0xffff_ffff),
            1.0,
            Some(0),
            Some(category_index),
            MappingQuality::Exact,
        )?;
        if let Some(labels) = labels
            && let Some(text) = labels.format_label(
                chart
                    .categories
                    .get(category_index)
                    .map(String::as_str)
                    .unwrap_or_default(),
                value,
                total,
            )
        {
            push_ods_pie_data_label(
                objects,
                chart_id,
                chart,
                frame,
                pie,
                center,
                radius,
                offset_x,
                offset_y,
                mid_angle,
                category_index,
                labels,
                &text,
            )?;
        }
        angle += sweep;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_ods_pie_data_label(
    objects: &mut Vec<Object>,
    chart_id: u32,
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    pie: Rect,
    center: f32,
    radius: f32,
    offset_x: f32,
    offset_y: f32,
    mid_angle: f32,
    category_index: usize,
    labels: &OdfChartDataLabels,
    text: &str,
) -> Result<(), Diagnostic> {
    let align = TextAlign::Center;
    let font_size = labels.font_size.max(8.0);
    let width = (text.chars().count() as f32 * font_size * 0.62 + 8.0).max(font_size * 2.0);
    let height = font_size * 1.35;
    let (cx, cy) = pie_label_anchor(
        pie.x + center,
        pie.y + center,
        offset_x,
        offset_y,
        radius,
        mid_angle,
        labels.position.unwrap_or(OdfLabelPosition::Outside),
    );
    push_ods_chart_shape(
        objects,
        chart_id,
        chart,
        frame,
        Rect {
            x: cx - width / 2.0,
            y: cy - height / 2.0,
            width,
            height,
        },
        Geometry::Rectangle,
        Paint::None,
        Paint::None,
        0.0,
        Some(0),
        Some(category_index),
        MappingQuality::Exact,
    )?;
    let object = objects.last_mut().expect("data label object was pushed");
    object.kind = ObjectKind::TextBox;
    object.text = Some(text.to_owned());
    object.visual = Visual::TextLayout {
        layout: TextLayout {
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
            font_family: labels.font_family.clone(),
            font_size,
            color: labels.color,
            bold: labels.bold,
            italic: false,
            align,
        }),
    };
    Ok(())
}

fn push_ods_chart_legend(
    objects: &mut Vec<Object>,
    chart_id: u32,
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    legend: &OdfChartLegend,
) -> Result<(), Diagnostic> {
    let labels = if chart.kind == OdsChartKind::Pie {
        chart.categories.clone()
    } else {
        chart
            .series_labels
            .iter()
            .cloned()
            .chain(chart.regressions.iter().map(|curve| curve.name.clone()))
            .collect()
    };
    let colors = if chart.kind == OdsChartKind::Pie {
        chart.colors.clone()
    } else {
        chart
            .series_colors
            .clone()
            .into_iter()
            .chain(chart.regressions.iter().map(|curve| curve.color))
            .collect()
    };
    if labels.is_empty() {
        return Ok(());
    }
    let horizontal = legend.is_horizontal();
    let bounds = resolve_legend_bounds(legend, frame.bounds, &labels);
    let key_size = legend.font_size * 0.6;
    let start = objects.len();
    push_ods_chart_shape(
        objects,
        chart_id,
        chart,
        frame,
        Rect {
            x: frame.bounds.x + bounds.x,
            y: frame.bounds.y + bounds.y,
            ..bounds
        },
        Geometry::Rectangle,
        Paint::None,
        if chart.kind == OdsChartKind::Pie {
            Paint::Solid(0x0000_00ff)
        } else {
            Paint::None
        },
        1.0,
        None,
        None,
        MappingQuality::Exact,
    )?;
    if horizontal {
        let item_gap = 8.0;
        let item_widths: Vec<f32> = labels
            .iter()
            .map(|label| {
                key_size + 4.0 + label.chars().count() as f32 * legend.font_size * 0.55 + item_gap
            })
            .collect();
        let total_width: f32 = item_widths.iter().sum();
        let mut x = frame.bounds.x + bounds.x + ((bounds.width - total_width) / 2.0).max(0.0);
        let y = frame.bounds.y + bounds.y;
        for (index, label) in labels.iter().enumerate() {
            let item_width = item_widths[index];
            let color = colors
                .get(index)
                .copied()
                .unwrap_or_else(|| super::office_chart_palette_color(index));
            push_ods_chart_shape(
                objects,
                chart_id,
                chart,
                frame,
                Rect {
                    x,
                    y: y + (bounds.height - key_size) / 2.0,
                    width: key_size,
                    height: if chart.kind != OdsChartKind::Pie && index >= chart.series.len() {
                        1.5
                    } else {
                        key_size
                    },
                },
                Geometry::Rectangle,
                Paint::Solid(color),
                Paint::None,
                0.0,
                None,
                None,
                MappingQuality::Exact,
            )?;
            push_ods_chart_shape(
                objects,
                chart_id,
                chart,
                frame,
                Rect {
                    x: x + key_size + 4.0,
                    y,
                    width: (item_width - key_size - 4.0).max(0.0),
                    height: bounds.height,
                },
                Geometry::Rectangle,
                Paint::None,
                Paint::None,
                0.0,
                None,
                None,
                MappingQuality::Exact,
            )?;
            let text = objects.last_mut().expect("legend text object was pushed");
            text.kind = ObjectKind::TextBox;
            text.text = Some(label.clone());
            text.visual = Visual::TextLayout {
                layout: TextLayout {
                    inset_left: 0.0,
                    inset_right: 0.0,
                    inset_top: 0.0,
                    inset_bottom: 0.0,
                    wrap: false,
                    vertical_align: crate::model::TextVerticalAlign::Center,
                    ..TextLayout::default()
                },
                visual: Box::new(Visual::Text {
                    geometry: Geometry::Rectangle,
                    fill: 0,
                    stroke: 0,
                    stroke_width: 0.0,
                    font_family: legend.font_family.clone(),
                    font_size: legend.font_size,
                    color: legend.color,
                    bold: false,
                    italic: false,
                    align: TextAlign::Start,
                }),
            };
            x += item_width;
        }
    } else {
        // Pie keys name categories; the custom legend height wraps them in row-major order.
        let rows = ((bounds.height / (legend.font_size * 1.2)).floor() as usize)
            .max(1)
            .min(labels.len());
        let columns = labels.len().div_ceil(rows);
        let row_height = bounds.height / rows as f32;
        let column_width = bounds.width / columns as f32;
        for (index, label) in labels.iter().enumerate() {
            let x = frame.bounds.x + bounds.x + (index % columns) as f32 * column_width + 3.0;
            let y = frame.bounds.y + bounds.y + (index / columns) as f32 * row_height;
            let color = colors
                .get(index)
                .copied()
                .unwrap_or_else(|| super::office_chart_palette_color(index));
            push_ods_chart_shape(
                objects,
                chart_id,
                chart,
                frame,
                Rect {
                    x,
                    y: y + (row_height - key_size) / 2.0,
                    width: key_size,
                    height: if chart.kind != OdsChartKind::Pie && index >= chart.series.len() {
                        1.5
                    } else {
                        key_size
                    },
                },
                Geometry::Rectangle,
                Paint::Solid(color),
                Paint::None,
                0.0,
                None,
                None,
                MappingQuality::Exact,
            )?;
            push_ods_chart_shape(
                objects,
                chart_id,
                chart,
                frame,
                Rect {
                    x: x + key_size + 4.0,
                    y,
                    width: (column_width - key_size - 10.0).max(0.0),
                    height: row_height,
                },
                Geometry::Rectangle,
                Paint::None,
                Paint::None,
                0.0,
                None,
                None,
                MappingQuality::Exact,
            )?;
            let text = objects.last_mut().expect("legend text object was pushed");
            text.kind = ObjectKind::TextBox;
            text.text = Some(label.clone());
            text.visual = Visual::TextLayout {
                layout: TextLayout {
                    inset_left: 0.0,
                    inset_right: 0.0,
                    inset_top: 0.0,
                    inset_bottom: 0.0,
                    wrap: false,
                    vertical_align: crate::model::TextVerticalAlign::Center,
                    ..TextLayout::default()
                },
                visual: Box::new(Visual::Text {
                    geometry: Geometry::Rectangle,
                    fill: 0,
                    stroke: 0,
                    stroke_width: 0.0,
                    font_family: legend.font_family.clone(),
                    font_size: legend.font_size,
                    color: legend.color,
                    bold: false,
                    italic: false,
                    align: TextAlign::Start,
                }),
            };
        }
    }
    for object in &mut objects[start..] {
        if let SourceLocator::Ods {
            path, row, column, ..
        } = &mut object.source.locator
        {
            *path = format!(
                "{}/chart:legend[1]",
                path.split("/chart:series[").next().unwrap_or("")
            );
            *row = None;
            *column = None;
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_ods_chart_shape(
    objects: &mut Vec<Object>,
    parent_numeric_id: u32,
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    bounds: Rect,
    geometry: Geometry,
    fill: Paint,
    stroke: Paint,
    stroke_width: f32,
    series: Option<usize>,
    category: Option<usize>,
    mapping: MappingQuality,
) -> Result<(), Diagnostic> {
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(&chart.source_part, "object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: Some(parent_numeric_id),
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: Some(format!("object:{parent_numeric_id}")),
        kind: ObjectKind::Shape,
        unit_index: frame.unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: ods_chart_source(chart, frame, series, category, mapping),
        visual: Visual::PaintedShape {
            geometry,
            fill,
            stroke,
            stroke_width,
        },
    });
    Ok(())
}

fn ods_chart_source(
    chart: &OdsBasicChart,
    frame: &OdsChartFrame,
    series: Option<usize>,
    category: Option<usize>,
    mapping: MappingQuality,
) -> SourceRef {
    SourceRef {
        part: chart.source_part.clone(),
        mapping,
        locator: SourceLocator::Ods {
            kind: "element",
            table_name: frame.table_name.clone(),
            row: series.and_then(|index| u32::try_from(index).ok()),
            column: category.and_then(|index| u32::try_from(index).ok()),
            element_id: frame.element_id.clone(),
            path: format!(
                "/office:document-content/office:body/office:spreadsheet/table:table[{}]/draw:frame[{}]/chart:chart[1]/chart:series[{}]/chart:data-point[{}]",
                frame.unit_index + 1,
                frame.frame_index + 1,
                series.unwrap_or(0) + 1,
                category.unwrap_or(0) + 1,
            ),
        },
    }
}

fn parse_ods_conditional_formats(
    bytes: &[u8],
    limits: Limits,
) -> Result<(Vec<OdsConditionalFormat>, bool), Diagnostic> {
    #[derive(Debug)]
    struct FormatState {
        depth: usize,
        table_name: String,
        range: CellRange,
        rules: Vec<OdsConditionalRule>,
    }

    #[derive(Debug)]
    enum ComplexState {
        ColorScale {
            depth: usize,
            entries: Vec<(OdsConditionalValue, u32)>,
        },
        DataBar {
            depth: usize,
            thresholds: Vec<OdsConditionalValue>,
            positive_color: u32,
            negative_color: u32,
            max_length: f64,
        },
    }

    let mut depth = 0_usize;
    let mut current = None::<FormatState>;
    let mut complex = None::<ComplexState>;
    let mut formats = Vec::new();
    let mut unsupported = false;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                match local_name(name) {
                    "conditional-format" if current.is_none() => {
                        let target =
                            optional_attribute(&attributes, "target-range-address", CONTENT_PART)?;
                        if let Some(target) = target {
                            if let Some((table_name, range)) = parse_ods_cell_range(&target) {
                                if empty {
                                    unsupported = true;
                                } else {
                                    current = Some(FormatState {
                                        depth,
                                        table_name,
                                        range,
                                        rules: Vec::new(),
                                    });
                                }
                            } else {
                                unsupported = true;
                            }
                        } else {
                            unsupported = true;
                        }
                    }
                    "condition" if current.is_some() && complex.is_none() => {
                        let expression = optional_attribute(&attributes, "value", CONTENT_PART)?
                            .or(optional_attribute(&attributes, "condition", CONTENT_PART)?);
                        let style_name =
                            optional_attribute(&attributes, "apply-style-name", CONTENT_PART)?;
                        if let Some(icon_type) =
                            expression.as_deref().and_then(parse_ods_icon_set_type)
                        {
                            current
                                .as_mut()
                                .expect("conditional format is present")
                                .rules
                                .push(OdsConditionalRule::IconSet { icon_type });
                            return Ok(());
                        }
                        match (
                            expression.as_deref().and_then(parse_ods_value_comparison),
                            style_name.filter(|name| !name.is_empty()),
                        ) {
                            (Some((operator, threshold)), Some(style_name)) => current
                                .as_mut()
                                .expect("conditional format is present")
                                .rules
                                .push(OdsConditionalRule::CellIs {
                                    operator,
                                    threshold,
                                    style_name,
                                }),
                            _ => unsupported = true,
                        }
                    }
                    "color-scale" if current.is_some() && complex.is_none() => {
                        if empty {
                            unsupported = true;
                        } else {
                            complex = Some(ComplexState::ColorScale {
                                depth,
                                entries: Vec::new(),
                            });
                        }
                    }
                    "color-scale-entry"
                        if matches!(complex.as_ref(), Some(ComplexState::ColorScale { .. })) =>
                    {
                        let value = parse_ods_conditional_value(&attributes)?;
                        let color = optional_attribute(&attributes, "color", CONTENT_PART)?
                            .as_deref()
                            .map(|value| parse_odf_color(value, CONTENT_PART))
                            .transpose()?
                            .flatten();
                        match (value, color) {
                            (Some(value), Some(color)) => {
                                if let Some(ComplexState::ColorScale { entries, .. }) =
                                    complex.as_mut()
                                {
                                    entries.push((value, color));
                                }
                            }
                            _ => unsupported = true,
                        }
                    }
                    "data-bar" if current.is_some() && complex.is_none() => {
                        let positive_color =
                            optional_attribute(&attributes, "positive-color", CONTENT_PART)?
                                .as_deref()
                                .map(|value| parse_odf_color(value, CONTENT_PART))
                                .transpose()?
                                .flatten();
                        let negative_color =
                            optional_attribute(&attributes, "negative-color", CONTENT_PART)?
                                .as_deref()
                                .map(|value| parse_odf_color(value, CONTENT_PART))
                                .transpose()?
                                .flatten();
                        let max_length =
                            optional_attribute(&attributes, "max-length", CONTENT_PART)?
                                .as_deref()
                                .and_then(|value| value.parse::<f64>().ok())
                                .filter(|value| {
                                    value.is_finite() && *value > 0.0 && *value <= 100.0
                                })
                                .unwrap_or(100.0);
                        if let Some(positive_color) = positive_color {
                            if empty {
                                unsupported = true;
                            } else {
                                complex = Some(ComplexState::DataBar {
                                    depth,
                                    thresholds: Vec::new(),
                                    positive_color,
                                    negative_color: negative_color.unwrap_or(positive_color),
                                    max_length,
                                });
                            }
                        } else {
                            unsupported = true;
                        }
                    }
                    "formatting-entry"
                        if matches!(complex.as_ref(), Some(ComplexState::DataBar { .. })) =>
                    {
                        match parse_ods_conditional_value(&attributes)? {
                            Some(value) => {
                                if let Some(ComplexState::DataBar { thresholds, .. }) =
                                    complex.as_mut()
                                {
                                    thresholds.push(value);
                                }
                            }
                            None => unsupported = true,
                        }
                    }
                    _ => {}
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if local == "color-scale"
                    && matches!(
                        complex.as_ref(),
                        Some(ComplexState::ColorScale {
                            depth: start,
                            ..
                        }) if *start == depth
                    )
                {
                    let Some(ComplexState::ColorScale { entries, .. }) = complex.take() else {
                        unreachable!("checked color-scale state")
                    };
                    if (2..=3).contains(&entries.len()) {
                        current
                            .as_mut()
                            .expect("conditional format is present")
                            .rules
                            .push(OdsConditionalRule::ColorScale { entries });
                    } else {
                        unsupported = true;
                    }
                } else if local == "data-bar"
                    && matches!(
                        complex.as_ref(),
                        Some(ComplexState::DataBar {
                            depth: start,
                            ..
                        }) if *start == depth
                    )
                {
                    let Some(ComplexState::DataBar {
                        thresholds,
                        positive_color,
                        negative_color,
                        max_length,
                        ..
                    }) = complex.take()
                    else {
                        unreachable!("checked data-bar state")
                    };
                    if thresholds.len() == 2 {
                        current
                            .as_mut()
                            .expect("conditional format is present")
                            .rules
                            .push(OdsConditionalRule::DataBar {
                                thresholds,
                                positive_color,
                                negative_color,
                                max_length,
                            });
                    } else {
                        unsupported = true;
                    }
                } else if local == "conditional-format"
                    && current.as_ref().is_some_and(|state| state.depth == depth)
                {
                    let state = current.take().ok_or_else(|| {
                        format_error(CONTENT_PART, "conditional-format state was lost")
                    })?;
                    if state.rules.is_empty() {
                        unsupported = true;
                    } else {
                        formats.push(OdsConditionalFormat {
                            table_name: state.table_name,
                            range: state.range,
                            rules: state.rules,
                        });
                    }
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, CONTENT_PART))?;
    Ok((formats, unsupported))
}

fn parse_ods_cell_range(value: &str) -> Option<(String, CellRange)> {
    let value = value.trim().trim_start_matches('[').trim_end_matches(']');
    let (start, end) = value.split_once(':').unwrap_or((value, value));
    let (table_name, start_row, start_column) = parse_ods_cell_address(start, None)?;
    let (end_table, end_row, end_column) = parse_ods_cell_address(end, Some(&table_name))?;
    if end_table != table_name {
        return None;
    }
    Some((
        table_name,
        CellRange {
            start_row: start_row.min(end_row),
            end_row: start_row.max(end_row),
            start_column: start_column.min(end_column),
            end_column: start_column.max(end_column),
        },
    ))
}

fn parse_ods_cell_address(
    value: &str,
    inherited_table: Option<&str>,
) -> Option<(String, u32, u32)> {
    let normalized = value
        .trim()
        .trim_start_matches('[')
        .trim_end_matches(']')
        .replace('$', "");
    let (table, address) = match normalized.rsplit_once('.') {
        Some((table, address)) if !table.is_empty() => {
            (table.trim_matches('\'').to_owned(), address)
        }
        Some((_, address)) => (inherited_table?.to_owned(), address),
        None => (inherited_table?.to_owned(), normalized.as_str()),
    };
    let split = address.bytes().position(|byte| byte.is_ascii_digit())?;
    let (column, row) = address.split_at(split);
    if column.is_empty()
        || row.starts_with('0')
        || !column.bytes().all(|byte| byte.is_ascii_alphabetic())
        || !row.bytes().all(|byte| byte.is_ascii_digit())
    {
        return None;
    }
    let column = column.bytes().try_fold(0_u32, |value, byte| {
        value
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
    })?;
    let row = row.parse::<u32>().ok()?.checked_sub(1)?;
    let column = column.checked_sub(1)?;
    (row < MAX_ROWS && column < MAX_COLUMNS).then_some((table, row, column))
}

fn parse_ods_value_comparison(value: &str) -> Option<(OdsCellOperator, f64)> {
    let expression = value.trim();
    let suffix = expression.strip_prefix("cell-content()")?.trim();
    let (operator, threshold) = [
        (">=", OdsCellOperator::GreaterThanOrEqual),
        ("<=", OdsCellOperator::LessThanOrEqual),
        ("!=", OdsCellOperator::NotEqual),
        ("<>", OdsCellOperator::NotEqual),
        (">", OdsCellOperator::GreaterThan),
        ("<", OdsCellOperator::LessThan),
        ("=", OdsCellOperator::Equal),
    ]
    .into_iter()
    .find_map(|(prefix, operator)| suffix.strip_prefix(prefix).map(|value| (operator, value)))?;
    let threshold = threshold.trim().parse::<f64>().ok()?;
    threshold.is_finite().then_some((operator, threshold))
}

fn parse_ods_conditional_value(
    attributes: &[XmlAttribute<'_>],
) -> Result<Option<OdsConditionalValue>, Diagnostic> {
    let Some(value_type) = optional_attribute(attributes, "type", CONTENT_PART)? else {
        return Ok(None);
    };
    let number = || {
        optional_attribute(attributes, "value", CONTENT_PART).map(|value| {
            value
                .as_deref()
                .and_then(|value| value.parse::<f64>().ok())
                .filter(|value| value.is_finite())
        })
    };
    let value = match value_type.as_str() {
        "minimum" | "auto-minimum" => Some(OdsConditionalValue::Minimum),
        "maximum" | "auto-maximum" => Some(OdsConditionalValue::Maximum),
        "number" | "formula" => number()?.map(OdsConditionalValue::Number),
        "percent" => number()?
            .filter(|value| (0.0..=100.0).contains(value))
            .map(OdsConditionalValue::Percent),
        "percentile" => number()?
            .filter(|value| (0.0..=100.0).contains(value))
            .map(OdsConditionalValue::Percentile),
        _ => None,
    };
    Ok(value)
}

/// Read-only ODS `table:content-validation` annotation. Rules are never executed.
#[derive(Debug, Default)]
struct OdsContentValidation {
    name: Option<String>,
    condition: Option<String>,
    show_error_message: bool,
    show_input_message: bool,
    help_title: Option<String>,
    help_text: String,
    error_title: Option<String>,
    error_text: String,
}

fn parse_ods_icon_set_type(value: &str) -> Option<String> {
    let value = value.trim();
    let body = ["iconset(", "icon-set(", "icon_set("]
        .iter()
        .find_map(|prefix| value.strip_prefix(*prefix))?;
    let body = body.trim().trim_end_matches(')');
    let icon_type = body.split(';').next()?.trim();
    (!icon_type.is_empty()).then(|| icon_type.to_ascii_lowercase())
}

fn parse_ods_content_validations(
    bytes: &[u8],
    limits: Limits,
) -> Result<(Vec<OdsContentValidation>, bool), Diagnostic> {
    #[derive(Debug)]
    enum MessageKind {
        Help,
        Error,
    }

    let mut depth = 0_usize;
    let mut current = None::<(usize, OdsContentValidation)>;
    let mut message: Option<(MessageKind, usize, Option<String>, String)> = None;
    let mut validations = Vec::new();
    let mut active_content = false;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if matches!(local, "validity-macros" | "event-listener") {
                    active_content = true;
                }
                if local == "content-validation" && current.is_none() {
                    let show = |attr: &str| -> Result<bool, Diagnostic> {
                        Ok(optional_attribute(&attributes, attr, CONTENT_PART)?
                            .is_some_and(|value| matches!(value.as_str(), "true" | "1")))
                    };
                    current = Some((
                        depth,
                        OdsContentValidation {
                            name: optional_attribute(&attributes, "name", CONTENT_PART)?,
                            condition: optional_attribute(&attributes, "condition", CONTENT_PART)?,
                            show_error_message: show("show-error-message")?,
                            show_input_message: show("show-input-message")?,
                            ..OdsContentValidation::default()
                        },
                    ));
                    if empty {
                        push_ods_content_validation(&mut current, &mut validations, limits)?;
                    }
                } else if matches!(local, "help-message" | "error-message") && current.is_some() {
                    let kind = if local == "help-message" {
                        MessageKind::Help
                    } else {
                        MessageKind::Error
                    };
                    message = Some((
                        kind,
                        depth,
                        optional_attribute(&attributes, "title", CONTENT_PART)?,
                        String::new(),
                    ));
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::Text(text) => {
                if let Some((_, _, _, body)) = message.as_mut() {
                    body.push_str(
                        &decode_xml_text(text).map_err(|error| with_part(error, CONTENT_PART))?,
                    );
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some((_, _, _, body)) = message.as_mut() {
                    body.push_str(text);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if matches!(local, "help-message" | "error-message")
                    && message
                        .as_ref()
                        .is_some_and(|(_, message_depth, _, _)| *message_depth == depth)
                {
                    let (kind, _, title, body) = message.take().unwrap();
                    if let Some((_, validation)) = current.as_mut() {
                        let body = body.trim().to_owned();
                        match kind {
                            MessageKind::Help => {
                                validation.help_title = title;
                                validation.help_text = body;
                            }
                            MessageKind::Error => {
                                validation.error_title = title;
                                validation.error_text = body;
                            }
                        }
                    }
                }
                if local == "content-validation"
                    && current
                        .as_ref()
                        .is_some_and(|(validation_depth, _)| *validation_depth == depth)
                {
                    push_ods_content_validation(&mut current, &mut validations, limits)?;
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, CONTENT_PART))?;
    Ok((validations, active_content))
}

fn push_ods_content_validation(
    current: &mut Option<(usize, OdsContentValidation)>,
    validations: &mut Vec<OdsContentValidation>,
    limits: Limits,
) -> Result<(), Diagnostic> {
    let Some((_, validation)) = current.take() else {
        return Ok(());
    };
    if validations.len() >= limits.max_document_objects {
        return Err(object_limit_error(
            CONTENT_PART,
            "ODS content-validation count exceeds the configured object limit",
        ));
    }
    validations.push(validation);
    Ok(())
}

fn ods_content_validation_diagnostic(validation: &OdsContentValidation) -> Diagnostic {
    let mut parts = Vec::new();
    if let Some(name) = validation.name.as_deref().filter(|name| !name.is_empty()) {
        parts.push(format!("name={name}"));
    }
    if let Some(condition) = validation
        .condition
        .as_deref()
        .filter(|value| !value.is_empty())
    {
        parts.push(format!("condition={condition}"));
    }
    if validation.show_error_message {
        match (
            validation.error_title.as_deref(),
            validation.error_text.as_str(),
        ) {
            (Some(title), text) if !text.is_empty() => {
                parts.push(format!("error={title}: {text}"));
            }
            (Some(title), _) => parts.push(format!("error={title}")),
            (None, text) if !text.is_empty() => parts.push(format!("error={text}")),
            _ => {}
        }
    }
    if validation.show_input_message {
        match (
            validation.help_title.as_deref(),
            validation.help_text.as_str(),
        ) {
            (Some(title), text) if !text.is_empty() => {
                parts.push(format!("help={title}: {text}"));
            }
            (Some(title), _) => parts.push(format!("help={title}")),
            (None, text) if !text.is_empty() => parts.push(format!("help={text}")),
            _ => {}
        }
    }
    if parts.is_empty() {
        parts.push("unnamed rule".to_owned());
    }
    Diagnostic::warning(
        DiagnosticCode::UnsupportedFeature,
        Phase::Parse,
        Fidelity::Approximate,
        format!(
            "ODS data validation is a read-only annotation and is not enforced: {}",
            parts.join("; ")
        ),
    )
    .in_part(CONTENT_PART)
}

#[cfg(test)]
fn parse_content_xml(bytes: &[u8], limits: Limits) -> Result<ParsedContent, Diagnostic> {
    let mut styles = OdsStyles::default();
    parse_styles_xml(bytes, limits, CONTENT_PART, &mut styles)?;
    parse_content_xml_with_styles(bytes, limits, &styles, &HashMap::new())
}

fn parse_content_xml_with_styles(
    bytes: &[u8],
    limits: Limits,
    styles: &OdsStyles,
    frozen_panes: &HashMap<String, FreezePanes>,
) -> Result<ParsedContent, Diagnostic> {
    let (conditional_formats, unsupported_conditional) =
        parse_ods_conditional_formats(bytes, limits)?;
    let (content_validations, validation_active_content) =
        parse_ods_content_validations(bytes, limits)?;
    let mut depth = 0_usize;
    let mut table = None;
    let mut row = None;
    let mut cell = None;
    let mut covered_cell_depth = None;
    let mut ignored_chart_depth = None;
    let mut units = Vec::new();
    let mut objects = Vec::new();
    let mut diagnostics = Vec::new();
    if unsupported_conditional {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Omitted,
                "some ODS conditional-format rules could not be rendered",
            )
            .in_part(CONTENT_PART),
        );
    }
    for validation in &content_validations {
        diagnostics.push(ods_content_validation_diagnostic(validation));
    }
    if validation_active_content {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ActiveContentBlocked,
                Phase::Security,
                Fidelity::Omitted,
                "ODS content-validation macros and event listeners are blocked",
            )
            .in_part(CONTENT_PART),
        );
    }
    // Reserve authored names, including later sheets, before naming an unnamed sheet.
    let mut reserved_names = HashSet::new();
    parse_xml(bytes, limits, |event| {
        if let XmlEvent::StartElement {
            name, attributes, ..
        } = event
            && local_name(name) == "table"
            && let Some(name) = optional_attribute(&attributes, "name", CONTENT_PART)?
        {
            reserved_names.insert(name);
        }
        Ok(())
    })
    .map_err(|error| with_part(error, CONTENT_PART))?;
    let mut table_names = HashSet::new();
    let mut unsupported_value_types = HashSet::new();
    let mut unsupported_drawing_reported = false;
    let mut materialized_text_bytes = 0_usize;
    let mut conditional_cells = Vec::new();

    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                if ignored_chart_depth.is_some() {
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }
                if local_name(name) == "chart" {
                    ignored_chart_depth = (!empty).then_some(depth);
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }
                if matches!(
                    name,
                    "draw:g"
                        | "draw:custom-shape"
                        | "draw:rect"
                        | "draw:ellipse"
                        | "draw:line"
                        | "draw:path"
                ) && !unsupported_drawing_reported
                {
                    diagnostics.push(
                        Diagnostic::warning(
                            DiagnosticCode::UnsupportedFeature,
                            Phase::Render,
                            Fidelity::Omitted,
                            "ODS drawing shapes are not rendered",
                        )
                        .in_part(CONTENT_PART),
                    );
                    unsupported_drawing_reported = true;
                }
                match local_name(name) {
                    "table" => {
                        if table.is_some() {
                            return Err(format_error(CONTENT_PART, "nested sheets are invalid"));
                        }
                        if units.len() >= limits.max_relationship_edges {
                            return Err(Diagnostic::fatal(
                                DiagnosticCode::RelationshipLimit,
                                Phase::Parse,
                                None,
                                "spreadsheet exceeds the configured sheet limit",
                            )
                            .in_part(CONTENT_PART));
                        }
                        let name = match optional_attribute(&attributes, "name", CONTENT_PART)? {
                            Some(name) if !name.is_empty() => name,
                            _ => {
                                let mut suffix = units.len() + 1;
                                let name = loop {
                                    let candidate = format!("Sheet{suffix}");
                                    if reserved_names.insert(candidate.clone()) {
                                        break candidate;
                                    }
                                    suffix += 1;
                                };
                                diagnostics.push(
                                    Diagnostic::warning(
                                        DiagnosticCode::FormatInvalid,
                                        Phase::Parse,
                                        Fidelity::Approximate,
                                        "ODS sheet has no name; a display name was assigned",
                                    )
                                    .in_part(CONTENT_PART)
                                    .with_detail("sheet", &name),
                                );
                                name
                            }
                        };
                        if !table_names.insert(name.clone()) {
                            return Err(format_error(
                                CONTENT_PART,
                                "sheet names must be non-empty and unique",
                            ));
                        }
                        let source_index = u32::try_from(units.len()).map_err(|_| {
                            format_error(CONTENT_PART, "sheet count exceeds supported range")
                        })?;
                        let mut tab_color = None;
                        let style_name =
                            optional_attribute(&attributes, "style-name", CONTENT_PART)?;
                        let mut style = style_name.as_deref();
                        for _ in 0..32 {
                            let Some((parent, color)) =
                                style.and_then(|name| styles.tables.get(name))
                            else {
                                break;
                            };
                            if let Some(color) = color {
                                match parse_odf_color(color, CONTENT_PART) {
                                    Ok(color) => tab_color = color,
                                    Err(_) => diagnostics.push(
                                        Diagnostic::warning(
                                            DiagnosticCode::FormatInvalid,
                                            Phase::Parse,
                                            Fidelity::Omitted,
                                            "invalid ODS worksheet tab color was ignored",
                                        )
                                        .in_part(CONTENT_PART),
                                    ),
                                }
                                break;
                            }
                            style = parent.as_deref();
                        }
                        let state = TableState {
                            tab_color,
                            depth,
                            source_index,
                            unit_index: source_index,
                            name,
                            next_row: 0,
                            next_source_row: 0,
                            used_rows: 0,
                            used_columns: 0,
                            next_defined_column: 0,
                            column_sizes: HashMap::new(),
                            column_cell_styles: Vec::new(),
                            row_sizes: Vec::new(),
                            placements: Vec::new(),
                            merges: Vec::new(),
                        };
                        if empty {
                            units.push(finish_table(state, &mut objects, frozen_panes)?);
                        } else {
                            table = Some(state);
                        }
                    }
                    "table-column" => {
                        let current_table = table.as_mut().ok_or_else(|| {
                            format_error(CONTENT_PART, "column is not inside a sheet")
                        })?;
                        if row.is_some() {
                            return Err(format_error(
                                CONTENT_PART,
                                "column definition appears after row content started",
                            ));
                        }
                        let repeat = repeat_attribute(
                            &attributes,
                            "number-columns-repeated",
                            MAX_COLUMNS,
                            CONTENT_PART,
                        )?;
                        let end = current_table
                            .next_defined_column
                            .checked_add(repeat)
                            .filter(|end| *end <= MAX_COLUMNS)
                            .ok_or_else(|| {
                                dimension_error(
                                    CONTENT_PART,
                                    "column definitions exceed the column limit",
                                )
                            })?;
                        let hidden = optional_attribute(&attributes, "visibility", CONTENT_PART)?
                            .as_deref()
                            .is_some_and(|visibility| visibility != "visible");
                        let width = if hidden {
                            0.0
                        } else {
                            optional_attribute(&attributes, "style-name", CONTENT_PART)?
                                .as_deref()
                                .and_then(|name| styles.columns.get(name))
                                .copied()
                                .unwrap_or(COLUMN_WIDTH)
                        };
                        let default_cell_style = optional_attribute(
                            &attributes,
                            "default-cell-style-name",
                            CONTENT_PART,
                        )?;
                        for column in current_table.next_defined_column..end {
                            current_table.column_sizes.insert(column, width);
                        }
                        if let Some(name) = default_cell_style {
                            current_table.column_cell_styles.push((
                                current_table.next_defined_column,
                                end,
                                name,
                            ));
                        }
                        current_table.next_defined_column = end;
                    }
                    "table-row" => {
                        if row.is_some() {
                            return Err(format_error(CONTENT_PART, "nested rows are invalid"));
                        }
                        let current_table = table.as_mut().ok_or_else(|| {
                            format_error(CONTENT_PART, "row is not inside a sheet")
                        })?;
                        let repeat = repeat_attribute(
                            &attributes,
                            "number-rows-repeated",
                            MAX_ROWS,
                            CONTENT_PART,
                        )?;
                        let end_row = current_table
                            .next_row
                            .checked_add(repeat)
                            .filter(|end| *end <= MAX_ROWS)
                            .ok_or_else(|| {
                                dimension_error(CONTENT_PART, "repeated rows exceed the row limit")
                            })?;
                        let source_index = current_table.next_source_row;
                        current_table.next_source_row = current_table
                            .next_source_row
                            .checked_add(1)
                            .ok_or_else(|| {
                                dimension_error(CONTENT_PART, "source row count overflow")
                            })?;
                        let row_style =
                            optional_attribute(&attributes, "style-name", CONTENT_PART)?
                                .as_deref()
                                .and_then(|name| styles.rows.get(name));
                        let state = RowState {
                            depth,
                            source_index,
                            first_row: current_table.next_row,
                            repeat,
                            next_column: 0,
                            next_source_cell: 0,
                            height: if optional_attribute(&attributes, "visibility", CONTENT_PART)?
                                .as_deref()
                                .is_some_and(|visibility| visibility != "visible")
                            {
                                0.0
                            } else {
                                row_style.map(|(height, _)| *height).unwrap_or(ROW_HEIGHT)
                            },
                            automatic_height: row_style.is_none_or(|(_, optimal)| *optimal),
                            optimal_height: row_style.is_some_and(|(_, optimal)| *optimal),
                            default_cell_style: optional_attribute(
                                &attributes,
                                "default-cell-style-name",
                                CONTENT_PART,
                            )?,
                            cells: Vec::new(),
                        };
                        if empty {
                            finish_row(
                                current_table,
                                state,
                                &mut objects,
                                limits.max_document_objects,
                                &mut materialized_text_bytes,
                                limits.max_total_uncompressed_bytes,
                                styles,
                                &conditional_formats,
                                &mut conditional_cells,
                                &mut diagnostics,
                            )?;
                            debug_assert_eq!(current_table.next_row, end_row);
                        } else {
                            row = Some(state);
                        }
                    }
                    "table-cell" => {
                        if cell.is_some() || covered_cell_depth.is_some() {
                            return Err(format_error(CONTENT_PART, "nested cells are invalid"));
                        }
                        let current_row = row.as_mut().ok_or_else(|| {
                            format_error(CONTENT_PART, "cell is not inside a row")
                        })?;
                        let repeat = repeat_attribute(
                            &attributes,
                            "number-columns-repeated",
                            MAX_COLUMNS,
                            CONTENT_PART,
                        )?;
                        let first_column = current_row.next_column;
                        current_row.next_column = current_row
                            .next_column
                            .checked_add(repeat)
                            .filter(|end| *end <= MAX_COLUMNS)
                            .ok_or_else(|| {
                                dimension_error(
                                    CONTENT_PART,
                                    "repeated cells exceed the column limit",
                                )
                            })?;
                        let source_index = current_row.next_source_cell;
                        current_row.next_source_cell =
                            current_row.next_source_cell.checked_add(1).ok_or_else(|| {
                                dimension_error(CONTENT_PART, "source cell count overflow")
                            })?;
                        let state = CellState {
                            depth,
                            source_index,
                            first_column,
                            repeat,
                            column_span: repeat_attribute(
                                &attributes,
                                "number-columns-spanned",
                                MAX_COLUMNS,
                                CONTENT_PART,
                            )?,
                            row_span: repeat_attribute(
                                &attributes,
                                "number-rows-spanned",
                                MAX_ROWS,
                                CONTENT_PART,
                            )?,
                            style_name: optional_attribute(
                                &attributes,
                                "style-name",
                                CONTENT_PART,
                            )?,
                            element_id: optional_attribute(&attributes, "id", CONTENT_PART)?,
                            value_type: optional_attribute(
                                &attributes,
                                "value-type",
                                CONTENT_PART,
                            )?,
                            value: optional_attribute(&attributes, "value", CONTENT_PART)?,
                            string_value: optional_attribute(
                                &attributes,
                                "string-value",
                                CONTENT_PART,
                            )?,
                            date_value: optional_attribute(
                                &attributes,
                                "date-value",
                                CONTENT_PART,
                            )?,
                            boolean_value: optional_attribute(
                                &attributes,
                                "boolean-value",
                                CONTENT_PART,
                            )?,
                            paragraph_depth: None,
                            paragraph_count: 0,
                            paragraph_text: String::new(),
                        };
                        if empty {
                            current_row.cells.push(finish_cell(
                                state,
                                &mut diagnostics,
                                &mut unsupported_value_types,
                            )?);
                        } else {
                            cell = Some(state);
                        }
                    }
                    "covered-table-cell" => {
                        if cell.is_some() || covered_cell_depth.is_some() {
                            return Err(format_error(CONTENT_PART, "nested cells are invalid"));
                        }
                        let current_row = row.as_mut().ok_or_else(|| {
                            format_error(CONTENT_PART, "covered cell is not inside a row")
                        })?;
                        let repeat = repeat_attribute(
                            &attributes,
                            "number-columns-repeated",
                            MAX_COLUMNS,
                            CONTENT_PART,
                        )?;
                        current_row.next_column = current_row
                            .next_column
                            .checked_add(repeat)
                            .filter(|end| *end <= MAX_COLUMNS)
                            .ok_or_else(|| {
                                dimension_error(
                                    CONTENT_PART,
                                    "covered cells exceed the column limit",
                                )
                            })?;
                        current_row.next_source_cell =
                            current_row.next_source_cell.checked_add(1).ok_or_else(|| {
                                dimension_error(CONTENT_PART, "source cell count overflow")
                            })?;
                        if !empty {
                            covered_cell_depth = Some(depth);
                        }
                    }
                    "p" => {
                        if let Some(current_cell) = cell.as_mut() {
                            if current_cell.paragraph_depth.is_some() {
                                return Err(format_error(
                                    CONTENT_PART,
                                    "nested cell paragraphs are invalid",
                                ));
                            }
                            if current_cell.paragraph_count != 0 {
                                append_text(&mut current_cell.paragraph_text, "\n", limits)?;
                            }
                            current_cell.paragraph_count =
                                current_cell.paragraph_count.checked_add(1).ok_or_else(|| {
                                    dimension_error(CONTENT_PART, "paragraph count overflow")
                                })?;
                            current_cell.paragraph_depth = (!empty).then_some(depth);
                        }
                    }
                    "s" => {
                        if let Some(current_cell) = cell
                            .as_mut()
                            .filter(|current| current.paragraph_depth.is_some())
                        {
                            let count = repeat_attribute(
                                &attributes,
                                "c",
                                u32::try_from(limits.max_xml_bytes).unwrap_or(u32::MAX),
                                CONTENT_PART,
                            )?;
                            append_spaces(&mut current_cell.paragraph_text, count, limits)?;
                        }
                    }
                    "tab" => {
                        if let Some(current_cell) = cell
                            .as_mut()
                            .filter(|current| current.paragraph_depth.is_some())
                        {
                            append_text(&mut current_cell.paragraph_text, "\t", limits)?;
                        }
                    }
                    "line-break" => {
                        if let Some(current_cell) = cell
                            .as_mut()
                            .filter(|current| current.paragraph_depth.is_some())
                        {
                            append_text(&mut current_cell.paragraph_text, "\n", limits)?;
                        }
                    }
                    _ => {}
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if let Some(start) = ignored_chart_depth {
                    if local_name(name) == "chart" && start == depth {
                        ignored_chart_depth = None;
                    }
                    return Ok(());
                }
                match local_name(name) {
                    "p" => {
                        if let Some(current_cell) = cell.as_mut()
                            && current_cell.paragraph_depth == Some(depth)
                        {
                            current_cell.paragraph_depth = None;
                        }
                    }
                    "table-cell" => {
                        if cell
                            .as_ref()
                            .is_some_and(|current_cell| current_cell.depth == depth)
                        {
                            let closed_cell = cell.take().ok_or_else(|| {
                                format_error(CONTENT_PART, "cell state was lost before closing")
                            })?;
                            let finished = finish_cell(
                                closed_cell,
                                &mut diagnostics,
                                &mut unsupported_value_types,
                            )?;
                            row.as_mut()
                                .ok_or_else(|| {
                                    format_error(CONTENT_PART, "closed cell has no parent row")
                                })?
                                .cells
                                .push(finished);
                        }
                    }
                    "covered-table-cell" => {
                        if covered_cell_depth == Some(depth) {
                            covered_cell_depth = None;
                        }
                    }
                    "table-row" => {
                        if row
                            .as_ref()
                            .is_some_and(|current_row| current_row.depth == depth)
                        {
                            let finished = row.take().ok_or_else(|| {
                                format_error(CONTENT_PART, "row state was lost before closing")
                            })?;
                            finish_row(
                                table.as_mut().ok_or_else(|| {
                                    format_error(CONTENT_PART, "closed row has no parent sheet")
                                })?,
                                finished,
                                &mut objects,
                                limits.max_document_objects,
                                &mut materialized_text_bytes,
                                limits.max_total_uncompressed_bytes,
                                styles,
                                &conditional_formats,
                                &mut conditional_cells,
                                &mut diagnostics,
                            )?;
                        }
                    }
                    "table"
                        if table
                            .as_ref()
                            .is_some_and(|current_table| current_table.depth == depth) =>
                    {
                        if row.is_some() || cell.is_some() || covered_cell_depth.is_some() {
                            return Err(format_error(
                                CONTENT_PART,
                                "sheet ended before its row or cell",
                            ));
                        }
                        let closed_table = table.take().ok_or_else(|| {
                            format_error(CONTENT_PART, "sheet state was lost before closing")
                        })?;
                        units.push(finish_table(closed_table, &mut objects, frozen_panes)?);
                    }
                    _ => {}
                }
            }
            XmlEvent::Text(text) => {
                if let Some(current_cell) = cell
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    let decoded =
                        decode_xml_text(text).map_err(|error| with_part(error, CONTENT_PART))?;
                    append_text(&mut current_cell.paragraph_text, &decoded, limits)?;
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(current_cell) = cell
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    append_text(&mut current_cell.paragraph_text, text, limits)?;
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, CONTENT_PART))?;

    apply_ods_conditional_overlays(
        &conditional_formats,
        &conditional_cells,
        &mut objects,
        limits.max_document_objects,
    )?;

    Ok(ParsedContent {
        units,
        objects,
        diagnostics,
    })
}

fn finish_cell(
    cell: CellState,
    diagnostics: &mut Vec<Diagnostic>,
    unsupported_value_types: &mut HashSet<String>,
) -> Result<CellTemplate, Diagnostic> {
    let displayed_text = (cell.paragraph_count != 0 && !cell.paragraph_text.is_empty())
        .then(|| cell.paragraph_text.clone());
    let mut numeric_value = None;
    let text = match cell.value_type.as_deref() {
        None => displayed_text,
        Some("string") => displayed_text
            .or(cell.string_value.clone())
            .or(Some(String::new())),
        Some("float" | "currency" | "percentage") => {
            let value = required_cell_value(cell.value.as_deref(), "value", cell.source_index)?;
            let number = value
                .parse::<f64>()
                .ok()
                .filter(|number| number.is_finite());
            if number.is_none() {
                return Err(format_error(
                    CONTENT_PART,
                    "numeric cell has an invalid office:value",
                ));
            }
            numeric_value = number;
            displayed_text.or_else(|| Some(value.to_owned()))
        }
        Some("date") => {
            let value =
                required_cell_value(cell.date_value.as_deref(), "date-value", cell.source_index)?;
            displayed_text.or_else(|| Some(value.to_owned()))
        }
        Some("boolean") => {
            let value = required_cell_value(
                cell.boolean_value.as_deref(),
                "boolean-value",
                cell.source_index,
            )?;
            let normalized = match value {
                "true" | "1" => "true",
                "false" | "0" => "false",
                _ => {
                    return Err(format_error(
                        CONTENT_PART,
                        "boolean cell has an invalid office:boolean-value",
                    ));
                }
            };
            displayed_text.or_else(|| Some(normalized.to_owned()))
        }
        Some(value_type) => {
            if unsupported_value_types.insert(value_type.to_owned()) {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Parse,
                        Fidelity::Approximate,
                        format!(
                            "ODS value type {value_type} uses its displayed paragraph text only"
                        ),
                    )
                    .in_part(CONTENT_PART),
                );
            }
            displayed_text
        }
    };

    Ok(CellTemplate {
        source_index: cell.source_index,
        first_column: cell.first_column,
        repeat: cell.repeat,
        column_span: cell.column_span,
        row_span: cell.row_span,
        style_name: cell.style_name,
        element_id: cell.element_id,
        text,
        numeric_value,
        default_align: if numeric_value.is_some() || cell.value_type.as_deref() == Some("date") {
            TextAlign::End
        } else {
            TextAlign::Start
        },
    })
}

#[allow(clippy::too_many_arguments)]
fn finish_row(
    table: &mut TableState,
    mut row: RowState,
    objects: &mut Vec<Object>,
    object_limit: usize,
    materialized_text_bytes: &mut usize,
    materialized_text_limit: usize,
    styles: &OdsStyles,
    conditional_formats: &[OdsConditionalFormat],
    conditional_cells: &mut Vec<OdsConditionalCell>,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<(), Diagnostic> {
    let material_cells = row.cells.iter().try_fold(0_usize, |count, cell| {
        if cell.text.is_none()
            && cell.element_id.is_none()
            && cell.style_name.is_none()
            && cell.row_span == 1
            && cell.column_span == 1
        {
            return Ok(count);
        }
        let repeat = usize::try_from(cell.repeat).map_err(|_| {
            object_limit_error(CONTENT_PART, "cell repetition exceeds supported range")
        })?;
        count
            .checked_add(repeat)
            .ok_or_else(|| object_limit_error(CONTENT_PART, "repeated cell object count overflow"))
    })?;
    let mut repeated_rows = usize::try_from(row.repeat)
        .map_err(|_| object_limit_error(CONTENT_PART, "row repetition exceeds supported range"))?;
    let added_objects = material_cells
        .checked_mul(repeated_rows)
        .ok_or_else(|| object_limit_error(CONTENT_PART, "repeated row object count overflow"))?;
    if objects
        .len()
        .checked_add(added_objects)
        .is_none_or(|count| count > object_limit)
    {
        if row.repeat == 1 || material_cells > object_limit.saturating_sub(objects.len()) {
            return Err(object_limit_error(
                CONTENT_PART,
                "repeated cells exceed the configured object limit",
            ));
        }
        // Keep the authored row and its logical extent without expanding an oversized run.
        repeated_rows = 1;
        table.used_rows = table.used_rows.max(row.first_row + row.repeat);
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ObjectLimit,
                Phase::Parse,
                Fidelity::Omitted,
                "ODS repeated row expansion exceeds the object limit; only the first row is rendered",
            )
            .in_part(CONTENT_PART)
            .with_detail("sheet", &table.name)
            .with_detail("firstOmittedRow", (row.first_row + 2).to_string())
            .with_detail("omittedRowCount", (row.repeat - 1).to_string()),
        );
    }

    for cell in &row.cells {
        let Some(text) = cell.text.as_deref() else {
            continue;
        };
        let repeated_columns = usize::try_from(cell.repeat).map_err(|_| {
            Diagnostic::fatal(
                DiagnosticCode::ZipTotalSizeLimit,
                Phase::Parse,
                None,
                "materialized repeated cell count exceeds supported range",
            )
            .in_part(CONTENT_PART)
        })?;
        let copies = repeated_columns.checked_mul(repeated_rows).ok_or_else(|| {
            Diagnostic::fatal(
                DiagnosticCode::ZipTotalSizeLimit,
                Phase::Parse,
                None,
                "materialized repeated cell count overflow",
            )
            .in_part(CONTENT_PART)
        })?;
        reserve_materialized_text_bytes(
            materialized_text_bytes,
            text.len(),
            copies,
            materialized_text_limit,
            CONTENT_PART,
        )?;
    }

    if row.automatic_height && row.height > 0.0 {
        // ponytail: explicit line breaks only; wrapped text needs measured column-width layout.
        for cell in row.cells.iter().filter(|cell| cell.row_span == 1) {
            let lines = cell
                .text
                .as_deref()
                .filter(|text| !text.is_empty())
                .map_or(0, |text| text.split('\n').count());
            // Undeclared rows retain the shared default; authored optimal rows also fit one line.
            if lines > 1 || (lines == 1 && row.optimal_height) {
                let font_size = cell
                    .style_name
                    .as_deref()
                    .and_then(|name| styles.cells.get(name))
                    .map_or_else(
                        || OdsCellStyle::default().font_size,
                        |style| style.font_size,
                    );
                // Match Visual::Text's 1.2 line spacing and 4px top/bottom insets.
                row.height = row.height.max(font_size * 1.2 * lines as f32 + 8.0);
            }
        }
    }
    if row.height != ROW_HEIGHT {
        table.row_sizes.push(SizeSpan {
            start: row.first_row,
            end: row.first_row + row.repeat - 1,
            size: row.height,
        });
    }
    for row_offset in 0..repeated_rows as u32 {
        let logical_row = row.first_row + row_offset;
        table.merges.retain(|merged| merged.end_row >= logical_row);
        for cell in &row.cells {
            if cell.text.is_none()
                && cell.element_id.is_none()
                && cell.style_name.is_none()
                && cell.row_span == 1
                && cell.column_span == 1
            {
                continue;
            }
            for column_offset in 0..cell.repeat {
                let logical_column = cell.first_column + column_offset;
                if table.merges.iter().any(|merged| {
                    logical_row >= merged.start_row
                        && logical_row <= merged.end_row
                        && logical_column >= merged.start_column
                        && logical_column <= merged.end_column
                }) {
                    continue;
                }
                let end_row = logical_row
                    .checked_add(cell.row_span - 1)
                    .filter(|end| *end < MAX_ROWS)
                    .ok_or_else(|| {
                        dimension_error(CONTENT_PART, "merged row span exceeds its limit")
                    })?;
                let end_column = logical_column
                    .checked_add(cell.column_span - 1)
                    .filter(|end| *end < MAX_COLUMNS)
                    .ok_or_else(|| {
                        dimension_error(CONTENT_PART, "merged column span exceeds its limit")
                    })?;
                let range = CellRange {
                    start_row: logical_row,
                    end_row,
                    start_column: logical_column,
                    end_column,
                };
                if (cell.row_span != 1 || cell.column_span != 1)
                    && table
                        .merges
                        .iter()
                        .any(|merged| ranges_overlap(*merged, range))
                {
                    return Err(format_error(CONTENT_PART, "merged cell ranges overlap"));
                }
                if cell.row_span != 1 || cell.column_span != 1 {
                    table.merges.push(range);
                }
                table.used_rows = table.used_rows.max(end_row + 1);
                table.used_columns = table.used_columns.max(end_column + 1);
                let visible_column = (logical_column..=end_column).any(|column| {
                    table
                        .column_sizes
                        .get(&column)
                        .copied()
                        .unwrap_or(COLUMN_WIDTH)
                        > 0.0
                });
                if !visible_column || (row.height == 0.0 && cell.row_span == 1) {
                    continue;
                }
                let base_style = cell
                    .style_name
                    .as_deref()
                    .or(row.default_cell_style.as_deref())
                    .or_else(|| {
                        table
                            .column_cell_styles
                            .iter()
                            .find(|(start, end, _)| {
                                logical_column >= *start && logical_column < *end
                            })
                            .map(|(_, _, name)| name.as_str())
                    })
                    .and_then(|name| styles.cells.get(name))
                    .cloned()
                    .unwrap_or_default();
                let mut style = conditional_cell_style(
                    &table.name,
                    logical_row,
                    logical_column,
                    cell.numeric_value,
                    &base_style,
                    styles,
                    conditional_formats,
                );
                if style.align.is_none() {
                    let mut parent = style.parent_style.as_deref();
                    // Match number-style inheritance's bound, including cyclic styles.
                    for _ in 0..32 {
                        let Some(inherited) = parent.and_then(|name| styles.cells.get(name)) else {
                            break;
                        };
                        if inherited.align.is_some() {
                            style.align = inherited.align;
                            break;
                        }
                        parent = inherited.parent_style.as_deref();
                    }
                }
                let object_index = objects.len();
                push_cell(
                    table,
                    &row,
                    cell,
                    logical_row,
                    logical_column,
                    &style,
                    objects,
                )?;
                if let Some(value) = cell.numeric_value {
                    let formatted = match number_style_name(&style, styles) {
                        Some(name) => resolve_number_style(name, styles, value)
                            .and_then(|number| number.decimal.as_ref())
                            .map(|format| format.render(value)),
                        // Only unstyled plain numeric text is General. Never reinterpret
                        // fractions, scientific notation, dates or currency literals.
                        None => cell
                            .text
                            .as_deref()
                            .filter(|text| {
                                text.bytes().all(|b| {
                                    b.is_ascii_digit() || matches!(b, b'-' | b'+' | b'.' | b',')
                                }) && text.replace(',', ".").parse::<f64>().ok() == Some(value)
                            })
                            .map(|_| value.to_string()),
                    };
                    if let Some(formatted) = formatted {
                        let old_len = objects[object_index].text.as_ref().map_or(0, String::len);
                        reserve_materialized_text_bytes(
                            materialized_text_bytes,
                            formatted.len().saturating_sub(old_len),
                            1,
                            materialized_text_limit,
                            CONTENT_PART,
                        )?;
                        objects[object_index].text = Some(formatted);
                    } else if number_style_name(&style, styles).is_some() {
                        const MESSAGE: &str =
                            "ODS unsupported numeric format preserves cached display text";
                        if !diagnostics.iter().any(|item| item.message == MESSAGE) {
                            diagnostics.push(
                                Diagnostic::warning(
                                    DiagnosticCode::UnsupportedFeature,
                                    Phase::Parse,
                                    Fidelity::Approximate,
                                    MESSAGE,
                                )
                                .in_part(CONTENT_PART),
                            );
                        }
                    }
                }
                if let Some(value) = cell.numeric_value
                    && let Some(text) = objects[object_index].text.as_deref()
                    && let Some(fill_character) = number_fill_character(&style, styles, text, value)
                {
                    let object = &mut objects[object_index];
                    object.visual = Visual::TextLayout {
                        layout: TextLayout {
                            fill_character: Some(fill_character),
                            wrap: false,
                            ..TextLayout::default()
                        },
                        visual: Box::new(std::mem::replace(&mut object.visual, Visual::None)),
                    };
                }
                if let Some(value) = cell.numeric_value {
                    conditional_cells.push(OdsConditionalCell {
                        object_index,
                        unit_index: table.unit_index,
                        table_name: table.name.clone(),
                        row: logical_row,
                        column: logical_column,
                        value,
                    });
                }
            }
        }
    }
    table.next_row = row.first_row + row.repeat;
    Ok(())
}

fn conditional_cell_style(
    table_name: &str,
    row: u32,
    column: u32,
    value: Option<f64>,
    base_style: &OdsCellStyle,
    styles: &OdsStyles,
    formats: &[OdsConditionalFormat],
) -> OdsCellStyle {
    let Some(value) = value else {
        return base_style.clone();
    };
    for format in formats.iter().filter(|format| {
        format.table_name == table_name
            && row >= format.range.start_row
            && row <= format.range.end_row
            && column >= format.range.start_column
            && column <= format.range.end_column
    }) {
        for rule in &format.rules {
            if let OdsConditionalRule::CellIs {
                operator,
                threshold,
                style_name,
            } = rule
            {
                let matches = ods_value_matches(value, *operator, *threshold);
                if matches && let Some(style) = styles.cells.get(style_name) {
                    return style.clone();
                }
            }
        }
    }
    base_style.clone()
}

fn apply_ods_conditional_overlays(
    formats: &[OdsConditionalFormat],
    cells: &[OdsConditionalCell],
    objects: &mut Vec<Object>,
    object_limit: usize,
) -> Result<(), Diagnostic> {
    for format in formats {
        let matching = cells
            .iter()
            .filter(|cell| ods_conditional_contains(format, cell))
            .collect::<Vec<_>>();
        let stats = ods_conditional_statistics(&matching);
        for cell in matching {
            for rule in &format.rules {
                match rule {
                    OdsConditionalRule::CellIs { .. } => {}
                    OdsConditionalRule::ColorScale { entries } => {
                        let fill = ods_color_scale_color(cell.value, entries, &stats);
                        if let Some(Object {
                            visual: Visual::Text { fill: current, .. },
                            ..
                        }) = objects.get_mut(cell.object_index)
                        {
                            *current = fill;
                        }
                    }
                    OdsConditionalRule::DataBar {
                        thresholds,
                        positive_color,
                        negative_color,
                        max_length,
                    } => {
                        let minimum = thresholds
                            .first()
                            .map_or(stats.minimum, |value| ods_conditional_value(*value, &stats));
                        let maximum = thresholds
                            .last()
                            .map_or(stats.maximum, |value| ods_conditional_value(*value, &stats));
                        if maximum <= minimum {
                            continue;
                        }
                        let position =
                            ((cell.value - minimum) / (maximum - minimum)).clamp(0.0, 1.0);
                        let axis = ((0.0 - minimum) / (maximum - minimum)).clamp(0.0, 1.0);
                        let (start, end, color) = if cell.value < 0.0 {
                            (position, axis, *negative_color)
                        } else {
                            (axis, position, *positive_color)
                        };
                        if end <= start {
                            continue;
                        }
                        let Some(cell_bounds) =
                            objects.get(cell.object_index).map(|object| object.bounds)
                        else {
                            return Err(format_error(
                                CONTENT_PART,
                                "conditional formatting references a missing cell object",
                            ));
                        };
                        let inset = cell_bounds.height.min(cell_bounds.width) * 0.10;
                        let available = (cell_bounds.width - inset * 2.0).max(0.0)
                            * (*max_length / 100.0) as f32;
                        push_ods_conditional_bar(
                            objects,
                            object_limit,
                            cell,
                            Rect {
                                x: cell_bounds.x + inset + available * start as f32,
                                y: cell_bounds.y + cell_bounds.height * 0.22,
                                width: available * (end - start) as f32,
                                height: cell_bounds.height * 0.56,
                            },
                            (color & 0xffff_ff00) | 0x99,
                        )?;
                    }
                    OdsConditionalRule::IconSet { icon_type } => {
                        let position = if stats.maximum > stats.minimum {
                            ((cell.value - stats.minimum) / (stats.maximum - stats.minimum))
                                .clamp(0.0, 1.0)
                        } else {
                            1.0
                        };
                        let (color, high) = if position >= 2.0 / 3.0 {
                            (0x00c8_53ff, true)
                        } else if position >= 1.0 / 3.0 {
                            (0xffd6_00ff, true)
                        } else {
                            (0xff17_44ff, false)
                        };
                        let Some(cell_bounds) =
                            objects.get(cell.object_index).map(|object| object.bounds)
                        else {
                            return Err(format_error(
                                CONTENT_PART,
                                "conditional formatting references a missing cell object",
                            ));
                        };
                        let size = cell_bounds.height.min(cell_bounds.width) * 0.36;
                        let geometry = if icon_type.contains("arrow") {
                            let (tip_y, base_y) = if high { (0.0, size) } else { (size, 0.0) };
                            Geometry::Path {
                                fill_rule: FillRule::NonZero,
                                commands: vec![
                                    PathCommand::MoveTo {
                                        x: size / 2.0,
                                        y: tip_y,
                                    },
                                    PathCommand::LineTo { x: size, y: base_y },
                                    PathCommand::LineTo { x: 0.0, y: base_y },
                                    PathCommand::LineTo {
                                        x: size / 2.0,
                                        y: tip_y,
                                    },
                                ],
                            }
                        } else {
                            Geometry::Ellipse
                        };
                        push_ods_conditional_icon(
                            objects,
                            object_limit,
                            cell,
                            Rect {
                                x: cell_bounds.x + cell_bounds.width - size - 1.0,
                                y: cell_bounds.y + (cell_bounds.height - size) / 2.0,
                                width: size,
                                height: size,
                            },
                            geometry,
                            color,
                            icon_type,
                        )?;
                    }
                }
            }
        }
    }
    Ok(())
}

fn ods_conditional_contains(format: &OdsConditionalFormat, cell: &OdsConditionalCell) -> bool {
    format.table_name == cell.table_name
        && cell.row >= format.range.start_row
        && cell.row <= format.range.end_row
        && cell.column >= format.range.start_column
        && cell.column <= format.range.end_column
}

fn ods_conditional_statistics(cells: &[&OdsConditionalCell]) -> OdsConditionalStats {
    let mut sorted = cells.iter().map(|cell| cell.value).collect::<Vec<_>>();
    sorted.sort_by(f64::total_cmp);
    OdsConditionalStats {
        minimum: sorted.first().copied().unwrap_or(0.0),
        maximum: sorted.last().copied().unwrap_or(0.0),
        sorted,
    }
}

fn ods_conditional_value(value: OdsConditionalValue, stats: &OdsConditionalStats) -> f64 {
    match value {
        OdsConditionalValue::Minimum => stats.minimum,
        OdsConditionalValue::Maximum => stats.maximum,
        OdsConditionalValue::Number(value) => value,
        OdsConditionalValue::Percent(percent) => {
            stats.minimum + (stats.maximum - stats.minimum) * percent / 100.0
        }
        OdsConditionalValue::Percentile(percent) => {
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

fn ods_value_matches(value: f64, operator: OdsCellOperator, threshold: f64) -> bool {
    match operator {
        OdsCellOperator::LessThan => value < threshold,
        OdsCellOperator::LessThanOrEqual => value <= threshold,
        OdsCellOperator::Equal => value == threshold,
        OdsCellOperator::NotEqual => value != threshold,
        OdsCellOperator::GreaterThanOrEqual => value >= threshold,
        OdsCellOperator::GreaterThan => value > threshold,
    }
}

fn ods_color_scale_color(
    value: f64,
    entries: &[(OdsConditionalValue, u32)],
    stats: &OdsConditionalStats,
) -> u32 {
    let points = entries
        .iter()
        .map(|(threshold, color)| (ods_conditional_value(*threshold, stats), *color))
        .collect::<Vec<_>>();
    if points
        .first()
        .is_some_and(|(threshold, _)| value <= *threshold)
    {
        return points[0].1;
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
            return interpolate_ods_color(start_color, end_color, progress);
        }
    }
    points.last().map_or(0xffff_ffff, |(_, color)| *color)
}

fn interpolate_ods_color(start: u32, end: u32, progress: f64) -> u32 {
    let channel = |shift: u32| {
        let start = f64::from((start >> shift) & 0xff);
        let end = f64::from((end >> shift) & 0xff);
        (start + (end - start) * progress).round() as u32
    };
    (channel(24) << 24) | (channel(16) << 16) | (channel(8) << 8) | channel(0)
}

fn push_ods_conditional_bar(
    objects: &mut Vec<Object>,
    object_limit: usize,
    cell: &OdsConditionalCell,
    bounds: Rect,
    color: u32,
) -> Result<(), Diagnostic> {
    if objects.len() >= object_limit {
        return Err(object_limit_error(
            CONTENT_PART,
            "ODS conditional formatting exceeds the configured object limit",
        ));
    }
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: None,
        kind: ObjectKind::Shape,
        unit_index: cell.unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping: MappingQuality::Derived,
            locator: SourceLocator::Ods {
                kind: "cell",
                table_name: cell.table_name.clone(),
                row: Some(cell.row),
                column: Some(cell.column),
                element_id: None,
                path: format!(
                    "/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[{}]/table:table-cell[{}]/calcext:data-bar",
                    cell.row + 1,
                    cell.column + 1
                ),
            },
        },
        visual: Visual::Shape {
            geometry: Geometry::Rectangle,
            fill: color,
            stroke: 0,
            stroke_width: 0.0,
        },
    });
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_ods_conditional_icon(
    objects: &mut Vec<Object>,
    object_limit: usize,
    cell: &OdsConditionalCell,
    bounds: Rect,
    geometry: Geometry,
    color: u32,
    icon_type: &str,
) -> Result<(), Diagnostic> {
    if objects.len() >= object_limit {
        return Err(object_limit_error(
            CONTENT_PART,
            "ODS conditional formatting exceeds the configured object limit",
        ));
    }
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: None,
        kind: ObjectKind::Shape,
        unit_index: cell.unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping: MappingQuality::Derived,
            locator: SourceLocator::Ods {
                kind: "cell",
                table_name: cell.table_name.clone(),
                row: Some(cell.row),
                column: Some(cell.column),
                element_id: None,
                path: format!(
                    "/office:document-content/office:body/office:spreadsheet/table:table/table:table-row[{}]/table:table-cell[{}]/calcext:icon-set[{icon_type}]",
                    cell.row + 1,
                    cell.column + 1
                ),
            },
        },
        visual: Visual::Shape {
            geometry,
            fill: color,
            stroke: 0,
            stroke_width: 0.0,
        },
    });
    Ok(())
}

fn push_cell(
    table: &mut TableState,
    row: &RowState,
    cell: &CellTemplate,
    logical_row: u32,
    logical_column: u32,
    style: &OdsCellStyle,
    objects: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| object_limit_error(CONTENT_PART, "object count exceeds supported range"))?;
    let repeated = row.repeat != 1 || cell.repeat != 1;
    let path = format!(
        "/office:document-content/office:body/office:spreadsheet/table:table[{}]/table:table-row[{}]/table:table-cell[{}]",
        table.source_index + 1,
        row.source_index + 1,
        cell.source_index + 1,
    );
    let object_index = objects.len();
    let visual = Visual::Text {
        geometry: Geometry::Rectangle,
        fill: style.fill,
        stroke: style.stroke,
        stroke_width: style.stroke_width,
        font_family: style.font_family.clone(),
        font_size: style.font_size,
        color: style.color,
        bold: style.bold,
        italic: style.italic,
        align: style.align.unwrap_or(cell.default_align),
    };
    let visual = if style.stroke_style == StrokeStyle::default() {
        visual
    } else {
        Visual::StrokeStyle {
            style: style.stroke_style.clone(),
            visual: Box::new(visual),
        }
    };
    objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: None,
        kind: ObjectKind::Cell,
        unit_index: table.unit_index,
        bounds: Rect {
            x: logical_column as f32 * COLUMN_WIDTH,
            y: logical_row as f32 * ROW_HEIGHT,
            width: COLUMN_WIDTH,
            height: ROW_HEIGHT,
        },
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: cell
            .text
            .as_deref()
            .map(|text| clone_materialized_text(text, CONTENT_PART))
            .transpose()?,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping: if repeated {
                MappingQuality::Derived
            } else {
                MappingQuality::Exact
            },
            locator: SourceLocator::Ods {
                kind: "cell",
                table_name: table.name.clone(),
                row: Some(logical_row),
                column: Some(logical_column),
                element_id: cell.element_id.clone(),
                path,
            },
        },
        visual,
    });
    table.placements.push(CellPlacement {
        object_index,
        row: logical_row,
        column: logical_column,
        row_span: cell.row_span,
        column_span: cell.column_span,
    });
    Ok(())
}

fn finish_table(
    table: TableState,
    objects: &mut [Object],
    frozen_panes: &HashMap<String, FreezePanes>,
) -> Result<Unit, Diagnostic> {
    let freeze = frozen_panes.get(&table.name).copied().unwrap_or_default();
    let rows = table.used_rows.max(freeze.rows).max(1);
    let columns = table.used_columns.max(freeze.columns).max(1);
    let column_axis = AxisLayout::new(
        COLUMN_WIDTH,
        table
            .column_sizes
            .into_iter()
            .map(|(column, size)| SizeSpan {
                start: column,
                end: column,
                size,
            })
            .collect(),
    );
    let row_axis = AxisLayout::new(ROW_HEIGHT, table.row_sizes);
    for placement in table.placements {
        let object = objects.get_mut(placement.object_index).ok_or_else(|| {
            format_error(CONTENT_PART, "cell placement references a missing object")
        })?;
        object.bounds = Rect {
            x: column_axis.offset(placement.column),
            y: row_axis.offset(placement.row),
            width: column_axis.span(
                placement.column,
                placement.column + placement.column_span - 1,
            ),
            height: row_axis.span(placement.row, placement.row + placement.row_span - 1),
        };
    }
    Ok(Unit {
        kind: UnitKind::Sheet,
        index: table.unit_index,
        id: format!("unit:{}", table.unit_index),
        name: table.name,
        width: column_axis.offset(columns).max(1.0),
        height: row_axis.offset(rows).max(1.0),
        rows,
        columns,
        frozen_rows: freeze.rows,
        frozen_columns: freeze.columns,
        frozen_width: column_axis.offset(freeze.columns),
        frozen_height: row_axis.offset(freeze.rows),
        row_axis: row_axis.descriptor(MAX_ROWS),
        column_axis: column_axis.descriptor(MAX_COLUMNS),
        show_grid_lines: true,
        tab_color: table.tab_color,
        sheet: None,
        slide: None,
    })
}

fn ranges_overlap(left: CellRange, right: CellRange) -> bool {
    left.start_row <= right.end_row
        && right.start_row <= left.end_row
        && left.start_column <= right.end_column
        && right.start_column <= left.end_column
}

fn repeat_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    maximum: u32,
    part: &str,
) -> Result<u32, Diagnostic> {
    let Some(value) = optional_attribute(attributes, name, part)? else {
        return Ok(1);
    };
    if value.starts_with('0') || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(dimension_error(
            part,
            format!("attribute {name} must be a canonical positive integer"),
        ));
    }
    value
        .parse::<u32>()
        .ok()
        .filter(|repeat| *repeat != 0 && *repeat <= maximum)
        .ok_or_else(|| dimension_error(part, format!("attribute {name} exceeds its limit")))
}

fn required_cell_value<'a>(
    value: Option<&'a str>,
    name: &str,
    source_index: u32,
) -> Result<&'a str, Diagnostic> {
    value.filter(|value| !value.is_empty()).ok_or_else(|| {
        format_error(
            CONTENT_PART,
            format!("source cell {} is missing office:{name}", source_index + 1),
        )
    })
}

fn append_text(target: &mut String, text: &str, limits: Limits) -> Result<(), Diagnostic> {
    if target
        .len()
        .checked_add(text.len())
        .is_none_or(|length| length > limits.max_xml_bytes)
    {
        return Err(Diagnostic::fatal(
            DiagnosticCode::XmlSizeLimit,
            Phase::Xml,
            None,
            "decoded cell text exceeds the configured XML byte limit",
        )
        .in_part(CONTENT_PART));
    }
    target.push_str(text);
    Ok(())
}

fn append_spaces(target: &mut String, count: u32, limits: Limits) -> Result<(), Diagnostic> {
    let count = usize::try_from(count)
        .map_err(|_| dimension_error(CONTENT_PART, "space repetition exceeds supported range"))?;
    if target
        .len()
        .checked_add(count)
        .is_none_or(|length| length > limits.max_xml_bytes)
    {
        return Err(Diagnostic::fatal(
            DiagnosticCode::XmlSizeLimit,
            Phase::Xml,
            None,
            "expanded cell text exceeds the configured XML byte limit",
        )
        .in_part(CONTENT_PART));
    }
    target.extend(std::iter::repeat_n(' ', count));
    Ok(())
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
    use super::{
        OdsCellStyle, OdsChartKind, apply_border, parse_content_xml, parse_ods_charts_in_xml,
        parse_ods_content_validations, parse_ods_icon_set_type,
    };
    use crate::diagnostic::DiagnosticCode;
    use crate::format::odf_chart::{OdfDataLabelNumber, OdfLabelPosition, OdfLegendPosition};
    use crate::limits::Limits;
    use crate::model::{Geometry, MappingQuality, ObjectKind, SourceLocator, UnitKind, Visual};

    #[test]
    fn supplied_date_styles_are_right_aligned() {
        let bytes = include_bytes!("../../tests/fixtures/oasis-1832-date-style.ods");
        let package = crate::package::Package::open(bytes, Limits::default()).unwrap();
        let document = super::parse(&package).unwrap();
        let cells = document
            .objects
            .iter()
            .filter(|cell| cell.text.is_some())
            .collect::<Vec<_>>();
        assert_eq!(cells.len(), 3);
        for (cell, expected) in cells
            .iter()
            .zip(["2020-05-23", "23.05. (Mai) 2020", "Sa 23.05.20"])
        {
            assert_eq!(cell.text.as_deref(), Some(expected));
            assert!(
                matches!(
                    &cell.visual,
                    Visual::Text {
                        align: crate::model::TextAlign::End,
                        ..
                    }
                ),
                "date must be right aligned: {:?}",
                cell.visual
            );
        }
    }

    #[test]
    fn supplied_engineering_numbers_are_right_aligned() {
        let bytes = include_bytes!("../../tests/fixtures/oasis-1828-3860-engineering.ods");
        let package = crate::package::Package::open(bytes, Limits::default()).unwrap();
        let document = super::parse(&package).unwrap();
        let cells = document
            .objects
            .iter()
            .filter(|object| object.text.is_some())
            .collect::<Vec<_>>();
        assert_eq!(cells.len(), 5);
        for (cell, expected) in cells.iter().zip([
            "250000000000.00",
            "2,50E+11",
            "250,0E+9",
            "2,500E11",
            "25,000E+10",
        ]) {
            assert_eq!(cell.text.as_deref(), Some(expected));
            assert!(
                matches!(
                    &cell.visual,
                    Visual::Text {
                        align: crate::model::TextAlign::End,
                        ..
                    }
                ),
                "{} must be right aligned: {:?}",
                expected,
                cell.visual
            );
        }
        let xml =
            String::from_utf8(package.required_part("content.xml").unwrap().to_vec()).unwrap();
        for (properties, expected) in [
            (
                "<style:paragraph-properties fo:text-align=\"start\"/>",
                crate::model::TextAlign::Start,
            ),
            (
                "<style:paragraph-properties fo:text-align=\"center\"/>",
                crate::model::TextAlign::Center,
            ),
        ] {
            let mut styles = super::OdsStyles::default();
            let parent = format!(
                r#"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo"><style:style style:name="Default" style:family="table-cell">{properties}</style:style></office:document-styles>"#
            );
            super::parse_styles_xml(
                parent.as_bytes(),
                Limits::default(),
                "styles.xml",
                &mut styles,
            )
            .unwrap();
            super::parse_styles_xml(
                xml.as_bytes(),
                Limits::default(),
                "content.xml",
                &mut styles,
            )
            .unwrap();
            let document = super::parse_content_xml_with_styles(
                xml.as_bytes(),
                Limits::default(),
                &styles,
                &Default::default(),
            )
            .unwrap();
            assert!(
                document
                    .objects
                    .iter()
                    .filter(|cell| cell.text.is_some())
                    .all(
                        |cell| matches!(&cell.visual, Visual::Text { align, .. } if *align == expected)
                    )
            );
        }
        let text_xml = xml.replace(
            "office:value-type=\"float\"",
            "office:value-type=\"string\"",
        );
        let mut styles = super::OdsStyles::default();
        super::parse_styles_xml(
            text_xml.as_bytes(),
            Limits::default(),
            "content.xml",
            &mut styles,
        )
        .unwrap();
        let document = super::parse_content_xml_with_styles(
            text_xml.as_bytes(),
            Limits::default(),
            &styles,
            &Default::default(),
        )
        .unwrap();
        assert!(
            document
                .objects
                .iter()
                .filter(|cell| cell.text.is_some())
                .all(|cell| matches!(
                    &cell.visual,
                    Visual::Text {
                        align: crate::model::TextAlign::Start,
                        ..
                    }
                ))
        );
    }

    #[test]
    fn parses_ods_scatter_area_and_radar_chart_classes() {
        let xml = br#"<office:document-content xmlns:office="office" xmlns:table="table" xmlns:chart="chart">
          <chart:chart chart:class="chart:scatter"><chart:plot-area><chart:series chart:values-cell-range-address="local.A1:local.A3"/></chart:plot-area><table:table table:name="local"><table:table-row><table:table-cell office:value="1"/></table:table-row><table:table-row><table:table-cell office:value="3"/></table:table-row><table:table-row><table:table-cell office:value="2"/></table:table-row></table:table></chart:chart>
          <chart:chart chart:class="chart:area"><chart:plot-area><chart:series chart:values-cell-range-address="local.A1:local.A3"/></chart:plot-area><table:table table:name="local"><table:table-row><table:table-cell office:value="1"/></table:table-row><table:table-row><table:table-cell office:value="3"/></table:table-row><table:table-row><table:table-cell office:value="2"/></table:table-row></table:table></chart:chart>
          <chart:chart chart:class="chart:radar"><chart:plot-area><chart:series chart:values-cell-range-address="local.A1:local.A3"/></chart:plot-area><table:table table:name="local"><table:table-row><table:table-cell office:value="1"/></table:table-row><table:table-row><table:table-cell office:value="3"/></table:table-row><table:table-row><table:table-cell office:value="2"/></table:table-row></table:table></chart:chart>
        </office:document-content>"#;
        let charts = parse_ods_charts_in_xml(xml, Limits::default(), "content.xml")
            .expect("charts")
            .into_iter()
            .map(|chart| chart.expect("chart has usable data"))
            .collect::<Vec<_>>();
        assert_eq!(charts.len(), 3);
        assert_eq!(charts[0].kind, OdsChartKind::Scatter);
        assert_eq!(charts[1].kind, OdsChartKind::Area);
        assert_eq!(charts[2].kind, OdsChartKind::Radar);
    }

    #[test]
    fn supplied_step_interpolation_draws_every_interval() {
        use crate::model::{Paint, Rect};
        let package = crate::package::Package::open(
            include_bytes!("../../tests/fixtures/oasis-3662-interpolation-step.ods"),
            Limits::default(),
        )
        .unwrap();
        for (index, command_count) in [2, 3, 3, 4, 4].into_iter().enumerate() {
            let part = format!("Object {}/content.xml", index + 1);
            let charts = parse_ods_charts_in_xml(
                &package.required_part(&part).unwrap(),
                Limits::default(),
                &part,
            )
            .unwrap();
            let chart = charts[0].as_ref().unwrap();
            assert_eq!(chart.domains[0], [-3., -1., 0., 5., 6., 9., 11.]);
            assert_eq!(chart.series[0], [5., 1., 2., -3., 5., 1., 2.]);
            // LibreOffice's original-file export includes one major-step margin.
            assert_eq!((chart.x_axis.minimum, chart.x_axis.maximum), (-4., 12.));
            assert_eq!((chart.y_axis.minimum, chart.y_axis.maximum), (-4., 6.));
            let frame = super::OdsChartFrame {
                unit_index: 0,
                table_name: "Sheet1".into(),
                frame_index: 0,
                element_id: None,
                bounds: Rect {
                    x: 0.,
                    y: 0.,
                    width: 442.,
                    height: 214.,
                },
                inline_chart_index: None,
                object_href: None,
            };
            let mut objects = Vec::new();
            super::push_ods_chart(
                chart,
                &frame,
                &mut objects,
                Limits::default().max_document_objects,
            )
            .unwrap();
            let lines = objects
                .iter()
                .filter_map(|object| match &object.visual {
                    Visual::PaintedShape {
                        geometry: Geometry::Path { commands, .. },
                        stroke: Paint::Solid(0x0045_86ff),
                        ..
                    } => Some(commands),
                    _ => None,
                })
                .collect::<Vec<_>>();
            assert_eq!(lines.len(), 6, "{part}: all six intervals must connect");
            for commands in lines {
                assert_eq!(commands.len(), command_count, "{part}");
                if index > 0 {
                    let points = commands
                        .iter()
                        .map(|command| match command {
                            crate::model::PathCommand::MoveTo { x, y }
                            | crate::model::PathCommand::LineTo { x, y } => (*x, *y),
                            _ => panic!("step must be linear"),
                        })
                        .collect::<Vec<_>>();
                    assert!(
                        points
                            .windows(2)
                            .all(|p| p[0].0 == p[1].0 || p[0].1 == p[1].1)
                    );
                    let first = points[0];
                    let last = *points.last().unwrap();
                    let expected = match index {
                        1 => vec![first, (first.0, last.1), last],
                        2 => vec![first, (last.0, first.1), last],
                        3 => vec![
                            first,
                            (first.0, (first.1 + last.1) / 2.),
                            (last.0, (first.1 + last.1) / 2.),
                            last,
                        ],
                        _ => vec![
                            first,
                            ((first.0 + last.0) / 2., first.1),
                            ((first.0 + last.0) / 2., last.1),
                            last,
                        ],
                    };
                    assert_eq!(points, expected, "{part}: step orientation");
                }
            }
        }
    }

    #[test]
    fn supplied_coordinate_region_xy_uses_native_axis_ranges() {
        let package = crate::package::Package::open(
            include_bytes!("../../tests/fixtures/oasis-3928-xy.ods"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        for part in ["Object 1/content.xml", "Object 2/content.xml"] {
            let charts = parse_ods_charts_in_xml(
                &package.required_part(part).unwrap(),
                Limits::default(),
                part,
            )
            .unwrap();
            let chart = charts[0].as_ref().unwrap();
            assert_eq!(
                (
                    chart.x_axis.minimum,
                    chart.x_axis.maximum,
                    chart.x_axis.interval
                ),
                (1.0, 10.0, 1.0)
            );
            assert_eq!(
                (
                    chart.y_axis.minimum,
                    chart.y_axis.maximum,
                    chart.y_axis.interval
                ),
                (0.0, 9.0, 1.0)
            );
            assert!((chart.plot_area.unwrap().width - 7.248 * 96.0 / 2.54).abs() < 0.01);
            assert_eq!(chart.grid_axes, [false, true]);
            let plot = chart.plot_area.unwrap();
            let markers = document
                .objects
                .iter()
                .filter(|o| {
                    o.source.part == part
                        && matches!(
                            o.visual,
                            Visual::PaintedShape {
                                geometry: Geometry::Rectangle,
                                fill: crate::model::Paint::Solid(0x0045_86ff),
                                ..
                            }
                        )
                        && o.bounds.width == 7.0
                })
                .collect::<Vec<_>>();
            assert_eq!(markers.len(), 4);
            let group = document
                .objects
                .iter()
                .find(|o| o.numeric_id == markers[0].parent_numeric_id.unwrap())
                .unwrap();
            for (marker, (x, y)) in
                markers
                    .iter()
                    .zip([(2.0, 5.0), (3.0, 4.0), (5.0, 8.0), (9.0, 3.0)])
            {
                assert!(
                    (marker.bounds.x + 3.5
                        - group.bounds.x
                        - plot.x
                        - (x - 1.0) / 9.0 * plot.width)
                        .abs()
                        < 0.01
                );
                assert!(
                    (marker.bounds.y + 3.5
                        - group.bounds.y
                        - plot.y
                        - (1.0 - y / 9.0) * plot.height)
                        .abs()
                        < 0.01
                );
            }
            assert!(
                !document.objects.iter().any(|o| o.source.part == part
                    && o.bounds.width == 0.5
                    && o.bounds.height == plot.height),
                "no x-axis grid was authored"
            );
        }
        let fonts = document
            .objects
            .iter()
            .filter_map(|o| match &o.visual {
                Visual::RichText { runs, .. } if o.text.as_deref() == Some("10") => {
                    Some(runs[0].font_size)
                }
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(fonts.len(), 2);
        assert!((fonts[0] - 40.0 / 3.0).abs() < 0.01);
        assert_eq!(fonts[1], 20.0);
    }

    #[test]
    fn supplied_coordinate_region_xy_3d_renders_ribbons() {
        let package = crate::package::Package::open(
            include_bytes!("../../tests/fixtures/oasis-3928-xy-3d.ods"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let ribbons = document.objects.iter().filter(|object| matches!(
                &object.visual,
                Visual::PaintedShape { geometry: Geometry::Path { commands, .. }, fill: crate::model::Paint::Solid(color), .. }
                    if commands.len() == 5 && (*color == 0x0045_86ff || *color == 0xff42_0eff)
            )).count();
        assert_eq!(
            ribbons, 12,
            "two charts, two series, three filled ribbon segments each"
        );
        for part in ["Object 1/content.xml", "Object 2/content.xml"] {
            let charts = parse_ods_charts_in_xml(
                &package.required_part(part).unwrap(),
                Limits::default(),
                part,
            )
            .unwrap();
            let chart = charts[0].as_ref().unwrap();
            assert_eq!(
                chart.series,
                [vec![5.0, 4.0, 8.0, 3.0], vec![1.0, 9.0, 7.0, 10.0]]
            );
            assert_eq!(chart.domains[0], [2.0, 3.0, 5.0, 9.0]);
            assert!(super::ods_chart_has_3d_ribbons(chart));
            assert_eq!(chart.series_colors, [0x0045_86ff, 0xff42_0eff]);
            let view = chart.projection.as_ref().unwrap();
            let plot = chart.plot_area.unwrap();
            let (a, b, c) = (
                view.project([0.0, 0.0, 0.5], plot),
                view.project([1.0 / 7.0, 0.0, 0.5], plot),
                view.project([3.0 / 7.0, 0.0, 0.5], plot),
            );
            assert!(
                ((c.0 - b.0) / (b.0 - a.0) - 2.0).abs() < 0.001,
                "numeric x spacing must survive projection"
            );
            assert!(
                (chart.axis_fonts[1]
                    - if part.contains("1/") {
                        40.0 / 3.0
                    } else {
                        20.0
                    })
                .abs()
                    < 0.01
            );
            let xml = String::from_utf8(package.required_part(part).unwrap().to_vec()).unwrap();
            let flat = parse_ods_charts_in_xml(
                xml.replace(
                    "chart:three-dimensional=\"true\"",
                    "chart:three-dimensional=\"false\"",
                )
                .as_bytes(),
                Limits::default(),
                part,
            )
            .unwrap();
            assert!(!super::ods_chart_has_3d_ribbons(flat[0].as_ref().unwrap()));
            let unsupported = parse_ods_charts_in_xml(
                xml.replace(
                    "dr3d:projection=\"parallel\"",
                    "dr3d:projection=\"perspective\"",
                )
                .as_bytes(),
                Limits::default(),
                part,
            )
            .unwrap();
            assert!(!super::ods_chart_has_3d_ribbons(
                unsupported[0].as_ref().unwrap()
            ));
        }
    }

    #[test]
    fn supplied_extrapolate_scatter_keeps_axes_points_and_regression_lines() {
        let bytes = include_bytes!("../../tests/fixtures/oasis-1148-extrapolate.ods");
        let package = crate::package::Package::open(bytes, Limits::default()).expect("ods package");
        let document = crate::format::ods::parse(&package).expect("Extrapolate.ods opens as ODS");
        let charts = parse_ods_charts_in_xml(
            &package.required_part("Object 1/content.xml").unwrap(),
            Limits::default(),
            "Object 1/content.xml",
        )
        .expect("chart part")
        .into_iter()
        .flatten()
        .collect::<Vec<_>>();
        assert_eq!(charts.len(), 1);
        let chart = &charts[0];
        assert_eq!(chart.kind, OdsChartKind::Scatter);
        assert_eq!(chart.title.as_deref(), Some("Extrapolate"));
        assert!((chart.x_axis.minimum - (-8.0)).abs() < 0.01);
        assert!((chart.x_axis.maximum - 12.0).abs() < 0.01);
        assert!((chart.x_axis.interval - 2.0).abs() < 0.01);
        assert!((chart.y_axis.minimum - (-4.0)).abs() < 0.01);
        assert!((chart.y_axis.maximum - 16.0).abs() < 0.01);
        assert!(chart.grid_axes[1]);
        assert_eq!(chart.domains.first().map(Vec::len), Some(6));
        assert_eq!(chart.series.first().map(Vec::len), Some(6));
        assert_eq!(chart.domains[0], [-2.0, 0.0, 1.0, 3.0, 4.0, 7.0]);
        assert_eq!(chart.series[0], [1.0, 2.0, 3.0, 4.0, 5.0, 10.0]);
        assert_eq!(chart.regressions.len(), 2);
        assert_eq!(chart.regressions[0].kind, "linear");
        assert_eq!(chart.regressions[1].kind, "exponential");
        assert_eq!(chart.series_style.symbol_name.as_deref(), Some("square"));
        assert!(chart.series_style.stroke_none);
        assert_eq!(chart.series_style.fill, 0x0000_ffff);

        let frame = super::OdsChartFrame {
            unit_index: 0,
            table_name: "Sheet1".to_owned(),
            frame_index: 0,
            element_id: None,
            bounds: crate::model::Rect {
                x: 0.0,
                y: 0.0,
                width: 600.0,
                height: 400.0,
            },
            inline_chart_index: None,
            object_href: None,
        };
        let mut objects = Vec::new();
        super::push_ods_chart(
            chart,
            &frame,
            &mut objects,
            Limits::default().max_document_objects,
        )
        .expect("chart renders");

        assert!(
            !objects.iter().any(|object| matches!(
                &object.visual,
                Visual::PaintedShape {
                    geometry: Geometry::Path { .. },
                    ..
                }
            )),
            "draw:stroke=none must keep the scatter series unconnected"
        );

        let rectangles = objects
            .iter()
            .filter(|object| {
                matches!(
                    &object.visual,
                    Visual::PaintedShape {
                        geometry: Geometry::Rectangle,
                        ..
                    }
                )
            })
            .count();
        assert!(
            rectangles >= 12,
            "axes, ticks, grid and square points: {rectangles}"
        );
        let curves = objects
            .iter()
            .filter_map(|object| {
                if let Visual::Effect {
                    clip: Some(Geometry::Rectangle),
                    visual,
                    ..
                } = &object.visual
                    && let Visual::PaintedShape {
                        geometry: Geometry::Path { commands, .. },
                        ..
                    } = visual.as_ref()
                {
                    Some((object.bounds, commands))
                } else {
                    None
                }
            })
            .collect::<Vec<_>>();
        assert_eq!(curves.len(), 2);
        for ((plot, commands), (start_y, end_y)) in curves
            .into_iter()
            .zip([(0.16393443, 8.809836), (1.1747601, 10.650358)])
        {
            let crate::model::PathCommand::MoveTo { x, y } = commands.first().unwrap() else {
                panic!("curve start")
            };
            assert!(
                (*x / plot.width - 0.3).abs() < 0.001,
                "curve starts at x=-2"
            );
            assert!((16.0 - *y / plot.height * 20.0 - start_y).abs() < 0.001);
            let crate::model::PathCommand::LineTo { x, y } = commands.last().unwrap() else {
                panic!("curve end")
            };
            assert!((*x / plot.width - 0.75).abs() < 0.001, "curve ends at x=7");
            assert!((16.0 - *y / plot.height * 20.0 - end_y).abs() < 0.001);
        }
        for label in ["y", "linear", "exponential"] {
            assert!(
                objects
                    .iter()
                    .any(|object| object.text.as_deref() == Some(label)),
                "legend: {label}"
            );
        }
        assert_eq!(
            objects
                .iter()
                .filter(|object| matches!(
                    &object.visual,
                    Visual::PaintedShape {
                        geometry: Geometry::Rectangle,
                        fill: crate::model::Paint::Solid(0x0000_ffff),
                        ..
                    }
                ))
                .count(),
            7,
            "six blue points and the blue legend key"
        );
        assert!(
            objects
                .iter()
                .any(|object| object.text.as_deref() == Some("Extrapolate")),
            "chart title is rendered"
        );
        assert!(
            objects
                .iter()
                .any(|object| object.text.as_deref() == Some("-8"))
                && objects
                    .iter()
                    .any(|object| object.text.as_deref() == Some("12"))
                && objects
                    .iter()
                    .any(|object| object.text.as_deref() == Some("-4"))
                && objects
                    .iter()
                    .any(|object| object.text.as_deref() == Some("16")),
            "axis min/max ticks are labeled"
        );
        let _ = document;
    }

    #[test]
    fn renders_ods_icon_set_markers_without_enforcing_rules() {
        assert_eq!(
            parse_ods_icon_set_type("iconset(3arrows;0;33;33;67;67;100)").as_deref(),
            Some("3arrows")
        );
        assert_eq!(
            parse_ods_icon_set_type("icon-set(3trafficlights)").as_deref(),
            Some("3trafficlights")
        );
        assert_eq!(parse_ods_icon_set_type("formula-is(1)"), None);

        let xml = br#"<office:document-content xmlns:office="office" xmlns:table="table" xmlns:text="text" xmlns:calcext="calcext">
          <office:body><office:spreadsheet>
            <table:table table:name="Sheet1"><table:table-row>
              <table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>
              <table:table-cell office:value-type="float" office:value="5"><text:p>5</text:p></table:table-cell>
              <table:table-cell office:value-type="float" office:value="9"><text:p>9</text:p></table:table-cell>
            </table:table-row></table:table>
            <calcext:conditional-formats>
              <calcext:conditional-format calcext:target-range-address="Sheet1.A1:Sheet1.C1">
                <calcext:condition calcext:apply-style-name="icons" calcext:value="iconset(3arrows)"/>
              </calcext:conditional-format>
            </calcext:conditional-formats>
          </office:spreadsheet></office:body></office:document-content>"#;
        let parsed = parse_content_xml(xml, Limits::default()).expect("content");
        let icons: Vec<_> = parsed
            .objects
            .iter()
            .filter(|object| {
                matches!(
                    object.source.locator,
                    SourceLocator::Ods {
                        ref path,
                        ..
                    } if path.contains("calcext:icon-set")
                )
            })
            .collect();
        assert_eq!(
            icons.len(),
            3,
            "{:?}",
            parsed.objects.iter().map(|o| &o.source)
        );
        assert!(icons.iter().any(|object| {
            matches!(
                &object.visual,
                Visual::Shape {
                    geometry: Geometry::Path { .. },
                    fill: 0x00c8_53ff,
                    ..
                }
            )
        }));
        assert!(
            !parsed
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("could not be rendered")),
            "{:?}",
            parsed.diagnostics
        );
    }

    #[test]
    fn ods_content_validations_are_read_only_annotations_with_prompt_error_metadata() {
        let xml = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
          xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
          <table:content-validations>
            <table:content-validation table:name="v1" table:condition="cell-content()&lt;=10"
              table:show-error-message="true" table:show-input-message="true">
              <table:help-message table:title="Hint"><text:p>Enter 1-10</text:p></table:help-message>
              <table:error-message table:title="Invalid"><text:p>Out of range</text:p></table:error-message>
            </table:content-validation>
          </table:content-validations>
          <office:body><office:spreadsheet>
            <table:table table:name="Sheet1"><table:table-row>
              <table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell>
            </table:table-row></table:table>
          </office:spreadsheet></office:body>
        </office:document-content>"#;
        let (validations, active) =
            parse_ods_content_validations(xml, Limits::default()).expect("validation parse");
        assert!(!active);
        assert_eq!(validations.len(), 1);
        let validation = &validations[0];
        assert_eq!(validation.name.as_deref(), Some("v1"));
        assert_eq!(validation.condition.as_deref(), Some("cell-content()<=10"));
        assert_eq!(validation.error_title.as_deref(), Some("Invalid"));
        assert_eq!(validation.error_text, "Out of range");
        assert_eq!(validation.help_title.as_deref(), Some("Hint"));
        assert_eq!(validation.help_text, "Enter 1-10");

        let parsed = parse_content_xml(xml, Limits::default()).expect("content parse");
        assert!(
            parsed.diagnostics.iter().any(|diagnostic| {
                diagnostic.message.contains("read-only annotation")
                    && diagnostic.message.contains("condition=cell-content()<=10")
                    && diagnostic.message.contains("error=Invalid: Out of range")
                    && diagnostic.message.contains("help=Hint: Enter 1-10")
            }),
            "{:?}",
            parsed.diagnostics
        );
        assert_eq!(parsed.objects.len(), 1);
    }

    #[test]
    fn ods_content_validation_macros_are_blocked_without_failing_the_sheet() {
        let xml = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
          xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
          <table:content-validations>
            <table:content-validation table:name="v1" table:condition="cell-content()&gt;0">
              <table:validity-macros/>
            </table:content-validation>
          </table:content-validations>
          <office:body><office:spreadsheet>
            <table:table table:name="Sheet1"><table:table-row>
              <table:table-cell office:value-type="string" office:string-value="ok"><text:p>ok</text:p></table:table-cell>
            </table:table-row></table:table>
          </office:spreadsheet></office:body>
        </office:document-content>"#;
        let (validations, active) =
            parse_ods_content_validations(xml, Limits::default()).expect("validation parse");
        assert!(active);
        assert_eq!(validations.len(), 1);
        let parsed = parse_content_xml(xml, Limits::default()).expect("content parse");
        assert!(parsed.diagnostics.iter().any(|diagnostic| {
            diagnostic.code == DiagnosticCode::ActiveContentBlocked
                && diagnostic.message.contains("content-validation macros")
        }));
        assert_eq!(parsed.objects.len(), 1);
    }

    #[test]
    fn supplied_pie_legend_preserves_labels_colors_and_custom_bounds() {
        let package = crate::package::Package::open(
            include_bytes!("../../tests/fixtures/oasis-3883-chart-legend.ods"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let labels: Vec<_> = document
            .objects
            .iter()
            .filter(|object| object.source.part == "Object 1/content.xml" && object.text.is_some())
            .collect();
        assert_eq!(
            labels
                .iter()
                .map(|object| object.text.as_deref().unwrap())
                .collect::<Vec<_>>(),
            ["Anne", "Tim", "Bob", "John", "Eve", "Georg", "James"]
        );
        let colors = [
            0x0045_86ff,
            0xff42_0eff,
            0xffd3_20ff,
            0x579d_1cff,
            0x7e00_21ff,
            0x83ca_ffff,
            0xd3d3_d3ff,
        ];
        for (index, label) in labels.iter().enumerate() {
            let key = &document.objects[label.numeric_id as usize - 1];
            assert!(matches!(key.visual, crate::model::Visual::PaintedShape {
                fill: crate::model::Paint::Solid(color), ..
            } if color == colors[index]));
            assert_eq!(label.bounds.y, labels[index / 3 * 3].bounds.y);
            if index >= 3 {
                assert!(label.bounds.y > labels[index - 3].bounds.y);
            }
        }
        let slice_colors: Vec<_> = document
            .objects
            .iter()
            .filter_map(|object| match object.visual {
                crate::model::Visual::PaintedShape {
                    geometry: crate::model::Geometry::Path { .. },
                    fill: crate::model::Paint::Solid(color),
                    ..
                } if object.source.part == "Object 1/content.xml" => Some(color),
                _ => None,
            })
            .collect();
        assert_eq!(slice_colors, colors);
        let left = (5.176 + 9.509) * 96.0 / 2.54;
        let top = (1.312 + 3.455) * 96.0 / 2.54;
        for label in labels {
            assert!(
                label.bounds.x >= left - 0.1
                    && label.bounds.x + label.bounds.width <= left + 5.533 * 96.0 / 2.54 + 0.1
            );
            assert!(
                label.bounds.y >= top - 0.1
                    && label.bounds.y + label.bounds.height <= top + 1.629 * 96.0 / 2.54 + 0.1
            );
        }
    }

    #[test]
    fn renders_exploded_pie_slice_from_pie_offset() {
        let chart = br##"<office:document-content xmlns:office="office" xmlns:chart="chart" xmlns:style="style" xmlns:text="text" xmlns:table="table" xmlns:draw="draw" xmlns:svg="svg">
          <office:automatic-styles>
            <style:style style:name="ch10" style:family="chart">
              <style:chart-properties chart:pie-offset="29" chart:solid-type="cuboid"/>
              <style:graphic-properties draw:fill-color="#ffd320"/>
            </style:style>
            <style:style style:name="ch8" style:family="chart">
              <style:chart-properties chart:solid-type="cuboid"/>
              <style:graphic-properties draw:fill-color="#004586"/>
            </style:style>
            <style:style style:name="ch11" style:family="chart">
              <style:chart-properties chart:solid-type="cuboid"/>
              <style:graphic-properties draw:fill-color="#579d1c"/>
            </style:style>
          </office:automatic-styles>
          <office:body><office:chart>
            <chart:chart svg:width="10.63cm" svg:height="8.989cm" chart:class="chart:circle">
              <chart:plot-area>
                <chart:coordinate-region svg:x="0cm" svg:y="0cm" svg:width="8cm" svg:height="8cm"/>
                <chart:series chart:values-cell-range-address="Sheet1.A1:Sheet1.A3">
                  <chart:data-point chart:style-name="ch8"/>
                  <chart:data-point chart:style-name="ch10"/>
                  <chart:data-point chart:style-name="ch11"/>
                </chart:series>
              </chart:plot-area>
              <table:table table:name="local-table">
                <table:table-row><table:table-cell office:value="5"/></table:table-row>
                <table:table-row><table:table-cell office:value="8"/></table:table-row>
                <table:table-row><table:table-cell office:value="3"/></table:table-row>
              </table:table>
            </chart:chart>
          </office:chart></office:body>
        </office:document-content>"##;
        let parsed = parse_ods_charts_in_xml(chart, Limits::default(), "Object 1/content.xml")
            .expect("chart xml parses");
        let chart = parsed.into_iter().next().flatten().expect("pie chart");
        assert_eq!(chart.point_explosions, vec![0.0, 0.29, 0.0]);

        let mut objects = Vec::new();
        let frame = super::OdsChartFrame {
            unit_index: 0,
            table_name: "Sheet1".to_owned(),
            frame_index: 0,
            element_id: None,
            bounds: crate::model::Rect {
                x: 0.0,
                y: 0.0,
                width: 8.0 * 96.0 / 2.54,
                height: 8.0 * 96.0 / 2.54,
            },
            inline_chart_index: None,
            object_href: None,
        };
        super::push_ods_chart(
            &chart,
            &frame,
            &mut objects,
            Limits::default().max_document_objects,
        )
        .expect("chart renders");
        let slices: Vec<_> = objects
            .iter()
            .filter(|object| {
                matches!(
                    object.visual,
                    crate::model::Visual::PaintedShape {
                        geometry: crate::model::Geometry::Path { .. },
                        ..
                    }
                )
            })
            .collect();
        assert_eq!(slices.len(), 3);
        let pie_center_x = frame.bounds.x + frame.bounds.width / 2.0;
        let pie_center_y = frame.bounds.y + frame.bounds.height / 2.0;
        let apex = |object: &crate::model::Object| match &object.visual {
            crate::model::Visual::PaintedShape {
                geometry: crate::model::Geometry::Path { commands, .. },
                ..
            } => match commands.first() {
                Some(crate::model::PathCommand::MoveTo { x, y }) => {
                    (object.bounds.x + *x, object.bounds.y + *y)
                }
                _ => (
                    object.bounds.x + object.bounds.width / 2.0,
                    object.bounds.y + object.bounds.height / 2.0,
                ),
            },
            _ => (
                object.bounds.x + object.bounds.width / 2.0,
                object.bounds.y + object.bounds.height / 2.0,
            ),
        };
        let distance_to_center = |object: &crate::model::Object| {
            let (x, y) = apex(object);
            ((x - pie_center_x).powi(2) + (y - pie_center_y).powi(2)).sqrt()
        };
        let exploded = &slices[1];
        let joined_distance = distance_to_center(&slices[0]);
        let exploded_distance = distance_to_center(exploded);
        let radius = frame.bounds.width / 2.0 * 0.94;
        assert!(
            exploded_distance > 0.2 * radius,
            "exploded apex distance {exploded_distance} should reflect pie-offset 0.29 (radius={radius})"
        );
        assert!(
            exploded_distance > joined_distance + 1.0,
            "exploded apex moves away from the pie center (joined={joined_distance}, exploded={exploded_distance})"
        );
        assert!(
            distance_to_center(&slices[2]) < radius * 0.05,
            "non-offset slices keep their apex at the pie center"
        );
    }

    #[test]
    fn renders_pie_percentage_labels_and_bottom_legend() {
        let chart = br##"<office:document-content xmlns:office="office" xmlns:chart="chart" xmlns:style="style" xmlns:text="text" xmlns:table="table" xmlns:fo="fo" xmlns:draw="draw" xmlns:svg="svg">
          <office:automatic-styles>
            <style:style style:name="ch3" style:family="chart">
              <style:chart-properties chart:data-label-number="percentage" chart:data-label-text="false" chart:data-label-symbol="false"/>
            </style:style>
            <style:style style:name="ch7" style:family="chart">
              <style:chart-properties chart:data-label-number="percentage" chart:data-label-text="false" chart:data-label-symbol="false" chart:label-position="outside"/>
              <style:text-properties fo:font-size="15pt"/>
            </style:style>
            <style:style style:name="ch2" style:family="chart">
              <style:text-properties fo:font-size="10pt"/>
            </style:style>
          </office:automatic-styles>
          <office:body><office:chart>
            <chart:chart svg:width="10.63cm" svg:height="8.989cm" chart:class="chart:circle">
              <chart:legend chart:legend-position="bottom" svg:x="3.738cm" svg:y="8.384cm" style:legend-expansion="wide" chart:style-name="ch2"/>
              <chart:plot-area chart:style-name="ch3">
                <chart:coordinate-region svg:x="1.999cm" svg:y="0.999cm" svg:width="4.999cm" svg:height="4.999cm"/>
                <chart:axis chart:dimension="x">
                  <chart:categories table:cell-range-address="Sheet1.A2:Sheet1.A5"/>
                </chart:axis>
                <chart:series chart:style-name="ch7" chart:values-cell-range-address="Sheet1.B2:Sheet1.B5"/>
              </chart:plot-area>
              <table:table table:name="local-table">
                <table:table-row><table:table-cell/><table:table-cell><text:p>y</text:p></table:table-cell></table:table-row>
                <table:table-row><table:table-cell><text:p>A</text:p></table:table-cell><table:table-cell office:value="5"/></table:table-row>
                <table:table-row><table:table-cell><text:p>B</text:p></table:table-cell><table:table-cell office:value="4"/></table:table-row>
                <table:table-row><table:table-cell><text:p>C</text:p></table:table-cell><table:table-cell office:value="8"/></table:table-row>
                <table:table-row><table:table-cell><text:p>D</text:p></table:table-cell><table:table-cell office:value="3"/></table:table-row>
              </table:table>
            </chart:chart>
          </office:chart></office:body>
        </office:document-content>"##;
        let parsed = parse_ods_charts_in_xml(chart, Limits::default(), "Object 1/content.xml")
            .expect("chart xml parses");
        let chart = parsed
            .into_iter()
            .next()
            .flatten()
            .expect("pie chart is present");
        assert_eq!(chart.kind, OdsChartKind::Pie);
        let labels = chart.data_labels.as_ref().expect("data labels configured");
        assert!(labels.is_enabled());
        assert_eq!(labels.number, Some(OdfDataLabelNumber::Percentage));
        assert_eq!(labels.position, Some(OdfLabelPosition::Outside));
        assert!((labels.font_size - 15.0 * 96.0 / 72.0).abs() < 0.01);
        let legend = chart.legend.as_ref().expect("legend configured");
        assert_eq!(legend.position, Some(OdfLegendPosition::Bottom));
        assert!(legend.expansion_wide);
        assert!(legend.width.is_none() && legend.height.is_none());

        let mut objects = Vec::new();
        let frame = super::OdsChartFrame {
            unit_index: 0,
            table_name: "Sheet1".to_owned(),
            frame_index: 0,
            element_id: None,
            bounds: crate::model::Rect {
                x: 0.0,
                y: 0.0,
                width: 10.63 * 96.0 / 2.54,
                height: 8.989 * 96.0 / 2.54,
            },
            inline_chart_index: None,
            object_href: None,
        };
        super::push_ods_chart(
            &chart,
            &frame,
            &mut objects,
            Limits::default().max_document_objects,
        )
        .expect("chart renders");
        let texts: Vec<_> = objects
            .iter()
            .filter_map(|object| object.text.as_deref())
            .collect();
        for expected in ["25%", "20%", "40%", "15%", "A", "B", "C", "D"] {
            assert!(texts.contains(&expected), "missing {expected} in {texts:?}");
        }
        let legend_labels: Vec<_> = objects
            .iter()
            .filter(|object| {
                object.text.is_some()
                    && matches!(
                        &object.source.locator,
                        SourceLocator::Ods { path, .. } if path.ends_with("/chart:legend[1]")
                    )
            })
            .collect();
        assert_eq!(legend_labels.len(), 4);
        let top = legend_labels
            .iter()
            .map(|o| o.bounds.y)
            .fold(f32::INFINITY, f32::min);
        let pie_bottom = frame.bounds.y + 5.998 * 96.0 / 2.54;
        assert!(
            top >= pie_bottom,
            "legend should sit below the pie (y={top})"
        );
        assert!(
            top + 1.0 >= 8.384 * 96.0 / 2.54,
            "legend keeps the authored bottom y"
        );
        assert!(
            legend_labels
                .windows(2)
                .all(|pair| pair[0].bounds.y == pair[1].bounds.y),
            "bottom legend items share one row"
        );
    }

    #[test]
    fn supplied_chart_outside_cells_is_inside_the_sheet_axes() {
        let package = crate::package::Package::open(
            include_bytes!("../../tests/fixtures/oasis-3883-chart-legend.ods"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let unit = &document.units[0];
        let chart: Vec<_> = document
            .objects
            .iter()
            .filter(|object| object.source.part == "Object 1/content.xml")
            .collect();
        assert_eq!(
            chart
                .iter()
                .filter(|object| matches!(
                    object.visual,
                    crate::model::Visual::PaintedShape {
                        geometry: crate::model::Geometry::Path { .. },
                        ..
                    }
                ))
                .count(),
            7,
            "all seven data slices"
        );
        for object in chart {
            assert!(unit.column_axis.offset(unit.columns) >= object.bounds.x + object.bounds.width);
            assert!(unit.row_axis.offset(unit.rows) >= object.bounds.y + object.bounds.height);
        }
    }

    #[test]
    fn supplied_number_fill_character_preserves_all_three_positions() {
        let package = crate::package::Package::open(
            include_bytes!("../../tests/fixtures/oasis-3765-number-fill-character.ods"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let cells: Vec<_> = document
            .objects
            .iter()
            .filter(|cell| cell.kind == ObjectKind::Cell)
            .collect();
        assert_eq!(cells.len(), 12);
        for (index, cell) in cells.iter().enumerate() {
            let fill = match &cell.visual {
                crate::model::Visual::TextLayout { layout, .. } => layout.fill_character,
                _ => None,
            };
            let text = cell.text.as_deref().unwrap();
            let offset = match index / 3 {
                0 => {
                    assert_eq!(fill, None);
                    continue;
                }
                1 => 0,
                2 => text.encode_utf16().count() - 1,
                _ => text.encode_utf16().count(),
            };
            assert_eq!(fill, Some((offset as u32, '―')));
            assert!(
                !text.contains('―'),
                "decorative fill must not alter cell text"
            );
        }
    }

    #[test]
    fn number_fill_maps_follow_inherited_column_row_and_cell_styles() {
        let xml = br#"<office:document-content><office:automatic-styles>
          <number:currency-style style:name="positive"><number:number/><number:fill-character>_</number:fill-character><number:currency-symbol>$</number:currency-symbol></number:currency-style>
          <number:currency-style style:name="signed"><number:text>-</number:text><number:fill-character>.</number:fill-character><number:number/><number:currency-symbol>$</number:currency-symbol><style:map style:condition="value()&gt;=0" style:apply-style-name="positive"/></number:currency-style>
          <style:style style:family="table-cell" style:name="base" style:data-style-name="signed"/>
          <style:style style:family="table-cell" style:name="child" style:parent-style-name="base"/>
          <style:style style:family="table-cell" style:name="plain"/>
          </office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1">
          <table:table-column table:number-columns-repeated="2" table:default-cell-style-name="child"/>
          <table:table-row><table:table-cell office:value-type="currency" office:value="7"><text:p>7$</text:p></table:table-cell><table:table-cell office:value-type="currency" office:value="-7"><text:p>-7$</text:p></table:table-cell></table:table-row>
          <table:table-row table:default-cell-style-name="plain"><table:table-cell office:value-type="currency" office:value="7"><text:p>7$</text:p></table:table-cell><table:table-cell table:style-name="child" office:value-type="currency" office:value="7"><text:p>7$</text:p></table:table-cell></table:table-row>
          </table:table></office:spreadsheet></office:body></office:document-content>"#;
        let content = parse_content_xml(xml, Limits::default()).unwrap();
        let fills: Vec<_> = content
            .objects
            .iter()
            .map(|object| match &object.visual {
                crate::model::Visual::TextLayout { layout, .. } => layout.fill_character,
                _ => None,
            })
            .collect();
        assert_eq!(
            fills,
            [Some((1, '_')), Some((1, '.')), None, Some((1, '_'))]
        );
    }

    #[test]
    fn opens_supplied_unnamed_sheet_and_preserves_line_break() {
        let package = crate::package::Package::open(
            include_bytes!("../../tests/fixtures/libreoffice-tdf75702-text-line-break.ods"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).expect("missing sheet name must be recoverable");
        assert_eq!(document.units.len(), 1);
        assert_eq!(document.units[0].name, "Sheet1");
        assert_eq!(document.objects[0].text.as_deref(), Some("line1\nline2"));
        assert!(
            document.objects[0].bounds.height >= 43.0,
            "both text lines must fit"
        );
        assert!(document.diagnostics.iter().any(|diagnostic| diagnostic.code
            == DiagnosticCode::FormatInvalid
            && diagnostic.fidelity == crate::diagnostic::Fidelity::Approximate));
    }

    #[test]
    fn supplied_diverse_fraction_preserves_underscores_and_fits_text() {
        let package = crate::package::Package::open(
            include_bytes!("../../tests/fixtures/oasis-3695-diverse-fraction.ods"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let cells: Vec<_> = document
            .objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Cell)
            .collect();
        assert_eq!(cells.len(), 16);
        for (index, expected) in [
            (10, "3_1/7"),
            (11, "1_3/4"),
            (14, "03__140914/995207"),
            (15, "01__3/4     "),
        ] {
            assert_eq!(cells[index].text.as_deref(), Some(expected));
            assert!(
                cells[index].bounds.height >= 25.6 - 0.001,
                "fraction underscores must fit inside the optimal row: {:?}",
                cells[index].bounds
            );
        }
    }

    #[test]
    fn supplied_extrapolate_optimal_rows_fit_single_line_text() {
        let package = crate::package::Package::open(
            include_bytes!("../../tests/fixtures/oasis-1148-extrapolate.ods"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let cells: Vec<_> = document
            .objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Cell)
            .collect();
        assert_eq!(cells.len(), 14);
        for cell in cells {
            assert!(
                cell.bounds.height >= 25.6 - 0.001,
                "{}: {:?}",
                cell.text.as_deref().unwrap(),
                cell.bounds
            );
        }
    }

    #[test]
    fn multiline_cells_preserve_explicit_and_hidden_row_heights() {
        let content = parse_content_xml(br#"<office:document-content>
          <office:automatic-styles><style:style style:family="table-row" style:name="fixed">
          <style:table-row-properties style:row-height="10px"/></style:style></office:automatic-styles>
          <office:body><office:spreadsheet><table:table table:name="Sheet1">
          <table:table-row table:style-name="fixed"><table:table-cell><text:p>a<text:line-break/>b</text:p></table:table-cell></table:table-row>
          <table:table-row table:visibility="collapse"><table:table-cell><text:p>a<text:line-break/>b</text:p></table:table-cell></table:table-row>
          </table:table></office:spreadsheet></office:body></office:document-content>"#, Limits::default()).unwrap();
        assert_eq!(content.objects.len(), 1);
        assert_eq!(content.objects[0].bounds.height, 10.0);
        assert_eq!(content.units[0].height, 10.0);
    }

    #[test]
    fn optimal_row_height_preserves_fixed_hidden_and_taller_rows() {
        for (optimal, height, visibility, expected) in [
            ("true", "10px", "visible", 25.6),
            ("1", "10px", "visible", 25.6),
            ("false", "10px", "visible", 10.0),
            ("0", "10px", "visible", 10.0),
            ("true", "40px", "visible", 40.0),
            ("true", "10px", "collapse", 0.0),
        ] {
            let xml = format!(
                r#"<office:document-content><office:automatic-styles>
              <style:style style:family="table-row" style:name="row">
              <style:table-row-properties style:row-height="{height}" style:use-optimal-row-height="{optimal}"/>
              </style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1">
              <table:table-row table:style-name="row" table:visibility="{visibility}" table:number-rows-repeated="2">
              <table:table-cell><text:p>x</text:p></table:table-cell></table:table-row>
              </table:table></office:spreadsheet></office:body></office:document-content>"#
            );
            let content = parse_content_xml(xml.as_bytes(), Limits::default()).unwrap();
            if expected == 0.0 {
                assert!(content.objects.is_empty());
            } else {
                assert_eq!(content.objects.len(), 2);
                assert!((content.objects[0].bounds.height - expected).abs() < 0.001);
                assert!((content.objects[1].bounds.y - expected).abs() < 0.001);
            }
        }
    }

    #[test]
    fn unnamed_flat_sheets_do_not_take_later_authored_names() {
        let document = super::parse_flat(br#"<office:document><office:body><office:spreadsheet>
          <table:table><table:table-row><table:table-cell><text:p>a</text:p></table:table-cell></table:table-row></table:table>
          <table:table table:name="Sheet1"/>
          <table:table table:name=""/>
          <table:table table:name="Sheet3"/>
          </office:spreadsheet></office:body></office:document>"#, Limits::default()).unwrap();
        assert_eq!(
            document
                .units
                .iter()
                .map(|unit| unit.name.as_str())
                .collect::<Vec<_>>(),
            ["Sheet2", "Sheet1", "Sheet4", "Sheet3"]
        );
        assert_eq!(document.objects[0].text.as_deref(), Some("a"));
    }

    #[test]
    fn opens_toolkit_sheet_with_bounded_repetition_and_preserves_later_rows() {
        let package = crate::package::Package::open(
            include_bytes!("../../tests/fixtures/toolkit-basic-sheet.ods"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).expect("oversized repetition must remain local");
        assert_eq!(document.units.len(), 3);
        assert_eq!(document.units[0].rows, 1_048_576);
        assert!(document.objects.len() < 100);
        assert!(
            document
                .objects
                .iter()
                .any(|object| object.text.as_deref() == Some("singleX"))
        );
        assert_eq!(
            document
                .objects
                .iter()
                .filter(|object| object.text.as_deref() == Some("blub"))
                .count(),
            4
        );
        assert_eq!(
            document
                .objects
                .iter()
                .filter(|object| object.text.as_deref() == Some("blubg"))
                .count(),
            4
        );
        assert!(document.objects.iter().any(|object| matches!(
            object.source.locator,
            SourceLocator::Ods {
                row: Some(1_048_575),
                column: Some(0),
                ..
            }
        )));
        assert!(
            document
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.code == DiagnosticCode::ObjectLimit
                    && diagnostic.fidelity == crate::diagnostic::Fidelity::Omitted)
        );
    }

    #[test]
    fn accepts_standard_odf_border_width_keywords() {
        for (value, expected) in [("thin", 1.0), ("medium", 3.0), ("thick", 5.0)] {
            let mut style = OdsCellStyle::default();
            apply_border(&format!("{value} solid #123456"), &mut style, "styles.xml")
                .expect("ODF border keyword");
            assert_eq!(style.stroke_width, expected);
            assert_eq!(style.stroke, 0x1234_56ff);
        }
        let mut style = OdsCellStyle::default();
        apply_border("1px dash-dot #123456", &mut style, "styles.xml").unwrap();
        assert_eq!(style.stroke_style.dash, [8.0, 4.0, 2.0, 4.0]);
    }

    #[test]
    fn parses_ods_cell_types_and_source_coordinates() {
        let content = parse_content_xml(
            br#"<office:document-content xmlns:office="urn:office" xmlns:table="urn:table" xmlns:text="urn:text">
                <office:body><office:spreadsheet><table:table table:name="Sheet1">
                  <table:table-row>
                    <table:table-cell office:value-type="string"><text:p>Hello<text:s text:c="2"/>ODS</text:p></table:table-cell>
                    <table:table-cell office:value-type="float" office:value="42.5"/>
                    <table:table-cell office:value-type="date" office:date-value="2026-07-15"/>
                    <table:table-cell office:value-type="boolean" office:boolean-value="1"/>
                  </table:table-row>
                </table:table></office:spreadsheet></office:body>
              </office:document-content>"#,
            Limits::default(),
        )
        .expect("valid ODS content");

        assert_eq!(content.units.len(), 1);
        assert_eq!(content.units[0].kind, UnitKind::Sheet);
        assert_eq!(content.units[0].name, "Sheet1");
        assert_eq!((content.units[0].rows, content.units[0].columns), (1, 4));
        assert_eq!(content.objects.len(), 4);
        assert!(
            content
                .objects
                .iter()
                .all(|object| object.kind == ObjectKind::Cell)
        );
        assert_eq!(content.objects[0].text.as_deref(), Some("Hello  ODS"));
        assert_eq!(content.objects[1].text.as_deref(), Some("42.5"));
        assert_eq!(content.objects[2].text.as_deref(), Some("2026-07-15"));
        assert_eq!(content.objects[3].text.as_deref(), Some("true"));
        assert_eq!(content.objects[0].bounds.x, 0.0);
        assert_eq!(content.objects[3].bounds.x, 288.0);
        assert_eq!(content.objects[0].source.mapping, MappingQuality::Exact);
        assert!(matches!(
            &content.objects[3].source.locator,
            SourceLocator::Ods {
                table_name,
                row: Some(0),
                column: Some(3),
                ..
            } if table_name == "Sheet1"
        ));
    }

    #[test]
    fn expands_repeated_material_cells_with_derived_mapping() {
        let content = parse_content_xml(
            br#"<office:document-content><office:body><office:spreadsheet>
                <table:table table:name="Repeated"><table:table-row table:number-rows-repeated="2">
                  <table:table-cell table:number-columns-repeated="2" office:value-type="string" office:string-value="x"/>
                </table:table-row></table:table>
              </office:spreadsheet></office:body></office:document-content>"#,
            Limits::default(),
        )
        .expect("bounded repetition");

        assert_eq!(content.objects.len(), 4);
        assert_eq!((content.units[0].rows, content.units[0].columns), (2, 2));
        assert!(
            content
                .objects
                .iter()
                .all(|object| object.source.mapping == MappingQuality::Derived)
        );
    }

    #[test]
    fn ignores_repeated_empty_paragraph_cells_at_the_sheet_tail() {
        let content = parse_content_xml(
            br#"<office:document-content><office:body><office:spreadsheet>
                <table:table table:name="Sheet1">
                  <table:table-row><table:table-cell office:value-type="string" office:string-value="x"/></table:table-row>
                  <table:table-row table:number-rows-repeated="1048575">
                    <table:table-cell table:number-columns-repeated="4"><text:p/></table:table-cell>
                  </table:table-row>
                </table:table>
              </office:spreadsheet></office:body></office:document-content>"#,
            Limits::default(),
        )
        .expect("a repeated empty tail must remain sparse");

        assert_eq!(content.objects.len(), 1);
        assert_eq!((content.units[0].rows, content.units[0].columns), (1, 1));
    }

    #[test]
    fn rejects_repetition_before_it_can_expand_objects() {
        let limits = Limits {
            max_document_objects: 4,
            ..Limits::default()
        };
        let error = parse_content_xml(
            br#"<office:document-content><office:body><office:spreadsheet>
                <table:table table:name="Bomb"><table:table-row>
                  <table:table-cell table:number-columns-repeated="5" office:value-type="string" office:string-value="x"/>
                </table:table-row></table:table>
              </office:spreadsheet></office:body></office:document-content>"#,
            limits,
        )
        .expect_err("object expansion must be bounded");
        assert_eq!(error.code, DiagnosticCode::ObjectLimit);

        let error = parse_content_xml(
            br#"<office:document-content><office:body><office:spreadsheet>
                <table:table table:name="Bomb"><table:table-row table:number-rows-repeated="1048577"/>
                </table:table></office:spreadsheet></office:body></office:document-content>"#,
            Limits::default(),
        )
        .expect_err("row dimensions must be bounded");
        assert_eq!(error.code, DiagnosticCode::LayoutBudgetExceeded);
    }

    #[test]
    fn diagnoses_remaining_unsupported_sheet_content_once() {
        let content = parse_content_xml(
            br#"<office:document-content xmlns:office="office" xmlns:table="table" xmlns:text="text" xmlns:draw="draw" xmlns:chart="chart" xmlns:calcext="calcext">
              <office:body><office:spreadsheet><table:table table:name="Sheet1">
                <calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="Sheet1.A1">
                  <calcext:condition calcext:apply-style-name="missing" calcext:value="formula-is(1)" calcext:base-cell-address="Sheet1.A1"/>
                </calcext:conditional-format></calcext:conditional-formats>
                <table:table-row><table:table-cell table:number-columns-spanned="2">
                  <text:p>value</text:p><draw:frame><draw:image/></draw:frame><draw:object/><draw:rect/><chart:chart/>
                </table:table-cell><table:covered-table-cell/></table:table-row>
              </table:table></office:spreadsheet></office:body>
            </office:document-content>"#,
            Limits::default(),
        )
        .expect("unsupported sheet content is diagnosed non-fatally");

        let unsupported = content
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.code == DiagnosticCode::UnsupportedFeature)
            .collect::<Vec<_>>();
        assert_eq!(unsupported.len(), 2);
        for needle in ["drawing", "conditional"] {
            assert_eq!(
                unsupported
                    .iter()
                    .filter(|diagnostic| diagnostic.message.contains(needle))
                    .count(),
                1,
                "{needle} diagnostic must be deduplicated"
            );
        }
    }

    #[test]
    fn parses_bar_line_and_pie_chart_series_from_inline_tables() {
        let bytes = br#"<office:document-content xmlns:office="office" xmlns:table="table" xmlns:chart="chart">
          <chart:chart chart:class="chart:bar"><chart:plot-area><chart:series chart:values-cell-range-address="local.A1:local.A2"/></chart:plot-area><table:table table:name="local"><table:table-row><table:table-cell office:value="10"/></table:table-row><table:table-row><table:table-cell office:value="20"/></table:table-row></table:table></chart:chart>
          <chart:chart chart:class="chart:line"><chart:plot-area><chart:series chart:values-cell-range-address="local.A1:local.A2"/></chart:plot-area><table:table table:name="local"><table:table-row><table:table-cell office:value="5"/></table:table-row><table:table-row><table:table-cell office:value="15"/></table:table-row></table:table></chart:chart>
          <chart:chart chart:class="chart:circle"><chart:plot-area><chart:series chart:values-cell-range-address="local.A1:local.A2"/></chart:plot-area><table:table table:name="local"><table:table-row><table:table-cell office:value="1"/></table:table-row><table:table-row><table:table-cell office:value="2"/></table:table-row></table:table></chart:chart>
        </office:document-content>"#;
        let charts = parse_ods_charts_in_xml(bytes, Limits::default(), "content.xml")
            .expect("valid inline charts")
            .into_iter()
            .map(|chart| chart.expect("chart has usable data"))
            .collect::<Vec<_>>();

        assert_eq!(charts.len(), 3);
        assert_eq!(charts[0].kind, OdsChartKind::Bar);
        assert_eq!(charts[1].kind, OdsChartKind::Line);
        assert_eq!(charts[2].kind, OdsChartKind::Pie);
        assert_eq!(charts[0].series, vec![vec![10.0, 20.0]]);
        assert_eq!(charts[1].series, vec![vec![5.0, 15.0]]);
        assert_eq!(charts[2].series, vec![vec![1.0, 2.0]]);
    }
}
