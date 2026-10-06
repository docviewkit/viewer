//! Native ODT text-flow parsing and bounded page layout.

use super::optional_xml_attribute as optional_attribute;
use std::collections::{HashMap, HashSet};

use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
use crate::font_metrics::FontMetricTable;
use crate::limits::Limits;
use crate::model::{
    AffineTransform, Document, DocumentFormat, DocumentKind, FillRule, Geometry, ImageCrop,
    MappingQuality, Object, ObjectKind, Paint, PathCommand, Rect, SourceLocator, SourceRef,
    StrokeStyle, TextAlign, TextLayout, TextOrientation, TextRun, TextTabLeader, TextTabStop, Unit,
    UnitKind, Visual,
};
use crate::package::Package;
use crate::xml::{XmlAttribute, XmlEvent, decode_xml_text, parse_xml};

use super::odf_math::{OdfMath, parse_odf_math};
use super::odp::{
    OdfConnector, OdfConnectorStyle, OdfShapeAnchor, OdpBasicChart, OdpChartKind,
    format_odp_chart_number, odf_connector_geometry, odp_chart_axis, odp_chart_axis_bounds,
    odp_chart_bar_bounds, odp_chart_category_label_bounds, odp_chart_plot,
    odp_chart_value_label_bounds, odp_chart_value_range, odp_chart_wall_fill,
    odp_pie_segment_geometry, parse_odf_connector, parse_odp_chart_part,
};
use super::presentation_image::{
    OdfImageTarget, OfficeImageError, clone_image_bytes, office_image_media_type,
    reserve_materialized_image_bytes, resolve_odf_image_target,
};
use super::{clone_materialized_text, local_name, reserve_materialized_text_bytes};

const STYLES_PART: &str = "styles.xml";
const CONTENT_PART: &str = "content.xml";
const CSS_PIXELS_PER_INCH: f32 = 96.0;
const DEFAULT_MARGIN: f32 = CSS_PIXELS_PER_INCH;
const FONT_SIZE: f32 = 16.0;
const LINE_HEIGHT: f32 = 24.0;

#[derive(Clone, Debug, PartialEq)]
struct OdtTextStyle {
    font_family: String,
    font_size: f32,
    color: u32,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    highlight: u32,
    baseline_shift: f32,
    letter_spacing: f32,
}

impl Default for OdtTextStyle {
    fn default() -> Self {
        Self {
            font_family: "Arial".to_owned(),
            font_size: FONT_SIZE,
            color: 0x0000_00ff,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            highlight: 0,
            baseline_shift: 0.0,
            letter_spacing: 0.0,
        }
    }
}

#[derive(Clone, Debug, Default)]
struct OdtTextStyleOverrides {
    font_family: Option<String>,
    font_size: Option<f32>,
    font_size_scale: Option<f32>,
    color: Option<u32>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strikethrough: Option<bool>,
    highlight: Option<u32>,
    baseline_shift: Option<f32>,
    letter_spacing: Option<f32>,
}

impl OdtTextStyleOverrides {
    fn apply(&self, style: &mut OdtTextStyle) {
        if let Some(value) = self.font_family.as_ref() {
            style.font_family.clone_from(value);
        }
        if let Some(value) = self.font_size {
            style.font_size = value;
        }
        if let Some(value) = self.font_size_scale {
            style.font_size *= value;
        }
        if let Some(value) = self.color {
            style.color = value;
        }
        if let Some(value) = self.bold {
            style.bold = value;
        }
        if let Some(value) = self.italic {
            style.italic = value;
        }
        if let Some(value) = self.underline {
            style.underline = value;
        }
        if let Some(value) = self.strikethrough {
            style.strikethrough = value;
        }
        if let Some(value) = self.highlight {
            style.highlight = value;
        }
        if let Some(value) = self.baseline_shift {
            style.baseline_shift = value;
        }
        if let Some(value) = self.letter_spacing {
            style.letter_spacing = value;
        }
    }
}

#[derive(Clone, Debug)]
struct OdtParagraphStyle {
    text: OdtTextStyle,
    align: TextAlign,
    line_height: f32,
    natural_line_height: bool,
    margin_top: f32,
    margin_bottom: f32,
    contextual_spacing: bool,
    common_style_name: Option<String>,
    margin_left: f32,
    margin_right: f32,
    text_indent: f32,
    page_break_before: bool,
    page_number: Option<u32>,
    page_break_after: bool,
    keep_with_next: bool,
    keep_together: bool,
    tab_stops: Vec<TextTabStop>,
    default_tab_stop: f32,
    drop_cap: Option<crate::model::TextDropCap>,
}

impl Default for OdtParagraphStyle {
    fn default() -> Self {
        Self {
            text: OdtTextStyle::default(),
            align: TextAlign::Start,
            line_height: LINE_HEIGHT,
            natural_line_height: true,
            margin_top: 0.0,
            margin_bottom: 0.0,
            contextual_spacing: false,
            common_style_name: None,
            margin_left: 0.0,
            margin_right: 0.0,
            text_indent: 0.0,
            page_break_before: false,
            page_number: None,
            page_break_after: false,
            keep_with_next: false,
            keep_together: false,
            tab_stops: Vec::new(),
            default_tab_stop: TextLayout::default().default_tab_stop,
            drop_cap: None,
        }
    }
}

#[derive(Clone, Debug)]
struct OdtStyle {
    family: String,
    master_page_name: Option<String>,
    paragraph: OdtParagraphStyle,
    graphic: OdtGraphicStyle,
    cell: OdtCellStyle,
    row: OdtRowStyle,
    column: OdtColumnStyle,
    text_overrides: OdtTextStyleOverrides,
    columns: Option<(u32, f32)>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum OdtWrap {
    RunThrough,
    NoWrap,
    Parallel,
}

#[derive(Clone, Debug)]
struct OdtGraphicStyle {
    explicit_stroke: bool,
    wrap: OdtWrap,
    margin_top: f32,
    allow_overlap: bool,
    horizontal_rel: Option<String>,
    vertical_rel: Option<String>,
    horizontal_pos: TextAlign,
    vertical_pos: TextAlign,
    fill: Paint,
    stroke: Paint,
    stroke_width: f32,
    marker_start: bool,
    marker_end: bool,
}

impl Default for OdtGraphicStyle {
    fn default() -> Self {
        Self {
            explicit_stroke: false,
            wrap: OdtWrap::RunThrough,
            margin_top: 0.0,
            allow_overlap: true,
            horizontal_rel: None,
            vertical_rel: None,
            horizontal_pos: TextAlign::Start,
            vertical_pos: TextAlign::Start,
            fill: Paint::Solid(0xf2f2_f2ff),
            stroke: Paint::Solid(0x8080_80ff),
            stroke_width: 1.0,
            marker_start: false,
            marker_end: false,
        }
    }
}

#[derive(Clone, Debug)]
struct OdtCellStyle {
    fill: u32,
    stroke: u32,
    stroke_width: f32,
    stroke_style: StrokeStyle,
    orientation: TextOrientation,
    padding: [f32; 4],
}

impl Default for OdtCellStyle {
    fn default() -> Self {
        Self {
            fill: 0xffff_ffff,
            stroke: 0x8080_80ff,
            stroke_width: 0.0,
            stroke_style: StrokeStyle::default(),
            orientation: TextOrientation::Horizontal,
            padding: [4.0; 4],
        }
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct OdtRowStyle {
    min_height: f32,
}

/// Authored `style:table-column-properties` size hints for one logical column.
#[derive(Clone, Copy, Debug, Default)]
struct OdtColumnStyle {
    /// `style:column-width` in CSS pixels.
    width: Option<f32>,
    /// `style:rel-column-width` star/percent weight.
    rel_width: Option<f32>,
}

impl OdtColumnStyle {
    fn size_hint(self) -> Option<f32> {
        self.rel_width
            .or(self.width)
            .filter(|value| value.is_finite() && *value > 0.0)
    }
}

#[derive(Clone, Debug)]
enum OdtListLevel {
    Number {
        format: String,
        prefix: String,
        suffix: String,
        start: u32,
    },
    Bullet {
        character: String,
        font_family: Option<String>,
    },
}

#[derive(Clone, Debug, Default)]
struct OdtListStyle {
    levels: HashMap<u32, OdtListLevel>,
    positions: HashMap<u32, OdtListPosition>,
}

#[derive(Clone, Debug)]
struct OdtListPosition {
    margin_left: f32,
    text_indent: f32,
    tab_stop: Option<f32>,
    followed_by: &'static str,
}

#[derive(Clone, Debug)]
struct MasterStoryParagraph {
    text: String,
    style: OdtParagraphStyle,
    runs: Vec<TextRun>,
    spans: Vec<(usize, OdtTextStyle)>,
}

/// One ODF master-page story slot (`header` / `header-left` / `header-first` …).
#[derive(Clone, Debug, Default)]
struct MasterStory {
    /// Element present in the package, even when its content is blank.
    present: bool,
    paragraphs: Vec<MasterStoryParagraph>,
    shapes: Vec<ShapeState>,
    table_objects: Vec<Object>,
    has_table: bool,
}

/// ODF first / left / right story slots for a master page.
///
/// Unlike Word's `select_word_story`, an absent first/left element falls back to
/// the default story; only an explicitly present (possibly blank) variant
/// overrides that page class.
#[derive(Clone, Debug, Default)]
struct MasterStoryVariants {
    default: MasterStory,
    left: MasterStory,
    first: MasterStory,
}

impl MasterStoryVariants {
    fn slot_mut(&mut self, kind: &str) -> &mut MasterStory {
        match kind {
            "header-first" | "footer-first" => &mut self.first,
            "header-left" | "footer-left" => &mut self.left,
            _ => &mut self.default,
        }
    }

    fn slot(&self, kind: &str) -> &MasterStory {
        match kind {
            "header-first" | "footer-first" => &self.first,
            "header-left" | "footer-left" => &self.left,
            _ => &self.default,
        }
    }

    fn select(&self, is_header: bool, first: bool, even: bool) -> (&'static str, &MasterStory) {
        if first && self.first.present {
            (
                if is_header {
                    "header-first"
                } else {
                    "footer-first"
                },
                &self.first,
            )
        } else if even && self.left.present {
            (
                if is_header {
                    "header-left"
                } else {
                    "footer-left"
                },
                &self.left,
            )
        } else {
            (if is_header { "header" } else { "footer" }, &self.default)
        }
    }
}

fn master_story_slot<'a>(page: &'a mut MasterPage, kind: &str) -> &'a mut MasterStory {
    if kind.starts_with("header") {
        page.headers.slot_mut(kind)
    } else {
        page.footers.slot_mut(kind)
    }
}

fn master_story_ref<'a>(page: &'a MasterPage, kind: &str) -> &'a MasterStory {
    if kind.starts_with("header") {
        page.headers.slot(kind)
    } else {
        page.footers.slot(kind)
    }
}

/// Western binding: unit 0 is page 1 (recto/right); odd units are left pages.
fn is_left_page(unit_index: u32) -> bool {
    unit_index % 2 == 1
}

fn master_story_kind(local: &str) -> Option<&'static str> {
    Some(match local {
        "header" => "header",
        "header-left" => "header-left",
        "header-first" => "header-first",
        "footer" => "footer",
        "footer-left" => "footer-left",
        "footer-first" => "footer-first",
        _ => return None,
    })
}

const MASTER_STORY_KINDS: [&str; 6] = [
    "header",
    "header-left",
    "header-first",
    "footer",
    "footer-left",
    "footer-first",
];

#[derive(Clone, Debug, Default)]
struct MasterPage {
    page_layout: Option<String>,
    next: Option<String>,
    headers: MasterStoryVariants,
    footers: MasterStoryVariants,
}

/// ODT story box metrics are separate from page margins and drawing shapes.
#[derive(Clone, Debug, Default)]
struct MasterStoryStyle {
    height: f32,
    min_height: f32,
    margin: [f32; 4], // top, right, bottom, left
    padding: [f32; 4],
    border: OdtCellStyle,
}

#[derive(Debug, Default)]
struct StyleCatalog<'a> {
    font_metrics: Option<&'a FontMetricTable>,
    page_backgrounds: HashMap<String, Paint>,
    story_styles: HashMap<String, [Option<MasterStoryStyle>; 2]>,
    styles: HashMap<String, OdtStyle>,
    default_paragraph: OdtParagraphStyle,
    default_graphic: OdtGraphicStyle,
    default_cell: OdtCellStyle,
    default_row: OdtRowStyle,
    default_table_column: OdtColumnStyle,
    font_faces: HashMap<String, String>,
    list_styles: HashMap<String, OdtListStyle>,
    master_pages: HashMap<String, MasterPage>,
    default_master_page: Option<String>,
    initial_master_page: Option<String>,
    master_page_starts: Vec<(u32, String)>,
}

impl StyleCatalog<'_> {
    fn paragraph(&self, name: Option<&str>) -> OdtParagraphStyle {
        name.and_then(|name| self.styles.get(name)).map_or_else(
            || self.default_paragraph.clone(),
            |style| style.paragraph.clone(),
        )
    }

    fn text(&self, name: Option<&str>, inherited: &OdtTextStyle) -> OdtTextStyle {
        let Some(style) = name
            .and_then(|name| self.styles.get(name))
            .filter(|style| matches!(style.family.as_str(), "text" | "paragraph"))
        else {
            return inherited.clone();
        };
        let mut resolved = inherited.clone();
        style.text_overrides.apply(&mut resolved);
        resolved
    }

    fn graphic(&self, name: Option<&str>) -> OdtGraphicStyle {
        name.and_then(|name| self.styles.get(name)).map_or_else(
            || self.default_graphic.clone(),
            |style| style.graphic.clone(),
        )
    }

    fn cell(&self, name: Option<&str>) -> OdtCellStyle {
        name.and_then(|name| self.styles.get(name))
            .map_or_else(|| self.default_cell.clone(), |style| style.cell.clone())
    }

    fn row(&self, name: Option<&str>) -> OdtRowStyle {
        name.and_then(|name| self.styles.get(name))
            .map_or(self.default_row, |style| style.row)
    }

    fn table_column(&self, name: Option<&str>) -> OdtColumnStyle {
        name.and_then(|name| self.styles.get(name))
            .map_or(self.default_table_column, |style| style.column)
    }

    fn master_page(&self, unit_index: u32) -> Option<(&str, &MasterPage)> {
        let start = self
            .master_page_starts
            .iter()
            .rev()
            .find(|(index, _)| *index <= unit_index);
        let mut name = start
            .map(|(_, name)| name.as_str())
            .or(self.initial_master_page.as_deref())
            .or(self.default_master_page.as_deref())?;
        for _ in start.map_or(0, |(index, _)| *index)..unit_index {
            let Some(next) = self
                .master_pages
                .get(name)
                .and_then(|master| master.next.as_deref())
            else {
                break;
            };
            name = next;
        }
        self.master_pages.get(name).map(|master| (name, master))
    }

    /// True when this unit starts a new run of its master page style.
    fn is_master_style_first_page(&self, unit_index: u32) -> bool {
        if unit_index == 0 {
            return true;
        }
        let current = self.master_page(unit_index).map(|(name, _)| name);
        let previous = self
            .master_page(unit_index.saturating_sub(1))
            .map(|(name, _)| name);
        current.is_some() && current != previous
    }

    fn master_stories(
        &self,
        unit_index: u32,
    ) -> Option<(&str, &'static str, &MasterStory, &'static str, &MasterStory)> {
        let (name, master) = self.master_page(unit_index)?;
        let first = self.is_master_style_first_page(unit_index);
        let even = is_left_page(unit_index);
        let (header_kind, header) = master.headers.select(true, first, even);
        let (footer_kind, footer) = master.footers.select(false, first, even);
        Some((name, header_kind, header, footer_kind, footer))
    }
}

#[derive(Clone, Copy, Debug)]
struct PageLayout {
    width: f32,
    height: f32,
    margin_top: f32,
    margin_right: f32,
    margin_bottom: f32,
    margin_left: f32,
}

impl Default for PageLayout {
    fn default() -> Self {
        Self {
            width: 8.5 * CSS_PIXELS_PER_INCH,
            height: 11.0 * CSS_PIXELS_PER_INCH,
            margin_top: DEFAULT_MARGIN,
            margin_right: DEFAULT_MARGIN,
            margin_bottom: DEFAULT_MARGIN,
            margin_left: DEFAULT_MARGIN,
        }
    }
}

#[derive(Debug)]
struct ParagraphState {
    depth: usize,
    index: u32,
    element_id: Option<String>,
    text: String,
    style: OdtParagraphStyle,
    runs: Vec<TextRun>,
    spans: Vec<(usize, OdtTextStyle)>,
    soft_page_breaks: Vec<usize>,
    source_text_start: u32,
    list_label: Option<String>,
    list_position: Option<OdtListPosition>,
    frames: Vec<FrameState>,
}

#[derive(Debug)]
struct FrameState {
    depth: usize,
    index: u32,
    source_element: &'static str,
    element_id: Option<String>,
    x: Option<f32>,
    y: Option<f32>,
    width: f32,
    height: f32,
    z: Option<i32>,
    inline: bool,
    page_anchored: bool,
    wrap: OdtWrap,
    transform: AffineTransform,
    graphic: OdtGraphicStyle,
    chart: Option<OdpBasicChart>,
    math: Option<OdfMath>,
    image_hrefs: Vec<String>,
    auto_width: bool,
    auto_height: bool,
    children: Vec<FrameState>,
    text: String,
    runs: Vec<TextRun>,
    text_style: OdtTextStyle,
    spans: Vec<(usize, OdtTextStyle)>,
    text_box_depth: Option<usize>,
    paragraph_depth: Option<usize>,
    paragraph_count: u32,
}

#[derive(Clone, Debug)]
struct ShapeState {
    depth: usize,
    index: u32,
    element: &'static str,
    element_id: Option<String>,
    geometry: Geometry,
    connector: Option<OdfConnector>,
    bounds: Rect,
    z: Option<i32>,
    flow_anchored: bool,
    transform: AffineTransform,
    graphic: OdtGraphicStyle,
    text: String,
    text_style: OdtParagraphStyle,
    runs: Vec<TextRun>,
    spans: Vec<(usize, OdtTextStyle)>,
    paragraph_depth: Option<usize>,
}

#[derive(Debug)]
struct DrawGroupState {
    depth: usize,
    character_anchored: bool,
    margin_top: f32,
    shapes: Vec<ShapeState>,
}

#[derive(Debug)]
struct ListContext {
    depth: usize,
    style_name: Option<String>,
    level: u32,
    next_value: u32,
    item_depth: Option<usize>,
    pending_label: Option<String>,
}

struct NoteState {
    depth: usize,
    order: usize,
    parent: Option<ParagraphState>,
    citation: String,
    citation_depth: Option<usize>,
    paragraphs: Vec<ParagraphState>,
    unit_index: u32,
    endnote: bool,
}

#[derive(Debug)]
struct TableState {
    depth: usize,
    index: u32,
    element_id: Option<String>,
    next_source_row: u32,
    declared_columns: u32,
    column_styles: Vec<OdtColumnStyle>,
    logical_rows: u32,
    projected_cells: usize,
    rows: Vec<TableRow>,
}

#[derive(Debug)]
struct RowState {
    depth: usize,
    source_index: u32,
    repeat: u32,
    next_column: u32,
    next_source_cell: u32,
    min_height: f32,
    cells: Vec<TableCell>,
}

#[derive(Debug)]
struct CellState {
    tables: Vec<TableState>,
    depth: usize,
    source_index: u32,
    first_column: u32,
    repeat: u32,
    element_id: Option<String>,
    style: OdtCellStyle,
    paragraph_depth: Option<usize>,
    paragraph_count: u32,
    text: String,
    paragraph_style: OdtParagraphStyle,
    column_span: u32,
    row_span: u32,
    image_hrefs: Vec<String>,
}

#[derive(Debug)]
struct TableRow {
    source_index: u32,
    first_row: u32,
    repeat: u32,
    columns: u32,
    min_height: f32,
    cells: Vec<TableCell>,
}

#[derive(Debug)]
struct TableCell {
    tables: Vec<TableState>,
    source_index: u32,
    first_column: u32,
    repeat: u32,
    element_id: Option<String>,
    style: OdtCellStyle,
    text: String,
    paragraph_style: OdtParagraphStyle,
    column_span: u32,
    row_span: u32,
    image_hrefs: Vec<String>,
}

#[derive(Debug)]
struct ParsedContent {
    materialized_image_bytes: usize,
    units: Vec<Unit>,
    objects: Vec<Object>,
    diagnostics: Vec<Diagnostic>,
}

#[cfg(test)]
pub fn parse(package: &Package<'_>) -> Result<Document, Diagnostic> {
    parse_with_font_metrics(package, &FontMetricTable::default())
}

pub fn parse_with_font_metrics(
    package: &Package<'_>,
    font_metrics: &FontMetricTable,
) -> Result<Document, Diagnostic> {
    let styles_bytes = package.required_part(STYLES_PART)?;
    let bytes = package.required_part(CONTENT_PART)?;
    let mut styles = StyleCatalog {
        font_metrics: Some(font_metrics),
        ..StyleCatalog::default()
    };
    parse_style_catalog_xml(&styles_bytes, package.limits(), STYLES_PART, &mut styles)?;
    parse_master_stories_xml(&styles_bytes, package.limits(), &mut styles)?;
    parse_style_catalog_xml(&bytes, package.limits(), CONTENT_PART, &mut styles)?;
    select_initial_master_page(&bytes, package.limits(), &mut styles)?;
    let layout = parse_page_layout_xml(
        &styles_bytes,
        package.limits(),
        styles.initial_master_page.as_deref(),
    )?;
    validate_page_layout(layout, package.limits())?;
    let background_diagnostics =
        parse_page_backgrounds_xml(&styles_bytes, package.limits(), Some(package), &mut styles)?;
    let story_diagnostics = prepare_master_tables(
        &styles_bytes,
        layout,
        package.limits(),
        &mut styles,
        Some(package),
    )?;
    let mut content =
        parse_content_xml(&bytes, layout, package.limits(), &mut styles, Some(package))?;
    content.diagnostics.extend(story_diagnostics);
    content.diagnostics.extend(background_diagnostics);
    append_page_backgrounds(&mut content, &styles, package.limits())?;

    Ok(Document {
        fatal: false,
        format: Some(DocumentFormat::Odt),
        kind: Some(DocumentKind::Text),
        units: content.units,
        outline: Vec::new(),
        objects: content.objects,
        embedded_fonts: Vec::new(),
        font_alternate_names: Vec::new(),
        diagnostics: content.diagnostics,
    })
}

#[cfg(test)]
pub(super) fn parse_flat(bytes: &[u8], limits: Limits) -> Result<Document, Diagnostic> {
    parse_flat_with_font_metrics(bytes, limits, &FontMetricTable::default())
}

pub(super) fn parse_flat_with_font_metrics(
    bytes: &[u8],
    limits: Limits,
    font_metrics: &FontMetricTable,
) -> Result<Document, Diagnostic> {
    let mut styles = StyleCatalog {
        font_metrics: Some(font_metrics),
        ..StyleCatalog::default()
    };
    parse_style_catalog_xml(bytes, limits, CONTENT_PART, &mut styles)?;
    parse_master_stories_xml(bytes, limits, &mut styles)?;
    select_initial_master_page(bytes, limits, &mut styles)?;
    let layout = parse_page_layout_xml(bytes, limits, styles.initial_master_page.as_deref())?;
    validate_page_layout(layout, limits)?;
    let background_diagnostics = parse_page_backgrounds_xml(bytes, limits, None, &mut styles)?;
    let story_diagnostics = prepare_master_tables(bytes, layout, limits, &mut styles, None)?;
    let mut content = parse_content_xml(bytes, layout, limits, &mut styles, None)?;
    content.diagnostics.extend(story_diagnostics);
    content.diagnostics.extend(background_diagnostics);
    append_page_backgrounds(&mut content, &styles, limits)?;
    let mut document = Document {
        fatal: false,
        format: Some(DocumentFormat::Odt),
        kind: Some(DocumentKind::Text),
        units: content.units,
        outline: Vec::new(),
        objects: content.objects,
        embedded_fonts: Vec::new(),
        font_alternate_names: Vec::new(),
        diagnostics: content.diagnostics,
    };
    document
        .diagnostics
        .extend(super::security::inspect_flat_odf(bytes, limits)?);
    Ok(document)
}

fn parse_content_xml(
    bytes: &[u8],
    layout: PageLayout,
    limits: Limits,
    styles: &mut StyleCatalog,
    package: Option<&Package<'_>>,
) -> Result<ParsedContent, Diagnostic> {
    parse_story_xml(bytes, layout, limits, styles, package, None)
}

// Resolve both sides before pagination: removing a previous paragraph's margin
// after it has already caused a page break is too late. ODF differs from DOCX:
// both paragraphs must opt in and automatic styles compare their common parent.
fn odt_paragraph_spacing(
    bytes: &[u8],
    limits: Limits,
    styles: &StyleCatalog,
    deleted_ids: &HashSet<String>,
) -> Result<Vec<(f32, f32)>, Diagnostic> {
    let mut spacing: Vec<(f32, f32)> = Vec::new();
    let mut previous: Option<(usize, OdtParagraphStyle, usize)> = None;
    let mut area = 0_usize;
    let mut depth = 0_usize;
    let mut skipped = None;
    let mut deleted = HashSet::new();
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if (matches!(local, "tracked-changes" | "annotation") || name.starts_with("draw:"))
                    && skipped.is_none()
                    && !empty
                {
                    skipped = Some(depth);
                }
                if matches!(local, "change-start" | "change-end")
                    && skipped.is_none()
                    && let Some(id) = optional_attribute(&attributes, "change-id", CONTENT_PART)?
                    && deleted_ids.contains(&id)
                {
                    if local == "change-start" {
                        deleted.insert(id);
                    } else {
                        deleted.remove(&id);
                    }
                }
                if matches!(local, "p" | "h") {
                    let style_name = optional_attribute(&attributes, "style-name", CONTENT_PART)?;
                    let style = styles.paragraph(style_name.as_deref());
                    let index = spacing.len();
                    spacing.push((style.margin_top, style.margin_bottom));
                    if skipped.is_none() && deleted.is_empty() {
                        let mut compared = style;
                        if let Some(conditional) =
                            optional_attribute(&attributes, "cond-style-name", CONTENT_PART)?
                        {
                            compared.common_style_name = Some(conditional);
                        }
                        if let Some((previous_index, previous_style, previous_area)) =
                            previous.as_ref().filter(|(_, previous_style, _)| {
                                !compared.page_break_before && !previous_style.page_break_after
                            })
                        {
                            if *previous_area == area
                                && compared.contextual_spacing
                                && previous_style.contextual_spacing
                                && compared.common_style_name.is_some()
                                && compared.common_style_name == previous_style.common_style_name
                            {
                                spacing[*previous_index].1 = 0.0;
                                spacing[index].0 = 0.0;
                            } else {
                                spacing[index].0 = super::collapsed_paragraph_space_before(
                                    spacing[index].0,
                                    spacing[*previous_index].1,
                                );
                            }
                        }
                        previous = Some((index, compared, area));
                    }
                } else if skipped.is_none() && deleted.is_empty() {
                    if local == "section" {
                        area += 1;
                    } else if matches!(
                        local,
                        "table"
                            | "table-cell"
                            | "text-box"
                            | "note"
                            | "header"
                            | "footer"
                            | "text"
                            | "list"
                            | "soft-page-break"
                    ) {
                        previous = None;
                        area += 1;
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if skipped == Some(depth) {
                    skipped = None;
                } else if skipped.is_none() && deleted.is_empty() {
                    match local_name(name) {
                        "section" => area += 1,
                        "table" | "table-cell" | "text-box" | "note" | "header" | "header-left"
                        | "header-first" | "footer" | "footer-left" | "footer-first" | "text"
                        | "list" => {
                            previous = None;
                            area += 1;
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
        Ok(())
    })?;
    Ok(spacing)
}

fn parse_story_xml(
    bytes: &[u8],
    mut layout: PageLayout,
    limits: Limits,
    styles: &mut StyleCatalog,
    package: Option<&Package<'_>>,
    story: Option<(&str, &str)>,
) -> Result<ParsedContent, Diagnostic> {
    let deleted_change_ids = deleted_change_ids(bytes, limits)?;
    let paragraph_spacing = odt_paragraph_spacing(bytes, limits, styles, &deleted_change_ids)?;
    let mut paragraph_spacing = paragraph_spacing.into_iter();

    let mut depth = 0_usize;
    let mut text_body_depth = None;
    let mut master_name = String::new();
    let mut skipped_metadata_depth = None;
    let mut active_deleted_changes = HashSet::new();
    let mut paragraph = None;
    let mut paragraph_count = 0_u32;
    let mut page_number_starts = Vec::<(u32, u32)>::new();
    let mut notes = Vec::<NoteState>::new();
    let mut completed_notes = Vec::<NoteState>::new();
    let mut note_count = 0_usize;
    let mut section: Option<(usize, (u32, f32), Vec<ParagraphState>)> = None;
    let mut frame: Option<FrameState> = None;
    let mut frame_count = 0_u32;
    let mut frame_stack = Vec::<FrameState>::new();
    let mut form_image_hrefs = HashMap::<String, String>::new();
    let mut shape: Option<ShapeState> = None;
    let mut shape_count = 0_u32;
    let mut group: Option<DrawGroupState> = None;
    let mut pending_groups = Vec::<DrawGroupState>::new();
    let mut lists = Vec::<ListContext>::new();
    let mut table = None;
    let mut table_count = 0_u32;
    let mut row = None;
    let mut cell = None;
    let mut covered_cell_depth = None;
    let mut page_field_depth = None;
    let mut page_field_slot: Option<&'static str> = None;
    let mut ruby: Option<(usize, String, String, Option<RubyCapture>)> = None;
    let mut skipped_nested_frame_depth = None;
    let mut skipped_shape_depth = None;
    let mut objects = Vec::new();
    let mut units = vec![page_unit(0, layout)];
    let mut unit_index = 0_u32;
    let mut diagnostics = vec![
        Diagnostic::warning(
            DiagnosticCode::ApproximateLayout,
            Phase::Layout,
            Fidelity::Approximate,
            "ODT text flow uses browser font metrics when supplied, with bounded estimates for missing metrics and authored tab stops",
        )
        .in_part(CONTENT_PART),
    ];
    let mut table_metrics_reported = false;
    let mut table_stack = Vec::<(TableState, Option<RowState>, Option<CellState>)>::new();
    let mut nested_frames_reported = false;
    let mut unsupported_chart_reported = false;
    let mut split_tables_reported = false;
    let mut clipped_rows_reported = false;
    let mut unsupported_shape_reported = false;
    let mut accepted_revision_policy_reported = false;
    let mut materialized_text_bytes = 0_usize;
    let mut materialized_image_bytes = 0_usize;

    let mut y = if story.is_some() {
        layout.margin_top
    } else {
        push_master_stories(0, layout, &mut objects, styles, limits.max_document_objects)?
    };

    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                let spacing = if matches!(local, "p" | "h") {
                    paragraph_spacing.next()
                } else {
                    None
                };

                if local == "note" && text_body_depth.is_some() && !empty && skipped_metadata_depth.is_none() && active_deleted_changes.is_empty() {
                    if note_count >= limits.max_document_objects { return Err(object_limit_error("ODT notes exceed object limit")); }
                    notes.push(NoteState { depth, order: note_count, parent: paragraph.take(), citation: String::new(), citation_depth: None,
                        paragraphs: Vec::new(), unit_index, endnote: optional_attribute(&attributes, "note-class", CONTENT_PART)?.as_deref() == Some("endnote") });
                    note_count += 1;
                    depth += 1;
                    return Ok(());
                }
                if local == "note-citation" && !notes.is_empty() {
                    notes.last_mut().unwrap().citation_depth = (!empty).then_some(depth);
                    if !empty { depth += 1; }
                    return Ok(());
                }
                if skipped_metadata_depth.is_some() {
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }

                // Revision metadata and comments are not part of the visible text story.
                if matches!(local, "tracked-changes" | "annotation") && text_body_depth.is_some() {
                    if !empty {
                        skipped_metadata_depth = Some(depth);
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }

                if matches!(local, "change-start" | "change-end")
                    && text_body_depth.is_some()
                {
                    if let Some(change_id) =
                        optional_attribute(&attributes, "change-id", CONTENT_PART)?
                        && deleted_change_ids.contains(&change_id)
                    {
                        if local == "change-start" {
                            active_deleted_changes.insert(change_id);
                        } else {
                            active_deleted_changes.remove(&change_id);
                        }
                        if !accepted_revision_policy_reported {
                            diagnostics.push(
                                Diagnostic::warning(
                                    DiagnosticCode::UnsupportedFeature,
                                    Phase::Render,
                                    Fidelity::Approximate,
                                    "ODT tracked changes use an accepted revision view: inserted content is visible and deleted content is hidden",
                                )
                                .in_part(CONTENT_PART),
                            );
                            accepted_revision_policy_reported = true;
                        }
                    }
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }

                if !active_deleted_changes.is_empty() {
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }

                if skipped_nested_frame_depth.is_some() {
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }

                if skipped_shape_depth.is_some() {
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }

                if section.is_some() && (matches!(local, "section" | "table" | "soft-page-break")
                    || name.starts_with("draw:"))
                {
                    let (_, _, paragraphs) = section.take().unwrap();
                    diagnostics.push(section_columns_fallback());
                    push_section_paragraphs(paragraphs, (1, 0.0), layout, &mut units,
                        &mut unit_index, &mut y, &mut objects, styles, limits,
                        &mut materialized_text_bytes, &mut diagnostics)?;
                }
                if local == "section" && text_body_depth.is_some() && table.is_none() && frame.is_none() && shape.is_none() && !empty {
                    let name = optional_attribute(&attributes, "style-name", CONTENT_PART)?;
                    if let Some(columns) = name.as_ref().and_then(|name| styles.styles.get(name))
                        .and_then(|style| style.columns) {
                        if columns.0 > 1 { section = Some((depth, columns, Vec::new())); }
                        else if columns.0 == 0 { diagnostics.push(section_columns_fallback()); }
                    }
                }

                if local == "image-frame" && text_body_depth.is_some() {
                    if let (Some(id), Some(href)) = (
                        optional_attribute(&attributes, "id", CONTENT_PART)?,
                        optional_attribute(&attributes, "image-data", CONTENT_PART)?,
                    ) {
                        form_image_hrefs.insert(id, href);
                    }
                }

                if let Some(current) = shape.as_mut() {
                    match local {
                        "enhanced-geometry" => {
                            if let Some(shape_type) =
                                optional_attribute(&attributes, "type", CONTENT_PART)?
                            {
                                current.geometry = match shape_type.as_str() {
                                    "ellipse" | "circle" => Geometry::Ellipse,
                                    "line" => Geometry::Line,
                                    _ => Geometry::Rectangle,
                                };
                            }
                        }
                        "p" | "h" => {
                            if !current.text.is_empty() {
                                append_shape_text(current, "\n", limits.max_xml_bytes)?;
                            }
                            current.paragraph_depth = (!empty).then_some(depth);
                        }
                        "span" if current.paragraph_depth.is_some() => {
                            let inherited = current
                                .spans
                                .last()
                                .map_or(&current.text_style.text, |(_, style)| style);
                            let style_name =
                                optional_attribute(&attributes, "style-name", CONTENT_PART)?;
                            let style = styles.text(style_name.as_deref(), inherited);
                            if !empty {
                                current.spans.push((depth, style));
                            }
                        }
                        "s" if current.paragraph_depth.is_some() => {
                            let count = optional_positive_u32(&attributes, "c", CONTENT_PART)?
                                .unwrap_or(1) as usize;
                            append_shape_repeated(current, ' ', count, limits.max_xml_bytes)?;
                        }
                        "tab" if current.paragraph_depth.is_some() => {
                            append_shape_text(current, "\t", limits.max_xml_bytes)?;
                        }
                        "line-break" if current.paragraph_depth.is_some() => {
                            append_shape_text(current, "\n", limits.max_xml_bytes)?;
                        }
                        _ => {}
                    }
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }

                if local == "g"
                    && text_body_depth.is_some()
                    && table.is_none()
                    && group.is_none()
                {
                    let graphic = styles.graphic(
                        optional_attribute(&attributes, "style-name", CONTENT_PART)?.as_deref(),
                    );
                    if !empty {
                        group = Some(DrawGroupState {
                            depth,
                            character_anchored: optional_attribute(
                                &attributes,
                                "anchor-type",
                                CONTENT_PART,
                            )?
                            .as_deref()
                                == Some("char"),
                            margin_top: graphic.margin_top,
                            shapes: Vec::new(),
                        });
                    }
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }

                if text_body_depth.is_some()
                    && table.is_none()
                    && matches!(
                        local,
                        "custom-shape"
                            | "rect"
                            | "ellipse"
                            | "line"
                            | "connector"
                            | "measure"
                            | "regular-polygon"
                            | "caption"
                    )
                {
                    let state =
                        start_shape(local, &attributes, styles, CONTENT_PART, depth, shape_count)?;
                    shape_count = shape_count.checked_add(1).ok_or_else(|| {
                        object_limit_error("ODT shape count exceeds supported range")
                    })?;
                    if empty {
                        if let Some(group) = group.as_mut() {
                            group.shapes.push(state);
                        } else {
                            let offset_y = if state.flow_anchored { y } else { 0.0 };
                            finish_shapes(vec![state], offset_y, unit_index, &mut objects, limits)?;
                        }
                    } else {
                        shape = Some(state);
                    }
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }

                if text_body_depth.is_some()
                    && table.is_none()
                    && matches!(local, "path" | "polygon" | "polyline" | "caption")
                {
                    if !unsupported_shape_reported {
                        diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::UnsupportedFeature,
                                Phase::Render,
                                Fidelity::Omitted,
                                "ODT freeform path, polygon, and caption shapes are not rendered",
                            )
                            .in_part(CONTENT_PART),
                        );
                        unsupported_shape_reported = true;
                    }
                    if !empty {
                        skipped_shape_depth = Some(depth);
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }

                if local == "master-page" {
                    master_name = optional_attribute(&attributes, "name", STYLES_PART)?.unwrap_or_default();
                }
                if text_body_depth.is_none() && story.map_or(local == "text", |(name, kind)| master_name == name && local == kind) {
                    text_body_depth = (!empty).then_some(depth);
                } else if matches!(local, "frame" | "control")
                    && text_body_depth.is_some()
                    && table.is_none()
                {
                    if frame.as_ref().is_some_and(|parent| parent.text_box_depth.is_none()
                        || !parent.image_hrefs.is_empty() || parent.chart.is_some()) {
                        if !nested_frames_reported {
                            diagnostics.push(
                                Diagnostic::warning(
                                    DiagnosticCode::UnsupportedFeature,
                                    Phase::Render,
                                    Fidelity::Omitted,
                                    "nested ODT frames are omitted from the static fallback layout",
                                )
                                .in_part(CONTENT_PART),
                            );
                            nested_frames_reported = true;
                        }
                        if !empty {
                            skipped_nested_frame_depth = Some(depth);
                            depth = depth.saturating_add(1);
                        }
                        return Ok(());
                    }
                    let width = optional_length(&attributes, "width", CONTENT_PART)?;
                    let height = optional_length(&attributes, "height", CONTENT_PART)?;
                    let auto_width = width.is_none();
                    let auto_height = height.is_none();
                    let width = width.unwrap_or(CSS_PIXELS_PER_INCH);
                    let height = height.unwrap_or(CSS_PIXELS_PER_INCH);
                    if width == 0.0 || height == 0.0 {
                        diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::ApproximateLayout,
                                Phase::Layout,
                                Fidelity::Approximate,
                                "ODT frame has a zero dimension; its mapped bounds are clamped to 1 pixel",
                            )
                            .in_part(CONTENT_PART),
                        );
                    }
                    let anchor_type =
                        optional_attribute(&attributes, "anchor-type", CONTENT_PART)?;
                    let inline = anchor_type.as_deref() == Some("as-char");
                    let mut graphic = styles.graphic(
                        optional_attribute(&attributes, "style-name", CONTENT_PART)?.as_deref(),
                    );
                    if local == "control" {
                        graphic.fill = Paint::Solid(0xffff_ffff);
                    }
                    if frame_count as usize >= limits.max_document_objects {
                        return Err(object_limit_error("ODT frame count exceeds the configured object limit"));
                    }
                    if let Some(parent) = frame.take() { frame_stack.push(parent); }
                    let state = FrameState {
                        depth,
                        index: frame_count,
                        source_element: if local == "control" { "control" } else { "frame" },
                        element_id: optional_attribute(&attributes, "id", CONTENT_PART)?
                            .or(optional_attribute(&attributes, "name", CONTENT_PART)?),
                        x: optional_signed_length(&attributes, "x", CONTENT_PART)?,
                        y: optional_signed_length(&attributes, "y", CONTENT_PART)?,
                        width,
                        height,
                        z: optional_z_index(&attributes, CONTENT_PART)?,
                        inline,
                        page_anchored: anchor_type.as_deref() == Some("page"),
                        wrap: if inline { OdtWrap::NoWrap } else { graphic.wrap },
                        transform: optional_attribute(&attributes, "transform", CONTENT_PART)?
                            .map(|value| parse_odf_transform(&value, CONTENT_PART))
                            .transpose()?
                            .unwrap_or_default(),
                        graphic,
                        chart: None,
                        math: None,
                        auto_width,
                        auto_height,
                        children: Vec::new(),
                        runs: Vec::new(),
                        text_style: OdtTextStyle::default(),
                        spans: Vec::new(),
                        image_hrefs: if local == "control" {
                            optional_attribute(&attributes, "control", CONTENT_PART)?
                                .and_then(|id| form_image_hrefs.get(&id).cloned())
                                .into_iter()
                                .collect()
                        } else {
                            Vec::new()
                        },
                        text: String::new(),
                        text_box_depth: None,
                        paragraph_depth: None,
                        paragraph_count: 0,
                    };
                    frame_count = frame_count.checked_add(1).ok_or_else(|| {
                        object_limit_error("ODT frame count exceeds supported range")
                    })?;
                    if empty {
                        if let Some(mut parent) = frame_stack.pop() {
                            parent.children.push(state);
                            frame = Some(parent);
                        } else {
                        finish_frame(
                            state,
                            layout,
                            &mut units,
                            &mut unit_index,
                            &mut y,
                            &mut objects,
                            &mut diagnostics,
                            package,
                            limits,
                            styles,
                            &mut materialized_image_bytes,
                            None,
                        )?;
                        }
                    } else {
                        frame = Some(state);
                    }
                } else if let Some(current) = frame.as_mut() {
                    match local {
                        "object" => {
                            let parsed_chart = if let Some(href) =
                                optional_attribute(&attributes, "href", CONTENT_PART)?
                            {
                                match resolve_odf_image_target(
                                    &href,
                                    limits.max_zip_path_bytes,
                                ) {
                                    Ok(OdfImageTarget::Embedded(target)) => {
                                        let part = if target.ends_with(".xml") {
                                            target
                                        } else {
                                            format!("{target}/content.xml")
                                        };
                                        if let Some(package) =
                                            package.filter(|package| package.has_part(&part))
                                        {
                                            match package.required_part(&part).and_then(|bytes| {
                                                current.math = parse_odf_math(&bytes, limits, &part)?;
                                                if current.math.is_some() {
                                                    Ok(None)
                                                } else {
                                                    parse_odp_chart_part(package, &part)
                                                }
                                            }) {
                                                Ok(chart) => chart,
                                                Err(error) => {
                                                    diagnostics.push(
                                                        Diagnostic::warning(
                                                            error.code,
                                                            error.phase,
                                                            Fidelity::Omitted,
                                                            format!(
                                                                "ODT embedded object could not be parsed and uses its static fallback: {}",
                                                                error.message
                                                            ),
                                                        )
                                                        .in_part(&part),
                                                    );
                                                    None
                                                }
                                            }
                                        } else {
                                            None
                                        }
                                    }
                                    Ok(OdfImageTarget::External) | Err(_) => None,
                                }
                            } else {
                                None
                            };
                            if let Some(chart) =
                                parsed_chart
                            {
                                current.chart = Some(chart);
                            } else if let Some(math) = current.math.as_ref() {
                                if math.approximate || math.text.is_empty() {
                                    diagnostics.push(Diagnostic::warning(
                                        DiagnosticCode::UnsupportedFeature, Phase::Render,
                                        if math.text.is_empty() { Fidelity::Omitted } else { Fidelity::Approximate },
                                        "ODT MathML was reduced to readable linear math text",
                                    ).in_part(&math.source_part));
                                }
                                current.text = math.text.clone();
                            } else if !unsupported_chart_reported {
                                diagnostics.push(
                                    Diagnostic::warning(
                                        DiagnosticCode::UnsupportedFeature,
                                        Phase::Render,
                                        Fidelity::Approximate,
                                        "ODT embedded object has no supported native chart data and uses its static fallback",
                                    )
                                    .in_part(CONTENT_PART),
                                );
                                unsupported_chart_reported = true;
                            }
                        }
                        "image" => {
                            if let Some(href) = optional_attribute(&attributes, "href", CONTENT_PART)? {
                                current.image_hrefs.push(href);
                            } else {
                                diagnostics.push(
                                    Diagnostic::warning(
                                        DiagnosticCode::UnsupportedFeature,
                                        Phase::Parse,
                                        Fidelity::Approximate,
                                        "ODT frame image has no href; a mapped placeholder is rendered",
                                    )
                                    .in_part(CONTENT_PART),
                                );
                            }
                        }
                        "text-box" => current.text_box_depth = (!empty).then_some(depth),
                        "p" | "h" if current.text_box_depth.is_some() => {
                            current.text_style = styles.paragraph(optional_attribute(&attributes, "style-name", CONTENT_PART)?.as_deref()).text;
                            if current.paragraph_count != 0 {
                                append_frame_text(current, "\n", limits.max_xml_bytes)?;
                            }
                            current.paragraph_count = current
                                .paragraph_count
                                .checked_add(1)
                                .ok_or_else(|| {
                                    object_limit_error(
                                        "ODT frame paragraph count exceeds supported range",
                                    )
                                })?;
                            current.paragraph_depth = (!empty).then_some(depth);
                        }
                        "span" if current.paragraph_depth.is_some() && !empty => {
                            let inherited = current.spans.last().map_or(&current.text_style, |(_, style)| style);
                            let style = styles.text(optional_attribute(&attributes, "style-name", CONTENT_PART)?.as_deref(), inherited);
                            current.spans.push((depth, style));
                        }
                        "s" if current.paragraph_depth.is_some() => {
                            let count = optional_positive_u32(&attributes, "c", CONTENT_PART)?
                                .unwrap_or(1) as usize;
                            if count > limits.max_xml_bytes { return Err(object_limit_error("ODT frame spaces exceed limit")); }
                            append_frame_text(current, &" ".repeat(count), limits.max_xml_bytes)?;
                        }
                        "tab" if current.paragraph_depth.is_some() => {
                            append_frame_text(current, "\t", limits.max_xml_bytes)?;
                        }
                        "line-break" if current.paragraph_depth.is_some() => {
                            append_frame_text(current, "\n", limits.max_xml_bytes)?;
                        }
                        _ => {}
                    }
                } else if local == "list" && text_body_depth.is_some() && table.is_none() {
                    let style_name = optional_attribute(&attributes, "style-name", CONTENT_PART)?
                        .or_else(|| lists.last().and_then(|list| list.style_name.clone()));
                    let level = u32::try_from(lists.len() + 1).map_err(|_| {
                        object_limit_error("ODT list nesting exceeds supported range")
                    })?;
                    let next_value = list_start(styles, style_name.as_deref(), level);
                    let context = ListContext {
                        depth,
                        style_name,
                        level,
                        next_value,
                        item_depth: None,
                        pending_label: None,
                    };
                    if !empty {
                        lists.push(context);
                    }
                } else if local == "list-item"
                    && text_body_depth.is_some()
                    && table.is_none()
                    && !lists.is_empty()
                {
                    let list = lists.last_mut().ok_or_else(|| {
                        format_error(CONTENT_PART, "ODT list parser state is missing")
                    })?;
                    if list.item_depth.is_some() {
                        return Err(format_error(CONTENT_PART, "nested ODT list items are invalid"));
                    }
                    let value = optional_positive_u32(&attributes, "start-value", CONTENT_PART)?
                        .unwrap_or(list.next_value);
                    list.pending_label = format_list_label(
                        styles,
                        list.style_name.as_deref(),
                        list.level,
                        value,
                    );
                    list.next_value = value.saturating_add(1);
                    if empty {
                        list.pending_label = None;
                    } else {
                        list.item_depth = Some(depth);
                    }
                } else if local == "table" && text_body_depth.is_some() {
                    if let Some(parent) = table.take() {
                        if cell.is_none() { return Err(format_error(CONTENT_PART, "nested ODT table must be inside a cell")); }
                        table_stack.push((parent, row.take(), cell.take()));
                    }
                    {
                        if paragraph.is_some() {
                            return Err(format_error(
                                CONTENT_PART,
                                "table starts before its enclosing paragraph closes",
                            ));
                        }
                        let state = TableState {
                            depth,
                            index: table_count,
                            element_id: optional_attribute(&attributes, "id", CONTENT_PART)?,
                            next_source_row: 0,
                            declared_columns: 0,
                            column_styles: Vec::new(),
                            logical_rows: 0,
                            projected_cells: 0,
                            rows: Vec::new(),
                        };
                        table_count = table_count.checked_add(1).ok_or_else(|| {
                            object_limit_error("ODT table count exceeds supported range")
                        })?;
                        if empty {
                            layout_table(
                                &state,
                                layout,
                                &mut units,
                                &mut unit_index,
                                &mut y,
                                &mut objects,
                                &mut diagnostics,
                                limits,
                                styles,
                                &mut split_tables_reported,
                                &mut clipped_rows_reported,
                                &mut table_metrics_reported,
                                package,
                                &mut materialized_image_bytes,
                            )?;
                        } else {
                            table = Some(state);
                        }
                    }
                } else if let Some(current_table) = table.as_mut() {
                    match local {
                        "table-column" => {
                            let repeat = bounded_repeat(
                                &attributes,
                                "number-columns-repeated",
                                limits.max_document_objects,
                            )?;
                            current_table.declared_columns = current_table
                                .declared_columns
                                .checked_add(repeat)
                                .filter(|count| {
                                    usize::try_from(*count)
                                        .is_ok_and(|count| count <= limits.max_document_objects)
                                })
                                .ok_or_else(|| {
                                    object_limit_error(
                                        "ODT declared table columns exceed the configured object limit",
                                    )
                                })?;
                            let column_style = styles.table_column(
                                optional_attribute(&attributes, "style-name", CONTENT_PART)?
                                    .as_deref(),
                            );
                            for _ in 0..repeat {
                                current_table.column_styles.push(column_style);
                            }
                        }
                        "table-row" => {
                            if row.is_some() {
                                return Err(format_error(CONTENT_PART, "nested table rows are invalid"));
                            }
                            let repeat = bounded_repeat(
                                &attributes,
                                "number-rows-repeated",
                                limits.max_document_objects,
                            )?;
                            let first_row = current_table.logical_rows;
                            current_table.logical_rows = current_table
                                .logical_rows
                                .checked_add(repeat)
                                .filter(|count| {
                                    usize::try_from(*count)
                                        .is_ok_and(|count| count <= limits.max_document_objects)
                                })
                                .ok_or_else(|| {
                                    object_limit_error(
                                        "ODT repeated table rows exceed the configured object limit",
                                    )
                                })?;
                            let source_index = current_table.next_source_row;
                            current_table.next_source_row = current_table
                                .next_source_row
                                .checked_add(1)
                                .ok_or_else(|| {
                                    object_limit_error("ODT source row count exceeds supported range")
                                })?;
                            let state = RowState {
                                depth,
                                source_index,
                                repeat,
                                next_column: 0,
                                next_source_cell: 0,
                                min_height: styles
                                    .row(
                                        optional_attribute(
                                            &attributes,
                                            "style-name",
                                            CONTENT_PART,
                                        )?
                                        .as_deref(),
                                    )
                                    .min_height,
                                cells: Vec::new(),
                            };
                            if empty {
                                finish_row(
                                    current_table,
                                    state,
                                    first_row,
                                    limits,
                                    &mut materialized_text_bytes,
                                )?;
                            } else {
                                row = Some(state);
                            }
                        }
                        "table-cell" => {
                            if cell.is_some() || covered_cell_depth.is_some() {
                                return Err(format_error(CONTENT_PART, "nested table cells are invalid"));
                            }
                            let current_row = row.as_mut().ok_or_else(|| {
                                format_error(CONTENT_PART, "table cell is not inside a row")
                            })?;
                            let repeat = bounded_repeat(
                                &attributes,
                                "number-columns-repeated",
                                limits.max_document_objects,
                            )?;
                            let first_column = current_row.next_column;
                            current_row.next_column = current_row
                                .next_column
                                .checked_add(repeat)
                                .filter(|count| {
                                    usize::try_from(*count)
                                        .is_ok_and(|count| count <= limits.max_document_objects)
                                })
                                .ok_or_else(|| {
                                    object_limit_error(
                                        "ODT repeated table cells exceed the configured object limit",
                                    )
                                })?;
                            let source_index = current_row.next_source_cell;
                            current_row.next_source_cell = current_row
                                .next_source_cell
                                .checked_add(1)
                                .ok_or_else(|| {
                                    object_limit_error("ODT source cell count exceeds supported range")
                                })?;
                            let column_span = bounded_repeat(
                                &attributes,
                                "number-columns-spanned",
                                limits.max_document_objects,
                            )?;
                            let row_span = bounded_repeat(
                                &attributes,
                                "number-rows-spanned",
                                limits.max_document_objects,
                            )?;
                            let state = CellState {
                                tables: Vec::new(),
                                depth,
                                source_index,
                                first_column,
                                repeat,
                                element_id: optional_attribute(&attributes, "id", CONTENT_PART)?,
                                style: styles.cell(
                                    optional_attribute(&attributes, "style-name", CONTENT_PART)?
                                        .as_deref(),
                                ),
                                paragraph_depth: None,
                                paragraph_count: 0,
                                text: String::new(),
                                paragraph_style: styles.paragraph(None),
                                column_span,
                                row_span,
                                image_hrefs: Vec::new(),
                            };
                            if empty {
                                current_row.cells.push(finish_cell(state));
                            } else {
                                cell = Some(state);
                            }
                        }
                        "covered-table-cell" => {
                            if cell.is_some() || covered_cell_depth.is_some() {
                                return Err(format_error(CONTENT_PART, "nested table cells are invalid"));
                            }
                            let current_row = row.as_mut().ok_or_else(|| {
                                format_error(CONTENT_PART, "covered table cell is not inside a row")
                            })?;
                            let repeat = bounded_repeat(
                                &attributes,
                                "number-columns-repeated",
                                limits.max_document_objects,
                            )?;
                            current_row.next_column = current_row
                                .next_column
                                .checked_add(repeat)
                                .filter(|count| {
                                    usize::try_from(*count)
                                        .is_ok_and(|count| count <= limits.max_document_objects)
                                })
                                .ok_or_else(|| {
                                    object_limit_error(
                                        "ODT covered cells exceed the configured object limit",
                                    )
                                })?;
                            current_row.next_source_cell = current_row
                                .next_source_cell
                                .checked_add(1)
                                .ok_or_else(|| {
                                    object_limit_error("ODT source cell count exceeds supported range")
                                })?;
                            if !empty {
                                covered_cell_depth = Some(depth);
                            }
                        }
                        "p" | "h" => {
                            if let Some(current) = cell.as_mut() {
                                // ponytail: the existing single-style cell visual uses its first paragraph;
                                // use rich cell runs when mixed paragraph/span styling is implemented.
                                if current.paragraph_count == 0 {
                                    current.paragraph_style = styles.paragraph(optional_attribute(&attributes, "style-name", CONTENT_PART)?.as_deref());
                                }
                                if current.paragraph_count != 0 {
                                    append_text(&mut current.text, "\n", limits.max_xml_bytes)?;
                                }
                                current.paragraph_count = current
                                    .paragraph_count
                                    .checked_add(1)
                                    .ok_or_else(|| {
                                        object_limit_error(
                                            "ODT cell paragraph count exceeds supported range",
                                        )
                                    })?;
                                if current.paragraph_depth.is_none() {
                                    current.paragraph_depth = (!empty).then_some(depth);
                                }
                            }
                        }
                        "s" => {
                            if let Some(current) = cell
                                .as_mut()
                                .filter(|current| current.paragraph_depth.is_some())
                            {
                                let count = optional_positive_u32(
                                    &attributes,
                                    "c",
                                    CONTENT_PART,
                                )?
                                .unwrap_or(1) as usize;
                                append_repeated(
                                    &mut current.text,
                                    ' ',
                                    count,
                                    limits.max_xml_bytes,
                                )?;
                            }
                        }
                        "tab" => {
                            if let Some(current) = cell
                                .as_mut()
                                .filter(|current| current.paragraph_depth.is_some())
                            {
                                append_text(&mut current.text, "\t", limits.max_xml_bytes)?;
                            }
                        }
                        "line-break" => {
                            if let Some(current) = cell
                                .as_mut()
                                .filter(|current| current.paragraph_depth.is_some())
                            {
                                append_text(&mut current.text, "\n", limits.max_xml_bytes)?;
                            }
                        }
                        "image" => {
                            if let Some(current) = cell.as_mut() {
                                if let Some(href) =
                                    optional_attribute(&attributes, "href", CONTENT_PART)?
                                {
                                    current.image_hrefs.push(href);
                                } else {
                                    diagnostics.push(
                                        Diagnostic::warning(
                                            DiagnosticCode::UnsupportedFeature,
                                            Phase::Render,
                                            Fidelity::Omitted,
                                            "ODT table image has no href and is omitted",
                                        )
                                        .in_part(CONTENT_PART),
                                    );
                                }
                            }
                        }
                        other if is_odf_dynamic_page_field(other).is_some() => {
                            if let Some(current) = cell
                                .as_mut()
                                .filter(|current| current.paragraph_depth.is_some())
                            {
                                let slot = is_odf_dynamic_page_field(other).unwrap();
                                if empty {
                                    append_text(&mut current.text, slot, limits.max_xml_bytes)?;
                                } else {
                                    page_field_depth = Some(depth);
                                    page_field_slot = Some(slot);
                                }
                            }
                        }
                        other if empty => {
                            if let Some(current) = cell
                                .as_mut()
                                .filter(|current| current.paragraph_depth.is_some())
                                && let Some(display) =
                                    odf_empty_field_text(other, &attributes, CONTENT_PART)?
                            {
                                append_text(&mut current.text, &display, limits.max_xml_bytes)?;
                            }
                        }
                        _ => {}
                    }
                } else if matches!(local, "p" | "h")
                    && text_body_depth.is_some()
                    && paragraph.is_none()
                {
                    let style_name = optional_attribute(&attributes, "style-name", CONTENT_PART)?;
                    let mut paragraph_style = styles.paragraph(style_name.as_deref());
                    // Master-page assignment starts a page even without fo:break-before.
                    // Keep logical numbering separate from the physical page count.
                    if story.is_none() && notes.is_empty() {
                        let master = style_name.as_deref()
                            .and_then(|name| styles.styles.get(name))
                            .and_then(|style| style.master_page_name.clone())
                            .filter(|name| styles.master_pages.contains_key(name));
                        if master.is_some()
                            || (paragraph_style.page_break_before && paragraph_style.page_number.is_some())
                        {
                            if let Some((_, _, paragraphs)) = section.take() {
                                diagnostics.push(section_columns_fallback());
                                push_section_paragraphs(paragraphs, (1, 0.0), layout, &mut units,
                                    &mut unit_index, &mut y, &mut objects, styles, limits,
                                    &mut materialized_text_bytes, &mut diagnostics)?;
                            }
                            let new_page = objects.iter().any(|object|
                                object.unit_index == unit_index && object.source.part == CONTENT_PART);

                            if new_page {
                                unit_index = unit_index.checked_add(1).ok_or_else(||
                                    format_error(CONTENT_PART, "page count exceeds the supported range"))?;
                            }
                            if let Some(name) = master {
                                let style_bytes = package.map(|p| p.required_part(STYLES_PART)).transpose()?;
                                layout = parse_page_layout_xml(
                                    style_bytes.as_deref().unwrap_or(bytes), limits, Some(&name),
                                )?;
                                validate_page_layout(layout, limits)?;
                                styles.master_page_starts.push((unit_index, name));
                            }
                            if new_page {
                                units.push(page_unit(unit_index, layout));
                            } else {
                                // Replace stories on an unused page (including the first page).
                                if let Some(start) = objects.iter().position(|object| object.unit_index == unit_index) {
                                    objects.truncate(start);
                                }
                                units[unit_index as usize] = page_unit(unit_index, layout);
                            }
                            y = push_master_stories(unit_index, layout, &mut objects, styles, limits.max_document_objects)?;
                            if let Some(number) = paragraph_style.page_number {
                                page_number_starts.push((unit_index, number));
                            }
                            paragraph_style.page_break_before = false;
                        }
                    }

                    if let Some((before, after)) = spacing {
                        paragraph_style.margin_top = before;
                        paragraph_style.margin_bottom = after;
                    }
                    let state = ParagraphState {
                        depth,
                        index: paragraph_count,
                        element_id: optional_attribute(&attributes, "id", CONTENT_PART)?,
                        text: String::new(),
                        style: paragraph_style,
                        runs: Vec::new(),
                        spans: Vec::new(),
                        soft_page_breaks: Vec::new(),
                        source_text_start: 0,
                        list_label: lists
                            .last_mut()
                            .and_then(|list| list.pending_label.take()),
                        list_position: lists.last().and_then(|list| {
                            format_list_label(styles, list.style_name.as_deref(), list.level, 1)?;
                            Some(list.style_name.as_deref().and_then(|name| styles.list_styles.get(name))
                                .and_then(|style| style.positions.get(&list.level)).cloned()
                                .unwrap_or(OdtListPosition {
                                    margin_left: list.level as f32 * 24.0,
                                    text_indent: -24.0,
                                    tab_stop: Some(list.level as f32 * 24.0),
                                    followed_by: "\t",
                                }))
                        }),
                        frames: Vec::new(),
                    };
                    paragraph_count = paragraph_count.checked_add(1).ok_or_else(|| {
                        format_error(CONTENT_PART, "paragraph count exceeds supported range")
                    })?;
                    if section.as_ref().is_some_and(|(_, _, paragraphs)|
                        paragraphs.len().saturating_add(objects.len()) >= limits.max_document_objects) {
                        return Err(object_limit_error("ODT section paragraphs exceed object limit"));
                    }
                    if empty {
                        if let Some(note) = notes.last_mut() { note.paragraphs.push(state); }
                        else if let Some((_, _, paragraphs)) = section.as_mut() {
                            paragraphs.push(state);
                        } else {
                        push_paragraph(
                            state,
                            layout,
                            &mut units,
                            &mut unit_index,
                            &mut y,
                            &mut objects,
                            styles,
                            limits.max_document_objects,
                            &mut materialized_text_bytes,
                            limits.max_total_uncompressed_bytes,
                        )?;
                        }
                    } else {
                        paragraph = Some(state);
                    }
                } else if let Some(current) = paragraph.as_mut() {
                    match local {
                        "span" => {
                            let style_name = optional_attribute(
                                &attributes,
                                "style-name",
                                CONTENT_PART,
                            )?;
                            let inherited = current
                                .spans
                                .last()
                                .map_or(&current.style.text, |(_, style)| style);
                            let style = styles.text(style_name.as_deref(), inherited);
                            if !empty {
                                current.spans.push((depth, style));
                            }
                        }
                        "s" => {
                            let count = optional_positive_u32(&attributes, "c", CONTENT_PART)?
                                .unwrap_or(1) as usize;
                            append_styled_repeated(current, ' ', count, limits.max_xml_bytes)?;
                        }
                        "tab" => append_styled_text(current, "\t", limits.max_xml_bytes)?,
                        "line-break" => append_styled_text(current, "\n", limits.max_xml_bytes)?,
                        "soft-page-break" => current.soft_page_breaks.push(current.text.len()),
                        other if is_odf_dynamic_page_field(other).is_some() => {
                            let slot = is_odf_dynamic_page_field(other).unwrap();
                            if empty {
                                append_styled_text(current, slot, limits.max_xml_bytes)?;
                            } else {
                                page_field_depth = Some(depth);
                                page_field_slot = Some(slot);
                            }
                        }
                        other if empty => {
                            if let Some(display) =
                                odf_empty_field_text(other, &attributes, CONTENT_PART)?
                            {
                                append_styled_text(current, &display, limits.max_xml_bytes)?;
                            }
                        }
                        "ruby" if !empty => {
                            ruby = Some((depth, String::new(), String::new(), None));
                        }
                        "ruby-base" => {
                            if let Some(entry) = ruby.as_mut() {
                                entry.3 = Some(RubyCapture::Base);
                            }
                        }
                        "ruby-text" => {
                            if let Some(entry) = ruby.as_mut() {
                                entry.3 = Some(RubyCapture::Text);
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
                let local = local_name(name);

                if page_field_depth == Some(depth) {
                    let slot = page_field_slot.take();
                    page_field_depth = None;
                    if let Some(slot) = slot {
                        if let Some(current) = paragraph.as_mut() {
                            append_styled_text(current, slot, limits.max_xml_bytes)?;
                        } else if let Some(current) = cell
                            .as_mut()
                            .filter(|current| current.paragraph_depth.is_some())
                        {
                            append_text(&mut current.text, slot, limits.max_xml_bytes)?;
                        }
                    }
                    return Ok(());
                }
                if matches!(local, "ruby-base" | "ruby-text")
                    && let Some((_, _, _, capture)) = ruby.as_mut()
                {
                    *capture = None;
                    return Ok(());
                }
                if local == "ruby"
                    && ruby
                        .as_ref()
                        .is_some_and(|(ruby_depth, _, _, _)| *ruby_depth == depth)
                {
                    let (_, base, ruby_text, _) = ruby.take().unwrap();
                    if let Some(current) = paragraph.as_mut() {
                        append_styled_text(current, &base, limits.max_xml_bytes)?;
                        if !ruby_text.is_empty() {
                            let mut ruby_style = current.style.text.clone();
                            ruby_style.font_size *= 0.55;
                            ruby_style.baseline_shift = current.style.text.font_size * 0.45;
                            append_text_run(
                                &mut current.text,
                                &mut current.runs,
                                &ruby_style,
                                &ruby_text,
                                limits.max_xml_bytes,
                            )?;
                        }
                    }
                    return Ok(());
                }

                if local == "note-citation" && let Some(note) = notes.last_mut() {
                    note.citation_depth = None;
                    return Ok(());
                }
                if local == "note" && notes.last().is_some_and(|note| note.depth == depth) {
                    let mut note = notes.pop().unwrap();
                    paragraph = note.parent.take();
                    if let Some(parent) = paragraph.as_mut() {
                        let mut citation_style = parent.style.text.clone();
                        citation_style.baseline_shift = citation_style.font_size * 0.35;
                        citation_style.font_size *= 0.65;
                        append_text_run(&mut parent.text, &mut parent.runs, &citation_style, &note.citation, limits.max_xml_bytes)?;
                    }
                    if let Some(first) = note.paragraphs.first_mut() {
                        first.list_label = Some(note.citation.clone());
                    }
                    completed_notes.push(note);
                    return Ok(());
                }
                if skipped_metadata_depth == Some(depth) {
                    skipped_metadata_depth = None;
                    return Ok(());
                }
                if skipped_metadata_depth.is_some() || !active_deleted_changes.is_empty() {
                    return Ok(());
                }

                if skipped_nested_frame_depth == Some(depth) {
                    skipped_nested_frame_depth = None;
                    return Ok(());
                }
                if skipped_nested_frame_depth.is_some() {
                    return Ok(());
                }

                if skipped_shape_depth == Some(depth) {
                    skipped_shape_depth = None;
                    return Ok(());
                }
                if skipped_shape_depth.is_some() {
                    return Ok(());
                }

                if shape.is_some() {
                    match local {
                        "p" | "h" => {
                            if let Some(current) = shape.as_mut()
                                && current.paragraph_depth == Some(depth)
                            {
                                current.paragraph_depth = None;
                            }
                        }
                        "span" => {
                            if let Some(current) = shape.as_mut()
                                && current
                                    .spans
                                    .last()
                                    .is_some_and(|(span_depth, _)| *span_depth == depth)
                            {
                                current.spans.pop();
                            }
                        }
                        "custom-shape" | "rect" | "ellipse" | "line" | "connector"
                            if shape
                                .as_ref()
                                .is_some_and(|current| current.depth == depth) =>
                        {
                            let finished = shape.take().ok_or_else(|| {
                                format_error(CONTENT_PART, "ODT shape state was lost")
                            })?;
                            if let Some(group) = group.as_mut() {
                                group.shapes.push(finished);
                            } else {
                                let offset_y = if finished.flow_anchored { y } else { 0.0 };
                                finish_shapes(
                                    vec![finished],
                                    offset_y,
                                    unit_index,
                                    &mut objects,
                                    limits,
                                )?;
                            }
                        }
                        _ => {}
                    }
                    return Ok(());
                }

                if local == "g"
                    && group
                        .as_ref()
                        .is_some_and(|current| current.depth == depth)
                {
                    let finished = group.take().ok_or_else(|| {
                        format_error(CONTENT_PART, "ODT draw group state was lost")
                    })?;
                    if paragraph.is_some() {
                        pending_groups.push(finished);
                    } else {
                        finish_group(finished, y, unit_index, &mut objects, limits)?;
                    }
                    return Ok(());
                }

                if frame.is_some() {
                    match local {
                        "p" | "h" => {
                            if let Some(current) = frame.as_mut()
                                && current.paragraph_depth == Some(depth)
                            {
                                current.paragraph_depth = None;
                            }
                        }
                        "span" => {
                            if let Some(current) = frame.as_mut() && current.spans.last().is_some_and(|(d, _)| *d == depth) { current.spans.pop(); }
                        }
                        "text-box" => {
                            if let Some(current) = frame.as_mut()
                                && current.text_box_depth == Some(depth)
                            {
                                current.text_box_depth = None;
                            }
                        }
                        "frame" | "control"
                            if frame
                                .as_ref()
                                .is_some_and(|current| current.depth == depth) =>
                        {
                            let mut finished = frame.take().ok_or_else(|| {
                                format_error(CONTENT_PART, "ODT frame state was lost")
                            })?;
                            if !finished.children.is_empty() {
                                if finished.auto_width {
                                    finished.width = finished.children.iter().map(|c| c.x.unwrap_or(0.0) + c.width).fold(0.0, f32::max);
                                }
                                if finished.auto_height {
                                    finished.height = finished.children.iter().map(|c| c.y.unwrap_or(0.0) + c.height).fold(0.0, f32::max);
                                }
                            }
                            // An unbounded, undecorated one-paragraph inline text box has
                            // the same line-flow semantics as its rich text in the host paragraph.
                            if frame_stack.is_empty() && finished.inline && finished.auto_width && finished.auto_height
                                && finished.children.is_empty() && finished.image_hrefs.is_empty() && finished.chart.is_none()
                                && finished.paragraph_count == 1 && !finished.graphic.explicit_stroke
                                && matches!(finished.graphic.fill, Paint::None) && finished.transform == AffineTransform::IDENTITY
                                && let Some(parent) = paragraph.as_mut() {
                                append_text(&mut parent.text, &finished.text, limits.max_xml_bytes)?;
                                parent.runs.extend(finished.runs);
                                return Ok(());
                            }
                            if let Some(mut parent) = frame_stack.pop() {
                                parent.children.push(finished);
                                frame = Some(parent);
                            } else if !finished.inline && !finished.page_anchored
                                && finished.wrap != OdtWrap::RunThrough
                                && notes.is_empty() && section.is_none()
                                && let Some(host) = paragraph.as_mut() {
                                host.frames.push(finished);
                            } else {
                                finish_frame(
                                    finished,
                                    layout,
                                    &mut units,
                                    &mut unit_index,
                                    &mut y,
                                    &mut objects,
                                    &mut diagnostics,
                                    package,
                                    limits,
                                    styles,
                                    &mut materialized_image_bytes,
                                    None,
                                )?;
                            }
                        }
                        _ => {}
                    }
                    return Ok(());
                }

                if local == "section" && section.as_ref().is_some_and(|(start, _, _)| *start == depth) {
                    let (_, columns, paragraphs) = section.take().unwrap();
                    push_section_paragraphs(paragraphs, columns, layout, &mut units,
                        &mut unit_index, &mut y, &mut objects, styles, limits,
                        &mut materialized_text_bytes, &mut diagnostics)?;
                }

                if table.is_some() {
                    match local {
                        "p" | "h" => {
                            if let Some(current) = cell.as_mut()
                                && current.paragraph_depth == Some(depth)
                            {
                                current.paragraph_depth = None;
                            }
                        }
                        "table-cell" => {
                            if cell
                                .as_ref()
                                .is_some_and(|current| current.depth == depth)
                            {
                                let finished = finish_cell(cell.take().ok_or_else(|| {
                                    format_error(CONTENT_PART, "table cell state was lost")
                                })?);
                                row.as_mut()
                                    .ok_or_else(|| {
                                        format_error(
                                            CONTENT_PART,
                                            "closed table cell has no parent row",
                                        )
                                    })?
                                    .cells
                                    .push(finished);
                            }
                        }
                        "covered-table-cell" if covered_cell_depth == Some(depth) => {
                            covered_cell_depth = None;
                        }
                        "table-row" => {
                            if row
                                .as_ref()
                                .is_some_and(|current| current.depth == depth)
                            {
                                let finished = row.take().ok_or_else(|| {
                                    format_error(CONTENT_PART, "table row state was lost")
                                })?;
                                let current_table = table.as_mut().ok_or_else(|| {
                                    format_error(CONTENT_PART, "closed row has no parent table")
                                })?;
                                let first_row = current_table
                                    .logical_rows
                                    .checked_sub(finished.repeat)
                                    .ok_or_else(|| {
                                        format_error(CONTENT_PART, "table row range underflow")
                                    })?;
                                finish_row(
                                    current_table,
                                    finished,
                                    first_row,
                                    limits,
                                    &mut materialized_text_bytes,
                                )?;
                            }
                        }
                        "table"
                            if table
                                .as_ref()
                                .is_some_and(|current| current.depth == depth) =>
                        {
                            if row.is_some() || cell.is_some() || covered_cell_depth.is_some() {
                                return Err(format_error(
                                    CONTENT_PART,
                                    "table closes before its row or cell",
                                ));
                            }
                            let finished = table.take().ok_or_else(|| {
                                format_error(CONTENT_PART, "table state was lost")
                            })?;
                            if let Some((parent, parent_row, mut parent_cell)) = table_stack.pop() {
                                parent_cell.as_mut().unwrap().tables.push(finished);
                                table = Some(parent); row = parent_row; cell = parent_cell;
                                return Ok(());
                            }
                            layout_table(
                                &finished,
                                layout,
                                &mut units,
                                &mut unit_index,
                                &mut y,
                                &mut objects,
                                &mut diagnostics,
                                limits,
                                styles,
                                &mut split_tables_reported,
                                &mut clipped_rows_reported,
                                &mut table_metrics_reported,
                                package,
                                &mut materialized_image_bytes,
                            )?;
                        }
                        _ => {}
                    }
                } else if matches!(local, "p" | "h")
                    && paragraph
                        .as_ref()
                        .is_some_and(|current| current.depth == depth)
                {
                    let mut finished = paragraph.take().ok_or_else(|| {
                        format_error(CONTENT_PART, "paragraph state was lost before closing")
                    })?;
                    // A cached break in an otherwise empty anchor paragraph belongs to
                    // its floating object, not to the text that follows that object.
                    let anchor_page_break = !finished.frames.is_empty() && finished.text.is_empty()
                        && finished.soft_page_breaks.first() == Some(&0);
                    if anchor_page_break {
                        if y > layout.margin_top {
                            unit_index = unit_index.checked_add(1).ok_or_else(||
                                format_error(CONTENT_PART, "page count exceeds supported range"))?;
                            units.push(page_unit(unit_index, layout));
                            y = push_master_stories(unit_index, layout, &mut objects, styles, limits.max_document_objects)?;
                        }
                        finished.soft_page_breaks.remove(0);
                    }
                    for frame in finished.frames.drain(..) {
                        let anchor_bottom = y + frame.y.unwrap_or(0.0) + frame.height;
                        finish_frame(frame, layout, &mut units, &mut unit_index, &mut y,
                            &mut objects, &mut diagnostics, package, limits, styles,
                            &mut materialized_image_bytes, None)?;
                        if anchor_page_break { y = y.max(anchor_bottom); }
                    }
                    if let Some(note) = notes.last_mut() { note.paragraphs.push(finished); }
                    else if let Some((_, _, paragraphs)) = section.as_mut() {
                        paragraphs.push(finished);
                    } else {
                    push_paragraph(
                        finished,
                        layout,
                        &mut units,
                        &mut unit_index,
                        &mut y,
                        &mut objects,
                        styles,
                        limits.max_document_objects,
                        &mut materialized_text_bytes,
                        limits.max_total_uncompressed_bytes,
                    )?;
                    }
                    for group in pending_groups.drain(..) {
                        finish_group(group, y, unit_index, &mut objects, limits)?;
                    }
                } else if local == "span"
                    && let Some(current) = paragraph.as_mut()
                    && current
                        .spans
                        .last()
                        .is_some_and(|(span_depth, _)| *span_depth == depth)
                {
                    current.spans.pop();
                } else if local == "list-item"
                    && let Some(list) = lists.last_mut()
                    && list.item_depth == Some(depth)
                {
                    list.item_depth = None;
                    list.pending_label = None;
                } else if local == "list"
                    && lists.last().is_some_and(|list| list.depth == depth)
                {
                    lists.pop();
                } else if local == story.map_or("text", |(_, kind)| kind) && text_body_depth == Some(depth) {
                    text_body_depth = None;
                }
            }
            XmlEvent::Text(text) => {
                if skipped_nested_frame_depth.is_some()
                    || skipped_shape_depth.is_some()
                    || covered_cell_depth.is_some()
                    || skipped_metadata_depth.is_some()
                    || !active_deleted_changes.is_empty()
                    || page_field_depth.is_some()
                {
                    return Ok(());
                }
                let decoded =
                    decode_xml_text(text).map_err(|error| with_part(error, CONTENT_PART))?;
                if let Some((_, base, ruby_text, capture)) = ruby.as_mut() {
                    match capture {
                        Some(RubyCapture::Base) => base.push_str(&decoded),
                        Some(RubyCapture::Text) => ruby_text.push_str(&decoded),
                        None => {}
                    }
                    return Ok(());
                }
                if let Some(note) = notes.last_mut().filter(|note| note.citation_depth.is_some()) {
                    append_text(&mut note.citation, &decode_xml_text(text).map_err(|error| with_part(error, CONTENT_PART))?, limits.max_xml_bytes)?;
                    return Ok(());
                }
                if let Some(current) = shape
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    append_shape_text(
                        current,
                        &decode_xml_text(text).map_err(|error| with_part(error, CONTENT_PART))?,
                        limits.max_xml_bytes,
                    )?;
                } else if group.is_some() {
                    return Ok(());
                } else if let Some(current) = frame
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    append_frame_text(current, &decode_xml_text(text).map_err(|error| with_part(error, CONTENT_PART))?, limits.max_xml_bytes)?;
                } else if frame.is_some() {
                    // XML indentation outside a frame paragraph is not host paragraph text.
                    return Ok(());
                } else if let Some(current) = cell
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    append_text(
                        &mut current.text,
                        &decode_xml_text(text).map_err(|error| with_part(error, CONTENT_PART))?,
                        limits.max_xml_bytes,
                    )?;
                } else if let Some(current) = paragraph.as_mut() {
                    append_styled_text(
                        current,
                        &decode_xml_text(text).map_err(|error| with_part(error, CONTENT_PART))?,
                        limits.max_xml_bytes,
                    )?;
                }
            }
            XmlEvent::Cdata(text) => {
                if skipped_nested_frame_depth.is_some()
                    || skipped_shape_depth.is_some()
                    || covered_cell_depth.is_some()
                    || skipped_metadata_depth.is_some()
                    || !active_deleted_changes.is_empty()
                {
                    return Ok(());
                }
                if let Some(note) = notes.last_mut().filter(|note| note.citation_depth.is_some()) {
                    append_text(&mut note.citation, text, limits.max_xml_bytes)?;
                    return Ok(());
                }
                if let Some(current) = shape
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    append_shape_text(current, text, limits.max_xml_bytes)?;
                } else if group.is_some() {
                    return Ok(());
                } else if let Some(current) = frame
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    append_frame_text(current, text, limits.max_xml_bytes)?;
                } else if frame.is_some() {
                    // XML indentation outside a frame paragraph is not host paragraph text.
                    return Ok(());
                } else if let Some(current) = cell
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    append_text(&mut current.text, text, limits.max_xml_bytes)?;
                } else if let Some(current) = paragraph.as_mut() {
                    append_styled_text(current, text, limits.max_xml_bytes)?;
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, CONTENT_PART))?;

    if table.is_some()
        || row.is_some()
        || cell.is_some()
        || covered_cell_depth.is_some()
        || frame.is_some()
        || shape.is_some()
        || group.is_some()
        || !pending_groups.is_empty()
        || skipped_shape_depth.is_some()
        || !lists.is_empty()
    {
        return Err(format_error(
            CONTENT_PART,
            "ODT content ends before its table structure closes",
        ));
    }

    completed_notes.sort_by_key(|note| note.order);
    for page in 0..units.len() as u32 {
        let start = objects.len();
        let mut note_y = layout.margin_top;
        let mut note_page = page;
        for note in completed_notes
            .iter_mut()
            .filter(|note| !note.endnote && note.unit_index == page)
        {
            for paragraph in std::mem::take(&mut note.paragraphs) {
                push_paragraph(
                    paragraph,
                    layout,
                    &mut units,
                    &mut note_page,
                    &mut note_y,
                    &mut objects,
                    styles,
                    limits.max_document_objects,
                    &mut materialized_text_bytes,
                    limits.max_total_uncompressed_bytes,
                )?;
            }
        }
        if objects.len() > start {
            let height = note_y - layout.margin_top;
            let top = layout.height - layout.margin_bottom - height;
            if note_page != page
                || objects[..start]
                    .iter()
                    .any(|o| o.unit_index == page && o.bounds.y + o.bounds.height > top)
            {
                diagnostics.push(Diagnostic::warning(DiagnosticCode::ApproximateLayout, Phase::Layout, Fidelity::Approximate,
                    "ODT footnotes exceed the remaining page area; body reflow or note continuation is approximate").in_part(CONTENT_PART));
            }
            for object in &mut objects[start..] {
                object.bounds.y += top - layout.margin_top;
            }
            if objects.len() >= limits.max_document_objects {
                return Err(object_limit_error(
                    "ODT note separator exceeds object limit",
                ));
            }
            let mut separator = objects[start].clone();
            separator.numeric_id = objects.len() as u32;
            separator.stable_id = format!("odt:footnote-separator:{page}");
            separator.kind = ObjectKind::Shape;
            separator.text = None;
            separator.bounds = Rect {
                x: layout.margin_left,
                y: top - 8.0,
                width: (layout.width - layout.margin_left - layout.margin_right) / 4.0,
                height: 0.0,
            };
            separator.visual = Visual::PaintedShape {
                geometry: Geometry::Line,
                fill: Paint::None,
                stroke: Paint::Solid(0x000000ff),
                stroke_width: 0.5,
            };
            objects.push(separator);
        }
    }
    for note in completed_notes.iter_mut().filter(|note| note.endnote) {
        for paragraph in std::mem::take(&mut note.paragraphs) {
            push_paragraph(
                paragraph,
                layout,
                &mut units,
                &mut unit_index,
                &mut y,
                &mut objects,
                styles,
                limits.max_document_objects,
                &mut materialized_text_bytes,
                limits.max_total_uncompressed_bytes,
            )?;
        }
    }

    let page_count = units.len() as u32;
    if story.is_none() {
        resolve_odf_page_fields(&mut objects, page_count, &page_number_starts);
    }
    Ok(ParsedContent {
        materialized_image_bytes,
        units,
        objects,
        diagnostics,
    })
}

fn deleted_change_ids(bytes: &[u8], limits: Limits) -> Result<HashSet<String>, Diagnostic> {
    let mut depth = 0_usize;
    let mut changed_region = None;
    let mut deleted = HashSet::new();
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                match local_name(name) {
                    "changed-region" if changed_region.is_none() => {
                        if let Some(id) = optional_attribute(&attributes, "id", CONTENT_PART)? {
                            changed_region = (!empty).then_some((depth, id));
                        }
                    }
                    "deletion" => {
                        if let Some((_, id)) = changed_region.as_ref() {
                            deleted.insert(id.clone());
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
                if local_name(name) == "changed-region"
                    && changed_region
                        .as_ref()
                        .is_some_and(|(region_depth, _)| *region_depth == depth)
                {
                    changed_region = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, CONTENT_PART))?;
    Ok(deleted)
}

fn prepare_master_tables(
    bytes: &[u8],
    layout: PageLayout,
    limits: Limits,
    styles: &mut StyleCatalog,
    package: Option<&Package<'_>>,
) -> Result<Vec<Diagnostic>, Diagnostic> {
    let mut diagnostics = Vec::new();
    for name in styles.master_pages.keys().cloned().collect::<Vec<_>>() {
        for kind in MASTER_STORY_KINDS {
            if !master_story_ref(&styles.master_pages[&name], kind).has_table {
                continue;
            }
            let content =
                parse_story_xml(bytes, layout, limits, styles, package, Some((&name, kind)))?;
            if content.objects.iter().any(|o| o.kind == ObjectKind::Table) {
                diagnostics.extend(content.diagnostics);
                let slot = master_story_slot(styles.master_pages.get_mut(&name).unwrap(), kind);
                slot.table_objects = content.objects;
                slot.paragraphs.clear();
                slot.shapes.clear();
            }
        }
    }
    Ok(diagnostics)
}

fn push_master_stories(
    unit_index: u32,
    layout: PageLayout,
    objects: &mut Vec<Object>,
    styles: &StyleCatalog,
    object_limit: usize,
) -> Result<f32, Diagnostic> {
    let Some((master_name, header_kind, header, footer_kind, footer)) =
        styles.master_stories(unit_index)
    else {
        return Ok(layout.margin_top);
    };
    let mut body_top = layout.margin_top;
    for (kind, is_header, content) in [
        (header_kind, true, &header.table_objects),
        (footer_kind, false, &footer.table_objects),
    ] {
        if content.is_empty() {
            continue;
        }
        if objects.len().saturating_add(content.len()) > object_limit {
            return Err(object_limit_error("ODT master table objects exceed limit"));
        }
        let top = content
            .iter()
            .map(|o| o.bounds.y)
            .fold(f32::INFINITY, f32::min);
        let bottom = content
            .iter()
            .map(|o| o.bounds.y + o.bounds.height)
            .fold(0.0, f32::max);
        let offset = if is_header {
            -top
        } else {
            layout.height - layout.margin_bottom - bottom
        };
        let base = objects.len() as u32;
        for template in content {
            let mut object = template.clone();
            object.numeric_id += base;
            object.parent_numeric_id = object.parent_numeric_id.map(|id| id + base);
            object.stable_id = format!("odt:{kind}:{unit_index}:{}", object.numeric_id);
            object.parent_stable_id = object
                .parent_numeric_id
                .map(|id| format!("odt:{kind}:{unit_index}:{id}"));
            object.unit_index = unit_index;
            object.bounds.y += offset;
            object.source.part = STYLES_PART.to_owned();
            objects.push(object);
        }
        if is_header {
            body_top = body_top.max(bottom + offset);
        }
    }

    for (kind, is_header, story) in [(header_kind, true, header), (footer_kind, false, footer)] {
        let mut x = layout.margin_left;
        let mut width = layout.width - layout.margin_left - layout.margin_right;
        let mut y = if is_header {
            0.0
        } else {
            layout.height - layout.margin_bottom
        };
        let mut available_height = if is_header {
            layout.margin_top
        } else {
            layout.margin_bottom
        };
        let box_style = styles
            .master_pages
            .get(master_name)
            .and_then(|master| master.page_layout.as_ref())
            .and_then(|name| styles.story_styles.get(name))
            .and_then(|pair| pair[usize::from(!is_header)].as_ref());
        if let Some(style) = box_style.filter(|_| story.present) {
            let natural_height: f32 = story
                .paragraphs
                .iter()
                .map(|p| {
                    effective_odt_line_height(&p.style, &p.runs, styles.font_metrics)
                        * p.text.split('\n').count().max(1) as f32
                })
                .sum();
            let inset_top = style.padding[0] + style.border.stroke_width;
            let inset_bottom = style.padding[2] + style.border.stroke_width;
            // ODF fixed height includes the story's spacing to the body.
            let height = (style.height.max(style.min_height) - style.margin[0] - style.margin[2])
                .max(natural_height + inset_top + inset_bottom);
            x += style.margin[3];
            width = (width - style.margin[1] - style.margin[3]).max(1.0);
            y = if is_header {
                layout.margin_top + style.margin[0]
            } else {
                layout.height - layout.margin_bottom - style.margin[2] - height
            };
            if objects.len() >= object_limit {
                return Err(object_limit_error("ODT story box exceeds object limit"));
            }
            let numeric_id = objects.len() as u32;
            let visual = Visual::Shape {
                geometry: Geometry::Rectangle,
                fill: style.border.fill,
                stroke: style.border.stroke,
                stroke_width: style.border.stroke_width,
            };
            let visual = if style.border.stroke_style == StrokeStyle::default() {
                visual
            } else {
                Visual::StrokeStyle {
                    style: style.border.stroke_style.clone(),
                    visual: Box::new(visual),
                }
            };
            objects.push(Object {
                numeric_id, parent_numeric_id: None,
                stable_id: format!("odt:{}-box:{unit_index}", if is_header { "header" } else { "footer" }),
                parent_stable_id: None, kind: ObjectKind::Shape, unit_index,
                bounds: Rect { x, y, width, height }, z: numeric_id as i32, text: None,
                source: SourceRef { part: STYLES_PART.to_owned(), mapping: MappingQuality::Derived,
                    locator: SourceLocator::Odt { kind: "element", element_id: None,
                        path: format!("/office:document-styles/office:automatic-styles/style:page-layout[@style:name='{}']/style:{}-style/style:header-footer-properties",
                            styles.master_pages[master_name].page_layout.as_deref().unwrap_or(""), if is_header { "header" } else { "footer" }),
                        row: None, column: None, text_range: None } }, visual,
            });
            if is_header {
                body_top = body_top.max(y + height + style.margin[2]);
            }
            x += style.padding[3] + style.border.stroke_width;
            width = (width - style.padding[1] - style.padding[3] - 2.0 * style.border.stroke_width)
                .max(1.0);
            y += inset_top;
            available_height = (height - inset_top - inset_bottom).max(1.0);
        }
        for (index, paragraph) in story.paragraphs.iter().enumerate() {
            if objects.len() >= object_limit {
                return Err(object_limit_error(
                    "ODT master-page story exceeds the configured object limit",
                ));
            }
            let line_height =
                effective_odt_line_height(&paragraph.style, &paragraph.runs, styles.font_metrics);
            let line_count = paragraph.text.split('\n').count().max(1);
            let natural_height = line_count as f32 * line_height;
            let height = natural_height.min(available_height.max(line_height));
            if !is_header && available_height <= 0.0 {
                y = (layout.height - height).max(0.0);
            }
            let numeric_id = u32::try_from(objects.len()).map_err(|_| {
                object_limit_error("ODT story object count exceeds supported range")
            })?;
            let text = paragraph.text.clone();
            let text_length = u32::try_from(text.chars().count()).unwrap_or(u32::MAX);
            let mut runs = paragraph.runs.clone();
            if runs.is_empty() && !text.is_empty() {
                let text_style = &paragraph.style.text;
                runs.push(TextRun {
                    paint: None,
                    east_asian_line_breaks: true,
                    text: text.clone(),
                    font_family: text_style.font_family.clone(),
                    font_size: text_style.font_size,
                    color: text_style.color,
                    bold: text_style.bold,
                    italic: text_style.italic,
                    underline: text_style.underline,
                    strikethrough: text_style.strikethrough,
                    highlight: text_style.highlight,
                    baseline_shift: text_style.baseline_shift,
                    letter_spacing: text_style.letter_spacing,
                    horizontal_scale: 1.0,
                });
            }
            objects.push(Object {
                numeric_id,
                parent_numeric_id: None,
                stable_id: format!("odt:{kind}:{unit_index}:{index}"),
                parent_stable_id: None,
                kind: ObjectKind::Paragraph,
                unit_index,
                bounds: Rect {
                    x,
                    y,
                    width,
                    height,
                },
                z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
                text: Some(text),
                source: SourceRef {
                    part: STYLES_PART.to_owned(),
                    mapping: MappingQuality::Derived,
                    locator: SourceLocator::Odt {
                        kind: "text-range",
                        element_id: None,
                        path: format!(
                            "/office:document-styles/office:master-styles/style:master-page/style:{kind}/text:p[{}]",
                            index + 1
                        ),
                        row: None,
                        column: None,
                        text_range: Some((0, text_length)),
                    },
                },
                visual: Visual::RichText {
                    geometry: Geometry::Rectangle,
                    fill: Paint::None,
                    stroke: Paint::None,
                    stroke_width: 0.0,
                    align: paragraph.style.align,
                    line_height,
                    runs,
                },
            });
            y += height;
        }
        if is_header {
            body_top = body_top.max(y);
        }
    }
    for (kind, is_header, story) in [(header_kind, true, header), (footer_kind, false, footer)] {
        let mut shapes = story.shapes.clone();
        let story_bottom = layout_master_shapes(&mut shapes, layout);
        finish_shapes_from(
            shapes,
            0.0,
            unit_index,
            objects,
            object_limit,
            ShapeSource::Master {
                master_name,
                story_kind: kind,
            },
        )?;
        if is_header {
            body_top = body_top.max(story_bottom);
        }
    }
    Ok(body_top)
}

fn layout_master_shapes(shapes: &mut [ShapeState], layout: PageLayout) -> f32 {
    for index in 0..shapes.len() {
        if shapes[index].graphic.horizontal_rel.as_deref() == Some("page-content") {
            shapes[index].bounds.x += layout.margin_left;
            if let Some(connector) = shapes[index].connector.as_mut() {
                connector.start.0 += layout.margin_left;
                connector.end.0 += layout.margin_left;
            }
        }
        if !shapes[index].graphic.allow_overlap {
            let mut y = shapes[index].bounds.y;
            for previous in &shapes[..index] {
                let horizontal_overlap = shapes[index].bounds.x
                    < previous.bounds.x + previous.bounds.width
                    && shapes[index].bounds.x + shapes[index].bounds.width > previous.bounds.x;
                let vertical_overlap = y < previous.bounds.y + previous.bounds.height
                    && y + shapes[index].bounds.height > previous.bounds.y;
                if horizontal_overlap && vertical_overlap {
                    y = y.max(previous.bounds.y + previous.bounds.height);
                }
            }
            let offset = y - shapes[index].bounds.y;
            shapes[index].bounds.y = y;
            if let Some(connector) = shapes[index].connector.as_mut() {
                connector.start.1 += offset;
                connector.end.1 += offset;
            }
        }
    }
    shapes
        .iter()
        .map(|shape| shape.bounds.y + shape.bounds.height)
        .fold(layout.margin_top, f32::max)
}

#[allow(clippy::too_many_arguments)]
fn finish_frame(
    frame: FrameState,
    layout: PageLayout,
    units: &mut Vec<Unit>,
    unit_index: &mut u32,
    y: &mut f32,
    objects: &mut Vec<Object>,
    diagnostics: &mut Vec<Diagnostic>,
    package: Option<&Package<'_>>,
    limits: Limits,
    styles: &StyleCatalog,
    materialized_image_bytes: &mut usize,
    parent: Option<(u32, &str, Rect)>,
) -> Result<(), Diagnostic> {
    if objects.len() >= limits.max_document_objects {
        return Err(object_limit_error(
            "ODT frame count exceeds the configured object limit",
        ));
    }
    if (frame.auto_width || frame.auto_height) && frame.children.is_empty() {
        diagnostics.push(Diagnostic::warning(DiagnosticCode::ApproximateLayout, Phase::Layout,
            Fidelity::Approximate, "ODT frame has no fixed dimensions or measurable child frames; a 1-inch mapped placeholder is used").in_part(CONTENT_PART));
    }
    let flow_positioned = parent.is_none()
        && !frame.page_anchored
        && frame
            .graphic
            .vertical_rel
            .as_deref()
            .is_none_or(|value| !matches!(value, "page" | "page-content"));
    let moves_page = flow_positioned
        && (frame.inline || frame.wrap != OdtWrap::RunThrough)
        && *y + frame.y.unwrap_or(0.0).max(0.0) + frame.height
            > layout.height - layout.margin_bottom
        && *y > layout.margin_top;
    if moves_page {
        *unit_index = unit_index
            .checked_add(1)
            .ok_or_else(|| format_error(CONTENT_PART, "page count exceeds supported range"))?;
        units.push(page_unit(*unit_index, layout));
        *y = push_master_stories(
            *unit_index,
            layout,
            objects,
            styles,
            limits.max_document_objects,
        )?;
    }
    let content_bounds = Rect {
        x: layout.margin_left,
        y: layout.margin_top,
        width: layout.width - layout.margin_left - layout.margin_right,
        height: layout.height - layout.margin_top - layout.margin_bottom,
    };
    let page_bounds = Rect {
        x: 0.0,
        y: 0.0,
        width: layout.width,
        height: layout.height,
    };
    let parent_bounds = parent.map(|(_, _, bounds)| bounds);
    let horizontal_reference = parent_bounds.unwrap_or_else(|| {
        if frame.graphic.horizontal_rel.as_deref() == Some("page") {
            page_bounds
        } else {
            content_bounds
        }
    });
    let vertical_reference =
        parent_bounds.unwrap_or_else(|| match frame.graphic.vertical_rel.as_deref() {
            Some("page") => page_bounds,
            Some("page-content") => content_bounds,
            Some("text" | "line") if frame.inline => Rect {
                y: *y,
                height: frame.height.max(frame.text_style.font_size * 1.5),
                ..content_bounds
            },
            _ if frame.page_anchored => content_bounds,
            _ => Rect {
                y: *y,
                ..content_bounds
            },
        });
    let width = frame.width.min(layout.width).max(1.0);
    let height = frame.height.min(layout.height).max(1.0);
    let bounds = Rect {
        x: match frame.graphic.horizontal_pos {
            TextAlign::End => horizontal_reference.x + horizontal_reference.width - width,
            TextAlign::Center => {
                horizontal_reference.x + (horizontal_reference.width - width) / 2.0
            }
            _ => horizontal_reference.x + frame.x.unwrap_or(0.0),
        },
        y: match frame.graphic.vertical_pos {
            TextAlign::End => vertical_reference.y + vertical_reference.height - height,
            TextAlign::Center => vertical_reference.y + (vertical_reference.height - height) / 2.0,
            _ => vertical_reference.y + frame.y.unwrap_or(0.0),
        },
        width,
        height,
    };
    let mapping = if frame.element_id.is_some() {
        MappingQuality::Exact
    } else {
        MappingQuality::Derived
    };
    let path = format!(
        "/office:document-content/office:body/office:text/draw:{}[{}]",
        frame.source_element,
        frame.index + 1
    );
    let mut image_visual = None;
    if frame.chart.is_none() && frame.math.is_none() && !frame.image_hrefs.is_empty() {
        if let Some(package) = package {
            for href in &frame.image_hrefs {
                let target = match resolve_odf_image_target(href, limits.max_zip_path_bytes) {
                    Ok(OdfImageTarget::Embedded(target)) => target,
                    Ok(OdfImageTarget::External) => {
                        diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::ExternalResourceBlocked,
                                Phase::Security,
                                Fidelity::Blocked,
                                "external ODT frame image was blocked",
                            )
                            .in_part(CONTENT_PART),
                        );
                        continue;
                    }
                    Err(message) => {
                        diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::UnsupportedFeature,
                                Phase::Parse,
                                Fidelity::Approximate,
                                format!("invalid ODT frame image reference: {message}"),
                            )
                            .in_part(CONTENT_PART),
                        );
                        continue;
                    }
                };
                let Some(bytes) = package.part(&target)? else {
                    diagnostics.push(
                        Diagnostic::warning(
                            DiagnosticCode::UnsupportedFeature,
                            Phase::Container,
                            Fidelity::Omitted,
                            "ODT frame image alternative is missing from the package",
                        )
                        .in_part(&target),
                    );
                    continue;
                };
                let declared = office_image_media_type(&target, &bytes);
                let media_type = match super::presentation_image::recover_office_image_signature(
                    declared, &bytes,
                ) {
                    Ok(media_type) => media_type,
                    Err(error) => {
                        let reason = match error {
                            OfficeImageError::UnsupportedFormat => "unsupported image format",
                            OfficeImageError::SignatureMismatch => "image signature mismatch",
                            OfficeImageError::DisabledByOffice => "format disabled by Office",
                        };
                        diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::UnsupportedFeature,
                                Phase::Render,
                                Fidelity::Approximate,
                                format!("ODT frame image alternative uses an {reason}"),
                            )
                            .in_part(&target),
                        );
                        continue;
                    }
                };
                if declared == Err(OfficeImageError::SignatureMismatch) {
                    diagnostics.push(Diagnostic::warning(DiagnosticCode::UnsupportedFeature,
                        Phase::Parse, Fidelity::Approximate,
                        format!("ODT image name disagrees with its signature; recovered as {media_type}"))
                        .in_part(&target));
                }
                reserve_materialized_image_bytes(
                    materialized_image_bytes,
                    bytes.len(),
                    limits.max_total_uncompressed_bytes,
                    &target,
                )?;
                image_visual = Some(Visual::Image {
                    media_type: media_type.to_owned(),
                    bytes: clone_image_bytes(&bytes, &target)?,
                    crop: ImageCrop::default(),
                });
                break;
            }
            if image_visual.is_none() {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Render,
                        Fidelity::Approximate,
                        "no supported ODT frame image alternative; a mapped placeholder is rendered",
                    )
                    .in_part(CONTENT_PART),
                );
            }
        } else {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Render,
                    Fidelity::Approximate,
                    "ODT frame image bytes are unavailable in parser-only mode; a mapped placeholder is rendered",
                )
                .in_part(CONTENT_PART),
            );
        }
    }
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| object_limit_error("ODT frame object count exceeds supported range"))?;
    let stable_id = format!("odt:frame:{}:visual", frame.index);
    let has_text = !frame.text.is_empty();
    let primary_kind = if frame.chart.is_some() {
        ObjectKind::Group
    } else if frame.math.is_some() {
        ObjectKind::TextBox
    } else if !frame.image_hrefs.is_empty() {
        ObjectKind::Image
    } else if !frame.children.is_empty() {
        ObjectKind::Group
    } else if has_text {
        ObjectKind::TextBox
    } else {
        ObjectKind::Shape
    };
    let primary_visual = transformed_visual(
        image_visual.unwrap_or(Visual::PaintedShape {
            geometry: Geometry::Rectangle,
            fill: if frame.math.is_some() {
                Paint::None
            } else if frame.chart.is_some() {
                Paint::Solid(0xffff_ffff)
            } else {
                frame.graphic.fill.clone()
            },
            stroke: if frame.math.is_some() {
                Paint::None
            } else if frame.chart.is_some() {
                Paint::Solid(0xd1d5_dbff)
            } else {
                frame.graphic.stroke.clone()
            },
            stroke_width: if frame.chart.is_some() {
                1.0
            } else {
                frame.graphic.stroke_width
            },
        }),
        frame.transform,
    );
    objects.push(Object {
        numeric_id,
        parent_numeric_id: parent.map(|(id, _, _)| id),
        stable_id: stable_id.clone(),
        parent_stable_id: parent.map(|(_, id, _)| id.to_owned()),
        kind: primary_kind,
        unit_index: *unit_index,
        bounds,
        z: frame
            .z
            .unwrap_or_else(|| i32::try_from(numeric_id).unwrap_or(i32::MAX)),
        text: None,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping,
            locator: SourceLocator::Odt {
                kind: "element",
                element_id: frame.element_id.clone(),
                path: path.clone(),
                row: None,
                column: None,
                text_range: None,
            },
        },
        visual: primary_visual,
    });
    if let Some(chart) = frame.chart.as_ref() {
        push_odt_chart(
            chart,
            frame.index,
            frame.element_id.as_deref(),
            numeric_id,
            &stable_id,
            *unit_index,
            bounds,
            objects,
            limits.max_document_objects,
        )?;
    }
    if has_text {
        if objects.len() >= limits.max_document_objects {
            return Err(object_limit_error(
                "ODT frame text exceeds the configured object limit",
            ));
        }
        let text_numeric_id = u32::try_from(objects.len())
            .map_err(|_| object_limit_error("ODT frame object count exceeds supported range"))?;
        let text_length = u32::try_from(frame.text.chars().count()).unwrap_or(u32::MAX);
        objects.push(Object {
            numeric_id: text_numeric_id,
            parent_numeric_id: Some(numeric_id),
            stable_id: format!("odt:frame:{}:text", frame.index),
            parent_stable_id: Some(stable_id.clone()),
            kind: ObjectKind::TextBox,
            unit_index: *unit_index,
            bounds,
            z: frame
                .z
                .unwrap_or_else(|| i32::try_from(text_numeric_id).unwrap_or(i32::MAX)),
            text: Some(frame.text.clone()),
            source: SourceRef {
                part: frame
                    .math
                    .as_ref()
                    .map_or_else(|| CONTENT_PART.to_owned(), |math| math.source_part.clone()),
                mapping,
                locator: SourceLocator::Odt {
                    kind: "text-range",
                    element_id: frame.element_id,
                    path: if frame.math.is_some() {
                        "/math:math[1]".to_owned()
                    } else {
                        path
                    },
                    row: None,
                    column: None,
                    text_range: Some((0, text_length)),
                },
            },
            visual: Visual::RichText {
                geometry: Geometry::Rectangle,
                fill: Paint::None,
                stroke: Paint::None,
                stroke_width: 0.0,
                align: TextAlign::Start,
                line_height: if frame.math.is_some() {
                    0.0
                } else {
                    LINE_HEIGHT
                },
                runs: vec![TextRun {
                    paint: None,
                    east_asian_line_breaks: true,
                    text: frame.text,
                    font_family: if frame.math.is_some() {
                        "Cambria Math"
                    } else {
                        "Arial"
                    }
                    .to_owned(),
                    font_size: FONT_SIZE,
                    color: 0x0000_00ff,
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
    }
    for child in frame.children {
        let mut child_y = bounds.y;
        finish_frame(
            child,
            layout,
            units,
            unit_index,
            &mut child_y,
            objects,
            diagnostics,
            package,
            limits,
            styles,
            materialized_image_bytes,
            Some((numeric_id, &stable_id, bounds)),
        )?;
    }
    if frame.wrap == OdtWrap::Parallel {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ApproximateLayout,
                Phase::Layout,
                Fidelity::Approximate,
                "ODT parallel frame wrap uses a deterministic top-and-bottom exclusion in static layout",
            )
            .in_part(CONTENT_PART),
        );
    }
    // Explicitly positioned floats can be accompanied by authored blank paragraphs.
    // Advance flow only for inline/unpositioned objects or a float moved to a new page.
    if flow_positioned
        && (moves_page || frame.inline || (frame.y.is_none() && frame.wrap != OdtWrap::RunThrough))
    {
        *y = (*y).max(bounds.y + bounds.height);
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_odt_chart(
    chart: &OdpBasicChart,
    frame_index: u32,
    element_id: Option<&str>,
    parent_numeric_id: u32,
    parent_stable_id: &str,
    unit_index: u32,
    bounds: Rect,
    objects: &mut Vec<Object>,
    object_limit: usize,
) -> Result<(), Diagnostic> {
    if chart.kind == OdpChartKind::Pie {
        return push_odt_pie_chart(
            chart,
            frame_index,
            element_id,
            parent_numeric_id,
            parent_stable_id,
            unit_index,
            bounds,
            objects,
            object_limit,
        );
    }
    let category_count = chart
        .series
        .iter()
        .map(|series| series.values.len())
        .max()
        .unwrap_or(0);
    let (minimum, maximum) = odp_chart_value_range(chart);
    if category_count == 0
        || (chart.kind != OdpChartKind::Bar && category_count < 2)
        || maximum <= 0.0
    {
        return Ok(());
    }
    let (axis_minimum, axis_maximum, tick_step) = odp_chart_axis(minimum, maximum);
    let tick_count = (((axis_maximum - axis_minimum) / tick_step).round() as usize).min(12);
    let point_count = chart
        .series
        .iter()
        .map(|series| {
            series
                .values
                .iter()
                .filter(|value| value.is_finite() && **value > 0.0)
                .count()
        })
        .sum::<usize>();
    let category_label_count = chart
        .categories
        .iter()
        .take(category_count)
        .filter(|category| !category.is_empty())
        .count();
    let labeled_series_count = chart
        .series
        .iter()
        .filter(|series| series.label.is_some())
        .count();
    let decoration_count = 2
        + tick_count
        + tick_count
        + 1
        + category_label_count
        + labeled_series_count * 2
        + usize::from(chart.title.is_some());
    if objects
        .len()
        .checked_add(point_count)
        .and_then(|required| required.checked_add(decoration_count))
        .is_none_or(|required| required > object_limit)
    {
        return Err(object_limit_error(
            "ODT chart objects exceed the configured object limit",
        ));
    }
    let plot = odp_chart_plot(chart, bounds);
    push_odt_chart_title(
        chart,
        frame_index,
        element_id,
        parent_numeric_id,
        parent_stable_id,
        unit_index,
        bounds,
        objects,
    )?;
    if let Some(wall) = chart.wall {
        push_odt_chart_shape(
            chart,
            frame_index,
            element_id,
            parent_numeric_id,
            parent_stable_id,
            unit_index,
            plot,
            Geometry::Rectangle,
            odp_chart_wall_fill(wall, plot),
            Paint::None,
            0.0,
            Some("chart:plot-area[1]/chart:wall[1]"),
            None,
            None,
            MappingQuality::Exact,
            objects,
        )?;
    }
    for (index, axis_bounds) in odp_chart_axis_bounds(chart, plot).into_iter().enumerate() {
        if index == 0 && !chart.value_axis_visible {
            continue;
        }
        push_odt_chart_shape(
            chart,
            frame_index,
            element_id,
            parent_numeric_id,
            parent_stable_id,
            unit_index,
            axis_bounds,
            Geometry::Line,
            Paint::None,
            Paint::Solid(0x6b72_80ff),
            1.0,
            None,
            None,
            None,
            MappingQuality::Derived,
            objects,
        )?;
    }
    let font_size = (bounds.height * 0.035).clamp(10.0, 14.0);
    let text_height = font_size * 1.4;
    for (category_index, category) in chart.categories.iter().take(category_count).enumerate() {
        if category.is_empty() {
            continue;
        }
        push_odt_chart_text(
            chart,
            frame_index,
            element_id,
            parent_numeric_id,
            parent_stable_id,
            unit_index,
            odp_chart_category_label_bounds(
                chart,
                plot,
                category_count,
                category_index,
                font_size,
                text_height,
            ),
            category,
            font_size,
            TextAlign::Center,
            "chart:plot-area[1]/chart:axis[1]/chart:categories[1]",
            None,
            Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
            MappingQuality::Exact,
            objects,
        )?;
    }
    for tick_index in 0..=tick_count {
        let value = axis_minimum + tick_step * tick_index as f32;
        let tick_y = plot.y + plot.height
            - plot.height * (value - axis_minimum) / (axis_maximum - axis_minimum);
        if tick_index != 0 {
            push_odt_chart_shape(
                chart,
                frame_index,
                element_id,
                parent_numeric_id,
                parent_stable_id,
                unit_index,
                Rect {
                    x: plot.x,
                    y: tick_y,
                    width: plot.width,
                    height: 0.01,
                },
                Geometry::Line,
                Paint::None,
                Paint::Solid(0xb3b3_b3ff),
                1.0,
                None,
                None,
                Some(u32::try_from(tick_index).unwrap_or(u32::MAX)),
                MappingQuality::Derived,
                objects,
            )?;
        }
        let (label_bounds, align) =
            odp_chart_value_label_bounds(chart, bounds, plot, font_size, text_height, tick_y);
        push_odt_chart_text(
            chart,
            frame_index,
            element_id,
            parent_numeric_id,
            parent_stable_id,
            unit_index,
            label_bounds,
            &format_odp_chart_number(value),
            font_size,
            align,
            "chart:plot-area[1]/chart:axis[2]",
            Some(u32::try_from(tick_index).unwrap_or(u32::MAX)),
            None,
            MappingQuality::Derived,
            objects,
        )?;
    }
    let legend_row_height = font_size * 1.55;
    let legend_y = plot.y + (plot.height - labeled_series_count as f32 * legend_row_height) / 2.0;
    for (legend_index, (series_index, series)) in chart
        .series
        .iter()
        .enumerate()
        .filter(|(_, series)| series.label.is_some())
        .enumerate()
    {
        let label = series.label.as_deref().unwrap_or_default();
        let swatch_size = font_size * 0.72;
        let row_y = legend_y + legend_index as f32 * legend_row_height;
        let swatch_x = plot.x + plot.width + bounds.width * 0.03;
        push_odt_chart_shape(
            chart,
            frame_index,
            element_id,
            parent_numeric_id,
            parent_stable_id,
            unit_index,
            Rect {
                x: swatch_x,
                y: row_y + (legend_row_height - swatch_size) / 2.0,
                width: swatch_size,
                height: swatch_size,
            },
            Geometry::Rectangle,
            Paint::Solid(
                series
                    .color
                    .unwrap_or_else(|| odt_chart_palette_color(series_index)),
            ),
            Paint::None,
            0.0,
            None,
            Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
            None,
            MappingQuality::Derived,
            objects,
        )?;
        let text_x = swatch_x + swatch_size + font_size * 0.35;
        push_odt_chart_text(
            chart,
            frame_index,
            element_id,
            parent_numeric_id,
            parent_stable_id,
            unit_index,
            Rect {
                x: text_x,
                y: row_y,
                width: (bounds.x + bounds.width - text_x).max(font_size),
                height: legend_row_height,
            },
            label,
            font_size,
            TextAlign::Start,
            "chart:legend[1]",
            Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
            None,
            MappingQuality::Exact,
            objects,
        )?;
    }
    if chart.kind == OdpChartKind::Bar {
        for (series_index, category_index, bar_bounds) in
            odp_chart_bar_bounds(chart, plot, axis_minimum, axis_maximum)
        {
            push_odt_chart_shape(
                chart,
                frame_index,
                element_id,
                parent_numeric_id,
                parent_stable_id,
                unit_index,
                bar_bounds,
                Geometry::Rectangle,
                Paint::Solid(
                    chart.series[series_index]
                        .color
                        .unwrap_or_else(|| odt_chart_palette_color(series_index)),
                ),
                Paint::None,
                0.0,
                None,
                Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                MappingQuality::Exact,
                objects,
            )?;
        }
        return Ok(());
    }
    let domain_bounds = (chart.kind == OdpChartKind::Scatter).then(|| {
        chart
            .series
            .iter()
            .flat_map(|series| series.domains.iter().copied())
            .filter(|value| value.is_finite())
            .fold((f32::INFINITY, f32::NEG_INFINITY), |(min, max), value| {
                (min.min(value), max.max(value))
            })
    });
    for (series_index, series) in chart.series.iter().enumerate() {
        let points = series
            .values
            .iter()
            .enumerate()
            .filter(|(_, value)| value.is_finite() && **value >= 0.0)
            .map(|(index, value)| {
                let x_ratio = domain_bounds
                    .and_then(|(min, max)| {
                        series
                            .domains
                            .get(index)
                            .filter(|domain| max > min && domain.is_finite())
                            .map(|domain| ((*domain - min) / (max - min)).clamp(0.0, 1.0))
                    })
                    .unwrap_or(index as f32 / (category_count - 1) as f32);
                (
                    plot.x + x_ratio * plot.width,
                    plot.y + plot.height
                        - plot.height
                            * ((*value - axis_minimum) / (axis_maximum - axis_minimum))
                                .clamp(0.0, 1.0),
                    index,
                )
            })
            .collect::<Vec<_>>();
        for segment in points.windows(2) {
            let (x0, y0, point_index) = segment[0];
            let (x1, y1, _) = segment[1];
            let segment_bounds = Rect {
                x: x0.min(x1),
                y: y0.min(y1),
                width: (x1 - x0).abs(),
                height: (y1 - y0).abs(),
            };
            push_odt_chart_shape(
                chart,
                frame_index,
                element_id,
                parent_numeric_id,
                parent_stable_id,
                unit_index,
                segment_bounds,
                Geometry::Path {
                    fill_rule: FillRule::NonZero,
                    commands: super::odf_chart::line_segment_commands(
                        (x0 - segment_bounds.x, y0 - segment_bounds.y),
                        (x1 - segment_bounds.x, y1 - segment_bounds.y),
                        None,
                    ),
                },
                Paint::None,
                Paint::Solid(
                    series
                        .color
                        .unwrap_or_else(|| odt_chart_palette_color(series_index)),
                ),
                2.0,
                None,
                Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                Some(u32::try_from(point_index).unwrap_or(u32::MAX)),
                MappingQuality::Exact,
                objects,
            )?;
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_odt_pie_chart(
    chart: &OdpBasicChart,
    frame_index: u32,
    element_id: Option<&str>,
    parent_numeric_id: u32,
    parent_stable_id: &str,
    unit_index: u32,
    bounds: Rect,
    objects: &mut Vec<Object>,
    object_limit: usize,
) -> Result<(), Diagnostic> {
    let Some(series) = chart.series.first() else {
        return Ok(());
    };
    let total = series
        .values
        .iter()
        .copied()
        .filter(|value| value.is_finite() && *value > 0.0)
        .sum::<f32>();
    let point_count = series
        .values
        .iter()
        .filter(|value| value.is_finite() && **value > 0.0)
        .count();
    if total <= 0.0 {
        return Ok(());
    }
    if objects
        .len()
        .saturating_add(point_count)
        .saturating_add(1)
        .saturating_add(usize::from(chart.title.is_some()))
        > object_limit
    {
        return Err(object_limit_error(
            "ODT pie chart objects exceed the configured object limit",
        ));
    }
    let has_title = chart.title.is_some();
    let plot = Rect {
        x: bounds.x + bounds.width * 0.08,
        y: bounds.y + bounds.height * if has_title { 0.18 } else { 0.06 },
        width: bounds.width * 0.86,
        height: bounds.height * if has_title { 0.66 } else { 0.78 },
    };
    push_odt_chart_title(
        chart,
        frame_index,
        element_id,
        parent_numeric_id,
        parent_stable_id,
        unit_index,
        bounds,
        objects,
    )?;
    if let Some(wall) = chart.wall {
        push_odt_chart_shape(
            chart,
            frame_index,
            element_id,
            parent_numeric_id,
            parent_stable_id,
            unit_index,
            plot,
            Geometry::Rectangle,
            odp_chart_wall_fill(wall, plot),
            Paint::None,
            0.0,
            Some("chart:plot-area[1]/chart:wall[1]"),
            None,
            None,
            MappingQuality::Exact,
            objects,
        )?;
    }
    let center_x = plot.x + plot.width / 2.0;
    let center_y = plot.y + plot.height / 2.0;
    let radius = plot.width.min(plot.height) * 0.42;
    let mut start_angle = -std::f32::consts::FRAC_PI_2;
    let labels = series
        .data_labels
        .as_ref()
        .filter(|labels| labels.is_enabled());
    for (category_index, value) in series.values.iter().copied().enumerate() {
        if !value.is_finite() || value <= 0.0 {
            continue;
        }
        let sweep = std::f32::consts::TAU * value / total;
        let mid_angle = start_angle + sweep / 2.0;
        let explosion = series
            .point_explosions
            .get(category_index)
            .copied()
            .unwrap_or(0.0);
        let (offset_x, offset_y) =
            crate::format::odf_chart::pie_explosion_offset(explosion, radius, mid_angle);
        let (segment_bounds, geometry) = odp_pie_segment_geometry(
            center_x + offset_x,
            center_y + offset_y,
            radius,
            start_angle,
            sweep,
        );
        push_odt_chart_shape(
            chart,
            frame_index,
            element_id,
            parent_numeric_id,
            parent_stable_id,
            unit_index,
            segment_bounds,
            geometry,
            Paint::Solid(odt_chart_palette_color(category_index)),
            Paint::Solid(0xffff_ffff),
            1.0,
            None,
            Some(0),
            Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
            MappingQuality::Exact,
            objects,
        )?;
        let category = chart
            .categories
            .get(category_index)
            .map(String::as_str)
            .unwrap_or_default();
        let custom = series
            .custom_labels
            .get(category_index)
            .and_then(Option::as_deref);
        let generated = labels.and_then(|labels| labels.format_label(category, value, total));
        if let Some(text) = custom.or(generated.as_deref()) {
            let font_size = labels.map_or(12.0, |labels| labels.font_size.max(8.0));
            let label_width = (radius * 0.9).max(font_size * 3.0);
            let label_height = font_size * 1.4;
            let (label_x, label_y) = if custom.is_some() {
                (
                    center_x + offset_x + radius * 0.62 * mid_angle.cos(),
                    center_y + offset_y + radius * 0.62 * mid_angle.sin(),
                )
            } else {
                crate::format::odf_chart::pie_label_anchor(
                    center_x,
                    center_y,
                    offset_x,
                    offset_y,
                    radius,
                    mid_angle,
                    labels
                        .and_then(|labels| labels.position)
                        .unwrap_or(crate::format::odf_chart::OdfLabelPosition::Outside),
                )
            };
            let source_path = format!(
                "chart:plot-area[1]/chart:series[1]/chart:data-point[{}]",
                category_index + 1
            );
            push_odt_chart_text(
                chart,
                frame_index,
                element_id,
                parent_numeric_id,
                parent_stable_id,
                unit_index,
                Rect {
                    x: label_x - label_width / 2.0,
                    y: label_y - label_height / 2.0,
                    width: label_width,
                    height: label_height,
                },
                text,
                font_size,
                TextAlign::Center,
                &source_path,
                Some(0),
                Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                MappingQuality::Exact,
                objects,
            )?;
        }
        start_angle += sweep;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_odt_chart_title(
    chart: &OdpBasicChart,
    frame_index: u32,
    element_id: Option<&str>,
    parent_numeric_id: u32,
    parent_stable_id: &str,
    unit_index: u32,
    bounds: Rect,
    objects: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    let Some(title) = chart.title.as_deref() else {
        return Ok(());
    };
    push_odt_chart_text(
        chart,
        frame_index,
        element_id,
        parent_numeric_id,
        parent_stable_id,
        unit_index,
        Rect {
            x: bounds.x + bounds.width * 0.08,
            y: bounds.y + bounds.height * 0.02,
            width: bounds.width * 0.84,
            height: bounds.height * 0.12,
        },
        title,
        (bounds.height * 0.055).clamp(14.0, 22.0),
        TextAlign::Center,
        "chart:title[1]",
        None,
        None,
        MappingQuality::Exact,
        objects,
    )
}

#[allow(clippy::too_many_arguments)]
fn push_odt_chart_text(
    chart: &OdpBasicChart,
    frame_index: u32,
    element_id: Option<&str>,
    parent_numeric_id: u32,
    parent_stable_id: &str,
    unit_index: u32,
    bounds: Rect,
    text: &str,
    font_size: f32,
    align: TextAlign,
    source_path: &str,
    row: Option<u32>,
    column: Option<u32>,
    mapping: MappingQuality,
    objects: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| object_limit_error("ODT chart object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: Some(parent_numeric_id),
        stable_id: format!("odt:frame:{frame_index}:chart:text:{numeric_id}"),
        parent_stable_id: Some(parent_stable_id.to_owned()),
        kind: ObjectKind::TextBox,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: Some(text.to_owned()),
        source: SourceRef {
            part: chart.source_part.clone(),
            mapping,
            locator: SourceLocator::Odt {
                kind: "element",
                element_id: element_id.map(str::to_owned),
                path: format!(
                    "/office:document-content/office:body/office:chart/chart:chart[1]/{source_path}"
                ),
                row,
                column,
                text_range: None,
            },
        },
        visual: Visual::RichText {
            geometry: Geometry::Rectangle,
            fill: Paint::None,
            stroke: Paint::None,
            stroke_width: 0.0,
            align,
            line_height: font_size * 1.2,
            runs: vec![TextRun {
                paint: None,
                east_asian_line_breaks: true,
                text: text.to_owned(),
                font_family: "Arial".to_owned(),
                font_size,
                color: 0x0000_00ff,
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

#[allow(clippy::too_many_arguments)]
fn push_odt_chart_shape(
    chart: &OdpBasicChart,
    frame_index: u32,
    element_id: Option<&str>,
    parent_numeric_id: u32,
    parent_stable_id: &str,
    unit_index: u32,
    bounds: Rect,
    geometry: Geometry,
    fill: Paint,
    stroke: Paint,
    stroke_width: f32,
    source_path: Option<&str>,
    row: Option<u32>,
    column: Option<u32>,
    mapping: MappingQuality,
    objects: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| object_limit_error("ODT chart object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: Some(parent_numeric_id),
        stable_id: format!("odt:frame:{frame_index}:chart:shape:{numeric_id}"),
        parent_stable_id: Some(parent_stable_id.to_owned()),
        kind: ObjectKind::Shape,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: chart.source_part.clone(),
            mapping,
            locator: SourceLocator::Odt {
                kind: "element",
                element_id: element_id.map(str::to_owned),
                path: format!(
                    "/office:document-content/office:body/office:chart/chart:chart[1]/{}",
                    source_path.map(str::to_owned).unwrap_or_else(|| format!(
                        "chart:plot-area[1]/chart:series[{}]/chart:data-point[{}]",
                        row.unwrap_or(0) + 1,
                        column.unwrap_or(0) + 1
                    ))
                ),
                row,
                column,
                text_range: None,
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

fn odt_chart_palette_color(index: usize) -> u32 {
    const PALETTE: [u32; 8] = [
        0x0045_86ff,
        0xff42_0eff,
        0xffd3_20ff,
        0x579d_1cff,
        0x7e00_21ff,
        0x83ca_ffff,
        0x3140_04ff,
        0xa5a5_a5ff,
    ];
    PALETTE[index % PALETTE.len()]
}

fn start_shape(
    local: &str,
    attributes: &[XmlAttribute<'_>],
    styles: &StyleCatalog,
    part: &str,
    depth: usize,
    index: u32,
) -> Result<ShapeState, Diagnostic> {
    let graphic = styles.graphic(optional_attribute(attributes, "style-name", part)?.as_deref());
    let connector = parse_odf_connector(local, attributes)?;
    let (bounds, geometry) = shape_bounds(local, attributes, part)?;
    let z = optional_z_index(attributes, part)?;
    Ok(ShapeState {
        depth,
        index,
        element: match local {
            "rect" => "rect",
            "ellipse" => "ellipse",
            "line" | "measure" => "line",
            "connector" => "connector",
            "regular-polygon" => "regular-polygon",
            "caption" => "caption",
            _ => "custom-shape",
        },
        element_id: optional_attribute(attributes, "id", part)?
            .or(optional_attribute(attributes, "name", part)?),
        geometry,
        connector,
        bounds,
        z,
        flow_anchored: matches!(
            optional_attribute(attributes, "anchor-type", part)?.as_deref(),
            Some("paragraph" | "char" | "as-char")
        ),
        transform: optional_attribute(attributes, "transform", part)?
            .map(|value| parse_odf_transform(&value, part))
            .transpose()?
            .unwrap_or_default(),
        graphic,
        text: String::new(),
        text_style: styles
            .paragraph(optional_attribute(attributes, "text-style-name", part)?.as_deref()),
        runs: Vec::new(),
        spans: Vec::new(),
        paragraph_depth: None,
    })
}

fn optional_z_index(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<Option<i32>, Diagnostic> {
    optional_attribute(attributes, "z-index", part)?
        .map(|value| {
            value
                .parse::<i32>()
                .ok()
                .filter(|value| *value >= 0)
                .ok_or_else(|| format_error(part, "ODT shape z-index is invalid"))
        })
        .transpose()
}

fn shape_bounds(
    kind: &str,
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<(Rect, Geometry), Diagnostic> {
    if matches!(kind, "line" | "connector" | "measure") {
        let x1 = optional_signed_length(attributes, "x1", part)?.unwrap_or(0.0);
        let y1 = optional_signed_length(attributes, "y1", part)?.unwrap_or(0.0);
        let x2 = optional_signed_length(attributes, "x2", part)?.unwrap_or(x1 + 1.0);
        let y2 = optional_signed_length(attributes, "y2", part)?.unwrap_or(y1 + 1.0);
        return Ok((
            Rect {
                x: x1.min(x2),
                y: y1.min(y2),
                width: (x2 - x1).abs().max(1.0),
                height: (y2 - y1).abs().max(1.0),
            },
            Geometry::Line,
        ));
    }
    let width = optional_length(attributes, "width", part)?.unwrap_or(CSS_PIXELS_PER_INCH);
    let height = optional_length(attributes, "height", part)?.unwrap_or(CSS_PIXELS_PER_INCH);
    if width <= 0.0 || height <= 0.0 {
        return Err(format_error(part, "ODT shape dimensions must be positive"));
    }
    let bounds = Rect {
        x: optional_signed_length(attributes, "x", part)?.unwrap_or(0.0),
        y: optional_signed_length(attributes, "y", part)?.unwrap_or(0.0),
        width,
        height,
    };
    let geometry = match kind {
        "ellipse" => Geometry::Ellipse,
        "regular-polygon" => {
            let corners = optional_positive_u32(attributes, "corners", part)?
                .unwrap_or(5)
                .clamp(3, 32);
            regular_polygon_geometry(corners, width, height)
        }
        "caption" | "rect" | "custom-shape" => Geometry::Rectangle,
        _ => Geometry::Rectangle,
    };
    Ok((bounds, geometry))
}

// Keep the ODT open-subpath closure: the shared polygon helper uses ClosePath,
// which changes the stroke join at the first vertex. These paths are not equivalent.
fn regular_polygon_geometry(corners: u32, width: f32, height: f32) -> Geometry {
    let center_x = width / 2.0;
    let center_y = height / 2.0;
    let mut commands = Vec::with_capacity(corners as usize + 1);
    for index in 0..corners {
        let angle =
            -std::f32::consts::FRAC_PI_2 + std::f32::consts::TAU * index as f32 / corners as f32;
        let x = center_x + center_x * angle.cos();
        let y = center_y + center_y * angle.sin();
        let command = if index == 0 {
            PathCommand::MoveTo { x, y }
        } else {
            PathCommand::LineTo { x, y }
        };
        commands.push(command);
    }
    if let Some(&PathCommand::MoveTo { x, y }) = commands.first() {
        commands.push(PathCommand::LineTo { x, y });
    }
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    }
}

fn parse_odf_transform(value: &str, part: &str) -> Result<AffineTransform, Diagnostic> {
    let mut remaining = value.trim();
    let mut transform = AffineTransform::IDENTITY;
    while !remaining.is_empty() {
        let open = remaining
            .find('(')
            .ok_or_else(|| format_error(part, "ODT transform is missing an opening parenthesis"))?;
        let close = remaining[open + 1..]
            .find(')')
            .map(|index| open + 1 + index)
            .ok_or_else(|| format_error(part, "ODT transform is missing a closing parenthesis"))?;
        let operation = remaining[..open].trim();
        let arguments = remaining[open + 1..close]
            .split(|character: char| character == ',' || character.is_ascii_whitespace())
            .filter(|argument| !argument.is_empty())
            .collect::<Vec<_>>();
        let number = |index: usize| -> Result<f32, Diagnostic> {
            arguments
                .get(index)
                .and_then(|value| value.parse::<f32>().ok())
                .filter(|value| value.is_finite())
                .ok_or_else(|| format_error(part, "ODT transform contains an invalid number"))
        };
        let length = |index: usize| -> Result<f32, Diagnostic> {
            arguments
                .get(index)
                .and_then(|value| parse_signed_length(value))
                .ok_or_else(|| format_error(part, "ODT transform contains an invalid length"))
        };
        let next = match operation {
            "rotate" if arguments.len() == 1 => {
                let angle = number(0)?;
                AffineTransform {
                    a: angle.cos(),
                    b: angle.sin(),
                    c: -angle.sin(),
                    d: angle.cos(),
                    e: 0.0,
                    f: 0.0,
                }
            }
            "translate" if matches!(arguments.len(), 1 | 2) => AffineTransform {
                e: length(0)?,
                f: if arguments.len() == 2 {
                    length(1)?
                } else {
                    0.0
                },
                ..AffineTransform::IDENTITY
            },
            "scale" if matches!(arguments.len(), 1 | 2) => {
                let x = number(0)?;
                AffineTransform {
                    a: x,
                    d: if arguments.len() == 2 { number(1)? } else { x },
                    ..AffineTransform::IDENTITY
                }
            }
            "matrix" if arguments.len() == 6 => AffineTransform {
                a: number(0)?,
                b: number(1)?,
                c: number(2)?,
                d: number(3)?,
                e: number(4)?,
                f: number(5)?,
            },
            _ => {
                return Err(format_error(
                    part,
                    "ODT transform operation is unsupported or malformed",
                ));
            }
        };
        transform = transform.concat(next);
        remaining = remaining[close + 1..].trim_start();
    }
    if !transform.is_valid() {
        return Err(format_error(part, "ODT transform is not finite"));
    }
    Ok(transform)
}

fn transformed_visual(visual: Visual, transform: AffineTransform) -> Visual {
    if transform == AffineTransform::IDENTITY {
        visual
    } else {
        Visual::Layer {
            transform,
            opacity: 1.0,
            blend_mode: crate::model::BlendMode::Normal,
            visual: Box::new(visual),
        }
    }
}

fn finish_group(
    group: DrawGroupState,
    flow_y: f32,
    unit_index: u32,
    objects: &mut Vec<Object>,
    limits: Limits,
) -> Result<(), Diagnostic> {
    let offset_y = if group.character_anchored {
        group
            .shapes
            .iter()
            .map(|shape| shape.bounds.y)
            .reduce(f32::min)
            .map_or(0.0, |minimum| flow_y + group.margin_top - minimum)
    } else {
        0.0
    };
    finish_shapes(group.shapes, offset_y, unit_index, objects, limits)
}

fn finish_shapes(
    shapes: Vec<ShapeState>,
    offset_y: f32,
    unit_index: u32,
    objects: &mut Vec<Object>,
    limits: Limits,
) -> Result<(), Diagnostic> {
    finish_shapes_from(
        shapes,
        offset_y,
        unit_index,
        objects,
        limits.max_document_objects,
        ShapeSource::Content,
    )
}

#[derive(Clone, Copy)]
enum ShapeSource<'a> {
    Content,
    Master {
        master_name: &'a str,
        story_kind: &'a str,
    },
}

fn finish_shapes_from(
    mut shapes: Vec<ShapeState>,
    offset_y: f32,
    unit_index: u32,
    objects: &mut Vec<Object>,
    object_limit: usize,
    source: ShapeSource<'_>,
) -> Result<(), Diagnostic> {
    for shape in &mut shapes {
        shape.bounds.y += offset_y;
        if let Some(connector) = shape.connector.as_mut() {
            connector.start.1 += offset_y;
            connector.end.1 += offset_y;
        }
    }
    let anchors = shapes
        .iter()
        .filter(|shape| shape.connector.is_none())
        .filter_map(|shape| {
            shape.element_id.clone().map(|element_id| {
                (
                    element_id,
                    OdfShapeAnchor {
                        bounds: shape.bounds,
                        transform: shape.transform,
                    },
                )
            })
        })
        .collect::<HashMap<_, _>>();
    for shape in shapes {
        finish_shape(shape, &anchors, unit_index, objects, object_limit, source)?;
    }
    Ok(())
}

fn finish_shape(
    shape: ShapeState,
    anchors: &HashMap<String, OdfShapeAnchor>,
    unit_index: u32,
    objects: &mut Vec<Object>,
    object_limit: usize,
    source: ShapeSource<'_>,
) -> Result<(), Diagnostic> {
    if objects.len() >= object_limit {
        return Err(object_limit_error(
            "ODT shape count exceeds the configured object limit",
        ));
    }
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| object_limit_error("ODT shape object count exceeds supported range"))?;
    let (part, stable_id, path) = match source {
        ShapeSource::Content => (
            CONTENT_PART,
            format!("odt:shape:{}", shape.index),
            format!(
                "/office:document-content/office:body/office:text/draw:{}[{}]",
                shape.element,
                shape.index + 1
            ),
        ),
        ShapeSource::Master {
            master_name,
            story_kind,
        } => (
            STYLES_PART,
            format!(
                "odt:{story_kind}:{unit_index}:{master_name}:shape:{}",
                shape.index
            ),
            format!(
                "/office:document-styles/office:master-styles/style:master-page[@style:name='{master_name}']/style:{story_kind}/draw:{}[{}]",
                shape.element,
                shape.index + 1
            ),
        ),
    };
    let geometry = shape.connector.as_ref().map_or_else(
        || shape.geometry.clone(),
        |connector| {
            odf_connector_geometry(
                connector,
                shape.bounds,
                shape.transform,
                anchors,
                OdfConnectorStyle {
                    stroke_width: shape.graphic.stroke_width,
                    dash: None,
                    marker_start: shape.graphic.marker_start,
                    marker_end: shape.graphic.marker_end,
                },
            )
        },
    );
    objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: stable_id.clone(),
        parent_stable_id: None,
        kind: ObjectKind::Shape,
        unit_index,
        bounds: shape.bounds,
        z: shape
            .z
            .unwrap_or_else(|| i32::try_from(numeric_id).unwrap_or(i32::MAX)),
        text: None,
        source: SourceRef {
            part: part.to_owned(),
            mapping: if shape.element_id.is_some() {
                MappingQuality::Exact
            } else {
                MappingQuality::Derived
            },
            locator: SourceLocator::Odt {
                kind: "element",
                element_id: shape.element_id.clone(),
                path: path.clone(),
                row: None,
                column: None,
                text_range: None,
            },
        },
        visual: transformed_visual(
            Visual::PaintedShape {
                geometry,
                fill: shape.graphic.fill.clone(),
                stroke: shape.graphic.stroke.clone(),
                stroke_width: shape.graphic.stroke_width,
            },
            shape.transform,
        ),
    });
    if !shape.text.is_empty() {
        if objects.len() >= object_limit {
            return Err(object_limit_error(
                "ODT shape text exceeds the configured object limit",
            ));
        }
        let text_id = u32::try_from(objects.len())
            .map_err(|_| object_limit_error("ODT shape text count exceeds supported range"))?;
        let text_length = u32::try_from(shape.text.chars().count()).unwrap_or(u32::MAX);
        objects.push(Object {
            numeric_id: text_id,
            parent_numeric_id: Some(numeric_id),
            stable_id: format!("{stable_id}:text"),
            parent_stable_id: Some(stable_id),
            kind: ObjectKind::TextBox,
            unit_index,
            bounds: shape.bounds,
            z: i32::try_from(text_id).unwrap_or(i32::MAX),
            text: Some(shape.text.clone()),
            source: SourceRef {
                part: part.to_owned(),
                mapping: MappingQuality::Derived,
                locator: SourceLocator::Odt {
                    kind: "text-range",
                    element_id: shape.element_id,
                    path,
                    row: None,
                    column: None,
                    text_range: Some((0, text_length)),
                },
            },
            visual: Visual::RichText {
                geometry: Geometry::Rectangle,
                fill: Paint::None,
                stroke: Paint::None,
                stroke_width: 0.0,
                align: shape.text_style.align,
                line_height: shape.text_style.line_height,
                runs: shape.runs,
            },
        });
    }
    Ok(())
}

fn list_start(styles: &StyleCatalog, name: Option<&str>, level: u32) -> u32 {
    name.and_then(|name| styles.list_styles.get(name))
        .and_then(|style| style.levels.get(&level))
        .and_then(|level| match level {
            OdtListLevel::Number { start, .. } => Some(*start),
            OdtListLevel::Bullet { .. } => None,
        })
        .unwrap_or(1)
}

fn format_list_label(
    styles: &StyleCatalog,
    name: Option<&str>,
    level: u32,
    value: u32,
) -> Option<String> {
    Some(
        match name
            .and_then(|name| styles.list_styles.get(name))
            .and_then(|style| style.levels.get(&level))
        {
            Some(OdtListLevel::Bullet {
                character,
                font_family,
            }) => super::normalize_symbol_font_character(character, font_family.as_deref()),
            Some(OdtListLevel::Number { format, .. }) if format.is_empty() => return None,
            Some(OdtListLevel::Number {
                format,
                prefix,
                suffix,
                ..
            }) => format!("{prefix}{}{suffix}", format_odt_number(value, format)),
            None => "•".to_owned(),
        },
    )
}

fn format_odt_number(value: u32, format: &str) -> String {
    if let Some(number) = super::zero_padded_number(value, format) {
        return number;
    }
    match format {
        "a" | "A" => {
            let uppercase = format == "A";
            let mut value = value.max(1);
            let mut result = String::new();
            while value > 0 {
                value -= 1;
                let character = char::from_u32(u32::from(b'a') + value % 26).unwrap_or('a');
                result.insert(
                    0,
                    if uppercase {
                        character.to_ascii_uppercase()
                    } else {
                        character
                    },
                );
                value /= 26;
            }
            result
        }
        "i" | "I" => {
            let mut value = value;
            let mut result = String::new();
            for (number, literal) in [
                (1000, "M"),
                (900, "CM"),
                (500, "D"),
                (400, "CD"),
                (100, "C"),
                (90, "XC"),
                (50, "L"),
                (40, "XL"),
                (10, "X"),
                (9, "IX"),
                (5, "V"),
                (4, "IV"),
                (1, "I"),
            ] {
                while value >= number {
                    result.push_str(literal);
                    value -= number;
                }
            }
            if format == "i" {
                result.to_ascii_lowercase()
            } else {
                result
            }
        }
        _ => value.to_string(),
    }
}

fn parse_page_backgrounds_xml(
    bytes: &[u8],
    limits: Limits,
    package: Option<&Package<'_>>,
    styles: &mut StyleCatalog,
) -> Result<Vec<Diagnostic>, Diagnostic> {
    let mut diagnostics = Vec::new();
    let mut layout = None;
    let mut properties_depth = None;
    let mut depth = 0_usize;
    let mut repeat = String::from("repeat");
    let mut tile_width = None;
    let mut tile_height = None;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "page-layout" {
                    layout = optional_attribute(&attributes, "name", STYLES_PART)?;
                }
                if local == "page-layout-properties" && layout.is_some() {
                    properties_depth = (!empty).then_some(depth);
                    repeat = optional_attribute(&attributes, "repeat", STYLES_PART)?
                        .unwrap_or_else(|| "repeat".to_owned());
                    tile_width = optional_length(&attributes, "fill-image-width", STYLES_PART)?
                        .filter(|v| *v > 0.0);
                    tile_height = optional_length(&attributes, "fill-image-height", STYLES_PART)?
                        .filter(|v| *v > 0.0);
                    if let Some(color) =
                        optional_attribute(&attributes, "background-color", STYLES_PART)?
                            .and_then(|v| parse_odf_color(&v))
                    {
                        styles
                            .page_backgrounds
                            .insert(layout.clone().unwrap(), Paint::Solid(color));
                    }
                }
                if local == "background-image"
                    && properties_depth.is_some_and(|start| depth == start + 1)
                {
                    if let Some(value) = optional_attribute(&attributes, "repeat", STYLES_PART)? {
                        repeat = value;
                    }
                    if let Some(href) = optional_attribute(&attributes, "href", STYLES_PART)?
                        .filter(|v| !v.is_empty())
                    {
                        // ponytail: positioned no-repeat images use tiling until intrinsic-size placement is shared.
                        if repeat == "no-repeat" {
                            diagnostics.push(
                                Diagnostic::warning(
                                    DiagnosticCode::UnsupportedFeature,
                                    Phase::Render,
                                    Fidelity::Approximate,
                                    "ODT positioned page background uses repeated image fill",
                                )
                                .in_part(STYLES_PART),
                            );
                        }
                        let image = (|| -> Result<Option<Paint>, Diagnostic> {
                            let target =
                                match resolve_odf_image_target(&href, limits.max_zip_path_bytes) {
                                    Ok(OdfImageTarget::Embedded(target)) => target,
                                    Ok(OdfImageTarget::External) => {
                                        diagnostics.push(
                                            Diagnostic::warning(
                                                DiagnosticCode::ExternalResourceBlocked,
                                                Phase::Security,
                                                Fidelity::Blocked,
                                                "external ODT page background was blocked",
                                            )
                                            .in_part(STYLES_PART),
                                        );
                                        return Ok(None);
                                    }
                                    Err(message) => return Err(format_error(STYLES_PART, message)),
                                };
                            let package = package.ok_or_else(|| {
                                format_error(
                                    STYLES_PART,
                                    "flat ODT page background has no packaged image",
                                )
                            })?;
                            let data = package.part(&target)?.ok_or_else(|| {
                                format_error(STYLES_PART, "ODT page background image is missing")
                            })?;
                            let declared = office_image_media_type(&target, &data);
                            let media_type =
                                super::presentation_image::recover_office_image_signature(
                                    declared, &data,
                                )
                                .map_err(|_| {
                                    format_error(
                                        STYLES_PART,
                                        "ODT page background image format is unsupported",
                                    )
                                })?;
                            if declared.is_err() {
                                diagnostics.push(
                                    Diagnostic::warning(
                                        DiagnosticCode::UnsupportedFeature,
                                        Phase::Parse,
                                        Fidelity::Approximate,
                                        "ODT background image recovered using its signature",
                                    )
                                    .in_part(&target),
                                );
                            }
                            Ok(Some(Paint::Image {
                                media_type: media_type.to_owned(),
                                bytes: data.into_vec(),
                                crop: ImageCrop::default(),
                                tile: repeat != "stretch",
                                tile_width,
                                tile_height,
                                mapping: None,
                            }))
                        })();
                        match image {
                            Ok(Some(image)) => {
                                styles
                                    .page_backgrounds
                                    .insert(layout.clone().unwrap(), image);
                            }
                            Ok(None) => {}
                            Err(error) if error.code == DiagnosticCode::FormatInvalid => {
                                diagnostics.push(
                                    Diagnostic::warning(
                                        error.code,
                                        Phase::Parse,
                                        Fidelity::Omitted,
                                        error.message,
                                    )
                                    .in_part(STYLES_PART),
                                )
                            }
                            Err(error) => return Err(error),
                        }
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if local_name(name) == "page-layout" {
                    layout = None;
                }
                if properties_depth == Some(depth) {
                    properties_depth = None;
                }
            }
            _ => {}
        }
        Ok(())
    })?;
    Ok(diagnostics)
}

fn append_page_backgrounds(
    content: &mut ParsedContent,
    styles: &StyleCatalog,
    limits: Limits,
) -> Result<(), Diagnostic> {
    let mut image_bytes = content.materialized_image_bytes;
    for unit in &content.units {
        let Some(fill) = styles
            .master_page(unit.index)
            .and_then(|(_, master)| master.page_layout.as_ref())
            .and_then(|name| styles.page_backgrounds.get(name))
        else {
            continue;
        };
        if matches!(fill, Paint::None) {
            continue;
        }
        if content.objects.len() >= limits.max_document_objects {
            return Err(object_limit_error(
                "ODT page backgrounds exceed the object limit",
            ));
        }
        if let Paint::Image { bytes, .. } = fill {
            reserve_materialized_image_bytes(
                &mut image_bytes,
                bytes.len(),
                limits.max_total_uncompressed_bytes,
                STYLES_PART,
            )?;
        }
        let numeric_id = content.objects.len() as u32;
        content.objects.push(Object {
            numeric_id,
            parent_numeric_id: None,
            stable_id: format!("odt:page:{}:background", unit.index),
            parent_stable_id: None,
            kind: ObjectKind::Shape,
            unit_index: unit.index,
            bounds: Rect {
                x: 0.0,
                y: 0.0,
                width: unit.width,
                height: unit.height,
            },
            z: -1,
            text: None,
            source: SourceRef {
                part: STYLES_PART.to_owned(),
                mapping: MappingQuality::Derived,
                locator: SourceLocator::Odt {
                    kind: "element",
                    element_id: None,
                    row: None,
                    column: None,
                    path: "/office:document-styles/office:automatic-styles/style:page-layout"
                        .to_owned(),
                    text_range: None,
                },
            },
            visual: Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: fill.clone(),
                stroke: Paint::None,
                stroke_width: 0.0,
            },
        });
    }
    Ok(())
}

fn parse_page_layout_xml(
    bytes: &[u8],
    limits: Limits,
    master_page_name: Option<&str>,
) -> Result<PageLayout, Diagnostic> {
    let mut current_layout = None;
    let mut layouts = HashMap::new();
    let mut master_layouts = Vec::new();

    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => match local_name(name) {
                "page-layout" => {
                    if current_layout.is_some() {
                        return Err(format_error(STYLES_PART, "nested page layouts are invalid"));
                    }
                    let name = required_attribute(&attributes, "name", STYLES_PART)?;
                    if empty {
                        return Err(format_error(STYLES_PART, "page layout has no properties"));
                    }
                    current_layout = Some(name);
                }
                "page-layout-properties" => {
                    let Some(name) = current_layout.as_ref() else {
                        return Ok(());
                    };
                    let uniform_margin = optional_length(&attributes, "margin", STYLES_PART)?;
                    let mut layout = PageLayout {
                        width: required_length(&attributes, "page-width", STYLES_PART)?,
                        height: required_length(&attributes, "page-height", STYLES_PART)?,
                        margin_top: uniform_margin.unwrap_or(DEFAULT_MARGIN),
                        margin_right: uniform_margin.unwrap_or(DEFAULT_MARGIN),
                        margin_bottom: uniform_margin.unwrap_or(DEFAULT_MARGIN),
                        margin_left: uniform_margin.unwrap_or(DEFAULT_MARGIN),
                    };
                    layout.margin_top = optional_length(&attributes, "margin-top", STYLES_PART)?
                        .unwrap_or(layout.margin_top);
                    layout.margin_right =
                        optional_length(&attributes, "margin-right", STYLES_PART)?
                            .unwrap_or(layout.margin_right);
                    layout.margin_bottom =
                        optional_length(&attributes, "margin-bottom", STYLES_PART)?
                            .unwrap_or(layout.margin_bottom);
                    layout.margin_left = optional_length(&attributes, "margin-left", STYLES_PART)?
                        .unwrap_or(layout.margin_left);
                    if layouts.insert(name.clone(), layout).is_some() {
                        return Err(format_error(
                            STYLES_PART,
                            format!("duplicate page layout properties: {name}"),
                        ));
                    }
                }
                "master-page" => {
                    master_layouts.push((
                        required_attribute(&attributes, "name", STYLES_PART)?,
                        required_attribute(&attributes, "page-layout-name", STYLES_PART)?,
                    ));
                }
                _ => {}
            },
            XmlEvent::EndElement { name } if local_name(name) == "page-layout" => {
                current_layout = None;
            }
            _ => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, STYLES_PART))?;

    let Some((_, layout_name)) = master_page_name
        .and_then(|name| master_layouts.iter().find(|(master, _)| master == name))
        .or_else(|| master_layouts.first())
    else {
        return Ok(PageLayout::default());
    };
    layouts.get(layout_name).copied().ok_or_else(|| {
        format_error(
            STYLES_PART,
            format!("master page references unknown page layout {layout_name}"),
        )
    })
}

fn parse_style_catalog_xml(
    bytes: &[u8],
    limits: Limits,
    part: &str,
    catalog: &mut StyleCatalog,
) -> Result<(), Diagnostic> {
    let mut depth = 0_usize;
    let mut automatic_styles = false;
    let mut current: Option<(usize, Option<String>, OdtStyle)> = None;
    let mut current_list: Option<(usize, String, OdtListStyle)> = None;
    let mut current_list_level = None;
    let mut current_list_bullet: Option<(usize, u32, String, Option<String>)> = None;
    let mut page_layout = None;
    let mut story_style = None;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                match local {
                    "automatic-styles" => automatic_styles = !empty,
                    "page-layout" => page_layout = optional_attribute(&attributes, "name", part)?,
                    "header-style" | "footer-style" if page_layout.is_some() => {
                        story_style = (!empty).then_some(usize::from(local == "footer-style"));
                    }
                    "header-footer-properties"
                        if page_layout.is_some() && story_style.is_some() =>
                    {
                        let mut style = MasterStoryStyle::default();
                        style.border.fill = 0;
                        style.border.stroke = 0x000000ff;
                        style.height = optional_length(&attributes, "height", part)?.unwrap_or(0.0);
                        // Some producers emit negative minimum heights; they cannot shrink content.
                        style.min_height = optional_attribute(&attributes, "min-height", part)?
                            .and_then(|value| parse_signed_length(&value))
                            .unwrap_or(0.0)
                            .max(0.0);
                        for (property, values) in [
                            ("margin", &mut style.margin),
                            ("padding", &mut style.padding),
                        ] {
                            values
                                .fill(optional_length(&attributes, property, part)?.unwrap_or(0.0));
                            for (index, side) in
                                ["top", "right", "bottom", "left"].iter().enumerate()
                            {
                                if let Some(value) = optional_length(
                                    &attributes,
                                    &format!("{property}-{side}"),
                                    part,
                                )? {
                                    values[index] = value;
                                }
                            }
                        }
                        if let Some(border) = optional_attribute(&attributes, "border", part)? {
                            apply_odt_cell_border(&border, &mut style.border);
                        }
                        if let Some(color) =
                            optional_attribute(&attributes, "background-color", part)?
                                .and_then(|v| parse_odf_color(&v))
                        {
                            style.border.fill = color;
                        }
                        catalog
                            .story_styles
                            .entry(page_layout.clone().unwrap())
                            .or_default()[story_style.unwrap()] = Some(style);
                    }
                    "list-style" => {
                        if current_list.is_some() {
                            return Err(format_error(part, "nested ODT list styles are invalid"));
                        }
                        let name = required_attribute(&attributes, "name", part)?;
                        if empty {
                            catalog.list_styles.insert(name, OdtListStyle::default());
                        } else {
                            current_list = Some((depth, name, OdtListStyle::default()));
                        }
                    }
                    "list-level-style-number" if current_list.is_some() => {
                        let level = optional_positive_u32(&attributes, "level", part)?
                            .ok_or_else(|| format_error(part, "ODT list level is missing"))?;
                        current_list_level = (!empty).then_some(level);
                        let format = optional_attribute(&attributes, "num-format", part)?
                            .unwrap_or_else(|| "1".to_owned());
                        let prefix = optional_attribute(&attributes, "num-prefix", part)?
                            .unwrap_or_default();
                        let suffix = optional_attribute(&attributes, "num-suffix", part)?
                            .unwrap_or_else(|| ".".to_owned());
                        let start =
                            optional_positive_u32(&attributes, "start-value", part)?.unwrap_or(1);
                        current_list
                            .as_mut()
                            .ok_or_else(|| format_error(part, "ODT list style state is missing"))?
                            .2
                            .levels
                            .insert(
                                level,
                                OdtListLevel::Number {
                                    format,
                                    prefix,
                                    suffix,
                                    start,
                                },
                            );
                    }
                    "list-level-style-bullet" if current_list.is_some() => {
                        let level = optional_positive_u32(&attributes, "level", part)?
                            .ok_or_else(|| format_error(part, "ODT list level is missing"))?;
                        current_list_level = (!empty).then_some(level);
                        let bullet = optional_attribute(&attributes, "bullet-char", part)?
                            .filter(|value| !value.is_empty())
                            .unwrap_or_else(|| "•".to_owned());
                        if empty {
                            current_list
                                .as_mut()
                                .ok_or_else(|| {
                                    format_error(part, "ODT list style state is missing")
                                })?
                                .2
                                .levels
                                .insert(
                                    level,
                                    OdtListLevel::Bullet {
                                        character: bullet,
                                        font_family: None,
                                    },
                                );
                        } else {
                            current_list_bullet = Some((depth, level, bullet, None));
                        }
                    }
                    "list-level-label-alignment"
                        if current_list.is_some() && current_list_level.is_some() =>
                    {
                        let position = OdtListPosition {
                            margin_left: optional_signed_length(&attributes, "margin-left", part)?
                                .unwrap_or(0.0),
                            text_indent: optional_signed_length(&attributes, "text-indent", part)?
                                .unwrap_or(0.0),
                            tab_stop: optional_length(&attributes, "list-tab-stop-position", part)?,
                            followed_by: match optional_attribute(
                                &attributes,
                                "label-followed-by",
                                part,
                            )?
                            .as_deref()
                            {
                                Some("nothing") => "",
                                Some("space") => " ",
                                _ => "\t",
                            },
                        };
                        current_list
                            .as_mut()
                            .unwrap()
                            .2
                            .positions
                            .insert(current_list_level.unwrap(), position);
                    }
                    "font-face" => {
                        let name = required_attribute(&attributes, "name", part)?;
                        let family = optional_attribute(&attributes, "font-family", part)?
                            .unwrap_or_else(|| name.clone());
                        catalog
                            .font_faces
                            .insert(name, super::odf_primary_font_family(&family).to_owned());
                    }
                    "style" | "default-style" => {
                        let name = if local == "default-style" {
                            None
                        } else {
                            Some(required_attribute(&attributes, "name", part)?)
                        };
                        let family = optional_attribute(&attributes, "family", part)?
                            .unwrap_or_else(|| "paragraph".to_owned());
                        let parent = optional_attribute(&attributes, "parent-style-name", part)?;
                        let inherited = parent
                            .as_deref()
                            .and_then(|parent| catalog.styles.get(parent));
                        let mut paragraph = inherited.map_or_else(
                            || catalog.default_paragraph.clone(),
                            |style| style.paragraph.clone(),
                        );
                        paragraph.common_style_name = if automatic_styles {
                            parent.clone()
                        } else {
                            name.clone()
                        };
                        let graphic = inherited.map_or_else(
                            || catalog.default_graphic.clone(),
                            |style| style.graphic.clone(),
                        );
                        let cell = inherited.map_or_else(
                            || catalog.default_cell.clone(),
                            |style| style.cell.clone(),
                        );
                        let row = inherited.map_or(catalog.default_row, |style| style.row);
                        let column =
                            inherited.map_or(catalog.default_table_column, |style| style.column);
                        let text_overrides = inherited
                            .map(|style| style.text_overrides.clone())
                            .unwrap_or_default();
                        let style = OdtStyle {
                            family,
                            master_page_name: optional_attribute(
                                &attributes,
                                "master-page-name",
                                part,
                            )?,
                            paragraph,
                            graphic,
                            cell,
                            row,
                            column,
                            text_overrides,
                            columns: inherited.and_then(|style| style.columns),
                        };
                        if empty {
                            commit_odt_style(catalog, name, style);
                        } else {
                            current = Some((depth, name, style));
                        }
                    }
                    "text-properties" if current_list_bullet.is_some() => {
                        let font_name = optional_attribute(&attributes, "font-name", part)?
                            .or(optional_attribute(&attributes, "font-name-asian", part)?)
                            .or(optional_attribute(&attributes, "font-name-complex", part)?);
                        let direct_family = optional_attribute(&attributes, "font-family", part)?
                            .or(optional_attribute(&attributes, "font-family-asian", part)?)
                            .or(optional_attribute(
                                &attributes,
                                "font-family-complex",
                                part,
                            )?);
                        current_list_bullet
                            .as_mut()
                            .ok_or_else(|| format_error(part, "ODT list bullet state is missing"))?
                            .3 = font_name
                            .as_deref()
                            .and_then(|name| catalog.font_faces.get(name).cloned())
                            .or(font_name)
                            .or_else(|| {
                                direct_family.map(|family| {
                                    super::odf_primary_font_family(&family).to_owned()
                                })
                            });
                    }
                    "text-properties" if current.is_some() => {
                        let current = current
                            .as_mut()
                            .ok_or_else(|| format_error(part, "ODT style state is missing"))?;
                        let style = &mut current.2.paragraph.text;
                        let overrides = &mut current.2.text_overrides;
                        if let Some(font_name) = optional_attribute(&attributes, "font-name", part)?
                        {
                            let family = catalog
                                .font_faces
                                .get(&font_name)
                                .cloned()
                                .unwrap_or(font_name);
                            style.font_family.clone_from(&family);
                            overrides.font_family = Some(family);
                        } else if let Some(font_family) =
                            optional_attribute(&attributes, "font-family", part)?
                        {
                            let family = super::odf_primary_font_family(&font_family).to_owned();
                            style.font_family.clone_from(&family);
                            overrides.font_family = Some(family);
                        }
                        if let Some(font_size) = optional_attribute(&attributes, "font-size", part)?
                            && let Some(font_size) = parse_length(&font_size)
                            && font_size > 0.0
                        {
                            style.font_size = font_size;
                            overrides.font_size = Some(font_size);
                        }
                        if let Some(color) = optional_attribute(&attributes, "color", part)?
                            && color != "transparent"
                        {
                            let color = parse_odf_color(&color)
                                .ok_or_else(|| format_error(part, "ODT text color is invalid"))?;
                            style.color = color;
                            overrides.color = Some(color);
                        }
                        if let Some(weight) = optional_attribute(&attributes, "font-weight", part)?
                        {
                            let bold =
                                weight == "bold" || weight.parse::<u32>().is_ok_and(|v| v >= 600);
                            style.bold = bold;
                            overrides.bold = Some(bold);
                        }
                        if let Some(font_style) =
                            optional_attribute(&attributes, "font-style", part)?
                        {
                            let italic = matches!(font_style.as_str(), "italic" | "oblique");
                            style.italic = italic;
                            overrides.italic = Some(italic);
                        }
                        if let Some(underline) =
                            optional_attribute(&attributes, "text-underline-style", part)?
                        {
                            let underline = underline != "none";
                            style.underline = underline;
                            overrides.underline = Some(underline);
                        }
                        if let Some(strike) =
                            optional_attribute(&attributes, "text-line-through-style", part)?
                        {
                            let strikethrough = strike != "none";
                            style.strikethrough = strikethrough;
                            overrides.strikethrough = Some(strikethrough);
                        }
                        if let Some(background) =
                            optional_attribute(&attributes, "background-color", part)?
                        {
                            let highlight = if background == "transparent" {
                                0
                            } else {
                                parse_odf_color(&background).ok_or_else(|| {
                                    format_error(part, "ODT text background color is invalid")
                                })?
                            };
                            style.highlight = highlight;
                            overrides.highlight = Some(highlight);
                        }
                        if let Some(position) =
                            optional_attribute(&attributes, "text-position", part)?
                        {
                            let mut values = position.split_ascii_whitespace();
                            let position = values.next().unwrap_or_default();
                            let baseline_shift = if position == "super" {
                                style.font_size * 0.35
                            } else if position == "sub" {
                                style.font_size * -0.20
                            } else {
                                position
                                    .strip_suffix('%')
                                    .and_then(|value| value.parse::<f32>().ok())
                                    .filter(|value| value.is_finite())
                                    .map_or(0.0, |value| style.font_size * value / 100.0)
                            };
                            style.baseline_shift = baseline_shift;
                            overrides.baseline_shift = Some(baseline_shift);
                            if let Some(scale) = values
                                .next()
                                .and_then(|value| value.strip_suffix('%'))
                                .and_then(|value| value.parse::<f32>().ok())
                                .filter(|value| value.is_finite() && *value > 0.0)
                            {
                                overrides.font_size_scale = Some(scale / 100.0);
                            }
                        }
                        if let Some(spacing) =
                            optional_attribute(&attributes, "letter-spacing", part)?
                            && spacing != "normal"
                        {
                            let letter_spacing =
                                parse_signed_length(&spacing).ok_or_else(|| {
                                    format_error(part, "ODT letter spacing is invalid")
                                })?;
                            style.letter_spacing = letter_spacing;
                            overrides.letter_spacing = Some(letter_spacing);
                        }
                    }
                    "columns"
                        if current
                            .as_ref()
                            .is_some_and(|(_, _, style)| style.family == "section") =>
                    {
                        let count =
                            optional_positive_u32(&attributes, "column-count", part)?.unwrap_or(1);
                        if count as usize > limits.max_document_objects {
                            return Err(object_limit_error(
                                "ODT section column count exceeds object limit",
                            ));
                        }
                        let gap = optional_length(&attributes, "column-gap", part)?.unwrap_or(0.0);
                        if current.as_ref().unwrap().2.columns != Some((0, 0.0)) {
                            current.as_mut().unwrap().2.columns = Some((count, gap));
                        }
                    }
                    // Unequal columns and non-balanced sections retain the existing flow.
                    "column" | "section-properties"
                        if current
                            .as_ref()
                            .is_some_and(|(_, _, style)| style.family == "section") =>
                    {
                        if local == "column"
                            || optional_attribute(&attributes, "dont-balance-text-columns", part)?
                                .as_deref()
                                == Some("true")
                            || ["margin-left", "margin-right"].iter().any(|name| {
                                optional_signed_length(&attributes, name, part)
                                    .ok()
                                    .flatten()
                                    .is_some_and(|v| v != 0.0)
                            })
                            || optional_attribute(&attributes, "writing-mode", part)?
                                .is_some_and(|v| v != "lr-tb")
                        {
                            current.as_mut().unwrap().2.columns = Some((0, 0.0));
                        }
                    }
                    "tab-stops" if current.is_some() => {
                        current.as_mut().unwrap().2.paragraph.tab_stops.clear();
                    }
                    "tab-stop" if current.is_some() => {
                        let position = required_length(&attributes, "position", part)?;
                        let align = match optional_attribute(&attributes, "type", part)?.as_deref()
                        {
                            Some("right") => TextAlign::End,
                            Some("center") => TextAlign::Center,
                            _ => TextAlign::Start,
                        };
                        let leader_style = optional_attribute(&attributes, "leader-style", part)?;
                        let leader = if leader_style.as_deref() == Some("none") {
                            TextTabLeader::None
                        } else {
                            match optional_attribute(&attributes, "leader-text", part)?.as_deref() {
                                Some(".") => TextTabLeader::Dot,
                                Some("-") => TextTabLeader::Hyphen,
                                Some("_") => TextTabLeader::Underscore,
                                Some("·") => TextTabLeader::MiddleDot,
                                _ => match leader_style.as_deref() {
                                    Some("dotted") => TextTabLeader::Dot,
                                    Some("dash" | "long-dash") => TextTabLeader::Hyphen,
                                    Some("solid") => TextTabLeader::Underscore,
                                    _ => TextTabLeader::None,
                                },
                            }
                        };
                        let stops = &mut current.as_mut().unwrap().2.paragraph.tab_stops;
                        stops.push(TextTabStop {
                            position,
                            align,
                            leader,
                        });
                    }
                    "drop-cap" if current.is_some() => {
                        let lines = optional_positive_u32(&attributes, "lines", part)?.unwrap_or(3);
                        let characters =
                            optional_positive_u32(&attributes, "length", part)?.unwrap_or(1);
                        let padding =
                            optional_length(&attributes, "distance", part)?.unwrap_or(0.0);
                        current
                            .as_mut()
                            .ok_or_else(|| format_error(part, "ODT style state is missing"))?
                            .2
                            .paragraph
                            .drop_cap = Some(crate::model::TextDropCap {
                            characters: characters.min(16),
                            lines: lines.min(16),
                            raised_lines: 0,
                            padding,
                            outdent: 0.0,
                        });
                    }
                    "paragraph-properties" if current.is_some() => {
                        let style = &mut current
                            .as_mut()
                            .ok_or_else(|| format_error(part, "ODT style state is missing"))?
                            .2
                            .paragraph;
                        if let Some(distance) =
                            optional_length(&attributes, "tab-stop-distance", part)?
                        {
                            style.default_tab_stop = distance.max(1.0);
                        }
                        if let Some(align) = optional_attribute(&attributes, "text-align", part)? {
                            style.align = super::odf_text_align(&align);
                        }
                        if let Some(line_height) =
                            optional_attribute(&attributes, "line-height", part)?
                        {
                            style.natural_line_height = line_height == "normal";
                            style.line_height =
                                if let Some(percent) = line_height.strip_suffix('%') {
                                    percent
                                        .parse::<f32>()
                                        .ok()
                                        .filter(|value| value.is_finite())
                                        .map_or(style.line_height, |value| {
                                            style.text.font_size * value / 100.0
                                        })
                                } else {
                                    parse_length(&line_height).unwrap_or(style.line_height)
                                }
                                .max(style.text.font_size);
                        }
                        if let Some(value) =
                            optional_attribute(&attributes, "contextual-spacing", part)?
                        {
                            style.contextual_spacing = matches!(value.as_str(), "true" | "1");
                        }
                        for (name, target) in [
                            ("margin-top", &mut style.margin_top),
                            ("margin-bottom", &mut style.margin_bottom),
                            ("margin-left", &mut style.margin_left),
                            ("margin-right", &mut style.margin_right),
                        ] {
                            if let Some(value) = optional_attribute(&attributes, name, part)? {
                                *target = parse_signed_length(&value)
                                    .ok_or_else(|| {
                                        format_error(part, format!("ODT {name} is invalid"))
                                    })?
                                    .max(0.0);
                            }
                        }
                        if let Some(indent) = optional_attribute(&attributes, "text-indent", part)?
                        {
                            style.text_indent = parse_signed_length(&indent)
                                .ok_or_else(|| format_error(part, "ODT text indent is invalid"))?;
                        }
                        if let Some(value) = optional_attribute(&attributes, "page-number", part)? {
                            style.page_number = value.parse::<u32>().ok();
                        }
                        style.page_break_before =
                            optional_attribute(&attributes, "break-before", part)?
                                .is_some_and(|value| value == "page");
                        style.page_break_after =
                            optional_attribute(&attributes, "break-after", part)?
                                .is_some_and(|value| value == "page");
                        style.keep_with_next =
                            optional_attribute(&attributes, "keep-with-next", part)?
                                .is_some_and(|value| matches!(value.as_str(), "always" | "true"));
                        style.keep_together =
                            optional_attribute(&attributes, "keep-together", part)?
                                .is_some_and(|value| matches!(value.as_str(), "always" | "true"));
                    }
                    "table-cell-properties" if current.is_some() => {
                        let cell = &mut current
                            .as_mut()
                            .ok_or_else(|| format_error(part, "ODT style state is missing"))?
                            .2
                            .cell;
                        if let Some(padding) = optional_length(&attributes, "padding", part)? {
                            cell.padding.fill(padding);
                        }
                        for (index, side) in ["top", "right", "bottom", "left"].iter().enumerate() {
                            if let Some(padding) =
                                optional_length(&attributes, &format!("padding-{side}"), part)?
                            {
                                cell.padding[index] = padding;
                            }
                        }
                        if let Some(background) =
                            optional_attribute(&attributes, "background-color", part)?
                            && background != "transparent"
                            && let Some(color) = parse_odf_color(&background)
                        {
                            cell.fill = color;
                        }
                        for name in [
                            "border",
                            "border-left",
                            "border-right",
                            "border-top",
                            "border-bottom",
                        ] {
                            if let Some(border) = optional_attribute(&attributes, name, part)? {
                                apply_odt_cell_border(&border, cell);
                                break;
                            }
                        }
                        if let Some(writing_mode) =
                            optional_attribute(&attributes, "writing-mode", part)?
                        {
                            cell.orientation = match writing_mode.as_str() {
                                "bt-lr" => TextOrientation::Rotated270,
                                "tb-rl" => TextOrientation::Rotated90,
                                "lr" | "lr-tb" | "rl" | "rl-tb" => TextOrientation::Horizontal,
                                _ => cell.orientation,
                            };
                        }
                    }
                    "table-row-properties" if current.is_some() => {
                        if let Some(min_height) =
                            optional_length(&attributes, "min-row-height", part)?
                        {
                            current
                                .as_mut()
                                .ok_or_else(|| format_error(part, "ODT style state is missing"))?
                                .2
                                .row
                                .min_height = min_height;
                        }
                    }
                    "table-column-properties" if current.is_some() => {
                        let column = &mut current
                            .as_mut()
                            .ok_or_else(|| format_error(part, "ODT style state is missing"))?
                            .2
                            .column;
                        if let Some(width) = optional_length(&attributes, "column-width", part)? {
                            column.width = Some(width);
                        }
                        if let Some(rel) =
                            optional_attribute(&attributes, "rel-column-width", part)?
                        {
                            column.rel_width =
                                Some(parse_rel_column_width(&rel).ok_or_else(|| {
                                    format_error(part, "ODT relative column width is invalid")
                                })?);
                        }
                    }
                    "graphic-properties" if current.is_some() => {
                        let graphic = &mut current
                            .as_mut()
                            .ok_or_else(|| format_error(part, "ODT style state is missing"))?
                            .2
                            .graphic;
                        if let Some(wrap) = optional_attribute(&attributes, "wrap", part)? {
                            graphic.wrap = match wrap.as_str() {
                                "run-through" => OdtWrap::RunThrough,
                                "none" => OdtWrap::NoWrap,
                                _ => OdtWrap::Parallel,
                            };
                        }
                        if let Some(margin_top) =
                            optional_attribute(&attributes, "margin-top", part)?
                        {
                            graphic.margin_top = parse_signed_length(&margin_top)
                                .ok_or_else(|| format_error(part, "ODT graphic margin is invalid"))?
                                .max(0.0);
                        }
                        if let Some(value) = optional_attribute(&attributes, "allow-overlap", part)?
                        {
                            graphic.allow_overlap = value != "false";
                        }
                        if let Some(value) =
                            optional_attribute(&attributes, "horizontal-rel", part)?
                        {
                            graphic.horizontal_rel = Some(value);
                        }
                        if let Some(value) = optional_attribute(&attributes, "vertical-rel", part)?
                        {
                            graphic.vertical_rel = Some(value);
                        }
                        for (name, position) in [
                            ("horizontal-pos", &mut graphic.horizontal_pos),
                            ("vertical-pos", &mut graphic.vertical_pos),
                        ] {
                            if let Some(value) = optional_attribute(&attributes, name, part)? {
                                *position = match value.as_str() {
                                    "right" | "bottom" => TextAlign::End,
                                    "center" | "middle" => TextAlign::Center,
                                    _ => TextAlign::Start,
                                };
                            }
                        }
                        if optional_attribute(&attributes, "fill", part)?.as_deref() == Some("none")
                        {
                            graphic.fill = Paint::None;
                        } else if let Some(color) =
                            optional_attribute(&attributes, "fill-color", part)?
                        {
                            graphic.fill =
                                Paint::Solid(parse_odf_color(&color).ok_or_else(|| {
                                    format_error(part, "ODT shape fill color is invalid")
                                })?);
                        }
                        if optional_attribute(&attributes, "stroke", part)?.is_some()
                            || optional_attribute(&attributes, "stroke-width", part)?.is_some()
                            || optional_attribute(&attributes, "border", part)?.is_some()
                        {
                            graphic.explicit_stroke = true;
                        }
                        if optional_attribute(&attributes, "stroke", part)?.as_deref()
                            == Some("none")
                        {
                            graphic.stroke = Paint::None;
                        } else if let Some(color) =
                            optional_attribute(&attributes, "stroke-color", part)?
                        {
                            graphic.stroke =
                                Paint::Solid(parse_odf_color(&color).ok_or_else(|| {
                                    format_error(part, "ODT shape stroke color is invalid")
                                })?);
                        }
                        if let Some(width) = optional_attribute(&attributes, "stroke-width", part)?
                        {
                            graphic.stroke_width = parse_length(&width)
                                .filter(|width| *width >= 0.0)
                                .ok_or_else(|| {
                                    format_error(part, "ODT shape stroke width is invalid")
                                })?;
                        }
                        graphic.marker_start =
                            optional_attribute(&attributes, "marker-start", part)?
                                .is_some_and(|marker| !marker.is_empty());
                        graphic.marker_end = optional_attribute(&attributes, "marker-end", part)?
                            .is_some_and(|marker| !marker.is_empty());
                    }
                    _ => {}
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if local_name(name) == "page-layout" {
                    page_layout = None;
                }
                if matches!(local_name(name), "header-style" | "footer-style") {
                    story_style = None;
                }
                if local_name(name) == "automatic-styles" {
                    automatic_styles = false;
                }
                if matches!(
                    local_name(name),
                    "list-level-style-bullet" | "list-level-style-number"
                ) {
                    current_list_level = None;
                }
                if local_name(name) == "list-level-style-bullet"
                    && current_list_bullet
                        .as_ref()
                        .is_some_and(|(bullet_depth, ..)| *bullet_depth == depth)
                {
                    let (_, level, character, font_family) = current_list_bullet
                        .take()
                        .ok_or_else(|| format_error(part, "ODT list bullet state is missing"))?;
                    current_list
                        .as_mut()
                        .ok_or_else(|| format_error(part, "ODT list style state is missing"))?
                        .2
                        .levels
                        .insert(
                            level,
                            OdtListLevel::Bullet {
                                character,
                                font_family,
                            },
                        );
                } else if local_name(name) == "list-style"
                    && current_list
                        .as_ref()
                        .is_some_and(|(style_depth, ..)| *style_depth == depth)
                {
                    let (_, name, style) = current_list
                        .take()
                        .ok_or_else(|| format_error(part, "ODT list style state is missing"))?;
                    catalog.list_styles.insert(name, style);
                } else if matches!(local_name(name), "style" | "default-style")
                    && current
                        .as_ref()
                        .is_some_and(|(style_depth, ..)| *style_depth == depth)
                {
                    let (_, name, style) = current
                        .take()
                        .ok_or_else(|| format_error(part, "ODT style state is missing"))?;
                    commit_odt_style(catalog, name, style);
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(())
}

fn commit_odt_style(catalog: &mut StyleCatalog, name: Option<String>, mut style: OdtStyle) {
    style
        .paragraph
        .tab_stops
        .sort_by(|a, b| a.position.total_cmp(&b.position));
    if let Some(name) = name {
        catalog.styles.insert(name, style);
    } else if style.family == "paragraph" {
        catalog.default_paragraph = style.paragraph;
    } else if style.family == "graphic" {
        catalog.default_graphic = style.graphic;
    } else if style.family == "table-cell" {
        catalog.default_cell = style.cell;
    } else if style.family == "table-row" {
        catalog.default_row = style.row;
    } else if style.family == "table-column" {
        catalog.default_table_column = style.column;
    }
}

fn select_initial_master_page(
    bytes: &[u8],
    limits: Limits,
    catalog: &mut StyleCatalog,
) -> Result<(), Diagnostic> {
    let mut depth = 0_usize;
    let mut text_depth = None;
    let mut first_paragraph_seen = false;
    let mut selected = None;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "text" && text_depth.is_none() {
                    text_depth = (!empty).then_some(depth);
                } else if matches!(local, "p" | "h")
                    && text_depth.is_some()
                    && !first_paragraph_seen
                {
                    first_paragraph_seen = true;
                    selected = optional_attribute(&attributes, "style-name", CONTENT_PART)?
                        .as_deref()
                        .and_then(|name| catalog.styles.get(name))
                        .and_then(|style| style.master_page_name.clone());
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if local_name(name) == "text" && text_depth == Some(depth) {
                    text_depth = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, CONTENT_PART))?;
    catalog.initial_master_page = selected.or_else(|| catalog.default_master_page.clone());
    Ok(())
}

fn parse_master_stories_xml(
    bytes: &[u8],
    limits: Limits,
    catalog: &mut StyleCatalog,
) -> Result<(), Diagnostic> {
    let mut depth = 0_usize;
    let mut master: Option<(usize, String, MasterPage, u32)> = None;
    let mut story: Option<(usize, &'static str)> = None;
    let mut paragraph: Option<(usize, MasterStoryParagraph)> = None;
    let mut shape: Option<ShapeState> = None;
    let mut page_field_depth = None;
    let mut page_field_slot: Option<&'static str> = None;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if let Some(current) = shape.as_mut() {
                    match local {
                        "enhanced-geometry" => {
                            if let Some(shape_type) =
                                optional_attribute(&attributes, "type", STYLES_PART)?
                            {
                                current.geometry = match shape_type.as_str() {
                                    "ellipse" | "circle" => Geometry::Ellipse,
                                    "line" => Geometry::Line,
                                    _ => Geometry::Rectangle,
                                };
                            }
                        }
                        "p" | "h" => {
                            if !current.text.is_empty() {
                                append_shape_text(current, "\n", limits.max_xml_bytes)?;
                            }
                            current.paragraph_depth = (!empty).then_some(depth);
                        }
                        "span" if current.paragraph_depth.is_some() => {
                            let inherited = current
                                .spans
                                .last()
                                .map_or(&current.text_style.text, |(_, style)| style);
                            let style_name =
                                optional_attribute(&attributes, "style-name", STYLES_PART)?;
                            let style = catalog.text(style_name.as_deref(), inherited);
                            if !empty {
                                current.spans.push((depth, style));
                            }
                        }
                        "s" if current.paragraph_depth.is_some() => {
                            let count = optional_positive_u32(&attributes, "c", STYLES_PART)?
                                .unwrap_or(1) as usize;
                            append_shape_repeated(current, ' ', count, limits.max_xml_bytes)?;
                        }
                        "tab" if current.paragraph_depth.is_some() => {
                            append_shape_text(current, "\t", limits.max_xml_bytes)?;
                        }
                        "line-break" if current.paragraph_depth.is_some() => {
                            append_shape_text(current, "\n", limits.max_xml_bytes)?;
                        }
                        _ => {}
                    }
                } else {
                    match local {
                        "master-page" if master.is_none() => {
                            let name = required_attribute(&attributes, "name", STYLES_PART)?;
                            if catalog.default_master_page.is_none() {
                                catalog.default_master_page = Some(name.clone());
                            }
                            let page = MasterPage {
                                page_layout: optional_attribute(
                                    &attributes,
                                    "page-layout-name",
                                    STYLES_PART,
                                )?,
                                next: optional_attribute(
                                    &attributes,
                                    "next-style-name",
                                    STYLES_PART,
                                )?,
                                ..MasterPage::default()
                            };
                            if empty {
                                catalog.master_pages.insert(name, page);
                            } else {
                                master = Some((depth, name, page, 0));
                            }
                        }
                        kind if master_story_kind(kind).is_some()
                            && master.is_some()
                            && story.is_none() =>
                        {
                            let kind = master_story_kind(kind).unwrap();
                            let current_master = master.as_mut().ok_or_else(|| {
                                format_error(STYLES_PART, "ODT master-page state is missing")
                            })?;
                            master_story_slot(&mut current_master.2, kind).present = true;
                            story = (!empty).then_some((depth, kind));
                        }
                        "custom-shape" | "rect" | "ellipse" | "line" | "connector" | "measure"
                        | "regular-polygon" | "caption"
                            if story.is_some() =>
                        {
                            let current_master = master.as_mut().ok_or_else(|| {
                                format_error(STYLES_PART, "ODT master-page state is missing")
                            })?;
                            let state = start_shape(
                                local,
                                &attributes,
                                catalog,
                                STYLES_PART,
                                depth,
                                current_master.3,
                            )?;
                            current_master.3 =
                                current_master.3.checked_add(1).ok_or_else(|| {
                                    object_limit_error(
                                        "ODT master shape count exceeds supported range",
                                    )
                                })?;
                            if empty {
                                push_master_shape(
                                    &mut current_master.2,
                                    story
                                        .ok_or_else(|| {
                                            format_error(STYLES_PART, "ODT master story is missing")
                                        })?
                                        .1,
                                    state,
                                );
                            } else {
                                shape = Some(state);
                            }
                        }
                        "table" if story.is_some() => {
                            let kind = story.unwrap().1;
                            master_story_slot(&mut master.as_mut().unwrap().2, kind).has_table =
                                true;
                        }
                        "p" | "h" if story.is_some() && paragraph.is_none() && !empty => {
                            let style_name =
                                optional_attribute(&attributes, "style-name", STYLES_PART)?;
                            let style = catalog.paragraph(style_name.as_deref());
                            paragraph = Some((
                                depth,
                                MasterStoryParagraph {
                                    text: String::new(),
                                    style,
                                    runs: Vec::new(),
                                    spans: Vec::new(),
                                },
                            ));
                        }
                        "span" if paragraph.is_some() => {
                            let (_, paragraph) = paragraph.as_mut().ok_or_else(|| {
                                format_error(
                                    STYLES_PART,
                                    "ODT master-page paragraph state is missing",
                                )
                            })?;
                            let style_name =
                                optional_attribute(&attributes, "style-name", STYLES_PART)?;
                            let inherited = paragraph
                                .spans
                                .last()
                                .map_or(&paragraph.style.text, |(_, style)| style);
                            let style = catalog.text(style_name.as_deref(), inherited);
                            if !empty {
                                paragraph.spans.push((depth, style));
                            }
                        }
                        "s" if paragraph.is_some() => {
                            let count = optional_positive_u32(&attributes, "c", STYLES_PART)?
                                .unwrap_or(1) as usize;
                            let (_, paragraph) = paragraph.as_mut().ok_or_else(|| {
                                format_error(
                                    STYLES_PART,
                                    "ODT master-page paragraph state is missing",
                                )
                            })?;
                            append_master_story_text(
                                paragraph,
                                &std::iter::repeat_n(' ', count).collect::<String>(),
                                limits.max_xml_bytes,
                            )?;
                        }
                        "tab" if paragraph.is_some() => {
                            let (_, paragraph) = paragraph.as_mut().ok_or_else(|| {
                                format_error(
                                    STYLES_PART,
                                    "ODT master-page paragraph state is missing",
                                )
                            })?;
                            append_master_story_text(paragraph, "\t", limits.max_xml_bytes)?;
                        }
                        "line-break" if paragraph.is_some() => {
                            let (_, paragraph) = paragraph.as_mut().ok_or_else(|| {
                                format_error(
                                    STYLES_PART,
                                    "ODT master-page paragraph state is missing",
                                )
                            })?;
                            append_master_story_text(paragraph, "\n", limits.max_xml_bytes)?;
                        }
                        other if paragraph.is_some() => {
                            let (_, paragraph) = paragraph.as_mut().ok_or_else(|| {
                                format_error(
                                    STYLES_PART,
                                    "ODT master-page paragraph state is missing",
                                )
                            })?;
                            if let Some(slot) = is_odf_dynamic_page_field(other) {
                                if empty {
                                    append_master_story_text(
                                        paragraph,
                                        slot,
                                        limits.max_xml_bytes,
                                    )?;
                                } else {
                                    page_field_depth = Some(depth);
                                    page_field_slot = Some(slot);
                                }
                            } else if empty
                                && let Some(display) =
                                    odf_empty_field_text(other, &attributes, STYLES_PART)?
                            {
                                append_master_story_text(
                                    paragraph,
                                    &display,
                                    limits.max_xml_bytes,
                                )?;
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
                let local = local_name(name);
                if page_field_depth == Some(depth) {
                    let slot = page_field_slot.take();
                    page_field_depth = None;
                    if let Some(slot) = slot
                        && let Some((_, paragraph)) = paragraph.as_mut()
                    {
                        append_master_story_text(paragraph, slot, limits.max_xml_bytes)?;
                    }
                    return Ok(());
                }
                if shape.is_some() {
                    match local {
                        "p" | "h" => {
                            if let Some(current) = shape.as_mut()
                                && current.paragraph_depth == Some(depth)
                            {
                                current.paragraph_depth = None;
                            }
                        }
                        "span" => {
                            if let Some(current) = shape.as_mut()
                                && current
                                    .spans
                                    .last()
                                    .is_some_and(|(span_depth, _)| *span_depth == depth)
                            {
                                current.spans.pop();
                            }
                        }
                        "custom-shape" | "rect" | "ellipse" | "line" | "connector"
                            if shape.as_ref().is_some_and(|current| current.depth == depth) =>
                        {
                            let finished = shape.take().ok_or_else(|| {
                                format_error(STYLES_PART, "ODT master shape state was lost")
                            })?;
                            let kind = story
                                .ok_or_else(|| {
                                    format_error(STYLES_PART, "ODT master story is missing")
                                })?
                                .1;
                            push_master_shape(
                                &mut master
                                    .as_mut()
                                    .ok_or_else(|| {
                                        format_error(
                                            STYLES_PART,
                                            "ODT master-page state is missing",
                                        )
                                    })?
                                    .2,
                                kind,
                                finished,
                            );
                        }
                        _ => {}
                    }
                } else {
                    if local == "span"
                        && let Some((_, paragraph)) = paragraph.as_mut()
                        && paragraph
                            .spans
                            .last()
                            .is_some_and(|(span_depth, _)| *span_depth == depth)
                    {
                        paragraph.spans.pop();
                    }
                    if matches!(local, "p" | "h")
                        && paragraph
                            .as_ref()
                            .is_some_and(|(paragraph_depth, _)| *paragraph_depth == depth)
                    {
                        let (_, text) = paragraph.take().ok_or_else(|| {
                            format_error(STYLES_PART, "ODT master-page paragraph state is missing")
                        })?;
                        if !text.text.trim().is_empty() {
                            let kind = story
                                .ok_or_else(|| {
                                    format_error(STYLES_PART, "ODT master story is missing")
                                })?
                                .1;
                            let page = &mut master
                                .as_mut()
                                .ok_or_else(|| {
                                    format_error(STYLES_PART, "ODT master-page state is missing")
                                })?
                                .2;
                            master_story_slot(page, kind).paragraphs.push(text);
                        }
                    }
                }
                if master_story_kind(local).is_some()
                    && story.is_some_and(|(story_depth, _)| story_depth == depth)
                {
                    story = None;
                } else if local == "master-page"
                    && master
                        .as_ref()
                        .is_some_and(|(master_depth, ..)| *master_depth == depth)
                {
                    let (_, name, page, _) = master.take().ok_or_else(|| {
                        format_error(STYLES_PART, "ODT master-page state was lost")
                    })?;
                    catalog.master_pages.insert(name, page);
                }
            }
            XmlEvent::Text(text) => {
                if page_field_depth.is_some() {
                    return Ok(());
                }
                if let Some(current) = shape
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    append_shape_text(
                        current,
                        &decode_xml_text(text).map_err(|error| with_part(error, STYLES_PART))?,
                        limits.max_xml_bytes,
                    )?;
                } else if let Some((_, paragraph)) = paragraph.as_mut() {
                    append_master_story_text(
                        paragraph,
                        &decode_xml_text(text).map_err(|error| with_part(error, STYLES_PART))?,
                        limits.max_xml_bytes,
                    )?;
                }
            }
            XmlEvent::Cdata(text) => {
                if page_field_depth.is_some() {
                    return Ok(());
                }
                if let Some(current) = shape
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    append_shape_text(current, text, limits.max_xml_bytes)?;
                } else if let Some((_, paragraph)) = paragraph.as_mut() {
                    append_master_story_text(paragraph, text, limits.max_xml_bytes)?;
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, STYLES_PART))?;
    if master.is_some() || story.is_some() || paragraph.is_some() || shape.is_some() {
        return Err(format_error(
            STYLES_PART,
            "ODT master-page content ends before its structure closes",
        ));
    }
    Ok(())
}

fn push_master_shape(page: &mut MasterPage, kind: &str, shape: ShapeState) {
    master_story_slot(page, kind).shapes.push(shape);
}

fn append_master_story_text(
    paragraph: &mut MasterStoryParagraph,
    value: &str,
    limit: usize,
) -> Result<(), Diagnostic> {
    let style = paragraph
        .spans
        .last()
        .map_or(&paragraph.style.text, |(_, style)| style);
    append_text_run(
        &mut paragraph.text,
        &mut paragraph.runs,
        style,
        value,
        limit,
    )
}

fn validate_page_layout(layout: PageLayout, limits: Limits) -> Result<(), Diagnostic> {
    let values = [
        layout.width,
        layout.height,
        layout.margin_top,
        layout.margin_right,
        layout.margin_bottom,
        layout.margin_left,
    ];
    if values
        .iter()
        .any(|value| !value.is_finite() || *value < 0.0)
        || layout.width <= layout.margin_left + layout.margin_right
        || layout.height <= layout.margin_top + layout.margin_bottom
    {
        return Err(format_error(STYLES_PART, "page layout is invalid"));
    }
    let pixels = (layout.width.ceil() as usize).saturating_mul(layout.height.ceil() as usize);
    if pixels > limits.max_render_pixels {
        return Err(Diagnostic::fatal(
            DiagnosticCode::LayoutBudgetExceeded,
            Phase::Layout,
            None,
            "ODT page size exceeds the configured render-pixel limit",
        )
        .in_part(STYLES_PART));
    }
    Ok(())
}

fn finish_cell(cell: CellState) -> TableCell {
    TableCell {
        tables: cell.tables,
        source_index: cell.source_index,
        first_column: cell.first_column,
        repeat: cell.repeat,
        element_id: cell.element_id,
        style: cell.style,
        text: cell.text,
        paragraph_style: cell.paragraph_style,
        column_span: cell.column_span,
        row_span: cell.row_span,
        image_hrefs: cell.image_hrefs,
    }
}

fn finish_row(
    table: &mut TableState,
    row: RowState,
    first_row: u32,
    limits: Limits,
    materialized_text_bytes: &mut usize,
) -> Result<(), Diagnostic> {
    let cells_per_row =
        row.cells.iter().try_fold(0_usize, |count, cell| {
            count
                .checked_add(usize::try_from(cell.repeat).map_err(|_| {
                    object_limit_error("ODT cell repetition exceeds supported range")
                })?)
                .ok_or_else(|| object_limit_error("ODT repeated cell count overflow"))
        })?;
    let added_cells = cells_per_row
        .checked_mul(
            usize::try_from(row.repeat)
                .map_err(|_| object_limit_error("ODT row repetition exceeds supported range"))?,
        )
        .ok_or_else(|| object_limit_error("ODT repeated table object count overflow"))?;
    table.projected_cells = table
        .projected_cells
        .checked_add(added_cells)
        .filter(|count| *count <= limits.max_document_objects)
        .ok_or_else(|| {
            object_limit_error("ODT repeated cells exceed the configured object limit")
        })?;
    let repeated_rows = usize::try_from(row.repeat).map_err(|_| {
        Diagnostic::fatal(
            DiagnosticCode::ZipTotalSizeLimit,
            Phase::Parse,
            None,
            "materialized repeated row count exceeds supported range",
        )
        .in_part(CONTENT_PART)
    })?;
    for cell in &row.cells {
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
            cell.text.len(),
            copies,
            limits.max_total_uncompressed_bytes,
            CONTENT_PART,
        )?;
    }
    table.rows.push(TableRow {
        source_index: row.source_index,
        first_row,
        repeat: row.repeat,
        columns: row.next_column,
        min_height: row.min_height,
        cells: row.cells,
    });
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn layout_table(
    table: &TableState,
    layout: PageLayout,
    units: &mut Vec<Unit>,
    unit_index: &mut u32,
    y: &mut f32,
    objects: &mut Vec<Object>,
    diagnostics: &mut Vec<Diagnostic>,
    limits: Limits,
    styles: &StyleCatalog,
    split_reported: &mut bool,
    clipped_reported: &mut bool,
    table_metrics_reported: &mut bool,
    package: Option<&Package<'_>>,
    materialized_image_bytes: &mut usize,
) -> Result<(), Diagnostic> {
    let maximum_table_parents = usize::try_from(table.logical_rows.max(1))
        .map_err(|_| object_limit_error("ODT logical row count exceeds supported range"))?;
    if objects
        .len()
        .checked_add(table.projected_cells)
        .and_then(|count| count.checked_add(maximum_table_parents))
        .is_none_or(|count| count > limits.max_document_objects)
    {
        return Err(object_limit_error(
            "ODT table expansion exceeds the configured object limit",
        ));
    }

    let column_count = table
        .rows
        .iter()
        .map(|row| row.columns)
        .max()
        .unwrap_or(0)
        .max(table.declared_columns)
        .max(table.column_styles.len() as u32)
        .max(1);
    let content_width = layout.width - layout.margin_left - layout.margin_right;
    let hints = (0..column_count)
        .map(|index| {
            table
                .column_styles
                .get(index as usize)
                .and_then(|style| style.size_hint())
                .unwrap_or(0.0)
        })
        .collect::<Vec<_>>();
    let authored_column_widths = hints.iter().any(|hint| *hint > 0.0);
    let column_sizes =
        super::normalized_odf_table_sizes(&hints, column_count as usize, content_width);
    if column_sizes
        .iter()
        .any(|width| !width.is_finite() || *width <= 0.0)
    {
        return Err(Diagnostic::fatal(
            DiagnosticCode::LayoutBudgetExceeded,
            Phase::Layout,
            None,
            "ODT table columns cannot be laid out within the page",
        )
        .in_part(CONTENT_PART));
    }
    if !authored_column_widths && !*table_metrics_reported {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ApproximateLayout,
                Phase::Layout,
                Fidelity::Approximate,
                "ODT table column widths use deterministic equal-width fallback layout",
            )
            .in_part(CONTENT_PART),
        );
        *table_metrics_reported = true;
    }
    let mut logical_row_heights = Vec::with_capacity(table.logical_rows as usize);
    for row in &table.rows {
        let height = table_row_height(row, &column_sizes, styles.font_metrics);
        for _ in 0..row.repeat {
            logical_row_heights.push(height);
        }
    }

    let table_path = format!(
        "/office:document-content/office:body/office:text/table:table[{}]",
        table.index + 1
    );
    let mut fragment_index = 0_u32;
    let mut fragment: Option<(usize, u32, String, f32)> = None;

    if table.rows.is_empty() {
        ensure_table_row_space(
            LINE_HEIGHT,
            false,
            layout,
            units,
            unit_index,
            y,
            &mut fragment,
            split_reported,
            diagnostics,
            objects,
            styles,
            limits.max_document_objects,
        )?;
        let (object_index, _, _, start_y) = start_table_fragment(
            &table,
            &table_path,
            fragment_index,
            *unit_index,
            Rect {
                x: layout.margin_left,
                y: *y,
                width: content_width,
                height: 0.0,
            },
            objects,
        )?;
        *y += LINE_HEIGHT;
        objects[object_index].bounds.height = *y - start_y;
        return Ok(());
    }

    for row in &table.rows {
        let measured_height = table_row_height(row, &column_sizes, styles.font_metrics);
        let content_height = layout.height - layout.margin_top - layout.margin_bottom;
        let row_height = measured_height.min(content_height);
        if measured_height > content_height && !*clipped_reported {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::LayoutBudgetExceeded,
                    Phase::Layout,
                    Fidelity::Approximate,
                    "an ODT table row is taller than a page and was clipped; cell text remains available through object mapping",
                )
                .in_part(CONTENT_PART),
            );
            *clipped_reported = true;
        }

        for row_offset in 0..row.repeat {
            let had_fragment = fragment.is_some();
            let moved_page = ensure_table_row_space(
                row_height,
                had_fragment,
                layout,
                units,
                unit_index,
                y,
                &mut fragment,
                split_reported,
                diagnostics,
                objects,
                styles,
                limits.max_document_objects,
            )?;
            if moved_page && had_fragment {
                fragment_index = fragment_index.checked_add(1).ok_or_else(|| {
                    object_limit_error("ODT table fragment count exceeds supported range")
                })?;
            }

            let (table_object_index, table_numeric_id, table_stable_id, fragment_y) =
                if let Some((index, numeric_id, stable_id, start_y)) = &fragment {
                    (*index, *numeric_id, stable_id.clone(), *start_y)
                } else {
                    let started = start_table_fragment(
                        &table,
                        &table_path,
                        fragment_index,
                        *unit_index,
                        Rect {
                            x: layout.margin_left,
                            y: *y,
                            width: content_width,
                            height: 0.0,
                        },
                        objects,
                    )?;
                    fragment = Some(started.clone());
                    started
                };
            let logical_row = row.first_row.checked_add(row_offset).ok_or_else(|| {
                object_limit_error("ODT logical row position exceeds supported range")
            })?;
            let row_y = *y;

            for cell in &row.cells {
                for column_offset in 0..cell.repeat {
                    let logical_column =
                        cell.first_column
                            .checked_add(column_offset)
                            .ok_or_else(|| {
                                object_limit_error(
                                    "ODT logical column position exceeds supported range",
                                )
                            })?;
                    let cell_object_index = objects.len();
                    let cell_x = layout.margin_left
                        + column_sizes
                            .iter()
                            .take(logical_column as usize)
                            .sum::<f32>();
                    let cell_width =
                        column_span_width(&column_sizes, logical_column, cell.column_span);
                    push_table_cell(
                        &table,
                        row,
                        cell,
                        logical_row,
                        logical_column,
                        *unit_index,
                        Rect {
                            x: cell_x,
                            y: row_y,
                            width: cell_width,
                            height: logical_row_heights
                                .iter()
                                .skip(logical_row as usize)
                                .take(cell.row_span as usize)
                                .sum::<f32>()
                                .min(layout.height - layout.margin_bottom - row_y)
                                .max(1.0),
                        },
                        table_numeric_id,
                        &table_stable_id,
                        objects,
                        styles.font_metrics,
                    )?;
                    push_table_cell_images(
                        cell,
                        Rect {
                            x: cell_x,
                            y: row_y,
                            width: cell_width,
                            height: logical_row_heights
                                .iter()
                                .skip(logical_row as usize)
                                .take(cell.row_span as usize)
                                .sum::<f32>()
                                .min(layout.height - layout.margin_bottom - row_y)
                                .max(1.0),
                        },
                        *unit_index,
                        table_numeric_id,
                        &table_stable_id,
                        package,
                        limits,
                        objects,
                        diagnostics,
                        materialized_image_bytes,
                    )?;
                    let nested_layout = PageLayout {
                        margin_left: cell_x + 2.0,
                        margin_right: layout.width - cell_x - cell_width + 2.0,
                        ..layout
                    };
                    let mut nested_y = row_y + 2.0;
                    for nested in &cell.tables {
                        let start = objects.len();
                        layout_table(
                            nested,
                            nested_layout,
                            units,
                            unit_index,
                            &mut nested_y,
                            objects,
                            diagnostics,
                            limits,
                            styles,
                            split_reported,
                            clipped_reported,
                            table_metrics_reported,
                            package,
                            materialized_image_bytes,
                        )?;
                        let parent_id = objects[cell_object_index].numeric_id;
                        let parent_stable_id = objects[cell_object_index].stable_id.clone();
                        for object in &mut objects[start..] {
                            if object.parent_numeric_id.is_none() {
                                object.parent_numeric_id = Some(parent_id);
                                object.parent_stable_id = Some(parent_stable_id.clone());
                            }
                        }
                    }
                }
            }

            *y += row_height;
            objects[table_object_index].bounds.height = *y - fragment_y;
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn ensure_table_row_space(
    height: f32,
    table_has_fragment: bool,
    layout: PageLayout,
    units: &mut Vec<Unit>,
    unit_index: &mut u32,
    y: &mut f32,
    fragment: &mut Option<(usize, u32, String, f32)>,
    split_reported: &mut bool,
    diagnostics: &mut Vec<Diagnostic>,
    objects: &mut Vec<Object>,
    styles: &StyleCatalog,
    object_limit: usize,
) -> Result<bool, Diagnostic> {
    if *y + height <= layout.height - layout.margin_bottom || *y <= layout.margin_top {
        return Ok(false);
    }
    *unit_index = unit_index
        .checked_add(1)
        .ok_or_else(|| format_error(CONTENT_PART, "page count exceeds the supported range"))?;
    units.push(page_unit(*unit_index, layout));
    *y = push_master_stories(*unit_index, layout, objects, styles, object_limit)?;
    *fragment = None;
    if table_has_fragment && !*split_reported {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ApproximateLayout,
                Phase::Layout,
                Fidelity::Approximate,
                "an ODT table was split into page-local fragments at row boundaries",
            )
            .in_part(CONTENT_PART),
        );
        *split_reported = true;
    }
    Ok(true)
}

fn start_table_fragment(
    table: &TableState,
    table_path: &str,
    fragment_index: u32,
    unit_index: u32,
    bounds: Rect,
    objects: &mut Vec<Object>,
) -> Result<(usize, u32, String, f32), Diagnostic> {
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| object_limit_error("ODT object count exceeds supported range"))?;
    let object_index = numeric_id as usize;
    let stable_id = format!("odt:table:{}:fragment:{fragment_index}", table.index);
    objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: stable_id.clone(),
        parent_stable_id: None,
        kind: ObjectKind::Table,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping: MappingQuality::Derived,
            locator: SourceLocator::Odt {
                kind: "element",
                element_id: table.element_id.clone(),
                path: table_path.to_owned(),
                row: None,
                column: None,
                text_range: None,
            },
        },
        visual: Visual::Shape {
            geometry: Geometry::Rectangle,
            fill: 0x0000_0000,
            stroke: 0x0000_0000,
            stroke_width: 0.0,
        },
    });
    Ok((object_index, numeric_id, stable_id, bounds.y))
}

#[allow(clippy::too_many_arguments)]
fn push_table_cell(
    table: &TableState,
    row: &TableRow,
    cell: &TableCell,
    logical_row: u32,
    logical_column: u32,
    unit_index: u32,
    bounds: Rect,
    table_numeric_id: u32,
    table_stable_id: &str,
    objects: &mut Vec<Object>,
    metrics: Option<&FontMetricTable>,
) -> Result<(), Diagnostic> {
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| object_limit_error("ODT object count exceeds supported range"))?;
    let stable_id = format!(
        "odt:table:{}:row:{logical_row}:column:{logical_column}",
        table.index
    );
    let text_length = u32::try_from(cell.text.chars().count()).unwrap_or(u32::MAX);
    let path = format!(
        "/office:document-content/office:body/office:text/table:table[{}]/table:table-row[{}]/table:table-cell[{}]",
        table.index + 1,
        row.source_index + 1,
        cell.source_index + 1,
    );
    let visual = Visual::Text {
        geometry: Geometry::Rectangle,
        fill: cell.style.fill,
        stroke: cell.style.stroke,
        stroke_width: cell.style.stroke_width,
        font_family: cell.paragraph_style.text.font_family.clone(),
        font_size: cell.paragraph_style.text.font_size,
        color: cell.paragraph_style.text.color,
        bold: cell.paragraph_style.text.bold,
        italic: cell.paragraph_style.text.italic,
        align: cell.paragraph_style.align,
    };
    let style = &cell.paragraph_style;
    let visual = Visual::TextLayout {
        layout: TextLayout {
            orientation: cell.style.orientation,
            inset_top: cell.style.padding[0] + style.margin_top,
            inset_right: cell.style.padding[1],
            inset_bottom: cell.style.padding[2] + style.margin_bottom,
            inset_left: cell.style.padding[3],
            paragraphs: vec![crate::model::TextParagraphLayout {
                align: style.align,
                margin_left: style.margin_left,
                margin_right: style.margin_right,
                first_line_indent: style.text_indent,
                line_height: effective_odt_line_height(style, &[], metrics),
                default_tab_stop: style.default_tab_stop,
                space_before: 0.0,
                space_after: 0.0,
                latin_line_break: true,
                hanging_punctuation: false,
                rule_above: None,
                rule_below: None,
                drop_cap: None,
            }],
            ..TextLayout::default()
        },
        visual: Box::new(visual),
    };
    let visual = if cell.style.stroke_style == StrokeStyle::default() {
        visual
    } else {
        Visual::StrokeStyle {
            style: cell.style.stroke_style.clone(),
            visual: Box::new(visual),
        }
    };
    objects.push(Object {
        numeric_id,
        parent_numeric_id: Some(table_numeric_id),
        stable_id,
        parent_stable_id: Some(table_stable_id.to_owned()),
        kind: ObjectKind::Cell,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: Some(clone_materialized_text(&cell.text, CONTENT_PART)?),
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping: if row.repeat != 1 || cell.repeat != 1 {
                MappingQuality::Derived
            } else {
                MappingQuality::Exact
            },
            locator: SourceLocator::Odt {
                kind: "table-cell",
                element_id: cell.element_id.clone(),
                path,
                row: Some(logical_row),
                column: Some(logical_column),
                text_range: Some((0, text_length)),
            },
        },
        visual,
    });
    Ok(())
}

fn push_table_cell_images(
    cell: &TableCell,
    bounds: Rect,
    unit_index: u32,
    table_numeric_id: u32,
    table_stable_id: &str,
    package: Option<&Package<'_>>,
    limits: Limits,
    objects: &mut Vec<Object>,
    diagnostics: &mut Vec<Diagnostic>,
    materialized_image_bytes: &mut usize,
) -> Result<(), Diagnostic> {
    if cell.image_hrefs.is_empty() {
        return Ok(());
    }
    let inset = 2.0;
    let image_bounds = Rect {
        x: bounds.x + inset,
        y: bounds.y + inset,
        width: (bounds.width - inset * 2.0).max(1.0),
        height: (bounds.height - inset * 2.0).max(1.0),
    };
    let Some(package) = package else {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Approximate,
                "ODT table image bytes are unavailable in parser-only mode; a mapped placeholder is rendered",
            )
            .in_part(CONTENT_PART),
        );
        return Ok(());
    };
    let mut image_visual = None;
    for href in &cell.image_hrefs {
        let target = match resolve_odf_image_target(href, limits.max_zip_path_bytes) {
            Ok(OdfImageTarget::Embedded(target)) => target,
            Ok(OdfImageTarget::External) => {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::ExternalResourceBlocked,
                        Phase::Security,
                        Fidelity::Blocked,
                        "external ODT table image was blocked",
                    )
                    .in_part(CONTENT_PART),
                );
                continue;
            }
            Err(message) => {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Parse,
                        Fidelity::Approximate,
                        format!("invalid ODT table image reference: {message}"),
                    )
                    .in_part(CONTENT_PART),
                );
                continue;
            }
        };
        let Some(bytes) = package.part(&target)? else {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Container,
                    Fidelity::Omitted,
                    "ODT table image is missing from the package",
                )
                .in_part(&target),
            );
            continue;
        };
        let declared = office_image_media_type(&target, &bytes);
        let media_type =
            match super::presentation_image::recover_office_image_signature(declared, &bytes) {
                Ok(media_type) => media_type,
                Err(error) => {
                    let reason = match error {
                        OfficeImageError::UnsupportedFormat => "unsupported image format",
                        OfficeImageError::SignatureMismatch => "image signature mismatch",
                        OfficeImageError::DisabledByOffice => "format disabled by Office",
                    };
                    diagnostics.push(
                        Diagnostic::warning(
                            DiagnosticCode::UnsupportedFeature,
                            Phase::Render,
                            Fidelity::Approximate,
                            format!("ODT table image uses an {reason}"),
                        )
                        .in_part(&target),
                    );
                    continue;
                }
            };
        reserve_materialized_image_bytes(
            materialized_image_bytes,
            bytes.len(),
            limits.max_total_uncompressed_bytes,
            &target,
        )?;
        image_visual = Some(Visual::Image {
            media_type: media_type.to_owned(),
            bytes: clone_image_bytes(&bytes, &target)?,
            crop: ImageCrop::default(),
        });
        break;
    }
    let Some(visual) = image_visual else {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Approximate,
                "no supported ODT table image alternative; a mapped placeholder is rendered",
            )
            .in_part(CONTENT_PART),
        );
        return Ok(());
    };
    if objects.len() >= limits.max_document_objects {
        return Err(object_limit_error(
            "ODT table image count exceeds the configured object limit",
        ));
    }
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| object_limit_error("ODT object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: Some(table_numeric_id),
        stable_id: format!("odt:table:image:{numeric_id}"),
        parent_stable_id: Some(table_stable_id.to_owned()),
        kind: ObjectKind::Image,
        unit_index,
        bounds: image_bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping: MappingQuality::Exact,
            locator: SourceLocator::Odt {
                kind: "table-cell",
                element_id: cell.element_id.clone(),
                path: format!(
                    "/office:document-content/office:body/office:text/table:table[1]/table:table-row[1]/table:table-cell[1]/draw:image[1]"
                ),
                row: None,
                column: None,
                text_range: None,
            },
        },
        visual,
    });
    Ok(())
}

fn column_span_width(column_sizes: &[f32], start: u32, span: u32) -> f32 {
    column_sizes
        .iter()
        .skip(start as usize)
        .take(span.max(1) as usize)
        .sum::<f32>()
        .max(1.0)
}

fn table_row_height(
    row: &TableRow,
    column_sizes: &[f32],
    metrics: Option<&FontMetricTable>,
) -> f32 {
    row.cells
        .iter()
        .map(|cell| {
            let width = column_span_width(column_sizes, cell.first_column, cell.column_span);
            let nested_height: f32 = cell
                .tables
                .iter()
                .map(|table| {
                    let columns = table
                        .declared_columns
                        .max(table.rows.iter().map(|r| r.columns).max().unwrap_or(1))
                        .max(1);
                    let nested_width = ((width - 4.0) / columns as f32).max(1.0);
                    let nested_sizes = vec![nested_width; columns as usize];
                    table
                        .rows
                        .iter()
                        .map(|r| table_row_height(r, &nested_sizes, metrics) * r.repeat as f32)
                        .sum::<f32>()
                        + 4.0
                })
                .sum();
            let style = &cell.paragraph_style;
            let text_style = &style.text;
            let characters: Vec<_> = cell
                .text
                .chars()
                .map(|c| {
                    (
                        c,
                        metrics
                            .and_then(|m| {
                                m.advance_em_at_size(
                                    &text_style.font_family,
                                    text_style.italic,
                                    text_style.bold,
                                    c,
                                    text_style.font_size,
                                )
                            })
                            .unwrap_or(0.5)
                            * text_style.font_size
                            + text_style.letter_spacing,
                    )
                })
                .collect();
            let available = (width
                - cell.style.padding[1]
                - cell.style.padding[3]
                - style.margin_left
                - style.margin_right)
                .max(1.0);
            // Cell content is currently one font style; paragraph runs use the same shared fitter.
            let lines = crate::text_layout::horizontal_line_breaks(
                &characters,
                (available - style.text_indent).max(1.0),
                available,
                style.default_tab_stop,
                false,
            )
            .len()
            .max(1);
            let text_height = lines as f32 * effective_odt_line_height(style, &[], metrics)
                + style.margin_top
                + style.margin_bottom
                + cell.style.padding[0]
                + cell.style.padding[2];
            text_height.max(nested_height)
        })
        .fold(LINE_HEIGHT, f32::max)
        .max(row.min_height)
}

fn bounded_repeat(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    maximum: usize,
) -> Result<u32, Diagnostic> {
    let Some(value) = optional_attribute(attributes, name, CONTENT_PART)? else {
        return Ok(1);
    };
    if value.starts_with('0') || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(format_error(
            CONTENT_PART,
            format!("attribute {name} must be a canonical positive integer"),
        ));
    }
    let repeat = value
        .parse::<u32>()
        .ok()
        .filter(|repeat| *repeat != 0)
        .ok_or_else(|| format_error(CONTENT_PART, format!("attribute {name} is invalid")))?;
    if usize::try_from(repeat).map_or(true, |repeat| repeat > maximum) {
        return Err(object_limit_error(format!(
            "attribute {name} exceeds the configured expansion limit"
        )));
    }
    Ok(repeat)
}

fn object_limit_error(message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::ObjectLimit, Phase::Parse, None, message)
        .in_part(CONTENT_PART)
}

#[allow(clippy::too_many_arguments)]
fn section_columns_fallback() -> Diagnostic {
    Diagnostic::warning(DiagnosticCode::ApproximateLayout, Phase::Layout, Fidelity::Approximate,
        "ODT section columns requiring nested content, unequal widths, paragraph splitting or pagination use single-column flow")
        .in_part(CONTENT_PART)
}

#[allow(clippy::too_many_arguments)]
fn push_section_paragraphs(
    paragraphs: Vec<ParagraphState>,
    (count, gap): (u32, f32),
    layout: PageLayout,
    units: &mut Vec<Unit>,
    unit_index: &mut u32,
    y: &mut f32,
    objects: &mut Vec<Object>,
    styles: &StyleCatalog,
    limits: Limits,
    materialized_text_bytes: &mut usize,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<(), Diagnostic> {
    let width = (layout.width
        - layout.margin_left
        - layout.margin_right
        - gap * count.saturating_sub(1) as f32)
        / count.max(1) as f32;
    let available = layout.height - layout.margin_bottom - *y;
    // ponytail: balance whole paragraphs on one page; diagnose and retain flow
    // when line-fragment or multi-page column layout is required.
    let heights: Vec<_> = paragraphs
        .iter()
        .map(|p| {
            let line_height = effective_odt_line_height(&p.style, &p.runs, styles.font_metrics);
            let text_width = width
                - p.style.margin_left
                - p.style.margin_right
                - p.list_position
                    .as_ref()
                    .map_or(0.0, |position| position.margin_left);
            let height =
                styled_paragraph_height(p, text_width.max(1.0), line_height, styles.font_metrics);
            (
                height,
                height + p.style.margin_top + p.style.margin_bottom,
                line_height,
            )
        })
        .collect();
    let eligible = count > 1
        && width > 0.0
        && paragraphs.len() >= count as usize
        && paragraphs
            .iter()
            .zip(&heights)
            .all(|(p, &(height, _, line))| {
                !p.style.page_break_before
                    && !p.style.page_break_after
                    && !p.style.keep_with_next
                    && p.soft_page_breaks.is_empty()
                    && height <= line + 0.001
            });
    let mut target = heights.iter().map(|(_, h, _)| *h).sum::<f32>();
    if eligible {
        let mut low = heights.iter().map(|(_, h, _)| *h).fold(0.0, f32::max);
        // Minimize the tallest contiguous column without splitting a paragraph.
        for _ in 0..24 {
            let mid = (low + target) / 2.0;
            let (mut used, mut columns) = (0.0, 1);
            for &(_, height, _) in &heights {
                if used > 0.0 && used + height > mid {
                    columns += 1;
                    used = 0.0;
                }
                used += height;
            }
            if columns <= count {
                target = mid;
            } else {
                low = mid;
            }
        }
    }
    if !eligible || target > available {
        if count > 1 {
            diagnostics.push(section_columns_fallback());
        }
        for paragraph in paragraphs {
            push_paragraph(
                paragraph,
                layout,
                units,
                unit_index,
                y,
                objects,
                styles,
                limits.max_document_objects,
                materialized_text_bytes,
                limits.max_total_uncompressed_bytes,
            )?;
        }
        return Ok(());
    }
    let top = *y;
    let (mut column, mut used, mut bottom) = (0, 0.0, top);
    for (paragraph, &(_, height, _)) in paragraphs.into_iter().zip(&heights) {
        if used > 0.0 && used + height > target + 0.001 && column + 1 < count {
            column += 1;
            used = 0.0;
        }
        let left = layout.margin_left + column as f32 * (width + gap);
        let column_layout = PageLayout {
            margin_left: left,
            margin_right: layout.width - left - width,
            ..layout
        };
        *y = top + used;
        push_paragraph(
            paragraph,
            column_layout,
            units,
            unit_index,
            y,
            objects,
            styles,
            limits.max_document_objects,
            materialized_text_bytes,
            limits.max_total_uncompressed_bytes,
        )?;
        used += height;
        bottom = bottom.max(*y);
    }
    *y = bottom;
    Ok(())
}

fn push_paragraph(
    mut paragraph: ParagraphState,
    layout: PageLayout,
    units: &mut Vec<Unit>,
    unit_index: &mut u32,
    y: &mut f32,
    objects: &mut Vec<Object>,
    styles: &StyleCatalog,
    object_limit: usize,
    materialized_text_bytes: &mut usize,
    materialized_text_limit: usize,
) -> Result<(), Diagnostic> {
    if paragraph.soft_page_breaks.is_empty() {
        return push_paragraph_segment(
            paragraph,
            layout,
            units,
            unit_index,
            y,
            objects,
            styles,
            object_limit,
            materialized_text_bytes,
            materialized_text_limit,
            false,
        );
    }
    let mut boundaries = std::mem::take(&mut paragraph.soft_page_breaks);
    boundaries.push(paragraph.text.len());
    let mut start = 0_usize;
    let mut source_start = paragraph.source_text_start;
    let mut list_label = paragraph.list_label.take();
    let segment_count = boundaries.len();
    for (segment_index, end) in boundaries.into_iter().enumerate() {
        if end < start || end > paragraph.text.len() || !paragraph.text.is_char_boundary(end) {
            return Err(format_error(
                CONTENT_PART,
                "ODT soft page break is not on a text boundary",
            ));
        }
        if segment_index > 0 && *y > layout.margin_top {
            *unit_index = unit_index.checked_add(1).ok_or_else(|| {
                format_error(CONTENT_PART, "page count exceeds the supported range")
            })?;
            units.push(page_unit(*unit_index, layout));
            *y = push_master_stories(*unit_index, layout, objects, styles, object_limit)?;
        }
        // A leading cached break has no text fragment before it.
        if start == end && segment_index + 1 < segment_count {
            continue;
        }
        let text = paragraph.text[start..end].to_owned();
        let text_length = u32::try_from(text.chars().count()).unwrap_or(u32::MAX);
        let mut style = paragraph.style.clone();
        style.page_break_before &= segment_index == 0;
        style.page_break_after &= segment_index + 1 == segment_count;
        let segment = ParagraphState {
            depth: paragraph.depth,
            index: paragraph.index,
            element_id: paragraph.element_id.clone(),
            text,
            style,
            runs: text_runs_in_range(&paragraph.runs, start, end),
            spans: Vec::new(),
            soft_page_breaks: Vec::new(),
            source_text_start: source_start,
            list_label: list_label.take(),
            list_position: paragraph.list_position.clone(),
            frames: Vec::new(),
        };
        push_paragraph_segment(
            segment,
            layout,
            units,
            unit_index,
            y,
            objects,
            styles,
            object_limit,
            materialized_text_bytes,
            materialized_text_limit,
            true,
        )?;
        source_start = source_start.saturating_add(text_length);
        start = end;
    }
    Ok(())
}

fn text_runs_in_range(runs: &[TextRun], start: usize, end: usize) -> Vec<TextRun> {
    let mut offset = 0_usize;
    runs.iter()
        .filter_map(|run| {
            let run_start = offset;
            let run_end = offset.saturating_add(run.text.len());
            offset = run_end;
            let overlap_start = start.max(run_start);
            let overlap_end = end.min(run_end);
            (overlap_start < overlap_end).then(|| {
                let mut segment = run.clone();
                segment.text =
                    run.text[overlap_start - run_start..overlap_end - run_start].to_owned();
                segment
            })
        })
        .collect()
}

#[allow(clippy::too_many_arguments)]
fn push_paragraph_segment(
    mut paragraph: ParagraphState,
    layout: PageLayout,
    units: &mut Vec<Unit>,
    unit_index: &mut u32,
    y: &mut f32,
    objects: &mut Vec<Object>,
    styles: &StyleCatalog,
    object_limit: usize,
    materialized_text_bytes: &mut usize,
    materialized_text_limit: usize,
    soft_page_segment: bool,
) -> Result<(), Diagnostic> {
    if objects.len() >= object_limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "document exceeds the configured object limit",
        )
        .in_part(CONTENT_PART));
    }
    if paragraph.style.page_break_before && *y > layout.margin_top {
        *unit_index = unit_index
            .checked_add(1)
            .ok_or_else(|| format_error(CONTENT_PART, "page count exceeds the supported range"))?;
        units.push(page_unit(*unit_index, layout));
        *y = push_master_stories(*unit_index, layout, objects, styles, object_limit)?;
    }
    let margin_left = if paragraph.list_position.is_some() {
        0.0
    } else {
        paragraph.style.margin_left
    };
    let x = layout.margin_left + margin_left;
    let width = layout.width
        - layout.margin_left
        - layout.margin_right
        - margin_left
        - paragraph.style.margin_right;
    let width = width.max(1.0);
    if let Some(position) = &paragraph.list_position {
        paragraph.style.margin_left = position.margin_left;
        paragraph.style.text_indent = if paragraph.list_label.is_some() {
            position.text_indent
        } else {
            0.0
        };
        if let Some(position) = position.tab_stop {
            paragraph.style.tab_stops.insert(
                0,
                TextTabStop {
                    position,
                    align: TextAlign::Start,
                    leader: TextTabLeader::None,
                },
            );
        }
    }
    if paragraph.runs.is_empty() && !paragraph.text.is_empty() {
        paragraph.runs.push(TextRun {
            paint: None,
            east_asian_line_breaks: true,
            text: paragraph.text.clone(),
            font_family: paragraph.style.text.font_family.clone(),
            font_size: paragraph.style.text.font_size,
            color: paragraph.style.text.color,
            bold: paragraph.style.text.bold,
            italic: paragraph.style.text.italic,
            underline: paragraph.style.text.underline,
            strikethrough: paragraph.style.text.strikethrough,
            highlight: paragraph.style.text.highlight,
            baseline_shift: paragraph.style.text.baseline_shift,
            letter_spacing: paragraph.style.text.letter_spacing,
            horizontal_scale: 1.0,
        });
    }
    let line_height =
        effective_odt_line_height(&paragraph.style, &paragraph.runs, styles.font_metrics);
    let mut height = styled_paragraph_height(
        &paragraph,
        width
            - paragraph
                .list_position
                .as_ref()
                .map_or(0.0, |p| p.margin_left),
        line_height,
        styles.font_metrics,
    );
    let usable_height = layout.height - layout.margin_top - layout.margin_bottom;
    if soft_page_segment {
        height = height.min(
            (layout.height
                - layout.margin_bottom
                - *y
                - paragraph.style.margin_top
                - paragraph.style.margin_bottom)
                .max(line_height),
        );
    } else if height > usable_height {
        return Err(Diagnostic::fatal(
            DiagnosticCode::LayoutBudgetExceeded,
            Phase::Layout,
            None,
            "a single ODT paragraph exceeds the usable page height",
        )
        .in_part(CONTENT_PART));
    }
    if *y == layout.margin_top {
        paragraph.style.margin_top = 0.0;
    }
    // Trailing paragraph spacing may extend past the page; only its text must fit.
    let required_height = paragraph.style.margin_top + height;
    if !soft_page_segment
        && *y + required_height > layout.height - layout.margin_bottom
        && *y > layout.margin_top
    {
        *unit_index = unit_index
            .checked_add(1)
            .ok_or_else(|| format_error(CONTENT_PART, "page count exceeds the supported range"))?;
        units.push(page_unit(*unit_index, layout));
        *y = push_master_stories(*unit_index, layout, objects, styles, object_limit)?;
        paragraph.style.margin_top = 0.0;
    }
    *y += paragraph.style.margin_top;

    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    let text_length = u32::try_from(paragraph.text.chars().count()).unwrap_or(u32::MAX);
    let text_end = paragraph.source_text_start.saturating_add(text_length);
    let path = format!(
        "/office:document-content/office:body/office:text/text:p[{}]",
        paragraph.index + 1
    );
    reserve_materialized_text_bytes(
        materialized_text_bytes,
        paragraph.text.len(),
        1,
        materialized_text_limit,
        CONTENT_PART,
    )?;
    let rich_text = Visual::RichText {
        geometry: Geometry::Rectangle,
        fill: Paint::None,
        stroke: Paint::None,
        stroke_width: 0.0,
        align: paragraph.style.align,
        line_height,
        runs: paragraph.runs,
    };
    let visual = Visual::TextLayout {
        layout: TextLayout {
            inset_left: 0.0,
            inset_right: 0.0,
            inset_top: 0.0,
            inset_bottom: 0.0,
            tab_stops: paragraph.style.tab_stops.clone(),
            default_tab_stop: paragraph.style.default_tab_stop,
            prefix: paragraph.list_label.as_ref().map(|label| {
                format!(
                    "{label}{}",
                    paragraph
                        .list_position
                        .as_ref()
                        .map_or("\t", |p| p.followed_by)
                )
            }),
            hanging_indent: if paragraph.list_label.is_some() {
                24.0
            } else {
                0.0
            },
            paragraphs: vec![crate::model::TextParagraphLayout {
                align: paragraph.style.align,
                // Non-list margins are already applied to the object bounds.
                margin_left: paragraph
                    .list_position
                    .as_ref()
                    .map_or(0.0, |p| p.margin_left),
                margin_right: 0.0,
                first_line_indent: paragraph.style.text_indent,
                default_tab_stop: paragraph.style.default_tab_stop.max(1.0),
                line_height,
                space_before: 0.0,
                space_after: 0.0,
                latin_line_break: true,
                hanging_punctuation: false,
                rule_above: None,
                rule_below: None,
                drop_cap: paragraph.style.drop_cap,
            }],
            ..TextLayout::default()
        },
        visual: Box::new(rich_text),
    };
    objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: None,
        kind: ObjectKind::Paragraph,
        unit_index: *unit_index,
        bounds: Rect {
            x,
            y: *y,
            width,
            height,
        },
        z: i32::try_from(paragraph.index).unwrap_or(i32::MAX),
        text: Some(paragraph.text),
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping: if paragraph.element_id.is_some() {
                MappingQuality::Exact
            } else {
                MappingQuality::Derived
            },
            locator: SourceLocator::Odt {
                kind: "text-range",
                element_id: paragraph.element_id,
                path,
                row: None,
                column: None,
                text_range: Some((paragraph.source_text_start, text_end)),
            },
        },
        visual,
    });
    *y += height + paragraph.style.margin_bottom;
    if paragraph.style.page_break_after {
        *unit_index = unit_index
            .checked_add(1)
            .ok_or_else(|| format_error(CONTENT_PART, "page count exceeds the supported range"))?;
        units.push(page_unit(*unit_index, layout));
        *y = push_master_stories(*unit_index, layout, objects, styles, object_limit)?;
    }
    Ok(())
}

fn page_unit(index: u32, layout: PageLayout) -> Unit {
    Unit {
        kind: UnitKind::Page,
        index,
        id: format!("unit:{index}"),
        name: format!("Page {}", index + 1),
        width: layout.width,
        height: layout.height,
        rows: 0,
        columns: 0,
        frozen_rows: 0,
        frozen_columns: 0,
        frozen_width: 0.0,
        frozen_height: 0.0,
        row_axis: crate::model::SheetAxis::default(),
        column_axis: crate::model::SheetAxis::default(),
        show_grid_lines: false,
        tab_color: None,
        sheet: None,
        slide: None,
    }
}

fn effective_odt_line_height(
    style: &OdtParagraphStyle,
    runs: &[TextRun],
    metrics: Option<&FontMetricTable>,
) -> f32 {
    let natural = |family: &str, italic, bold, size| {
        metrics
            .and_then(|m| m.line_height_em(family, italic, bold, 1.0))
            .unwrap_or(1.2)
            * size
    };
    let base = if style.natural_line_height {
        natural(
            &style.text.font_family,
            style.text.italic,
            style.text.bold,
            style.text.font_size,
        )
    } else {
        style.line_height
    };
    runs.iter()
        .map(|run| natural(&run.font_family, run.italic, run.bold, run.font_size))
        .fold(base, f32::max)
        .max(1.0)
}

fn styled_paragraph_height(
    paragraph: &ParagraphState,
    width: f32,
    line_height: f32,
    metrics: Option<&FontMetricTable>,
) -> f32 {
    let advance = |run: &TextRun, character| {
        (metrics
            .and_then(|m| {
                m.advance_em_at_size(
                    &run.font_family,
                    run.italic,
                    run.bold,
                    character,
                    run.font_size,
                )
            })
            .unwrap_or(0.5)
            * run.font_size
            + run.letter_spacing)
            .max(0.0)
    };
    if paragraph.style.tab_stops.is_empty() || !paragraph.text.contains('\t') {
        let characters: Vec<_> = paragraph
            .runs
            .iter()
            .flat_map(|run| run.text.chars().map(|c| (c, advance(run, c))))
            .collect();
        let mut lines = crate::text_layout::horizontal_line_breaks(
            &characters,
            width - paragraph.style.text_indent.max(0.0),
            width,
            paragraph.style.default_tab_stop,
            false,
        )
        .len();
        if paragraph.text.ends_with('\n') {
            lines += 1;
        }
        return lines as f32 * line_height;
    }
    // Authored tab stops are not equivalent to the shared fitter's uniform tabs.
    let mut lines = 1_usize;
    let mut line_width = paragraph.style.text_indent.max(0.0);
    for run in &paragraph.runs {
        for character in run.text.chars() {
            let character_width = advance(run, character);
            if character == '\n' {
                lines = lines.saturating_add(1);
                line_width = 0.0;
            } else {
                if line_width > 0.0
                    && line_width + character_width > width
                    && !crate::text_layout::is_line_start_prohibited(character)
                {
                    lines = lines.saturating_add(1);
                    line_width = 0.0;
                }
                if character == '\t' {
                    line_width = paragraph
                        .style
                        .tab_stops
                        .iter()
                        .find(|stop| stop.position > line_width)
                        .map_or_else(
                            || {
                                ((line_width / paragraph.style.default_tab_stop).floor() + 1.0)
                                    * paragraph.style.default_tab_stop
                            },
                            |stop| stop.position,
                        );
                } else {
                    line_width += character_width;
                }
            }
        }
    }
    lines.max(1) as f32 * line_height
}

fn append_text(target: &mut String, value: &str, limit: usize) -> Result<(), Diagnostic> {
    if target.len().saturating_add(value.len()) > limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::XmlSizeLimit,
            Phase::Parse,
            None,
            "ODT paragraph text exceeds the configured XML byte limit",
        )
        .in_part(CONTENT_PART));
    }
    target.push_str(value);
    Ok(())
}

fn append_frame_text(frame: &mut FrameState, value: &str, limit: usize) -> Result<(), Diagnostic> {
    let style = frame
        .spans
        .last()
        .map_or(&frame.text_style, |(_, style)| style);
    append_text_run(&mut frame.text, &mut frame.runs, style, value, limit)
}

fn append_styled_text(
    paragraph: &mut ParagraphState,
    value: &str,
    limit: usize,
) -> Result<(), Diagnostic> {
    let style = paragraph
        .spans
        .last()
        .map_or(&paragraph.style.text, |(_, style)| style);
    append_text_run(
        &mut paragraph.text,
        &mut paragraph.runs,
        style,
        value,
        limit,
    )
}

fn append_shape_text(shape: &mut ShapeState, value: &str, limit: usize) -> Result<(), Diagnostic> {
    let style = shape
        .spans
        .last()
        .map_or(&shape.text_style.text, |(_, style)| style);
    append_text_run(&mut shape.text, &mut shape.runs, style, value, limit)
}

fn append_text_run(
    text: &mut String,
    runs: &mut Vec<TextRun>,
    style: &OdtTextStyle,
    value: &str,
    limit: usize,
) -> Result<(), Diagnostic> {
    append_text(text, value, limit)?;
    if value.is_empty() {
        return Ok(());
    }
    if let Some(run) = runs.last_mut()
        && same_text_run_style(run, style)
    {
        run.text.push_str(value);
    } else {
        runs.push(TextRun {
            paint: None,
            east_asian_line_breaks: true,
            text: value.to_owned(),
            font_family: style.font_family.clone(),
            font_size: style.font_size,
            color: style.color,
            bold: style.bold,
            italic: style.italic,
            underline: style.underline,
            strikethrough: style.strikethrough,
            highlight: style.highlight,
            baseline_shift: style.baseline_shift,
            letter_spacing: style.letter_spacing,
            horizontal_scale: 1.0,
        });
    }
    Ok(())
}

fn append_styled_repeated(
    paragraph: &mut ParagraphState,
    value: char,
    count: usize,
    limit: usize,
) -> Result<(), Diagnostic> {
    if paragraph.text.len().saturating_add(count) > limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::XmlSizeLimit,
            Phase::Parse,
            None,
            "ODT repeated paragraph text exceeds the configured XML byte limit",
        )
        .in_part(CONTENT_PART));
    }
    let value = std::iter::repeat_n(value, count).collect::<String>();
    append_styled_text(paragraph, &value, limit)
}

fn append_shape_repeated(
    shape: &mut ShapeState,
    value: char,
    count: usize,
    limit: usize,
) -> Result<(), Diagnostic> {
    if shape.text.len().saturating_add(count) > limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::XmlSizeLimit,
            Phase::Parse,
            None,
            "ODT repeated shape text exceeds the configured XML byte limit",
        )
        .in_part(CONTENT_PART));
    }
    append_shape_text(
        shape,
        &std::iter::repeat_n(value, count).collect::<String>(),
        limit,
    )
}

fn same_text_run_style(run: &TextRun, style: &OdtTextStyle) -> bool {
    run.font_family == style.font_family
        && run.font_size == style.font_size
        && run.color == style.color
        && run.bold == style.bold
        && run.italic == style.italic
        && run.underline == style.underline
        && run.strikethrough == style.strikethrough
        && run.highlight == style.highlight
        && run.baseline_shift == style.baseline_shift
        && run.letter_spacing == style.letter_spacing
}

fn append_repeated(
    target: &mut String,
    value: char,
    count: usize,
    limit: usize,
) -> Result<(), Diagnostic> {
    if target.len().saturating_add(count) > limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::XmlSizeLimit,
            Phase::Parse,
            None,
            "ODT repeated text exceeds the configured XML byte limit",
        )
        .in_part(CONTENT_PART));
    }
    target.extend(std::iter::repeat_n(value, count));
    Ok(())
}

fn required_length(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<f32, Diagnostic> {
    optional_length(attributes, name, part)?
        .ok_or_else(|| format_error(part, format!("element is missing {name}")))
}

fn optional_length(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<f32>, Diagnostic> {
    let Some(value) = optional_attribute(attributes, name, part)? else {
        return Ok(None);
    };
    parse_length(&value)
        .map(Some)
        .ok_or_else(|| format_error(part, format!("attribute {name} is not an ODF length")))
}

fn optional_signed_length(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<f32>, Diagnostic> {
    let Some(value) = optional_attribute(attributes, name, part)? else {
        return Ok(None);
    };
    parse_signed_length(&value)
        .map(Some)
        .ok_or_else(|| format_error(part, format!("attribute {name} is not an ODF length")))
}

fn parse_length(value: &str) -> Option<f32> {
    let value = value.trim();
    [
        ("in", CSS_PIXELS_PER_INCH),
        ("cm", CSS_PIXELS_PER_INCH / 2.54),
        ("mm", CSS_PIXELS_PER_INCH / 25.4),
        ("pt", CSS_PIXELS_PER_INCH / 72.0),
        ("pc", CSS_PIXELS_PER_INCH / 6.0),
        ("px", 1.0),
    ]
    .iter()
    .find_map(|(suffix, scale)| {
        value
            .strip_suffix(suffix)
            .and_then(|number| number.parse::<f32>().ok())
            .map(|number| number * scale)
            .filter(|number| number.is_finite() && *number >= 0.0)
    })
}

/// ODF `style:rel-column-width`: `n*`, optional `+` minimum marker, or `n%`.
fn parse_rel_column_width(value: &str) -> Option<f32> {
    let body = value
        .trim()
        .strip_suffix('+')
        .unwrap_or(value.trim())
        .trim();
    let digits = body.strip_suffix('*').unwrap_or(body).trim();
    let weight = if body.ends_with('%') {
        digits
            .strip_suffix('%')
            .unwrap_or(digits)
            .trim()
            .parse::<f32>()
            .ok()?
            / 100.0
    } else {
        digits.parse::<f32>().ok()?
    };
    (weight.is_finite() && weight > 0.0).then_some(weight)
}

/// Placeholder for `text:page-number` / `text:page-count` while stories are laid out.
const ODT_PAGE_NUMBER_SLOT: &str = "\u{E000}";
const ODT_PAGE_COUNT_SLOT: &str = "\u{E001}";

#[derive(Clone, Copy, Debug)]
enum RubyCapture {
    Base,
    Text,
}

fn is_odf_dynamic_page_field(local: &str) -> Option<&'static str> {
    match local {
        "page-number" => Some(ODT_PAGE_NUMBER_SLOT),
        "page-count" => Some(ODT_PAGE_COUNT_SLOT),
        _ => None,
    }
}

/// Display text for empty ODF field elements that carry their value in attributes.
fn odf_empty_field_text(
    local: &str,
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<Option<String>, Diagnostic> {
    let string_value = || -> Result<Option<String>, Diagnostic> {
        optional_attribute(attributes, "string-value", part)
    };
    let display = match local {
        "date" | "creation-date" | "modification-date" | "print-date" => {
            optional_attribute(attributes, "date-value", part)?
                .map(|value| format_odf_date_value(&value))
        }
        "time" | "creation-time" | "modification-time" | "print-time" => {
            optional_attribute(attributes, "time-value", part)?
                .map(|value| format_odf_time_value(&value))
        }
        "author"
        | "author-initials"
        | "author-name"
        | "sender-firstname"
        | "sender-lastname"
        | "sender-initials"
        | "sender-title"
        | "sender-position"
        | "sender-email"
        | "sender-phone-private"
        | "sender-street"
        | "sender-city"
        | "sender-postal-code"
        | "sender-country"
        | "sender-state"
        | "title"
        | "subject"
        | "file-name"
        | "template-name"
        | "sheet-name"
        | "page-style-name"
        | "chapter"
        | "page-continuation-string" => string_value()?,
        _ => None,
    };
    Ok(display.filter(|text| !text.is_empty()))
}

fn format_odf_date_value(value: &str) -> String {
    value
        .split_once('T')
        .map_or(value.trim(), |(date, _)| date)
        .to_owned()
}

fn format_odf_time_value(value: &str) -> String {
    let value = value.trim();
    value
        .split_once('.')
        .map_or(value, |(time, _)| time)
        .to_owned()
}

fn resolve_odf_page_fields(objects: &mut [Object], page_count: u32, starts: &[(u32, u32)]) {
    fn replace(text: &mut String, number: &str, count: &str) {
        if text.contains(ODT_PAGE_NUMBER_SLOT) || text.contains(ODT_PAGE_COUNT_SLOT) {
            *text = text
                .replace(ODT_PAGE_NUMBER_SLOT, number)
                .replace(ODT_PAGE_COUNT_SLOT, count);
        }
    }
    fn resolve_visual(visual: &mut Visual, number: &str, count: &str) {
        match visual {
            Visual::RichText { runs, .. } => {
                for run in runs {
                    replace(&mut run.text, number, count);
                }
            }
            Visual::TextLayout { visual, .. }
            | Visual::Layer { visual, .. }
            | Visual::Effect { visual, .. }
            | Visual::TextEffects { visual, .. }
            | Visual::StrokeStyle { visual, .. } => resolve_visual(visual, number, count),
            _ => {}
        }
    }
    let count = page_count.max(1).to_string();
    for object in objects {
        let number = starts
            .iter()
            .rev()
            .find(|(page, _)| *page <= object.unit_index)
            .map_or(object.unit_index + 1, |(page, start)| {
                start.saturating_add(object.unit_index - page)
            })
            .to_string();
        if let Some(text) = object.text.as_mut() {
            replace(text, &number, &count);
        }
        resolve_visual(&mut object.visual, &number, &count);
    }
}

fn parse_signed_length(value: &str) -> Option<f32> {
    let value = value.trim();
    [
        ("in", CSS_PIXELS_PER_INCH),
        ("cm", CSS_PIXELS_PER_INCH / 2.54),
        ("mm", CSS_PIXELS_PER_INCH / 25.4),
        ("pt", CSS_PIXELS_PER_INCH / 72.0),
        ("pc", CSS_PIXELS_PER_INCH / 6.0),
        ("px", 1.0),
    ]
    .iter()
    .find_map(|(suffix, scale)| {
        value
            .strip_suffix(suffix)
            .and_then(|number| number.parse::<f32>().ok())
            .map(|number| number * scale)
            .filter(|number| number.is_finite())
    })
}

fn apply_odt_cell_border(value: &str, style: &mut OdtCellStyle) {
    if value.eq_ignore_ascii_case("none") {
        style.stroke_width = 0.0;
        style.stroke_style = StrokeStyle::default();
        return;
    }
    let mut tokens = value.split_ascii_whitespace();
    let Some(width) = tokens.next().and_then(|width| match width {
        "thin" => Some(1.0),
        "medium" => Some(3.0),
        "thick" => Some(5.0),
        _ => parse_length(width),
    }) else {
        return;
    };
    let pattern = tokens.next().unwrap_or("solid");
    style.stroke_width = width.max(0.5);
    style.stroke_style = super::odf_border_stroke_style(pattern, style.stroke_width);
    if let Some(color) = value.split_ascii_whitespace().find_map(parse_odf_color) {
        style.stroke = color;
    }
}

fn parse_odf_color(value: &str) -> Option<u32> {
    let value = value.strip_prefix('#')?;
    if value.len() != 6 {
        return None;
    }
    u32::from_str_radix(value, 16)
        .ok()
        .map(|rgb| (rgb << 8) | 0xff)
}

fn optional_positive_u32(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<u32>, Diagnostic> {
    let Some(value) = optional_attribute(attributes, name, part)? else {
        return Ok(None);
    };
    value
        .parse::<u32>()
        .ok()
        .filter(|value| *value != 0)
        .map(Some)
        .ok_or_else(|| format_error(part, format!("attribute {name} is not a positive integer")))
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
        CSS_PIXELS_PER_INCH, LINE_HEIGHT, PageLayout, STYLES_PART, StyleCatalog, parse_content_xml,
        parse_flat, parse_length, parse_page_layout_xml, parse_style_catalog_xml, push_odt_chart,
    };
    use crate::diagnostic::DiagnosticCode;
    use crate::format::odp::{OdpBasicChart, OdpChartKind, OdpChartSeries};
    use crate::format::presentation_image::stored_zip;
    use crate::limits::Limits;
    use crate::model::{
        Geometry, MappingQuality, ObjectKind, Rect, SourceLocator, TextAlign, Visual,
    };
    use crate::package::Package;

    fn paragraph_visual(object: &crate::model::Object) -> &Visual {
        let Visual::TextLayout { layout, visual } = &object.visual else {
            panic!("flow paragraphs need an explicit layout");
        };
        assert_eq!(
            (
                layout.inset_left,
                layout.inset_right,
                layout.inset_top,
                layout.inset_bottom
            ),
            (0.0, 0.0, 0.0, 0.0)
        );
        visual
    }

    #[test]
    fn contextual_spacing_uses_common_styles_and_both_flags_with_area_boundaries() {
        let xml = br#"<office:document xmlns:office="office" xmlns:style="style" xmlns:text="text" xmlns:fo="fo">
          <office:styles>
            <style:style style:name="A" style:family="paragraph"><style:paragraph-properties fo:margin-top="10pt" fo:margin-bottom="20pt" style:contextual-spacing="true"/></style:style>
            <style:style style:name="B" style:family="paragraph" style:parent-style-name="A"/>
          </office:styles>
          <office:automatic-styles>
            <style:style style:name="P1" style:family="paragraph" style:parent-style-name="A"/>
            <style:style style:name="P2" style:family="paragraph" style:parent-style-name="A"/>
            <style:style style:name="Off" style:family="paragraph" style:parent-style-name="A"><style:paragraph-properties style:contextual-spacing="false"/></style:style>
          </office:automatic-styles>
          <office:body><office:text>
            <text:p text:style-name="P1">one</text:p><text:p text:style-name="P2">two</text:p>
            <text:section><text:p text:style-name="A">three</text:p><text:p text:style-name="A"/></text:section>
            <text:p text:style-name="P1">five</text:p><text:p text:style-name="Off">six</text:p>
            <text:p text:style-name="P2">seven</text:p><text:p text:style-name="B">eight</text:p>
            <text:p text:style-name="B" text:cond-style-name="A">nine</text:p><text:p text:style-name="P1">ten</text:p>
          </office:text></office:body>
        </office:document>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        let spacing =
            super::odt_paragraph_spacing(xml, Limits::default(), &styles, &Default::default())
                .unwrap();
        let before = 10.0 * (CSS_PIXELS_PER_INCH / 72.0);
        let after = 20.0 * (CSS_PIXELS_PER_INCH / 72.0);
        assert_eq!(
            spacing,
            vec![
                (before, 0.0),
                (0.0, after),
                (0.0, 0.0),
                (0.0, after),
                (0.0, after),
                (0.0, after),
                (0.0, after),
                (0.0, after),
                (0.0, 0.0),
                (0.0, after)
            ]
        );
        let separated = br#"<office:text xmlns:office="office" xmlns:text="text" xmlns:draw="draw">
          <text:p text:style-name="P1">one</text:p>
          <office:annotation><text:p text:style-name="B">comment</text:p></office:annotation>
          <draw:frame><draw:text-box><text:p text:style-name="B">frame</text:p></draw:text-box></draw:frame>
          <text:change-start text:change-id="deleted"/><text:p text:style-name="B">deleted</text:p><text:change-end text:change-id="deleted"/>
          <text:p text:style-name="P2">two</text:p>
        </office:text>"#;
        let spacing = super::odt_paragraph_spacing(
            separated,
            Limits::default(),
            &mut styles,
            &std::collections::HashSet::from(["deleted".to_owned()]),
        )
        .unwrap();
        assert_eq!(spacing[0], (before, 0.0));
        assert_eq!(spacing[4], (0.0, after));
        styles
            .styles
            .get_mut("P2")
            .unwrap()
            .paragraph
            .page_break_before = true;
        let spacing = super::odt_paragraph_spacing(
            separated,
            Limits::default(),
            &styles,
            &std::collections::HashSet::from(["deleted".to_owned()]),
        )
        .unwrap();
        assert_eq!(spacing[0], (before, after));
        assert_eq!(spacing[4], (before, after));
    }

    #[test]
    fn supplied_page_number_zero_restarts_without_changing_page_count() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/oasis-3923-page-number-zero.odt"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let texts: Vec<_> = document
            .objects
            .iter()
            .filter_map(|o| o.text.as_deref().map(|text| (o.unit_index, text)))
            .collect();
        assert_eq!(document.units.len(), 4, "{texts:?}");
        for (page, number) in [(0, "0"), (1, "1"), (2, "0")] {
            assert!(texts.contains(&(page, number)), "{texts:?}");
        }
        for (page, header) in [
            (0, "Seite 0 von 4"),
            (1, "Seite 1 von 4"),
            (2, "Seite 0 von 4"),
            (3, "Seite 1 von 4"),
        ] {
            assert!(texts.contains(&(page, header)), "{texts:?}");
        }
        assert!(document.units[0].height > document.units[0].width);
        assert!(document.units[2].width > document.units[2].height);
        for object in document
            .objects
            .iter()
            .filter(|o| o.kind == ObjectKind::Paragraph)
        {
            let visual = match &object.visual {
                Visual::TextLayout { visual, .. } => visual.as_ref(),
                visual => visual,
            };
            let Visual::RichText { runs, .. } = visual else {
                panic!("paragraph runs");
            };
            let painted: String = runs.iter().map(|run| run.text.as_str()).collect();
            assert_eq!(Some(painted.as_str()), object.text.as_deref());
            assert!(!painted.contains(['\u{E000}', '\u{E001}']));
        }
        // The authored restart, not the cached field text, controls numbering.
        for (value, expected) in [("7", [7, 8, 7, 8]), ("auto", [1, 2, 3, 4])] {
            let content = package.required_part("content.xml").unwrap();
            let content = std::str::from_utf8(&content).unwrap().replace(
                "style:page-number=\"0\"",
                &format!("style:page-number=\"{value}\""),
            );
            let styles = package.required_part("styles.xml").unwrap();
            let bytes = stored_zip(&[
                ("mimetype", b"application/vnd.oasis.opendocument.text"),
                ("content.xml", content.as_bytes()),
                ("styles.xml", &styles),
            ]);
            let variant = Package::open(&bytes, Limits::default()).unwrap();
            let parsed = super::parse(&variant).unwrap();
            assert_eq!(parsed.units.len(), 4);
            for (page, number) in expected.into_iter().enumerate() {
                let header = format!("Seite {number} von 4");
                assert!(
                    parsed
                        .objects
                        .iter()
                        .any(|o| o.unit_index == page as u32 && o.text.as_deref() == Some(&header))
                );
            }
        }
    }

    #[test]
    fn supplied_contextual_spacing_section_stays_on_one_page() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/oasis-contextual-spacing-section.odt"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        assert_eq!(document.units.len(), 1);
        let paragraphs: Vec<_> = document
            .objects
            .iter()
            .filter(|o| o.kind == ObjectKind::Paragraph)
            .collect();
        assert_eq!(paragraphs.len(), 6);
        assert!((paragraphs[0].bounds.y - 20.0 * CSS_PIXELS_PER_INCH / 25.4).abs() < 0.01);
        for p in &paragraphs {
            paragraph_visual(p);
            assert!((p.bounds.width - 170.01 * CSS_PIXELS_PER_INCH / 25.4).abs() < 0.1);
        }
        for pair in [
            (&paragraphs[0], &paragraphs[1]),
            (&paragraphs[2], &paragraphs[3]),
            (&paragraphs[4], &paragraphs[5]),
        ] {
            assert!((pair.1.bounds.y - pair.0.bounds.y - pair.0.bounds.height).abs() < 0.01);
        }
        for pair in [
            (&paragraphs[1], &paragraphs[2]),
            (&paragraphs[3], &paragraphs[4]),
        ] {
            assert!(
                (pair.1.bounds.y
                    - pair.0.bounds.y
                    - pair.0.bounds.height
                    - 10.0 * CSS_PIXELS_PER_INCH / 25.4)
                    .abs()
                    < 0.01
            );
        }
    }

    #[test]
    fn renders_real_embedded_math_without_replacement_image() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/math_OOo311.odt"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let formula = document
            .objects
            .iter()
            .find(|o| o.text.as_deref() == Some("1+1"))
            .expect("embedded MathML must render 1+1 exactly once, without annotation text");
        assert_eq!(formula.source.part, "Object 1/content.xml");
        assert!(
            formula.bounds.y < 100.0,
            "inline formula must stay in the first text line: {:?}",
            formula.bounds
        );
        assert!(
            matches!(&formula.visual, Visual::RichText { runs, .. } if runs.len() == 1 && runs[0].text == "1+1")
        );
        assert!(
            !document
                .objects
                .iter()
                .any(|o| matches!(o.visual, Visual::Image { .. }))
        );
        assert!(
            !document
                .diagnostics
                .iter()
                .any(|d| d.message.contains("static fallback")
                    || d.message.contains("image alternative"))
        );
    }

    #[test]
    fn real_annotation_stays_out_of_body_text() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/annotation-body.odt"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let body: Vec<_> = document
            .objects
            .iter()
            .filter_map(|object| object.text.as_deref())
            .collect();
        assert_eq!(body, ["Hallo"]);
    }

    #[test]
    fn corpus_footer_nested_tables_are_visible() {
        let bytes = include_bytes!("../../tests/fixtures/corpus-nestedTableInFooter.odt");
        let package = Package::open(bytes, Limits::default()).unwrap();
        let document = super::parse(&package).unwrap();
        let tables: Vec<_> = document
            .objects
            .iter()
            .filter(|o| o.kind == ObjectKind::Table)
            .collect();
        assert_eq!(tables.len(), 2);
        assert!(tables.iter().all(|o| o.bounds.y > 900.0));
    }

    #[test]
    fn corpus_auto_inline_heading_flows_with_paragraph() {
        let document = super::parse_flat(
            include_bytes!("../../tests/fixtures/corpus-tdf48459.fodt"),
            Limits::default(),
        )
        .unwrap();
        assert!(
            document
                .objects
                .iter()
                .any(|o| o.text.as_deref() == Some("Heading and paragraph")),
            "inline heading must participate in its parent paragraph: {:?}",
            document.objects.iter().map(|o| &o.text).collect::<Vec<_>>()
        );
        assert!(
            !document
                .objects
                .iter()
                .any(|o| o.kind == ObjectKind::TextBox)
        );
    }

    #[test]
    fn corpus_nested_footnotes_stay_out_of_body() {
        let document = super::parse_flat(
            include_bytes!("../../tests/fixtures/corpus-nested_footnote.fodt"),
            Limits::default(),
        )
        .unwrap();
        let body = document
            .objects
            .iter()
            .find(|o| o.text.as_deref().is_some_and(|t| t.starts_with("Lorem")))
            .unwrap();
        assert!(!body.text.as_ref().unwrap().contains("Vestibulum"));
        for prefix in ["Vestibulum", "Proin"] {
            let note = document
                .objects
                .iter()
                .find(|o| o.text.as_deref().is_some_and(|t| t.contains(prefix)))
                .unwrap();
            assert!(note.bounds.y > 900.0, "{} at {}", prefix, note.bounds.y);
        }
    }

    #[test]
    fn supplied_odt_preserves_section_columns_and_tab_leaders() {
        let bytes = include_bytes!("../../tests/fixtures/word-columns-tab-leaders.odt");
        let package = Package::open(bytes, Limits::default()).unwrap();
        let document = super::parse(&package).unwrap();
        let find = |prefix: &str| {
            document
                .objects
                .iter()
                .find(|object| {
                    object
                        .text
                        .as_deref()
                        .is_some_and(|text| text.starts_with(prefix))
                })
                .unwrap()
        };
        let bullet = find("Bullet (Alt");
        let copyright = find("Copyright (Alt");
        let greek = find("Greek delta");
        let umlaut = find("Umlauted A");
        assert!(
            (bullet.bounds.y - greek.bounds.y).abs() < 0.01,
            "Greek must start the second column: {:?} vs {:?}",
            bullet.bounds,
            greek.bounds
        );
        assert!((greek.bounds.x - bullet.bounds.x - bullet.bounds.width - 48.0).abs() < 0.01);
        assert_eq!(bullet.bounds.x, copyright.bounds.x);
        assert_eq!(bullet.bounds.x, umlaut.bounds.x);
        assert!((umlaut.bounds.y - copyright.bounds.y - copyright.bounds.height).abs() < 0.01);
        assert!(umlaut.bounds.width > bullet.bounds.width * 2.0);
        use crate::model::TextTabLeader;
        for (object, leader) in [
            (bullet, TextTabLeader::Dot),
            (copyright, TextTabLeader::Hyphen),
            (greek, TextTabLeader::Dot),
            (umlaut, TextTabLeader::Underscore),
        ] {
            let Visual::TextLayout { layout, .. } = &object.visual else {
                panic!("missing authored tabs")
            };
            assert_eq!(layout.tab_stops.len(), 1);
            assert_eq!(layout.tab_stops[0].position, 192.0);
            assert_eq!(layout.tab_stops[0].leader, leader);
        }
    }

    #[test]
    fn supplied_odt_renders_all_seven_nested_graphics() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/word-columns-tab-leaders.odt"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let images: Vec<_> = document
            .objects
            .iter()
            .filter(|o| o.kind == ObjectKind::Image)
            .collect();
        assert_eq!(
            images.len(),
            7,
            "all original JPEG, PNG and EMF pictures must survive nested text boxes"
        );
        for (index, ext) in ["jpeg", "png", "png", "png", "png", "emf", "emf"]
            .iter()
            .enumerate()
        {
            let bytes = package
                .part(&format!("media/image{}.{ext}", index + 1))
                .unwrap()
                .unwrap();
            let image = images.iter().find(|o| matches!(&o.visual, Visual::Image { bytes: actual, .. } if actual.as_slice() == &*bytes)).unwrap_or_else(|| panic!("missing original image {}", index+1));
            let parent = document
                .objects
                .iter()
                .find(|o| Some(o.numeric_id) == image.parent_numeric_id)
                .expect("retain containing frame");
            assert_eq!(
                parent.bounds, image.bounds,
                "auto-sized wrapper must fit its picture"
            );
            assert_eq!(parent.unit_index, image.unit_index);
            let page = &document.units[image.unit_index as usize];
            assert!(image.bounds.x >= 0.0 && image.bounds.y >= 0.0);
            assert!(image.bounds.x + image.bounds.width <= page.width + 0.01);
            assert!(image.bounds.y + image.bounds.height <= page.height + 0.01);
        }
        let picture = |name: &str| {
            images.iter().find(|o| matches!(&o.source.locator, SourceLocator::Odt { element_id: Some(id), .. } if id == name)).unwrap()
        };
        assert!((picture("Picture 7").bounds.x - 4.5006 * 96.0).abs() < 0.01);
        assert!((picture("Picture 7").bounds.y - 1.7506 * 96.0).abs() < 0.01);
        assert!(
            !document
                .diagnostics
                .iter()
                .any(|d| d.message.contains("nested ODT frames are omitted"))
        );
    }

    #[test]
    fn section_column_fallback_preserves_order_and_tab_style_boundaries() {
        let styles = br#"<office:styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
          <style:style style:name="S" style:family="section"><style:section-properties>
            <style:columns fo:column-count="2" fo:column-gap="0.5in"/>
          </style:section-properties></style:style>
          <style:style style:name="Tabs" style:family="paragraph"><style:paragraph-properties style:tab-stop-distance="0.5in">
            <style:tab-stops><style:tab-stop style:position="2in" style:leader-style="dotted"/></style:tab-stops>
          </style:paragraph-properties></style:style>
          <style:style style:name="Clear" style:family="paragraph" style:parent-style-name="Tabs"><style:paragraph-properties><style:tab-stops/></style:paragraph-properties></style:style>
          <style:style style:name="None" style:family="paragraph" style:parent-style-name="Tabs"><style:paragraph-properties><style:tab-stops><style:tab-stop style:position="1in" style:type="right" style:leader-style="none" style:leader-text="."/></style:tab-stops></style:paragraph-properties></style:style>
        </office:styles>"#;
        let mut catalog = StyleCatalog::default();
        parse_style_catalog_xml(styles, Limits::default(), STYLES_PART, &mut catalog).unwrap();
        assert_eq!(catalog.paragraph(Some("Tabs")).tab_stops[0].position, 192.0);
        assert!(catalog.paragraph(Some("Clear")).tab_stops.is_empty());
        let none = catalog.paragraph(Some("None"));
        assert_eq!(none.default_tab_stop, 48.0);
        assert_eq!(none.tab_stops[0].align, TextAlign::End);
        assert_eq!(none.tab_stops[0].leader, crate::model::TextTabLeader::None);
        for middle in [
            "<text:section><text:p>Middle</text:p></text:section>",
            "<text:p>Middle<text:soft-page-break/>continued</text:p>",
            "<table:table><table:table-row><table:table-cell><text:p>Middle</text:p></table:table-cell></table:table-row></table:table>",
        ] {
            let xml = format!(
                r#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:table="table"><office:body><office:text>
              <text:section text:style-name="S"><text:p>First</text:p>{middle}<text:p>Last</text:p></text:section><text:p>After</text:p>
            </office:text></office:body></office:document-content>"#
            );
            let parsed = parse_content_xml(
                xml.as_bytes(),
                PageLayout::default(),
                Limits::default(),
                &mut catalog,
                None,
            )
            .unwrap();
            let texts: Vec<_> = parsed
                .objects
                .iter()
                .filter_map(|o| o.text.as_deref())
                .collect();
            let index = |text| texts.iter().position(|value| *value == text).unwrap();
            assert!(
                index("First") < index("Middle")
                    && index("Middle") < index("Last")
                    && index("Last") < index("After")
            );
            assert!(
                parsed
                    .diagnostics
                    .iter()
                    .any(|d| d.message.contains("section columns"))
            );
        }
    }

    #[test]
    fn renders_native_odt_line_scatter_and_pie_charts() {
        for kind in [OdpChartKind::Line, OdpChartKind::Scatter, OdpChartKind::Pie] {
            let chart = OdpBasicChart {
                kind,
                title: None,
                border_visible: true,
                categories: vec!["1".to_owned(), "2".to_owned()],
                series: vec![OdpChartSeries {
                    label: Some("Series".to_owned()),
                    color: Some(0x1234_56ff),
                    values: vec![1.0, 2.0],
                    domains: vec![10.0, 20.0],
                    custom_labels: Vec::new(),
                    point_explosions: Vec::new(),
                    data_labels: None,
                }],
                show_legend: true,
                stacked: false,
                category_axis_at_end: false,
                value_axis_at_end: false,
                plot_area: None,
                value_axis_visible: true,
                wall: None,
                source_part: "Object 1/content.xml".to_owned(),
            };
            let mut objects = Vec::new();
            push_odt_chart(
                &chart,
                0,
                None,
                0,
                "parent",
                0,
                Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 400.0,
                    height: 300.0,
                },
                &mut objects,
                100,
            )
            .expect("supported ODT chart");
            assert!(objects.iter().any(|object| {
                object.source.mapping == MappingQuality::Exact
                    && matches!(
                        object.visual,
                        Visual::PaintedShape {
                            geometry: Geometry::Path { .. },
                            ..
                        }
                    )
            }));
        }
    }

    #[test]
    fn converts_odf_lengths() {
        assert_eq!(parse_length("8.5in"), Some(816.0));
        assert_eq!(parse_length("72pt"), Some(96.0));
        assert_eq!(super::parse_signed_length("-2cm"), Some(-2.0 * 96.0 / 2.54));
        assert_eq!(parse_length("50%"), None);
    }

    #[test]
    fn corpus_header_footer_first_and_left_variants_render() {
        let bytes = include_bytes!("../../tests/fixtures/corpus-header-footer-first.odt");
        let package = Package::open(bytes, Limits::default()).unwrap();
        let document = super::parse(&package).unwrap();
        assert_eq!(
            document.units.len(),
            7,
            "OASIS 3789 fixture lays out as 7 pages"
        );
        let stories: Vec<(u32, &str)> = document
            .objects
            .iter()
            .filter(|object| object.source.part == STYLES_PART)
            .filter_map(|object| {
                object
                    .text
                    .as_deref()
                    .filter(|text| text.contains("Page Header") || text.contains("Page Footer"))
                    .map(|text| (object.unit_index, text))
            })
            .collect();
        let by_unit = |unit: u32| -> Vec<&str> {
            stories
                .iter()
                .filter(|(index, _)| *index == unit)
                .map(|(_, text)| *text)
                .collect()
        };
        assert_eq!(
            by_unit(0),
            [
                "First Page Header (First Page)",
                "First Page Footer (First Page)"
            ]
        );
        assert_eq!(
            by_unit(1),
            ["First Page Header (Index)", "First Page Footer (Index)"]
        );
        assert_eq!(
            by_unit(2),
            [
                "First Page Header (Default Page Style)",
                "First Page Footer (Default Page Style)"
            ]
        );
        assert_eq!(
            by_unit(3),
            [
                "Left Page Header (Default Page Style)",
                "Left Page Footer (Default Page Style)"
            ]
        );
        assert_eq!(
            by_unit(4),
            [
                "Right Page Header (Default Page Style)",
                "Right Page Footer (Default Page Style)"
            ]
        );
        assert_eq!(
            by_unit(5),
            [
                "Left Page Header (Default Page Style)",
                "Left Page Footer (Default Page Style)"
            ]
        );
        assert_eq!(
            by_unit(6),
            [
                "Right Page Header (Default Page Style)",
                "Right Page Footer (Default Page Style)"
            ]
        );
    }

    #[test]
    fn corpus_first_line_indent_does_not_shift_paragraph_bounds() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/oasis-3937-background-border.odt"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let paragraph = |prefix: &str| {
            document
                .objects
                .iter()
                .find(|o| o.text.as_deref().is_some_and(|t| t.starts_with(prefix)))
                .unwrap()
        };
        let plain = paragraph("Er hörte");
        let indented = paragraph("Oder gehörten");
        assert!((indented.bounds.x - plain.bounds.x).abs() < 0.01);
        assert!((indented.bounds.width - plain.bounds.width).abs() < 0.01);
        let Visual::TextLayout { layout, .. } = &indented.visual else {
            panic!("paragraph layout")
        };
        assert!((layout.paragraphs[0].first_line_indent - 4.99 * 96.0 / 25.4).abs() < 0.01);
    }

    #[test]
    fn corpus_header_footer_layout_borders_render() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/corpus-header-footer-first.odt"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        for (unit, header_color, footer_color) in [
            (0, 0x32cd32ff, 0x800000ff),
            (1, 0x0000ffff, 0x9932ccff),
            (2, 0xff4500ff, 0xbdb76bff),
            (3, 0xff4500ff, 0xbdb76bff),
            (4, 0xff4500ff, 0xbdb76bff),
            (5, 0xff4500ff, 0xbdb76bff),
            (6, 0xff4500ff, 0xbdb76bff),
        ] {
            for (kind, color) in [("header", header_color), ("footer", footer_color)] {
                let border = document
                    .objects
                    .iter()
                    .find(|o| {
                        o.unit_index == unit && o.stable_id == format!("odt:{kind}-box:{unit}")
                    })
                    .expect("authored header/footer border must be rendered");
                assert!(matches!(border.visual, Visual::Shape { stroke: c, .. } if c == color));
                assert!((border.bounds.width - 108.01 * 96.0 / 25.4).abs() < 0.1);
                assert!((border.bounds.height - 30.01 * 96.0 / 25.4).abs() < 0.1);
                let text = document
                    .objects
                    .iter()
                    .find(|o| {
                        o.unit_index == unit
                            && o.text.as_deref().is_some_and(|t| {
                                t.contains(if kind == "header" {
                                    "Page Header"
                                } else {
                                    "Page Footer"
                                })
                            })
                    })
                    .unwrap();
                assert!(text.bounds.x > border.bounds.x);
                assert!(text.bounds.y > border.bounds.y);
                assert!(
                    text.bounds.y + text.bounds.height < border.bounds.y + border.bounds.height
                );
            }
        }
    }

    #[test]
    fn odf_master_story_variants_fall_back_unlike_word() {
        use super::{MasterStoryParagraph, MasterStoryVariants, OdtParagraphStyle};
        let mut variants = MasterStoryVariants::default();
        assert!(!variants.select(true, true, false).1.present);
        assert!(!variants.select(true, false, true).1.present);
        assert!(!variants.select(true, false, false).1.present);

        // Present blank first variant wins and does not inherit the default text.
        variants.default.paragraphs.push(MasterStoryParagraph {
            text: "default".into(),
            style: OdtParagraphStyle::default(),
            runs: Vec::new(),
            spans: Vec::new(),
        });
        variants.default.present = true;
        variants.first.present = true;
        let (kind, story) = variants.select(true, true, false);
        assert_eq!(kind, "header-first");
        assert!(story.paragraphs.is_empty());

        // Absent left variant falls back to default (Word would keep it missing).
        let (kind, story) = variants.select(true, false, true);
        assert_eq!(kind, "header");
        assert_eq!(story.paragraphs[0].text, "default");

        variants.left.present = true;
        let (kind, story) = variants.select(true, false, true);
        assert_eq!(kind, "header-left");
        assert!(story.paragraphs.is_empty());
    }

    #[test]
    fn selects_the_named_master_page_layout() {
        let styles = br#"<office:document-styles>
          <office:styles><style:default-page-layout>
            <style:page-layout-properties style:writing-mode="lr-tb"/>
          </style:default-page-layout></office:styles>
          <office:automatic-styles><style:page-layout style:name="Mpm1">
            <style:page-layout-properties fo:page-width="200mm" fo:page-height="100mm" fo:margin="10mm"/>
          </style:page-layout><style:page-layout style:name="Mpm2">
            <style:page-layout-properties fo:page-width="300mm" fo:page-height="150mm" fo:margin="10mm"/>
          </style:page-layout></office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Standard" style:page-layout-name="Mpm1"/>
            <style:master-page style:name="First" style:page-layout-name="Mpm2"/>
          </office:master-styles>
        </office:document-styles>"#;

        let layout = parse_page_layout_xml(styles, Limits::default(), Some("First"))
            .expect("the selected master page must choose its own named page layout");
        assert_eq!(layout.width, 300.0 * 96.0 / 25.4);
        assert_eq!(layout.height, 150.0 * 96.0 / 25.4);
    }

    #[test]
    fn preserves_odt_paragraph_and_span_styles_as_rich_text() {
        let styles_xml = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
          <office:styles>
            <style:style style:name="Body" style:family="paragraph">
              <style:paragraph-properties fo:text-align="center" fo:line-height="150%" fo:margin-left="12pt"/>
              <style:text-properties style:font-name="Liberation Sans" fo:font-size="12pt" fo:color="#123456" fo:font-weight="bold"/>
            </style:style>
            <style:style style:name="Emphasis" style:family="text">
              <style:text-properties style:font-name="Liberation Serif" fo:font-size="14pt" fo:color="#ff0000" style:text-underline-style="solid" fo:background-color="#ffff00"/>
            </style:style>
          </office:styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text"><office:body><office:text>
          <text:p text:style-name="Body">Body <text:span text:style-name="Emphasis">emphasis</text:span></text:p>
        </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), "styles.xml", &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();
        let paragraph = parsed
            .objects
            .iter()
            .find(|object| object.kind == ObjectKind::Paragraph)
            .unwrap();
        assert!(paragraph.bounds.x > PageLayout::default().margin_left);
        let Visual::RichText {
            align,
            line_height,
            runs,
            ..
        } = paragraph_visual(paragraph)
        else {
            panic!("ODT paragraph must use rich text");
        };
        assert_eq!(*align, TextAlign::Center);
        assert_eq!(*line_height, 24.0);
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].font_family, "Liberation Sans");
        assert!(runs[0].bold);
        assert_eq!(runs[1].font_family, "Liberation Serif");
        assert!(runs[1].underline);
        assert_eq!(runs[1].highlight, 0xffff_00ff);
    }

    #[test]
    fn applies_odt_subscript_relative_font_sizes() {
        let styles_xml = br#"<office:document-styles xmlns:office="office" xmlns:style="style">
          <office:styles>
            <style:style style:name="DefaultSubscript" style:family="text"><style:text-properties style:text-position="sub 66%"/></style:style>
            <style:style style:name="LargeSubscript" style:family="text" style:parent-style-name="DefaultSubscript"><style:text-properties style:text-position="sub 80%"/></style:style>
            <style:style style:name="CustomSubscript" style:family="text" style:parent-style-name="LargeSubscript"><style:text-properties style:text-position="-25% 80%"/></style:style>
          </office:styles>
        </office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:style="style" xmlns:text="text">
          <office:automatic-styles><style:style style:name="T1" style:family="text"><style:text-properties style:text-position="sub 58%"/></style:style></office:automatic-styles>
          <office:body><office:text><text:p>A<text:span text:style-name="T1">one</text:span><text:span text:style-name="DefaultSubscript">two</text:span><text:span text:style-name="LargeSubscript">three</text:span><text:span text:style-name="CustomSubscript">four</text:span></text:p></office:text></office:body>
        </office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), "styles.xml", &mut styles).unwrap();
        parse_style_catalog_xml(content, Limits::default(), "content.xml", &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();
        let Visual::RichText { runs, .. } = paragraph_visual(&parsed.objects[0]) else {
            panic!("ODT paragraph must use rich text");
        };
        let base_size = 16.0;
        assert_eq!(runs.len(), 5);
        assert!((runs[1].font_size - base_size * 0.58).abs() < 0.001);
        assert!((runs[2].font_size - base_size * 0.66).abs() < 0.001);
        assert!((runs[3].font_size - base_size * 0.80).abs() < 0.001);
        assert!((runs[4].font_size - base_size * 0.80).abs() < 0.001);
        assert!(runs[4].baseline_shift < 0.0);
    }

    #[test]
    fn lays_out_table_object_tree_and_source_mapping() {
        let xml = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:table="table" xmlns:xml="xml">
          <office:body><office:text>
            <text:p xml:id="before">Before</text:p>
            <table:table xml:id="prices">
              <table:table-row>
                <table:table-cell xml:id="a1"><text:p>A1</text:p></table:table-cell>
                <table:table-cell><text:p>B1</text:p></table:table-cell>
              </table:table-row>
              <table:table-row>
                <table:table-cell table:number-columns-repeated="2"><text:p>next</text:p></table:table-cell>
              </table:table-row>
            </table:table>
            <text:p>After</text:p>
          </office:text></office:body>
        </office:document-content>"#;

        let parsed = parse_content_xml(
            xml,
            PageLayout::default(),
            Limits::default(),
            &mut StyleCatalog::default(),
            None,
        )
        .unwrap();
        assert_eq!(parsed.units.len(), 1);
        assert_eq!(parsed.objects.len(), 7);
        assert_eq!(parsed.objects[0].kind, ObjectKind::Paragraph);
        assert_eq!(parsed.objects[1].kind, ObjectKind::Table);
        assert!(matches!(
            parsed.objects[1].source.locator,
            SourceLocator::Odt {
                kind: "element",
                row: None,
                column: None,
                text_range: None,
                ..
            }
        ));
        assert_eq!(parsed.objects[2].kind, ObjectKind::Cell);
        assert_eq!(parsed.objects[2].parent_numeric_id, Some(1));
        assert_eq!(
            parsed.objects[2].parent_stable_id.as_deref(),
            Some("odt:table:0:fragment:0")
        );
        assert_eq!(parsed.objects[2].text.as_deref(), Some("A1"));
        assert!(parsed.objects[2].bounds.width > 0.0);
        assert!(parsed.objects[2].bounds.height > 0.0);
        assert_eq!(
            parsed.objects[2].source.locator,
            SourceLocator::Odt {
                kind: "table-cell",
                element_id: Some("a1".to_owned()),
                path: "/office:document-content/office:body/office:text/table:table[1]/table:table-row[1]/table:table-cell[1]".to_owned(),
                row: Some(0),
                column: Some(0),
                text_range: Some((0, 2)),
            }
        );
        assert_eq!(parsed.objects[4].source.mapping, MappingQuality::Derived);
        assert_eq!(parsed.objects[5].kind, ObjectKind::Cell);
        assert_eq!(parsed.objects[6].kind, ObjectKind::Paragraph);
        assert_eq!(parsed.objects[6].text.as_deref(), Some("After"));
    }

    #[test]
    fn rejects_repeated_table_dimensions_before_materializing_objects() {
        let xml = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:table="table">
          <office:body><office:text><table:table><table:table-row table:number-rows-repeated="3">
            <table:table-cell table:number-columns-repeated="2"><text:p>x</text:p></table:table-cell>
          </table:table-row></table:table></office:text></office:body>
        </office:document-content>"#;
        let limits = Limits {
            max_document_objects: 5,
            ..Limits::default()
        };

        let error = parse_content_xml(
            xml,
            PageLayout::default(),
            limits,
            &mut StyleCatalog::default(),
            None,
        )
        .unwrap_err();
        assert_eq!(error.code, DiagnosticCode::ObjectLimit);
    }

    #[test]
    fn preserves_plain_paragraph_flow() {
        for closing in "、。，．！？：；）］】》〉」』〕”’,.;:!?)]}…".chars()
        {
            let xml = format!(
                r#"<office:document-content xmlns:office="office" xmlns:text="text"><office:body><office:text><text:p>中文{closing}</text:p></office:text></office:body></office:document-content>"#
            );
            let page = PageLayout {
                width: super::FONT_SIZE + 2.0 * super::DEFAULT_MARGIN,
                ..PageLayout::default()
            };
            let parsed = parse_content_xml(
                xml.as_bytes(),
                page,
                Limits::default(),
                &mut StyleCatalog::default(),
                None,
            )
            .unwrap();
            assert!(
                (parsed.objects[0].bounds.height - super::FONT_SIZE * 1.2).abs() < 0.01,
                "{closing}: {:?}",
                parsed.objects[0].bounds,
            );
        }
        let xml = br#"<office:document-content xmlns:office="office" xmlns:text="text">
          <office:body><office:text><text:p xml:id="intro">Hello<text:s text:c="2"/>world</text:p></office:text></office:body>
        </office:document-content>"#;

        let parsed = parse_content_xml(
            xml,
            PageLayout::default(),
            Limits::default(),
            &mut StyleCatalog::default(),
            None,
        )
        .unwrap();
        assert_eq!(parsed.objects.len(), 1);
        assert_eq!(parsed.objects[0].kind, ObjectKind::Paragraph);
        assert_eq!(parsed.objects[0].text.as_deref(), Some("Hello  world"));
        assert!(matches!(
            parsed.objects[0].source.locator,
            SourceLocator::Odt {
                kind: "text-range",
                row: None,
                column: None,
                text_range: Some((0, 12)),
                ..
            }
        ));
    }

    #[test]
    fn omits_content_inside_tracked_deletions() {
        let xml = br#"<office:document xmlns:office="office" xmlns:text="text" xmlns:table="table" office:mimetype="application/vnd.oasis.opendocument.text">
          <office:body><office:text>
            <text:tracked-changes><text:changed-region text:id="ct1"><text:deletion/></text:changed-region></text:tracked-changes>
            <text:p><text:change-start text:change-id="ct1"/></text:p>
            <table:table><table:table-row><table:table-cell><text:p>Deleted table</text:p></table:table-cell></table:table-row></table:table>
            <text:p><text:change-end text:change-id="ct1"/></text:p>
          </office:text></office:body>
        </office:document>"#;

        let document = parse_flat(xml, Limits::default()).unwrap();

        assert!(
            document
                .objects
                .iter()
                .all(|object| object.text.as_deref().is_none_or(str::is_empty))
        );
        assert!(
            document
                .objects
                .iter()
                .all(|object| !matches!(object.kind, ObjectKind::Table | ObjectKind::Cell))
        );
    }

    #[test]
    fn cached_leading_page_break_does_not_create_an_empty_paragraph() {
        let xml = br#"<office:document-content xmlns:office="office" xmlns:text="text"><office:body><office:text>
          <text:p>First</text:p><text:p><text:soft-page-break/>Second</text:p>
        </office:text></office:body></office:document-content>"#;
        let parsed = parse_content_xml(
            xml,
            PageLayout::default(),
            Limits::default(),
            &mut StyleCatalog::default(),
            None,
        )
        .unwrap();
        assert_eq!(parsed.units.len(), 2);
        assert_eq!(
            parsed
                .objects
                .iter()
                .filter_map(|o| o.text.as_deref())
                .collect::<Vec<_>>(),
            vec!["First", "Second"]
        );
    }

    #[test]
    fn trailing_paragraph_spacing_does_not_push_fitting_text_to_next_page() {
        let xml = br#"<office:document-content xmlns:office="office" xmlns:text="text"><office:body><office:text>
          <text:p>First</text:p><text:p>Second</text:p>
        </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        styles.default_paragraph.margin_bottom = 20.0;
        let page = PageLayout {
            height: super::DEFAULT_MARGIN * 2.0 + super::LINE_HEIGHT * 2.0 + 21.0,
            ..PageLayout::default()
        };
        let parsed = parse_content_xml(xml, page, Limits::default(), &mut styles, None).unwrap();
        assert_eq!(parsed.units.len(), 1);
        assert!(
            parsed
                .objects
                .iter()
                .all(|o| o.bounds.y + o.bounds.height <= page.height - page.margin_bottom)
        );
    }

    #[test]
    fn flat_odt_merges_text_style_overrides_and_preserves_soft_page_breaks() {
        let xml = br#"<office:document xmlns:office="office" xmlns:style="style" xmlns:text="text" xmlns:fo="fo" office:mimetype="application/vnd.oasis.opendocument.text">
          <office:styles>
            <style:default-style style:family="paragraph"><style:text-properties style:font-name="Times New Roman" fo:font-size="12pt"/></style:default-style>
            <style:style style:name="Heading" style:family="paragraph"><style:text-properties style:font-name="Liberation Sans" fo:font-size="18pt" fo:font-weight="bold"/></style:style>
          </office:styles>
          <office:automatic-styles>
            <style:style style:name="T1" style:family="text"><style:text-properties fo:font-size="49pt"/></style:style>
            <style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="20.814cm" fo:page-height="27.94cm" fo:margin="2cm"/></style:page-layout>
          </office:automatic-styles>
          <office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles>
          <office:body><office:text>
            <text:h text:style-name="Heading"><text:span text:style-name="T1">Heading</text:span></text:h>
            <text:p><text:span text:style-name="T1">First page</text:span><text:soft-page-break/><text:span text:style-name="T1">Second page</text:span><text:soft-page-break/><text:span text:style-name="T1">Third page</text:span></text:p>
          </office:text></office:body>
        </office:document>"#;

        let document = parse_flat(xml, Limits::default()).unwrap();
        assert_eq!(document.units.len(), 3);
        let text_objects = document
            .objects
            .iter()
            .filter(|object| object.text.is_some())
            .collect::<Vec<_>>();
        assert_eq!(
            text_objects
                .iter()
                .map(|object| (object.unit_index, object.text.as_deref().unwrap()))
                .collect::<Vec<_>>(),
            vec![
                (0, "Heading"),
                (0, "First page"),
                (1, "Second page"),
                (2, "Third page"),
            ]
        );
        let Visual::RichText {
            line_height,
            runs: heading_runs,
            ..
        } = paragraph_visual(text_objects[0])
        else {
            panic!("heading must remain rich text");
        };
        assert_eq!(heading_runs[0].font_family, "Liberation Sans");
        assert_eq!(heading_runs[0].font_size, 49.0 * CSS_PIXELS_PER_INCH / 72.0);
        assert!(*line_height >= heading_runs[0].font_size * 1.15);
        let Visual::RichText {
            runs: body_runs, ..
        } = paragraph_visual(text_objects[1])
        else {
            panic!("body must remain rich text");
        };
        assert_eq!(body_runs[0].font_family, "Times New Roman");
    }

    #[test]
    fn reports_table_pagination_merge_and_image_fallbacks() {
        let xml = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:table="table" xmlns:draw="draw">
          <office:body><office:text><table:table>
            <table:table-row><table:table-cell table:number-columns-spanned="2"><text:p>one</text:p><draw:image/></table:table-cell></table:table-row>
            <table:table-row><table:table-cell><text:p>two</text:p></table:table-cell></table:table-row>
            <table:table-row><table:table-cell><text:p>three</text:p></table:table-cell></table:table-row>
            <table:table-row><table:table-cell><text:p>four</text:p></table:table-cell></table:table-row>
          </table:table></office:text></office:body>
        </office:document-content>"#;
        let layout = PageLayout {
            width: 200.0,
            height: 100.0,
            margin_top: 10.0,
            margin_right: 10.0,
            margin_bottom: 10.0,
            margin_left: 10.0,
        };

        let parsed = parse_content_xml(
            xml,
            layout,
            Limits::default(),
            &mut StyleCatalog::default(),
            None,
        )
        .unwrap();
        assert_eq!(parsed.units.len(), 2);
        let messages = parsed
            .diagnostics
            .iter()
            .map(|diagnostic| diagnostic.message.as_str())
            .collect::<Vec<_>>();
        assert!(messages.iter().any(|message| message.contains("split")));
        assert!(
            !messages
                .iter()
                .any(|message| message.contains("unmerged fallback"))
        );
        // Images with a resolvable href are rendered; href-less table images stay diagnosed.
        assert!(
            messages
                .iter()
                .any(|message| message.contains("table image has no href")),
            "{messages:?}"
        );
    }

    #[test]
    fn renders_body_frames_and_diagnoses_unsupported_paths_once() {
        let xml = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:draw="draw">
          <office:body><office:text>
            <text:p>before<draw:frame><draw:image/></draw:frame><draw:frame><draw:image/></draw:frame></text:p>
            <draw:custom-shape/><draw:path/><draw:custom-shape/>
          </office:text></office:body>
        </office:document-content>"#;

        let parsed = parse_content_xml(
            xml,
            PageLayout::default(),
            Limits::default(),
            &mut StyleCatalog::default(),
            None,
        )
        .unwrap();
        let unsupported = parsed
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.code == DiagnosticCode::UnsupportedFeature)
            .collect::<Vec<_>>();
        assert_eq!(
            parsed
                .objects
                .iter()
                .filter(|object| object.kind == ObjectKind::Shape)
                .count(),
            4,
            "frames and supported custom shapes remain visible and addressable"
        );
        assert_eq!(
            unsupported
                .iter()
                .filter(|diagnostic| diagnostic.message.contains("freeform path"))
                .count(),
            1,
            "unsupported freeform shapes remain deduplicated"
        );
    }

    #[test]
    fn renders_character_anchored_group_connectors_and_labels_in_flow() {
        let styles_xml = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:fo="fo">
          <office:styles>
            <style:style style:name="Group" style:family="graphic"><style:graphic-properties fo:margin-top="20px"/></style:style>
            <style:style style:name="Connector" style:family="graphic"><style:graphic-properties draw:fill="none" draw:stroke="solid" svg:stroke-color="#616161" svg:stroke-width="2px" draw:marker-end="Arrow"/></style:style>
            <style:style style:name="Label" style:family="paragraph"/>
            <style:style style:name="Red" style:family="text"><style:text-properties fo:color="#c11e25"/></style:style>
          </office:styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:draw="draw" xmlns:svg="svg"><office:body><office:text>
          <text:p>Line 1<text:line-break/>Line 2<draw:g text:anchor-type="char" draw:style-name="Group">
            <draw:rect draw:id="a" svg:x="0px" svg:y="-20px" svg:width="20px" svg:height="20px"/>
            <draw:rect draw:id="b" svg:x="60px" svg:y="40px" svg:width="20px" svg:height="20px"/>
            <draw:connector draw:style-name="Connector" draw:text-style-name="Label" svg:x1="20px" svg:y1="-10px" svg:x2="60px" svg:y2="50px" draw:start-shape="a" draw:start-glue-point="7" draw:end-shape="b" draw:end-glue-point="5"><text:p><text:span text:style-name="Red">No</text:span></text:p></draw:connector>
          </draw:g></text:p>
        </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();

        let connector = parsed
            .objects
            .iter()
            .find(|object| {
                matches!(
                    &object.source.locator,
                    SourceLocator::Odt { path, .. } if path.contains("draw:connector")
                ) && object.kind == ObjectKind::Shape
            })
            .expect("connector remains a painted source-mapped shape");
        let Visual::PaintedShape {
            geometry: Geometry::Path { commands, .. },
            ..
        } = &connector.visual
        else {
            panic!("default ODF connector must preserve its routed path and marker");
        };
        assert!(commands.len() >= 6);
        let label = parsed
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("No"))
            .expect("connector label remains visible");
        let Visual::RichText { runs, .. } = &label.visual else {
            panic!("connector label must remain styled text");
        };
        assert_eq!(runs[0].color, 0xc11e_25ff);
        let paragraph = parsed
            .objects
            .iter()
            .find(|object| object.kind == ObjectKind::Paragraph)
            .unwrap();
        let group_top = parsed
            .objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Shape)
            .map(|object| object.bounds.y)
            .reduce(f32::min)
            .unwrap();
        assert_eq!(
            group_top,
            paragraph.bounds.y + paragraph.bounds.height + 20.0
        );
    }

    #[test]
    fn renders_shapes_from_the_selected_master_page_header() {
        let styles_xml = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:text="text" xmlns:draw="draw" xmlns:svg="svg" xmlns:loext="loext">
          <office:automatic-styles>
            <style:style style:name="FirstParagraph" style:family="paragraph" style:master-page-name="First"/>
            <style:style style:name="Red" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#ff0000" draw:stroke="none" style:horizontal-rel="page-content" loext:allow-overlap="false"/></style:style>
            <style:style style:name="Green" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#92d050" draw:stroke="none" style:horizontal-rel="page" loext:allow-overlap="false"/></style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Standard"><style:header><text:p>standard</text:p></style:header></style:master-page>
            <style:master-page style:name="First"><style:header><text:p>
              <draw:custom-shape draw:name="red" draw:style-name="Red" draw:z-index="1" svg:x="0in" svg:y="1in" svg:width="4in" svg:height="1in"/>
              <draw:custom-shape draw:name="green" draw:style-name="Green" draw:z-index="0" svg:x="1in" svg:y="0.8in" svg:width="4in" svg:height="0.25in"/>
            </text:p></style:header></style:master-page>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text"><office:body><office:text>
          <text:p text:style-name="FirstParagraph">body text</text:p>
        </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        assert!(!styles.graphic(Some("Red")).allow_overlap);
        assert!(!styles.graphic(Some("Green")).allow_overlap);
        super::parse_master_stories_xml(styles_xml, Limits::default(), &mut styles).unwrap();
        super::select_initial_master_page(content, Limits::default(), &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();

        let master_shapes = parsed
            .objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Shape && object.source.part == STYLES_PART)
            .collect::<Vec<_>>();
        assert_eq!(master_shapes.len(), 2);
        assert!(matches!(
            master_shapes[0].visual,
            Visual::PaintedShape {
                fill: crate::model::Paint::Solid(0xff00_00ff),
                ..
            }
        ));
        assert!(matches!(
            master_shapes[1].visual,
            Visual::PaintedShape {
                fill: crate::model::Paint::Solid(0x92d0_50ff),
                ..
            }
        ));
        assert_eq!(master_shapes[0].bounds.x, PageLayout::default().margin_left);
        assert_eq!(
            master_shapes[1].bounds.y,
            master_shapes[0].bounds.y + master_shapes[0].bounds.height
        );
        let body = parsed
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("body text"))
            .unwrap();
        assert!(body.bounds.y >= master_shapes[1].bounds.y + master_shapes[1].bounds.height);
        assert!(
            !parsed
                .objects
                .iter()
                .any(|object| object.text.as_deref() == Some("standard"))
        );
    }

    #[test]
    fn renders_master_header_with_paragraph_font_size() {
        let styles_xml = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:text="text" xmlns:fo="fo">
          <office:styles>
            <style:style style:name="Header" style:family="paragraph">
              <style:text-properties fo:font-size="40pt"/>
            </style:style>
          </office:styles>
          <office:master-styles>
            <style:master-page style:name="Standard">
              <style:header>
                <text:p text:style-name="Header">Seite <text:page-number text:select-page="current">1</text:page-number><text:s/>von <text:page-count>4</text:page-count></text:p>
              </style:header>
            </style:master-page>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text"><office:body><office:text>
          <text:p>body text</text:p>
        </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        super::parse_master_stories_xml(styles_xml, Limits::default(), &mut styles).unwrap();
        super::select_initial_master_page(content, Limits::default(), &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();

        let header = parsed
            .objects
            .iter()
            .find(|object| {
                object.source.part == STYLES_PART
                    && object
                        .text
                        .as_deref()
                        .is_some_and(|text| text.starts_with("Seite "))
            })
            .expect("master header paragraph is rendered");
        // Dynamic ODF page fields ignore stale cached results and report the
        // laid-out page number and page count of this single-page document.
        assert_eq!(header.text.as_deref(), Some("Seite 1 von 1"));
        let Visual::RichText {
            runs, line_height, ..
        } = &header.visual
        else {
            panic!("master header paragraphs use rich text");
        };
        let expected_size = 40.0 * CSS_PIXELS_PER_INCH / 72.0;
        assert!((runs[0].font_size - expected_size).abs() < 0.01);
        assert!(*line_height >= expected_size * 1.15);
    }

    #[test]
    fn maps_odt_drop_cap_and_ruby_onto_shared_text_layout() {
        let styles_xml =
            br#"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
          <office:styles>
            <style:style style:name="Drop" style:family="paragraph">
              <style:paragraph-properties>
                <style:drop-cap style:lines="3" style:length="1" style:distance="2pt"/>
              </style:paragraph-properties>
            </style:style>
          </office:styles>
        </office:document-styles>"#;
        let content = r#"<office:document-content xmlns:office="office" xmlns:text="text">
          <office:body><office:text>
            <text:p text:style-name="Drop">Tale of <text:ruby><text:ruby-base>漢字</text:ruby-base><text:ruby-text>かんじ</text:ruby-text></text:ruby></text:p>
          </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        let parsed = parse_content_xml(
            content.as_bytes(),
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();
        let paragraph = parsed
            .objects
            .iter()
            .find(|object| object.kind == ObjectKind::Paragraph)
            .expect("drop-cap paragraph");
        let Visual::TextLayout { layout, visual } = &paragraph.visual else {
            panic!("paragraphs use shared text layout");
        };
        let cap = layout.paragraphs[0]
            .drop_cap
            .expect("style:drop-cap reaches TextDropCap");
        assert_eq!(cap.characters, 1);
        assert_eq!(cap.lines, 3);
        assert!((cap.padding - 2.0 * CSS_PIXELS_PER_INCH / 72.0).abs() < 0.01);
        assert_eq!(paragraph.text.as_deref(), Some("Tale of 漢字かんじ"));
        let Visual::RichText { runs, .. } = &**visual else {
            panic!("nested rich text");
        };
        let ruby = runs
            .iter()
            .find(|run| run.text == "かんじ")
            .expect("ruby text run");
        assert!(ruby.baseline_shift > 0.0);
        assert!(ruby.font_size < runs[0].font_size);
    }

    #[test]
    fn renders_odt_measure_and_regular_polygon_shapes() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text">
          <office:body><office:text>
            <draw:measure svg:x1="1in" svg:y1="1in" svg:x2="3in" svg:y2="1in"/>
            <draw:regular-polygon svg:x="1in" svg:y="2in" svg:width="1in" svg:height="1in" draw:corners="5"/>
          </office:text></office:body></office:document-content>"#;
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut StyleCatalog::default(),
            None,
        )
        .expect("measure and polygon parse");
        assert!(
            parsed.objects.iter().any(|object| matches!(
                &object.visual,
                Visual::PaintedShape {
                    geometry: Geometry::Line,
                    ..
                }
            )),
            "measure must paint a line: {:?}",
            parsed.objects
        );
        assert!(
            parsed.objects.iter().any(|object| matches!(
                &object.visual,
                Visual::PaintedShape {
                    geometry: Geometry::Path { commands, .. },
                    ..
                } if commands.len() >= 6
            )),
            "regular-polygon must paint a pentagon path: {:?}",
            parsed.objects
        );
    }

    #[test]
    fn renders_odt_table_cell_images_with_source_mapping() {
        const PNG: &[u8] = b"\x89PNG\r\n\x1a\n";
        let bytes = crate::format::presentation_image::stored_zip(&[
            ("mimetype", b"application/vnd.oasis.opendocument.text"),
            (
                "styles.xml",
                br#"<office:document-styles xmlns:office="office"/>"#,
            ),
            (
                "content.xml",
                br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:table="table" xmlns:draw="draw" xmlns:xlink="http://www.w3.org/1999/xlink">
                  <office:body><office:text>
                    <table:table><table:table-row>
                      <table:table-cell><draw:frame><draw:image xlink:href="Pictures/cell.png"/></draw:frame><text:p>InCell</text:p></table:table-cell>
                    </table:table-row></table:table>
                  </office:text></office:body></office:document-content>"#,
            ),
            ("Pictures/cell.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).expect("zip package");
        let document = super::parse(&package).expect("table image document");
        let image = document
            .objects
            .iter()
            .find(|object| object.kind == ObjectKind::Image)
            .expect("table cell image is rendered");
        assert!(matches!(&image.visual, Visual::Image { bytes, .. } if bytes.as_slice() == PNG));
        assert!(
            matches!(&image.source.locator, SourceLocator::Odt { path, .. } if path.contains("draw:image")),
            "{:?}",
            image.source
        );
        assert!(
            document
                .objects
                .iter()
                .any(|object| object.text.as_deref() == Some("InCell"))
        );
        assert!(
            !document.diagnostics.iter().any(|diagnostic| diagnostic
                .message
                .contains("not rendered in the table fallback")),
            "{:?}",
            document.diagnostics
        );
    }

    #[test]
    fn empty_odf_fields_use_their_value_attributes() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:office2="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
          <office:body><office:text>
            <text:p>On <text:date office:date-value="2024-01-15T00:00:00"/> by <text:author office:string-value="Ada"/></text:p>
          </office:text></office:body></office:document-content>"#;
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut StyleCatalog::default(),
            None,
        )
        .unwrap();
        let paragraph = parsed
            .objects
            .iter()
            .find(|object| {
                object
                    .text
                    .as_deref()
                    .is_some_and(|text| text.starts_with("On "))
            })
            .expect("field paragraph");
        assert_eq!(
            paragraph.text.as_deref(),
            Some("On 2024-01-15 by Ada"),
            "{:?}",
            paragraph.text
        );
    }

    #[test]
    fn empty_odf_page_fields_resolve_against_the_laid_out_document() {
        let styles_xml = br#"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:text="text">
          <office:master-styles>
            <style:master-page style:name="Standard">
              <style:footer>
                <text:p><text:page-number text:select-page="current"/>/<text:page-count/></text:p>
              </style:footer>
            </style:master-page>
          </office:master-styles>
        </office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:fo="fo" xmlns:style="style">
          <office:automatic-styles>
            <style:style style:name="P" style:family="paragraph">
              <style:paragraph-properties fo:margin-top="0in" fo:margin-bottom="0in"/>
              <style:text-properties fo:font-size="12pt"/>
            </style:style>
          </office:automatic-styles>
          <office:body><office:text>
            <text:p text:style-name="P">one</text:p>
            <text:p text:style-name="P">two</text:p>
          </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        super::parse_master_stories_xml(styles_xml, Limits::default(), &mut styles).unwrap();
        super::select_initial_master_page(content, Limits::default(), &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();
        let footers: Vec<_> = parsed
            .objects
            .iter()
            .filter(|object| {
                object.source.part == STYLES_PART
                    && object
                        .text
                        .as_deref()
                        .is_some_and(|text| text.contains('/'))
            })
            .collect();
        assert!(!footers.is_empty(), "{:?}", parsed.objects);
        for footer in &footers {
            let text = footer.text.as_deref().unwrap();
            assert!(
                text.starts_with(&format!("{}/", footer.unit_index + 1)),
                "footer page slot must match its page: {text} on unit {}",
                footer.unit_index
            );
            assert!(
                text.ends_with(&format!("/{}", parsed.units.len())),
                "footer page count must match the laid-out document: {text}",
            );
        }
    }

    #[test]
    fn omits_nested_frames_with_image_alternatives() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink"><office:body><office:text>
          <draw:frame svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in">
            <draw:image xlink:href="Pictures/preferred.png"/>
            <draw:image xlink:href="Pictures/fallback.png"/>
            <draw:text-box><text:p>Outer<draw:frame svg:width="1in" svg:height="1in"><draw:text-box><text:p>Nested</text:p></draw:text-box></draw:frame></text:p></draw:text-box>
          </draw:frame>
        </office:text></office:body></office:document-content>"#;
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut StyleCatalog::default(),
            None,
        )
        .expect("nested frames and image alternatives are non-fatal");

        assert!(
            parsed
                .diagnostics
                .iter()
                .any(|diagnostic| { diagnostic.message.contains("nested ODT frames are omitted") })
        );
    }

    #[test]
    fn nested_text_box_images_keep_geometry_text_and_resource_isolation() {
        const PNG: &[u8] = b"\x89PNG\r\n\x1a\n";
        let bytes = stored_zip(&[("Pictures/good.png", PNG)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        for dimensions in ["", "svg:width=\"2in\" svg:height=\"1in\""] {
            let xml = format!(
                r#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink"><office:body><office:text>
              <draw:frame svg:x="10px" svg:y="20px" {dimensions}><draw:text-box><text:p>Outer<draw:frame svg:x="3px" svg:y="4px" svg:width="16px" svg:height="10px"><draw:image xlink:href="https://example.invalid/image.png"/><draw:image xlink:href="Pictures/missing.png"/><draw:image xlink:href="Pictures/good.png"/></draw:frame>Tail</text:p></draw:text-box></draw:frame><text:p>After</text:p>
            </office:text></office:body></office:document-content>"#
            );
            let parsed = parse_content_xml(
                xml.as_bytes(),
                PageLayout::default(),
                Limits::default(),
                &mut StyleCatalog::default(),
                Some(&package),
            )
            .unwrap();
            let image = parsed
                .objects
                .iter()
                .find(|o| o.kind == ObjectKind::Image)
                .unwrap();
            assert_eq!(
                image.bounds,
                Rect {
                    x: 109.0,
                    y: 120.0,
                    width: 16.0,
                    height: 10.0
                }
            );
            assert!(matches!(&image.visual, Visual::Image { bytes, .. } if bytes == PNG));
            let parent = &parsed.objects[image.parent_numeric_id.unwrap() as usize];
            assert_eq!(
                (parent.bounds.width, parent.bounds.height),
                if dimensions.is_empty() {
                    (19.0, 14.0)
                } else {
                    (192.0, 96.0)
                }
            );
            assert!(
                parsed
                    .objects
                    .iter()
                    .any(|o| o.text.as_deref() == Some("OuterTail"))
            );
            assert!(
                parsed
                    .objects
                    .iter()
                    .any(|o| o.text.as_deref() == Some("After"))
            );
            assert!(
                parsed
                    .diagnostics
                    .iter()
                    .any(|d| d.code == DiagnosticCode::ExternalResourceBlocked)
            );
            assert!(
                parsed
                    .diagnostics
                    .iter()
                    .any(|d| d.message.contains("missing from the package"))
            );
        }
    }

    #[test]
    fn uses_supported_fallback_for_odt_frame_image_alternatives() {
        const PNG: &[u8] = b"\x89PNG\r\n\x1a\n";
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink"><office:body><office:text>
          <draw:frame svg:width="2in" svg:height="1in">
            <draw:image xlink:href="Pictures/preferred.pdf"/>
            <draw:image xlink:href="Pictures/fallback.png"/>
          </draw:frame>
        </office:text></office:body></office:document-content>"#;
        let bytes = stored_zip(&[
            ("Pictures/preferred.pdf", b"%PDF-1.4"),
            ("Pictures/fallback.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut StyleCatalog::default(),
            Some(&package),
        )
        .unwrap();

        let frame = parsed
            .objects
            .iter()
            .find(|object| object.stable_id == "odt:frame:0:visual")
            .unwrap();
        assert!(matches!(
            &frame.visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/png" && bytes == PNG
        ));
    }

    #[test]
    fn renders_paragraph_anchored_form_images_below_higher_z_shapes() {
        const PNG: &[u8] = b"\x89PNG\r\n\x1a\n";
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:draw="draw" xmlns:form="form" xmlns:svg="svg"><office:body><office:text>
          <form:form><form:image-frame xml:id="control1" form:image-data="Pictures/map.png"/></form:form>
          <text:p>Before</text:p>
          <text:p>
            <draw:control text:anchor-type="paragraph" draw:control="control1" draw:z-index="0" svg:x="10px" svg:y="20px" svg:width="100px" svg:height="50px"/>
            <draw:custom-shape text:anchor-type="paragraph" draw:z-index="1" svg:x="20px" svg:y="30px" svg:width="80px" svg:height="20px"/>
          </text:p>
        </office:text></office:body></office:document-content>"#;
        let bytes = stored_zip(&[("Pictures/map.png", PNG)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut StyleCatalog::default(),
            Some(&package),
        )
        .unwrap();
        let image = parsed
            .objects
            .iter()
            .find(|object| object.kind == ObjectKind::Image)
            .expect("the form image must use the existing ODT image pipeline");
        let shape = parsed
            .objects
            .iter()
            .find(|object| object.kind == ObjectKind::Shape)
            .expect("the overlapping shape remains visible");

        assert_eq!(shape.bounds.y, image.bounds.y + 10.0);
        assert!(shape.z > image.z);
    }

    #[test]
    fn keeps_text_when_a_frame_image_is_missing_from_the_package() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink"><office:body><office:text>
          <text:p>Before<draw:frame svg:width="2in" svg:height="1in"><draw:image xlink:href="Pictures/missing.png"/></draw:frame></text:p>
        </office:text></office:body></office:document-content>"#;
        let bytes = stored_zip(&[("unrelated", b"")]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut StyleCatalog::default(),
            Some(&package),
        )
        .expect("a missing frame image must not suppress usable ODT text");

        assert!(
            parsed
                .objects
                .iter()
                .any(|object| object.text.as_deref() == Some("Before"))
        );
        assert!(parsed.diagnostics.iter().any(|diagnostic| {
            diagnostic.message.contains("missing from the package")
                && diagnostic.location.part.as_deref() == Some("Pictures/missing.png")
        }));
    }

    fn sample_document() -> crate::model::Document {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/file-sample_1MB.odt"),
            Limits::default(),
        )
        .unwrap();
        super::parse(&package).unwrap()
    }

    #[test]
    fn supplied_sample_resolves_primary_font_family() {
        let document = sample_document();
        let heading = document
            .objects
            .iter()
            .find(|o| {
                o.text
                    .as_deref()
                    .is_some_and(|t| t.starts_with("Lorem ipsum dolor"))
            })
            .unwrap();
        let Visual::TextLayout { visual, .. } = &heading.visual else {
            panic!("heading layout")
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("heading text")
        };
        assert_eq!(runs[0].font_family, "Liberation Sans");
    }

    #[test]
    fn supplied_sample_preserves_justified_body() {
        let document = sample_document();
        let body = document
            .objects
            .iter()
            .find(|o| {
                o.text
                    .as_deref()
                    .is_some_and(|t| t.starts_with("Vestibulum neque"))
            })
            .unwrap();
        let Visual::TextLayout { layout, .. } = &body.visual else {
            panic!("paragraph layout")
        };
        assert_eq!(layout.paragraphs[0].align, TextAlign::Justify);
    }

    #[test]
    fn supplied_sample_preserves_list_body_and_unnumbered_heading_positions() {
        let document = sample_document();
        let body = document
            .objects
            .iter()
            .find(|o| {
                o.text
                    .as_deref()
                    .is_some_and(|t| t.starts_with("Vestibulum neque"))
            })
            .unwrap();
        let heading = document
            .objects
            .iter()
            .find(|o| {
                o.text
                    .as_deref()
                    .is_some_and(|t| t.starts_with("Lorem ipsum dolor"))
            })
            .unwrap();
        assert_eq!(
            heading.bounds.x, body.bounds.x,
            "suppressed numbering must not add a list indent"
        );
        let bullet = document
            .objects
            .iter()
            .find(|o| {
                o.text
                    .as_deref()
                    .is_some_and(|t| t.starts_with("Maecenas non lorem"))
            })
            .unwrap();
        let Visual::TextLayout { layout, .. } = &bullet.visual else {
            panic!("list layout")
        };
        assert!(
            (bullet.bounds.x + layout.paragraphs[0].margin_left - body.bounds.x - 48.0).abs()
                < 0.01
        );
        assert_eq!(layout.paragraphs[0].first_line_indent, -24.0);
        assert!(layout.tab_stops.iter().any(|stop| stop.position == 48.0));
    }

    #[test]
    fn supplied_sample_table_uses_text_metrics_and_paragraph_spacing() {
        let document = sample_document();
        let cell = |text| {
            document
                .objects
                .iter()
                .find(|o| o.kind == ObjectKind::Cell && o.text.as_deref() == Some(text))
                .unwrap()
        };
        assert!(
            (cell("1").bounds.height - cell("3").bounds.height).abs() < 0.01,
            "both body rows fit one line with the same paragraph spacing"
        );
        assert!(
            cell("3").bounds.height > 30.0,
            "cell padding and paragraph spacing must contribute"
        );
    }

    #[test]
    fn supplied_sample_image_follows_its_anchor_page_and_reserves_space() {
        let document = sample_document();
        let image = document
            .objects
            .iter()
            .find(|o| o.kind == ObjectKind::Image && o.bounds.width > 600.0)
            .unwrap();
        let previous = document
            .objects
            .iter()
            .take_while(|o| o.numeric_id != image.numeric_id)
            .filter(|o| {
                o.kind == ObjectKind::Paragraph && o.text.as_deref().is_some_and(|t| !t.is_empty())
            })
            .last()
            .unwrap();
        assert_eq!(
            image.unit_index,
            previous.unit_index + 1,
            "the cached anchor break must move the image with its paragraph"
        );
        assert!(
            image.bounds.y + image.bounds.height
                < document.units[image.unit_index as usize].height - 70.0
        );
        let following = document
            .objects
            .iter()
            .find(|o| {
                o.kind == ObjectKind::Paragraph
                    && o.unit_index == image.unit_index
                    && o.text
                        .as_deref()
                        .is_some_and(|t| t.starts_with("Maecenas mauris"))
            })
            .unwrap();
        assert!(following.bounds.y >= image.bounds.y + image.bounds.height);
    }

    #[test]
    fn supplied_sample_suppresses_empty_number_formats_but_keeps_bullets() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/file-sample_1MB.odt"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let mut headings = 0;
        let mut bullets = 0;
        for object in &document.objects {
            let Visual::TextLayout { layout, .. } = &object.visual else {
                continue;
            };
            if object.text.as_deref().is_some_and(|text| {
                text.starts_with("Lorem ipsum dolor sit amet")
                    || text.starts_with("Cras fringilla ipsum magna")
                    || text.starts_with("Maecenas mauris lectus")
            }) {
                headings += 1;
                assert_eq!(layout.prefix, None, "{:?}", object.text);
                assert_eq!(layout.hanging_indent, 0.0);
            }
            if let Some(prefix) = &layout.prefix {
                assert_eq!(
                    prefix, "•\t",
                    "unexpected generated number: {:?}",
                    object.text
                );
                bullets += 1;
            }
        }
        assert!(headings >= 4, "all supplied headings must survive");
        assert_eq!(bullets, 6);
    }

    #[test]
    fn renders_list_labels_and_real_table_spans() {
        let styles_xml = br#"<office:document-styles xmlns:office="office" xmlns:text="text" xmlns:style="style" xmlns:svg="svg">
          <office:font-face-decls><style:font-face style:name="Bullet Face" svg:font-family="Wingdings"/></office:font-face-decls>
          <office:styles><text:list-style style:name="L1">
            <text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/>
            <text:list-level-style-bullet text:level="2" text:bullet-char="q"><style:text-properties style:font-name="Bullet Face"/></text:list-level-style-bullet>
          </text:list-style></office:styles>
        </office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:table="table"><office:body><office:text>
          <text:list text:style-name="L1"><text:list-item><text:p>First</text:p><text:list><text:list-item><text:p>Nested</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>Second</text:p></text:list-item></text:list>
          <table:table><table:table-column table:number-columns-repeated="2"/>
            <table:table-row><table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2"><text:p>Span</text:p></table:table-cell><table:covered-table-cell/></table:table-row>
            <table:table-row><table:covered-table-cell/><table:covered-table-cell/></table:table-row>
          </table:table>
        </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();
        let labels = parsed
            .objects
            .iter()
            .filter_map(|object| match &object.visual {
                Visual::TextLayout { layout, .. } => layout.prefix.as_deref(),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(labels, vec!["1.\t", "❑\t", "2.\t"]);
        let span = parsed
            .objects
            .iter()
            .find(|object| {
                object.kind == ObjectKind::Cell && object.text.as_deref() == Some("Span")
            })
            .unwrap();
        let content_width = PageLayout::default().width
            - PageLayout::default().margin_left
            - PageLayout::default().margin_right;
        assert_eq!(span.bounds.width, content_width);
        assert!(span.bounds.height >= LINE_HEIGHT * 2.0);
    }

    #[test]
    fn authored_odt_table_column_widths_preserve_authored_ratios() {
        let styles_xml =
            br#"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
          <office:automatic-styles>
            <style:style style:name="coNarrow" style:family="table-column">
              <style:table-column-properties style:column-width="1cm"/>
            </style:style>
            <style:style style:name="coWide" style:family="table-column">
              <style:table-column-properties style:column-width="3cm"/>
            </style:style>
          </office:automatic-styles>
        </office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:table="table"><office:body><office:text>
          <table:table>
            <table:table-column table:style-name="coNarrow"/>
            <table:table-column table:style-name="coWide"/>
            <table:table-row>
              <table:table-cell><text:p>Narrow</text:p></table:table-cell>
              <table:table-cell><text:p>Wide</text:p></table:table-cell>
            </table:table-row>
          </table:table>
        </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();
        let narrow = parsed
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Narrow"))
            .unwrap();
        let wide = parsed
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Wide"))
            .unwrap();
        assert!(
            (wide.bounds.width - narrow.bounds.width * 3.0).abs() < 0.05,
            "authored 1:3 column widths must survive layout: {} vs {}",
            narrow.bounds.width,
            wide.bounds.width
        );
        assert!((wide.bounds.x - (narrow.bounds.x + narrow.bounds.width)).abs() < 0.05);
        assert!(
            !parsed
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("equal-width fallback")),
            "{:?}",
            parsed.diagnostics
        );
    }

    #[test]
    fn relative_odt_table_column_widths_use_star_weights() {
        let styles_xml = br#"<office:document-styles xmlns:office="office" xmlns:style="style">
          <office:styles>
            <style:style style:name="co1" style:family="table-column">
              <style:table-column-properties style:rel-column-width="1*"/>
            </style:style>
            <style:style style:name="co2" style:family="table-column">
              <style:table-column-properties style:rel-column-width="3*+"/>
            </style:style>
          </office:styles>
        </office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:table="table"><office:body><office:text>
          <table:table>
            <table:table-column table:style-name="co1"/>
            <table:table-column table:style-name="co2"/>
            <table:table-row>
              <table:table-cell><text:p>One</text:p></table:table-cell>
              <table:table-cell><text:p>Three</text:p></table:table-cell>
            </table:table-row>
          </table:table>
        </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();
        let one = parsed
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("One"))
            .unwrap();
        let three = parsed
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Three"))
            .unwrap();
        assert!(
            (three.bounds.width - one.bounds.width * 3.0).abs() < 0.05,
            "star weights 1* and 3*+ must keep a 1:3 ratio: {} vs {}",
            one.bounds.width,
            three.bounds.width
        );
    }

    #[test]
    fn missing_odt_table_column_widths_keep_equal_width_fallback() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:table="table"><office:body><office:text>
          <table:table><table:table-column table:number-columns-repeated="2"/>
            <table:table-row>
              <table:table-cell><text:p>Left</text:p></table:table-cell>
              <table:table-cell><text:p>Right</text:p></table:table-cell>
            </table:table-row>
          </table:table>
        </office:text></office:body></office:document-content>"#;
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut StyleCatalog::default(),
            None,
        )
        .unwrap();
        let left = parsed
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Left"))
            .unwrap();
        let right = parsed
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Right"))
            .unwrap();
        assert!((left.bounds.width - right.bounds.width).abs() < 0.05);
        assert!(
            parsed
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.message.contains("equal-width fallback")),
            "{:?}",
            parsed.diagnostics
        );
    }

    #[test]
    fn preserves_zero_padded_number_formats() {
        let styles_xml = br#"<office:document-styles xmlns:office="office" xmlns:text="text" xmlns:style="style"><office:styles>
          <text:list-style style:name="L2"><text:list-level-style-number text:level="1" style:num-format="01, 02, 03, ..." style:num-suffix="."/></text:list-style>
          <text:list-style style:name="L5"><text:list-level-style-number text:level="1" style:num-format="00001, 00002, 00003, ..." style:num-suffix="."/></text:list-style>
        </office:styles></office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text"><office:body><office:text>
          <text:list text:style-name="L2"><text:list-item><text:p>Two digits</text:p></text:list-item></text:list>
          <text:list text:style-name="L5"><text:list-item><text:p>Five digits</text:p></text:list-item></text:list>
        </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();
        let labels = parsed
            .objects
            .iter()
            .filter_map(|object| match &object.visual {
                Visual::TextLayout { layout, .. } => layout.prefix.as_deref(),
                _ => None,
            })
            .collect::<Vec<_>>();

        assert_eq!(labels, ["01.\t", "00001.\t"]);
    }

    #[test]
    fn supplied_border_table_inherits_paragraph_font() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/odt-border-types.odt"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let cell = document
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("solid"))
            .unwrap();
        assert!(matches!(&cell.visual, Visual::TextLayout { visual, .. }
                if matches!(visual.as_ref(), Visual::Text { font_family, .. } if font_family == "Liberation Serif")));
    }

    #[test]
    fn supplied_merged_cells_without_borders_remain_unpainted() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/odt-merged-cells-without-border.odt"),
            Limits::default(),
        )
        .unwrap();
        let document = super::parse(&package).unwrap();
        let cells = document
            .objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Cell)
            .collect::<Vec<_>>();
        assert_eq!(cells.len(), 3);
        for cell in cells {
            assert!(
                matches!(&cell.visual, Visual::TextLayout { visual, .. }
                    if matches!(visual.as_ref(), Visual::Text { stroke_width, .. } if *stroke_width == 0.0))
                    || matches!(&cell.visual, Visual::PaintedShape { stroke_width, .. } if *stroke_width == 0.0),
                "no border was authored: {:?}",
                cell.visual
            );
        }
    }

    #[test]
    fn preserves_authored_table_border_styles() {
        let styles_xml = br#"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo"><office:styles>
          <style:style style:name="Dotted" style:family="table-cell"><style:table-cell-properties fo:border="1px dotted #112233"/></style:style>
          <style:style style:name="DashDot" style:family="table-cell"><style:table-cell-properties fo:border="1px dash-dot #112233"/></style:style>
          <style:style style:name="Double" style:family="table-cell"><style:table-cell-properties fo:border="1px double #112233"/></style:style>
        </office:styles></office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:table="table" xmlns:text="text"><office:body><office:text><table:table><table:table-row>
          <table:table-cell table:style-name="Dotted"><text:p>Dotted</text:p></table:table-cell>
          <table:table-cell table:style-name="DashDot"><text:p>Dash dot</text:p></table:table-cell>
          <table:table-cell table:style-name="Double"><text:p>Double</text:p></table:table-cell>
        </table:table-row></table:table></office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();
        let cells = parsed
            .objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Cell)
            .collect::<Vec<_>>();

        let Visual::StrokeStyle { style, .. } = &cells[0].visual else {
            panic!("dotted cell border must retain its stroke style");
        };
        assert_eq!(style.dash, [1.0, 2.0]);
        let Visual::StrokeStyle { style, .. } = &cells[1].visual else {
            panic!("dash-dot cell border must retain its stroke style");
        };
        assert_eq!(style.dash, [8.0, 4.0, 2.0, 4.0]);
        let Visual::StrokeStyle { style, .. } = &cells[2].visual else {
            panic!("double cell border must retain its stroke style");
        };
        assert_eq!(style.compound, crate::model::LineCompound::Double);
    }

    #[test]
    fn preserves_table_cell_writing_modes() {
        let styles_xml = br#"<office:document-content xmlns:office="office" xmlns:style="style"><office:automatic-styles>
          <style:style style:name="Table1.A1" style:family="table-cell"><style:table-cell-properties style:writing-mode="bt-lr"/></style:style>
          <style:style style:name="Table1.C1" style:family="table-cell"><style:table-cell-properties style:writing-mode="tb-rl"/></style:style>
          <style:style style:name="Table1.1" style:family="table-row"><style:table-row-properties style:min-row-height="2cm"/></style:style>
        </office:automatic-styles></office:document-content>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:table="table" xmlns:text="text"><office:body><office:text><table:table><table:table-row table:style-name="Table1.1">
          <table:table-cell table:style-name="Table1.A1"><text:p>AAA.</text:p></table:table-cell>
          <table:table-cell table:style-name="Table1.C1"><text:p>CCC.</text:p></table:table-cell>
        </table:table-row></table:table></office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();
        let cells = parsed
            .objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Cell)
            .collect::<Vec<_>>();

        let orientations = cells
            .iter()
            .map(|cell| match &cell.visual {
                Visual::TextLayout { layout, .. } => layout.orientation,
                _ => crate::model::TextOrientation::Horizontal,
            })
            .collect::<Vec<_>>();
        assert_eq!(
            orientations,
            [
                crate::model::TextOrientation::Rotated270,
                crate::model::TextOrientation::Rotated90
            ]
        );
        assert!(
            cells
                .iter()
                .all(|cell| cell.bounds.height >= 2.0 * CSS_PIXELS_PER_INCH / 2.54)
        );
    }

    #[test]
    fn applies_frame_transforms_and_renders_common_shapes() {
        let styles_xml = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:fo="fo" xmlns:svg="svg">
          <office:styles><style:style style:name="Graphic" style:family="graphic"><style:graphic-properties style:wrap="run-through" draw:fill="solid" draw:fill-color="#123456" draw:stroke="solid" svg:stroke-color="#654321" svg:stroke-width="2pt"/></style:style></office:styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:text="text" xmlns:draw="draw" xmlns:svg="svg"><office:body><office:text>
          <draw:frame draw:style-name="Graphic" draw:transform="rotate(0.5)" svg:x="10pt" svg:y="20pt" svg:width="30pt" svg:height="40pt"><draw:text-box><text:p>Rotated</text:p></draw:text-box></draw:frame>
          <draw:rect draw:style-name="Graphic" draw:transform="translate(5pt 0pt)" svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in"/>
          <draw:ellipse draw:style-name="Graphic" svg:x="3in" svg:y="1in" svg:width="1in" svg:height="1in"/>
          <draw:line draw:style-name="Graphic" svg:x1="1in" svg:y1="3in" svg:x2="3in" svg:y2="3in"/>
        </office:text></office:body></office:document-content>"#;
        let mut styles = StyleCatalog::default();
        parse_style_catalog_xml(styles_xml, Limits::default(), STYLES_PART, &mut styles).unwrap();
        let parsed = parse_content_xml(
            content,
            PageLayout::default(),
            Limits::default(),
            &mut styles,
            None,
        )
        .unwrap();
        let transformed = parsed
            .objects
            .iter()
            .find(|object| object.stable_id == "odt:frame:0:visual")
            .unwrap();
        assert!(matches!(transformed.visual, Visual::Layer { .. }));
        let shapes = parsed
            .objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Shape)
            .collect::<Vec<_>>();
        assert_eq!(shapes.len(), 3);
        assert!(matches!(shapes[0].visual, Visual::Layer { .. }));
    }
}
