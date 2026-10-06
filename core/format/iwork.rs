//! Bounded native detection for modern single-file iWork packages.
//!
//! Pages decodes its saved pagination, native text, and bounded inline images.
//! Numbers decodes native sheets, tables, tile storage, and cell values. Keynote
//! decodes a bounded static text/image subset and falls back per slide to an
//! embedded preview.
//! The adapter never converts to another office format.

use std::borrow::Cow;
use std::collections::{HashMap, HashSet};

use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
use crate::font_metrics::FontMetricTable;
use crate::format::embedded_media::embedded_media_type;
use crate::format::presentation_image::office_image_media_type;
use crate::limits::Limits;
use crate::model::{
    AffineTransform, Document, DocumentFormat, DocumentKind, EmbeddedFont, FillRule, Geometry,
    GradientStop, ImageCrop, LineAlignment, LineCap, LineJoin, MappingQuality, Object, ObjectKind,
    Paint, PathCommand, Rect, Reflection, Shadow, SourceLocator, SourceRef, StrokeStyle, TextAlign,
    TextAutoFit, TextHorizontalOverflow, TextLayout, TextOrientation, TextParagraphLayout,
    TextParagraphRule, TextRun, TextTabLeader, TextTabStop, TextVerticalAlign,
    TextVerticalOverflow, Unit, UnitKind, Visual,
};
use crate::package::Package;
use crate::xml::{XmlEvent, decode_xml_text, parse_xml};

const DOCUMENT_COMPONENT: &str = "Index/Document.iwa";
const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
const STYLESHEET_COMPONENT: &str = "Index/DocumentStylesheet.iwa";
const THEME_STYLESHEET_COMPONENT: &str = "Index/ThemeStylesheet.iwa";
const CALCULATION_ENGINE_COMPONENT: &str = "Index/CalculationEngine.iwa";
const PROPERTIES_PART: &str = "Metadata/Properties.plist";
const PAGES_DOCUMENT_TYPE: u64 = 10_000;
const PAGES_FLOATING_DRAWABLES_TYPE: u64 = 10_010;
const PAGES_LAYOUT_STATE_TYPE: u64 = 10_131;
const PAGES_VIEW_STATE_TYPE: u64 = 10_133;
const PAGES_TOC_TYPE: u64 = 2_240;
const PAGES_TOC_ATTACHMENT_TYPE: u64 = 2_241;
const PAGES_TOC_LAYOUT_TYPE: u64 = 2_242;
const IWORK_TOC_PARAGRAPH_STYLE_TYPE: u64 = 2_026;
const IWORK_TEXT_FIELD_TYPE: u64 = 2_010;
const IWORK_NUMBER_ATTACHMENT_TYPE: u64 = 2_043;
const PAGES_SECTION_TYPE: u64 = 10_011;
const PAGES_SECTION_TEMPLATE_TYPE: u64 = 10_143;
const NUMBERS_OR_KEYNOTE_DOCUMENT_TYPE: u64 = 1;
const NUMBERS_SHEET_TYPE: u64 = 2;
const NUMBERS_TABLE_INFO_TYPE: u64 = 6_000;
const NUMBERS_TABLE_MODEL_TYPE: u64 = 6_001;
const NUMBERS_TILE_TYPE: u64 = 6_002;
const NUMBERS_TILE_ROW_CAPACITY: u32 = 256;
const NUMBERS_DATA_LIST_TYPE: u64 = 6_005;
const NUMBERS_HEADER_STORAGE_TYPE: u64 = 6_006;
const NUMBERS_STROKE_SIDECAR_TYPE: u64 = 6_305;
const NUMBERS_STROKE_LAYER_TYPE: u64 = 6_306;
const KEYNOTE_SHOW_TYPE: u64 = 2;
const KEYNOTE_SLIDE_NODE_TYPE: u64 = 4;
const KEYNOTE_SLIDE_TYPE: u64 = 5;
const KEYNOTE_PLACEHOLDER_TYPE: u64 = 7;
const IWORK_TEXT_STORAGE_TYPE: u64 = 2_001;
const IWORK_TEXT_FLOW_TYPE: u64 = 2_410;
const IWORK_DRAWABLE_ATTACHMENT_TYPE: u64 = 2_003;
const IWORK_TEXT_SHAPE_TYPE: u64 = 2_011;
const IWORK_CAPTION_TYPE: u64 = 633;
const IWORK_CAPTION_PLACEMENT_TYPE: u64 = 634;
const IWORK_CHARACTER_STYLE_TYPE: u64 = 2_021;
const PAGES_DROP_CAP_STYLE_TYPE: u64 = 10_024;
const IWORK_PARAGRAPH_STYLE_TYPE: u64 = 2_022;
const IWORK_LIST_STYLE_TYPE: u64 = 2_023;
const IWORK_TEXT_SHAPE_STYLE_TYPE: u64 = 2_025;
const IWORK_SHAPE_TYPE: u64 = 3_004;
const IWORK_IMAGE_TYPE: u64 = 3_005;
const IWORK_MASK_TYPE: u64 = 3_006;
const IWORK_MEDIA_TYPE: u64 = 3_007;
const IWORK_GROUP_TYPE: u64 = 3_008;
const IWORK_CONNECTION_LINE_TYPE: u64 = 3_009;
const IWORK_SHAPE_STYLE_TYPE: u64 = 3_015;
const IWORK_MEDIA_STYLE_TYPE: u64 = 3_016;
const IWORK_CHART_TYPE: u64 = 5_021;
const IWORK_CHART_STYLE_TYPE: u64 = 5_022;
const IWORK_CHART_TITLE_TYPE: u64 = 5_023;
const IWORK_CHART_AXIS_STYLE_TYPE: u64 = 5_026;
const IWORK_CHART_AXIS_NONSTYLE_TYPE: u64 = 5_027;
const IWORK_CHART_SERIES_STYLE_TYPE: u64 = 5_028;
const IWORK_CHART_SERIES_NONSTYLE_TYPE: u64 = 5_029;
const IWORK_TABLE_TYPE: u64 = 6_000;
const KEYNOTE_SLIDE_STYLE_TYPE: u64 = 9;
const IWORK_PACKAGE_METADATA_TYPE: u64 = 11_006;
const MAX_IWA_PROTOBUF_FIELDS: usize = 1_024;
const IWORK_POINT_TO_CSS_PIXEL: f32 = 96.0 / 72.0;
const KEYNOTE_PINGFANG_GLYPH_SCALE: f32 = 0.99;
const KEYNOTE_HELVETICA_GLYPH_SCALE: f32 = 0.94;
const BINARY_PLIST_TRAILER_BYTES: usize = 32;
const CANONICAL_PLIST_DOCTYPE: &[u8] = br#"<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">"#;

#[derive(Clone)]
struct IwaArchive {
    identifier: u64,
    messages: Vec<IwaMessage>,
}

#[derive(Clone)]
struct IwaMessage {
    message_type: u64,
    payload: Vec<u8>,
    data_references: Vec<u64>,
}

struct IwaMessageInfo {
    message_type: u64,
    payload_length: usize,
    data_references: Vec<u64>,
}

struct IworkPreview {
    part: String,
    media_type: &'static str,
    width: f32,
    height: f32,
    bytes: Vec<u8>,
}

struct KeynoteShow {
    slide_nodes: Vec<u64>,
    size: Option<(f32, f32)>,
}

#[derive(Clone, Copy, Debug, Default)]
struct KeynoteGeometry {
    x: f32,
    y: f32,
    width: f32,
    height: f32,
    rotation_degrees: f32,
    resize_flags: Option<u64>,
    flip_horizontal: bool,
    flip_vertical: bool,
}

#[derive(Default)]
struct KeynoteNativeStats {
    slides: usize,
    unsupported_drawables: usize,
}

#[derive(Clone, Debug)]
struct KeynoteTextStyle {
    font_family: String,
    font_size: f32,
    color: u32,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    highlight: u32,
    baseline_shift: f32,
    tracking: f32,
    all_caps: bool,
    align: TextAlign,
    natural_alignment: bool,
    keep_lines_together: bool,
    widow_control: bool,
    line_height_multiple: Option<f32>,
    margin_left: f32,
    margin_right: f32,
    first_line_indent: f32,
    default_tab_stop: f32,
    tab_stops: Vec<TextTabStop>,
    space_before: f32,
    space_after: f32,
    list_level: usize,
    rule_enabled: bool,
    border_positions: Option<u64>,
    rule_offset: (f32, f32),
    rule_width: f32,
    rule_stroke: Option<(Paint, f32, StrokeStyle)>,
}

#[derive(Clone, Debug)]
struct IworkTableCellStyle {
    fill: Paint,
    wrap: bool,
}

#[derive(Clone, Debug, PartialEq)]
struct NumbersStroke {
    color: u32,
    width: f32,
    style: StrokeStyle,
}

#[derive(Clone, Debug)]
struct NumbersTableBorder {
    horizontal: bool,
    boundary: u32,
    origin: u32,
    length: u32,
    stroke: NumbersStroke,
}

impl Default for IworkTableCellStyle {
    fn default() -> Self {
        Self {
            fill: Paint::None,
            wrap: false,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum KeynoteListMarkerKind {
    Text(String),
    Number(u64),
}

#[derive(Clone, Debug, PartialEq)]
struct KeynoteListMarker {
    kind: KeynoteListMarkerKind,
    level: usize,
    font_family: Option<String>,
    scale: f32,
    color: Option<u32>,
    indent: f32,
    text_indent: f32,
}

impl Default for KeynoteTextStyle {
    fn default() -> Self {
        Self {
            font_family: "Arial".to_owned(),
            font_size: 28.0,
            color: 0x0000_00ff,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            highlight: 0,
            baseline_shift: 0.0,
            tracking: 0.0,
            all_caps: false,
            align: TextAlign::Start,
            natural_alignment: true,
            keep_lines_together: false,
            widow_control: true,
            line_height_multiple: None,
            margin_left: 0.0,
            margin_right: 0.0,
            first_line_indent: 0.0,
            default_tab_stop: 36.0,
            tab_stops: Vec::new(),
            space_before: 0.0,
            space_after: 0.0,
            list_level: 0,
            rule_enabled: false,
            border_positions: None,
            rule_offset: (0.0, 0.0),
            rule_width: 1.0,
            rule_stroke: None,
        }
    }
}

impl KeynoteTextStyle {
    fn paragraph_indents(&self, list_offset: Option<f32>) -> (f32, f32) {
        let left = list_offset.unwrap_or(self.margin_left);
        // TSWP's first-line position is absolute; the shared model is relative
        // to the body indent. Generated list markers start at the frame origin.
        (
            left,
            list_offset.map_or(self.first_line_indent, |_| 0.0) - left,
        )
    }

    fn border_padding(&self) -> f32 {
        // Pages stores historical deltas: a saved 0 is 6pt in its inspector;
        // changing the isolated supplied title to 10pt writes a delta of 4.
        (6.0 + self.rule_offset.1).max(0.0)
    }

    fn horizontal_border(&self, position: u64) -> Option<TextParagraphRule> {
        if self.border_positions? & position == 0 {
            return None;
        }
        let (Paint::Solid(color), stroke_width, _) = self.rule_stroke.as_ref()? else {
            return None;
        };
        Some(TextParagraphRule {
            color: *color,
            stroke_width: *stroke_width,
            offset_x: 0.0,
            offset_y: if position == 1 {
                -self.border_padding()
            } else {
                self.border_padding()
            },
            width: self.rule_width,
        })
    }

    fn rule_above(&self) -> Option<TextParagraphRule> {
        if self.border_positions.is_some() {
            return self.horizontal_border(1);
        }
        let (Paint::Solid(color), stroke_width, _) = self.rule_stroke.as_ref()? else {
            return None;
        };
        self.rule_enabled.then_some(TextParagraphRule {
            color: *color,
            stroke_width: *stroke_width,
            offset_x: self.rule_offset.0,
            offset_y: self.rule_offset.1,
            width: self.rule_width,
        })
    }
}

#[derive(Clone, Copy, Debug)]
struct KeynoteStyleChange {
    character_index: usize,
    identifier: Option<u64>,
}

#[derive(Clone, Copy, Debug)]
struct KeynoteParagraphDataChange {
    character_index: usize,
    list_level: usize,
}

#[derive(Default)]
struct KeynoteTextStorage {
    text: String,
    paragraph_styles: Vec<KeynoteStyleChange>,
    paragraph_data: Vec<KeynoteParagraphDataChange>,
    list_styles: Vec<KeynoteStyleChange>,
    character_styles: Vec<KeynoteStyleChange>,
    drop_cap_styles: Vec<KeynoteStyleChange>,
    attachments: Vec<KeynoteStyleChange>,
    sections: Vec<KeynoteStyleChange>,
}

#[derive(Clone, Copy, Debug)]
struct PagesDocumentLayout<'a> {
    frame_targets: &'a [(u64, PagesTargetHint)],
    font_metrics: &'a FontMetricTable,
    body_storage: u64,
    width: f32,
    height: f32,
    margin_left: f32,
    margin_right: f32,
    margin_top: f32,
    margin_bottom: f32,
    header_distance: f32,
    footer_distance: f32,
    view_scale: f32,
}

#[derive(Clone, Debug)]
struct PagesTargetHint {
    range_start: usize,
    range_end: usize,
    anchored_start: usize,
    anchored_end: usize,
    origin: Option<(f32, f32)>,
    size: Option<(f32, f32)>,
    column_count: u32,
}

#[derive(Clone, Debug)]
struct PagesPageHint {
    targets: Vec<PagesTargetHint>,
    flow_targets: Vec<(u64, PagesTargetHint)>,
    flow_diagnostics: Vec<Diagnostic>,
    attachment_positions: Vec<(f32, f32)>,
    toc_ranges: Vec<(usize, usize)>,
}

#[derive(Clone, Copy)]
struct PagesAttachment {
    drawable: u64,
    horizontal_offset_type: u64,
    horizontal_offset: f32,
}

#[derive(Clone, Debug)]
struct NumbersCell {
    row: u32,
    column: u32,
    text: Option<String>,
    text_storage: Option<u64>,
    span: bool,
    part: String,
    style_index: Option<u32>,
    text_style_index: Option<u32>,
}

#[derive(Clone, Debug)]
struct IworkCellText {
    text: String,
    storage: Option<u64>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct NumbersMerge {
    row: u32,
    start_column: u32,
    end_column: u32,
}

#[derive(Clone, Debug)]
struct KeynoteDrawableStyle {
    fill: Paint,
    image_fill: Option<(u64, bool, Option<(f32, f32)>)>,
    stroke: Paint,
    stroke_width: f32,
    stroke_style: StrokeStyle,
    opacity: f32,
    shadow: Option<Shadow>,
    reflection: Option<Reflection>,
}

impl Default for KeynoteDrawableStyle {
    fn default() -> Self {
        Self {
            fill: Paint::None,
            image_fill: None,
            stroke: Paint::None,
            stroke_width: 0.0,
            stroke_style: StrokeStyle::default(),
            opacity: 1.0,
            shadow: None,
            reflection: None,
        }
    }
}

#[derive(Clone, Debug)]
struct KeynoteImage {
    geometry: KeynoteGeometry,
    data_identifiers: Vec<u64>,
    style_identifier: Option<u64>,
    mask_identifier: Option<u64>,
}

#[derive(Clone, Debug)]
struct KeynoteMask {
    geometry: KeynoteGeometry,
    clip: Geometry,
    relative_to_image: bool,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum KeynoteChartKind {
    Column,
    Bar,
    Line,
    Area,
    Pie,
}

#[derive(Clone, Debug)]
struct KeynoteChart {
    geometry: KeynoteGeometry,
    kind: KeynoteChartKind,
    stacked: bool,
    title: Option<String>,
    series_names: Vec<String>,
    categories: Vec<String>,
    show_data_labels: Vec<bool>,
    show_legend: bool,
    background: Paint,
    background_stroke: (Paint, f32, StrokeStyle),
    plot_fill: Paint,
    bar_gaps: Option<(f32, f32)>,
    title_style: KeynoteTextStyle,
    value_axis: IworkChartAxis,
    category_axis: IworkChartAxis,
    colors: Vec<u32>,
    series: Vec<Vec<f32>>,
}

#[derive(Clone, Debug)]
struct IworkChartAxis {
    text: KeynoteTextStyle,
    show_axis: bool,
    show_labels: bool,
    grid: Option<(Paint, f32, StrokeStyle)>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum SharedDocumentKind {
    Numbers,
    Keynote,
}

/// Detects a modern, single-file iWork package from its contents and renders
/// its verified root preview. Filenames and caller-supplied MIME types are not
/// consulted.
pub fn detect_and_parse(bytes: &[u8], limits: Limits) -> Result<Option<Document>, Diagnostic> {
    detect_and_parse_with_font_metrics(bytes, limits, &FontMetricTable::default())
}

pub fn detect_and_parse_with_font_metrics(
    bytes: &[u8],
    limits: Limits,
    font_metrics: &FontMetricTable,
) -> Result<Option<Document>, Diagnostic> {
    if !bytes.starts_with(b"PK") {
        return Ok(None);
    }

    let package = Package::open_iwork(bytes, limits)?;
    let has_document = package.has_part(DOCUMENT_COMPONENT);
    let has_properties = package.has_part(PROPERTIES_PART);
    let has_iwa_component = package.entry_names().any(is_iwa_component);

    if !has_document {
        if has_properties && package.has_part("Index.zip") {
            return Err(directory_package_error("Index.zip"));
        }
        if let Some(prefix) = directory_package_prefix(&package) {
            return Err(directory_package_error(prefix));
        }
        if has_properties && has_iwa_component {
            return Err(format_error(
                DOCUMENT_COMPONENT,
                "iWork package is missing its root Document.iwa component",
            ));
        }
        return Ok(None);
    }
    if !has_properties {
        return Err(format_error(
            PROPERTIES_PART,
            "iWork package is missing Metadata/Properties.plist",
        ));
    }
    if !has_iwa_component {
        return Err(format_error(
            DOCUMENT_COMPONENT,
            "iWork package contains no Index/*.iwa components",
        ));
    }
    let properties = package.required_part(PROPERTIES_PART)?;
    validate_properties_plist(&properties, limits)?;

    let document_component = package.required_part(DOCUMENT_COMPONENT)?;
    let document_archives =
        parse_iwa_archives(&package, DOCUMENT_COMPONENT, &document_component, limits)?;
    let document_root = root_message(&document_archives, DOCUMENT_COMPONENT)?;
    if document_archives[0].identifier != 1 {
        return Err(iwa_error(
            DOCUMENT_COMPONENT,
            "iWork Document.iwa root ArchiveInfo identifier must be 1",
        ));
    }

    let keynote_marker = package
        .entry_names()
        .find(|name| is_keynote_component(name))
        .map(str::to_owned);
    if let Some(marker) = keynote_marker.as_deref() {
        let component = package.required_part(marker)?;
        let marker_archives = parse_iwa_archives(&package, marker, &component, limits)?;
        let marker_root = root_message(&marker_archives, marker)?;
        if marker_archives[0].identifier == 0 || !matches!(marker_root.message_type, 5 | 6) {
            return Err(format_error(
                marker,
                "iWork Keynote marker component has no native Slide identity",
            ));
        }
    }

    let (format, kind, unit_kind) = match document_root.message_type {
        NUMBERS_OR_KEYNOTE_DOCUMENT_TYPE => {
            let shared_kind = shared_document_kind(&document_root.payload)?;
            match (shared_kind, keynote_marker.is_some()) {
                (SharedDocumentKind::Keynote, _) => (
                    DocumentFormat::Keynote,
                    DocumentKind::Presentation,
                    UnitKind::Slide,
                ),
                (SharedDocumentKind::Numbers, false) => (
                    DocumentFormat::Numbers,
                    DocumentKind::Spreadsheet,
                    UnitKind::Sheet,
                ),
                (SharedDocumentKind::Numbers, true) => {
                    return Err(format_error(
                        DOCUMENT_COMPONENT,
                        "iWork package has contradictory Numbers and Keynote native identities",
                    ));
                }
            }
        }
        PAGES_DOCUMENT_TYPE if keynote_marker.is_none() => {
            (DocumentFormat::Pages, DocumentKind::Text, UnitKind::Page)
        }
        PAGES_DOCUMENT_TYPE => {
            return Err(format_error(
                DOCUMENT_COMPONENT,
                "iWork package has contradictory Pages and Keynote component markers",
            ));
        }
        message_type => {
            return Err(Diagnostic::fatal(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                None,
                format!(
                    "IWORK_MESSAGE_TYPE_UNSUPPORTED: unsupported iWork root message type {message_type}"
                ),
            )
            .in_part(DOCUMENT_COMPONENT));
        }
    };

    if let Some(marker) = [".iwpv2", ".iwph"]
        .into_iter()
        .find(|marker| package.has_part(marker))
    {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ZipEncrypted,
            Phase::Security,
            None,
            "IWORK_PASSWORD_PROTECTED_UNSUPPORTED: password-protected iWork documents are not supported",
        )
        .in_part(marker));
    }

    if format == DocumentFormat::Numbers {
        return numbers_document(&package, &document_archives, limits).map(Some);
    }

    let preview_part = ["preview.jpg", "preview.png"]
        .into_iter()
        .find(|part| package.has_part(part))
        .ok_or_else(|| {
            Diagnostic::fatal(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                None,
                "IWORK_PREVIEW_MISSING: this iWork package has no supported root preview.jpg or preview.png",
            )
        })?;
    let preview = package.required_part(preview_part)?;
    let (media_type, width, height) = preview_properties(preview_part, &preview, limits)?;

    let root_preview = IworkPreview {
        part: preview_part.to_owned(),
        media_type,
        width,
        height,
        bytes: preview.into_vec(),
    };
    if format == DocumentFormat::Pages
        && let Some(document) = pages_document(&package, &document_archives, limits, font_metrics)?
    {
        return Ok(Some(document));
    }
    if format == DocumentFormat::Keynote
        && let Some(document) =
            keynote_preview_document(&package, &document_archives, &root_preview, limits)?
    {
        return Ok(Some(document));
    }
    let source = SourceRef {
        part: preview_part.to_owned(),
        mapping: MappingQuality::Exact,
        locator: SourceLocator::Iwork {
            kind: "preview",
            component: preview_part.to_owned(),
        },
    };
    Ok(Some(Document {
        fatal: false,
        format: Some(format),
        kind: Some(kind),
        units: vec![Unit {
            kind: unit_kind,
            index: 0,
            id: "iwork-preview-unit".to_owned(),
            name: "Preview".to_owned(),
            width,
            height,
            rows: 0,
            columns: 0,
            frozen_rows: 0,
            frozen_columns: 0,
            frozen_width: 0.0,
            frozen_height: 0.0,
            row_axis: crate::model::SheetAxis::default(),
            column_axis: crate::model::SheetAxis::default(),
            show_grid_lines: unit_kind == UnitKind::Sheet,
            tab_color: None,
            sheet: None,
            slide: None,
        }],
        outline: Vec::new(),
        objects: vec![Object {
            numeric_id: 0,
            parent_numeric_id: None,
            stable_id: "iwork-preview".to_owned(),
            parent_stable_id: None,
            kind: ObjectKind::Image,
            unit_index: 0,
            bounds: Rect {
                x: 0.0,
                y: 0.0,
                width,
                height,
            },
            z: 0,
            text: None,
            source,
            visual: Visual::Image {
                media_type: media_type.to_owned(),
                bytes: root_preview.bytes,
                crop: ImageCrop::default(),
            },
        }],
        embedded_fonts: Vec::new(),
        font_alternate_names: Vec::new(),
        diagnostics: vec![
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Approximate,
                "IWORK_PREVIEW_ONLY: showing the embedded first-page, first-sheet, or first-slide preview; native text, additional units, search, and object inspection are unavailable",
            )
            .in_part(preview_part),
        ],
    }))
}

fn is_iwa_component(name: &str) -> bool {
    name.strip_prefix("Index/")
        .is_some_and(|relative| !relative.is_empty() && relative.ends_with(".iwa"))
}

fn is_keynote_component(name: &str) -> bool {
    let Some(relative) = name.strip_prefix("Index/") else {
        return false;
    };
    if relative.contains('/') || !relative.ends_with(".iwa") {
        return false;
    }
    let stem = &relative[..relative.len() - 4];
    ["Slide", "MasterSlide", "TemplateSlide"]
        .into_iter()
        .any(|prefix| keynote_component_stem(stem, prefix))
}

fn keynote_component_stem(stem: &str, prefix: &str) -> bool {
    if stem == prefix {
        return true;
    }
    let Some(suffix) = stem.strip_prefix(prefix) else {
        return false;
    };
    if !suffix.is_empty() && suffix.bytes().all(|byte| byte.is_ascii_digit()) {
        return true;
    }
    suffix.strip_prefix('-').is_some_and(|identifier| {
        identifier
            .split('-')
            .all(|segment| !segment.is_empty() && segment.bytes().all(|byte| byte.is_ascii_digit()))
    })
}

fn directory_package_prefix<'a>(package: &'a Package<'_>) -> Option<&'a str> {
    package.entry_names().find_map(|name| {
        let prefix = name
            .strip_suffix("/Index/Document.iwa")
            .or_else(|| name.strip_suffix("/Index.zip"))?;
        if prefix.is_empty() {
            return None;
        }
        let properties = format!("{prefix}/{PROPERTIES_PART}");
        package.has_part(&properties).then_some(prefix)
    })
}

fn directory_package_error(part: &str) -> Diagnostic {
    Diagnostic::fatal(
        DiagnosticCode::UnsupportedFeature,
        Phase::Container,
        None,
        "IWORK_DIRECTORY_PACKAGE_UNSUPPORTED: iWork directory packages are not supported; save the document as a single file",
    )
    .in_part(part)
}

fn validate_properties_plist(bytes: &[u8], limits: Limits) -> Result<(), Diagnostic> {
    if bytes.len() > limits.max_xml_bytes {
        return Err(Diagnostic::fatal(
            DiagnosticCode::XmlSizeLimit,
            Phase::Parse,
            None,
            "iWork metadata property list exceeds the configured byte limit",
        )
        .in_part(PROPERTIES_PART));
    }
    if bytes.starts_with(b"bplist00") {
        validate_binary_plist(bytes, limits)
    } else {
        validate_xml_plist(bytes, limits)
    }
}

#[derive(Debug)]
enum XmlPlistNodeKind {
    Plist { values: usize },
    Array,
    Dict { expects_key: bool },
    Key,
    String,
    Data,
    Date,
    Real,
    Integer,
    Boolean,
}

#[derive(Debug)]
struct XmlPlistNode {
    name: &'static str,
    kind: XmlPlistNodeKind,
    text: String,
}

fn validate_xml_plist(bytes: &[u8], limits: Limits) -> Result<(), Diagnostic> {
    let xml = strip_canonical_plist_doctype(bytes)?;
    let mut stack = Vec::<XmlPlistNode>::new();
    let mut root_seen = false;
    parse_xml(xml.as_ref(), limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let node = xml_plist_node(name)?;
                if stack.is_empty() {
                    if root_seen || name != "plist" {
                        return Err(plist_error("XML property list root must be plist"));
                    }
                    root_seen = true;
                    if attributes.len() != 1 || attributes[0].name != "version" {
                        return Err(plist_error(
                            "XML plist root must declare only version=\"1.0\"",
                        ));
                    }
                    let version = decode_xml_text(attributes[0].value)
                        .map_err(|error| error.in_part(PROPERTIES_PART))?;
                    if version.as_ref() != "1.0" {
                        return Err(plist_error("XML plist version must be 1.0"));
                    }
                } else {
                    if !attributes.is_empty() {
                        return Err(plist_error("XML plist values cannot carry attributes"));
                    }
                    xml_plist_accept_child(stack.last_mut().unwrap(), name)?;
                }
                if empty {
                    validate_completed_xml_plist_node(&node)?;
                } else {
                    stack.try_reserve(1).map_err(|_| {
                        Diagnostic::fatal(
                            DiagnosticCode::AllocationFailed,
                            Phase::Parse,
                            None,
                            "unable to grow the XML plist validation stack",
                        )
                        .in_part(PROPERTIES_PART)
                    })?;
                    stack.push(node);
                }
            }
            XmlEvent::EndElement { name } => {
                let node = stack
                    .pop()
                    .ok_or_else(|| plist_error("XML plist closing element has no open value"))?;
                if node.name != name {
                    return Err(plist_error("XML plist element stack is inconsistent"));
                }
                validate_completed_xml_plist_node(&node)?;
            }
            XmlEvent::Text(value) => {
                let value =
                    decode_xml_text(value).map_err(|error| error.in_part(PROPERTIES_PART))?;
                append_xml_plist_text(&mut stack, value.as_ref(), limits)?;
            }
            XmlEvent::Cdata(value) => append_xml_plist_text(&mut stack, value, limits)?,
        }
        Ok(())
    })
    .map_err(|error| {
        if error.location.part.is_some() {
            error
        } else {
            error.in_part(PROPERTIES_PART)
        }
    })?;
    if !root_seen || !stack.is_empty() {
        return Err(plist_error("XML property list is structurally incomplete"));
    }
    Ok(())
}

fn strip_canonical_plist_doctype(bytes: &[u8]) -> Result<Cow<'_, [u8]>, Diagnostic> {
    let root = find_subslice(bytes, b"<plist").ok_or_else(|| {
        plist_error("iWork metadata is not a recognizable binary or XML property list")
    })?;
    let prefix = &bytes[..root];
    let Some(doctype) = find_subslice(prefix, b"<!DOCTYPE") else {
        return Ok(Cow::Borrowed(bytes));
    };
    if !prefix[doctype..].starts_with(CANONICAL_PLIST_DOCTYPE) {
        return Err(plist_error(
            "XML plist contains a non-canonical or unsafe document type declaration",
        ));
    }
    let end = doctype
        .checked_add(CANONICAL_PLIST_DOCTYPE.len())
        .ok_or_else(|| plist_error("XML plist document type offset overflows"))?;
    if find_subslice(&prefix[end..], b"<!DOCTYPE").is_some() {
        return Err(plist_error(
            "XML plist contains multiple document type declarations",
        ));
    }
    let mut normalized = Vec::new();
    normalized
        .try_reserve_exact(bytes.len() - CANONICAL_PLIST_DOCTYPE.len())
        .map_err(|_| {
            Diagnostic::fatal(
                DiagnosticCode::AllocationFailed,
                Phase::Parse,
                None,
                "unable to normalize the bounded XML plist declaration",
            )
            .in_part(PROPERTIES_PART)
        })?;
    normalized.extend_from_slice(&bytes[..doctype]);
    normalized.extend_from_slice(&bytes[end..]);
    Ok(Cow::Owned(normalized))
}

fn xml_plist_node(name: &str) -> Result<XmlPlistNode, Diagnostic> {
    let (name, kind) = match name {
        "plist" => ("plist", XmlPlistNodeKind::Plist { values: 0 }),
        "array" => ("array", XmlPlistNodeKind::Array),
        "dict" => ("dict", XmlPlistNodeKind::Dict { expects_key: true }),
        "key" => ("key", XmlPlistNodeKind::Key),
        "string" => ("string", XmlPlistNodeKind::String),
        "data" => ("data", XmlPlistNodeKind::Data),
        "date" => ("date", XmlPlistNodeKind::Date),
        "real" => ("real", XmlPlistNodeKind::Real),
        "integer" => ("integer", XmlPlistNodeKind::Integer),
        "true" => ("true", XmlPlistNodeKind::Boolean),
        "false" => ("false", XmlPlistNodeKind::Boolean),
        _ => return Err(plist_error("XML plist contains an unsupported element")),
    };
    Ok(XmlPlistNode {
        name,
        kind,
        text: String::new(),
    })
}

fn is_xml_plist_value(name: &str) -> bool {
    matches!(
        name,
        "array" | "dict" | "string" | "data" | "date" | "real" | "integer" | "true" | "false"
    )
}

fn xml_plist_accept_child(parent: &mut XmlPlistNode, child: &str) -> Result<(), Diagnostic> {
    match &mut parent.kind {
        XmlPlistNodeKind::Plist { values } => {
            if !is_xml_plist_value(child) || *values != 0 {
                return Err(plist_error("XML plist must contain exactly one root value"));
            }
            *values = 1;
        }
        XmlPlistNodeKind::Array => {
            if !is_xml_plist_value(child) {
                return Err(plist_error("XML plist array contains a non-value element"));
            }
        }
        XmlPlistNodeKind::Dict { expects_key } => {
            if *expects_key {
                if child != "key" {
                    return Err(plist_error(
                        "XML plist dictionary value has no preceding key",
                    ));
                }
                *expects_key = false;
            } else {
                if !is_xml_plist_value(child) {
                    return Err(plist_error("XML plist dictionary key has no value"));
                }
                *expects_key = true;
            }
        }
        _ => return Err(plist_error("XML plist scalar contains a child element")),
    }
    Ok(())
}

fn append_xml_plist_text(
    stack: &mut [XmlPlistNode],
    value: &str,
    limits: Limits,
) -> Result<(), Diagnostic> {
    let Some(node) = stack.last_mut() else {
        if value.trim().is_empty() {
            return Ok(());
        }
        return Err(plist_error("XML plist has text outside its root value"));
    };
    match node.kind {
        XmlPlistNodeKind::Key
        | XmlPlistNodeKind::String
        | XmlPlistNodeKind::Data
        | XmlPlistNodeKind::Date
        | XmlPlistNodeKind::Real
        | XmlPlistNodeKind::Integer => {
            let next = node
                .text
                .len()
                .checked_add(value.len())
                .filter(|length| *length <= limits.max_xml_bytes)
                .ok_or_else(|| plist_error("XML plist scalar text exceeds the byte limit"))?;
            node.text.try_reserve(next - node.text.len()).map_err(|_| {
                Diagnostic::fatal(
                    DiagnosticCode::AllocationFailed,
                    Phase::Parse,
                    None,
                    "unable to grow XML plist scalar text",
                )
                .in_part(PROPERTIES_PART)
            })?;
            node.text.push_str(value);
        }
        _ if value.trim().is_empty() => {}
        _ => {
            return Err(plist_error(
                "XML plist container contains non-whitespace text",
            ));
        }
    }
    Ok(())
}

fn validate_completed_xml_plist_node(node: &XmlPlistNode) -> Result<(), Diagnostic> {
    match &node.kind {
        XmlPlistNodeKind::Plist { values } if *values != 1 => {
            Err(plist_error("XML plist must contain exactly one root value"))
        }
        XmlPlistNodeKind::Dict { expects_key } if !expects_key => Err(plist_error(
            "XML plist dictionary ends after a key without a value",
        )),
        XmlPlistNodeKind::Integer if node.text.trim().parse::<i128>().is_err() => {
            Err(plist_error("XML plist integer is invalid"))
        }
        XmlPlistNodeKind::Real
            if node
                .text
                .trim()
                .parse::<f64>()
                .ok()
                .is_none_or(|value| !value.is_finite()) =>
        {
            Err(plist_error("XML plist real value is invalid"))
        }
        XmlPlistNodeKind::Date if !valid_xml_plist_date(node.text.trim()) => {
            Err(plist_error("XML plist date is invalid"))
        }
        XmlPlistNodeKind::Data if !valid_xml_plist_base64(&node.text) => {
            Err(plist_error("XML plist data is not canonical base64"))
        }
        _ => Ok(()),
    }
}

fn valid_xml_plist_date(value: &str) -> bool {
    let bytes = value.as_bytes();
    if bytes.len() < 20
        || bytes[4] != b'-'
        || bytes[7] != b'-'
        || bytes[10] != b'T'
        || bytes[13] != b':'
        || bytes[16] != b':'
        || bytes.last() != Some(&b'Z')
    {
        return false;
    }
    if !bytes[..4].iter().all(u8::is_ascii_digit)
        || !bytes[5..7].iter().all(u8::is_ascii_digit)
        || !bytes[8..10].iter().all(u8::is_ascii_digit)
        || !bytes[11..13].iter().all(u8::is_ascii_digit)
        || !bytes[14..16].iter().all(u8::is_ascii_digit)
        || !bytes[17..19].iter().all(u8::is_ascii_digit)
    {
        return false;
    }
    let fraction = &bytes[19..bytes.len() - 1];
    if !fraction.is_empty()
        && (fraction[0] != b'.'
            || fraction.len() == 1
            || !fraction[1..].iter().all(u8::is_ascii_digit))
    {
        return false;
    }
    let parse = |range: std::ops::Range<usize>| {
        std::str::from_utf8(&bytes[range])
            .ok()
            .and_then(|value| value.parse::<u32>().ok())
    };
    matches!(parse(5..7), Some(1..=12))
        && matches!(parse(8..10), Some(1..=31))
        && matches!(parse(11..13), Some(0..=23))
        && matches!(parse(14..16), Some(0..=59))
        && matches!(parse(17..19), Some(0..=59))
}

fn valid_xml_plist_base64(value: &str) -> bool {
    let mut significant = 0_usize;
    let mut padding = 0_usize;
    let mut saw_padding = false;
    for byte in value.bytes() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        significant += 1;
        match byte {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'+' | b'/' if !saw_padding => {}
            b'=' if padding < 2 => {
                saw_padding = true;
                padding += 1;
            }
            _ => return false,
        }
    }
    significant.is_multiple_of(4) && (significant != 0 || padding == 0)
}

fn validate_binary_plist(bytes: &[u8], limits: Limits) -> Result<(), Diagnostic> {
    let minimum = 8_usize
        .checked_add(1)
        .and_then(|value| value.checked_add(BINARY_PLIST_TRAILER_BYTES))
        .unwrap();
    if bytes.len() < minimum {
        return Err(plist_error("binary plist is truncated"));
    }
    let trailer_start = bytes.len() - BINARY_PLIST_TRAILER_BYTES;
    let trailer = &bytes[trailer_start..];
    if trailer[..6].iter().any(|byte| *byte != 0) {
        return Err(plist_error(
            "binary plist trailer reserved bytes are nonzero",
        ));
    }
    let offset_width = usize::from(trailer[6]);
    let reference_width = usize::from(trailer[7]);
    if !(1..=8).contains(&offset_width) || !(1..=8).contains(&reference_width) {
        return Err(plist_error(
            "binary plist uses an invalid offset or reference width",
        ));
    }
    let object_count = usize::try_from(read_be_u64(trailer, 8, 8)?)
        .map_err(|_| plist_error("binary plist object count is not addressable"))?;
    let top_object = usize::try_from(read_be_u64(trailer, 16, 8)?)
        .map_err(|_| plist_error("binary plist top object is not addressable"))?;
    let offset_table = usize::try_from(read_be_u64(trailer, 24, 8)?)
        .map_err(|_| plist_error("binary plist offset table is not addressable"))?;
    if object_count == 0 || object_count > limits.max_xml_nodes || top_object >= object_count {
        return Err(plist_error("binary plist object table bounds are invalid"));
    }
    if offset_table < 9 || offset_table >= trailer_start {
        return Err(plist_error("binary plist offset table position is invalid"));
    }
    if !binary_width_can_hold(offset_width, offset_table - 1)
        || !binary_width_can_hold(reference_width, object_count - 1)
    {
        return Err(plist_error(
            "binary plist width cannot represent its table bounds",
        ));
    }
    let table_bytes = object_count
        .checked_mul(offset_width)
        .ok_or_else(|| plist_error("binary plist offset table size overflows"))?;
    if offset_table
        .checked_add(table_bytes)
        .is_none_or(|end| end != trailer_start)
    {
        return Err(plist_error(
            "binary plist offset table does not end at its trailer",
        ));
    }

    let mut offsets = Vec::new();
    offsets.try_reserve_exact(object_count).map_err(|_| {
        Diagnostic::fatal(
            DiagnosticCode::AllocationFailed,
            Phase::Parse,
            None,
            "unable to allocate the bounded binary plist offset table",
        )
        .in_part(PROPERTIES_PART)
    })?;
    for index in 0..object_count {
        let position = offset_table + index * offset_width;
        let offset = read_be_usize(bytes, position, offset_width)?;
        if !(8..offset_table).contains(&offset) {
            return Err(plist_error(
                "binary plist object offset is outside the object table",
            ));
        }
        offsets.push(offset);
    }
    let mut ordered = offsets
        .iter()
        .copied()
        .enumerate()
        .map(|(identifier, offset)| (offset, identifier))
        .collect::<Vec<_>>();
    ordered.sort_unstable_by_key(|entry| entry.0);
    if ordered.windows(2).any(|pair| pair[0].0 == pair[1].0) {
        return Err(plist_error("binary plist object offsets are duplicated"));
    }
    let mut boundaries = vec![offset_table; object_count];
    for pair in ordered.windows(2) {
        boundaries[pair[0].1] = pair[1].0;
    }
    if bytes[offsets[top_object]] >> 4 != 0x0d {
        return Err(plist_error(
            "iWork metadata binary plist root must be a dictionary",
        ));
    }

    let mut states = Vec::new();
    states.try_reserve_exact(object_count).map_err(|_| {
        Diagnostic::fatal(
            DiagnosticCode::AllocationFailed,
            Phase::Parse,
            None,
            "unable to allocate the bounded binary plist visit map",
        )
        .in_part(PROPERTIES_PART)
    })?;
    states.resize(object_count, 0_u8);
    let mut stack = vec![(top_object, 1_usize, false)];
    let mut edges = 0_usize;
    while let Some((identifier, depth, exiting)) = stack.pop() {
        if exiting {
            states[identifier] = 2;
            continue;
        }
        match states[identifier] {
            2 => continue,
            1 => return Err(plist_error("binary plist object graph contains a cycle")),
            _ => {}
        }
        if depth > limits.max_xml_depth {
            return Err(plist_error(
                "binary plist object graph exceeds the depth limit",
            ));
        }
        states[identifier] = 1;
        let references = binary_plist_references(
            bytes,
            offsets[identifier],
            boundaries[identifier],
            reference_width,
            &offsets,
            limits,
        )?;
        edges = edges
            .checked_add(references.len())
            .filter(|value| *value <= limits.max_relationship_edges)
            .ok_or_else(|| plist_error("binary plist object graph exceeds the edge limit"))?;
        stack.try_reserve(references.len() + 1).map_err(|_| {
            Diagnostic::fatal(
                DiagnosticCode::AllocationFailed,
                Phase::Parse,
                None,
                "unable to grow the bounded binary plist traversal stack",
            )
            .in_part(PROPERTIES_PART)
        })?;
        stack.push((identifier, depth, true));
        for reference in references.into_iter().rev() {
            stack.push((reference, depth + 1, false));
        }
    }
    Ok(())
}

fn binary_plist_references(
    bytes: &[u8],
    offset: usize,
    boundary: usize,
    reference_width: usize,
    object_offsets: &[usize],
    limits: Limits,
) -> Result<Vec<usize>, Diagnostic> {
    let object_count = object_offsets.len();
    let marker = *bytes
        .get(offset)
        .ok_or_else(|| plist_error("binary plist object marker is truncated"))?;
    let class = marker >> 4;
    let info = marker & 0x0f;
    let cursor = offset + 1;
    match class {
        0x0 if matches!(marker, 0x00 | 0x08 | 0x09) => {
            ensure_binary_object_end(cursor, 0, boundary)?;
            Ok(Vec::new())
        }
        0x1 => {
            let size = binary_power_of_two_size(info, 4, "integer")?;
            ensure_binary_object_end(cursor, size, boundary)?;
            Ok(Vec::new())
        }
        0x2 => {
            let size = binary_power_of_two_size(info, 3, "real")?;
            if !matches!(size, 4 | 8) {
                return Err(plist_error("binary plist real width is invalid"));
            }
            ensure_binary_object_end(cursor, size, boundary)?;
            Ok(Vec::new())
        }
        0x3 if marker == 0x33 => {
            ensure_binary_object_end(cursor, 8, boundary)?;
            Ok(Vec::new())
        }
        0x4..=0x6 => {
            let (count, data_start) = binary_plist_count(bytes, cursor, info, boundary)?;
            let unit = if class == 0x6 { 2 } else { 1 };
            let size = count
                .checked_mul(unit)
                .ok_or_else(|| plist_error("binary plist scalar size overflows"))?;
            let end = ensure_binary_object_end(data_start, size, boundary)?;
            if class == 0x5 && bytes[data_start..end].iter().any(|byte| !byte.is_ascii()) {
                return Err(plist_error(
                    "binary plist ASCII string contains a non-ASCII byte",
                ));
            }
            if class == 0x6 {
                let units = bytes[data_start..end]
                    .chunks_exact(2)
                    .map(|pair| u16::from_be_bytes([pair[0], pair[1]]));
                if char::decode_utf16(units).any(|value| value.is_err()) {
                    return Err(plist_error("binary plist string contains invalid UTF-16"));
                }
            }
            Ok(Vec::new())
        }
        0x8 => {
            ensure_binary_object_end(cursor, usize::from(info) + 1, boundary)?;
            Ok(Vec::new())
        }
        0x0a..=0x0d => {
            let (count, references_start) = binary_plist_count(bytes, cursor, info, boundary)?;
            if count > limits.max_relationship_edges {
                return Err(plist_error(
                    "binary plist collection exceeds the edge limit",
                ));
            }
            let reference_count = if class == 0x0d {
                count
                    .checked_mul(2)
                    .ok_or_else(|| plist_error("binary plist dictionary size overflows"))?
            } else {
                count
            };
            let reference_bytes = reference_count
                .checked_mul(reference_width)
                .ok_or_else(|| plist_error("binary plist reference table size overflows"))?;
            ensure_binary_object_end(references_start, reference_bytes, boundary)?;
            let mut references = Vec::new();
            references.try_reserve_exact(reference_count).map_err(|_| {
                Diagnostic::fatal(
                    DiagnosticCode::AllocationFailed,
                    Phase::Parse,
                    None,
                    "unable to allocate binary plist references",
                )
                .in_part(PROPERTIES_PART)
            })?;
            for index in 0..reference_count {
                let position = references_start + index * reference_width;
                let reference = read_be_usize(bytes, position, reference_width)?;
                if reference >= object_count {
                    return Err(plist_error(
                        "binary plist reference exceeds the object table",
                    ));
                }
                if class == 0x0d && index < count {
                    let key_class = bytes[object_offsets[reference]] >> 4;
                    if !matches!(key_class, 0x05 | 0x06) {
                        return Err(plist_error("binary plist dictionary key is not a string"));
                    }
                }
                references.push(reference);
            }
            Ok(references)
        }
        _ => Err(plist_error(
            "binary plist contains an unsupported object marker",
        )),
    }
}

fn binary_plist_count(
    bytes: &[u8],
    cursor: usize,
    info: u8,
    boundary: usize,
) -> Result<(usize, usize), Diagnostic> {
    if info < 0x0f {
        return Ok((usize::from(info), cursor));
    }
    let marker = *bytes
        .get(cursor)
        .filter(|_| cursor < boundary)
        .ok_or_else(|| plist_error("binary plist extended count is truncated"))?;
    if marker >> 4 != 0x01 {
        return Err(plist_error("binary plist extended count is not an integer"));
    }
    let width = binary_power_of_two_size(marker & 0x0f, 3, "count")?;
    let value_start = cursor + 1;
    ensure_binary_object_end(value_start, width, boundary)?;
    let value = read_be_usize(bytes, value_start, width)?;
    Ok((value, value_start + width))
}

fn binary_power_of_two_size(
    info: u8,
    maximum_exponent: u8,
    kind: &str,
) -> Result<usize, Diagnostic> {
    if info > maximum_exponent {
        return Err(plist_error(format!("binary plist {kind} width is invalid")));
    }
    1_usize
        .checked_shl(u32::from(info))
        .ok_or_else(|| plist_error(format!("binary plist {kind} width overflows")))
}

fn ensure_binary_object_end(
    start: usize,
    length: usize,
    boundary: usize,
) -> Result<usize, Diagnostic> {
    start
        .checked_add(length)
        .filter(|end| *end <= boundary)
        .ok_or_else(|| plist_error("binary plist object overlaps the next object or offset table"))
}

fn binary_width_can_hold(width: usize, value: usize) -> bool {
    width >= 8 || value < (1_usize << (width * 8))
}

fn read_be_u64(bytes: &[u8], offset: usize, width: usize) -> Result<u64, Diagnostic> {
    let end = offset
        .checked_add(width)
        .ok_or_else(|| plist_error("binary plist integer offset overflows"))?;
    let slice = bytes
        .get(offset..end)
        .ok_or_else(|| plist_error("binary plist integer is truncated"))?;
    let mut value = 0_u64;
    for byte in slice {
        value = (value << 8) | u64::from(*byte);
    }
    Ok(value)
}

fn read_be_usize(bytes: &[u8], offset: usize, width: usize) -> Result<usize, Diagnostic> {
    usize::try_from(read_be_u64(bytes, offset, width)?)
        .map_err(|_| plist_error("binary plist integer is not addressable"))
}

fn find_subslice(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

fn plist_error(message: impl Into<String>) -> Diagnostic {
    format_error(PROPERTIES_PART, message)
}

fn parse_iwa_archives(
    package: &Package<'_>,
    part: &str,
    bytes: &[u8],
    limits: Limits,
) -> Result<Vec<IwaArchive>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut expanded = Vec::new();
    while cursor < bytes.len() {
        let header_end = cursor
            .checked_add(4)
            .ok_or_else(|| iwa_error(part, "IWA chunk header offset overflow"))?;
        let Some(header) = bytes.get(cursor..header_end) else {
            return Err(iwa_error(part, "IWA chunk header is truncated"));
        };
        if header[0] != 0 {
            return Err(iwa_error(
                part,
                "IWA chunk type must be the native Snappy data type 0x00",
            ));
        }
        let compressed_length =
            usize::from(header[1]) | (usize::from(header[2]) << 8) | (usize::from(header[3]) << 16);
        if compressed_length == 0 {
            return Err(iwa_error(part, "IWA contains an empty Snappy chunk"));
        }
        let chunk_end = header_end
            .checked_add(compressed_length)
            .ok_or_else(|| iwa_error(part, "IWA compressed chunk length overflow"))?;
        let Some(chunk) = bytes.get(header_end..chunk_end) else {
            return Err(iwa_error(part, "IWA compressed chunk is truncated"));
        };
        let (declared_length, _) = read_varint(chunk)
            .map_err(|message| iwa_error(part, format!("invalid IWA Snappy length: {message}")))?;
        let declared_length = usize::try_from(declared_length)
            .map_err(|_| iwa_error(part, "IWA Snappy output length is not addressable"))?;
        let operations = snappy_operation_cost(chunk, declared_length, part)?;
        let (expanded_bytes_left, operations_left) = package.remaining_iwork_budget();
        if declared_length > expanded_bytes_left {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ZipTotalSizeLimit,
                Phase::Parse,
                None,
                "IWA cumulative Snappy output exhausts the shared document expanded-byte budget",
            )
            .in_part(part));
        }
        if operations > operations_left {
            return Err(snappy_operation_limit(part));
        }
        if !package.consume_iwork_budget(declared_length, operations) {
            return Err(iwa_error(
                part,
                "IWA shared document resource budget changed unexpectedly",
            ));
        }
        let chunk = decompress_snappy(chunk, limits, part)?;
        let combined_length = expanded
            .len()
            .checked_add(chunk.len())
            .filter(|length| *length <= limits.max_total_uncompressed_bytes)
            .ok_or_else(|| {
                Diagnostic::fatal(
                    DiagnosticCode::ZipTotalSizeLimit,
                    Phase::Parse,
                    None,
                    "IWA component exceeds the configured expanded-byte limit",
                )
                .in_part(part)
            })?;
        expanded
            .try_reserve_exact(combined_length - expanded.len())
            .map_err(|_| {
                Diagnostic::fatal(
                    DiagnosticCode::AllocationFailed,
                    Phase::Parse,
                    None,
                    "unable to allocate the bounded IWA component stream",
                )
                .in_part(part)
            })?;
        expanded.extend_from_slice(&chunk);
        cursor = chunk_end;
    }
    if expanded.is_empty() {
        return Err(iwa_error(part, "IWA component contains no chunks"));
    }
    parse_iwa_archive_stream(&expanded, part, limits)
}

fn parse_iwa_archive_stream(
    bytes: &[u8],
    part: &str,
    limits: Limits,
) -> Result<Vec<IwaArchive>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut archives = Vec::new();
    while cursor < bytes.len() {
        if archives.len() >= limits.max_relationship_edges {
            return Err(iwa_error(part, "IWA component contains too many archives"));
        }
        let (archive_length, prefix_length) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid IWA ArchiveInfo length: {message}"))
        })?;
        cursor += prefix_length;
        let archive_length = usize::try_from(archive_length)
            .map_err(|_| iwa_error(part, "IWA ArchiveInfo length is not addressable"))?;
        let archive_end = cursor
            .checked_add(archive_length)
            .ok_or_else(|| iwa_error(part, "IWA ArchiveInfo length overflow"))?;
        let archive_info = bytes
            .get(cursor..archive_end)
            .ok_or_else(|| iwa_error(part, "IWA ArchiveInfo is truncated"))?;
        let (identifier, message_infos) = parse_archive_info(archive_info, part)?;
        if archives.is_empty() && message_infos.len() != 1 {
            return Err(iwa_error(
                part,
                "IWA root ArchiveInfo must contain exactly one MessageInfo",
            ));
        }
        cursor = archive_end;

        let mut messages = Vec::new();
        messages
            .try_reserve_exact(message_infos.len())
            .map_err(|_| {
                Diagnostic::fatal(
                    DiagnosticCode::AllocationFailed,
                    Phase::Parse,
                    None,
                    "unable to allocate the bounded IWA message list",
                )
                .in_part(part)
            })?;
        for info in message_infos {
            let payload_end = cursor
                .checked_add(info.payload_length)
                .ok_or_else(|| iwa_error(part, "IWA message payload length overflow"))?;
            let payload = bytes
                .get(cursor..payload_end)
                .ok_or_else(|| iwa_error(part, "IWA message payload is truncated"))?;
            let mut owned_payload = Vec::new();
            owned_payload
                .try_reserve_exact(payload.len())
                .map_err(|_| {
                    Diagnostic::fatal(
                        DiagnosticCode::AllocationFailed,
                        Phase::Parse,
                        None,
                        "unable to allocate the bounded IWA message payload",
                    )
                    .in_part(part)
                })?;
            owned_payload.extend_from_slice(payload);
            messages.push(IwaMessage {
                message_type: info.message_type,
                payload: owned_payload,
                data_references: info.data_references,
            });
            cursor = payload_end;
        }
        if messages.is_empty() {
            return Err(iwa_error(part, "IWA ArchiveInfo contains no MessageInfo"));
        }
        archives.push(IwaArchive {
            identifier,
            messages,
        });
    }
    Ok(archives)
}

fn root_message<'a>(archives: &'a [IwaArchive], part: &str) -> Result<&'a IwaMessage, Diagnostic> {
    let root = archives
        .first()
        .ok_or_else(|| iwa_error(part, "IWA component contains no archives"))?;
    if root.messages.len() != 1 {
        return Err(iwa_error(
            part,
            "IWA root ArchiveInfo must contain exactly one MessageInfo",
        ));
    }
    Ok(&root.messages[0])
}

fn decompress_snappy(bytes: &[u8], limits: Limits, part: &str) -> Result<Vec<u8>, Diagnostic> {
    let (expected_length, mut cursor) = read_varint(bytes)
        .map_err(|message| iwa_error(part, format!("invalid Snappy length: {message}")))?;
    let expected_length = usize::try_from(expected_length)
        .map_err(|_| iwa_error(part, "Snappy output length is not addressable"))?;
    if expected_length > limits.max_entry_uncompressed_bytes {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ZipEntryTooLarge,
            Phase::Parse,
            None,
            "IWA Snappy chunk exceeds the configured expanded-byte limit",
        )
        .in_part(part));
    }
    let mut output = Vec::new();
    output.try_reserve_exact(expected_length).map_err(|_| {
        Diagnostic::fatal(
            DiagnosticCode::AllocationFailed,
            Phase::Parse,
            None,
            "unable to allocate the bounded IWA Snappy output buffer",
        )
        .in_part(part)
    })?;

    while cursor < bytes.len() {
        let tag = bytes[cursor];
        cursor += 1;
        match tag & 0x03 {
            0 => {
                let code = usize::from(tag >> 2);
                let literal_length = if code < 60 {
                    code + 1
                } else {
                    let length_bytes = code - 59;
                    let end = cursor
                        .checked_add(length_bytes)
                        .ok_or_else(|| iwa_error(part, "Snappy literal length offset overflow"))?;
                    let Some(encoded) = bytes.get(cursor..end) else {
                        return Err(iwa_error(part, "Snappy literal length is truncated"));
                    };
                    cursor = end;
                    let mut length = 0_usize;
                    for (index, byte) in encoded.iter().copied().enumerate() {
                        length |= usize::from(byte) << (index * 8);
                    }
                    length
                        .checked_add(1)
                        .ok_or_else(|| iwa_error(part, "Snappy literal length overflow"))?
                };
                let end = cursor
                    .checked_add(literal_length)
                    .ok_or_else(|| iwa_error(part, "Snappy literal offset overflow"))?;
                let Some(literal) = bytes.get(cursor..end) else {
                    return Err(iwa_error(part, "Snappy literal is truncated"));
                };
                ensure_output_capacity(output.len(), literal_length, expected_length, part)?;
                output.extend_from_slice(literal);
                cursor = end;
            }
            1 => {
                let Some(&low_offset) = bytes.get(cursor) else {
                    return Err(iwa_error(part, "Snappy one-byte copy is truncated"));
                };
                cursor += 1;
                let length = 4 + usize::from((tag >> 2) & 0x07);
                let offset = ((usize::from(tag) & 0xe0) << 3) | usize::from(low_offset);
                copy_snappy(&mut output, offset, length, expected_length, part)?;
            }
            2 => {
                let end = cursor
                    .checked_add(2)
                    .ok_or_else(|| iwa_error(part, "Snappy two-byte copy offset overflow"))?;
                let Some(offset) = bytes.get(cursor..end) else {
                    return Err(iwa_error(part, "Snappy two-byte copy is truncated"));
                };
                cursor = end;
                copy_snappy(
                    &mut output,
                    usize::from(u16::from_le_bytes([offset[0], offset[1]])),
                    1 + usize::from(tag >> 2),
                    expected_length,
                    part,
                )?;
            }
            3 => {
                let end = cursor
                    .checked_add(4)
                    .ok_or_else(|| iwa_error(part, "Snappy four-byte copy offset overflow"))?;
                let Some(offset) = bytes.get(cursor..end) else {
                    return Err(iwa_error(part, "Snappy four-byte copy is truncated"));
                };
                cursor = end;
                let offset = usize::try_from(u32::from_le_bytes([
                    offset[0], offset[1], offset[2], offset[3],
                ]))
                .map_err(|_| iwa_error(part, "Snappy copy offset is not addressable"))?;
                copy_snappy(
                    &mut output,
                    offset,
                    1 + usize::from(tag >> 2),
                    expected_length,
                    part,
                )?;
            }
            _ => unreachable!(),
        }
    }
    if output.len() != expected_length {
        return Err(iwa_error(
            part,
            "Snappy output length does not match its declaration",
        ));
    }
    Ok(output)
}

fn snappy_operation_cost(
    bytes: &[u8],
    expected_length: usize,
    part: &str,
) -> Result<usize, Diagnostic> {
    bytes
        .len()
        .checked_add(expected_length)
        .ok_or_else(|| snappy_operation_limit(part))
}

fn snappy_operation_limit(part: &str) -> Diagnostic {
    Diagnostic::fatal(
        DiagnosticCode::ZipDeflateInvalid,
        Phase::Parse,
        None,
        "IWA Snappy operation budget exhausted",
    )
    .in_part(part)
}

fn copy_snappy(
    output: &mut Vec<u8>,
    offset: usize,
    length: usize,
    expected_length: usize,
    part: &str,
) -> Result<(), Diagnostic> {
    if offset == 0 || offset > output.len() {
        return Err(iwa_error(
            part,
            "Snappy copy offset is outside prior output",
        ));
    }
    ensure_output_capacity(output.len(), length, expected_length, part)?;
    for _ in 0..length {
        let byte = output[output.len() - offset];
        output.push(byte);
    }
    Ok(())
}

fn ensure_output_capacity(
    current: usize,
    additional: usize,
    expected: usize,
    part: &str,
) -> Result<(), Diagnostic> {
    if current
        .checked_add(additional)
        .is_none_or(|length| length > expected)
    {
        return Err(iwa_error(
            part,
            "Snappy command exceeds the declared output length",
        ));
    }
    Ok(())
}

fn parse_archive_info(bytes: &[u8], part: &str) -> Result<(u64, Vec<IwaMessageInfo>), Diagnostic> {
    let mut cursor = 0_usize;
    let mut identifier = None;
    let mut message_infos = Vec::new();
    let mut field_count = 0_usize;
    while cursor < bytes.len() {
        field_count += 1;
        if field_count > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(part, "IWA ArchiveInfo contains too many fields"));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid ArchiveInfo field key: {message}"))
        })?;
        cursor += consumed;
        let field = key >> 3;
        let wire = key & 7;
        if field == 0 {
            return Err(iwa_error(
                part,
                "IWA ArchiveInfo contains protobuf field zero",
            ));
        }
        match (field, wire) {
            (1, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid ArchiveInfo identifier: {message}"))
                })?;
                if identifier.replace(value).is_some() {
                    return Err(iwa_error(part, "IWA ArchiveInfo repeats its identifier"));
                }
                cursor += consumed;
            }
            (2, 2) => {
                let (length, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid MessageInfo length: {message}"))
                })?;
                cursor += consumed;
                let length = usize::try_from(length)
                    .map_err(|_| iwa_error(part, "MessageInfo length is not addressable"))?;
                let end = cursor
                    .checked_add(length)
                    .ok_or_else(|| iwa_error(part, "MessageInfo length overflow"))?;
                let Some(info) = bytes.get(cursor..end) else {
                    return Err(iwa_error(part, "MessageInfo is truncated"));
                };
                message_infos.push(parse_message_info(info, part)?);
                cursor = end;
            }
            _ => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    let identifier =
        identifier.ok_or_else(|| iwa_error(part, "IWA ArchiveInfo has no identifier"))?;
    if message_infos.is_empty() {
        return Err(iwa_error(part, "IWA ArchiveInfo has no MessageInfo"));
    }
    Ok((identifier, message_infos))
}

fn parse_message_info(bytes: &[u8], part: &str) -> Result<IwaMessageInfo, Diagnostic> {
    let mut cursor = 0_usize;
    let mut message_type = None;
    let mut payload_length = None;
    let mut data_references = Vec::new();
    let mut field_count = 0_usize;
    while cursor < bytes.len() {
        field_count += 1;
        if field_count > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(part, "IWA MessageInfo contains too many fields"));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid MessageInfo field key: {message}"))
        })?;
        cursor += consumed;
        let field = key >> 3;
        let wire = key & 7;
        if field == 0 {
            return Err(iwa_error(
                part,
                "IWA MessageInfo contains protobuf field zero",
            ));
        }
        match (field, wire) {
            (1, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid iWork message type: {message}"))
                })?;
                if message_type.replace(value).is_some() {
                    return Err(iwa_error(part, "IWA MessageInfo repeats its message type"));
                }
                cursor += consumed;
            }
            (3, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid iWork message length: {message}"))
                })?;
                let value = usize::try_from(value)
                    .map_err(|_| iwa_error(part, "iWork message length is not addressable"))?;
                if payload_length.replace(value).is_some() {
                    return Err(iwa_error(
                        part,
                        "IWA MessageInfo repeats its payload length",
                    ));
                }
                cursor += consumed;
            }
            (6, 2) => {
                let packed =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "IWA data reference list")?;
                let mut packed_cursor = 0_usize;
                while packed_cursor < packed.len() {
                    let (identifier, consumed) =
                        read_varint(&packed[packed_cursor..]).map_err(|message| {
                            iwa_error(part, format!("invalid IWA data reference: {message}"))
                        })?;
                    if identifier == 0 {
                        return Err(iwa_error(part, "IWA data reference identifier is zero"));
                    }
                    data_references.push(identifier);
                    packed_cursor += consumed;
                }
            }
            (6, 0) => {
                let (identifier, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid IWA data reference: {message}"))
                })?;
                if identifier == 0 {
                    return Err(iwa_error(part, "IWA data reference identifier is zero"));
                }
                data_references.push(identifier);
                cursor += consumed;
            }
            _ => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(IwaMessageInfo {
        message_type: message_type
            .ok_or_else(|| iwa_error(part, "IWA MessageInfo has no message type"))?,
        payload_length: payload_length
            .ok_or_else(|| iwa_error(part, "IWA MessageInfo has no payload length"))?,
        data_references,
    })
}

fn shared_document_kind(bytes: &[u8]) -> Result<SharedDocumentKind, Diagnostic> {
    let mut cursor = 0_usize;
    let mut sheet_references = 0_usize;
    let mut show_references = 0_usize;
    let mut document_stylesheet = 0_usize;
    let mut annotation_storage = 0_usize;
    let mut calculation_engine = 0_usize;
    let mut deprecated_or_keynote_super = 0_usize;
    let mut numbers_super = 0_usize;
    let mut field_count = 0_usize;
    while cursor < bytes.len() {
        field_count += 1;
        if field_count > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                DOCUMENT_COMPONENT,
                "iWork root document payload contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                DOCUMENT_COMPONENT,
                format!("invalid iWork root document field key: {message}"),
            )
        })?;
        cursor += consumed;
        let field = key >> 3;
        let wire = key & 7;
        if field == 0 {
            return Err(iwa_error(
                DOCUMENT_COMPONENT,
                "iWork root document payload contains protobuf field zero",
            ));
        }
        match (field, wire) {
            (field @ (1 | 2 | 4 | 5 | 6), 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    DOCUMENT_COMPONENT,
                    "iWork root reference",
                )?;
                parse_iwa_reference(reference)?;
                let counter = match field {
                    1 => &mut sheet_references,
                    2 => &mut show_references,
                    4 => &mut document_stylesheet,
                    5 => &mut annotation_storage,
                    6 => &mut calculation_engine,
                    _ => unreachable!(),
                };
                *counter = counter.checked_add(1).ok_or_else(|| {
                    iwa_error(DOCUMENT_COMPONENT, "iWork reference count overflows")
                })?;
            }
            (field @ (3 | 8), 2) => {
                let message = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    DOCUMENT_COMPONENT,
                    "iWork document super archive",
                )?;
                if message.is_empty() {
                    return Err(iwa_error(
                        DOCUMENT_COMPONENT,
                        "iWork document super archive is empty",
                    ));
                }
                validate_bounded_protobuf_message(message, DOCUMENT_COMPONENT)?;
                let counter = if field == 3 {
                    &mut deprecated_or_keynote_super
                } else {
                    &mut numbers_super
                };
                *counter = counter.checked_add(1).ok_or_else(|| {
                    iwa_error(DOCUMENT_COMPONENT, "iWork super archive count overflows")
                })?;
            }
            (1 | 2 | 3 | 4 | 5 | 6 | 8, _) => {
                return Err(iwa_error(
                    DOCUMENT_COMPONENT,
                    "iWork root identity field has an invalid protobuf wire type",
                ));
            }
            _ => skip_protobuf_value(bytes, &mut cursor, wire, DOCUMENT_COMPONENT)?,
        }
    }
    let numbers = sheet_references >= 1
        && show_references == 0
        && document_stylesheet == 1
        && annotation_storage == 1
        && calculation_engine == 1
        && deprecated_or_keynote_super <= 1
        && numbers_super == 1;
    let keynote = sheet_references == 0
        && show_references == 1
        && document_stylesheet == 0
        && annotation_storage == 0
        && calculation_engine == 0
        && deprecated_or_keynote_super == 1
        && numbers_super == 0;
    match (numbers, keynote) {
        (true, false) => Ok(SharedDocumentKind::Numbers),
        (false, true) => Ok(SharedDocumentKind::Keynote),
        _ => Err(iwa_error(
            DOCUMENT_COMPONENT,
            "iWork root document has ambiguous Numbers/Keynote native identity",
        )),
    }
}

fn read_iwa_length_delimited<'a>(
    bytes: &'a [u8],
    cursor: &mut usize,
    part: &str,
    label: &str,
) -> Result<&'a [u8], Diagnostic> {
    let (length, consumed) = read_varint(&bytes[*cursor..])
        .map_err(|message| iwa_error(part, format!("invalid {label} length: {message}")))?;
    *cursor += consumed;
    let length = usize::try_from(length)
        .map_err(|_| iwa_error(part, format!("{label} length is not addressable")))?;
    let end = cursor
        .checked_add(length)
        .ok_or_else(|| iwa_error(part, format!("{label} length overflows")))?;
    let value = bytes
        .get(*cursor..end)
        .ok_or_else(|| iwa_error(part, format!("{label} is truncated")))?;
    *cursor = end;
    Ok(value)
}

fn validate_bounded_protobuf_message(bytes: &[u8], part: &str) -> Result<(), Diagnostic> {
    let mut cursor = 0_usize;
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                part,
                "nested iWork protobuf contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid nested protobuf key: {message}"))
        })?;
        cursor += consumed;
        if key >> 3 == 0 {
            return Err(iwa_error(part, "nested iWork protobuf contains field zero"));
        }
        skip_protobuf_value(bytes, &mut cursor, key & 7, part)?;
    }
    Ok(())
}

fn parse_iwa_reference(bytes: &[u8]) -> Result<u64, Diagnostic> {
    let mut cursor = 0_usize;
    let mut identifier = None;
    let mut field_count = 0_usize;
    while cursor < bytes.len() {
        field_count += 1;
        if field_count > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                DOCUMENT_COMPONENT,
                "iWork native reference contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                DOCUMENT_COMPONENT,
                format!("invalid iWork native reference key: {message}"),
            )
        })?;
        cursor += consumed;
        let field = key >> 3;
        let wire = key & 7;
        if field == 0 {
            return Err(iwa_error(
                DOCUMENT_COMPONENT,
                "iWork native reference contains protobuf field zero",
            ));
        }
        if (field, wire) == (1, 0) {
            let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                iwa_error(
                    DOCUMENT_COMPONENT,
                    format!("invalid iWork native reference identifier: {message}"),
                )
            })?;
            if value == 0 || identifier.replace(value).is_some() {
                return Err(iwa_error(
                    DOCUMENT_COMPONENT,
                    "iWork native reference identifier is zero or repeated",
                ));
            }
            cursor += consumed;
        } else {
            skip_protobuf_value(bytes, &mut cursor, wire, DOCUMENT_COMPONENT)?;
        }
    }
    identifier.ok_or_else(|| {
        iwa_error(
            DOCUMENT_COMPONENT,
            "iWork native reference has no identifier",
        )
    })
}

fn skip_protobuf_value(
    bytes: &[u8],
    cursor: &mut usize,
    wire: u64,
    part: &str,
) -> Result<(), Diagnostic> {
    let length = match wire {
        0 => {
            let (_, consumed) = read_varint(&bytes[*cursor..]).map_err(|message| {
                iwa_error(part, format!("invalid protobuf varint: {message}"))
            })?;
            *cursor += consumed;
            return Ok(());
        }
        1 => 8,
        2 => {
            let (length, consumed) = read_varint(&bytes[*cursor..]).map_err(|message| {
                iwa_error(part, format!("invalid protobuf field length: {message}"))
            })?;
            *cursor += consumed;
            usize::try_from(length)
                .map_err(|_| iwa_error(part, "protobuf field length is not addressable"))?
        }
        5 => 4,
        _ => {
            return Err(iwa_error(
                part,
                "unsupported protobuf wire type in IWA header",
            ));
        }
    };
    let end = cursor
        .checked_add(length)
        .ok_or_else(|| iwa_error(part, "protobuf field length overflow"))?;
    if end > bytes.len() {
        return Err(iwa_error(part, "protobuf field is truncated"));
    }
    *cursor = end;
    Ok(())
}

fn read_varint(bytes: &[u8]) -> Result<(u64, usize), &'static str> {
    let mut value = 0_u64;
    for (index, byte) in bytes.iter().copied().take(10).enumerate() {
        if index == 9 && byte > 1 {
            return Err("varint overflows u64");
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            return Ok((value, index + 1));
        }
    }
    if bytes.len() < 10 {
        Err("varint is truncated")
    } else {
        Err("varint exceeds ten bytes")
    }
}

struct NumbersArchive {
    part: String,
    archive: IwaArchive,
}

struct NumbersMessageSpace {
    archives: Vec<NumbersArchive>,
}

impl NumbersMessageSpace {
    fn message(
        &self,
        identifier: u64,
        message_type: u64,
    ) -> Result<Option<(&str, &IwaMessage)>, Diagnostic> {
        let mut found = None;
        for entry in self
            .archives
            .iter()
            .filter(|entry| entry.archive.identifier == identifier)
        {
            for message in entry
                .archive
                .messages
                .iter()
                .filter(|message| message.message_type == message_type)
            {
                if found.replace((entry.part.as_str(), message)).is_some() {
                    return Err(iwa_error(
                        &entry.part,
                        format!(
                            "Numbers object {identifier} resolves to multiple type {message_type} messages"
                        ),
                    ));
                }
            }
        }
        Ok(found)
    }

    fn any_message(&self, identifier: u64) -> Result<Option<(&str, &IwaMessage)>, Diagnostic> {
        let mut found = None;
        for entry in self
            .archives
            .iter()
            .filter(|entry| entry.archive.identifier == identifier)
        {
            for message in &entry.archive.messages {
                if found.replace((entry.part.as_str(), message)).is_some() {
                    return Err(iwa_error(
                        &entry.part,
                        format!("Numbers object {identifier} resolves to multiple messages"),
                    ));
                }
            }
        }
        Ok(found)
    }
}

fn numbers_document(
    package: &Package<'_>,
    document_archives: &[IwaArchive],
    limits: Limits,
) -> Result<Document, Diagnostic> {
    let mut archives = document_archives
        .iter()
        .cloned()
        .map(|archive| NumbersArchive {
            part: DOCUMENT_COMPONENT.to_owned(),
            archive,
        })
        .collect::<Vec<_>>();
    let parts = package
        .entry_names()
        .filter(|part| is_iwa_component(part) && *part != DOCUMENT_COMPONENT)
        .map(str::to_owned)
        .collect::<Vec<_>>();
    for part in parts {
        let bytes = package.required_part(&part)?;
        for archive in parse_iwa_archives(package, &part, &bytes, limits)? {
            archives.push(NumbersArchive {
                part: part.clone(),
                archive,
            });
        }
    }
    let stylesheet_archives = archives
        .iter()
        .filter(|entry| {
            matches!(
                entry.part.as_str(),
                STYLESHEET_COMPONENT | THEME_STYLESHEET_COMPONENT
            )
        })
        .map(|entry| entry.archive.clone())
        .collect::<Vec<_>>();
    // TSCH title/series archives and cached grids have the same semantics in
    // Numbers, Pages and Keynote; sheet ownership stays in this adapter.
    let chart_styles = archives
        .iter()
        .filter(|entry| {
            entry.archive.messages.iter().any(|message| {
                matches!(
                    message.message_type,
                    IWORK_CHART_STYLE_TYPE
                        ..=IWORK_CHART_SERIES_NONSTYLE_TYPE
                            | IWORK_PARAGRAPH_STYLE_TYPE
                            | IWORK_CHARACTER_STYLE_TYPE
                )
            })
        })
        .map(|entry| entry.archive.clone())
        .collect::<Vec<_>>();
    let messages = NumbersMessageSpace { archives };
    let root = root_message(document_archives, DOCUMENT_COMPONENT)?;
    let sheet_references = numbers_references(&root.payload, 1, DOCUMENT_COMPONENT)?;
    if sheet_references.is_empty() {
        return Err(Diagnostic::fatal(
            DiagnosticCode::UnsupportedFeature,
            Phase::Parse,
            None,
            "NUMBERS_NATIVE_SHEETS_MISSING: Numbers document contains no native sheet references",
        )
        .in_part(DOCUMENT_COMPONENT));
    }

    let mut diagnostics = vec![
        Diagnostic::warning(
            DiagnosticCode::UnsupportedFeature,
            Phase::Render,
            Fidelity::Approximate,
            "NUMBERS_NATIVE_STYLE_PARTIAL: native sheets and cells are rendered; advanced Numbers table styles, formulas, and floating canvas objects are not yet fully reproduced",
        )
        .in_part(DOCUMENT_COMPONENT),
    ];
    let mut units = Vec::new();
    let mut objects = Vec::new();
    for sheet_identifier in sheet_references {
        let (sheet_part, sheet) = messages
            .message(sheet_identifier, NUMBERS_SHEET_TYPE)?
            .ok_or_else(|| {
                iwa_error(
                    DOCUMENT_COMPONENT,
                    format!("Numbers sheet reference {sheet_identifier} is missing"),
                )
            })?;
        let sheet_name = numbers_first_bytes(&sheet.payload, 1, sheet_part)?
            .and_then(|bytes| std::str::from_utf8(bytes).ok())
            .filter(|name| !name.is_empty())
            .unwrap_or("Sheet")
            .to_owned();
        let first_unit = units.len();
        let drawable_references = numbers_references(&sheet.payload, 2, sheet_part)?;
        for table_identifier in drawable_references.iter().copied() {
            let Some((table_part, table_info)) =
                messages.message(table_identifier, NUMBERS_TABLE_INFO_TYPE)?
            else {
                continue;
            };
            let Some(table_model_reference) =
                numbers_references(&table_info.payload, 2, table_part)?
                    .into_iter()
                    .next()
            else {
                return Err(iwa_error(
                    table_part,
                    format!("Numbers table {table_identifier} has no table-model reference"),
                ));
            };
            let (model_part, model) = messages
                .message(table_model_reference, NUMBERS_TABLE_MODEL_TYPE)?
                .ok_or_else(|| {
                    iwa_error(
                        table_part,
                        format!("Numbers table model {table_model_reference} is missing"),
                    )
                })?;
            let unit_index = u32::try_from(units.len())
                .map_err(|_| iwa_error(model_part, "Numbers table count exceeds u32"))?;
            let table_index = units.len() - first_unit;
            let name = if table_index == 0 {
                sheet_name.clone()
            } else {
                format!("{sheet_name} – Table {}", table_index + 1)
            };
            let (unit, table_diagnostics) = numbers_table(
                &messages,
                model,
                model_part,
                unit_index,
                name,
                limits,
                &stylesheet_archives,
                &mut objects,
            )?;
            units.push(unit);
            diagnostics.extend(table_diagnostics);
        }
        let first_chart_object = objects.len();
        let unit_index = u32::try_from(first_unit)
            .map_err(|_| iwa_error(sheet_part, "Numbers sheet count exceeds u32"))?;
        for identifier in drawable_references {
            let Some((part, message)) = messages.message(identifier, IWORK_CHART_TYPE)? else {
                continue;
            };
            let chart = match keynote_chart(&message.payload, &chart_styles, &chart_styles, part) {
                Ok(Some(chart)) => chart,
                result => {
                    let reason = match result {
                        Err(error) if error.code != DiagnosticCode::FormatInvalid => {
                            return Err(error);
                        }
                        Err(error) => error.to_string(),
                        _ => "unsupported chart type or missing cached data".to_owned(),
                    };
                    diagnostics.push(
                        Diagnostic::warning(
                            DiagnosticCode::UnsupportedFeature,
                            Phase::Parse,
                            Fidelity::Omitted,
                            format!("NUMBERS_CHART_OMITTED: chart {identifier}: {reason}"),
                        )
                        .in_part(part),
                    );
                    continue;
                }
            };
            keynote_push_chart(
                chart,
                identifier,
                unit_index,
                part,
                limits,
                &mut objects,
                0.0,
                0.0,
                100,
            )?;
        }
        if objects.len() > first_chart_object {
            // Tables are already in CSS pixels; only new chart objects need the
            // shared point conversion. Keep the existing table-per-unit contract.
            scale_iwork_point_objects(&mut objects[first_chart_object..]);
            if first_unit == units.len() {
                units.push(Unit {
                    kind: UnitKind::Sheet,
                    index: unit_index,
                    id: format!("numbers-sheet-{sheet_identifier}"),
                    name: sheet_name,
                    width: 0.0,
                    height: 0.0,
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
                });
            }
            let unit = &mut units[first_unit];
            for object in &objects[first_chart_object..] {
                unit.width = unit.width.max(object.bounds.x + object.bounds.width);
                unit.height = unit.height.max(object.bounds.y + object.bounds.height);
            }
        }
    }
    if units.is_empty() {
        return Err(Diagnostic::fatal(
            DiagnosticCode::UnsupportedFeature,
            Phase::Parse,
            None,
            "NUMBERS_NATIVE_TABLES_MISSING: Numbers document contains no native tables",
        )
        .in_part(DOCUMENT_COMPONENT));
    }

    Ok(Document {
        fatal: false,
        format: Some(DocumentFormat::Numbers),
        kind: Some(DocumentKind::Spreadsheet),
        units,
        outline: Vec::new(),
        objects,
        embedded_fonts: Vec::new(),
        font_alternate_names: Vec::new(),
        diagnostics,
    })
}

fn numbers_table(
    messages: &NumbersMessageSpace,
    model: &IwaMessage,
    model_part: &str,
    unit_index: u32,
    name: String,
    limits: Limits,
    stylesheet_archives: &[IwaArchive],
    objects: &mut Vec<Object>,
) -> Result<(Unit, Vec<Diagnostic>), Diagnostic> {
    let mut diagnostics = Vec::new();
    let (rows, columns, mut cells) = numbers_table_cells(messages, model, model_part, limits)?;
    let store = numbers_first_bytes(&model.payload, 4, model_part)?
        .ok_or_else(|| iwa_error(model_part, "Numbers table has no data store"))?;
    let merges = numbers_table_merges(messages, model, store, rows, columns, model_part, limits)?;
    let row_references = numbers_first_bytes(store, 1, model_part)?
        .map(|headers| numbers_references(headers, 2, model_part))
        .transpose()?
        .unwrap_or_default();
    let column_references = numbers_references(store, 2, model_part)?;
    let row_count = usize::try_from(rows)
        .map_err(|_| iwa_error(model_part, "Numbers row count is not addressable"))?;
    let column_count = usize::try_from(columns)
        .map_err(|_| iwa_error(model_part, "Numbers column count is not addressable"))?;
    let row_heights = iwork_table_header_sizes(messages, &row_references, row_count)?
        .into_iter()
        .map(|size| size.unwrap_or(15.0) * IWORK_POINT_TO_CSS_PIXEL)
        .collect::<Vec<_>>();
    let column_widths = iwork_table_header_sizes(messages, &column_references, column_count)?
        .into_iter()
        .map(|size| size.unwrap_or(72.0) * IWORK_POINT_TO_CSS_PIXEL)
        .collect::<Vec<_>>();
    let mut row_offsets = Vec::with_capacity(row_heights.len() + 1);
    row_offsets.push(0.0);
    for height in &row_heights {
        row_offsets.push(row_offsets.last().copied().unwrap_or(0.0) + height);
    }
    let mut column_offsets = Vec::with_capacity(column_widths.len() + 1);
    column_offsets.push(0.0);
    for width in &column_widths {
        column_offsets.push(column_offsets.last().copied().unwrap_or(0.0) + width);
    }
    let width = column_offsets.last().copied().unwrap_or(0.0);
    let height = row_offsets.last().copied().unwrap_or(0.0);
    let table_cell_styles = keynote_table_cell_styles(
        messages,
        stylesheet_archives,
        model,
        model_part,
        width / IWORK_POINT_TO_CSS_PIXEL,
        height / IWORK_POINT_TO_CSS_PIXEL,
        limits,
    )?;
    let cell_text_styles =
        iwork_table_text_styles(messages, stylesheet_archives, model, model_part, limits)?;
    let borders = iwork_table_borders(messages, model, model_part, rows, columns, &mut diagnostics)
        .unwrap_or_default();
    let occupied_cells = cells
        .iter()
        .filter(|cell| cell.text.as_deref().is_some_and(|text| !text.is_empty()))
        .map(|cell| (cell.row, cell.column))
        .collect::<HashSet<_>>();
    cells.sort_by_key(|cell| cell.text.is_some());
    for cell in cells {
        if cell.span {
            continue;
        }
        let row = usize::try_from(cell.row)
            .map_err(|_| iwa_error(&cell.part, "Numbers row is not addressable"))?;
        let column = usize::try_from(cell.column)
            .map_err(|_| iwa_error(&cell.part, "Numbers column is not addressable"))?;
        let (end_row, end_column) = merges
            .get(&(cell.row, cell.column))
            .copied()
            .unwrap_or((cell.row.saturating_add(1), cell.column.saturating_add(1)));
        let end_row = usize::try_from(end_row)
            .map_err(|_| iwa_error(&cell.part, "Numbers merged row is not addressable"))?;
        let end_column = usize::try_from(end_column)
            .map_err(|_| iwa_error(&cell.part, "Numbers merged column is not addressable"))?;
        if end_row >= row_offsets.len() || end_column >= column_offsets.len() {
            continue;
        }
        let table_cell_style = cell
            .style_index
            .and_then(|index| usize::try_from(index).ok())
            .and_then(|index| table_cell_styles.get(index))
            .cloned()
            .unwrap_or_default();
        let fill = table_cell_style.fill;
        let mut visual = if let Some(text) = cell.text.as_ref() {
            let style = cell
                .text_style_index
                .and_then(|index| usize::try_from(index).ok())
                .and_then(|index| cell_text_styles.get(index))
                .and_then(Clone::clone)
                .map(Ok)
                .unwrap_or_else(|| {
                    pages_table_text_style(
                        model,
                        model_part,
                        stylesheet_archives,
                        cell.row,
                        cell.column,
                        rows,
                    )
                })?;
            let previous_occupied = cell
                .column
                .checked_sub(1)
                .is_some_and(|column| occupied_cells.contains(&(cell.row, column)));
            let next_occupied = u32::try_from(end_column)
                .ok()
                .is_some_and(|column| occupied_cells.contains(&(cell.row, column)));
            let clip_horizontal_overflow = crate::model::spreadsheet_text_overflow_is_clipped(
                style.align,
                previous_occupied,
                next_occupied,
            );
            Visual::TextLayout {
                layout: TextLayout {
                    wrap: table_cell_style.wrap,
                    horizontal_overflow: if table_cell_style.wrap || clip_horizontal_overflow {
                        TextHorizontalOverflow::Clip
                    } else {
                        TextHorizontalOverflow::Overflow
                    },
                    ..TextLayout::default()
                },
                visual: Box::new(Visual::RichText {
                    geometry: Geometry::Rectangle,
                    fill,
                    stroke: Paint::None,
                    stroke_width: 0.0,
                    align: style.align,
                    line_height: style
                        .line_height_multiple
                        .map_or(style.font_size * 1.2, |multiple| {
                            if multiple > 4.0 {
                                multiple
                            } else {
                                style.font_size * multiple
                            }
                        })
                        .max(style.font_size),
                    runs: vec![keynote_text_run(iwork_visible_table_text(text), &style)],
                }),
            }
        } else if !matches!(fill, Paint::None) {
            Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill,
                stroke: Paint::None,
                stroke_width: 0.0,
            }
        } else {
            continue;
        };
        scale_iwork_visual(&mut visual);
        if objects.len() >= limits.max_document_objects {
            return Err(iwa_error(
                &cell.part,
                "Numbers cells exceed the configured object limit",
            ));
        }
        let numeric_id = u32::try_from(objects.len())
            .map_err(|_| iwa_error(&cell.part, "Numbers object count exceeds u32"))?;
        let address = numbers_cell_address(cell.row, cell.column);
        objects.push(Object {
            numeric_id,
            parent_numeric_id: None,
            stable_id: format!("numbers:{unit_index}:{address}"),
            parent_stable_id: None,
            kind: ObjectKind::Cell,
            unit_index,
            bounds: Rect {
                x: column_offsets[column],
                y: row_offsets[row],
                width: column_offsets[end_column] - column_offsets[column],
                height: row_offsets[end_row] - row_offsets[row],
            },
            z: i32::try_from(numeric_id)
                .map_err(|_| iwa_error(&cell.part, "Numbers z-order exceeds i32"))?,
            text: cell.text,
            source: SourceRef {
                part: cell.part.clone(),
                mapping: MappingQuality::Exact,
                locator: SourceLocator::Iwork {
                    kind: "cell",
                    component: format!("{}#{address}", cell.part),
                },
            },
            visual,
        });
    }

    for border in borders {
        if objects.len() >= limits.max_document_objects {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::ObjectLimit,
                    Phase::Render,
                    Fidelity::Omitted,
                    "Numbers table borders were omitted to preserve cell content within the object limit",
                )
                .in_part(model_part),
            );
            break;
        }
        let Some((bounds, mut visual)) =
            iwork_table_border_visual(&border, &column_offsets, &row_offsets, 0)
        else {
            continue;
        };
        scale_iwork_visual(&mut visual);
        let numeric_id = u32::try_from(objects.len())
            .map_err(|_| iwa_error(model_part, "Numbers object count exceeds u32"))?;
        let orientation = if border.horizontal {
            "horizontal"
        } else {
            "vertical"
        };
        objects.push(Object {
            numeric_id,
            parent_numeric_id: None,
            stable_id: format!(
                "numbers:{unit_index}:border:{orientation}:{}:{}:{}",
                border.boundary, border.origin, border.length
            ),
            parent_stable_id: None,
            kind: ObjectKind::Shape,
            unit_index,
            bounds,
            z: i32::try_from(numeric_id)
                .map_err(|_| iwa_error(model_part, "Numbers z-order exceeds i32"))?,
            text: None,
            source: SourceRef {
                part: model_part.to_owned(),
                mapping: MappingQuality::Exact,
                locator: SourceLocator::Iwork {
                    kind: "cell",
                    component: format!("{model_part}#border-{orientation}"),
                },
            },
            visual,
        });
    }

    Ok((
        Unit {
            kind: UnitKind::Sheet,
            index: unit_index,
            id: format!("numbers-table-{unit_index}"),
            name,
            width,
            height,
            rows,
            columns,
            frozen_rows: 0,
            frozen_columns: 0,
            frozen_width: 0.0,
            frozen_height: 0.0,
            row_axis: crate::model::SheetAxis::from_sizes(&row_heights, 20.0),
            column_axis: crate::model::SheetAxis::from_sizes(&column_widths, 96.0),
            show_grid_lines: true,
            tab_color: None,
            sheet: None,
            slide: None,
        },
        diagnostics,
    ))
}

fn iwork_table_borders(
    messages: &NumbersMessageSpace,
    model: &IwaMessage,
    part: &str,
    rows: u32,
    columns: u32,
    diagnostics: &mut Vec<Diagnostic>,
) -> Option<Vec<NumbersTableBorder>> {
    let result = numbers_references(&model.payload, 49, part).and_then(|references| {
        if references.is_empty() {
            Ok(None)
        } else {
            numbers_table_borders(messages, model, part, rows, columns).map(Some)
        }
    });
    match result {
        Ok(borders) => borders,
        Err(error) => {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::FormatInvalid,
                    Phase::Parse,
                    Fidelity::Omitted,
                    format!("NUMBERS_STROKE_SIDECAR_INVALID: {error}"),
                )
                .in_part(part),
            );
            // Authored but unreadable borders must not replace usable cell content.
            Some(Vec::new())
        }
    }
}

// TST sidecar strokes share geometry across Numbers, Pages and Keynote.
// Offsets are in the caller's units; stroke widths remain in iWork points.
fn iwork_table_border_visual(
    border: &NumbersTableBorder,
    columns: &[f32],
    rows: &[f32],
    row_start: u32,
) -> Option<(Rect, Visual)> {
    let row_end = row_start.checked_add(u32::try_from(rows.len().checked_sub(1)?).ok()?)?;
    let end = border.origin.saturating_add(border.length);
    let bounds = if border.horizontal {
        if !(row_start..=row_end).contains(&border.boundary) {
            return None;
        }
        Rect {
            x: *columns.get(border.origin as usize)?,
            y: *rows.get((border.boundary - row_start) as usize)?,
            width: *columns.get(end as usize)? - *columns.get(border.origin as usize)?,
            height: 0.0,
        }
    } else {
        let start = border.origin.max(row_start);
        let end = end.min(row_end);
        if start >= end {
            return None;
        }
        Rect {
            x: *columns.get(border.boundary as usize)?,
            y: *rows.get((start - row_start) as usize)?,
            width: 0.0,
            height: *rows.get((end - row_start) as usize)?
                - *rows.get((start - row_start) as usize)?,
        }
    };
    let mut visual = Visual::PaintedShape {
        geometry: Geometry::Line,
        fill: Paint::None,
        stroke: Paint::Solid(border.stroke.color),
        stroke_width: border.stroke.width,
    };
    if border.stroke.style != StrokeStyle::default() {
        visual = Visual::StrokeStyle {
            style: border.stroke.style.clone(),
            visual: Box::new(visual),
        };
    }
    Some((bounds, visual))
}

fn numbers_table_borders(
    messages: &NumbersMessageSpace,
    model: &IwaMessage,
    model_part: &str,
    rows: u32,
    columns: u32,
) -> Result<Vec<NumbersTableBorder>, Diagnostic> {
    let Some(sidecar_reference) = numbers_references(&model.payload, 49, model_part)?
        .into_iter()
        .next()
    else {
        return Ok(Vec::new());
    };
    let (sidecar_part, sidecar) = messages
        .message(sidecar_reference, NUMBERS_STROKE_SIDECAR_TYPE)?
        .ok_or_else(|| {
            iwa_error(
                model_part,
                format!("Numbers stroke sidecar {sidecar_reference} is missing"),
            )
        })?;
    let mut runs = Vec::new();
    for (field, horizontal, trailing) in [
        (6, true, false),
        (4, false, false),
        (5, false, true),
        (7, true, true),
    ] {
        for layer_reference in numbers_references(&sidecar.payload, field, sidecar_part)? {
            let (layer_part, layer) = messages
                .message(layer_reference, NUMBERS_STROKE_LAYER_TYPE)?
                .ok_or_else(|| {
                    iwa_error(
                        sidecar_part,
                        format!("Numbers stroke layer {layer_reference} is missing"),
                    )
                })?;
            let Some(index) = numbers_varint(&layer.payload, 1, layer_part)?
                .and_then(|value| u32::try_from(value).ok())
            else {
                return Err(iwa_error(
                    layer_part,
                    "Numbers stroke layer has no valid index",
                ));
            };
            let Some(boundary) = index.checked_add(u32::from(trailing)) else {
                continue;
            };
            let segment_limit = if horizontal { columns } else { rows };
            let boundary_limit = if horizontal { rows } else { columns };
            if boundary > boundary_limit {
                continue;
            }
            for run in numbers_bytes(&layer.payload, 2, layer_part)? {
                let origin = numbers_varint(run, 1, layer_part)?
                    .and_then(|value| u32::try_from(value).ok())
                    .ok_or_else(|| iwa_error(layer_part, "Numbers stroke run has no origin"))?;
                let length = numbers_varint(run, 2, layer_part)?
                    .and_then(|value| u32::try_from(value).ok())
                    .ok_or_else(|| iwa_error(layer_part, "Numbers stroke run has no length"))?;
                let order = numbers_varint(run, 4, layer_part)?.unwrap_or(0);
                let end = origin.saturating_add(length).min(segment_limit);
                if origin >= end {
                    continue;
                }
                let stroke = numbers_first_bytes(run, 3, layer_part)?
                    .map(|stroke| numbers_stroke(stroke, layer_part))
                    .transpose()?
                    .flatten();
                runs.push((order, horizontal, boundary, origin, end, stroke));
            }
        }
    }
    runs.sort_by_key(|run| run.0);
    let mut segments = HashMap::new();
    for (_, horizontal, boundary, origin, end, stroke) in runs {
        for segment in origin..end {
            let key = (horizontal, boundary, segment);
            if let Some(stroke) = &stroke {
                segments.insert(key, stroke.clone());
            } else {
                segments.remove(&key);
            }
        }
    }
    let mut segments = segments.into_iter().collect::<Vec<_>>();
    segments
        .sort_by_key(|((horizontal, boundary, segment), _)| (!*horizontal, *boundary, *segment));
    let mut borders: Vec<NumbersTableBorder> = Vec::new();
    for ((horizontal, boundary, origin), stroke) in segments {
        if let Some(previous) = borders.last_mut()
            && previous.horizontal == horizontal
            && previous.boundary == boundary
            && previous.origin.saturating_add(previous.length) == origin
            && previous.stroke == stroke
        {
            previous.length = previous.length.saturating_add(1);
        } else {
            borders.push(NumbersTableBorder {
                horizontal,
                boundary,
                origin,
                length: 1,
                stroke,
            });
        }
    }
    Ok(borders)
}

fn numbers_stroke(bytes: &[u8], part: &str) -> Result<Option<NumbersStroke>, Diagnostic> {
    let (paint, width, style) = keynote_stroke(bytes, part)?;
    Ok(match paint {
        Paint::Solid(color) if width > 0.0 => Some(NumbersStroke {
            color,
            width,
            style,
        }),
        _ => None,
    })
}

fn numbers_table_cells(
    messages: &NumbersMessageSpace,
    model: &IwaMessage,
    model_part: &str,
    limits: Limits,
) -> Result<(u32, u32, Vec<NumbersCell>), Diagnostic> {
    let rows = numbers_varint(&model.payload, 6, model_part)?
        .and_then(|value| u32::try_from(value).ok())
        .filter(|value| *value > 0)
        .ok_or_else(|| iwa_error(model_part, "Numbers table has no valid row count"))?;
    let columns = numbers_varint(&model.payload, 7, model_part)?
        .and_then(|value| u32::try_from(value).ok())
        .filter(|value| *value > 0)
        .ok_or_else(|| iwa_error(model_part, "Numbers table has no valid column count"))?;
    let store = numbers_first_bytes(&model.payload, 4, model_part)?
        .ok_or_else(|| iwa_error(model_part, "Numbers table has no data store"))?;
    let strings = numbers_data_list_reference(messages, store, 4, model_part, limits)?;
    let rich_strings = numbers_data_list_reference(messages, store, 17, model_part, limits)?;
    let tile_storage = numbers_first_bytes(store, 3, model_part)?
        .ok_or_else(|| iwa_error(model_part, "Numbers table has no tile storage"))?;
    let mut tiles = Vec::new();
    for entry in numbers_bytes(tile_storage, 1, model_part)? {
        let tile_order = numbers_varint(entry, 1, model_part)?
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| iwa_error(model_part, "Numbers tile entry has no valid order"))?;
        let Some(reference) = numbers_references(entry, 2, model_part)?.into_iter().next() else {
            return Err(iwa_error(model_part, "Numbers tile entry has no reference"));
        };
        tiles.push((tile_order, reference));
    }
    tiles.sort_unstable_by_key(|(order, _)| *order);

    let mut cells = Vec::new();
    for (tile_order, tile_reference) in tiles {
        let (tile_part, tile) = messages
            .message(tile_reference, NUMBERS_TILE_TYPE)?
            .ok_or_else(|| {
                iwa_error(
                    model_part,
                    format!("Numbers tile {tile_reference} is missing"),
                )
            })?;
        let row_base = tile_order
            .checked_mul(NUMBERS_TILE_ROW_CAPACITY)
            .ok_or_else(|| iwa_error(tile_part, "Numbers tile row base overflows"))?;
        let modern = numbers_varint(&tile.payload, 7, tile_part)?.unwrap_or(0) > 0;
        for row_info in numbers_bytes(&tile.payload, 5, tile_part)? {
            let local_row = numbers_varint(row_info, 1, tile_part)?
                .and_then(|value| u32::try_from(value).ok())
                .ok_or_else(|| iwa_error(tile_part, "Numbers tile row has no row index"))?;
            let row = row_base
                .checked_add(local_row)
                .ok_or_else(|| iwa_error(tile_part, "Numbers row index overflows"))?;
            if row >= rows {
                return Err(iwa_error(
                    tile_part,
                    "Numbers tile row exceeds table bounds",
                ));
            }
            for (column, cell) in numbers_tile_row_cells(row_info, modern, tile_part)? {
                if column >= columns {
                    return Err(iwa_error(
                        tile_part,
                        "Numbers tile cell exceeds table bounds",
                    ));
                }
                if cells.len() >= limits.max_document_objects {
                    return Err(iwa_error(
                        tile_part,
                        "Numbers cells exceed the configured object limit",
                    ));
                }
                let (style_index, text_style_index) = numbers_cell_style_indexes(cell, tile_part)?;
                let text = numbers_cell_text(cell, &strings, &rich_strings, tile_part)?;
                cells.push(NumbersCell {
                    row,
                    column,
                    text_storage: text.as_ref().and_then(|text| text.storage),
                    text: text.map(|text| text.text),
                    span: cell.get(1).copied() == Some(1),
                    part: tile_part.to_owned(),
                    style_index,
                    text_style_index,
                });
            }
        }
    }
    Ok((rows, columns, cells))
}

fn numbers_data_list_reference(
    messages: &NumbersMessageSpace,
    store: &[u8],
    field: u64,
    part: &str,
    limits: Limits,
) -> Result<Vec<Option<IworkCellText>>, Diagnostic> {
    let Some(reference) = numbers_references(store, field, part)?.into_iter().next() else {
        return Ok(Vec::new());
    };
    let (list_part, list) = messages
        .message(reference, NUMBERS_DATA_LIST_TYPE)?
        .ok_or_else(|| iwa_error(part, format!("Numbers data list {reference} is missing")))?;
    let list_type = numbers_varint(&list.payload, 1, list_part)?.unwrap_or(0);
    let mut values = Vec::new();
    for entry in numbers_bytes(&list.payload, 3, list_part)? {
        let index = numbers_varint(entry, 1, list_part)?
            .and_then(|value| usize::try_from(value).ok())
            .ok_or_else(|| iwa_error(list_part, "Numbers data-list entry has no valid index"))?;
        if index >= limits.max_document_objects {
            return Err(iwa_error(
                list_part,
                "Numbers data-list index exceeds the configured object limit",
            ));
        }
        if values.len() <= index {
            values.resize(index + 1, None);
        }
        values[index] = match list_type {
            1 => numbers_first_bytes(entry, 3, list_part)?.map(|bytes| IworkCellText {
                text: String::from_utf8_lossy(bytes).into_owned(),
                storage: None,
            }),
            8 => {
                let reference = numbers_references(entry, 9, list_part)?.into_iter().next();
                reference
                    .map(|reference| numbers_rich_text(messages, reference, list_part))
                    .transpose()?
            }
            _ => None,
        };
    }
    Ok(values)
}

fn numbers_rich_text(
    messages: &NumbersMessageSpace,
    identifier: u64,
    part: &str,
) -> Result<IworkCellText, Diagnostic> {
    let (payload_part, payload) = messages
        .any_message(identifier)?
        .ok_or_else(|| iwa_error(part, format!("Numbers rich text {identifier} is missing")))?;
    let storage_reference = numbers_references(&payload.payload, 1, payload_part)?
        .into_iter()
        .next()
        .ok_or_else(|| iwa_error(payload_part, "Numbers rich text has no storage reference"))?;
    let (storage_part, storage) = messages
        .message(storage_reference, IWORK_TEXT_STORAGE_TYPE)?
        .ok_or_else(|| {
            iwa_error(
                payload_part,
                format!("Numbers text storage {storage_reference} is missing"),
            )
        })?;
    Ok(IworkCellText {
        text: numbers_bytes(&storage.payload, 3, storage_part)?
            .into_iter()
            .map(|bytes| String::from_utf8_lossy(bytes))
            .collect::<String>(),
        storage: Some(storage_reference),
    })
}

fn numbers_tile_row_cells<'a>(
    row: &'a [u8],
    modern: bool,
    part: &str,
) -> Result<Vec<(u32, &'a [u8])>, Diagnostic> {
    let count = numbers_varint(row, 2, part)?
        .and_then(|value| usize::try_from(value).ok())
        .ok_or_else(|| iwa_error(part, "Numbers tile row has no cell count"))?;
    let wide = numbers_varint(row, 8, part)?.unwrap_or(0) > 0;
    let width = if wide { 4_usize } else { 1_usize };
    let selected = if modern {
        (
            numbers_first_bytes(row, 6, part)?,
            numbers_first_bytes(row, 7, part)?,
        )
    } else {
        (
            numbers_first_bytes(row, 3, part)?,
            numbers_first_bytes(row, 4, part)?,
        )
    };
    let fallback = if modern {
        (
            numbers_first_bytes(row, 3, part)?,
            numbers_first_bytes(row, 4, part)?,
        )
    } else {
        (
            numbers_first_bytes(row, 6, part)?,
            numbers_first_bytes(row, 7, part)?,
        )
    };
    let (storage, offsets) = match selected {
        (Some(storage), Some(offsets)) => (storage, offsets),
        _ => match fallback {
            (Some(storage), Some(offsets)) => (storage, offsets),
            _ if count == 0 => return Ok(Vec::new()),
            _ => return Err(iwa_error(part, "Numbers tile row storage is missing")),
        },
    };
    if offsets.len() % 2 != 0 {
        return Err(iwa_error(part, "Numbers tile row offsets are truncated"));
    }
    let mut present = Vec::new();
    for (column, encoded) in offsets.chunks_exact(2).enumerate() {
        let offset = u16::from_le_bytes([encoded[0], encoded[1]]);
        if offset != u16::MAX {
            let byte_offset = usize::from(offset)
                .checked_mul(width)
                .ok_or_else(|| iwa_error(part, "Numbers cell offset overflows"))?;
            if byte_offset >= storage.len() {
                return Err(iwa_error(part, "Numbers cell offset exceeds row storage"));
            }
            present.push((
                u32::try_from(column)
                    .map_err(|_| iwa_error(part, "Numbers column index exceeds u32"))?,
                byte_offset,
            ));
        }
    }
    if present.len() != count {
        return Err(iwa_error(
            part,
            format!(
                "Numbers tile row declares {count} cells but contains {} offsets",
                present.len()
            ),
        ));
    }
    let mut cells = Vec::with_capacity(present.len());
    for (index, (column, start)) in present.iter().copied().enumerate() {
        let end = present
            .get(index + 1)
            .map_or(storage.len(), |(_, offset)| *offset);
        let cell = storage
            .get(start..end)
            .ok_or_else(|| iwa_error(part, "Numbers cell storage range is invalid"))?;
        cells.push((column, cell));
    }
    Ok(cells)
}

fn numbers_table_merges(
    messages: &NumbersMessageSpace,
    model: &IwaMessage,
    store: &[u8],
    rows: u32,
    columns: u32,
    part: &str,
    limits: Limits,
) -> Result<HashMap<(u32, u32), (u32, u32)>, Diagnostic> {
    let mut merges = HashMap::new();
    if let Some(owner) = numbers_first_bytes(&model.payload, 47, part)?
        && let Some(formula_store) = numbers_first_bytes(owner, 2, part)?
    {
        for pair in numbers_bytes(formula_store, 3, part)? {
            let Some(formula) = numbers_first_bytes(pair, 2, part)? else {
                continue;
            };
            let Some(ast) = numbers_first_bytes(formula, 1, part)? else {
                continue;
            };
            let Some(node) = numbers_bytes(ast, 1, part)?.into_iter().next() else {
                continue;
            };
            if numbers_varint(node, 1, part)? != Some(67) {
                continue;
            }
            let Some(tract) = numbers_first_bytes(node, 40, part)? else {
                continue;
            };
            let Some(column_range) = numbers_bytes(tract, 3, part)?.into_iter().next() else {
                continue;
            };
            let Some(row_range) = numbers_bytes(tract, 4, part)?.into_iter().next() else {
                continue;
            };
            let Some(column) =
                numbers_varint(column_range, 1, part)?.and_then(|value| u32::try_from(value).ok())
            else {
                continue;
            };
            let Some(row) =
                numbers_varint(row_range, 1, part)?.and_then(|value| u32::try_from(value).ok())
            else {
                continue;
            };
            let end_column = numbers_varint(column_range, 2, part)?
                .and_then(|value| u32::try_from(value).ok())
                .unwrap_or(column)
                .saturating_add(1);
            let end_row = numbers_varint(row_range, 2, part)?
                .and_then(|value| u32::try_from(value).ok())
                .unwrap_or(row)
                .saturating_add(1);
            if end_row > rows || end_column > columns {
                continue;
            }
            if merges.len() >= limits.max_document_objects {
                return Err(iwa_error(
                    part,
                    "Numbers merge regions exceed the configured object limit",
                ));
            }
            merges.insert((row, column), (end_row, end_column));
        }
    }
    if !merges.is_empty() {
        return Ok(merges);
    }
    let Some(reference) = numbers_references(store, 13, part)?.into_iter().next() else {
        return Ok(merges);
    };
    let Some((merge_part, map)) = messages.any_message(reference)? else {
        return Ok(merges);
    };
    for range in numbers_bytes(&map.payload, 1, merge_part)? {
        let Some(origin) = numbers_first_bytes(range, 1, merge_part)? else {
            continue;
        };
        let Some(size) = numbers_first_bytes(range, 2, merge_part)? else {
            continue;
        };
        let Some(origin) = keynote_fixed32_field(origin, 1, merge_part)?.map(f32::to_bits) else {
            continue;
        };
        let Some(size) = keynote_fixed32_field(size, 1, merge_part)?.map(f32::to_bits) else {
            continue;
        };
        let row = origin & 0xffff;
        let column = origin >> 16;
        let row_count = size & 0xffff;
        let column_count = size >> 16;
        let Some(end_row) = row.checked_add(row_count) else {
            continue;
        };
        let Some(end_column) = column.checked_add(column_count) else {
            continue;
        };
        if row_count == 0 || column_count == 0 || end_row > rows || end_column > columns {
            continue;
        }
        if merges.len() >= limits.max_document_objects {
            return Err(iwa_error(
                merge_part,
                "Numbers merge regions exceed the configured object limit",
            ));
        }
        merges.insert((row, column), (end_row, end_column));
    }
    Ok(merges)
}

fn numbers_cell_text(
    bytes: &[u8],
    strings: &[Option<IworkCellText>],
    rich_strings: &[Option<IworkCellText>],
    part: &str,
) -> Result<Option<IworkCellText>, Diagnostic> {
    let version = *bytes
        .first()
        .ok_or_else(|| iwa_error(part, "Numbers cell storage is empty"))?;
    if version == 5 {
        return numbers_new_cell_text(bytes, strings, rich_strings, part);
    }
    if version > 4 {
        return Err(iwa_error(
            part,
            format!("unsupported Numbers cell storage version {version}"),
        ));
    }
    let cell_type = *bytes
        .get(if version == 4 { 1 } else { 2 })
        .ok_or_else(|| iwa_error(part, "Numbers old cell header is truncated"))?;
    let flags = numbers_le_u32(bytes, 4, part)?;
    let mut offset = if version > 1 { 12 } else { 8 };
    if flags & 0x0002 != 0 {
        offset += 4;
    }
    offset += (flags & if version > 1 { 0x0d8c } else { 0x018c }).count_ones() as usize * 4;
    let rich_index = if flags & 0x0200 != 0 {
        let value = numbers_le_u32(bytes, offset, part)?;
        offset += 4;
        Some(value)
    } else {
        None
    };
    offset += (flags & if version > 1 { 0x3000 } else { 0x1000 }).count_ones() as usize * 4;
    let string_index = if flags & 0x0010 != 0 {
        let value = numbers_le_u32(bytes, offset, part)?;
        offset += 4;
        Some(value)
    } else {
        None
    };
    let number = if flags & 0x0020 != 0 {
        let value = numbers_le_f64(bytes, offset, part)?;
        offset += 8;
        Some(value)
    } else {
        None
    };
    let date = if flags & 0x0040 != 0 {
        Some(numbers_le_f64(bytes, offset, part)?)
    } else {
        None
    };
    numbers_typed_cell_text(
        cell_type,
        string_index,
        rich_index,
        number,
        None,
        date,
        strings,
        rich_strings,
        part,
    )
}

fn numbers_new_cell_text(
    bytes: &[u8],
    strings: &[Option<IworkCellText>],
    rich_strings: &[Option<IworkCellText>],
    part: &str,
) -> Result<Option<IworkCellText>, Diagnostic> {
    let cell_type = *bytes
        .get(1)
        .ok_or_else(|| iwa_error(part, "Numbers new cell header is truncated"))?;
    let fields = numbers_le_u32(bytes, 8, part)?;
    let mut offset = 12_usize;
    let decimal = if fields & 0x0001 != 0 {
        let value = numbers_decimal128(bytes, offset, part)?;
        offset += 16;
        Some(value)
    } else {
        None
    };
    let number = if fields & 0x0002 != 0 {
        let value = numbers_le_f64(bytes, offset, part)?;
        offset += 8;
        Some(value)
    } else {
        None
    };
    let date = if fields & 0x0004 != 0 {
        let value = numbers_le_f64(bytes, offset, part)?;
        offset += 8;
        Some(value)
    } else {
        None
    };
    let string_index = if fields & 0x0008 != 0 {
        let value = numbers_le_u32(bytes, offset, part)?;
        offset += 4;
        Some(value)
    } else {
        None
    };
    let rich_index = if fields & 0x0010 != 0 {
        Some(numbers_le_u32(bytes, offset, part)?)
    } else {
        None
    };
    numbers_typed_cell_text(
        cell_type,
        string_index,
        rich_index,
        number,
        decimal,
        date,
        strings,
        rich_strings,
        part,
    )
}

fn numbers_cell_style_indexes(
    bytes: &[u8],
    part: &str,
) -> Result<(Option<u32>, Option<u32>), Diagnostic> {
    if bytes.first().copied() != Some(5) {
        return Ok((None, None));
    }
    let fields = numbers_le_u32(bytes, 8, part)?;
    let mut offset = 12_usize;
    for (flag, width) in [(0x0001_u32, 16_usize), (0x0002, 8), (0x0004, 8)] {
        if fields & flag != 0 {
            offset = offset
                .checked_add(width)
                .ok_or_else(|| iwa_error(part, "Numbers cell style offset overflows"))?;
        }
    }
    for flag in [0x0008_u32, 0x0010] {
        if fields & flag != 0 {
            offset = offset
                .checked_add(4)
                .ok_or_else(|| iwa_error(part, "Numbers cell style offset overflows"))?;
        }
    }
    let cell_style = if fields & 0x0020 != 0 {
        let value = numbers_le_u32(bytes, offset, part)?;
        offset = offset
            .checked_add(4)
            .ok_or_else(|| iwa_error(part, "Numbers text style offset overflows"))?;
        Some(value)
    } else {
        None
    };
    let text_style = if fields & 0x0040 != 0 {
        Some(numbers_le_u32(bytes, offset, part)?)
    } else {
        None
    };
    Ok((cell_style, text_style))
}

#[allow(clippy::too_many_arguments)]
fn numbers_typed_cell_text(
    cell_type: u8,
    string_index: Option<u32>,
    rich_index: Option<u32>,
    number: Option<f64>,
    decimal: Option<f64>,
    date: Option<f64>,
    strings: &[Option<IworkCellText>],
    rich_strings: &[Option<IworkCellText>],
    part: &str,
) -> Result<Option<IworkCellText>, Diagnostic> {
    let list_value = |values: &[Option<IworkCellText>], index: Option<u32>| {
        index
            .and_then(|index| usize::try_from(index).ok())
            .and_then(|index| values.get(index))
            .and_then(Clone::clone)
    };
    let value = match cell_type {
        0 | 1 => None,
        2 | 10 => decimal.or(number).map(numbers_format_number),
        3 => return Ok(list_value(strings, string_index)),
        5 => date.map(numbers_format_date),
        6 => Some(if number.unwrap_or(0.0) > 0.0 {
            "TRUE".to_owned()
        } else {
            "FALSE".to_owned()
        }),
        7 => number.map(numbers_format_number),
        8 => Some("#ERROR!".to_owned()),
        9 => return Ok(list_value(rich_strings, rich_index)),
        _ => {
            return Err(iwa_error(
                part,
                format!("unsupported Numbers cell type {cell_type}"),
            ));
        }
    };
    Ok(value.map(|text| IworkCellText {
        text,
        storage: None,
    }))
}

fn numbers_decimal128(bytes: &[u8], offset: usize, part: &str) -> Result<f64, Diagnostic> {
    let value = bytes
        .get(offset..offset + 16)
        .ok_or_else(|| iwa_error(part, "Numbers decimal128 value is truncated"))?;
    let exponent = (u16::from(value[15] & 0x7f) << 7) | u16::from(value[14] >> 1);
    let mut mantissa = f64::from(value[14] & 1);
    for byte in value[..14].iter().rev() {
        mantissa = mantissa * 256.0 + f64::from(*byte);
    }
    let sign = if value[15] & 0x80 != 0 { -1.0 } else { 1.0 };
    let decoded = sign * mantissa * 10_f64.powi(i32::from(exponent) - 0x1820);
    if !decoded.is_finite() {
        return Err(iwa_error(part, "Numbers decimal128 value is not finite"));
    }
    Ok(decoded)
}

fn numbers_format_number(value: f64) -> String {
    if value.fract().abs() < f64::EPSILON {
        format!("{value:.0}")
    } else {
        let text = format!("{value:.15}");
        text.trim_end_matches('0').trim_end_matches('.').to_owned()
    }
}

fn numbers_format_date(seconds_since_2001: f64) -> String {
    let days_since_unix = (seconds_since_2001 / 86_400.0).floor() as i64 + 11_323;
    let z = days_since_unix + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let day_of_era = z - era * 146_097;
    let year_of_era =
        (day_of_era - day_of_era / 1_460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let mut year = year_of_era + era * 400;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_prime = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_prime + 2) / 5 + 1;
    let month = month_prime + if month_prime < 10 { 3 } else { -9 };
    year += i64::from(month <= 2);
    format!("{year:04}-{month:02}-{day:02}")
}

fn numbers_cell_address(row: u32, column: u32) -> String {
    let mut value = column + 1;
    let mut column_name = String::new();
    while value > 0 {
        let remainder = (value - 1) % 26;
        column_name.insert(0, char::from(b'A' + remainder as u8));
        value = (value - 1) / 26;
    }
    format!("{column_name}{}", row + 1)
}

fn numbers_le_u32(bytes: &[u8], offset: usize, part: &str) -> Result<u32, Diagnostic> {
    let value: [u8; 4] = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| iwa_error(part, "Numbers cell u32 value is truncated"))?
        .try_into()
        .map_err(|_| iwa_error(part, "Numbers cell u32 value is invalid"))?;
    Ok(u32::from_le_bytes(value))
}

fn numbers_le_f64(bytes: &[u8], offset: usize, part: &str) -> Result<f64, Diagnostic> {
    let value: [u8; 8] = bytes
        .get(offset..offset + 8)
        .ok_or_else(|| iwa_error(part, "Numbers cell f64 value is truncated"))?
        .try_into()
        .map_err(|_| iwa_error(part, "Numbers cell f64 value is invalid"))?;
    let value = f64::from_le_bytes(value);
    if !value.is_finite() {
        return Err(iwa_error(part, "Numbers cell f64 value is not finite"));
    }
    Ok(value)
}

fn numbers_references(bytes: &[u8], field: u64, part: &str) -> Result<Vec<u64>, Diagnostic> {
    numbers_bytes(bytes, field, part)?
        .into_iter()
        .map(parse_iwa_reference)
        .collect()
}

fn numbers_varint(bytes: &[u8], field: u64, part: &str) -> Result<Option<u64>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(part, "Numbers protobuf contains too many fields"));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Numbers protobuf key: {message}"))
        })?;
        cursor += consumed;
        if key >> 3 == field && key & 7 == 0 {
            let (value, _) = read_varint(&bytes[cursor..]).map_err(|message| {
                iwa_error(part, format!("invalid Numbers protobuf varint: {message}"))
            })?;
            return Ok(Some(value));
        }
        skip_protobuf_value(bytes, &mut cursor, key & 7, part)?;
    }
    Ok(None)
}

fn numbers_first_bytes<'a>(
    bytes: &'a [u8],
    field: u64,
    part: &str,
) -> Result<Option<&'a [u8]>, Diagnostic> {
    Ok(numbers_bytes(bytes, field, part)?.into_iter().next())
}

fn numbers_bytes<'a>(bytes: &'a [u8], field: u64, part: &str) -> Result<Vec<&'a [u8]>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut fields = 0_usize;
    let mut values = Vec::new();
    while cursor < bytes.len() {
        fields += 1;
        if fields > Limits::HARD_MAX.max_document_objects {
            return Err(iwa_error(part, "Numbers protobuf contains too many fields"));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Numbers protobuf key: {message}"))
        })?;
        cursor += consumed;
        if key >> 3 == field && key & 7 == 2 {
            values.push(read_iwa_length_delimited(
                bytes,
                &mut cursor,
                part,
                "Numbers protobuf field",
            )?);
        } else {
            skip_protobuf_value(bytes, &mut cursor, key & 7, part)?;
        }
    }
    Ok(values)
}

fn pages_document(
    package: &Package<'_>,
    document_archives: &[IwaArchive],
    limits: Limits,
    font_metrics: &FontMetricTable,
) -> Result<Option<Document>, Diagnostic> {
    let root = root_message(document_archives, DOCUMENT_COMPONENT)?;
    let floating_drawables = numbers_references(&root.payload, 3, DOCUMENT_COMPONENT)?
        .into_iter()
        .next()
        .map(|identifier| pages_floating_drawables(document_archives, identifier))
        .transpose()?
        .unwrap_or_default();
    let Some(layout) = pages_document_layout(&root.payload, font_metrics)? else {
        return Ok(None);
    };
    let Some(body) = archive_message(
        document_archives,
        layout.body_storage,
        IWORK_TEXT_STORAGE_TYPE,
        DOCUMENT_COMPONENT,
    )?
    else {
        return Ok(None);
    };
    let body = keynote_storage_text(&body.payload, DOCUMENT_COMPONENT, limits)?;
    let (page_hints, view_scale) = pages_layout_hints(package, limits)?;
    if page_hints.is_empty() {
        return Ok(None);
    }
    let layout = PagesDocumentLayout {
        view_scale,
        ..layout
    };

    let mut components: Vec<String> = package
        .entry_names()
        .filter(|part| {
            matches!(
                *part,
                STYLESHEET_COMPONENT | THEME_STYLESHEET_COMPONENT | CALCULATION_ENGINE_COMPONENT
            ) || iwork_object_container_component(part)
                || (part.starts_with("Index/Tables/") && part.ends_with(".iwa"))
        })
        .map(str::to_owned)
        .collect();
    components.sort();
    components.dedup();
    let mut archives = Vec::new();
    let mut message_archives = document_archives
        .iter()
        .cloned()
        .map(|archive| NumbersArchive {
            part: DOCUMENT_COMPONENT.to_owned(),
            archive,
        })
        .collect::<Vec<_>>();
    for component in components {
        let bytes = package.required_part(&component)?;
        let component_archives = parse_iwa_archives(package, &component, &bytes, limits)?;
        archives.extend(component_archives.iter().cloned());
        message_archives.extend(
            component_archives
                .into_iter()
                .map(|archive| NumbersArchive {
                    part: component.clone(),
                    archive,
                }),
        );
    }
    archives.extend_from_slice(document_archives);
    let messages = NumbersMessageSpace {
        archives: message_archives,
    };
    let data_files = iwork_data_files(package, limits)?;

    let mut diagnostics = Vec::new();
    for id in body
        .drop_cap_styles
        .iter()
        .filter_map(|change| change.identifier)
    {
        if let Err(diagnostic) = pages_drop_cap(&archives, id, DOCUMENT_COMPONENT) {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Parse,
                    Fidelity::Approximate,
                    diagnostic.message,
                )
                .in_part(DOCUMENT_COMPONENT),
            );
        }
    }
    let mut units = Vec::new();
    let mut objects = Vec::new();
    let mut embedded_fonts = Vec::new();
    units.try_reserve_exact(page_hints.len()).map_err(|_| {
        Diagnostic::fatal(
            DiagnosticCode::AllocationFailed,
            Phase::Parse,
            None,
            "unable to allocate the bounded Pages page list",
        )
        .in_part(DOCUMENT_COMPONENT)
    })?;
    for (index, page) in page_hints.iter().enumerate() {
        diagnostics.extend(page.flow_diagnostics.iter().cloned());
        let unit_index = u32::try_from(index)
            .map_err(|_| iwa_error(DOCUMENT_COMPONENT, "Pages page index overflows"))?;
        let target_bottom = page
            .targets
            .iter()
            .filter_map(|target| {
                target
                    .origin
                    .zip(target.size)
                    .map(|((_, y), (_, height))| y + height)
            })
            .fold(0.0_f32, f32::max);
        let mut layout = PagesDocumentLayout {
            margin_top: if target_bottom > 0.0 {
                layout
                    .margin_top
                    .max((layout.height - layout.margin_bottom - target_bottom).max(0.0))
            } else {
                layout.margin_top
            },
            ..layout
        };
        if let Some(header_bottom) = pages_header_objects(
            &mut diagnostics,
            package,
            &archives,
            &messages,
            &data_files,
            &body,
            page,
            unit_index,
            layout,
            limits,
            &mut objects,
            &mut embedded_fonts,
        )? {
            layout.margin_top = layout.margin_top.max(header_bottom + 12.0);
        }
        units.push(Unit {
            kind: UnitKind::Page,
            index: unit_index,
            id: format!("pages-page-{}", index + 1),
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
        });

        for (target_index, target) in page.targets.iter().enumerate() {
            let runs = pages_text_runs_for_range(
                &body,
                &archives,
                DOCUMENT_COMPONENT,
                target.range_start,
                target.range_end,
                limits,
            )?;
            let text = runs.iter().map(|run| run.text.as_str()).collect::<String>();
            if text.is_empty() {
                continue;
            }
            let paragraphs = pages_paragraph_layouts(
                &body,
                &runs,
                &archives,
                DOCUMENT_COMPONENT,
                target.range_start,
                target.range_end,
                &messages,
                limits,
                layout.view_scale,
                layout.font_metrics,
                false,
            )?;
            let fallback_y = if target.anchored_start < target.range_start
                && page
                    .attachment_positions
                    .iter()
                    .any(|(_, y)| *y <= layout.margin_top + 1.0)
            {
                page.attachment_positions
                    .iter()
                    .map(|(_, y)| *y)
                    .fold(layout.margin_top, f32::max)
                    + 321.15
            } else {
                layout.margin_top
            };
            let (target_width, target_height) = target.size.unwrap_or((
                layout.width - layout.margin_left - layout.margin_right,
                layout.height - fallback_y - layout.margin_bottom,
            ));
            let (target_x, target_y) = target
                .origin
                .map_or((layout.margin_left, fallback_y), |(x, y)| {
                    (layout.margin_left + x, layout.margin_top + y)
                });
            pages_push_object(
                &mut objects,
                limits,
                unit_index,
                layout.body_storage,
                ObjectKind::TextBox,
                Rect {
                    x: target_x,
                    y: target_y,
                    width: target_width.max(1.0),
                    height: target_height.max(1.0),
                },
                10,
                Some(text),
                "body",
                MappingQuality::Exact,
                Visual::TextLayout {
                    layout: TextLayout {
                        vertical_align: TextVerticalAlign::Top,
                        column_count: target.column_count,
                        column_spacing: if target.column_count > 1 { 36.0 } else { 0.0 },
                        // As with linked frames, the saved range already selects
                        // this page's text; approximate font boxes must not clip it.
                        inset_left: 0.0,
                        inset_right: 0.0,
                        inset_top: 0.0,
                        inset_bottom: 0.0,
                        paragraphs,
                        ..TextLayout::default()
                    },
                    visual: Box::new(Visual::RichText {
                        geometry: Geometry::Rectangle,
                        fill: Paint::None,
                        stroke: Paint::None,
                        stroke_width: 0.0,
                        align: TextAlign::Start,
                        line_height: 0.0,
                        runs,
                    }),
                },
            )?;
            if let Some(object) = objects.last_mut() {
                object
                    .stable_id
                    .push_str(&format!("-fragment-{target_index}"));
            }
        }

        pages_attachment_objects(
            &mut diagnostics,
            package,
            &archives,
            &messages,
            &data_files,
            &body,
            page,
            index
                .checked_sub(1)
                .and_then(|previous| page_hints.get(previous)),
            unit_index,
            layout,
            limits,
            &mut objects,
            &mut embedded_fonts,
        )?;
        if let Some(drawables) = floating_drawables.get(index) {
            let mut frame_targets = Vec::new();
            for (flow, target) in &page.flow_targets {
                let mut frames = Vec::new();
                for identifier in drawables {
                    if let Some(shape) = archive_message(
                        &archives,
                        *identifier,
                        IWORK_TEXT_SHAPE_TYPE,
                        DOCUMENT_COMPONENT,
                    )? && keynote_reference_field(&shape.payload, 3, DOCUMENT_COMPONENT)?
                        == Some(*flow)
                    {
                        frames.push(*identifier);
                    }
                }
                if frames.len() == 1 {
                    frame_targets.push((frames[0], target.clone()));
                } else {
                    let result = pages_flow_frame_targets(
                        &archives, &messages, *flow, target, &frames, layout, limits,
                    );
                    if let Ok(targets) = &result {
                        frame_targets.extend(targets.iter().cloned());
                    }
                    diagnostics.push(Diagnostic::warning(DiagnosticCode::UnsupportedFeature,
                        Phase::Layout, Fidelity::Approximate,
                        match result {
                            Ok(_) => "PAGES_TEXT_FLOW_APPROXIMATE: same-page frames use measured text allocation".to_owned(),
                            Err(error) => format!("PAGES_TEXT_FLOW_APPROXIMATE: {}", error.message),
                        },
                    ).in_part(DOCUMENT_COMPONENT));
                }
            }
            let layout = PagesDocumentLayout {
                frame_targets: &frame_targets,
                ..layout
            };
            let mut path = Vec::new();
            let first_object = objects.len();
            let mut wrap_regions = Vec::new();
            for (z, identifier) in drawables.iter().copied().enumerate() {
                keynote_native_drawable(
                    &mut diagnostics,
                    package,
                    &archives,
                    &archives,
                    &messages,
                    &data_files,
                    identifier,
                    unit_index,
                    DOCUMENT_COMPONENT,
                    limits,
                    &mut objects,
                    &mut path,
                    layout.width,
                    layout.height,
                    0.0,
                    0.0,
                    i32::try_from(z).unwrap_or(i32::MAX),
                    true,
                    &mut embedded_fonts,
                    IWORK_POINT_TO_CSS_PIXEL,
                    Some(layout),
                )?;
                // Reuse the materialized crop bounds even when the image has no bytes.
                if let Some(object) = objects.last().filter(|object| {
                    object.stable_id
                        == format!("iwork-keynote-{unit_index}-Index/Document.iwa-{identifier}")
                }) {
                    match pages_drawable_wrap_region(&archives, identifier, object.bounds) {
                        Ok(Some(region)) => wrap_regions.push((object.stable_id.clone(), region)),
                        Ok(None) => {}
                        Err(error) => diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::UnsupportedFeature,
                                Phase::Layout,
                                Fidelity::Approximate,
                                format!("PAGES_TEXT_WRAP_UNAVAILABLE: {}", error.message),
                            )
                            .in_part(DOCUMENT_COMPONENT),
                        ),
                    }
                }
            }
            pages_apply_text_wrap(
                &mut objects[first_object..],
                &wrap_regions,
                limits,
                &mut diagnostics,
            );
            if !wrap_regions.is_empty() {
                diagnostics.push(Diagnostic::warning(DiagnosticCode::UnsupportedFeature, Phase::Layout,
                    Fidelity::Approximate, "PAGES_TEXT_WRAP_APPROXIMATE: floating objects use rectangular exclusions; contour and directional wrapping may be approximated")
                    .in_part(DOCUMENT_COMPONENT));
            }
        }
    }

    diagnostics.push(
        Diagnostic::warning(
            DiagnosticCode::UnsupportedFeature,
            Phase::Render,
            Fidelity::Approximate,
            "PAGES_NATIVE_STATIC: rendered saved pagination, styled body text, and supported inline images; advanced tables and drawing effects may be approximated",
        ).in_part(DOCUMENT_COMPONENT),
    );
    let mut document = Document {
        fatal: false,
        format: Some(DocumentFormat::Pages),
        kind: Some(DocumentKind::Text),
        units,
        outline: Vec::new(),
        objects,
        embedded_fonts,
        font_alternate_names: Vec::new(),
        diagnostics,
    };
    scale_iwork_point_document(&mut document);
    Ok(Some(document))
}

fn pages_drawable_wrap_region(
    archives: &[IwaArchive],
    identifier: u64,
    bounds: Rect,
) -> Result<Option<Rect>, Diagnostic> {
    let part = DOCUMENT_COMPONENT;
    let drawable = if let Some(image) =
        archive_message(archives, identifier, IWORK_IMAGE_TYPE, part)?
    {
        keynote_nested_message(&image.payload, 1, part)?
    } else if let Some(text) = archive_message(archives, identifier, IWORK_TEXT_SHAPE_TYPE, part)? {
        match keynote_nested_message(&text.payload, 1, part)? {
            Some(shape) => keynote_nested_message(shape, 1, part)?,
            None => None,
        }
    } else {
        None
    };
    let Some(drawable) = drawable else {
        return Ok(None);
    };
    let Some(wrap) = keynote_nested_message(drawable, 3, part)? else {
        return Ok(None);
    };
    if keynote_varint_field(wrap, 1, part)? != Some(1) {
        return Ok(None);
    }
    if keynote_drawable_geometry(drawable, part)?
        .rotation_degrees
        .abs()
        > f32::EPSILON
    {
        return Err(iwa_error(part, "rotated text wrap is not supported"));
    }
    let margin = keynote_fixed32_field(wrap, 4, part)?
        .unwrap_or(0.0)
        .max(0.0);
    Ok(Some(Rect {
        x: bounds.x - margin,
        y: bounds.y - margin,
        width: bounds.width + 2.0 * margin,
        height: bounds.height + 2.0 * margin,
    }))
}

fn pages_apply_text_wrap(
    objects: &mut [Object],
    regions: &[(String, Rect)],
    limits: Limits,
    diagnostics: &mut Vec<Diagnostic>,
) {
    let mut checks = 0;
    for object in objects {
        let Visual::TextLayout { layout, .. } = &mut object.visual else {
            continue;
        };
        if layout.vertical_align != TextVerticalAlign::Top || layout.rotation_degrees != 0.0 {
            continue;
        }
        let bounds = object.bounds;
        for (owner, region) in regions {
            if checks >= limits.max_relationship_edges || layout.wrap_regions.len() >= 1024 {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Layout,
                        Fidelity::Approximate,
                        "PAGES_TEXT_WRAP_LIMIT: remaining exclusions were skipped",
                    )
                    .in_part(DOCUMENT_COMPONENT),
                );
                return;
            }
            checks += 1;
            if owner != &object.stable_id
                && region.x < bounds.x + bounds.width
                && region.x + region.width > bounds.x
                && region.y < bounds.y + bounds.height
                && region.y + region.height > bounds.y
            {
                layout.wrap_regions.push(Rect {
                    x: region.x - bounds.x - layout.inset_left - layout.margin_left,
                    y: region.y - bounds.y - layout.inset_top,
                    ..*region
                });
            }
        }
    }
}

fn pages_floating_drawables(
    archives: &[IwaArchive],
    identifier: u64,
) -> Result<Vec<Vec<u64>>, Diagnostic> {
    let Some(message) = archive_message(
        archives,
        identifier,
        PAGES_FLOATING_DRAWABLES_TYPE,
        DOCUMENT_COMPONENT,
    )?
    else {
        return Ok(Vec::new());
    };
    let mut pages = Vec::<Vec<u64>>::new();
    for group in numbers_bytes(&message.payload, 1, DOCUMENT_COMPONENT)? {
        let page_index = numbers_varint(group, 1, DOCUMENT_COMPONENT)?
            .and_then(|value| usize::try_from(value).ok())
            .ok_or_else(|| iwa_error(DOCUMENT_COMPONENT, "Pages drawable page index is invalid"))?;
        if page_index >= pages.len() {
            pages.resize_with(page_index + 1, Vec::new);
        }
        for field in [2, 4, 3] {
            for entry in numbers_bytes(group, field, DOCUMENT_COMPONENT)? {
                let Some(reference) = numbers_first_bytes(entry, 1, DOCUMENT_COMPONENT)? else {
                    continue;
                };
                let drawable = parse_iwa_reference(reference)?;
                if !pages[page_index].contains(&drawable) {
                    pages[page_index].push(drawable);
                }
            }
        }
    }
    Ok(pages)
}

#[allow(clippy::too_many_arguments)]
fn pages_flow_frame_targets(
    archives: &[IwaArchive],
    messages: &NumbersMessageSpace,
    flow: u64,
    target: &PagesTargetHint,
    frames: &[u64],
    layout: PagesDocumentLayout,
    limits: Limits,
) -> Result<Vec<(u64, PagesTargetHint)>, Diagnostic> {
    let part = DOCUMENT_COMPONENT;
    if frames.is_empty() || frames.len() > 64 || target.column_count as usize != frames.len() {
        return Err(iwa_error(
            part,
            "text flow has no unambiguous top-level frame sequence",
        ));
    }
    let flow = archive_message(archives, flow, IWORK_TEXT_FLOW_TYPE, part)?
        .ok_or_else(|| iwa_error(part, "text flow is unavailable"))?;
    let ordered = numbers_references(&flow.payload, 2, part)?
        .into_iter()
        .filter(|identifier| frames.contains(identifier))
        .collect::<Vec<_>>();
    if ordered.len() != frames.len() || ordered.iter().collect::<HashSet<_>>().len() != frames.len()
    {
        return Err(iwa_error(part, "text flow frame order is incomplete"));
    }
    let storage_id = keynote_reference_field(&flow.payload, 1, part)?
        .ok_or_else(|| iwa_error(part, "text flow has no shared story"))?;
    let storage = archive_message(archives, storage_id, IWORK_TEXT_STORAGE_TYPE, part)?
        .ok_or_else(|| iwa_error(part, "shared story is unavailable"))?;
    let storage = keynote_storage_text(&storage.payload, part, limits)?;
    let end_byte = keynote_utf16_byte_offset(&storage.text, target.range_end)
        .ok_or_else(|| iwa_error(part, "text flow range exceeds the story"))?;
    let mut start = target.range_start;
    let mut result = Vec::new();
    for (index, frame) in ordered.iter().copied().enumerate() {
        let mut end = start;
        if index + 1 == ordered.len() {
            // The saved page boundary remains authoritative, including overflow.
            end = target.range_end;
        } else {
            let shape = archive_message(archives, frame, IWORK_TEXT_SHAPE_TYPE, part)?
                .ok_or_else(|| iwa_error(part, "text flow frame is unavailable"))?;
            let (geometry, _, style_id, _, _, _) = keynote_text_shape(&shape.payload, part)?;
            let box_layout = style_id
                .map(|id| keynote_text_layout_style(archives, id, part))
                .transpose()?
                .flatten()
                .unwrap_or_default();
            let width = geometry.width - box_layout.inset_left - box_layout.inset_right;
            let height = geometry.height - box_layout.inset_top - box_layout.inset_bottom;
            if !width.is_finite()
                || !height.is_finite()
                || width <= 0.0
                || height <= 0.0
                || box_layout.column_count > 1
            {
                return Err(iwa_error(
                    part,
                    "text flow frame geometry needs an unsupported layout",
                ));
            }
            let start_byte = keynote_utf16_byte_offset(&storage.text, start)
                .filter(|byte| *byte <= end_byte)
                .ok_or_else(|| iwa_error(part, "text flow start is invalid"))?;
            let mut used = 0.0;
            let mut previous_after = 0.0;
            for text in storage.text[start_byte..end_byte].split_inclusive(['\n', '\u{2028}']) {
                let paragraph_end = end + text.encode_utf16().count();
                let runs = pages_text_runs_for_range(
                    &storage,
                    archives,
                    part,
                    end,
                    paragraph_end,
                    limits,
                )?;
                // ponytail: dynamic fields/case expansions need a source-to-glyph
                // index map; never guess a UTF-16 boundary for those stories.
                if runs
                    .iter()
                    .map(|run| run.text.encode_utf16().count())
                    .sum::<usize>()
                    != paragraph_end - end
                    || text.contains('\t')
                {
                    return Err(iwa_error(
                        part,
                        "text flow needs a source-to-display range map",
                    ));
                }
                let paragraphs = pages_paragraph_layouts(
                    &storage,
                    &runs,
                    archives,
                    part,
                    end,
                    paragraph_end,
                    messages,
                    limits,
                    layout.view_scale,
                    layout.font_metrics,
                    true,
                )?;
                let paragraph = &paragraphs[0];
                let style = iwork_inherited_style_at(&storage.paragraph_styles, end)
                    .map(|id| keynote_text_style(archives, id, part))
                    .transpose()?
                    .unwrap_or_default();
                let characters = runs
                    .iter()
                    .flat_map(|run| {
                        run.text.chars().map(|character| {
                            let advance = layout
                                .font_metrics
                                .advance_em_at_size(
                                    &run.font_family,
                                    run.italic,
                                    run.bold,
                                    character,
                                    run.font_size * IWORK_POINT_TO_CSS_PIXEL,
                                )
                                .map_or_else(
                                    || {
                                        keynote_character_advance(
                                            character,
                                            run.font_size,
                                            run.letter_spacing,
                                        )
                                    },
                                    |em| em * run.font_size + run.letter_spacing,
                                );
                            (
                                if character == '\u{2028}' {
                                    '\n'
                                } else {
                                    character
                                },
                                advance.max(0.0),
                            )
                        })
                    })
                    .collect::<Vec<_>>();
                let available_width = width - paragraph.margin_left - paragraph.margin_right;
                let breaks = crate::text_layout::horizontal_line_breaks(
                    &characters,
                    available_width - paragraph.first_line_indent,
                    available_width,
                    paragraph.default_tab_stop,
                    false,
                );
                let before = if end == start {
                    0.0
                } else {
                    super::collapsed_paragraph_space_before(paragraph.space_before, previous_after)
                };
                let mut fits = ((height - used - before).max(0.0) / paragraph.line_height.max(0.1))
                    .floor() as usize;
                if fits < breaks.len() {
                    if end > start
                        && (style.keep_lines_together || (style.widow_control && fits < 2))
                    {
                        break;
                    }
                    if style.widow_control && breaks.len() - fits == 1 && fits > 1 {
                        fits -= 1;
                    }
                    if end > start && style.widow_control && fits < 2 {
                        break;
                    }
                    // An over-height first line must still make forward progress.
                    let count = breaks[fits.max(1).min(breaks.len()) - 1];
                    end += characters[..count]
                        .iter()
                        .map(|(character, _)| character.len_utf16())
                        .sum::<usize>();
                    break;
                }
                used +=
                    before + breaks.len() as f32 * paragraph.line_height + paragraph.space_after;
                previous_after = paragraph.space_after;
                end = paragraph_end;
            }
        }
        result.push((
            frame,
            PagesTargetHint {
                range_start: start,
                range_end: end,
                ..target.clone()
            },
        ));
        start = end;
    }
    Ok(result)
}

#[allow(clippy::too_many_arguments)]
fn pages_header_objects(
    diagnostics: &mut Vec<Diagnostic>,
    package: &Package<'_>,
    archives: &[IwaArchive],
    messages: &NumbersMessageSpace,
    data_files: &[(u64, String)],
    body: &KeynoteTextStorage,
    page: &PagesPageHint,
    unit_index: u32,
    layout: PagesDocumentLayout,
    limits: Limits,
    objects: &mut Vec<Object>,
    embedded_fonts: &mut Vec<EmbeddedFont>,
) -> Result<Option<f32>, Diagnostic> {
    let Some(page_start) = page.targets.first().map(|target| target.range_start) else {
        return Ok(None);
    };
    let Some(section) = body
        .sections
        .iter()
        .rev()
        .find(|change| change.character_index <= page_start && change.identifier.is_some())
    else {
        return Ok(None);
    };
    let section_identifier = section.identifier.unwrap();
    let Some(section_message) = archive_message(
        archives,
        section_identifier,
        PAGES_SECTION_TYPE,
        DOCUMENT_COMPONENT,
    )?
    else {
        return Ok(None);
    };
    let first_page = page.targets.iter().any(|target| {
        target.range_start <= section.character_index && section.character_index < target.range_end
    });
    let different_first = numbers_varint(&section_message.payload, 18, DOCUMENT_COMPONENT)?
        .is_some_and(|value| value != 0);
    let different_even_odd = numbers_varint(&section_message.payload, 19, DOCUMENT_COMPONENT)?
        .is_some_and(|value| value != 0);
    let template_field = if first_page && different_first {
        23
    } else if different_even_odd && (unit_index + 1).is_multiple_of(2) {
        24
    } else {
        25
    };
    let Some(template_identifier) =
        numbers_references(&section_message.payload, template_field, DOCUMENT_COMPONENT)?
            .into_iter()
            .next()
    else {
        return Ok(None);
    };
    let Some(template) = archive_message(
        archives,
        template_identifier,
        PAGES_SECTION_TEMPLATE_TYPE,
        DOCUMENT_COMPONENT,
    )?
    else {
        return Ok(None);
    };
    let mut path = Vec::new();
    for (z, identifier) in numbers_references(&template.payload, 3, DOCUMENT_COMPONENT)?
        .into_iter()
        .enumerate()
    {
        keynote_native_drawable(
            diagnostics,
            package,
            archives,
            archives,
            messages,
            data_files,
            identifier,
            unit_index,
            DOCUMENT_COMPONENT,
            limits,
            objects,
            &mut path,
            layout.width,
            layout.height,
            0.0,
            0.0,
            i32::MIN.saturating_add(i32::try_from(z).unwrap_or(i32::MAX)),
            true,
            embedded_fonts,
            IWORK_POINT_TO_CSS_PIXEL,
            None,
        )?;
    }
    // Section artwork stays local even when its header/footer text is linked.
    let mut header_template = template;
    let mut inherited_section = section_message;
    for previous in body.sections.iter().rev().filter(|previous| {
        previous.character_index < section.character_index && previous.identifier.is_some()
    }) {
        if numbers_varint(&inherited_section.payload, 17, DOCUMENT_COMPONENT)? != Some(1) {
            break;
        }
        let Some(previous_section) = archive_message(
            archives,
            previous.identifier.unwrap(),
            PAGES_SECTION_TYPE,
            DOCUMENT_COMPONENT,
        )?
        else {
            break;
        };
        inherited_section = previous_section;
        if let Some(identifier) = keynote_reference_field(
            &previous_section.payload,
            template_field,
            DOCUMENT_COMPONENT,
        )? && let Some(previous_template) = archive_message(
            archives,
            identifier,
            PAGES_SECTION_TEMPLATE_TYPE,
            DOCUMENT_COMPONENT,
        )? {
            header_template = previous_template;
        }
    }
    let header_y = layout.header_distance;
    let content_width = (layout.width - layout.margin_left - layout.margin_right).max(1.0);
    let mut header_bottom = None::<f32>;
    let page_number = (unit_index + 1).to_string();
    for (field, is_header) in [(1, true), (2, false)] {
        for (slot, storage_identifier) in
            numbers_references(&header_template.payload, field, DOCUMENT_COMPONENT)?
                .into_iter()
                .enumerate()
        {
            let Some(storage_message) = archive_message(
                archives,
                storage_identifier,
                IWORK_TEXT_STORAGE_TYPE,
                DOCUMENT_COMPONENT,
            )?
            else {
                continue;
            };
            let storage =
                keynote_storage_text(&storage_message.payload, DOCUMENT_COMPONENT, limits)?;
            let range_end = storage.text.encode_utf16().count();
            if range_end == 0 {
                continue;
            }
            let runs = pages_text_runs_for_range_with_page_number(
                &storage,
                archives,
                DOCUMENT_COMPONENT,
                0,
                range_end,
                limits,
                Some(&page_number),
            )?;
            let mut paragraphs = pages_paragraph_layouts(
                &storage,
                &runs,
                archives,
                DOCUMENT_COMPONENT,
                0,
                range_end,
                messages,
                limits,
                layout.view_scale,
                layout.font_metrics,
                false,
            )?;
            let mut position = 0;
            for (text, paragraph) in storage
                .text
                .split_inclusive(['\n', '\u{2028}'])
                .zip(&mut paragraphs)
            {
                let style = iwork_inherited_style_at(&storage.paragraph_styles, position)
                    .map(|identifier| keynote_text_style(archives, identifier, DOCUMENT_COMPONENT))
                    .transpose()?
                    .unwrap_or_default();
                if style.natural_alignment {
                    paragraph.align = match slot {
                        1 => TextAlign::Center,
                        2 => TextAlign::End,
                        _ => TextAlign::Start,
                    };
                }
                position += text.encode_utf16().count();
            }
            let text_height =
                keynote_text_layout_height(&runs, &paragraphs, content_width).max(1.0);
            let text_y = if is_header {
                header_y
            } else {
                layout.height - layout.footer_distance - text_height
            };
            let text = runs.iter().map(|run| run.text.as_str()).collect::<String>();
            if !text.is_empty() {
                pages_push_object(
                    objects,
                    limits,
                    unit_index,
                    storage_identifier,
                    ObjectKind::TextBox,
                    Rect {
                        x: layout.margin_left,
                        y: text_y,
                        width: content_width,
                        height: text_height,
                    },
                    20,
                    Some(text),
                    "body",
                    MappingQuality::Exact,
                    Visual::TextLayout {
                        layout: TextLayout {
                            vertical_align: TextVerticalAlign::Top,
                            inset_left: 0.0,
                            inset_right: 0.0,
                            inset_top: 0.0,
                            inset_bottom: 0.0,
                            paragraphs,
                            horizontal_overflow: TextHorizontalOverflow::Clip,
                            vertical_overflow: TextVerticalOverflow::Clip,
                            ..TextLayout::default()
                        },
                        visual: Box::new(Visual::RichText {
                            geometry: Geometry::Rectangle,
                            fill: Paint::None,
                            stroke: Paint::None,
                            stroke_width: 0.0,
                            align: TextAlign::Start,
                            line_height: 0.0,
                            runs,
                        }),
                    },
                )?;
            }
            let header_page = PagesPageHint {
                flow_targets: Vec::new(),
                flow_diagnostics: Vec::new(),
                targets: vec![PagesTargetHint {
                    range_start: 0,
                    range_end,
                    anchored_start: 0,
                    anchored_end: range_end,
                    origin: None,
                    size: Some((content_width, text_height)),
                    column_count: 1,
                }],
                attachment_positions: Vec::new(),
                toc_ranges: Vec::new(),
            };
            let first_object = objects.len();
            pages_attachment_objects(
                diagnostics,
                package,
                archives,
                messages,
                data_files,
                &storage,
                &header_page,
                None,
                unit_index,
                PagesDocumentLayout {
                    margin_top: text_y,
                    ..layout
                },
                limits,
                objects,
                embedded_fonts,
            )?;
            if is_header {
                let content_bottom = objects[first_object..]
                    .iter()
                    .map(|object| object.bounds.y + object.bounds.height)
                    .reduce(f32::max)
                    .unwrap_or(text_y + text_height);
                header_bottom = Some(content_bottom.max(header_bottom.unwrap_or(0.0)));
            }
        }
    }
    Ok(header_bottom)
}

fn scale_iwork_point_document(document: &mut Document) {
    for unit in &mut document.units {
        unit.width *= IWORK_POINT_TO_CSS_PIXEL;
        unit.height *= IWORK_POINT_TO_CSS_PIXEL;
        unit.frozen_width *= IWORK_POINT_TO_CSS_PIXEL;
        unit.frozen_height *= IWORK_POINT_TO_CSS_PIXEL;
    }
    scale_iwork_point_objects(&mut document.objects);
}

fn scale_iwork_point_objects(objects: &mut [Object]) {
    for object in objects {
        let final_embedded_pdf_visual = matches!(
            &object.source.locator,
            SourceLocator::Iwork { component, .. } if component.starts_with("embedded-pdf-archive-")
        );
        if final_embedded_pdf_visual {
            continue;
        }
        object.bounds.x *= IWORK_POINT_TO_CSS_PIXEL;
        object.bounds.y *= IWORK_POINT_TO_CSS_PIXEL;
        object.bounds.width *= IWORK_POINT_TO_CSS_PIXEL;
        object.bounds.height *= IWORK_POINT_TO_CSS_PIXEL;
        scale_iwork_visual(&mut object.visual);
    }
}

fn scale_iwork_path(commands: &mut [PathCommand]) {
    for command in commands {
        match command {
            PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => {
                *x *= IWORK_POINT_TO_CSS_PIXEL;
                *y *= IWORK_POINT_TO_CSS_PIXEL;
            }
            PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => {
                *cpx *= IWORK_POINT_TO_CSS_PIXEL;
                *cpy *= IWORK_POINT_TO_CSS_PIXEL;
                *x *= IWORK_POINT_TO_CSS_PIXEL;
                *y *= IWORK_POINT_TO_CSS_PIXEL;
            }
            PathCommand::BezierCurveTo {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                *cp1x *= IWORK_POINT_TO_CSS_PIXEL;
                *cp1y *= IWORK_POINT_TO_CSS_PIXEL;
                *cp2x *= IWORK_POINT_TO_CSS_PIXEL;
                *cp2y *= IWORK_POINT_TO_CSS_PIXEL;
                *x *= IWORK_POINT_TO_CSS_PIXEL;
                *y *= IWORK_POINT_TO_CSS_PIXEL;
            }
            PathCommand::ClosePath => {}
        }
    }
}

fn scale_iwork_geometry(geometry: &mut Geometry) {
    match geometry {
        Geometry::RoundedRectangle { radius_x, radius_y } => {
            *radius_x *= IWORK_POINT_TO_CSS_PIXEL;
            *radius_y *= IWORK_POINT_TO_CSS_PIXEL;
        }
        Geometry::Path { commands, .. } => scale_iwork_path(commands),
        Geometry::LayeredPath { layers } => {
            for layer in layers {
                scale_iwork_path(&mut layer.commands);
            }
        }
        Geometry::Rectangle | Geometry::Ellipse | Geometry::Line => {}
    }
}

fn scale_iwork_paint(paint: &mut Paint) {
    match paint {
        Paint::MappedGradient { paint, .. } => scale_iwork_paint(paint),
        Paint::LinearGradient { x0, y0, x1, y1, .. } => {
            *x0 *= IWORK_POINT_TO_CSS_PIXEL;
            *y0 *= IWORK_POINT_TO_CSS_PIXEL;
            *x1 *= IWORK_POINT_TO_CSS_PIXEL;
            *y1 *= IWORK_POINT_TO_CSS_PIXEL;
        }
        Paint::RadialGradient {
            x0,
            y0,
            r0,
            x1,
            y1,
            r1,
            ..
        } => {
            *x0 *= IWORK_POINT_TO_CSS_PIXEL;
            *y0 *= IWORK_POINT_TO_CSS_PIXEL;
            *r0 *= IWORK_POINT_TO_CSS_PIXEL;
            *x1 *= IWORK_POINT_TO_CSS_PIXEL;
            *y1 *= IWORK_POINT_TO_CSS_PIXEL;
            *r1 *= IWORK_POINT_TO_CSS_PIXEL;
        }
        Paint::RectGradient {
            center_x, center_y, ..
        } => {
            *center_x *= IWORK_POINT_TO_CSS_PIXEL;
            *center_y *= IWORK_POINT_TO_CSS_PIXEL;
        }
        Paint::CircleGradient {
            x0,
            y0,
            r0,
            x1,
            y1,
            r1,
            ..
        } => {
            *x0 *= IWORK_POINT_TO_CSS_PIXEL;
            *y0 *= IWORK_POINT_TO_CSS_PIXEL;
            *r0 *= IWORK_POINT_TO_CSS_PIXEL;
            *x1 *= IWORK_POINT_TO_CSS_PIXEL;
            *y1 *= IWORK_POINT_TO_CSS_PIXEL;
            *r1 *= IWORK_POINT_TO_CSS_PIXEL;
        }
        Paint::XpsGradient {
            start_x,
            start_y,
            end_x,
            end_y,
            radius_x,
            radius_y,
            relative,
            transform,
            ..
        } => {
            if !*relative {
                *start_x *= IWORK_POINT_TO_CSS_PIXEL;
                *start_y *= IWORK_POINT_TO_CSS_PIXEL;
                *end_x *= IWORK_POINT_TO_CSS_PIXEL;
                *end_y *= IWORK_POINT_TO_CSS_PIXEL;
                *radius_x *= IWORK_POINT_TO_CSS_PIXEL;
                *radius_y *= IWORK_POINT_TO_CSS_PIXEL;
            }
            transform.e *= IWORK_POINT_TO_CSS_PIXEL;
            transform.f *= IWORK_POINT_TO_CSS_PIXEL;
        }
        Paint::None
        | Paint::ShapeGradient { .. }
        | Paint::Solid(_)
        | Paint::Pattern { .. }
        | Paint::Image { .. }
        | Paint::Visual { .. } => {}
    }
}

fn scale_iwork_visual(visual: &mut Visual) {
    match visual {
        Visual::Shape {
            geometry,
            stroke_width,
            ..
        } => {
            scale_iwork_geometry(geometry);
            *stroke_width *= IWORK_POINT_TO_CSS_PIXEL;
        }
        Visual::Text {
            geometry,
            stroke_width,
            font_size,
            ..
        } => {
            scale_iwork_geometry(geometry);
            *stroke_width *= IWORK_POINT_TO_CSS_PIXEL;
            *font_size *= IWORK_POINT_TO_CSS_PIXEL;
        }
        Visual::Layer {
            transform, visual, ..
        } => {
            transform.e *= IWORK_POINT_TO_CSS_PIXEL;
            transform.f *= IWORK_POINT_TO_CSS_PIXEL;
            scale_iwork_visual(visual);
        }
        Visual::PaintedShape {
            geometry,
            fill,
            stroke,
            stroke_width,
        } => {
            scale_iwork_geometry(geometry);
            scale_iwork_paint(fill);
            scale_iwork_paint(stroke);
            *stroke_width *= IWORK_POINT_TO_CSS_PIXEL;
        }
        Visual::RichText {
            geometry,
            fill,
            stroke,
            stroke_width,
            line_height,
            runs,
            ..
        } => {
            scale_iwork_geometry(geometry);
            scale_iwork_paint(fill);
            scale_iwork_paint(stroke);
            *stroke_width *= IWORK_POINT_TO_CSS_PIXEL;
            *line_height *= IWORK_POINT_TO_CSS_PIXEL;
            for run in runs {
                run.font_size *= IWORK_POINT_TO_CSS_PIXEL;
                run.baseline_shift *= IWORK_POINT_TO_CSS_PIXEL;
                run.letter_spacing *= IWORK_POINT_TO_CSS_PIXEL;
            }
        }
        Visual::Effect {
            shadow,
            clip,
            visual,
        } => {
            if let Some(shadow) = shadow {
                shadow.blur *= IWORK_POINT_TO_CSS_PIXEL;
                shadow.offset_x *= IWORK_POINT_TO_CSS_PIXEL;
                shadow.offset_y *= IWORK_POINT_TO_CSS_PIXEL;
            }
            if let Some(clip) = clip {
                scale_iwork_geometry(clip);
            }
            scale_iwork_visual(visual);
        }
        Visual::TextLayout { layout, visual } => {
            for rect in &mut layout.wrap_regions {
                rect.x *= IWORK_POINT_TO_CSS_PIXEL;
                rect.y *= IWORK_POINT_TO_CSS_PIXEL;
                rect.width *= IWORK_POINT_TO_CSS_PIXEL;
                rect.height *= IWORK_POINT_TO_CSS_PIXEL;
            }
            for tab in &mut layout.tab_stops {
                tab.position *= IWORK_POINT_TO_CSS_PIXEL;
            }
            for value in [
                &mut layout.default_tab_stop,
                &mut layout.hanging_indent,
                &mut layout.paragraph_spacing,
                &mut layout.inset_left,
                &mut layout.inset_right,
                &mut layout.inset_top,
                &mut layout.inset_bottom,
                &mut layout.margin_left,
                &mut layout.margin_right,
                &mut layout.first_line_indent,
                &mut layout.column_spacing,
                &mut layout.text_stroke_width,
                &mut layout.text_baseline,
            ] {
                *value *= IWORK_POINT_TO_CSS_PIXEL;
            }
            for paragraph in &mut layout.paragraphs {
                paragraph.margin_left *= IWORK_POINT_TO_CSS_PIXEL;
                paragraph.margin_right *= IWORK_POINT_TO_CSS_PIXEL;
                paragraph.first_line_indent *= IWORK_POINT_TO_CSS_PIXEL;
                paragraph.default_tab_stop *= IWORK_POINT_TO_CSS_PIXEL;
                paragraph.line_height *= IWORK_POINT_TO_CSS_PIXEL;
                paragraph.space_before *= IWORK_POINT_TO_CSS_PIXEL;
                paragraph.space_after *= IWORK_POINT_TO_CSS_PIXEL;
                for rule in [&mut paragraph.rule_above, &mut paragraph.rule_below]
                    .into_iter()
                    .flatten()
                {
                    rule.stroke_width *= IWORK_POINT_TO_CSS_PIXEL;
                    rule.offset_x *= IWORK_POINT_TO_CSS_PIXEL;
                    rule.offset_y *= IWORK_POINT_TO_CSS_PIXEL;
                }
            }
            scale_iwork_visual(visual);
        }
        Visual::TextEffects { effects, visual } => {
            for shadow in effects
                .iter_mut()
                .flat_map(|effect| [&mut effect.shadow, &mut effect.inner_shadow])
                .flatten()
            {
                shadow.blur *= IWORK_POINT_TO_CSS_PIXEL;
                shadow.offset_x *= IWORK_POINT_TO_CSS_PIXEL;
                shadow.offset_y *= IWORK_POINT_TO_CSS_PIXEL;
            }
            scale_iwork_visual(visual);
        }
        Visual::StrokeStyle { style, visual } => {
            for dash in &mut style.dash {
                *dash *= IWORK_POINT_TO_CSS_PIXEL;
            }
            scale_iwork_visual(visual);
        }
        Visual::AdvancedEffect {
            outer_shadow,
            inner_shadow,
            glow,
            reflection,
            soft_edge,
            three_d,
            visual,
        } => {
            if let Some(effect) = outer_shadow {
                effect.shadow.blur *= IWORK_POINT_TO_CSS_PIXEL;
                effect.shadow.offset_x *= IWORK_POINT_TO_CSS_PIXEL;
                effect.shadow.offset_y *= IWORK_POINT_TO_CSS_PIXEL;
            }
            if let Some(shadow) = inner_shadow {
                shadow.blur *= IWORK_POINT_TO_CSS_PIXEL;
                shadow.offset_x *= IWORK_POINT_TO_CSS_PIXEL;
                shadow.offset_y *= IWORK_POINT_TO_CSS_PIXEL;
            }
            if let Some(glow) = glow {
                glow.radius *= IWORK_POINT_TO_CSS_PIXEL;
            }
            if let Some(reflection) = reflection {
                reflection.blur *= IWORK_POINT_TO_CSS_PIXEL;
                reflection.distance *= IWORK_POINT_TO_CSS_PIXEL;
            }
            if let Some(soft_edge) = soft_edge {
                *soft_edge *= IWORK_POINT_TO_CSS_PIXEL;
            }
            if let Some(three_d) = three_d {
                three_d.z *= IWORK_POINT_TO_CSS_PIXEL;
                three_d.extrusion_height *= IWORK_POINT_TO_CSS_PIXEL;
                three_d.contour_width *= IWORK_POINT_TO_CSS_PIXEL;
                if let Some(bevel) = &mut three_d.bevel_top {
                    bevel.width *= IWORK_POINT_TO_CSS_PIXEL;
                    bevel.height *= IWORK_POINT_TO_CSS_PIXEL;
                }
                if let Some(bevel) = &mut three_d.bevel_bottom {
                    bevel.width *= IWORK_POINT_TO_CSS_PIXEL;
                    bevel.height *= IWORK_POINT_TO_CSS_PIXEL;
                }
                if let Some(backdrop) = &mut three_d.backdrop {
                    backdrop.anchor_x *= IWORK_POINT_TO_CSS_PIXEL;
                    backdrop.anchor_y *= IWORK_POINT_TO_CSS_PIXEL;
                    backdrop.anchor_z *= IWORK_POINT_TO_CSS_PIXEL;
                }
                if let Some(flat_text_z) = &mut three_d.flat_text_z {
                    *flat_text_z *= IWORK_POINT_TO_CSS_PIXEL;
                }
            }
            scale_iwork_visual(visual);
        }
        Visual::Media { poster, .. } => scale_iwork_visual(poster),
        Visual::ImageColorChange { visual, .. }
        | Visual::ImageAdjustment { visual, .. }
        | Visual::ColorManagedImage { visual, .. } => scale_iwork_visual(visual),
        Visual::Group { children } => {
            for child in children {
                child.bounds.x *= IWORK_POINT_TO_CSS_PIXEL;
                child.bounds.y *= IWORK_POINT_TO_CSS_PIXEL;
                child.bounds.width *= IWORK_POINT_TO_CSS_PIXEL;
                child.bounds.height *= IWORK_POINT_TO_CSS_PIXEL;
                scale_iwork_visual(&mut child.visual);
            }
        }
        Visual::OpacityMask { visual, .. } => scale_iwork_visual(visual),
        Visual::None
        | Visual::Image { .. }
        | Visual::ImageWithFallback { .. }
        | Visual::MaskedImage { .. } => {}
    }
}

fn pages_document_layout<'a>(
    bytes: &[u8],
    font_metrics: &'a FontMetricTable,
) -> Result<Option<PagesDocumentLayout<'a>>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut body_storage = None;
    let mut values = [None; 8];
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                DOCUMENT_COMPONENT,
                format!("invalid Pages document key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (4, 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    DOCUMENT_COMPONENT,
                    "Pages body storage reference",
                )?;
                body_storage = Some(parse_iwa_reference(reference)?);
            }
            (field @ 30..=37, 5) => {
                values[usize::try_from(field - 30).unwrap()] = Some(keynote_fixed32(
                    bytes,
                    &mut cursor,
                    DOCUMENT_COMPONENT,
                    "Pages document dimension",
                )?);
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, DOCUMENT_COMPONENT)?,
        }
    }
    let Some(body_storage) = body_storage else {
        return Ok(None);
    };
    let [
        Some(width),
        Some(height),
        Some(margin_left),
        Some(margin_right),
        Some(margin_top),
        Some(margin_bottom),
        header_margin,
        footer_margin,
    ] = values
    else {
        return Ok(None);
    };
    if width <= 0.0 || height <= 0.0 {
        return Err(iwa_error(
            DOCUMENT_COMPONENT,
            "Pages document dimensions are invalid",
        ));
    }
    Ok(Some(PagesDocumentLayout {
        frame_targets: &[],
        font_metrics,
        body_storage,
        width,
        height,
        margin_left: margin_left.max(0.0),
        margin_right: margin_right.max(0.0),
        margin_top: margin_top.max(0.0),
        margin_bottom: margin_bottom.max(footer_margin.unwrap_or(0.0)).max(0.0),
        header_distance: header_margin.unwrap_or(0.0).max(0.0),
        footer_distance: footer_margin.unwrap_or(0.0).max(0.0),
        view_scale: 1.0,
    }))
}

fn keynote_point_path(
    bytes: &[u8],
    width: f32,
    height: f32,
    part: &str,
) -> Result<Geometry, Diagnostic> {
    let kind = keynote_varint_field(bytes, 1, part)?.unwrap_or(1);
    if kind == 200 {
        // iWork stores independent arm widths in the path's original coordinate space.
        // DrawingML/legacy adjustments have different units and stay in their adapters.
        let (source_width, source_height) = keynote_nested_message(bytes, 3, part)?
            .map(|size| iwork_float_pair(size, part))
            .transpose()?
            .filter(|(width, height)| *width > 0.0 && *height > 0.0)
            .unwrap_or((width.max(1.0), height.max(1.0)));
        let (arm_width, arm_height) = keynote_nested_message(bytes, 2, part)?
            .map(|point| iwork_float_pair(point, part))
            .transpose()?
            .unwrap_or((source_width / 2.0, source_height / 2.0));
        let left = width * (1.0 - (arm_width / source_width).clamp(0.0, 1.0)) / 2.0;
        let top = height * (1.0 - (arm_height / source_height).clamp(0.0, 1.0)) / 2.0;
        return Ok(super::cross_geometry(
            width,
            height,
            left,
            width - left,
            top,
            height - top,
        ));
    }
    if kind == 100 {
        let (points, inner_radius) = keynote_nested_message(bytes, 2, part)?
            .map(|point| iwork_float_pair(point, part))
            .transpose()?
            .map(|(points, radius)| {
                (
                    (points.round() as usize).clamp(3, 100),
                    (radius * 0.5).clamp(0.0, 0.5),
                )
            })
            .unwrap_or((5, 0.190_983));
        let mut commands = (0..points * 2)
            .map(|index| {
                let angle = -std::f32::consts::FRAC_PI_2
                    + std::f32::consts::PI * index as f32 / points as f32;
                let radius = if index % 2 == 0 { 0.5 } else { inner_radius };
                let x = width * (0.5 + radius * angle.cos());
                let y = height * (0.5 + radius * angle.sin());
                if index == 0 {
                    PathCommand::MoveTo { x, y }
                } else {
                    PathCommand::LineTo { x, y }
                }
            })
            .collect::<Vec<_>>();
        commands.push(PathCommand::ClosePath);
        return Ok(Geometry::Path {
            fill_rule: FillRule::NonZero,
            commands,
        });
    }
    let points: &[(f32, f32)] = match kind {
        0 => &[
            (0.0, 0.5),
            (0.2, 0.0),
            (0.2, 0.25),
            (1.0, 0.25),
            (1.0, 0.75),
            (0.2, 0.75),
            (0.2, 1.0),
        ],
        10 => &[
            (0.0, 0.5),
            (0.2, 0.0),
            (0.2, 0.25),
            (0.8, 0.25),
            (0.8, 0.0),
            (1.0, 0.5),
            (0.8, 1.0),
            (0.8, 0.75),
            (0.2, 0.75),
            (0.2, 1.0),
        ],
        _ => &[
            (0.0, 0.25),
            (0.8, 0.25),
            (0.8, 0.0),
            (1.0, 0.5),
            (0.8, 1.0),
            (0.8, 0.75),
            (0.0, 0.75),
        ],
    };
    let mut commands = points
        .iter()
        .enumerate()
        .map(|(index, (x, y))| {
            if index == 0 {
                PathCommand::MoveTo {
                    x: x * width,
                    y: y * height,
                }
            } else {
                PathCommand::LineTo {
                    x: x * width,
                    y: y * height,
                }
            }
        })
        .collect::<Vec<_>>();
    commands.push(PathCommand::ClosePath);
    Ok(Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    })
}

fn pages_layout_hints(
    package: &Package<'_>,
    limits: Limits,
) -> Result<(Vec<PagesPageHint>, f32), Diagnostic> {
    let mut components: Vec<String> = package
        .entry_names()
        .filter(|part| part.starts_with("Index/ViewState") && part.ends_with(".iwa"))
        .map(str::to_owned)
        .collect();
    components.sort();
    for component in components {
        let bytes = package.required_part(&component)?;
        let archives = parse_iwa_archives(package, &component, &bytes, limits)?;
        let view_scale = archives
            .iter()
            .flat_map(|archive| &archive.messages)
            .find(|message| message.message_type == PAGES_VIEW_STATE_TYPE)
            .map(|message| pages_view_scale(&message.payload, &component))
            .transpose()?
            .unwrap_or(1.0);
        if let Some(message) = archives
            .iter()
            .flat_map(|archive| &archive.messages)
            .find(|message| message.message_type == PAGES_LAYOUT_STATE_TYPE)
        {
            return Ok((
                pages_parse_layout_state(&message.payload, &component, &archives)?,
                view_scale,
            ));
        }
    }
    Ok((Vec::new(), 1.0))
}

fn pages_view_scale(bytes: &[u8], part: &str) -> Result<f32, Diagnostic> {
    Ok(keynote_fixed32_field(bytes, 15, part)?
        .filter(|scale| scale.is_finite() && (0.1..=10.0).contains(scale))
        .unwrap_or(1.0))
}

fn pages_parse_layout_state(
    bytes: &[u8],
    part: &str,
    archives: &[IwaArchive],
) -> Result<Vec<PagesPageHint>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut pages = Vec::new();
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Pages layout-state key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (5, 2) => {
                let section = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Pages section layout hint",
                )?;
                pages_parse_section_hint(section, part, archives, &mut pages)?;
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    for pair in pages.windows(2) {
        let previous_end = pair[0].targets.last().map(|target| target.range_end);
        let next_start = pair[1].targets.first().map(|target| target.range_start);
        if previous_end
            .zip(next_start)
            .is_some_and(|(end, start)| end > start)
        {
            return Err(iwa_error(part, "Pages saved page ranges overlap"));
        }
    }
    Ok(pages)
}

fn pages_parse_section_hint(
    bytes: &[u8],
    part: &str,
    archives: &[IwaArchive],
    pages: &mut Vec<PagesPageHint>,
) -> Result<(), Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Pages section-hint key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let page =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Pages page layout hint")?;
                if let Some(page) = pages_parse_page_hint(page, part, archives)? {
                    pages.push(page);
                }
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(())
}

fn pages_parse_page_hint(
    bytes: &[u8],
    part: &str,
    archives: &[IwaArchive],
) -> Result<Option<PagesPageHint>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut targets = Vec::new();
    let mut attachment_positions = Vec::new();
    let mut toc_ranges = Vec::new();
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Pages page-hint key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (2, 2) => {
                let target = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Pages target layout hint",
                )?;
                if let Some(target) = pages_parse_target_hint(target, part)? {
                    if targets.last().is_some_and(|previous: &PagesTargetHint| {
                        previous.range_end > target.range_start
                    }) {
                        return Err(iwa_error(part, "Pages saved target ranges overlap"));
                    }
                    targets.push(target);
                }
            }
            (12, 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Pages attachment layout reference",
                )?;
                if let Some(layout) = archive_message(
                    archives,
                    parse_iwa_reference(reference)?,
                    PAGES_TOC_LAYOUT_TYPE,
                    part,
                )? && let Some(range) = numbers_first_bytes(&layout.payload, 1, part)?
                {
                    toc_ranges.push(pages_parse_range(range, part)?);
                }
            }
            (8, 2) => {
                let anchor = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Pages anchored attachment position",
                )?;
                if let Some(position) = pages_parse_anchor_position(anchor, part)? {
                    attachment_positions.push(position);
                }
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    if targets.is_empty() {
        return Ok(None);
    }
    let mut flow_diagnostics = Vec::new();
    let flow_targets = (|| {
        let flows = numbers_references(bytes, 14, part)?;
        let hints = numbers_bytes(bytes, 15, part)?;
        if flows.len() != hints.len() {
            return Err(iwa_error(
                part,
                "text flow references and page ranges do not match",
            ));
        }
        flows
            .into_iter()
            .zip(hints)
            .map(|(flow, target)| {
                pages_parse_target_hint(target, part)
                    .map(|target| target.map(|target| (flow, target)))
            })
            .collect::<Result<Vec<_>, _>>()
            .map(|targets| targets.into_iter().flatten().collect())
    })()
    .unwrap_or_else(|error| {
        flow_diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Layout,
                Fidelity::Approximate,
                format!("PAGES_TEXT_FLOW_APPROXIMATE: {}", error.message),
            )
            .in_part(part),
        );
        Vec::new()
    });
    Ok(Some(PagesPageHint {
        flow_targets,
        flow_diagnostics,
        targets,
        attachment_positions,
        toc_ranges,
    }))
}

fn pages_parse_target_hint(
    bytes: &[u8],
    part: &str,
) -> Result<Option<PagesTargetHint>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut range = None;
    let mut anchored = None;
    let mut origin = None;
    let mut size = None;
    let mut column_count = 1_u32;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Pages target-hint key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (field @ (3 | 5), 2) => {
                let encoded =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Pages text range")?;
                let value = pages_parse_range(encoded, part)?;
                if field == 3 {
                    range = Some(value);
                } else {
                    anchored = Some(value);
                }
            }
            (field @ (8 | 9), 2) => {
                let encoded =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Pages target geometry")?;
                let value = pages_double_pair(encoded, part)?;
                if field == 8 {
                    origin = value;
                } else {
                    size = value;
                }
            }
            (6, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(
                        part,
                        format!("invalid Pages target column count: {message}"),
                    )
                })?;
                cursor += consumed;
                column_count = u32::try_from(value).unwrap_or(64).clamp(1, 64);
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    let Some((range_start, range_end)) = range else {
        return Ok(None);
    };
    let (anchored_start, anchored_end) = anchored.unwrap_or((range_start, range_end));
    Ok(Some(PagesTargetHint {
        range_start,
        range_end,
        anchored_start,
        anchored_end,
        origin,
        size,
        column_count,
    }))
}

fn pages_parse_range(bytes: &[u8], part: &str) -> Result<(usize, usize), Diagnostic> {
    let mut cursor = 0_usize;
    let mut location = None;
    let mut length = None;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Pages range key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (field @ (1 | 2), 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Pages range value: {message}"))
                })?;
                cursor += consumed;
                let value = usize::try_from(value)
                    .map_err(|_| iwa_error(part, "Pages range value is too large"))?;
                if field == 1 {
                    location = Some(value);
                } else {
                    length = Some(value);
                }
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    let location = location.unwrap_or(0);
    let end = location
        .checked_add(length.unwrap_or(0))
        .ok_or_else(|| iwa_error(part, "Pages range overflows"))?;
    Ok((location, end))
}

fn pages_parse_anchor_position(bytes: &[u8], part: &str) -> Result<Option<(f32, f32)>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Pages anchor key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (2, 2) => {
                let point =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Pages anchor point")?;
                return iwork_float_pair(point, part).map(Some);
            }
            (4, 2) => {
                let point = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Pages precise anchor point",
                )?;
                return pages_double_pair(point, part);
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(None)
}

fn pages_double_pair(bytes: &[u8], part: &str) -> Result<Option<(f32, f32)>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut values = [0.0_f32; 2];
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Pages double point key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (field @ (1 | 2), 1) => {
                let end = cursor
                    .checked_add(8)
                    .ok_or_else(|| iwa_error(part, "Pages double point offset overflows"))?;
                let encoded: [u8; 8] = bytes
                    .get(cursor..end)
                    .ok_or_else(|| iwa_error(part, "Pages double point is truncated"))?
                    .try_into()
                    .map_err(|_| iwa_error(part, "Pages double point is invalid"))?;
                let value = f64::from_le_bytes(encoded);
                if value == f64::INFINITY {
                    return Ok(None);
                }
                if !value.is_finite() {
                    return Err(iwa_error(part, "Pages double point is not finite"));
                }
                let value = value as f32;
                if !value.is_finite() {
                    return Err(iwa_error(
                        part,
                        "Pages double point is outside the renderable range",
                    ));
                }
                values[usize::try_from(field - 1).unwrap()] = value;
                cursor = end;
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(Some((values[0], values[1])))
}

#[allow(clippy::too_many_arguments)]
fn pages_attachment_objects(
    diagnostics: &mut Vec<Diagnostic>,
    package: &Package<'_>,
    archives: &[IwaArchive],
    messages: &NumbersMessageSpace,
    data_files: &[(u64, String)],
    storage: &KeynoteTextStorage,
    page: &PagesPageHint,
    previous_page: Option<&PagesPageHint>,
    unit_index: u32,
    layout: PagesDocumentLayout,
    limits: Limits,
    objects: &mut Vec<Object>,
    embedded_fonts: &mut Vec<EmbeddedFont>,
) -> Result<(), Diagnostic> {
    let mut attachments = Vec::new();
    for change in &storage.attachments {
        let Some(identifier) = change.identifier else {
            continue;
        };
        let inside = page.targets.iter().any(|target| {
            target.anchored_start <= change.character_index
                && change.character_index < target.anchored_end
        });
        let at_trailing_boundary = !inside
            && page
                .targets
                .last()
                .is_some_and(|target| target.anchored_end == change.character_index);
        if (inside || at_trailing_boundary)
            && let Some(attachment) = archive_message(
                archives,
                identifier,
                PAGES_TOC_ATTACHMENT_TYPE,
                DOCUMENT_COMPONENT,
            )?
        {
            let first = objects.len();
            let result = pages_toc_objects(
                archives,
                messages,
                storage,
                attachment,
                change.character_index,
                page,
                unit_index,
                layout,
                limits,
                objects,
            );
            match result {
                Ok(()) if objects.len() > first => {}
                Err(error) if error.code != DiagnosticCode::FormatInvalid => return Err(error),
                result => diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Render,
                        Fidelity::Omitted,
                        format!(
                            "PAGES_TOC_OMITTED: {}",
                            result.err().map_or_else(
                                || "missing TOC content or saved range".to_owned(),
                                |error| error.message
                            )
                        ),
                    )
                    .in_part(DOCUMENT_COMPONENT),
                ),
            }
            continue;
        }
        let trailing_table = if at_trailing_boundary {
            let Some(drawable) =
                pages_attachment_drawable(archives, identifier, DOCUMENT_COMPONENT)?
            else {
                continue;
            };
            messages.message(drawable, IWORK_TABLE_TYPE)?.is_some()
        } else {
            false
        };
        if inside || trailing_table {
            attachments.push((change.character_index, identifier));
        }
    }
    let mut positioned = page.attachment_positions.iter().copied();
    for (character_index, attachment_identifier) in attachments {
        let Some(attachment) =
            pages_attachment(archives, attachment_identifier, DOCUMENT_COMPONENT)?
        else {
            continue;
        };
        let drawable_identifier = attachment.drawable;
        if let Some(group) = archive_message(
            archives,
            drawable_identifier,
            IWORK_GROUP_TYPE,
            DOCUMENT_COMPONENT,
        )? {
            let (_, children) = keynote_group(&group.payload, DOCUMENT_COMPONENT)?;
            let Some((anchor_x, anchor_y)) = positioned.next() else {
                continue;
            };
            for child in children {
                if let Some(image_message) =
                    archive_message(archives, child, IWORK_IMAGE_TYPE, DOCUMENT_COMPONENT)?
                {
                    let image = keynote_image(&image_message.payload, DOCUMENT_COMPONENT)?;
                    let bounds = Rect {
                        x: anchor_x + image.geometry.x,
                        y: anchor_y + image.geometry.y,
                        width: image.geometry.width.max(1.0),
                        height: image.geometry.height.max(1.0),
                    };
                    if let Some((part, bytes)) = pages_pdf_data(package, data_files, &image)?
                        && iwork_push_pdf_objects(
                            &part,
                            &bytes,
                            child,
                            unit_index,
                            bounds,
                            50,
                            "inline-image",
                            IWORK_POINT_TO_CSS_PIXEL,
                            true,
                            limits,
                            objects,
                            embedded_fonts,
                        )?
                    {
                        continue;
                    }
                    let Some((part, media_type, bytes)) =
                        pages_image_data(package, data_files, &image, limits)?
                    else {
                        continue;
                    };
                    pages_push_object(
                        objects,
                        limits,
                        unit_index,
                        child,
                        ObjectKind::Image,
                        bounds,
                        50,
                        None,
                        "inline-image",
                        MappingQuality::Exact,
                        Visual::Image {
                            media_type,
                            bytes,
                            crop: ImageCrop::default(),
                        },
                    )?;
                    if let Some(object) = objects.last_mut() {
                        object.source.part = part;
                    }
                } else if let Some(chart_message) =
                    archive_message(archives, child, IWORK_CHART_TYPE, DOCUMENT_COMPONENT)?
                    && let Some(chart) = keynote_chart(
                        &chart_message.payload,
                        archives,
                        archives,
                        DOCUMENT_COMPONENT,
                    )?
                {
                    keynote_push_chart(
                        chart,
                        child,
                        unit_index,
                        DOCUMENT_COMPONENT,
                        limits,
                        objects,
                        anchor_x,
                        anchor_y,
                        50,
                    )?;
                }
            }
        } else if let Some(image_message) = archive_message(
            archives,
            drawable_identifier,
            IWORK_IMAGE_TYPE,
            DOCUMENT_COMPONENT,
        )? {
            let image = keynote_image(&image_message.payload, DOCUMENT_COMPONENT)?;
            let target = page.targets.iter().find(|target| {
                target.anchored_start <= character_index && character_index < target.anchored_end
            });
            let range_start = target.map_or(character_index, |target| target.range_start);
            let target_y = target
                .and_then(|target| target.origin)
                .map_or(layout.margin_top, |(_, y)| layout.margin_top + y);
            let character_byte = keynote_utf16_byte_offset(&storage.text, character_index)
                .ok_or_else(|| {
                    iwa_error(
                        DOCUMENT_COMPONENT,
                        "Pages attachment splits a UTF-16 surrogate pair",
                    )
                })?;
            let paragraph_byte = storage.text[..character_byte]
                .char_indices()
                .rev()
                .find_map(|(index, character)| {
                    matches!(character, '\n' | '\u{2028}').then_some(index + character.len_utf8())
                })
                .unwrap_or(0);
            let paragraph_start = storage.text[..paragraph_byte]
                .encode_utf16()
                .count()
                .max(range_start);
            let mut x = layout.margin_left + attachment.horizontal_offset;
            for previous in storage.attachments.iter().filter(|previous| {
                paragraph_start <= previous.character_index
                    && previous.character_index < character_index
            }) {
                let Some(previous_identifier) = previous.identifier else {
                    continue;
                };
                let Some(previous_drawable) =
                    pages_attachment_drawable(archives, previous_identifier, DOCUMENT_COMPONENT)?
                else {
                    continue;
                };
                let Some(previous_image) = archive_message(
                    archives,
                    previous_drawable,
                    IWORK_IMAGE_TYPE,
                    DOCUMENT_COMPONENT,
                )?
                else {
                    continue;
                };
                x += keynote_image(&previous_image.payload, DOCUMENT_COMPONENT)?
                    .geometry
                    .width;
            }
            let y = pages_table_top(
                storage,
                archives,
                range_start,
                paragraph_start,
                layout,
                target_y,
                messages,
                limits,
            )?;
            let first_object = objects.len();
            keynote_native_drawable(
                diagnostics,
                package,
                archives,
                archives,
                messages,
                data_files,
                drawable_identifier,
                unit_index,
                DOCUMENT_COMPONENT,
                limits,
                objects,
                &mut Vec::new(),
                layout.width,
                layout.height,
                x - image.geometry.x,
                y - image.geometry.y,
                50,
                true,
                embedded_fonts,
                IWORK_POINT_TO_CSS_PIXEL,
                None,
            )?;
            for object in &mut objects[first_object..] {
                if let SourceLocator::Iwork { kind, .. } = &mut object.source.locator {
                    *kind = "inline-image";
                }
            }
        } else if let Some((table_part, table)) =
            messages.message(drawable_identifier, IWORK_TABLE_TYPE)?
        {
            let geometry = pages_table_geometry(&table.payload, DOCUMENT_COMPONENT)?;
            let Some(model_identifier) = numbers_references(&table.payload, 2, table_part)?
                .into_iter()
                .next()
            else {
                continue;
            };
            let Some((model_part, model)) =
                messages.message(model_identifier, NUMBERS_TABLE_MODEL_TYPE)?
            else {
                continue;
            };
            let (rows, columns, cells) = numbers_table_cells(messages, model, model_part, limits)?;
            let target = page.targets.iter().find(|target| {
                target.anchored_start <= character_index && character_index <= target.anchored_end
            });
            let table_width = target
                .and_then(|target| target.size)
                .map(|(width, _)| width)
                .filter(|width| width.is_finite() && *width > 0.0)
                .unwrap_or(geometry.width)
                .max(1.0);
            let column_widths =
                pages_table_column_widths(messages, model, model_part, columns, table_width)?;
            let mut row_heights =
                pages_table_row_heights(messages, model, model_part, rows, geometry.height)?;
            let header_rows = numbers_varint(&model.payload, 9, model_part)?
                .and_then(|value| usize::try_from(value).ok())
                .unwrap_or(0);
            for height in row_heights.iter_mut().skip(header_rows) {
                *height *= layout.view_scale;
            }
            iwork_table_image_row_heights(archives, &cells, &mut row_heights, limits)?;
            let x = pages_table_left(
                storage,
                archives,
                character_index,
                attachment,
                layout,
                table_width,
            )?;
            let range_start = target.map_or(character_index, |target| target.range_start);
            let target_y = target
                .and_then(|target| target.origin)
                .map_or(layout.margin_top, |(_, y)| layout.margin_top + y);
            let y = pages_table_top(
                storage,
                archives,
                range_start,
                character_index,
                layout,
                target_y,
                messages,
                limits,
            )?;
            let fitting_end = |start: usize, available: f32| {
                let mut height = 0.0_f32;
                let mut end = start;
                while let Some(row_height) = row_heights.get(end)
                    && height + row_height <= available + 0.5
                {
                    height += row_height;
                    end += 1;
                }
                end
            };
            let continuation = previous_page.is_some_and(|previous| {
                previous
                    .targets
                    .last()
                    .is_some_and(|target| target.anchored_end == character_index)
                    && page
                        .targets
                        .first()
                        .is_some_and(|target| target.anchored_start == character_index)
            });
            let trailing_fragment = page
                .targets
                .last()
                .is_some_and(|target| target.anchored_end == character_index);
            let row_start = if continuation {
                let previous = previous_page.unwrap();
                let previous_target = previous.targets.last().unwrap();
                let previous_y = pages_table_top(
                    storage,
                    archives,
                    previous_target.range_start,
                    character_index,
                    layout,
                    previous_target
                        .origin
                        .map_or(layout.margin_top, |(_, y)| layout.margin_top + y),
                    messages,
                    limits,
                )?;
                fitting_end(
                    0,
                    (layout.height - layout.margin_bottom - previous_y).max(0.0),
                )
            } else {
                0
            };
            let row_end = if trailing_fragment || continuation {
                fitting_end(
                    row_start,
                    (layout.height - layout.margin_bottom - y).max(0.0),
                )
            } else {
                row_heights.len()
            };
            if row_start >= row_end {
                continue;
            }
            let fragment_rows = u32::try_from(row_end - row_start)
                .map_err(|_| iwa_error(model_part, "Pages table fragment is too large"))?;
            let fragment_heights = &row_heights[row_start..row_end];
            let fragment_height = fragment_heights.iter().sum::<f32>();
            let row_start_u32 = u32::try_from(row_start)
                .map_err(|_| iwa_error(model_part, "Pages table fragment start is too large"))?;
            let row_end_u32 = u32::try_from(row_end)
                .map_err(|_| iwa_error(model_part, "Pages table fragment end is too large"))?;
            let fragment_cells = cells
                .iter()
                .filter(|cell| row_start_u32 <= cell.row && cell.row < row_end_u32)
                .cloned()
                .map(|mut cell| {
                    cell.row -= row_start_u32;
                    cell
                })
                .collect::<Vec<_>>();
            let cell_fills = keynote_table_cell_fills(
                messages,
                archives,
                model,
                model_part,
                table_width,
                geometry.height,
                limits,
            )?;
            let cell_text_styles =
                iwork_table_text_styles(messages, archives, model, model_part, limits)?;
            let banded_fills = keynote_table_banded_fills(
                archives,
                model,
                model_part,
                table_width,
                geometry.height,
            )?;
            let fragment_merges =
                pages_table_trailing_merges(model, model_part, rows, columns, &cells)?
                    .into_iter()
                    .filter(|merge| row_start_u32 <= merge.row && merge.row < row_end_u32)
                    .map(|mut merge| {
                        merge.row -= row_start_u32;
                        merge
                    })
                    .collect::<Vec<_>>();
            let borders =
                iwork_table_borders(messages, model, model_part, rows, columns, diagnostics);
            let cell_attachments = pages_push_table_grid(
                objects,
                limits,
                unit_index,
                drawable_identifier,
                Rect {
                    x,
                    y,
                    width: table_width,
                    height: fragment_height.max(1.0),
                },
                model,
                model_part,
                fragment_rows,
                &column_widths,
                fragment_heights,
                &fragment_cells,
                archives,
                &cell_fills,
                &cell_text_styles,
                banded_fills.as_ref(),
                &fragment_merges,
                borders.as_deref(),
                row_start_u32,
            )?;
            let first_image = objects.len();
            iwork_table_cell_images(
                diagnostics,
                package,
                archives,
                archives,
                messages,
                data_files,
                &cell_attachments,
                unit_index,
                limits,
                objects,
                embedded_fonts,
                IWORK_POINT_TO_CSS_PIXEL,
            )?;
            for object in &mut objects[first_image..] {
                if let SourceLocator::Iwork { kind, .. } = &mut object.source.locator {
                    *kind = "inline-image";
                }
            }
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn pages_toc_objects(
    archives: &[IwaArchive],
    messages: &NumbersMessageSpace,
    body: &KeynoteTextStorage,
    attachment: &IwaMessage,
    character_index: usize,
    page: &PagesPageHint,
    unit_index: u32,
    layout: PagesDocumentLayout,
    limits: Limits,
    objects: &mut Vec<Object>,
) -> Result<(), Diagnostic> {
    let part = DOCUMENT_COMPONENT;
    let Some(base) = numbers_first_bytes(&attachment.payload, 1, part)? else {
        return Ok(());
    };
    let Some(identifier) = keynote_reference_field(base, 1, part)? else {
        return Ok(());
    };
    let Some(toc) = archive_message(archives, identifier, PAGES_TOC_TYPE, part)? else {
        return Ok(());
    };
    let Some(shape) = numbers_first_bytes(&toc.payload, 1, part)? else {
        return Ok(());
    };
    let Some(storage_identifier) = keynote_reference_field(shape, 2, part)? else {
        return Ok(());
    };
    let Some(message) =
        archive_message(archives, storage_identifier, IWORK_TEXT_STORAGE_TYPE, part)?
    else {
        return Ok(());
    };
    let mut storage = keynote_storage_text(&message.payload, part, limits)?;
    // Cached TOC labels already contain their generated heading text.
    storage.list_styles.clear();
    let Some(target) = page.targets.iter().find(|target| {
        target.anchored_start <= character_index && character_index <= target.anchored_end
    }) else {
        return Ok(());
    };
    let (x, origin_y) = target.origin.unwrap_or((0.0, 0.0));
    let width = target.size.map_or(
        layout.width - layout.margin_left - layout.margin_right,
        |size| size.0,
    );
    let mut y = pages_table_top(
        body,
        archives,
        target.range_start,
        character_index,
        layout,
        layout.margin_top + origin_y,
        messages,
        limits,
    )?;
    for &(start, end) in &page.toc_ranges {
        let start_byte = keynote_utf16_byte_offset(&storage.text, start)
            .ok_or_else(|| iwa_error(part, "Pages TOC range start is invalid"))?;
        let end_byte = keynote_utf16_byte_offset(&storage.text, end)
            .filter(|end| *end >= start_byte)
            .ok_or_else(|| iwa_error(part, "Pages TOC range end is invalid"))?;
        let mut position = start;
        for paragraph in storage.text[start_byte..end_byte].split_inclusive(['\n', '\u{2028}']) {
            let next = position + paragraph.encode_utf16().count();
            let runs = pages_text_runs_for_range(&storage, archives, part, position, next, limits)?;
            let paragraphs = pages_paragraph_layouts(
                &storage,
                &runs,
                archives,
                part,
                position,
                next,
                messages,
                limits,
                layout.view_scale,
                layout.font_metrics,
                false,
            )?;
            let style = iwork_inherited_style_at(&storage.paragraph_styles, position)
                .map(|id| keynote_text_style(archives, id, part))
                .transpose()?
                .unwrap_or_default();
            let height = keynote_text_layout_height(
                &runs,
                &paragraphs,
                (width - style.margin_left - style.margin_right).max(1.0),
            )
            .max(1.0);
            let text = runs.iter().map(|run| run.text.as_str()).collect::<String>();
            pages_push_object(
                objects,
                limits,
                unit_index,
                storage_identifier,
                ObjectKind::TextBox,
                Rect {
                    x: layout.margin_left + x,
                    y,
                    width,
                    height,
                },
                30,
                Some(text),
                "body",
                MappingQuality::Approximate,
                Visual::TextLayout {
                    layout: TextLayout {
                        inset_left: 0.0,
                        inset_right: 0.0,
                        inset_top: 0.0,
                        inset_bottom: 0.0,
                        // TOC labels are already expanded; their sole tab separates the page field.
                        tab_stops: style
                            .tab_stops
                            .into_iter()
                            .filter(|tab| tab.align == TextAlign::End)
                            .collect(),
                        paragraphs,
                        ..TextLayout::default()
                    },
                    visual: Box::new(Visual::RichText {
                        geometry: Geometry::Rectangle,
                        fill: Paint::None,
                        stroke: Paint::None,
                        stroke_width: 0.0,
                        align: TextAlign::Start,
                        line_height: 0.0,
                        runs,
                    }),
                },
            )?;
            objects
                .last_mut()
                .unwrap()
                .stable_id
                .push_str(&format!("-toc-{position}"));
            y += height;
            position = next;
        }
    }
    Ok(())
}

fn pages_attachment_drawable(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
) -> Result<Option<u64>, Diagnostic> {
    Ok(pages_attachment(archives, identifier, part)?.map(|attachment| attachment.drawable))
}

fn pages_attachment(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
) -> Result<Option<PagesAttachment>, Diagnostic> {
    let Some(attachment) =
        archive_message(archives, identifier, IWORK_DRAWABLE_ATTACHMENT_TYPE, part)?
    else {
        return Ok(None);
    };
    let mut cursor = 0_usize;
    let mut drawable = None;
    let mut horizontal_offset_type = 0_u64;
    let mut horizontal_offset = 0.0_f32;
    while cursor < attachment.payload.len() {
        let (key, consumed) = read_varint(&attachment.payload[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Pages attachment key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let reference = read_iwa_length_delimited(
                    &attachment.payload,
                    &mut cursor,
                    part,
                    "Pages attachment drawable",
                )?;
                drawable = Some(parse_iwa_reference(reference)?);
            }
            (2, 0) => {
                let (value, consumed) =
                    read_varint(&attachment.payload[cursor..]).map_err(|message| {
                        iwa_error(
                            part,
                            format!("invalid Pages horizontal offset type: {message}"),
                        )
                    })?;
                cursor += consumed;
                horizontal_offset_type = value;
            }
            (3, 5) => {
                horizontal_offset = keynote_fixed32(
                    &attachment.payload,
                    &mut cursor,
                    part,
                    "Pages horizontal offset",
                )?;
            }
            (_, wire) => skip_protobuf_value(&attachment.payload, &mut cursor, wire, part)?,
        }
    }
    Ok(drawable.map(|drawable| PagesAttachment {
        drawable,
        horizontal_offset_type,
        horizontal_offset,
    }))
}

fn pages_table_geometry(bytes: &[u8], part: &str) -> Result<KeynoteGeometry, Diagnostic> {
    let Some(table_info) = keynote_nested_message(bytes, 1, part)? else {
        return Ok(KeynoteGeometry::default());
    };
    keynote_drawable_geometry(table_info, part)
}

fn pages_image_data(
    package: &Package<'_>,
    data_files: &[(u64, String)],
    image: &KeynoteImage,
    limits: Limits,
) -> Result<Option<(String, String, Vec<u8>)>, Diagnostic> {
    for identifier in &image.data_identifiers {
        let Some((_, part)) = data_files
            .iter()
            .find(|(candidate, _)| candidate == identifier)
        else {
            continue;
        };
        let bytes = package.required_part(part)?;
        let Ok((media_type, _, _)) = iwork_image_properties(part, &bytes, limits) else {
            continue;
        };
        return Ok(Some((
            part.clone(),
            media_type.to_owned(),
            bytes.into_vec(),
        )));
    }
    Ok(None)
}

fn pages_pdf_data(
    package: &Package<'_>,
    data_files: &[(u64, String)],
    image: &KeynoteImage,
) -> Result<Option<(String, Vec<u8>)>, Diagnostic> {
    for identifier in &image.data_identifiers {
        let Some((_, part)) = data_files
            .iter()
            .find(|(candidate, _)| candidate == identifier)
        else {
            continue;
        };
        if part
            .rsplit_once('.')
            .is_none_or(|(_, extension)| !extension.eq_ignore_ascii_case("pdf"))
        {
            continue;
        }
        return Ok(Some((
            part.clone(),
            package.required_part(part)?.into_vec(),
        )));
    }
    Ok(None)
}

#[allow(clippy::too_many_arguments)]
fn iwork_push_pdf_objects(
    part: &str,
    bytes: &[u8],
    identifier: u64,
    unit_index: u32,
    bounds: Rect,
    base_z: i32,
    source_kind: &'static str,
    parent_coordinate_scale: f32,
    visual_coordinates_are_final: bool,
    limits: Limits,
    objects: &mut Vec<Object>,
    embedded_fonts: &mut Vec<EmbeddedFont>,
) -> Result<bool, Diagnostic> {
    let Some(mut pdf) = crate::format::pdf::detect_and_parse(bytes, limits)? else {
        return Ok(false);
    };
    if pdf.fatal || pdf.units.len() != 1 {
        return Ok(false);
    }
    let source_unit = &pdf.units[0];
    if source_unit.width <= 0.0 || source_unit.height <= 0.0 {
        return Ok(false);
    }
    let object_count = objects
        .len()
        .checked_add(pdf.objects.len())
        .ok_or_else(|| iwa_error(part, "iWork embedded PDF object count overflows"))?;
    if object_count > limits.max_document_objects {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "iWork embedded PDF exceeds the configured object limit",
        )
        .in_part(part));
    }
    let transform = AffineTransform {
        a: bounds.width / source_unit.width * parent_coordinate_scale,
        b: 0.0,
        c: 0.0,
        d: bounds.height / source_unit.height * parent_coordinate_scale,
        e: bounds.x
            * if visual_coordinates_are_final {
                parent_coordinate_scale
            } else {
                1.0
            },
        f: bounds.y
            * if visual_coordinates_are_final {
                parent_coordinate_scale
            } else {
                1.0
            },
    };
    let font_families = pdf
        .embedded_fonts
        .iter_mut()
        .map(|font| {
            let authored = font.family.clone();
            font.family = iwork_pdf_font_family(part, identifier, &authored);
            (authored, font.family.clone())
        })
        .collect::<HashMap<_, _>>();
    for (index, mut source) in pdf.objects.into_iter().enumerate() {
        rename_iwork_pdf_visual_fonts(&mut source.visual, &font_families);
        let numeric_id = u32::try_from(objects.len())
            .ok()
            .and_then(|value| value.checked_add(1))
            .ok_or_else(|| iwa_error(part, "Pages embedded PDF object identifier overflows"))?;
        objects.push(Object {
            numeric_id,
            parent_numeric_id: None,
            stable_id: format!("iwork-{unit_index}-{identifier}-pdf-{index}"),
            parent_stable_id: None,
            kind: source.kind,
            unit_index,
            bounds: if visual_coordinates_are_final {
                source.bounds
            } else {
                Rect {
                    x: bounds.x + source.bounds.x * bounds.width / source_unit.width,
                    y: bounds.y + source.bounds.y * bounds.height / source_unit.height,
                    width: source.bounds.width * bounds.width / source_unit.width,
                    height: source.bounds.height * bounds.height / source_unit.height,
                }
            },
            z: base_z.saturating_add(source.z),
            text: source.text,
            source: SourceRef {
                part: part.to_owned(),
                mapping: MappingQuality::Exact,
                locator: SourceLocator::Iwork {
                    kind: source_kind,
                    component: if visual_coordinates_are_final {
                        format!("embedded-pdf-archive-{identifier}")
                    } else {
                        format!("archive-{identifier}")
                    },
                },
            },
            visual: Visual::Layer {
                transform,
                opacity: 1.0,
                blend_mode: crate::model::BlendMode::Normal,
                visual: Box::new(source.visual),
            },
        });
    }
    for font in pdf.embedded_fonts {
        if !embedded_fonts.iter().any(|existing| {
            existing.family == font.family
                && existing.style == font.style
                && existing.weight == font.weight
        }) {
            embedded_fonts.push(font);
        }
    }
    Ok(true)
}

fn iwork_pdf_font_family(part: &str, identifier: u64, family: &str) -> String {
    format!(
        "iWork PDF {:08x} {identifier} {family}",
        crate::zip::crc32(part.as_bytes())
    )
}

fn rename_iwork_pdf_visual_fonts(visual: &mut Visual, families: &HashMap<String, String>) {
    match visual {
        Visual::Text { font_family, .. } => {
            if let Some(replacement) = families.get(font_family) {
                *font_family = replacement.clone();
            }
        }
        Visual::RichText { runs, .. } => {
            for run in runs {
                if let Some(replacement) = families.get(&run.font_family) {
                    run.font_family = replacement.clone();
                }
            }
        }
        Visual::Media { poster: visual, .. }
        | Visual::ImageColorChange { visual, .. }
        | Visual::ImageAdjustment { visual, .. }
        | Visual::ColorManagedImage { visual, .. }
        | Visual::Layer { visual, .. }
        | Visual::Effect { visual, .. }
        | Visual::TextLayout { visual, .. }
        | Visual::TextEffects { visual, .. }
        | Visual::StrokeStyle { visual, .. }
        | Visual::AdvancedEffect { visual, .. } => {
            rename_iwork_pdf_visual_fonts(visual, families);
        }
        Visual::Group { children } => {
            for child in children {
                rename_iwork_pdf_visual_fonts(&mut child.visual, families);
            }
        }
        Visual::OpacityMask { visual, .. } => rename_iwork_pdf_visual_fonts(visual, families),
        Visual::None
        | Visual::Shape { .. }
        | Visual::Image { .. }
        | Visual::ImageWithFallback { .. }
        | Visual::MaskedImage { .. }
        | Visual::PaintedShape { .. } => {}
    }
}

fn pages_table_left(
    storage: &KeynoteTextStorage,
    archives: &[IwaArchive],
    character_index: usize,
    attachment: PagesAttachment,
    layout: PagesDocumentLayout,
    table_width: f32,
) -> Result<f32, Diagnostic> {
    let align = iwork_inherited_style_at(&storage.paragraph_styles, character_index)
        .map(|identifier| keynote_text_style(archives, identifier, DOCUMENT_COMPONENT))
        .transpose()?
        .map_or(TextAlign::Start, |style| style.align);
    let content_width = (layout.width - layout.margin_left - layout.margin_right).max(1.0);
    let x = match attachment.horizontal_offset_type {
        0 => {
            layout.margin_left + (content_width - table_width) / 2.0 + attachment.horizontal_offset
        }
        _ => match align {
            TextAlign::Center => {
                layout.margin_left
                    + (content_width - table_width) / 2.0
                    + attachment.horizontal_offset
            }
            TextAlign::End => {
                layout.width - layout.margin_right - table_width - attachment.horizontal_offset
            }
            TextAlign::Start
            | TextAlign::Justify
            | TextAlign::Distribute
            | TextAlign::MediumKashida
            | TextAlign::HighKashida
            | TextAlign::LowKashida
            | TextAlign::ThaiDistribute => layout.margin_left + attachment.horizontal_offset,
        },
    };
    Ok(x.clamp(0.0, (layout.width - table_width).max(0.0)))
}

#[allow(clippy::too_many_arguments)]
fn pages_table_top(
    storage: &KeynoteTextStorage,
    archives: &[IwaArchive],
    range_start: usize,
    attachment_index: usize,
    layout: PagesDocumentLayout,
    target_y: f32,
    messages: &NumbersMessageSpace,
    limits: Limits,
) -> Result<f32, Diagnostic> {
    let runs = pages_text_runs_for_range(
        storage,
        archives,
        DOCUMENT_COMPONENT,
        range_start,
        attachment_index,
        limits,
    )?;
    let paragraphs = pages_paragraph_layouts(
        storage,
        &runs,
        archives,
        DOCUMENT_COMPONENT,
        range_start,
        attachment_index,
        messages,
        limits,
        layout.view_scale,
        layout.font_metrics,
        false,
    )?;
    let height = keynote_text_layout_height(
        &runs,
        &paragraphs,
        (layout.width - layout.margin_left - layout.margin_right).max(1.0),
    );
    Ok((target_y + height).clamp(target_y, layout.height - layout.margin_bottom))
}

fn pages_table_column_widths(
    messages: &NumbersMessageSpace,
    model: &IwaMessage,
    model_part: &str,
    columns: u32,
    table_width: f32,
) -> Result<Vec<f32>, Diagnostic> {
    let Some(store) = numbers_first_bytes(&model.payload, 4, model_part)? else {
        return pages_table_header_sizes(messages, &[], columns, table_width, model_part);
    };
    let Some(header_reference) = numbers_references(store, 2, model_part)?.into_iter().next()
    else {
        return pages_table_header_sizes(messages, &[], columns, table_width, model_part);
    };
    pages_table_header_sizes(
        messages,
        &[header_reference],
        columns,
        table_width,
        model_part,
    )
}

fn pages_table_row_heights(
    messages: &NumbersMessageSpace,
    model: &IwaMessage,
    model_part: &str,
    rows: u32,
    table_height: f32,
) -> Result<Vec<f32>, Diagnostic> {
    let Some(store) = numbers_first_bytes(&model.payload, 4, model_part)? else {
        return pages_table_header_sizes(messages, &[], rows, table_height, model_part);
    };
    let references = numbers_first_bytes(store, 1, model_part)?
        .map(|headers| numbers_references(headers, 2, model_part))
        .transpose()?
        .unwrap_or_default();
    pages_table_header_sizes(messages, &references, rows, table_height, model_part)
}

fn iwork_table_image_row_heights(
    archives: &[IwaArchive],
    cells: &[NumbersCell],
    heights: &mut [f32],
    limits: Limits,
) -> Result<bool, Diagnostic> {
    let mut found = false;
    for cell in cells {
        let Some(height) = heights.get_mut(cell.row as usize) else {
            continue;
        };
        let Some(storage_identifier) = cell.text_storage else {
            continue;
        };
        let Some(message) = archive_message(
            archives,
            storage_identifier,
            IWORK_TEXT_STORAGE_TYPE,
            &cell.part,
        )?
        else {
            continue;
        };
        let storage = keynote_storage_text(&message.payload, &cell.part, limits)?;
        for change in storage.attachments {
            let Some(identifier) = change.identifier else {
                continue;
            };
            let Some(drawable) = pages_attachment_drawable(archives, identifier, &cell.part)?
            else {
                continue;
            };
            let Some(image) = archive_message(archives, drawable, IWORK_IMAGE_TYPE, &cell.part)?
            else {
                continue;
            };
            *height = height.max(keynote_image(&image.payload, &cell.part)?.geometry.height + 6.0);
            found = true;
        }
    }
    Ok(found)
}

fn pages_table_header_sizes(
    messages: &NumbersMessageSpace,
    header_references: &[u64],
    count: u32,
    total: f32,
    part: &str,
) -> Result<Vec<f32>, Diagnostic> {
    let count = usize::try_from(count)
        .map_err(|_| iwa_error(part, "Pages table header count is not addressable"))?;
    let fallback = || vec![total / count as f32; count];
    let Some(mut sizes) = iwork_table_header_sizes(messages, header_references, count)?
        .into_iter()
        .collect::<Option<Vec<_>>>()
    else {
        return Ok(fallback());
    };
    let authored_total = sizes.iter().sum::<f32>();
    if !authored_total.is_finite() || authored_total <= 0.0 {
        return Ok(fallback());
    }
    let scale = total / authored_total;
    for size in &mut sizes {
        *size *= scale;
    }
    Ok(sizes)
}

fn iwork_table_header_sizes(
    messages: &NumbersMessageSpace,
    header_references: &[u64],
    count: usize,
) -> Result<Vec<Option<f32>>, Diagnostic> {
    let mut sizes = vec![None; count];
    for reference in header_references {
        let Some((header_part, header)) =
            messages.message(*reference, NUMBERS_HEADER_STORAGE_TYPE)?
        else {
            continue;
        };
        for entry in numbers_bytes(&header.payload, 2, header_part)? {
            let Some(index) = numbers_varint(entry, 1, header_part)?
                .and_then(|value| usize::try_from(value).ok())
            else {
                continue;
            };
            if index >= sizes.len() {
                continue;
            }
            sizes[index] = keynote_fixed32_field(entry, 2, header_part)?
                .filter(|size| size.is_finite() && *size > 0.0);
        }
    }
    Ok(sizes)
}

fn pages_table_text_style(
    model: &IwaMessage,
    model_part: &str,
    archives: &[IwaArchive],
    row: u32,
    column: u32,
    rows: u32,
) -> Result<KeynoteTextStyle, Diagnostic> {
    let header_rows = numbers_varint(&model.payload, 9, model_part)?
        .and_then(|value| u32::try_from(value).ok())
        .unwrap_or(0);
    let header_columns = numbers_varint(&model.payload, 10, model_part)?
        .and_then(|value| u32::try_from(value).ok())
        .unwrap_or(0);
    let footer_rows = numbers_varint(&model.payload, 11, model_part)?
        .and_then(|value| u32::try_from(value).ok())
        .unwrap_or(0);
    let style_field = if row < header_rows {
        25
    } else if column < header_columns {
        26
    } else if row >= rows.saturating_sub(footer_rows) {
        27
    } else {
        24
    };
    let authored = numbers_references(&model.payload, style_field, model_part)?
        .into_iter()
        .next()
        .map(|identifier| keynote_text_style(archives, identifier, model_part))
        .transpose()?;
    let has_authored = authored.is_some();
    let mut style = authored.unwrap_or_else(|| KeynoteTextStyle {
        font_family: "Helvetica".to_owned(),
        font_size: 10.5,
        ..KeynoteTextStyle::default()
    });
    if !has_authored
        && (row < header_rows || column < header_columns || row >= rows.saturating_sub(footer_rows))
    {
        style.align = TextAlign::Center;
        style.bold = true;
    }
    Ok(style)
}

fn iwork_table_text_styles(
    messages: &NumbersMessageSpace,
    archives: &[IwaArchive],
    model: &IwaMessage,
    model_part: &str,
    limits: Limits,
) -> Result<Vec<Option<KeynoteTextStyle>>, Diagnostic> {
    let Some(store) = numbers_first_bytes(&model.payload, 4, model_part)? else {
        return Ok(Vec::new());
    };
    let Some(list_reference) = numbers_references(store, 5, model_part)?.into_iter().next() else {
        return Ok(Vec::new());
    };
    let Some((list_part, list)) = messages.message(list_reference, NUMBERS_DATA_LIST_TYPE)? else {
        return Ok(Vec::new());
    };
    let mut styles = Vec::new();
    for entry in numbers_bytes(&list.payload, 3, list_part)? {
        let Some(index) =
            numbers_varint(entry, 1, list_part)?.and_then(|value| usize::try_from(value).ok())
        else {
            continue;
        };
        if index >= limits.max_document_objects {
            return Err(iwa_error(
                list_part,
                "iWork cell text-style index exceeds the configured object limit",
            ));
        }
        let Some(reference) = numbers_references(entry, 4, list_part)?.into_iter().next() else {
            continue;
        };
        if archive_message(archives, reference, IWORK_PARAGRAPH_STYLE_TYPE, list_part)?.is_none()
            && archive_message(archives, reference, IWORK_CHARACTER_STYLE_TYPE, list_part)?
                .is_none()
        {
            continue;
        }
        if styles.len() <= index {
            styles.resize(index + 1, None);
        }
        let style = keynote_text_style(archives, reference, list_part)?;
        styles[index] = Some(style);
    }
    Ok(styles)
}

fn keynote_table_cell_styles(
    messages: &NumbersMessageSpace,
    stylesheet_archives: &[IwaArchive],
    model: &IwaMessage,
    model_part: &str,
    width: f32,
    height: f32,
    limits: Limits,
) -> Result<Vec<IworkTableCellStyle>, Diagnostic> {
    let Some(store) = numbers_first_bytes(&model.payload, 4, model_part)? else {
        return Ok(Vec::new());
    };
    let Some(list_reference) = numbers_references(store, 5, model_part)?.into_iter().next() else {
        return Ok(Vec::new());
    };
    let Some((list_part, list)) = messages.message(list_reference, NUMBERS_DATA_LIST_TYPE)? else {
        return Ok(Vec::new());
    };
    let mut styles = Vec::new();
    for entry in numbers_bytes(&list.payload, 3, list_part)? {
        let Some(index) =
            numbers_varint(entry, 1, list_part)?.and_then(|value| usize::try_from(value).ok())
        else {
            continue;
        };
        if index >= limits.max_document_objects {
            return Err(iwa_error(
                list_part,
                "Numbers cell-style index exceeds the configured object limit",
            ));
        }
        let Some(reference) = numbers_references(entry, 4, list_part)?.into_iter().next() else {
            continue;
        };
        if styles.len() <= index {
            styles.resize(index + 1, IworkTableCellStyle::default());
        }
        styles[index] = keynote_table_cell_style(
            stylesheet_archives,
            reference,
            list_part,
            width,
            height,
            &mut Vec::new(),
        )?;
    }
    Ok(styles)
}

fn keynote_table_cell_fills(
    messages: &NumbersMessageSpace,
    stylesheet_archives: &[IwaArchive],
    model: &IwaMessage,
    model_part: &str,
    width: f32,
    height: f32,
    limits: Limits,
) -> Result<Vec<Paint>, Diagnostic> {
    keynote_table_cell_styles(
        messages,
        stylesheet_archives,
        model,
        model_part,
        width,
        height,
        limits,
    )
    .map(|styles| styles.into_iter().map(|style| style.fill).collect())
}

fn keynote_table_cell_fill(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
    width: f32,
    height: f32,
    path: &mut Vec<u64>,
) -> Result<Paint, Diagnostic> {
    keynote_table_cell_style(archives, identifier, part, width, height, path)
        .map(|style| style.fill)
}

fn keynote_table_cell_style(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
    width: f32,
    height: f32,
    path: &mut Vec<u64>,
) -> Result<IworkTableCellStyle, Diagnostic> {
    if path.contains(&identifier) {
        return Err(iwa_error(
            part,
            "Keynote table cell style inheritance is cyclic",
        ));
    }
    if path.len() >= 32 {
        return Err(iwa_error(
            part,
            "Keynote table cell style inheritance is too deep",
        ));
    }
    let Some(style) = archive_message(archives, identifier, 6_004, part)? else {
        return Ok(IworkTableCellStyle::default());
    };
    path.push(identifier);
    let mut result = IworkTableCellStyle::default();
    if let Some(super_style) = numbers_first_bytes(&style.payload, 1, part)?
        && let Some(reference) = numbers_references(super_style, 3, part)?.into_iter().next()
    {
        result = keynote_table_cell_style(archives, reference, part, width, height, path)?;
    }
    if let Some(properties) = numbers_first_bytes(&style.payload, 11, part)? {
        if let Some(authored_fill) = numbers_first_bytes(properties, 1, part)? {
            result.fill = keynote_fill(authored_fill, part, width, height)?;
        }
        if let Some(wrap) = numbers_varint(properties, 3, part)? {
            result.wrap = wrap != 0;
        }
    }
    path.pop();
    Ok(result)
}

fn keynote_table_banded_fills(
    archives: &[IwaArchive],
    model: &IwaMessage,
    part: &str,
    width: f32,
    height: f32,
) -> Result<Option<(Paint, Paint)>, Diagnostic> {
    let Some(body_style) = numbers_references(&model.payload, 18, part)?
        .into_iter()
        .next()
    else {
        return Ok(None);
    };
    let Some(table_style) = numbers_references(&model.payload, 3, part)?
        .into_iter()
        .next()
    else {
        return Ok(None);
    };
    let body_fill =
        keynote_table_cell_fill(archives, body_style, part, width, height, &mut Vec::new())?;
    let mut enabled = false;
    let mut banded_fill = Paint::None;
    let mut header_column_divider = false;
    let mut header_row_divider = false;
    keynote_apply_table_style_properties(
        archives,
        table_style,
        part,
        width,
        height,
        &mut enabled,
        &mut banded_fill,
        &mut header_column_divider,
        &mut header_row_divider,
        &mut Vec::new(),
    )?;
    Ok((enabled && !matches!(banded_fill, Paint::None)).then_some((body_fill, banded_fill)))
}

#[allow(clippy::too_many_arguments)]
fn keynote_apply_table_style_properties(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
    width: f32,
    height: f32,
    enabled: &mut bool,
    fill: &mut Paint,
    header_column_divider: &mut bool,
    header_row_divider: &mut bool,
    path: &mut Vec<u64>,
) -> Result<(), Diagnostic> {
    if path.len() >= 32 || path.contains(&identifier) {
        return Err(iwa_error(
            part,
            "Keynote table style inheritance is cyclic or too deep",
        ));
    }
    let Some(style) = archive_message(archives, identifier, 6_003, part)? else {
        return Ok(());
    };
    path.push(identifier);
    if let Some(super_style) = numbers_first_bytes(&style.payload, 1, part)?
        && let Some(parent) = numbers_references(super_style, 3, part)?.into_iter().next()
    {
        keynote_apply_table_style_properties(
            archives,
            parent,
            part,
            width,
            height,
            enabled,
            fill,
            header_column_divider,
            header_row_divider,
            path,
        )?;
    }
    if let Some(properties) = numbers_first_bytes(&style.payload, 11, part)? {
        if let Some(value) = numbers_varint(properties, 1, part)? {
            *enabled = value != 0;
        }
        if let Some(authored_fill) = numbers_first_bytes(properties, 2, part)? {
            *fill = keynote_fill(authored_fill, part, width, height)?;
        }
        if let Some(value) = numbers_varint(properties, 42, part)? {
            *header_column_divider = value != 0;
        }
        if let Some(value) = numbers_varint(properties, 43, part)? {
            *header_row_divider = value != 0;
        }
    }
    path.pop();
    Ok(())
}

fn keynote_table_header_dividers(
    archives: &[IwaArchive],
    model: &IwaMessage,
    part: &str,
) -> Result<(bool, bool), Diagnostic> {
    let Some(table_style) = numbers_references(&model.payload, 3, part)?
        .into_iter()
        .next()
    else {
        return Ok((false, false));
    };
    let mut enabled = false;
    let mut fill = Paint::None;
    let mut header_column_divider = false;
    let mut header_row_divider = false;
    keynote_apply_table_style_properties(
        archives,
        table_style,
        part,
        1.0,
        1.0,
        &mut enabled,
        &mut fill,
        &mut header_column_divider,
        &mut header_row_divider,
        &mut Vec::new(),
    )?;
    Ok((header_column_divider, header_row_divider))
}

fn pages_table_trailing_merges(
    model: &IwaMessage,
    model_part: &str,
    rows: u32,
    columns: u32,
    cells: &[NumbersCell],
) -> Result<Vec<NumbersMerge>, Diagnostic> {
    let header_rows = numbers_varint(&model.payload, 9, model_part)?
        .and_then(|value| u32::try_from(value).ok())
        .unwrap_or(0);
    let footer_rows = numbers_varint(&model.payload, 11, model_part)?
        .and_then(|value| u32::try_from(value).ok())
        .unwrap_or(0);
    let mut body_style_counts = HashMap::<u32, usize>::new();
    for cell in cells.iter().filter(|cell| {
        cell.row >= header_rows
            && cell.row < rows.saturating_sub(footer_rows)
            && cell.text.is_none()
    }) {
        if let Some(style) = cell.style_index {
            *body_style_counts.entry(style).or_default() += 1;
        }
    }
    let body_style = body_style_counts
        .into_iter()
        .max_by_key(|(_, count)| *count)
        .map(|(style, _)| style);
    let mut merges = Vec::new();
    for row in 0..rows {
        let Some(source) = cells
            .iter()
            .filter(|cell| cell.row == row && cell.text.is_some())
            .max_by_key(|cell| cell.column)
        else {
            continue;
        };
        if source.column + 1 >= columns || (row >= header_rows && source.style_index == body_style)
        {
            continue;
        }
        let trailing = cells
            .iter()
            .filter(|cell| cell.row == row && cell.column > source.column)
            .collect::<Vec<_>>();
        if trailing.len() != usize::try_from(columns - source.column - 1).unwrap_or(usize::MAX)
            || trailing
                .iter()
                .any(|cell| cell.text.is_some() || cell.style_index != source.style_index)
        {
            continue;
        }
        merges.push(NumbersMerge {
            row,
            start_column: source.column,
            end_column: columns,
        });
    }
    Ok(merges)
}

fn pages_table_grouped_header_label(
    cells: &[NumbersCell],
    row: u32,
    column: u32,
    rows: u32,
) -> bool {
    let Some(next_row) = row.checked_add(1).filter(|next_row| *next_row < rows) else {
        return false;
    };
    cells
        .iter()
        .any(|cell| cell.row == row && cell.column == column && cell.text.is_some())
        && cells
            .iter()
            .all(|cell| cell.row != next_row || cell.column != column || cell.text.is_none())
}

#[allow(clippy::too_many_arguments)]
fn pages_push_table_grid(
    objects: &mut Vec<Object>,
    limits: Limits,
    unit_index: u32,
    identifier: u64,
    bounds: Rect,
    model: &IwaMessage,
    model_part: &str,
    rows: u32,
    column_widths: &[f32],
    row_heights: &[f32],
    cells: &[NumbersCell],
    text_style_archives: &[IwaArchive],
    cell_fills: &[Paint],
    cell_text_styles: &[Option<KeynoteTextStyle>],
    banded_fills: Option<&(Paint, Paint)>,
    merges: &[NumbersMerge],
    borders: Option<&[NumbersTableBorder]>,
    row_start: u32,
) -> Result<Vec<(u64, Rect)>, Diagnostic> {
    let mut attachments = Vec::new();
    if rows == 0 || column_widths.is_empty() || row_heights.len() != rows as usize {
        return Ok(attachments);
    }
    let header_rows = numbers_varint(&model.payload, 9, model_part)?
        .and_then(|value| u32::try_from(value).ok())
        .unwrap_or(0);
    let footer_rows = numbers_varint(&model.payload, 11, model_part)?
        .and_then(|value| u32::try_from(value).ok())
        .unwrap_or(0);
    let (header_column_divider, header_row_divider) =
        keynote_table_header_dividers(text_style_archives, model, model_part)?;
    let mut columns = Vec::with_capacity(column_widths.len() + 1);
    columns.push(0.0);
    for width in column_widths {
        columns.push(columns.last().copied().unwrap_or(0.0) + width);
    }
    if let Some(last) = columns.last_mut() {
        *last = bounds.width;
    }
    let mut row_offsets = Vec::with_capacity(row_heights.len() + 1);
    row_offsets.push(0.0);
    for height in row_heights {
        row_offsets.push(row_offsets.last().copied().unwrap_or(0.0) + height);
    }
    if let Some(last) = row_offsets.last_mut() {
        *last = bounds.height;
    }
    pages_push_object(
        objects,
        limits,
        unit_index,
        identifier,
        ObjectKind::Table,
        bounds,
        40,
        None,
        "table",
        MappingQuality::Approximate,
        Visual::StrokeStyle {
            style: StrokeStyle {
                alignment: LineAlignment::Inset,
                ..StrokeStyle::default()
            },
            visual: Box::new(Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: Paint::Solid(0xffff_ffff),
                stroke: if borders.is_some() {
                    Paint::None
                } else {
                    Paint::Solid(0x0000_00ff)
                },
                stroke_width: if borders.is_some() { 0.0 } else { 2.0 },
            }),
        },
    )?;
    for (serial, cell) in cells.iter().enumerate() {
        let Some(mut fill) = cell
            .style_index
            .and_then(|index| usize::try_from(index).ok())
            .and_then(|index| cell_fills.get(index))
            .filter(|fill| !matches!(fill, Paint::None))
            .cloned()
        else {
            continue;
        };
        let Ok(column) = usize::try_from(cell.column) else {
            continue;
        };
        let Ok(row) = usize::try_from(cell.row) else {
            continue;
        };
        if column + 1 >= columns.len() || row + 1 >= row_offsets.len() {
            continue;
        }
        if cell.row >= header_rows
            && cell.row < rows.saturating_sub(footer_rows)
            && (cell.row - header_rows) % 2 == 1
            && let Some((body_fill, banded_fill)) = banded_fills
            && fill == *body_fill
        {
            fill = banded_fill.clone();
        }
        pages_push_object(
            objects,
            limits,
            unit_index,
            identifier.saturating_add(1_000 + serial as u64),
            ObjectKind::Cell,
            Rect {
                x: bounds.x + columns[column],
                y: bounds.y + row_offsets[row],
                width: columns[column + 1] - columns[column],
                height: row_offsets[row + 1] - row_offsets[row],
            },
            41,
            None,
            "table-cell",
            MappingQuality::Exact,
            Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill,
                stroke: Paint::None,
                stroke_width: 0.0,
            },
        )?;
    }
    if borders.is_none() {
        for (serial, x) in columns
            .iter()
            .copied()
            .skip(1)
            .take(columns.len().saturating_sub(2))
            .enumerate()
        {
            pages_push_table_line(
                objects,
                limits,
                unit_index,
                identifier.saturating_add(serial as u64 + 1),
                Rect {
                    x: bounds.x + x,
                    y: bounds.y,
                    width: 0.6,
                    height: bounds.height,
                },
                true,
                header_column_divider && serial == 0,
            )?;
        }
        for (serial, y) in row_offsets
            .iter()
            .copied()
            .skip(1)
            .take(row_offsets.len().saturating_sub(2))
            .enumerate()
        {
            let row = serial as u32;
            if header_column_divider && pages_table_grouped_header_label(cells, row, 0, rows) {
                continue;
            }
            pages_push_table_line(
                objects,
                limits,
                unit_index,
                identifier.saturating_add(100 + serial as u64),
                Rect {
                    x: bounds.x,
                    y: bounds.y + y,
                    width: bounds.width,
                    height: 0.6,
                },
                false,
                header_row_divider && serial == 0,
            )?;
        }
    }
    for (serial, merge) in merges.iter().enumerate() {
        let Some(source) = cells
            .iter()
            .find(|cell| cell.row == merge.row && cell.column == merge.start_column)
        else {
            continue;
        };
        let Ok(row) = usize::try_from(merge.row) else {
            continue;
        };
        let Ok(start_column) = usize::try_from(merge.start_column) else {
            continue;
        };
        let Ok(end_column) = usize::try_from(merge.end_column) else {
            continue;
        };
        if row + 1 >= row_offsets.len()
            || start_column >= columns.len()
            || end_column >= columns.len()
        {
            continue;
        }
        let fill = source
            .style_index
            .and_then(|index| usize::try_from(index).ok())
            .and_then(|index| cell_fills.get(index))
            .filter(|fill| !matches!(fill, Paint::None))
            .cloned()
            .unwrap_or(Paint::Solid(0xffff_ffff));
        pages_push_object(
            objects,
            limits,
            unit_index,
            identifier.saturating_add(400 + serial as u64),
            ObjectKind::Cell,
            Rect {
                x: bounds.x + columns[start_column],
                y: bounds.y + row_offsets[row],
                width: columns[end_column] - columns[start_column],
                height: row_offsets[row + 1] - row_offsets[row],
            },
            42,
            None,
            "table-cell",
            MappingQuality::Derived,
            Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill,
                stroke: if borders.is_some() {
                    Paint::None
                } else {
                    Paint::Solid(0xd1d1_d1ff)
                },
                stroke_width: if borders.is_some() { 0.0 } else { 0.6 },
            },
        )?;
    }
    for (serial, border) in borders.unwrap_or_default().iter().enumerate() {
        let Some((mut line_bounds, visual)) =
            iwork_table_border_visual(border, &columns, &row_offsets, row_start)
        else {
            continue;
        };
        line_bounds.x += bounds.x;
        line_bounds.y += bounds.y;
        pages_push_object(
            objects,
            limits,
            unit_index,
            identifier.saturating_add(serial as u64),
            ObjectKind::Shape,
            line_bounds,
            43,
            None,
            "table-grid",
            MappingQuality::Exact,
            visual,
        )?;
    }
    for (serial, cell) in cells.iter().enumerate() {
        let Some(text) = cell.text.as_ref() else {
            continue;
        };
        let column_index = cell.column;
        let Ok(column) = usize::try_from(column_index) else {
            continue;
        };
        if cell.row >= rows || column + 1 >= columns.len() {
            continue;
        }
        let row_index = usize::try_from(cell.row)
            .map_err(|_| iwa_error(model_part, "Pages table row is not addressable"))?;
        let row_height = row_offsets[row_index + 1] - row_offsets[row_index];
        let x = columns[column];
        let end_column = merges
            .iter()
            .find(|merge| merge.row == cell.row && merge.start_column == cell.column)
            .and_then(|merge| usize::try_from(merge.end_column).ok())
            .filter(|end| *end < columns.len())
            .unwrap_or(column + 1);
        let width = columns[end_column] - x;
        let mut style = cell
            .text_style_index
            .and_then(|index| usize::try_from(index).ok())
            .and_then(|index| cell_text_styles.get(index))
            .and_then(Clone::clone)
            .map(Ok)
            .unwrap_or_else(|| {
                pages_table_text_style(
                    model,
                    model_part,
                    text_style_archives,
                    cell.row,
                    column_index,
                    rows,
                )
            })?;
        if borders.is_none() && end_column > column + 1 && cell.row >= header_rows {
            style.align = TextAlign::End;
            style.color = 0xffff_ffff;
            style.bold = true;
        }
        let footer = cell.row >= rows.saturating_sub(footer_rows);
        let vertical_padding = if footer { 9.0 } else { 3.0 };
        let runs = vec![keynote_text_run(iwork_visible_table_text(text), &style)];
        let grouped_header_label = borders.is_none()
            && header_column_divider
            && column_index == 0
            && pages_table_grouped_header_label(cells, cell.row, column_index, rows);
        let (text_row, text_height) = if grouped_header_label {
            (
                row_index,
                row_offsets[row_index + 2] - row_offsets[row_index],
            )
        } else {
            (row_index, row_height)
        };
        let visual = Visual::RichText {
            geometry: Geometry::Rectangle,
            fill: Paint::None,
            stroke: Paint::None,
            stroke_width: 0.0,
            align: style.align,
            line_height: style
                .line_height_multiple
                .map_or(style.font_size * 1.2, |multiple| {
                    if multiple > 4.0 {
                        multiple
                    } else {
                        style.font_size * multiple
                    }
                })
                .max(style.font_size),
            runs,
        };
        pages_push_object(
            objects,
            limits,
            unit_index,
            identifier.saturating_add(200 + serial as u64),
            ObjectKind::Cell,
            Rect {
                x: bounds.x + x + 4.0,
                y: bounds.y
                    + row_offsets[text_row]
                    + if grouped_header_label {
                        0.0
                    } else {
                        vertical_padding
                    },
                width: (width - 8.0).max(1.0),
                height: if grouped_header_label {
                    text_height
                } else {
                    row_height - vertical_padding - 3.0
                }
                .max(1.0),
            },
            43,
            Some(text.clone()),
            "table-cell",
            MappingQuality::Approximate,
            if grouped_header_label {
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
                    visual: Box::new(visual),
                }
            } else {
                visual
            },
        )?;
        if let Some(storage) = cell.text_storage {
            attachments.push((storage, objects.last().unwrap().bounds));
        }
    }
    Ok(attachments)
}

#[allow(clippy::too_many_arguments)]
fn iwork_table_cell_images(
    diagnostics: &mut Vec<Diagnostic>,
    package: &Package<'_>,
    archives: &[IwaArchive],
    stylesheet_archives: &[IwaArchive],
    messages: &NumbersMessageSpace,
    data_files: &[(u64, String)],
    cells: &[(u64, Rect)],
    unit_index: u32,
    limits: Limits,
    objects: &mut Vec<Object>,
    embedded_fonts: &mut Vec<EmbeddedFont>,
    embedded_pdf_scale: f32,
) -> Result<(), Diagnostic> {
    for (identifier, bounds) in cells {
        let Some((part, message)) = messages.message(*identifier, IWORK_TEXT_STORAGE_TYPE)? else {
            continue;
        };
        let storage = keynote_storage_text(&message.payload, part, limits)?;
        let mut x = bounds.x;
        for change in &storage.attachments {
            let Some(attachment_identifier) = change.identifier else {
                continue;
            };
            let result: Result<usize, Diagnostic> = (|| {
                let Some(attachment) = pages_attachment(archives, attachment_identifier, part)?
                else {
                    return Ok(1);
                };
                // Cell text attachments share image semantics, but not floating table/group layout.
                let Some(image) =
                    archive_message(archives, attachment.drawable, IWORK_IMAGE_TYPE, part)?
                else {
                    return Ok(1);
                };
                let geometry = keynote_image(&image.payload, part)?.geometry;
                let omitted = keynote_native_drawable(
                    diagnostics,
                    package,
                    archives,
                    stylesheet_archives,
                    messages,
                    data_files,
                    attachment.drawable,
                    unit_index,
                    part,
                    limits,
                    objects,
                    &mut Vec::new(),
                    bounds.width,
                    bounds.height,
                    x + attachment.horizontal_offset,
                    bounds.y,
                    44,
                    true,
                    embedded_fonts,
                    embedded_pdf_scale,
                    None,
                )?;
                x += geometry.width;
                Ok(omitted)
            })();
            match result {
                Ok(0) => {}
                Err(error)
                    if !matches!(
                        error.code,
                        DiagnosticCode::FormatInvalid
                            | DiagnosticCode::UnsupportedFormat
                            | DiagnosticCode::ZipCrcMismatch
                            | DiagnosticCode::ZipDeflateInvalid
                    ) =>
                {
                    return Err(error);
                }
                result => diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Render,
                        Fidelity::Omitted,
                        format!(
                            "IWORK_CELL_ATTACHMENT_OMITTED: attachment {attachment_identifier}: {}",
                            result.err().map_or_else(
                                || "missing or unsupported drawable".to_owned(),
                                |error| error.message
                            )
                        ),
                    )
                    .in_part(part),
                ),
            }
        }
    }
    Ok(())
}

fn pages_push_table_line(
    objects: &mut Vec<Object>,
    limits: Limits,
    unit_index: u32,
    identifier: u64,
    bounds: Rect,
    vertical: bool,
    strong: bool,
) -> Result<(), Diagnostic> {
    let commands = if vertical {
        vec![
            PathCommand::MoveTo { x: 0.0, y: 0.0 },
            PathCommand::LineTo {
                x: 0.0,
                y: bounds.height,
            },
        ]
    } else {
        vec![
            PathCommand::MoveTo { x: 0.0, y: 0.0 },
            PathCommand::LineTo {
                x: bounds.width,
                y: 0.0,
            },
        ]
    };
    pages_push_object(
        objects,
        limits,
        unit_index,
        identifier,
        ObjectKind::Shape,
        bounds,
        41,
        None,
        "table-grid",
        MappingQuality::Approximate,
        Visual::PaintedShape {
            geometry: Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands,
            },
            fill: Paint::None,
            stroke: Paint::Solid(if strong { 0x0000_00ff } else { 0xd1d1_d1ff }),
            stroke_width: if strong { 2.0 } else { 0.6 },
        },
    )
}

#[allow(clippy::too_many_arguments)]
fn pages_push_object(
    objects: &mut Vec<Object>,
    limits: Limits,
    unit_index: u32,
    identifier: u64,
    kind: ObjectKind,
    bounds: Rect,
    z: i32,
    text: Option<String>,
    source_kind: &'static str,
    mapping: MappingQuality,
    visual: Visual,
) -> Result<(), Diagnostic> {
    if objects.len() >= limits.max_document_objects {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "Pages native object count exceeds the configured limit",
        )
        .in_part(DOCUMENT_COMPONENT));
    }
    let numeric_id = u32::try_from(objects.len())
        .ok()
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| {
            iwa_error(
                DOCUMENT_COMPONENT,
                "Pages native object identifier overflows",
            )
        })?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("iwork-pages-{unit_index}-{source_kind}-{identifier}"),
        parent_stable_id: None,
        kind,
        unit_index,
        bounds,
        z,
        text,
        source: SourceRef {
            part: DOCUMENT_COMPONENT.to_owned(),
            mapping,
            locator: SourceLocator::Iwork {
                kind: source_kind,
                component: format!("archive-{identifier}"),
            },
        },
        visual,
    });
    Ok(())
}

fn keynote_preview_document(
    package: &Package<'_>,
    document_archives: &[IwaArchive],
    root_preview: &IworkPreview,
    limits: Limits,
) -> Result<Option<Document>, Diagnostic> {
    let root = root_message(document_archives, DOCUMENT_COMPONENT)?;
    let Some(show_identifier) = keynote_document_show_reference(&root.payload)? else {
        return Ok(None);
    };
    let Some(show) = archive_message(
        document_archives,
        show_identifier,
        KEYNOTE_SHOW_TYPE,
        DOCUMENT_COMPONENT,
    )?
    else {
        return Ok(None);
    };
    let show = keynote_show(&show.payload)?;
    let mut slide_nodes = show.slide_nodes;
    let legacy_slide_components = slide_nodes.is_empty();
    if legacy_slide_components {
        slide_nodes = legacy_keynote_slide_identifiers(package, limits)?;
    }
    if slide_nodes.is_empty() {
        return Ok(None);
    }

    let data_files = iwork_data_files(package, limits)?;
    let mut stylesheet_archives = Vec::new();
    let mut stylesheet_components: Vec<String> = package
        .entry_names()
        .filter(|part| keynote_stylesheet_component(part))
        .map(str::to_owned)
        .collect();
    stylesheet_components
        .sort_by_key(|component| usize::from(component.starts_with("Index/DocumentStylesheet")));
    for component in stylesheet_components {
        let bytes = package.required_part(&component)?;
        stylesheet_archives.extend(parse_iwa_archives(package, &component, &bytes, limits)?);
    }
    let mut object_archives = Vec::new();
    for component in package
        .entry_names()
        .filter(|part| iwork_object_container_component(part))
    {
        let bytes = package.required_part(component)?;
        object_archives.extend(parse_iwa_archives(package, component, &bytes, limits)?);
    }
    let mut message_archives = document_archives
        .iter()
        .cloned()
        .map(|archive| NumbersArchive {
            part: DOCUMENT_COMPONENT.to_owned(),
            archive,
        })
        .collect::<Vec<_>>();
    let table_components = package
        .entry_names()
        .filter(|part| {
            *part == CALCULATION_ENGINE_COMPONENT
                || (part.starts_with("Index/Tables/") && part.ends_with(".iwa"))
        })
        .map(str::to_owned)
        .collect::<Vec<_>>();
    for component in table_components {
        let bytes = package.required_part(&component)?;
        message_archives.extend(
            parse_iwa_archives(package, &component, &bytes, limits)?
                .into_iter()
                .map(|archive| NumbersArchive {
                    part: component.clone(),
                    archive,
                }),
        );
    }
    let messages = NumbersMessageSpace {
        archives: message_archives,
    };
    let native_point_size = show.size.is_some();
    let embedded_pdf_scale = if native_point_size {
        IWORK_POINT_TO_CSS_PIXEL
    } else {
        1.0
    };
    let (width, height) = show
        .size
        .unwrap_or((root_preview.width, root_preview.height));
    let mut units = Vec::new();
    let mut objects = Vec::new();
    let mut embedded_fonts = Vec::new();
    units.try_reserve_exact(slide_nodes.len()).map_err(|_| {
        Diagnostic::fatal(
            DiagnosticCode::AllocationFailed,
            Phase::Parse,
            None,
            "unable to allocate the bounded Keynote slide list",
        )
        .in_part(DOCUMENT_COMPONENT)
    })?;
    objects.try_reserve_exact(slide_nodes.len()).map_err(|_| {
        Diagnostic::fatal(
            DiagnosticCode::AllocationFailed,
            Phase::Parse,
            None,
            "unable to allocate the bounded Keynote slide previews",
        )
        .in_part(DOCUMENT_COMPONENT)
    })?;

    let mut missing_previews = 0_usize;
    let mut diagnostics = Vec::new();
    let mut native = KeynoteNativeStats::default();
    for (index, slide_node_identifier) in slide_nodes.iter().copied().enumerate() {
        let unit_index = u32::try_from(index)
            .map_err(|_| iwa_error(DOCUMENT_COMPONENT, "Keynote slide index exceeds u32"))?;
        units.push(Unit {
            kind: UnitKind::Slide,
            index: unit_index,
            id: format!("iwork-slide-{slide_node_identifier}"),
            name: format!("Slide {}", index + 1),
            width,
            height,
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
        });

        let slide_node = archive_message(
            document_archives,
            slide_node_identifier,
            KEYNOTE_SLIDE_NODE_TYPE,
            DOCUMENT_COMPONENT,
        )?;
        if let Some(slide_identifier) = if legacy_slide_components {
            Some(slide_node_identifier)
        } else {
            slide_node
                .map(|message| keynote_slide_reference(&message.payload))
                .transpose()?
                .flatten()
        } && let Some(part) = keynote_slide_component(package, slide_identifier, index)?
        {
            let bytes = package.required_part(&part)?;
            let mut archives = parse_iwa_archives(package, &part, &bytes, limits)?;
            archives.extend_from_slice(&object_archives);
            let before = objects.len();
            let unsupported = keynote_native_slide_objects(
                &mut diagnostics,
                package,
                &archives,
                &stylesheet_archives,
                &messages,
                &data_files,
                slide_identifier,
                unit_index,
                width,
                height,
                &part,
                limits,
                &mut objects,
                &mut embedded_fonts,
                embedded_pdf_scale,
            )?;
            if objects.len() > before {
                native.slides += 1;
                native.unsupported_drawables =
                    native.unsupported_drawables.saturating_add(unsupported);
                continue;
            }
        }

        let preview = if index == 0 {
            Some((
                root_preview.part.clone(),
                root_preview.media_type,
                root_preview.bytes.clone(),
            ))
        } else {
            let part = slide_node.and_then(|message| {
                message.data_references.iter().find_map(|identifier| {
                    data_files.iter().find_map(|(current, part)| {
                        (*current == *identifier && iwork_slide_preview_part(part))
                            .then(|| part.clone())
                    })
                })
            });
            match part {
                Some(part) if package.has_part(&part) => {
                    let bytes = package.required_part(&part)?;
                    let (media_type, _, _) = iwork_image_properties(&part, &bytes, limits)?;
                    Some((part, media_type, bytes.into_vec()))
                }
                _ => None,
            }
        };

        let Some((part, media_type, bytes)) = preview else {
            missing_previews += 1;
            continue;
        };
        objects.push(Object {
            numeric_id: u32::try_from(objects.len())
                .ok()
                .and_then(|value| value.checked_add(1))
                .ok_or_else(|| iwa_error(DOCUMENT_COMPONENT, "Keynote object id overflows"))?,
            parent_numeric_id: None,
            stable_id: format!("iwork-slide-preview-{slide_node_identifier}"),
            parent_stable_id: None,
            kind: ObjectKind::Image,
            unit_index,
            bounds: Rect {
                x: 0.0,
                y: 0.0,
                width,
                height,
            },
            z: 0,
            text: None,
            source: SourceRef {
                part,
                mapping: MappingQuality::Exact,
                locator: SourceLocator::Iwork {
                    kind: "slide-preview",
                    component: format!("slide-node-{slide_node_identifier}"),
                },
            },
            visual: Visual::Image {
                media_type: media_type.to_owned(),
                bytes,
                crop: ImageCrop::default(),
            },
        });
    }

    if native.slides == 0 {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Approximate,
                "IWORK_KEYNOTE_PREVIEW_ONLY: showing embedded slide previews; native text, search, object inspection, builds, animations, and transitions are unavailable",
            )
            .in_part(DOCUMENT_COMPONENT),
        );
    } else {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Approximate,
                format!(
                    "IWORK_KEYNOTE_NATIVE_STATIC_SUBSET: rendered {} slides from native text and image objects; themes, advanced styling, masks, charts, builds, animations, and transitions may be approximate or unavailable",
                    native.slides
                ),
            )
            .in_part(DOCUMENT_COMPONENT),
        );
        if native.unsupported_drawables != 0 {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Render,
                    Fidelity::Approximate,
                    format!(
                        "IWORK_KEYNOTE_DRAWABLES_UNSUPPORTED: {} native drawable references were not rendered",
                        native.unsupported_drawables
                    ),
                )
                .in_part(DOCUMENT_COMPONENT),
            );
        }
    }
    if missing_previews != 0 {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Approximate,
                format!(
                    "IWORK_SLIDE_PREVIEW_MISSING: {missing_previews} Keynote slides have no supported embedded preview"
                ),
            )
            .in_part(DOCUMENT_COMPONENT),
        );
    }
    let mut document = Document {
        fatal: false,
        format: Some(DocumentFormat::Keynote),
        kind: Some(DocumentKind::Presentation),
        units,
        outline: Vec::new(),
        objects,
        embedded_fonts,
        font_alternate_names: Vec::new(),
        diagnostics,
    };
    if native_point_size {
        scale_iwork_point_document(&mut document);
    }
    Ok(Some(document))
}

fn keynote_slide_reference(bytes: &[u8]) -> Result<Option<u64>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut slide = None;
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                DOCUMENT_COMPONENT,
                "Keynote slide node contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                DOCUMENT_COMPONENT,
                format!("invalid Keynote slide-node field key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (2, 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    DOCUMENT_COMPONENT,
                    "Keynote slide reference",
                )?;
                let identifier = parse_iwa_reference(reference)?;
                if slide.replace(identifier).is_some() {
                    return Err(iwa_error(
                        DOCUMENT_COMPONENT,
                        "Keynote slide node repeats its slide reference",
                    ));
                }
            }
            (2, _) => {
                return Err(iwa_error(
                    DOCUMENT_COMPONENT,
                    "Keynote slide reference has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, DOCUMENT_COMPONENT)?,
        }
    }
    Ok(slide)
}

fn keynote_slide_component(
    package: &Package<'_>,
    slide_identifier: u64,
    slide_index: usize,
) -> Result<Option<String>, Diagnostic> {
    let part = keynote_object_component(package, "Slide", slide_identifier)?;
    Ok(part.or_else(|| {
        let legacy = format!("Index/Slide{}.iwa", slide_index + 1);
        package.has_part(&legacy).then_some(legacy)
    }))
}

fn legacy_keynote_slide_identifiers(
    package: &Package<'_>,
    limits: Limits,
) -> Result<Vec<u64>, Diagnostic> {
    let mut parts = package
        .entry_names()
        .filter_map(|part| {
            part.strip_prefix("Index/Slide")
                .and_then(|suffix| suffix.strip_suffix(".iwa"))
                .filter(|suffix| {
                    !suffix.is_empty() && suffix.bytes().all(|byte| byte.is_ascii_digit())
                })
                .and_then(|suffix| suffix.parse::<usize>().ok())
                .map(|index| (index, part.to_owned()))
        })
        .collect::<Vec<_>>();
    parts.sort_unstable_by_key(|(index, _)| *index);

    let mut identifiers = Vec::new();
    for (_, part) in parts {
        let bytes = package.required_part(&part)?;
        let archives = parse_iwa_archives(package, &part, &bytes, limits)?;
        let root = root_message(&archives, &part)?;
        if root.message_type == KEYNOTE_SLIDE_TYPE {
            identifiers.push(archives[0].identifier);
        }
    }
    Ok(identifiers)
}

fn keynote_object_component(
    package: &Package<'_>,
    component: &str,
    identifier: u64,
) -> Result<Option<String>, Diagnostic> {
    let exact = format!("Index/{component}-{identifier}.iwa");
    if package.has_part(&exact) {
        return Ok(Some(exact));
    }
    let prefix = format!("Index/{component}-{identifier}-");
    let mut found = None;
    for part in package
        .entry_names()
        .filter(|part| part.starts_with(&prefix) && part.ends_with(".iwa"))
    {
        if found.replace(part.to_owned()).is_some() {
            return Err(iwa_error(
                DOCUMENT_COMPONENT,
                format!("Keynote {component} identifier resolves to multiple components"),
            ));
        }
    }
    Ok(found.or_else(|| {
        let fallback = format!("Index/{component}.iwa");
        package.has_part(&fallback).then_some(fallback)
    }))
}

fn keynote_stylesheet_component(part: &str) -> bool {
    [THEME_STYLESHEET_COMPONENT, STYLESHEET_COMPONENT]
        .into_iter()
        .any(|component| {
            part == component
                || component.strip_suffix(".iwa").is_some_and(|prefix| {
                    part.strip_prefix(prefix)
                        .and_then(|suffix| suffix.strip_prefix('-'))
                        .and_then(|suffix| suffix.strip_suffix(".iwa"))
                        .is_some_and(|identifier| !identifier.is_empty())
                })
        })
}

fn iwork_object_container_component(part: &str) -> bool {
    part == "Index/ObjectContainer.iwa"
        || (part.starts_with("Index/ObjectContainer-") && part.ends_with(".iwa"))
}

#[allow(clippy::too_many_arguments)]
fn keynote_native_slide_objects(
    diagnostics: &mut Vec<Diagnostic>,
    package: &Package<'_>,
    archives: &[IwaArchive],
    stylesheet_archives: &[IwaArchive],
    messages: &NumbersMessageSpace,
    data_files: &[(u64, String)],
    slide_identifier: u64,
    unit_index: u32,
    slide_width: f32,
    slide_height: f32,
    part: &str,
    limits: Limits,
    objects: &mut Vec<Object>,
    embedded_fonts: &mut Vec<EmbeddedFont>,
    embedded_pdf_scale: f32,
) -> Result<usize, Diagnostic> {
    let Some(slide) = archive_message(archives, slide_identifier, KEYNOTE_SLIDE_TYPE, part)? else {
        return Ok(0);
    };
    let slide_start = objects.len();
    let mut master_range = slide_start..slide_start;
    let mut unsupported = 0_usize;
    if let Some(master_identifier) = keynote_slide_master_reference(&slide.payload, part)?
        && let Some(master_part) =
            keynote_object_component(package, "TemplateSlide", master_identifier)?
    {
        let bytes = package.required_part(&master_part)?;
        let master_archives = parse_iwa_archives(package, &master_part, &bytes, limits)?;
        if let Some(master) = archive_message(
            &master_archives,
            master_identifier,
            KEYNOTE_SLIDE_TYPE,
            &master_part,
        )? {
            unsupported = unsupported.saturating_add(keynote_native_slide_archive_objects(
                diagnostics,
                package,
                &master_archives,
                stylesheet_archives,
                messages,
                data_files,
                &master.payload,
                unit_index,
                slide_width,
                slide_height,
                &master_part,
                limits,
                objects,
                i32::MIN,
                -1_000_000,
                "slide-background",
                false,
                embedded_fonts,
                embedded_pdf_scale,
            )?);
            master_range = slide_start..objects.len();
        }
    }
    let slide_object_start = objects.len();
    unsupported = unsupported.saturating_add(keynote_native_slide_archive_objects(
        diagnostics,
        package,
        archives,
        stylesheet_archives,
        messages,
        data_files,
        &slide.payload,
        unit_index,
        slide_width,
        slide_height,
        part,
        limits,
        objects,
        i32::MIN.saturating_add(1),
        0,
        "slide-background",
        true,
        embedded_fonts,
        embedded_pdf_scale,
    )?);
    let slide_text_bounds = objects[slide_object_start..]
        .iter()
        .filter(|object| object.kind == ObjectKind::TextBox && object.text.is_some())
        .map(|object| object.bounds)
        .collect::<Vec<_>>();
    for object in &mut objects[master_range] {
        if object.kind == ObjectKind::TextBox
            && slide_text_bounds
                .iter()
                .any(|bounds| keynote_bounds_overlap(object.bounds, *bounds))
        {
            object.text = None;
            object.visual = Visual::None;
        }
    }
    if objects.len() == slide_start {
        objects.truncate(slide_start);
    }
    Ok(unsupported)
}

fn keynote_bounds_overlap(left: Rect, right: Rect) -> bool {
    let intersection_width = (left.x + left.width).min(right.x + right.width) - left.x.max(right.x);
    let intersection_height =
        (left.y + left.height).min(right.y + right.height) - left.y.max(right.y);
    intersection_width > left.width.min(right.width) * 0.8
        && intersection_height > left.height.min(right.height) * 0.8
}

#[allow(clippy::too_many_arguments)]
fn keynote_native_slide_archive_objects(
    diagnostics: &mut Vec<Diagnostic>,
    package: &Package<'_>,
    archives: &[IwaArchive],
    stylesheet_archives: &[IwaArchive],
    messages: &NumbersMessageSpace,
    data_files: &[(u64, String)],
    slide: &[u8],
    unit_index: u32,
    slide_width: f32,
    slide_height: f32,
    part: &str,
    limits: Limits,
    objects: &mut Vec<Object>,
    background_z: i32,
    drawable_z: i32,
    background_kind: &'static str,
    render_placeholder_content: bool,
    embedded_fonts: &mut Vec<EmbeddedFont>,
    embedded_pdf_scale: f32,
) -> Result<usize, Diagnostic> {
    if let Some(style_identifier) = keynote_slide_style_reference(slide, part)?
        && let Some(fill) = keynote_slide_background(
            package,
            stylesheet_archives,
            data_files,
            style_identifier,
            STYLESHEET_COMPONENT,
            slide_width,
            slide_height,
            limits,
        )?
    {
        let kind = if matches!(fill, Visual::Image { .. }) {
            ObjectKind::Image
        } else {
            ObjectKind::Shape
        };
        push_keynote_native_object(
            objects,
            limits,
            style_identifier,
            unit_index,
            part,
            background_kind,
            kind,
            Rect {
                x: 0.0,
                y: 0.0,
                width: slide_width,
                height: slide_height,
            },
            background_z,
            None,
            MappingQuality::Exact,
            fill,
        )?;
    }
    let drawables = keynote_slide_drawables(slide, part)?;
    let mut path = Vec::new();
    let mut unsupported = 0_usize;
    for (z, identifier) in drawables.into_iter().enumerate() {
        unsupported = unsupported.saturating_add(keynote_native_drawable(
            diagnostics,
            package,
            archives,
            stylesheet_archives,
            messages,
            data_files,
            identifier,
            unit_index,
            part,
            limits,
            objects,
            &mut path,
            slide_width,
            slide_height,
            0.0,
            0.0,
            drawable_z.saturating_add(i32::try_from(z).unwrap_or(i32::MAX)),
            render_placeholder_content,
            embedded_fonts,
            embedded_pdf_scale,
            None,
        )?);
    }
    Ok(unsupported)
}

fn keynote_slide_master_reference(bytes: &[u8], part: &str) -> Result<Option<u64>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote slide master key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (17, 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote master slide reference",
                )?;
                return Ok(Some(parse_iwa_reference(reference)?));
            }
            (17, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote master slide reference has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(None)
}

fn keynote_slide_style_reference(bytes: &[u8], part: &str) -> Result<Option<u64>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote slide-style key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote slide-style reference",
                )?;
                return Ok(Some(parse_iwa_reference(reference)?));
            }
            (1, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote slide-style reference has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(None)
}

#[allow(clippy::too_many_arguments)]
fn keynote_slide_background(
    package: &Package<'_>,
    archives: &[IwaArchive],
    data_files: &[(u64, String)],
    style_identifier: u64,
    part: &str,
    width: f32,
    height: f32,
    limits: Limits,
) -> Result<Option<Visual>, Diagnostic> {
    let Some(style) = archive_message(archives, style_identifier, KEYNOTE_SLIDE_STYLE_TYPE, part)?
    else {
        return Ok(None);
    };
    let mut cursor = 0_usize;
    while cursor < style.payload.len() {
        let (key, consumed) = read_varint(&style.payload[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote slide style key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (11, 2) => {
                let properties = read_iwa_length_delimited(
                    &style.payload,
                    &mut cursor,
                    part,
                    "Keynote slide-style properties",
                )?;
                let mut properties_cursor = 0_usize;
                while properties_cursor < properties.len() {
                    let (properties_key, consumed) = read_varint(&properties[properties_cursor..])
                        .map_err(|message| {
                            iwa_error(
                                part,
                                format!("invalid Keynote slide properties key: {message}"),
                            )
                        })?;
                    properties_cursor += consumed;
                    match (properties_key >> 3, properties_key & 7) {
                        (1, 2) => {
                            let fill = read_iwa_length_delimited(
                                properties,
                                &mut properties_cursor,
                                part,
                                "Keynote slide fill",
                            )?;
                            return keynote_background_visual(
                                package, data_files, fill, part, width, height, limits,
                            );
                        }
                        (_, wire) => {
                            skip_protobuf_value(properties, &mut properties_cursor, wire, part)?
                        }
                    }
                }
            }
            (_, wire) => skip_protobuf_value(&style.payload, &mut cursor, wire, part)?,
        }
    }
    Ok(None)
}

fn keynote_background_visual(
    package: &Package<'_>,
    data_files: &[(u64, String)],
    bytes: &[u8],
    part: &str,
    width: f32,
    height: f32,
    limits: Limits,
) -> Result<Option<Visual>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote background fill key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (3, 2) => {
                let image_fill = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote background image fill",
                )?;
                let identifier = keynote_reference_field(image_fill, 6, part)?
                    .or(keynote_reference_field(image_fill, 1, part)?)
                    .or(keynote_reference_field(image_fill, 7, part)?)
                    .or(keynote_reference_field(image_fill, 5, part)?);
                let Some(identifier) = identifier else {
                    return Ok(None);
                };
                let Some(data_part) = data_files.iter().find_map(|(current, data_part)| {
                    (*current == identifier).then_some(data_part.as_str())
                }) else {
                    return Ok(None);
                };
                if !package.has_part(data_part) {
                    return Ok(None);
                }
                let image = package.required_part(data_part)?;
                let (media_type, _, _) = iwork_image_properties(data_part, &image, limits)?;
                return Ok(Some(Visual::Image {
                    media_type: media_type.to_owned(),
                    bytes: image.into_vec(),
                    crop: ImageCrop::default(),
                }));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    let fill = keynote_fill(bytes, part, width, height)?;
    Ok(
        (!matches!(fill, Paint::None)).then_some(Visual::PaintedShape {
            geometry: Geometry::Rectangle,
            fill,
            stroke: Paint::None,
            stroke_width: 0.0,
        }),
    )
}

fn keynote_reference_field(
    bytes: &[u8],
    field: u64,
    part: &str,
) -> Result<Option<u64>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote reference key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (current, 2) if current == field => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote object reference",
                )?;
                return Ok(Some(parse_iwa_reference(reference)?));
            }
            (current, _) if current == field => {
                return Err(iwa_error(
                    part,
                    "Keynote object reference has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(None)
}

fn keynote_image_fill(
    package: &Package<'_>,
    data_files: &[(u64, String)],
    identifier: u64,
    tile: bool,
    tile_size: Option<(f32, f32)>,
    limits: Limits,
) -> Result<Option<Paint>, Diagnostic> {
    let Some(part) = data_files
        .iter()
        .find_map(|(current, part)| (*current == identifier).then_some(part))
    else {
        return Ok(None);
    };
    if !package.has_part(part) {
        return Ok(None);
    }
    let bytes = package.required_part(part)?;
    let media_type = match iwork_image_properties(part, &bytes, limits) {
        Ok((media_type, _, _)) => media_type,
        Err(_) => match office_image_media_type(part, &bytes) {
            Ok(media_type) => media_type,
            Err(_) => return Ok(None),
        },
    };
    Ok(Some(Paint::Image {
        mapping: None,
        media_type: media_type.to_owned(),
        bytes: bytes.into_vec(),
        crop: ImageCrop::default(),
        tile,
        tile_width: tile_size.map(|(width, _)| width),
        tile_height: tile_size.map(|(_, height)| height),
    }))
}

fn keynote_materialize_image_fill(
    diagnostics: &mut Vec<Diagnostic>,
    package: &Package<'_>,
    data_files: &[(u64, String)],
    style: &mut KeynoteDrawableStyle,
    limits: Limits,
    part: &str,
) -> Result<(), Diagnostic> {
    if let Some((identifier, tile, tile_size)) = style.image_fill {
        if let Some(fill) =
            keynote_image_fill(package, data_files, identifier, tile, tile_size, limits)?
        {
            style.fill = fill;
        } else {
            diagnostics.push(Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Approximate,
                format!("IWORK_IMAGE_FILL_UNAVAILABLE: image data {identifier} is unavailable; retaining the authored reference color when present"),
            ).in_part(part));
        }
    }
    Ok(())
}

fn keynote_slide_drawables(bytes: &[u8], part: &str) -> Result<Vec<u64>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut owned = Vec::new();
    let mut z_order = Vec::new();
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(part, "Keynote slide contains too many fields"));
        }
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote slide key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (field @ (7 | 42), 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote drawable reference",
                )?;
                let target = if field == 42 {
                    &mut z_order
                } else {
                    &mut owned
                };
                target.push(parse_iwa_reference(reference)?);
            }
            (7 | 42, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote drawable reference has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(if z_order.is_empty() { owned } else { z_order })
}

#[allow(clippy::too_many_arguments)]
fn keynote_native_drawable(
    diagnostics: &mut Vec<Diagnostic>,
    package: &Package<'_>,
    archives: &[IwaArchive],
    stylesheet_archives: &[IwaArchive],
    messages: &NumbersMessageSpace,
    data_files: &[(u64, String)],
    identifier: u64,
    unit_index: u32,
    part: &str,
    limits: Limits,
    objects: &mut Vec<Object>,
    path: &mut Vec<u64>,
    slide_width: f32,
    slide_height: f32,
    offset_x: f32,
    offset_y: f32,
    z: i32,
    render_placeholder_content: bool,
    embedded_fonts: &mut Vec<EmbeddedFont>,
    embedded_pdf_scale: f32,
    pages_layout: Option<PagesDocumentLayout<'_>>,
) -> Result<usize, Diagnostic> {
    if path.len() >= limits.max_xml_depth || path.contains(&identifier) {
        return Ok(1);
    }
    path.push(identifier);
    let result = if let Some(group) = archive_message(archives, identifier, IWORK_GROUP_TYPE, part)?
    {
        let (geometry, children) = keynote_group(&group.payload, part)?;
        let first_object = objects.len();
        let mut unsupported = 0_usize;
        for child in children {
            unsupported = unsupported.saturating_add(keynote_native_drawable(
                diagnostics,
                package,
                archives,
                stylesheet_archives,
                messages,
                data_files,
                child,
                unit_index,
                part,
                limits,
                objects,
                path,
                slide_width,
                slide_height,
                offset_x + geometry.x,
                offset_y + geometry.y,
                z,
                render_placeholder_content,
                embedded_fonts,
                embedded_pdf_scale,
                pages_layout,
            )?);
        }
        keynote_align_group_text_to_shapes(objects, first_object);
        Ok(unsupported)
    } else if let Some(chart) = archive_message(archives, identifier, IWORK_CHART_TYPE, part)? {
        let Some(chart) = keynote_chart(&chart.payload, archives, stylesheet_archives, part)?
        else {
            path.pop();
            return Ok(1);
        };
        keynote_push_chart(
            chart, identifier, unit_index, part, limits, objects, offset_x, offset_y, z,
        )?;
        Ok(0)
    } else if let Some(shape) = keynote_shape_payload(archives, identifier, part)? {
        let (geometry, style_identifier, shape_geometry) = keynote_shape(shape, part)?;
        let bounds = keynote_bounds(geometry, offset_x, offset_y, None);
        let mut style = style_identifier
            .map(|style_identifier| {
                keynote_drawable_style(
                    stylesheet_archives,
                    style_identifier,
                    STYLESHEET_COMPONENT,
                    bounds.width,
                    bounds.height,
                )
            })
            .transpose()?
            .unwrap_or_default();
        keynote_materialize_image_fill(diagnostics, package, data_files, &mut style, limits, part)?;
        let visual = keynote_wrap_geometry_transform(
            geometry,
            bounds,
            keynote_wrap_style(
                &style,
                None,
                Visual::PaintedShape {
                    geometry: shape_geometry,
                    fill: style.fill.clone(),
                    stroke: style.stroke.clone(),
                    stroke_width: style.stroke_width,
                },
            ),
        );
        push_keynote_native_object(
            objects,
            limits,
            identifier,
            unit_index,
            part,
            "shape",
            ObjectKind::Shape,
            bounds,
            z,
            None,
            MappingQuality::Exact,
            visual,
        )?;
        Ok(0)
    } else if let Some(media) = archive_message(archives, identifier, IWORK_MEDIA_TYPE, part)? {
        let drawable = keynote_nested_message(&media.payload, 1, part)?.unwrap_or_default();
        let geometry = keynote_drawable_geometry(drawable, part)?;
        let bounds = keynote_bounds(geometry, offset_x, offset_y, None);
        let mut media_data = None;
        let mut poster = None;
        for data_identifier in &media.data_references {
            let Some(data_part) = data_files.iter().find_map(|(current, data_part)| {
                (*current == *data_identifier).then_some(data_part)
            }) else {
                continue;
            };
            if !package.has_part(data_part) {
                continue;
            }
            let bytes = package.required_part(data_part)?;
            if let Ok((media_type, _, _)) = iwork_image_properties(data_part, &bytes, limits) {
                poster.get_or_insert_with(|| Visual::Image {
                    media_type: media_type.to_owned(),
                    bytes: bytes.into_vec(),
                    crop: ImageCrop::default(),
                });
            } else if let Ok((kind, media_type)) = embedded_media_type(data_part, &bytes, None) {
                media_data.get_or_insert_with(|| (kind, media_type.to_owned(), bytes.into_vec()));
            }
        }
        let Some((kind, media_type, bytes)) = media_data else {
            path.pop();
            return Ok(1);
        };
        let style_identifier = keynote_reference_field(drawable, 2, part)?;
        let style = style_identifier
            .map(|style_identifier| {
                keynote_media_style(stylesheet_archives, style_identifier, STYLESHEET_COMPONENT)
            })
            .transpose()?
            .unwrap_or_default();
        let visual = keynote_wrap_style(
            &style,
            None,
            Visual::Media {
                kind,
                media_type,
                bytes,
                poster: Box::new(poster.unwrap_or(Visual::None)),
            },
        );
        push_keynote_native_object(
            objects,
            limits,
            identifier,
            unit_index,
            part,
            "media",
            ObjectKind::Image,
            bounds,
            z,
            None,
            MappingQuality::Exact,
            visual,
        )?;
        Ok(0)
    } else if let Some(image_message) =
        archive_message(archives, identifier, IWORK_IMAGE_TYPE, part)?
    {
        // ImageArchive flags bit 0 marks a media placeholder, not master artwork.
        if !render_placeholder_content
            && keynote_varint_field(&image_message.payload, 7, part)?.unwrap_or(0) & 1 != 0
        {
            path.pop();
            return Ok(0);
        }
        let image = keynote_image(&image_message.payload, part)?;
        let mask = image
            .mask_identifier
            .map(|mask_identifier| {
                archive_message(archives, mask_identifier, IWORK_MASK_TYPE, part)?
                    .map(|mask| keynote_mask(&mask.payload, identifier, part))
                    .transpose()
            })
            .transpose()?
            .flatten();
        let display_geometry = mask.as_ref().map_or(image.geometry, |mask| {
            if mask.relative_to_image {
                KeynoteGeometry {
                    x: image.geometry.x + mask.geometry.x,
                    y: image.geometry.y + mask.geometry.y,
                    width: mask.geometry.width,
                    height: mask.geometry.height,
                    ..image.geometry
                }
            } else {
                mask.geometry
            }
        });
        // Captions are independent text, including when the image data is absent.
        // TSA.CaptionInfoArchive wraps the same TSWP shape used by ordinary text boxes.
        let caption_result = (|| -> Result<(), Diagnostic> {
            let Some(drawable) = keynote_nested_message(&image_message.payload, 1, part)? else {
                return Ok(());
            };
            if keynote_varint_field(drawable, 13, part)?.unwrap_or(0) != 0 {
                return Ok(());
            }
            let Some(caption_id) = keynote_reference_field(drawable, 11, part)? else {
                return Ok(());
            };
            let Some(caption) = archive_message(archives, caption_id, IWORK_CAPTION_TYPE, part)?
            else {
                return Ok(());
            };
            let shape = keynote_nested_message(&caption.payload, 1, part)?
                .ok_or_else(|| iwa_error(part, "caption has no text shape"))?;
            let placement = keynote_reference_field(&caption.payload, 2, part)?
                .map(|id| archive_message(archives, id, IWORK_CAPTION_PLACEMENT_TYPE, part))
                .transpose()?
                .flatten();
            let anchors = placement
                .map(|placement| {
                    Ok::<_, Diagnostic>((
                        keynote_varint_field(&placement.payload, 1, part)?,
                        keynote_varint_field(&placement.payload, 2, part)?,
                    ))
                })
                .transpose()?;
            // ponytail: bottom captions only; retain text below for other placements
            // until top/side anchors have native-reference coverage.
            if anchors != Some((Some(1), Some(7))) || display_geometry.rotation_degrees != 0.0 {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Layout,
                        Fidelity::Approximate,
                        "IWORK_CAPTION_PLACEMENT_APPROXIMATE: caption retained below its image",
                    )
                    .in_part(part),
                );
            }
            keynote_native_text_object(
                diagnostics,
                messages,
                package,
                archives,
                stylesheet_archives,
                data_files,
                caption_id,
                shape,
                unit_index,
                part,
                limits,
                objects,
                slide_width,
                slide_height,
                offset_x,
                offset_y,
                z,
                pages_layout,
                Some(display_geometry),
            )?;
            Ok(())
        })();
        if let Err(error) = caption_result {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Layout,
                    Fidelity::Approximate,
                    format!("IWORK_CAPTION_UNAVAILABLE: {}", error.message),
                )
                .in_part(part),
            );
        }
        let bounds = keynote_bounds(display_geometry, offset_x, offset_y, None);
        let crop = mask
            .as_ref()
            .map(|mask| {
                if mask.relative_to_image {
                    keynote_relative_image_crop(image.geometry, mask.geometry)
                } else {
                    keynote_image_crop(image.geometry, mask.geometry)
                }
            })
            .unwrap_or_default();
        let style = image
            .style_identifier
            .map(|style_identifier| {
                keynote_media_style(stylesheet_archives, style_identifier, STYLESHEET_COMPONENT)
            })
            .transpose()?
            .unwrap_or_default();
        if let Some((pdf_part, pdf_bytes)) = pages_pdf_data(package, data_files, &image)? {
            let first = objects.len();
            let final_pdf_visual = display_geometry.rotation_degrees.abs() <= f32::EPSILON
                && mask.is_none()
                && style.stroke_style.dash.is_empty()
                && style.stroke_style.cap == LineCap::Flat
                && style.stroke_style.join == LineJoin::Miter
                && style.shadow.is_none()
                && style.reflection.is_none()
                && style.opacity >= 1.0;
            if iwork_push_pdf_objects(
                &pdf_part,
                &pdf_bytes,
                identifier,
                unit_index,
                bounds,
                z,
                "image",
                embedded_pdf_scale,
                final_pdf_visual,
                limits,
                objects,
                embedded_fonts,
            )? {
                for object in &mut objects[first..] {
                    let visual = std::mem::replace(&mut object.visual, Visual::None);
                    object.visual = keynote_wrap_geometry_transform(
                        display_geometry,
                        bounds,
                        keynote_wrap_style(
                            &style,
                            mask.as_ref().map(|mask| mask.clip.clone()),
                            visual,
                        ),
                    );
                }
                path.pop();
                return Ok(0);
            }
        }
        let data_part = image
            .data_identifiers
            .iter()
            .chain(&image_message.data_references)
            .find_map(|identifier| {
                data_files.iter().find_map(|(current, data_part)| {
                    (*current == *identifier
                        && matches!(
                            data_part
                                .to_ascii_lowercase()
                                .rsplit_once('.')
                                .map(|(_, extension)| extension),
                            Some("jpg" | "jpeg" | "png" | "tif" | "tiff")
                        ))
                    .then_some(data_part)
                })
            })
            .filter(|data_part| package.has_part(data_part));
        let visual = if let Some(data_part) = data_part {
            let bytes = package.required_part(data_part)?;
            let (media_type, _, _) = iwork_image_properties(data_part, &bytes, limits)?;
            keynote_wrap_style(
                &style,
                mask.map(|mask| mask.clip),
                Visual::Image {
                    media_type: media_type.to_owned(),
                    bytes: bytes.into_vec(),
                    crop,
                },
            )
        } else {
            diagnostics.push(Diagnostic::warning(DiagnosticCode::UnsupportedFeature, Phase::Render,
                Fidelity::Approximate, format!("IWORK_IMAGE_UNAVAILABLE: image {identifier} has no packaged resource; retaining its empty frame"))
                .in_part(part));
            Visual::None
        };
        push_keynote_native_object(
            objects,
            limits,
            identifier,
            unit_index,
            part,
            "image",
            ObjectKind::Image,
            bounds,
            z,
            None,
            MappingQuality::Exact,
            visual,
        )?;
        Ok(usize::from(data_part.is_none()))
    } else if let Some(placeholder) =
        archive_message(archives, identifier, KEYNOTE_PLACEHOLDER_TYPE, part)?
    {
        if !render_placeholder_content {
            path.pop();
            return Ok(0);
        }
        let Some(shape) = keynote_placeholder_shape(&placeholder.payload, part)? else {
            path.pop();
            return Ok(1);
        };
        let rendered = keynote_native_text_object(
            diagnostics,
            messages,
            package,
            archives,
            stylesheet_archives,
            data_files,
            identifier,
            shape,
            unit_index,
            part,
            limits,
            objects,
            slide_width,
            slide_height,
            offset_x,
            offset_y,
            z,
            pages_layout,
            None,
        )?;
        Ok(usize::from(!rendered))
    } else if let Some(shape) = archive_message(archives, identifier, IWORK_TEXT_SHAPE_TYPE, part)?
    {
        let rendered = keynote_native_text_object(
            diagnostics,
            messages,
            package,
            archives,
            stylesheet_archives,
            data_files,
            identifier,
            &shape.payload,
            unit_index,
            part,
            limits,
            objects,
            slide_width,
            slide_height,
            offset_x,
            offset_y,
            z,
            pages_layout,
            None,
        )?;
        Ok(usize::from(!rendered))
    } else if let Some((table_part, table)) = messages.message(identifier, IWORK_TABLE_TYPE)? {
        let geometry = pages_table_geometry(&table.payload, table_part)?;
        let mut bounds = keynote_bounds(geometry, offset_x, offset_y, None);
        let Some(model_identifier) = numbers_references(&table.payload, 2, table_part)?
            .into_iter()
            .next()
        else {
            path.pop();
            return Ok(1);
        };
        let Some((model_part, model)) =
            messages.message(model_identifier, NUMBERS_TABLE_MODEL_TYPE)?
        else {
            path.pop();
            return Ok(1);
        };
        let (rows, columns, cells) = numbers_table_cells(messages, model, model_part, limits)?;
        let column_widths =
            pages_table_column_widths(messages, model, model_part, columns, bounds.width)?;
        let mut row_heights =
            pages_table_row_heights(messages, model, model_part, rows, bounds.height)?;
        if iwork_table_image_row_heights(archives, &cells, &mut row_heights, limits)? {
            bounds.height = row_heights.iter().sum();
        }
        let cell_fills = keynote_table_cell_fills(
            messages,
            stylesheet_archives,
            model,
            model_part,
            bounds.width,
            bounds.height,
            limits,
        )?;
        let cell_text_styles =
            iwork_table_text_styles(messages, stylesheet_archives, model, model_part, limits)?;
        let banded_fills = keynote_table_banded_fills(
            stylesheet_archives,
            model,
            model_part,
            bounds.width,
            bounds.height,
        )?;
        let merges = pages_table_trailing_merges(model, model_part, rows, columns, &cells)?;
        let first = objects.len();
        let borders = iwork_table_borders(messages, model, model_part, rows, columns, diagnostics);
        let cell_attachments = pages_push_table_grid(
            objects,
            limits,
            unit_index,
            identifier,
            bounds,
            model,
            model_part,
            rows,
            &column_widths,
            &row_heights,
            &cells,
            stylesheet_archives,
            &cell_fills,
            &cell_text_styles,
            banded_fills.as_ref(),
            &merges,
            borders.as_deref(),
            0,
        )?;
        iwork_table_cell_images(
            diagnostics,
            package,
            archives,
            stylesheet_archives,
            messages,
            data_files,
            &cell_attachments,
            unit_index,
            limits,
            objects,
            embedded_fonts,
            embedded_pdf_scale,
        )?;
        for (serial, object) in objects[first..].iter_mut().enumerate() {
            object.z = z.saturating_add(object.z.saturating_sub(40));
            object.stable_id = format!("iwork-keynote-table-{unit_index}-{identifier}-{serial}");
            object.source.part = table_part.to_owned();
        }
        Ok(usize::from(objects.len() == first))
    } else {
        Ok(1)
    };
    path.pop();
    result
}

fn keynote_align_group_text_to_shapes(objects: &mut [Object], first_object: usize) {
    let shapes = objects
        .iter()
        .enumerate()
        .skip(first_object)
        .filter(|(_, object)| object.kind == ObjectKind::Shape)
        .map(|(_, object)| object.bounds)
        .collect::<Vec<_>>();
    let mut assignments = Vec::new();
    for text_index in first_object..objects.len() {
        if objects[text_index].kind != ObjectKind::TextBox
            || !matches!(
                keynote_text_vertical_align(&objects[text_index].visual),
                Some(TextVerticalAlign::Center | TextVerticalAlign::Bottom)
            )
        {
            continue;
        }
        let text = objects[text_index].bounds;
        let text_center_x = text.x + text.width / 2.0;
        let text_center_y = text.y + text.height / 2.0;
        let candidate = shapes
            .iter()
            .filter_map(|shape| {
                let width_ratio = text.width / shape.width.max(f32::EPSILON);
                let height_ratio = text.height / shape.height.max(f32::EPSILON);
                if !(0.35..=1.5).contains(&width_ratio) || !(0.35..=1.5).contains(&height_ratio) {
                    return None;
                }
                let delta_x = (text_center_x - (shape.x + shape.width / 2.0)).abs();
                let delta_y = (text_center_y - (shape.y + shape.height / 2.0)).abs();
                if delta_x > shape.width / 2.0 + 0.01 || delta_y > shape.height / 2.0 + 0.01 {
                    return None;
                }
                Some((*shape, delta_x + delta_y))
            })
            .min_by(|(_, left), (_, right)| left.total_cmp(right))
            .map(|(shape, _)| shape);
        if let Some(shape) = candidate {
            assignments.push((text_index, shape));
        }
    }
    for (text_index, shape) in &assignments {
        if assignments
            .iter()
            .filter(|(_, candidate)| candidate == shape)
            .count()
            == 1
        {
            objects[*text_index].bounds.y = shape.y;
            objects[*text_index].bounds.height = shape.height;
        }
    }
}

fn keynote_text_vertical_align(visual: &Visual) -> Option<TextVerticalAlign> {
    match visual {
        Visual::TextLayout { layout, .. } => Some(layout.vertical_align),
        Visual::Layer { visual, .. }
        | Visual::Effect { visual, .. }
        | Visual::TextEffects { visual, .. }
        | Visual::StrokeStyle { visual, .. }
        | Visual::AdvancedEffect { visual, .. }
        | Visual::ColorManagedImage { visual, .. }
        | Visual::ImageAdjustment { visual, .. }
        | Visual::OpacityMask { visual, .. } => keynote_text_vertical_align(visual),
        Visual::None
        | Visual::Shape { .. }
        | Visual::Text { .. }
        | Visual::Image { .. }
        | Visual::ImageWithFallback { .. }
        | Visual::MaskedImage { .. }
        | Visual::Media { .. }
        | Visual::ImageColorChange { .. }
        | Visual::PaintedShape { .. }
        | Visual::RichText { .. } => None,
        Visual::Group { .. } => None,
    }
}

fn keynote_placeholder_shape<'a>(
    bytes: &'a [u8],
    part: &str,
) -> Result<Option<&'a [u8]>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote placeholder key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                return read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote placeholder shape",
                )
                .map(Some);
            }
            (1, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote placeholder shape has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(None)
}

#[allow(clippy::too_many_arguments)]
fn keynote_native_text_object(
    diagnostics: &mut Vec<Diagnostic>,
    messages: &NumbersMessageSpace,
    package: &Package<'_>,
    archives: &[IwaArchive],
    stylesheet_archives: &[IwaArchive],
    data_files: &[(u64, String)],
    identifier: u64,
    shape: &[u8],
    unit_index: u32,
    part: &str,
    limits: Limits,
    objects: &mut Vec<Object>,
    slide_width: f32,
    slide_height: f32,
    offset_x: f32,
    offset_y: f32,
    z: i32,
    pages_layout: Option<PagesDocumentLayout<'_>>,
    caption_parent: Option<KeynoteGeometry>,
) -> Result<bool, Diagnostic> {
    let (
        mut geometry,
        storage_identifier,
        shape_style_identifier,
        shape_geometry,
        auto_width,
        auto_height,
    ) = keynote_text_shape(shape, part)?;
    if let Some(parent) = caption_parent {
        // Standalone TSA captions use an 8 pt gap from the displayed (masked)
        // image edge, not from the uncropped image or the saved auto-size anchor.
        geometry.x = parent.x;
        geometry.y = parent.y + parent.height + 8.0;
        geometry.width = parent.width;
    }
    let auto_sized = auto_width || auto_height;
    let flow = keynote_reference_field(shape, 3, part)?;
    let saved_target = pages_layout.and_then(|layout| {
        layout
            .frame_targets
            .iter()
            .find(|(frame, _)| *frame == identifier)
            .map(|(_, target)| target)
    });
    let flow_storage = flow
        .filter(|_| saved_target.is_some())
        .map(|identifier| archive_message(archives, identifier, IWORK_TEXT_FLOW_TYPE, part))
        .transpose()?
        .flatten()
        .map(|message| keynote_reference_field(&message.payload, 1, part))
        .transpose()?
        .flatten();
    // Use a shared story only with a validated frame range; otherwise retain
    // per-shape content instead of duplicating the complete story in every frame.
    let storage_identifier = flow_storage.or(storage_identifier);
    let storage = storage_identifier
        .map(|storage_identifier| {
            archive_message(archives, storage_identifier, IWORK_TEXT_STORAGE_TYPE, part)
        })
        .transpose()?
        .flatten()
        .map(|storage| keynote_storage_text(&storage.payload, part, limits))
        .transpose()?
        .unwrap_or_default();
    let flow_target = saved_target.filter(|target| {
            let valid = target.range_start <= target.range_end
                && keynote_utf16_byte_offset(&storage.text, target.range_start).is_some()
                && keynote_utf16_byte_offset(&storage.text, target.range_end).is_some();
            if !valid {
                diagnostics.push(Diagnostic::warning(DiagnosticCode::UnsupportedFeature,
                    Phase::Layout, Fidelity::Approximate,
                    "PAGES_TEXT_FLOW_APPROXIMATE: invalid saved text range; retaining available story content",
                ).in_part(part));
            }
            valid
        });
    let range_start = flow_target.map_or(0, |target| target.range_start);
    let range_end = flow_target.map_or_else(
        || storage.text.encode_utf16().count(),
        |target| target.range_end,
    );
    let first_paragraph_style = iwork_inherited_style_at(&storage.paragraph_styles, range_start);
    let mut fallback_style = first_paragraph_style
        .map(|identifier| keynote_text_style(stylesheet_archives, identifier, STYLESHEET_COMPONENT))
        .transpose()?
        .unwrap_or_default();
    let normalized_text = normalize_keynote_text(&storage.text);
    let lines = normalized_text.lines().count().max(1) as f32;
    if first_paragraph_style.is_none() && geometry.height > 0.0 {
        fallback_style.font_size = (geometry.height / lines * 0.72).clamp(12.0, 54.0);
    }
    let mut runs = if flow_target.is_some() {
        pages_text_runs_for_range(
            &storage,
            stylesheet_archives,
            STYLESHEET_COMPONENT,
            range_start,
            range_end,
            limits,
        )?
    } else {
        keynote_text_runs(
            &storage,
            stylesheet_archives,
            STYLESHEET_COMPONENT,
            &fallback_style,
            limits,
        )?
    };
    let starts_with_list_marker = runs
        .iter()
        .flat_map(|run| run.text.chars())
        .find(|character| !character.is_whitespace())
        .is_some_and(|character| matches!(character, '•' | '◦' | '▪' | '‣'));
    let leading_space_indent = if geometry.height <= 0.0 {
        let indent = keynote_leading_space_indent(&mut runs);
        if starts_with_list_marker { 0.0 } else { indent }
    } else {
        0.0
    };
    let text = keynote_semantic_text(&runs);
    if text.trim().is_empty() {
        let bounds = keynote_bounds(geometry, offset_x, offset_y, None);
        let mut drawable_style = shape_style_identifier
            .map(|style_identifier| {
                keynote_drawable_style(
                    stylesheet_archives,
                    style_identifier,
                    STYLESHEET_COMPONENT,
                    bounds.width,
                    bounds.height,
                )
            })
            .transpose()?
            .unwrap_or_default();
        keynote_materialize_image_fill(
            diagnostics,
            package,
            data_files,
            &mut drawable_style,
            limits,
            part,
        )?;
        if matches!(drawable_style.fill, Paint::None)
            && matches!(drawable_style.stroke, Paint::None)
        {
            return Ok(false);
        }
        let visual = keynote_wrap_geometry_transform(
            geometry,
            bounds,
            keynote_wrap_style(
                &drawable_style,
                None,
                Visual::PaintedShape {
                    geometry: shape_geometry,
                    fill: drawable_style.fill.clone(),
                    stroke: drawable_style.stroke.clone(),
                    stroke_width: drawable_style.stroke_width,
                },
            ),
        );
        push_keynote_native_object(
            objects,
            limits,
            identifier,
            unit_index,
            part,
            "text-box",
            ObjectKind::Shape,
            bounds,
            z,
            None,
            MappingQuality::Derived,
            visual,
        )?;
        return Ok(true);
    }
    let font_size = runs
        .iter()
        .map(|run| run.font_size)
        .fold(fallback_style.font_size, f32::max);
    let explicit_box_multiline = geometry.height > 0.0
        && (text.contains('\n')
            || (geometry.width > 8.0
                && keynote_wrapped_line_count(&runs, geometry.width - 8.0) > 1));
    let uses_pingfang = runs
        .iter()
        .any(|run| run.font_family.starts_with("PingFangSC-"));
    let line_height = fallback_style.line_height_multiple.map_or_else(
        || {
            font_size
                * if geometry.height <= 0.0 {
                    1.29
                } else if explicit_box_multiline {
                    if auto_sized { 1.2 } else { 1.4 }
                } else {
                    1.0
                }
        },
        |spacing| {
            if spacing > 4.0 {
                spacing
            } else {
                font_size * spacing * 1.2
            }
        },
    );
    let mut text_layout = shape_style_identifier
        .map(|style_identifier| {
            keynote_text_layout_style(stylesheet_archives, style_identifier, STYLESHEET_COMPONENT)
        })
        .transpose()?
        .flatten();
    if text.contains('\t') {
        let layout = text_layout.get_or_insert_with(TextLayout::default);
        layout.default_tab_stop = fallback_style.default_tab_stop;
        layout.tab_stops = fallback_style.tab_stops.clone();
    }
    let anchor_adjustment =
        if geometry.height > 0.0 && geometry.resize_flags.is_some_and(|flags| flags & 2 == 0) {
            match text_layout.as_ref().map(|layout| layout.vertical_align) {
                Some(TextVerticalAlign::Center) => -geometry.height / 2.0,
                Some(TextVerticalAlign::Bottom) => -geometry.height,
                _ => 0.0,
            }
        } else {
            0.0
        };
    let vertical_adjustment = anchor_adjustment
        + if geometry.height > 0.0 {
            0.0
        } else if part.starts_with("Index/TemplateSlide") && !text.contains('\n') {
            font_size * 0.2
        } else if explicit_box_multiline {
            -(line_height - font_size * 0.72).max(0.0) + if text.contains('\n') { 1.5 } else { 0.0 }
        } else if geometry.height > font_size * 2.0 && uses_pingfang {
            -font_size * 0.015
        } else {
            fallback_style.line_height_multiple.map_or(0.0, |_| {
                ((line_height - font_size * 1.2).max(0.0)) / 2.0
                    - if uses_pingfang { font_size * 0.1 } else { 0.0 }
            })
        };
    let mut bounds = keynote_text_bounds(
        geometry,
        (offset_x, offset_y),
        (slide_width, slide_height),
        &runs,
        fallback_style.align,
        line_height,
        vertical_adjustment,
        auto_width,
    );
    if caption_parent.is_some() {
        bounds.y = offset_y + geometry.y;
    }
    let mut drawable_style = shape_style_identifier
        .map(|style_identifier| {
            keynote_drawable_style(
                stylesheet_archives,
                style_identifier,
                STYLESHEET_COMPONENT,
                bounds.width,
                bounds.height,
            )
        })
        .transpose()?
        .unwrap_or_default();
    keynote_materialize_image_fill(
        diagnostics,
        package,
        data_files,
        &mut drawable_style,
        limits,
        part,
    )?;
    let paragraph_indents = if explicit_box_multiline {
        keynote_paragraph_indents(
            &storage,
            stylesheet_archives,
            STYLESHEET_COMPONENT,
            &fallback_style,
        )?
    } else {
        Vec::new()
    };
    let paragraph_layouts = if let Some(layout) = pages_layout.filter(|_| {
        flow_target.is_some()
            || caption_parent.is_some()
            || fallback_style
                .border_positions
                .is_some_and(|positions| positions & 3 != 0)
            || paragraph_indents
                .iter()
                .any(|paragraph| paragraph.rule_above.is_some() || paragraph.rule_below.is_some())
    }) {
        let mut paragraphs = pages_paragraph_layouts(
            &storage,
            &runs,
            stylesheet_archives,
            STYLESHEET_COMPONENT,
            range_start,
            range_end,
            messages,
            limits,
            layout.view_scale,
            layout.font_metrics,
            true,
        )?;
        if let Some(first) = paragraphs.first_mut() {
            let start = keynote_utf16_byte_offset(&storage.text, range_start).unwrap_or(0);
            if start > 0 && !storage.text[..start].ends_with(['\n', '\u{2028}']) {
                first.rule_above = None;
            }
            first.space_before = first
                .rule_above
                .map_or(0.0, |rule| (-rule.offset_y).max(0.0));
        }
        let text_layout = text_layout.get_or_insert_with(TextLayout::default);
        // Saved ranges already select this frame's text. Approximate font
        // metrics must not erase its final line a second time.
        text_layout.vertical_overflow = TextVerticalOverflow::Overflow;
        let end = keynote_utf16_byte_offset(&storage.text, range_end).unwrap_or(storage.text.len());
        text_layout.continues_after =
            end < storage.text.len() && !storage.text[..end].ends_with(['\n', '\u{2028}']);
        if text_layout.continues_after
            && let Some(last) = paragraphs.last_mut()
        {
            last.rule_below = None;
        }
        Some(paragraphs)
    } else {
        explicit_box_multiline.then(|| {
            keynote_text_paragraph_layouts(
                &runs,
                fallback_style.align,
                line_height,
                fallback_style.line_height_multiple.is_some(),
                &paragraph_indents,
            )
        })
    };
    if geometry.height > 0.0 && font_size >= 48.0 {
        for run in &mut runs {
            if run.font_family == "PingFangSC-Regular" {
                let compensation = run.font_size * (1.0 - KEYNOTE_PINGFANG_GLYPH_SCALE);
                run.baseline_shift += compensation;
                run.font_size *= KEYNOTE_PINGFANG_GLYPH_SCALE;
            }
        }
    }
    if auto_sized {
        for run in &mut runs {
            if run.font_family == "Helvetica" {
                run.font_size *= KEYNOTE_HELVETICA_GLYPH_SCALE;
            }
        }
    }
    let text_visual = Visual::RichText {
        geometry: shape_geometry,
        fill: drawable_style.fill.clone(),
        stroke: drawable_style.stroke.clone(),
        stroke_width: drawable_style.stroke_width,
        align: fallback_style.align,
        line_height,
        runs,
    };
    if leading_space_indent > 0.0
        && let Some(layout) = text_layout.as_mut()
    {
        layout.first_line_indent += leading_space_indent;
    }
    if part.starts_with("Index/TemplateSlide")
        && !text.contains('\n')
        && font_size >= 48.0
        && let Some(layout) = text_layout.as_mut()
    {
        layout.vertical_align = TextVerticalAlign::Center;
    }
    if let (Some(layout), Some(paragraphs)) = (text_layout.as_mut(), paragraph_layouts) {
        layout.paragraphs = paragraphs;
    }
    if auto_width && let Some(layout) = text_layout.as_mut() {
        layout.wrap = false;
    }
    let text_visual = match text_layout {
        Some(layout) => Visual::TextLayout {
            layout,
            visual: Box::new(text_visual),
        },
        None => text_visual,
    };
    let mut effect_style = drawable_style.clone();
    if matches!(effect_style.fill, Paint::None) && matches!(effect_style.stroke, Paint::None) {
        effect_style.shadow = None;
        effect_style.reflection = None;
    }
    let visual = keynote_wrap_geometry_transform(
        geometry,
        bounds,
        keynote_wrap_style(&effect_style, None, text_visual),
    );
    push_keynote_native_object(
        objects,
        limits,
        identifier,
        unit_index,
        part,
        "text-box",
        ObjectKind::TextBox,
        bounds,
        z,
        Some(text),
        MappingQuality::Derived,
        visual,
    )?;
    Ok(true)
}

#[allow(clippy::too_many_arguments)]
fn keynote_push_chart(
    chart: KeynoteChart,
    identifier: u64,
    unit_index: u32,
    part: &str,
    limits: Limits,
    objects: &mut Vec<Object>,
    offset_x: f32,
    offset_y: f32,
    z: i32,
) -> Result<(), Diagnostic> {
    let authored = keynote_bounds(chart.geometry, offset_x, offset_y, None);
    let cartesian = chart.kind != KeynoteChartKind::Pie;
    let horizontal = chart.kind == KeynoteChartKind::Bar;
    let label_width = if horizontal {
        chart
            .categories
            .iter()
            .map(|text| {
                keynote_bounds(
                    KeynoteGeometry::default(),
                    0.0,
                    0.0,
                    Some((text, chart.category_axis.text.font_size)),
                )
                .width
            })
            .fold(0.0_f32, f32::max)
    } else {
        iwork_chart_value_label_width(&chart)
    };
    let show_left_labels = if horizontal {
        chart.category_axis.show_labels
    } else {
        chart.value_axis.show_labels
    };
    let left = if show_left_labels {
        label_width + 8.0
    } else {
        0.0
    };
    let title_height = if chart.title.is_some() {
        chart.title_style.font_size * 1.2
    } else {
        0.0
    };
    let top = if title_height > 0.0 {
        title_height + 6.0
    } else {
        chart.value_axis.text.font_size * 0.6
    };
    let bottom = if horizontal {
        chart.value_axis.text.font_size
    } else {
        chart.category_axis.text.font_size
    } * 1.2
        + 5.0;
    // Cartesian TSCH drawable geometry is the plot frame; labels/title extend
    // outside it. Pie layout has a different frame contract and remains separate.
    let bounds = if cartesian {
        Rect {
            x: authored.x - left,
            y: authored.y - top,
            width: authored.width + left,
            height: authored.height + top + bottom,
        }
    } else {
        authored
    };
    push_keynote_native_object(
        objects,
        limits,
        identifier,
        unit_index,
        part,
        "chart",
        ObjectKind::Group,
        bounds,
        z,
        None,
        MappingQuality::Derived,
        Visual::StrokeStyle {
            style: chart.background_stroke.2.clone(),
            visual: Box::new(Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: chart.background.clone(),
                stroke: chart.background_stroke.0.clone(),
                stroke_width: chart.background_stroke.1,
            }),
        },
    )?;
    let parent_numeric_id = objects.last().map(|object| object.numeric_id);
    let parent_stable_id = objects.last().map(|object| object.stable_id.clone());
    let padding_x = (bounds.width * 0.07).min(28.0);
    let padding_y = (bounds.height * 0.08).min(22.0);
    let mut plot = if cartesian {
        authored
    } else {
        Rect {
            x: bounds.x + padding_x,
            y: bounds.y + padding_y,
            width: (bounds.width - padding_x * 2.0).max(1.0),
            height: (bounds.height - padding_y * 2.0).max(1.0),
        }
    };
    let legend = chart.show_legend.then(|| {
        let width = (bounds.width * 0.28).clamp(70.0, 120.0);
        if !cartesian {
            plot.width = (plot.width - width - 12.0).max(1.0);
        }
        Rect {
            x: plot.x + plot.width + 12.0,
            y: bounds.y + (bounds.height - chart.series_names.len() as f32 * 18.0 - 8.0) / 2.0,
            width,
            height: chart.series_names.len() as f32 * 18.0 + 8.0,
        }
    });
    let mut serial = 0_u32;
    if chart.plot_fill != Paint::None {
        keynote_push_chart_child(
            identifier,
            unit_index,
            part,
            limits,
            objects,
            plot,
            z,
            parent_numeric_id,
            parent_stable_id.as_deref(),
            &mut serial,
            Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: chart.plot_fill.clone(),
                stroke: Paint::None,
                stroke_width: 0.0,
            },
        )?;
    }
    if let Some(title) = chart.title.as_ref() {
        keynote_push_chart_child(
            identifier,
            unit_index,
            part,
            limits,
            objects,
            Rect {
                x: if cartesian { plot.x } else { bounds.x },
                y: bounds.y,
                width: if cartesian { plot.width } else { bounds.width },
                height: title_height,
            },
            z,
            parent_numeric_id,
            parent_stable_id.as_deref(),
            &mut serial,
            keynote_chart_label_visual(title.clone(), TextAlign::Center, &chart.title_style),
        )?;
    }
    match chart.kind {
        KeynoteChartKind::Column => keynote_push_bar_chart(
            &chart,
            false,
            identifier,
            unit_index,
            part,
            limits,
            objects,
            plot,
            z,
            parent_numeric_id,
            parent_stable_id.as_deref(),
            &mut serial,
        ),
        KeynoteChartKind::Bar => keynote_push_bar_chart(
            &chart,
            true,
            identifier,
            unit_index,
            part,
            limits,
            objects,
            plot,
            z,
            parent_numeric_id,
            parent_stable_id.as_deref(),
            &mut serial,
        ),
        KeynoteChartKind::Line | KeynoteChartKind::Area => keynote_push_line_chart(
            &chart,
            identifier,
            unit_index,
            part,
            limits,
            objects,
            plot,
            z,
            parent_numeric_id,
            parent_stable_id.as_deref(),
            &mut serial,
        ),
        KeynoteChartKind::Pie => keynote_push_pie_chart(
            &chart,
            identifier,
            unit_index,
            part,
            limits,
            objects,
            plot,
            z,
            parent_numeric_id,
            parent_stable_id.as_deref(),
            &mut serial,
        ),
    }?;
    if let Some(legend) = legend {
        keynote_push_chart_legend(
            &chart,
            identifier,
            unit_index,
            part,
            limits,
            objects,
            legend,
            z,
            parent_numeric_id,
            parent_stable_id.as_deref(),
            &mut serial,
        )?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn keynote_push_chart_legend(
    chart: &KeynoteChart,
    identifier: u64,
    unit_index: u32,
    part: &str,
    limits: Limits,
    objects: &mut Vec<Object>,
    bounds: Rect,
    z: i32,
    parent_numeric_id: Option<u32>,
    parent_stable_id: Option<&str>,
    serial: &mut u32,
) -> Result<(), Diagnostic> {
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
        Visual::PaintedShape {
            geometry: Geometry::Rectangle,
            fill: Paint::Solid(0xffff_ffff),
            stroke: Paint::Solid(0x3333_33ff),
            stroke_width: 1.0,
        },
    )?;
    for (index, name) in chart.series_names.iter().enumerate() {
        let y = bounds.y + 5.0 + index as f32 * 18.0;
        keynote_push_chart_child(
            identifier,
            unit_index,
            part,
            limits,
            objects,
            Rect {
                x: bounds.x + 6.0,
                y: y + 2.0,
                width: 12.0,
                height: 12.0,
            },
            z,
            parent_numeric_id,
            parent_stable_id,
            serial,
            Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: Paint::Solid(keynote_chart_color(chart, index)),
                stroke: Paint::Solid(0x3333_33ff),
                stroke_width: 1.0,
            },
        )?;
        keynote_push_chart_child(
            identifier,
            unit_index,
            part,
            limits,
            objects,
            Rect {
                x: bounds.x + 23.0,
                y,
                width: (bounds.width - 27.0).max(1.0),
                height: 16.0,
            },
            z,
            parent_numeric_id,
            parent_stable_id,
            serial,
            keynote_chart_label_visual(name.clone(), TextAlign::Start, &chart.category_axis.text),
        )?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn keynote_push_bar_chart(
    chart: &KeynoteChart,
    horizontal: bool,
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
    let categories = chart.series.iter().map(Vec::len).max().unwrap_or(0);
    if categories == 0 {
        return Ok(());
    }
    let (minimum, maximum) = keynote_chart_range(&chart.series, chart.stacked);
    let span = (maximum - minimum).max(f32::EPSILON);
    let value_label_width = iwork_chart_value_label_width(chart);
    let baseline = if horizontal {
        plot.x + (0.0 - minimum) / span * plot.width
    } else {
        plot.y + (maximum - 0.0) / span * plot.height
    };
    keynote_push_chart_axes(
        chart,
        identifier,
        unit_index,
        part,
        limits,
        objects,
        plot,
        horizontal,
        baseline,
        z,
        parent_numeric_id,
        parent_stable_id,
        serial,
    )?;
    let category_span = if horizontal {
        plot.height / categories as f32
    } else {
        plot.width / categories as f32
    };
    for tick in (0..=4).filter(|_| chart.value_axis.show_labels) {
        let ratio = tick as f32 / 4.0;
        let value = minimum + span * ratio;
        let bounds = if horizontal {
            Rect {
                x: plot.x + plot.width * ratio - 20.0,
                y: plot.y + plot.height + 2.0,
                width: 40.0,
                height: 12.0,
            }
        } else {
            Rect {
                x: plot.x - value_label_width - 8.0,
                y: plot.y + plot.height * (1.0 - ratio) - chart.value_axis.text.font_size * 0.6,
                width: value_label_width + 2.0,
                height: chart.value_axis.text.font_size * 1.2,
            }
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
                keynote_chart_number(value, false),
                if horizontal {
                    TextAlign::Center
                } else {
                    TextAlign::End
                },
                &chart.value_axis.text,
            ),
        )?;
    }
    for category in (0..categories).filter(|_| chart.category_axis.show_labels) {
        let Some(label) = chart
            .categories
            .get(category)
            .filter(|label| !label.is_empty())
        else {
            continue;
        };
        let bounds = if horizontal {
            Rect {
                x: plot.x - 100.0,
                y: plot.y + category_span * (category as f32 + 0.5) - 6.0,
                width: 94.0,
                height: 12.0,
            }
        } else {
            Rect {
                x: plot.x + category_span * category as f32,
                y: plot.y + plot.height + 3.0,
                width: category_span,
                height: chart.category_axis.text.font_size * 1.2 + 2.0,
            }
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
    let series_count = if chart.stacked {
        1
    } else {
        chart.series.len().max(1)
    } as f32;
    let (grouped_span, group_margin, series_step) = if let Some((bar_gap, set_gap)) = chart.bar_gaps
    {
        let width = category_span / (series_count + (series_count - 1.0) * bar_gap + set_gap);
        (width, width * set_gap / 2.0, width * (1.0 + bar_gap))
    } else {
        (
            category_span * 0.72 / series_count,
            category_span * 0.14,
            category_span * 0.72 / series_count,
        )
    };
    let mut positive = vec![0.0_f32; categories];
    let mut negative = vec![0.0_f32; categories];
    for (series_index, series) in chart.series.iter().enumerate() {
        for category in 0..categories {
            let value = series.get(category).copied().unwrap_or(0.0);
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
            let (start_position, end_position) = if horizontal {
                (
                    plot.x + (start - minimum) / span * plot.width,
                    plot.x + (end - minimum) / span * plot.width,
                )
            } else {
                (
                    plot.y + (maximum - start) / span * plot.height,
                    plot.y + (maximum - end) / span * plot.height,
                )
            };
            let group_offset = if chart.stacked {
                group_margin
            } else {
                group_margin + series_step * series_index as f32
            };
            let bar_bounds = if horizontal {
                Rect {
                    x: start_position.min(end_position),
                    y: plot.y + category as f32 * category_span + group_offset,
                    width: (end_position - start_position).abs().max(0.5),
                    height: grouped_span.max(0.5),
                }
            } else {
                Rect {
                    x: plot.x + category as f32 * category_span + group_offset,
                    y: start_position.min(end_position),
                    width: grouped_span.max(0.5),
                    height: (end_position - start_position).abs().max(0.5),
                }
            };
            keynote_push_chart_child(
                identifier,
                unit_index,
                part,
                limits,
                objects,
                bar_bounds,
                z,
                parent_numeric_id,
                parent_stable_id,
                serial,
                Visual::PaintedShape {
                    geometry: Geometry::Rectangle,
                    fill: Paint::Solid(keynote_chart_color(chart, series_index)),
                    stroke: Paint::None,
                    stroke_width: 0.0,
                },
            )?;
            if chart
                .show_data_labels
                .get(series_index)
                .copied()
                .unwrap_or(false)
            {
                let bounds = if horizontal {
                    Rect {
                        x: bar_bounds.x + bar_bounds.width + 3.0,
                        y: bar_bounds.y + bar_bounds.height / 2.0 - 6.0,
                        width: 40.0,
                        height: 12.0,
                    }
                } else {
                    Rect {
                        x: bar_bounds.x - bar_bounds.width * 0.25,
                        y: bar_bounds.y - 14.0,
                        width: bar_bounds.width * 1.5,
                        height: 12.0,
                    }
                };
                let percent = chart
                    .series
                    .iter()
                    .flatten()
                    .all(|value| value.abs() <= 1.0 + f32::EPSILON);
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
                        keynote_chart_number(value, percent),
                        TextAlign::Center,
                        &chart.value_axis.text,
                    ),
                )?;
            }
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn keynote_push_line_chart(
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
    let categories = chart.series.iter().map(Vec::len).max().unwrap_or(0);
    if categories == 0 {
        return Ok(());
    }
    let (minimum, maximum) = keynote_chart_range(&chart.series, chart.stacked);
    let span = (maximum - minimum).max(f32::EPSILON);
    let baseline = plot.y + (maximum - 0.0) / span * plot.height;
    keynote_push_chart_axes(
        chart,
        identifier,
        unit_index,
        part,
        limits,
        objects,
        plot,
        false,
        baseline,
        z,
        parent_numeric_id,
        parent_stable_id,
        serial,
    )?;
    let mut cumulative = vec![0.0_f32; categories];
    for (series_index, series) in chart.series.iter().enumerate() {
        let mut commands = Vec::new();
        let mut points = Vec::new();
        for (category, cumulative_value) in cumulative.iter_mut().enumerate() {
            let value = series.get(category).copied().unwrap_or(0.0);
            let plotted = if chart.stacked {
                *cumulative_value += value;
                *cumulative_value
            } else {
                value
            };
            let x = if categories == 1 {
                plot.width / 2.0
            } else {
                category as f32 * plot.width / (categories - 1) as f32
            };
            let y = (maximum - plotted) / span * plot.height;
            points.push((x, y));
        }
        if chart.kind == KeynoteChartKind::Area {
            commands.push(PathCommand::MoveTo {
                x: points.first().map_or(0.0, |point| point.0),
                y: (baseline - plot.y).clamp(0.0, plot.height),
            });
        }
        for (index, (x, y)) in points.iter().copied().enumerate() {
            commands.push(if index == 0 && chart.kind == KeynoteChartKind::Line {
                PathCommand::MoveTo { x, y }
            } else {
                PathCommand::LineTo { x, y }
            });
        }
        if chart.kind == KeynoteChartKind::Area {
            commands.push(PathCommand::LineTo {
                x: points.last().map_or(plot.width, |point| point.0),
                y: (baseline - plot.y).clamp(0.0, plot.height),
            });
            commands.push(PathCommand::ClosePath);
        }
        let color = keynote_chart_color(chart, series_index);
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
                geometry: Geometry::Path {
                    fill_rule: FillRule::NonZero,
                    commands,
                },
                fill: if chart.kind == KeynoteChartKind::Area {
                    Paint::Solid(keynote_color_opacity(color, 0.55))
                } else {
                    Paint::None
                },
                stroke: Paint::Solid(color),
                stroke_width: 2.0,
            },
        )?;
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn keynote_push_pie_chart(
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
    let Some(series) = chart.series.first() else {
        return Ok(());
    };
    let values: Vec<f32> = series.iter().map(|value| value.max(0.0)).collect();
    let total: f32 = values.iter().sum();
    if total <= f32::EPSILON {
        return Ok(());
    }
    let size = plot.width.min(plot.height) * 0.58;
    let pie = Rect {
        x: plot.x + (plot.width - size) / 2.0,
        y: plot.y + (plot.height - size) / 2.0,
        width: size,
        height: size,
    };
    let center = size / 2.0;
    let radius = center;
    let mut start_angle = -std::f32::consts::FRAC_PI_2;
    for (index, value) in values.into_iter().enumerate() {
        if value <= f32::EPSILON {
            continue;
        }
        let end_angle = start_angle + std::f32::consts::TAU * value / total;
        let segments = (((end_angle - start_angle).abs() / std::f32::consts::TAU) * 48.0)
            .ceil()
            .max(2.0) as usize;
        let mut commands = vec![PathCommand::MoveTo {
            x: center,
            y: center,
        }];
        for segment in 0..=segments {
            let ratio = segment as f32 / segments as f32;
            let angle = start_angle + (end_angle - start_angle) * ratio;
            commands.push(PathCommand::LineTo {
                x: center + radius * angle.cos(),
                y: center + radius * angle.sin(),
            });
        }
        commands.push(PathCommand::ClosePath);
        keynote_push_chart_child(
            identifier,
            unit_index,
            part,
            limits,
            objects,
            pie,
            z,
            parent_numeric_id,
            parent_stable_id,
            serial,
            Visual::PaintedShape {
                geometry: Geometry::Path {
                    fill_rule: FillRule::NonZero,
                    commands,
                },
                fill: Paint::Solid(keynote_chart_color(chart, index)),
                stroke: Paint::Solid(0xffff_ffff),
                stroke_width: 1.0,
            },
        )?;
        if let Some(category) = chart
            .categories
            .get(index)
            .filter(|value| !value.is_empty())
        {
            let label = format!("{category}\n{:.0}%", value / total * 100.0);
            let style = KeynoteTextStyle {
                font_family: "Helvetica".to_owned(),
                font_size: 9.0,
                align: TextAlign::Center,
                ..KeynoteTextStyle::default()
            };
            let midpoint = (start_angle + end_angle) / 2.0;
            let label_width = plot.width * 0.38;
            let label_height = 26.0;
            let label_radius = size * 0.78;
            keynote_push_chart_child(
                identifier,
                unit_index,
                part,
                limits,
                objects,
                Rect {
                    x: pie.x + center + label_radius * midpoint.cos() - label_width / 2.0,
                    y: pie.y + center + label_radius * midpoint.sin() - label_height / 2.0,
                    width: label_width,
                    height: label_height,
                },
                z,
                parent_numeric_id,
                parent_stable_id,
                serial,
                Visual::RichText {
                    geometry: Geometry::Rectangle,
                    fill: Paint::None,
                    stroke: Paint::None,
                    stroke_width: 0.0,
                    align: TextAlign::Center,
                    line_height: 10.8,
                    runs: vec![keynote_text_run(label, &style)],
                },
            )?;
        }
        start_angle = end_angle;
    }
    Ok(())
}

fn keynote_chart_range(series: &[Vec<f32>], stacked: bool) -> (f32, f32) {
    if stacked {
        let categories = series.iter().map(Vec::len).max().unwrap_or(0);
        let mut minimum = 0.0_f32;
        let mut maximum = 0.0_f32;
        for category in 0..categories {
            let mut positive = 0.0_f32;
            let mut negative = 0.0_f32;
            for values in series {
                let value = values.get(category).copied().unwrap_or(0.0);
                if value >= 0.0 {
                    positive += value;
                } else {
                    negative += value;
                }
            }
            minimum = minimum.min(negative);
            maximum = maximum.max(positive);
        }
        keynote_chart_nice_range(minimum, maximum)
    } else {
        let mut minimum = 0.0_f32;
        let mut maximum = 0.0_f32;
        for value in series.iter().flatten().copied() {
            minimum = minimum.min(value);
            maximum = maximum.max(value);
        }
        keynote_chart_nice_range(minimum, maximum)
    }
}

fn keynote_chart_nice_range(minimum: f32, maximum: f32) -> (f32, f32) {
    if maximum - minimum <= f32::EPSILON {
        return (minimum, minimum + 1.0);
    }
    if minimum >= 0.0 {
        let magnitude = 10.0_f32.powf(maximum.log10().floor());
        return (0.0, (maximum / magnitude).ceil() * magnitude);
    }
    (minimum, maximum)
}

fn keynote_chart_number(value: f32, percent: bool) -> String {
    if percent {
        return format!("{:.0}%", value * 100.0);
    }
    let mut text = format!("{value:.2}");
    while text.ends_with('0') {
        text.pop();
    }
    if text.ends_with('.') {
        text.pop();
    }
    text
}

fn keynote_chart_label_visual(text: String, align: TextAlign, style: &KeynoteTextStyle) -> Visual {
    Visual::TextLayout {
        layout: TextLayout {
            inset_left: 0.0,
            inset_right: 0.0,
            inset_top: 0.0,
            inset_bottom: 0.0,
            wrap: false,
            ..TextLayout::default()
        },
        visual: Box::new(Visual::RichText {
            geometry: Geometry::Rectangle,
            fill: Paint::None,
            stroke: Paint::None,
            stroke_width: 0.0,
            align,
            line_height: style.font_size * 1.2,
            runs: vec![keynote_text_run(text, style)],
        }),
    }
}

fn iwork_chart_value_label_width(chart: &KeynoteChart) -> f32 {
    let (minimum, maximum) = keynote_chart_range(&chart.series, chart.stacked);
    // ponytail: reuse the existing bounded text-width estimate; browser font
    // metrics can replace it when chart label layout receives measured advances.
    [minimum, maximum]
        .iter()
        .map(|value| {
            keynote_bounds(
                KeynoteGeometry::default(),
                0.0,
                0.0,
                Some((
                    &keynote_chart_number(*value, false),
                    chart.value_axis.text.font_size,
                )),
            )
            .width
        })
        .fold(0.0_f32, f32::max)
}

#[allow(clippy::too_many_arguments)]
fn keynote_push_chart_axes(
    chart: &KeynoteChart,
    identifier: u64,
    unit_index: u32,
    part: &str,
    limits: Limits,
    objects: &mut Vec<Object>,
    plot: Rect,
    horizontal: bool,
    baseline: f32,
    z: i32,
    parent_numeric_id: Option<u32>,
    parent_stable_id: Option<&str>,
    serial: &mut u32,
) -> Result<(), Diagnostic> {
    if let Some((stroke, stroke_width, style)) = &chart.value_axis.grid {
        let mut commands = Vec::new();
        for tick in 0..=4 {
            let ratio = tick as f32 / 4.0;
            let (start, end) = if horizontal {
                ((plot.width * ratio, 0.0), (plot.width * ratio, plot.height))
            } else {
                (
                    (0.0, plot.height * ratio),
                    (plot.width, plot.height * ratio),
                )
            };
            commands.push(PathCommand::MoveTo {
                x: start.0,
                y: start.1,
            });
            commands.push(PathCommand::LineTo { x: end.0, y: end.1 });
        }
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
                        commands,
                    },
                    fill: Paint::None,
                    stroke: stroke.clone(),
                    stroke_width: *stroke_width,
                }),
            },
        )?;
    }
    let mut commands = if horizontal {
        vec![
            PathCommand::MoveTo {
                x: (baseline - plot.x).clamp(0.0, plot.width),
                y: 0.0,
            },
            PathCommand::LineTo {
                x: (baseline - plot.x).clamp(0.0, plot.width),
                y: plot.height,
            },
            PathCommand::MoveTo {
                x: 0.0,
                y: plot.height,
            },
            PathCommand::LineTo {
                x: plot.width,
                y: plot.height,
            },
        ]
    } else {
        vec![
            PathCommand::MoveTo {
                x: 0.0,
                y: (baseline - plot.y).clamp(0.0, plot.height),
            },
            PathCommand::LineTo {
                x: plot.width,
                y: (baseline - plot.y).clamp(0.0, plot.height),
            },
            PathCommand::MoveTo { x: 0.0, y: 0.0 },
            PathCommand::LineTo {
                x: 0.0,
                y: plot.height,
            },
        ]
    };
    // The first segment is the category baseline, the second the value axis.
    if !chart.value_axis.show_axis {
        commands.truncate(2);
    }
    if !chart.category_axis.show_axis
        || (chart.value_axis.grid.is_some()
            && (baseline
                - if horizontal {
                    plot.x
                } else {
                    plot.y + plot.height
                })
            .abs()
                < 0.001)
    {
        commands.drain(..2);
    }
    if commands.is_empty() {
        return Ok(());
    }
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
            geometry: Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands,
            },
            fill: Paint::None,
            stroke: Paint::Solid(0x7777_77ff),
            stroke_width: 1.0,
        },
    )
}

#[allow(clippy::too_many_arguments)]
fn keynote_push_chart_child(
    identifier: u64,
    unit_index: u32,
    part: &str,
    limits: Limits,
    objects: &mut Vec<Object>,
    bounds: Rect,
    z: i32,
    parent_numeric_id: Option<u32>,
    parent_stable_id: Option<&str>,
    serial: &mut u32,
    visual: Visual,
) -> Result<(), Diagnostic> {
    if objects.len() >= limits.max_document_objects {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "Keynote native object count exceeds the configured limit",
        )
        .in_part(part));
    }
    *serial = serial
        .checked_add(1)
        .ok_or_else(|| iwa_error(part, "Keynote chart object identifier overflows"))?;
    let numeric_id = u32::try_from(objects.len())
        .ok()
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| iwa_error(part, "Keynote native object identifier overflows"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id,
        stable_id: format!(
            "iwork-keynote-chart-{unit_index}-{part}-{identifier}-{}",
            *serial
        ),
        parent_stable_id: parent_stable_id.map(str::to_owned),
        kind: ObjectKind::Shape,
        unit_index,
        bounds,
        z: z.saturating_add(*serial as i32),
        text: None,
        source: SourceRef {
            part: part.to_owned(),
            mapping: MappingQuality::Derived,
            locator: SourceLocator::Iwork {
                kind: "chart-series",
                component: format!("archive-{identifier}"),
            },
        },
        visual,
    });
    Ok(())
}

fn keynote_chart_color(chart: &KeynoteChart, index: usize) -> u32 {
    const PALETTE: [u32; 8] = [
        0x4f81_bdff,
        0xc050_4dff,
        0x9bbb_59ff,
        0x8064_a2ff,
        0x4bacc6ff,
        0xf796_46ff,
        0x3355_88ff,
        0x9436_34ff,
    ];
    chart
        .colors
        .get(index)
        .copied()
        .filter(|color| *color != 0)
        .unwrap_or(PALETTE[index % PALETTE.len()])
}

#[allow(clippy::too_many_arguments)]
fn push_keynote_native_object(
    objects: &mut Vec<Object>,
    limits: Limits,
    identifier: u64,
    unit_index: u32,
    part: &str,
    source_kind: &'static str,
    kind: ObjectKind,
    bounds: Rect,
    z: i32,
    text: Option<String>,
    mapping: MappingQuality,
    visual: Visual,
) -> Result<(), Diagnostic> {
    if objects.len() >= limits.max_document_objects {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "Keynote native object count exceeds the configured limit",
        )
        .in_part(part));
    }
    let numeric_id = u32::try_from(objects.len())
        .ok()
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| iwa_error(part, "Keynote native object identifier overflows"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("iwork-keynote-{unit_index}-{part}-{identifier}"),
        parent_stable_id: None,
        kind,
        unit_index,
        bounds,
        z,
        text,
        source: SourceRef {
            part: part.to_owned(),
            mapping,
            locator: SourceLocator::Iwork {
                kind: source_kind,
                component: format!("archive-{identifier}"),
            },
        },
        visual,
    });
    Ok(())
}

fn keynote_bounds(
    geometry: KeynoteGeometry,
    offset_x: f32,
    offset_y: f32,
    text: Option<(&str, f32)>,
) -> Rect {
    let (estimated_width, estimated_height) = text.map_or((1.0, 1.0), |(text, font_size)| {
        let longest = text
            .lines()
            .map(str::chars)
            .map(Iterator::count)
            .max()
            .unwrap_or(1) as f32;
        let lines = text.lines().count().max(1) as f32;
        (longest * font_size * 0.56, lines * font_size * 1.25)
    });
    Rect {
        x: offset_x + geometry.x,
        y: offset_y + geometry.y,
        width: if geometry.width > 0.0 {
            geometry.width
        } else {
            estimated_width
        },
        height: if geometry.height > 0.0 {
            geometry.height
        } else {
            estimated_height
        },
    }
}

fn keynote_wrap_geometry_transform(
    geometry: KeynoteGeometry,
    bounds: Rect,
    visual: Visual,
) -> Visual {
    let radians = (-geometry.rotation_degrees).rem_euclid(360.0).to_radians();
    if radians.abs() <= f32::EPSILON && !geometry.flip_horizontal && !geometry.flip_vertical {
        return visual;
    }
    let cosine = radians.cos();
    let sine = radians.sin();
    let scale_x = if geometry.flip_horizontal { -1.0 } else { 1.0 };
    let scale_y = if geometry.flip_vertical { -1.0 } else { 1.0 };
    let a = cosine * scale_x;
    let b = sine * scale_x;
    let c = -sine * scale_y;
    let d = cosine * scale_y;
    let center_x = bounds.x + bounds.width / 2.0;
    let center_y = bounds.y + bounds.height / 2.0;
    Visual::Layer {
        transform: AffineTransform {
            a,
            b,
            c,
            d,
            e: center_x - a * center_x - c * center_y,
            f: center_y - b * center_x - d * center_y,
        },
        opacity: 1.0,
        blend_mode: crate::model::BlendMode::Normal,
        visual: Box::new(visual),
    }
}

fn keynote_text_bounds(
    geometry: KeynoteGeometry,
    offset: (f32, f32),
    slide_size: (f32, f32),
    runs: &[TextRun],
    align: TextAlign,
    line_height: f32,
    vertical_adjustment: f32,
    auto_width: bool,
) -> Rect {
    const HORIZONTAL_INSETS: f32 = 8.0;
    const VERTICAL_INSETS: f32 = 8.0;

    let (offset_x, offset_y) = offset;
    let (slide_width, slide_height) = slide_size;
    let anchor_x = offset_x + geometry.x;
    let anchor_y = offset_y + geometry.y;
    let available_width = match align {
        TextAlign::Center => 2.0 * anchor_x.min((slide_width - anchor_x).max(0.0)),
        TextAlign::End => anchor_x,
        TextAlign::Start
        | TextAlign::Justify
        | TextAlign::Distribute
        | TextAlign::MediumKashida
        | TextAlign::HighKashida
        | TextAlign::LowKashida
        | TextAlign::ThaiDistribute => slide_width - anchor_x,
    }
    .max(1.0);
    let natural_width = keynote_text_natural_width(runs) * 1.03 + HORIZONTAL_INSETS;
    let width = if geometry.width > 0.0 && !auto_width {
        geometry.width
    } else if natural_width > available_width {
        (available_width - 24.0).max(1.0)
    } else {
        natural_width.max(1.0)
    };
    let content_width = (width - HORIZONTAL_INSETS).max(1.0);
    let height = if geometry.height > 0.0 {
        geometry.height
    } else {
        keynote_wrapped_line_count(runs, content_width) as f32 * line_height + VERTICAL_INSETS
    };
    let x = if geometry.width > 0.0 && !auto_width {
        anchor_x
    } else {
        match align {
            TextAlign::Center => anchor_x - width / 2.0,
            TextAlign::End => anchor_x - width,
            TextAlign::Start
            | TextAlign::Justify
            | TextAlign::Distribute
            | TextAlign::MediumKashida
            | TextAlign::HighKashida
            | TextAlign::LowKashida
            | TextAlign::ThaiDistribute => anchor_x,
        }
    };
    let y = if geometry.height > 0.0 {
        anchor_y + vertical_adjustment
    } else {
        anchor_y - height / 2.0 + vertical_adjustment
    };
    Rect {
        x: x.clamp(0.0, slide_width),
        y: y.clamp(0.0, slide_height),
        width: width.min(slide_width),
        height: height.min(slide_height),
    }
}

fn keynote_text_natural_width(runs: &[TextRun]) -> f32 {
    let mut current = 0.0_f32;
    let mut maximum = 0.0_f32;
    for run in runs {
        for character in run.text.chars() {
            if character == '\n' {
                maximum = maximum.max(current);
                current = 0.0;
            } else {
                current += keynote_character_advance(character, run.font_size, run.letter_spacing);
            }
        }
    }
    maximum.max(current)
}

fn keynote_text_paragraph_layouts(
    runs: &[TextRun],
    align: TextAlign,
    fallback_line_height: f32,
    authored_line_height: bool,
    paragraph_indents: &[TextParagraphLayout],
) -> Vec<TextParagraphLayout> {
    let mut layouts = Vec::new();
    let mut paragraph_line_height = 0.0_f32;
    for run in runs {
        let natural_line_height = run.font_size
            * if run.font_family == "PingFangSC-Regular" {
                1.4
            } else {
                1.2
            };
        for segment in run.text.split_inclusive('\n') {
            paragraph_line_height = paragraph_line_height.max(natural_line_height);
            if segment.ends_with('\n') {
                let line_height = if paragraph_line_height > 0.0 {
                    if authored_line_height {
                        fallback_line_height
                    } else {
                        paragraph_line_height
                    }
                } else {
                    fallback_line_height
                };
                layouts.push(keynote_text_paragraph_layout(
                    align,
                    line_height,
                    if layouts.is_empty() || line_height >= fallback_line_height {
                        0.0
                    } else {
                        3.0
                    },
                    paragraph_indents.get(layouts.len()),
                ));
                paragraph_line_height = 0.0;
            }
        }
    }
    if paragraph_line_height > 0.0 || layouts.is_empty() {
        let line_height = if paragraph_line_height > 0.0 {
            if authored_line_height {
                fallback_line_height
            } else {
                paragraph_line_height
            }
        } else {
            fallback_line_height
        };
        layouts.push(keynote_text_paragraph_layout(
            align,
            line_height,
            if layouts.is_empty() || line_height >= fallback_line_height {
                0.0
            } else {
                3.0
            },
            paragraph_indents.get(layouts.len()),
        ));
    }
    for (index, (layout, authored)) in layouts.iter_mut().zip(paragraph_indents).enumerate() {
        // Keynote suppresses paragraph-before spacing at the top of a text frame.
        layout.space_before = if index == 0 {
            0.0
        } else {
            authored.space_before
        };
        layout.space_after = authored.space_after;
    }
    layouts
}

fn keynote_paragraph_indents(
    storage: &KeynoteTextStorage,
    archives: &[IwaArchive],
    part: &str,
    fallback_style: &KeynoteTextStyle,
) -> Result<Vec<TextParagraphLayout>, Diagnostic> {
    let mut offsets = Vec::new();
    let mut position = 0_usize;
    for paragraph in storage.text.split_inclusive(['\n', '\u{2028}']) {
        let mut style = iwork_inherited_style_at(&storage.paragraph_styles, position)
            .map(|identifier| keynote_text_style(archives, identifier, part))
            .transpose()?
            .unwrap_or_else(|| fallback_style.clone());
        style.list_level = iwork_paragraph_list_level_at(&storage.paragraph_data, position);
        let offset = iwork_inherited_style_at(&storage.list_styles, position)
            .map(|identifier| {
                keynote_list_marker(
                    archives,
                    identifier,
                    style.list_level,
                    part,
                    &mut Vec::new(),
                )
            })
            .transpose()?
            .flatten()
            .map(|marker| style.margin_left + marker.indent + marker.text_indent * style.font_size);
        let mut layout = keynote_text_paragraph_layout(style.align, 0.0, 0.0, None);
        (layout.margin_left, layout.first_line_indent) = style.paragraph_indents(offset);
        layout.margin_right = style.margin_right;
        layout.default_tab_stop = style.default_tab_stop;
        layout.space_before = style.space_before;
        layout.space_after = style.space_after;
        layout.rule_above = style.rule_above();
        layout.rule_below = style.horizontal_border(2);
        offsets.push(layout);
        position = position.saturating_add(paragraph.encode_utf16().count());
    }
    Ok(offsets)
}

fn keynote_text_paragraph_layout(
    align: TextAlign,
    line_height: f32,
    space_before: f32,
    body_layout: Option<&TextParagraphLayout>,
) -> TextParagraphLayout {
    let body_offset = body_layout.map_or(0.0, |layout| layout.margin_left);
    TextParagraphLayout {
        align,
        margin_left: body_offset,
        margin_right: 0.0,
        first_line_indent: -body_offset,
        default_tab_stop: body_layout.map_or(36.0, |layout| layout.default_tab_stop),
        line_height,
        space_before,
        space_after: 0.0,
        latin_line_break: true,
        hanging_punctuation: false,
        rule_above: None,
        rule_below: None,
        drop_cap: None,
    }
}

fn keynote_wrapped_line_count(runs: &[TextRun], maximum_width: f32) -> usize {
    let mut lines = 1_usize;
    let mut current = 0.0_f32;
    for run in runs {
        for character in run.text.chars() {
            if character == '\n' {
                lines = lines.saturating_add(1);
                current = 0.0;
                continue;
            }
            let advance = keynote_character_advance(character, run.font_size, run.letter_spacing);
            if current > 0.0 && current + advance > maximum_width {
                lines = lines.saturating_add(1);
                current = 0.0;
            }
            current += advance;
        }
    }
    lines
}

fn keynote_text_layout_height(
    runs: &[TextRun],
    paragraphs: &[TextParagraphLayout],
    maximum_width: f32,
) -> f32 {
    if runs.is_empty() {
        return 0.0;
    }
    let paragraph_metrics = |index: usize| {
        paragraphs
            .get(index)
            .or_else(|| paragraphs.last())
            .map_or((12.6, 0.0, 0.0), |paragraph| {
                (
                    paragraph.line_height,
                    paragraph.space_before,
                    paragraph.space_after,
                )
            })
    };
    let mut height = 0.0_f32;
    let mut width = 0.0_f32;
    let mut paragraph_index = 0_usize;
    let mut paragraph_started = false;
    for run in runs {
        for character in run.text.chars() {
            let (line_height, space_before, space_after) = paragraph_metrics(paragraph_index);
            if !paragraph_started {
                height += space_before;
                paragraph_started = true;
            }
            if character == '\n' {
                height += line_height + space_after;
                width = 0.0;
                paragraph_index = paragraph_index.saturating_add(1);
                paragraph_started = false;
                continue;
            }
            let advance = keynote_character_advance(character, run.font_size, run.letter_spacing);
            if width > 0.0 && width + advance > maximum_width {
                height += line_height;
                width = 0.0;
            }
            width += advance;
        }
    }
    if paragraph_started {
        let (line_height, _, space_after) = paragraph_metrics(paragraph_index);
        height += line_height + space_after;
    }
    height
}

fn keynote_character_advance(character: char, font_size: f32, letter_spacing: f32) -> f32 {
    let ratio = if character == '\t' {
        2.0
    } else if character.is_ascii_whitespace() {
        0.33
    } else if character.is_ascii_punctuation() {
        0.4
    } else if character.is_ascii() {
        0.56
    } else {
        1.0
    };
    (font_size * ratio + letter_spacing).max(0.0)
}

fn keynote_group(bytes: &[u8], part: &str) -> Result<(KeynoteGeometry, Vec<u64>), Diagnostic> {
    let mut cursor = 0_usize;
    let mut geometry = KeynoteGeometry::default();
    let mut children = Vec::new();
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(part, "Keynote group contains too many fields"));
        }
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote group key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let drawable =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote group drawable")?;
                geometry = keynote_drawable_geometry(drawable, part)?;
            }
            (2, 2) => {
                let reference =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote group child")?;
                children.push(parse_iwa_reference(reference)?);
            }
            (1 | 2, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote group field has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok((geometry, children))
}

fn keynote_image(bytes: &[u8], part: &str) -> Result<KeynoteImage, Diagnostic> {
    let mut cursor = 0_usize;
    let mut geometry = KeynoteGeometry::default();
    let mut data = None;
    let mut original = None;
    let mut adjusted = None;
    let mut enhanced = None;
    let mut style = None;
    let mut mask = None;
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(part, "Keynote image contains too many fields"));
        }
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote image key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let drawable =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote image drawable")?;
                geometry = keynote_drawable_geometry(drawable, part)?;
            }
            (field @ (3 | 5 | 11 | 12 | 13 | 15 | 16 | 17), 2) => {
                let reference =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote image reference")?;
                let reference = Some(parse_iwa_reference(reference)?);
                match field {
                    3 => style = reference,
                    5 => mask = reference,
                    11 => data = reference,
                    12 => {}
                    13 => original = reference,
                    15 => adjusted = reference,
                    16 => {}
                    17 => enhanced = reference,
                    _ => unreachable!(),
                }
            }
            (1 | 3 | 5 | 11 | 12 | 13 | 15 | 16 | 17, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote image field has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(KeynoteImage {
        geometry,
        data_identifiers: [adjusted, enhanced, data, original]
            .into_iter()
            .flatten()
            .collect(),
        style_identifier: style,
        mask_identifier: mask,
    })
}

fn keynote_mask(
    bytes: &[u8],
    image_identifier: u64,
    part: &str,
) -> Result<KeynoteMask, Diagnostic> {
    let mut cursor = 0_usize;
    let mut geometry = KeynoteGeometry::default();
    let mut path_source = None;
    let mut relative_to_image = false;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote mask key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let drawable =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote mask drawable")?;
                geometry = keynote_drawable_geometry(drawable, part)?;
                relative_to_image =
                    keynote_reference_field(drawable, 2, part)? == Some(image_identifier);
            }
            (2, 2) => {
                path_source = Some(read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote mask path",
                )?);
            }
            (1 | 2, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote mask field has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    let clip = path_source
        .map(|path| keynote_path_source(path, geometry.width, geometry.height, part))
        .transpose()?
        .flatten()
        .unwrap_or(Geometry::Rectangle);
    Ok(KeynoteMask {
        geometry,
        clip,
        relative_to_image,
    })
}

fn keynote_image_crop(image: KeynoteGeometry, mask: KeynoteGeometry) -> ImageCrop {
    if image.width <= f32::EPSILON || image.height <= f32::EPSILON {
        return ImageCrop::default();
    }
    let crop = ImageCrop {
        left: (mask.x - image.x) / image.width,
        top: (mask.y - image.y) / image.height,
        right: (image.x + image.width - mask.x - mask.width) / image.width,
        bottom: (image.y + image.height - mask.y - mask.height) / image.height,
    };
    if crop.is_valid() {
        crop
    } else {
        ImageCrop::default()
    }
}

fn keynote_relative_image_crop(image: KeynoteGeometry, mask: KeynoteGeometry) -> ImageCrop {
    if image.width <= f32::EPSILON || image.height <= f32::EPSILON {
        return ImageCrop::default();
    }
    ImageCrop {
        left: (mask.x / image.width).clamp(0.0, 1.0),
        right: ((image.width - mask.x - mask.width) / image.width).clamp(0.0, 1.0),
        top: (mask.y / image.height).clamp(0.0, 1.0),
        bottom: ((image.height - mask.y - mask.height) / image.height).clamp(0.0, 1.0),
    }
}

fn keynote_shape(
    bytes: &[u8],
    part: &str,
) -> Result<(KeynoteGeometry, Option<u64>, Geometry), Diagnostic> {
    let mut cursor = 0_usize;
    let mut geometry = KeynoteGeometry::default();
    let mut style = None;
    let mut path_source = None;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote shape key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let drawable =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote shape drawable")?;
                geometry = keynote_drawable_geometry(drawable, part)?;
            }
            (2, 2) => {
                let reference =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote shape style")?;
                style = Some(parse_iwa_reference(reference)?);
            }
            (3, 2) => {
                path_source = Some(read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote shape path",
                )?);
            }
            (1..=3, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote shape field has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    let shape = path_source
        .map(|path| keynote_path_source(path, geometry.width, geometry.height, part))
        .transpose()?
        .flatten()
        .unwrap_or(Geometry::Rectangle);
    Ok((geometry, style, shape))
}

fn keynote_chart(
    bytes: &[u8],
    archives: &[IwaArchive],
    stylesheet_archives: &[IwaArchive],
    part: &str,
) -> Result<Option<KeynoteChart>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut geometry = KeynoteGeometry::default();
    let mut archive = None;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote chart key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let drawable =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote chart drawable")?;
                geometry = keynote_drawable_geometry(drawable, part)?;
            }
            (10_000, 2) => {
                archive = Some(read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote chart archive",
                )?);
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    let Some(archive) = archive else {
        return Ok(None);
    };
    let mut cursor = 0_usize;
    let mut chart_type = 0_u64;
    let mut direction = 0_u64;
    let mut title_reference = None;
    let mut style_references = Vec::new();
    let mut row_names = Vec::new();
    let mut column_names = Vec::new();
    let mut rows = Vec::new();
    while cursor < archive.len() {
        let (key, consumed) = read_varint(&archive[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote chart-archive key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 0) => {
                let (value, consumed) = read_varint(&archive[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote chart type: {message}"))
                })?;
                cursor += consumed;
                chart_type = value;
            }
            (5, 0) => {
                let (value, consumed) = read_varint(&archive[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote chart direction: {message}"))
                })?;
                cursor += consumed;
                direction = value;
            }
            (7, 2) => {
                let grid =
                    read_iwa_length_delimited(archive, &mut cursor, part, "Keynote chart grid")?;
                (row_names, column_names, rows) = keynote_chart_grid(grid, part)?;
            }
            (10, 2) => {
                let reference = read_iwa_length_delimited(
                    archive,
                    &mut cursor,
                    part,
                    "Keynote chart title reference",
                )?;
                title_reference = Some(parse_iwa_reference(reference)?);
            }
            (18, 2) => {
                let styles = read_iwa_length_delimited(
                    archive,
                    &mut cursor,
                    part,
                    "Keynote chart series styles",
                )?;
                style_references = keynote_chart_style_references(styles, part)?;
            }
            (_, wire) => skip_protobuf_value(archive, &mut cursor, wire, part)?,
        }
    }
    let (kind, stacked, three_d) = match chart_type {
        1 => (KeynoteChartKind::Column, false, false),
        2 => (KeynoteChartKind::Bar, false, false),
        3 => (KeynoteChartKind::Line, false, false),
        4 => (KeynoteChartKind::Area, false, false),
        5 => (KeynoteChartKind::Pie, false, false),
        6 => (KeynoteChartKind::Column, true, false),
        7 => (KeynoteChartKind::Bar, true, false),
        8 => (KeynoteChartKind::Area, true, false),
        12 => (KeynoteChartKind::Column, false, true),
        13 => (KeynoteChartKind::Bar, false, true),
        14 => (KeynoteChartKind::Line, false, true),
        15 => (KeynoteChartKind::Area, false, true),
        16 => (KeynoteChartKind::Pie, false, true),
        17 => (KeynoteChartKind::Column, true, true),
        18 => (KeynoteChartKind::Bar, true, true),
        19 => (KeynoteChartKind::Area, true, true),
        _ => return Ok(None),
    };
    let by_column = direction == 2
        || (kind == KeynoteChartKind::Pie
            && rows.len() > 1
            && rows.iter().all(|series| series.len() == 1));
    let (series_names, categories) = if by_column {
        rows = keynote_transpose_chart_series(rows);
        (column_names, row_names)
    } else {
        (row_names, column_names)
    };
    rows.retain(|series| !series.is_empty());
    if rows.is_empty() {
        return Ok(None);
    }
    let title = title_reference
        .map(|identifier| keynote_chart_title(archives, identifier, part))
        .transpose()?
        .flatten();
    let colors =
        keynote_chart_series_colors(stylesheet_archives, &style_references, kind, three_d, part)?;
    let chart_style = iwork_chart_properties(
        stylesheet_archives,
        numbers_references(archive, 9, part)?
            .first()
            .copied()
            .unwrap_or(0),
        IWORK_CHART_STYLE_TYPE,
        part,
    )?;
    let chart_nonstyle = iwork_chart_properties(
        archives,
        title_reference.unwrap_or(0),
        IWORK_CHART_TITLE_TYPE,
        part,
    )?;
    let paragraph_styles = numbers_references(archive, 20, part)?;
    let title_style = iwork_chart_text_style(
        &chart_style,
        20,
        &paragraph_styles,
        stylesheet_archives,
        part,
        KeynoteTextStyle {
            font_family: "Helvetica".to_owned(),
            font_size: 12.0,
            bold: true,
            ..KeynoteTextStyle::default()
        },
    )?;
    let axis = |value: bool| -> Result<IworkChartAxis, Diagnostic> {
        let style = iwork_chart_properties(
            stylesheet_archives,
            numbers_references(archive, if value { 13 } else { 15 }, part)?
                .first()
                .copied()
                .unwrap_or(0),
            IWORK_CHART_AXIS_STYLE_TYPE,
            part,
        )?;
        let nonstyle = iwork_chart_properties(
            archives,
            numbers_references(archive, if value { 14 } else { 16 }, part)?
                .first()
                .copied()
                .unwrap_or(0),
            IWORK_CHART_AXIS_NONSTYLE_TYPE,
            part,
        )?;
        let text = iwork_chart_text_style(
            &style,
            if value { 8 } else { 6 },
            &paragraph_styles,
            stylesheet_archives,
            part,
            KeynoteTextStyle {
                font_family: "Helvetica".to_owned(),
                font_size: 9.0,
                ..KeynoteTextStyle::default()
            },
        )?;
        let grid = if numbers_varint(&style, if value { 28 } else { 27 }, part)? == Some(1) {
            numbers_first_bytes(&style, if value { 17 } else { 16 }, part)?
                .map(|bytes| keynote_stroke(bytes, part))
                .transpose()?
        } else {
            None
        };
        Ok(IworkChartAxis {
            text,
            grid,
            show_axis: numbers_varint(&style, if value { 25 } else { 24 }, part)?.unwrap_or(1) != 0,
            show_labels: numbers_varint(&nonstyle, if value { 11 } else { 9 }, part)?.unwrap_or(1)
                != 0,
        })
    };
    let series_nonstyles = numbers_first_bytes(archive, 19, part)?
        .map(|bytes| keynote_chart_style_references(bytes, part))
        .transpose()?
        .unwrap_or_default();
    let show_data_labels = (0..rows.len())
        .map(|index| {
            let properties = iwork_chart_properties(
                archives,
                series_nonstyles.get(index).copied().unwrap_or(0),
                IWORK_CHART_SERIES_NONSTYLE_TYPE,
                part,
            )?;
            let field = match kind {
                KeynoteChartKind::Bar | KeynoteChartKind::Column => 39,
                KeynoteChartKind::Area => 38,
                KeynoteChartKind::Line => 42,
                KeynoteChartKind::Pie => 44,
            };
            Ok(numbers_varint(&properties, field, part)?
                .or(numbers_varint(&properties, 41, part)?)
                .unwrap_or(0)
                != 0)
        })
        .collect::<Result<Vec<_>, Diagnostic>>()?;
    let background = numbers_first_bytes(&chart_style, 8, part)?
        .map(|bytes| keynote_fill(bytes, part, geometry.width, geometry.height))
        .transpose()?
        .unwrap_or(Paint::None);
    let background_stroke = numbers_first_bytes(&chart_style, 9, part)?
        .map(|bytes| keynote_stroke(bytes, part))
        .transpose()?
        .unwrap_or((Paint::None, 0.0, StrokeStyle::default()));
    let plot_fill = numbers_first_bytes(&chart_style, 14, part)?
        .map(|bytes| keynote_fill(bytes, part, geometry.width, geometry.height))
        .transpose()?
        .unwrap_or(Paint::None);
    let bar_gaps = keynote_fixed32_field(&chart_style, 17, part)?
        .map(|gap| {
            Ok((
                keynote_fixed32_field(&chart_style, 16, part)?
                    .unwrap_or(0.0)
                    .max(0.0)
                    / 100.0,
                gap.max(0.0) / 100.0,
            ))
        })
        .transpose()?;
    let show_legend = numbers_varint(&chart_nonstyle, 20, part)?
        .map_or(series_names.len() > 1, |value| value != 0);
    Ok(Some(KeynoteChart {
        geometry,
        kind,
        stacked,
        title,
        series_names,
        categories,
        show_data_labels,
        show_legend,
        background,
        background_stroke,
        plot_fill,
        bar_gaps,
        title_style,
        value_axis: axis(true)?,
        category_axis: axis(false)?,
        colors,
        series: rows,
    }))
}

// TSCH generated property archives inherit by field presence, including explicit
// false and empty fill/stroke values. First-field readers preserve child overrides.
fn iwork_chart_properties(
    archives: &[IwaArchive],
    mut identifier: u64,
    message_type: u64,
    part: &str,
) -> Result<Vec<u8>, Diagnostic> {
    let mut result = Vec::new();
    let mut path = Vec::new();
    while identifier != 0 && path.len() < 64 && !path.contains(&identifier) {
        let Some(style) = archive_message(archives, identifier, message_type, part)? else {
            break;
        };
        path.push(identifier);
        if let Some(properties) = numbers_first_bytes(&style.payload, 10_000, part)? {
            result.extend_from_slice(properties);
        }
        identifier = numbers_first_bytes(&style.payload, 1, part)?
            .map(|base| keynote_style_parent(base, part))
            .transpose()?
            .flatten()
            .unwrap_or(0);
    }
    Ok(result)
}

fn iwork_chart_text_style(
    properties: &[u8],
    field: u64,
    references: &[u64],
    archives: &[IwaArchive],
    part: &str,
    fallback: KeynoteTextStyle,
) -> Result<KeynoteTextStyle, Diagnostic> {
    let index = numbers_varint(properties, field, part)?.or(numbers_varint(properties, 7, part)?);
    match index
        .and_then(|index| usize::try_from(index).ok())
        .and_then(|index| references.get(index))
    {
        Some(identifier) => keynote_text_style(archives, *identifier, part),
        None => Ok(fallback),
    }
}

fn keynote_chart_style_references(bytes: &[u8], part: &str) -> Result<Vec<u64>, Diagnostic> {
    let mut references = Vec::new();
    for entry in numbers_bytes(bytes, 2, part)? {
        let Some(index) = numbers_varint(entry, 1, part)? else {
            continue;
        };
        let Some(reference) = numbers_first_bytes(entry, 2, part)? else {
            continue;
        };
        let index = usize::try_from(index)
            .map_err(|_| iwa_error(part, "Keynote chart series index is too large"))?;
        if index >= MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                part,
                "Keynote chart contains too many series styles",
            ));
        }
        references.resize(references.len().max(index + 1), 0);
        references[index] = parse_iwa_reference(reference)?;
    }
    Ok(references)
}

fn keynote_chart_series_colors(
    archives: &[IwaArchive],
    references: &[u64],
    kind: KeynoteChartKind,
    three_d: bool,
    part: &str,
) -> Result<Vec<u32>, Diagnostic> {
    let fill_field = match (kind, three_d) {
        (KeynoteChartKind::Area, true) => 6,
        (KeynoteChartKind::Bar, true) => 7,
        (KeynoteChartKind::Column, true) => 8,
        (KeynoteChartKind::Line, true) => 9,
        (KeynoteChartKind::Pie, true) => 10,
        (KeynoteChartKind::Area, false) => 11,
        (KeynoteChartKind::Bar, false) => 12,
        (KeynoteChartKind::Column, false) => 13,
        (KeynoteChartKind::Pie, false) => 17,
        (KeynoteChartKind::Line, false) => 14,
    };
    let mut colors = Vec::with_capacity(references.len());
    for identifier in references {
        colors.push(
            keynote_chart_series_color(archives, *identifier, fill_field, part, &mut Vec::new())?
                .unwrap_or(0),
        );
    }
    Ok(colors)
}

fn keynote_chart_series_color(
    archives: &[IwaArchive],
    identifier: u64,
    fill_field: u64,
    part: &str,
    path: &mut Vec<u64>,
) -> Result<Option<u32>, Diagnostic> {
    if path.len() >= 64 || path.contains(&identifier) {
        return Ok(None);
    }
    let Some(style) = archive_message(archives, identifier, IWORK_CHART_SERIES_STYLE_TYPE, part)?
    else {
        return Ok(None);
    };
    path.push(identifier);
    let properties = numbers_first_bytes(&style.payload, 10_000, part)?;
    let color = properties
        .map(|properties| {
            numbers_first_bytes(properties, fill_field, part)?.map_or_else(
                || numbers_first_bytes(properties, 14, part),
                |fill| Ok(Some(fill)),
            )
        })
        .transpose()?
        .flatten()
        .map(|fill| numbers_first_bytes(fill, 1, part))
        .transpose()?
        .flatten()
        .map(|color| keynote_color(color, part))
        .transpose()?;
    let color = if color.is_some() {
        color
    } else if let Some(parent) = keynote_nested_message(&style.payload, 1, part)?
        .map(|base| keynote_style_parent(base, part))
        .transpose()?
        .flatten()
    {
        keynote_chart_series_color(archives, parent, fill_field, part, path)?
    } else {
        None
    };
    path.pop();
    Ok(color)
}

fn keynote_chart_grid(
    bytes: &[u8],
    part: &str,
) -> Result<(Vec<String>, Vec<String>, Vec<Vec<f32>>), Diagnostic> {
    let mut cursor = 0_usize;
    let mut row_names = Vec::new();
    let mut column_names = Vec::new();
    let mut rows = Vec::new();
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote chart-grid key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (field @ (1 | 2), 2) => {
                let name =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote chart grid name")?;
                if let Ok(name) = std::str::from_utf8(name) {
                    if field == 1 {
                        row_names.push(name.to_owned());
                    } else {
                        column_names.push(name.to_owned());
                    }
                }
            }
            (3, 2) => {
                let row = read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote chart row")?;
                rows.push(keynote_chart_row(row, part)?);
                if rows.len() >= MAX_IWA_PROTOBUF_FIELDS {
                    break;
                }
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok((row_names, column_names, rows))
}

fn keynote_chart_title(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
) -> Result<Option<String>, Diagnostic> {
    let Some(title) = archive_message(archives, identifier, IWORK_CHART_TITLE_TYPE, part)? else {
        return Ok(None);
    };
    let Some(properties) = keynote_nested_message(&title.payload, 10_000, part)? else {
        return Ok(None);
    };
    if numbers_varint(properties, 21, part)? == Some(0) {
        return Ok(None);
    }
    Ok(numbers_first_bytes(properties, 23, part)?
        .and_then(|value| std::str::from_utf8(value).ok())
        .filter(|value| !value.is_empty())
        .map(str::to_owned))
}

fn keynote_chart_row(bytes: &[u8], part: &str) -> Result<Vec<f32>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut values = Vec::new();
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote chart-row key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let value =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote chart value")?;
                values.push(keynote_chart_value(value, part)?.unwrap_or(0.0));
                if values.len() >= MAX_IWA_PROTOBUF_FIELDS {
                    break;
                }
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(values)
}

fn keynote_chart_value(bytes: &[u8], part: &str) -> Result<Option<f32>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote chart-value key: {message}"))
        })?;
        cursor += consumed;
        if key >> 3 == 1 && key & 7 == 1 {
            let end = cursor
                .checked_add(8)
                .filter(|end| *end <= bytes.len())
                .ok_or_else(|| iwa_error(part, "truncated Keynote chart numeric value"))?;
            let value = f64::from_le_bytes(
                bytes[cursor..end]
                    .try_into()
                    .map_err(|_| iwa_error(part, "invalid Keynote chart numeric value"))?,
            );
            return Ok(value.is_finite().then_some(value as f32));
        }
        skip_protobuf_value(bytes, &mut cursor, key & 7, part)?;
    }
    Ok(None)
}

fn keynote_transpose_chart_series(rows: Vec<Vec<f32>>) -> Vec<Vec<f32>> {
    let columns = rows.iter().map(Vec::len).max().unwrap_or(0);
    (0..columns)
        .map(|column| {
            rows.iter()
                .map(|row| row.get(column).copied().unwrap_or(0.0))
                .collect()
        })
        .collect()
}

#[allow(clippy::type_complexity)]
fn keynote_text_shape(
    bytes: &[u8],
    part: &str,
) -> Result<
    (
        KeynoteGeometry,
        Option<u64>,
        Option<u64>,
        Geometry,
        bool,
        bool,
    ),
    Diagnostic,
> {
    let mut cursor = 0_usize;
    let mut geometry = KeynoteGeometry::default();
    let mut storage = None;
    let mut style = None;
    let mut shape_geometry = Geometry::Rectangle;
    let mut auto_width = false;
    let mut auto_height = false;
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                part,
                "Keynote text shape contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote text-shape key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let shape =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote native shape")?;
                let (authored_geometry, _, parsed_shape_geometry) = keynote_shape(shape, part)?;
                auto_width = authored_geometry.width <= f32::EPSILON;
                auto_height = authored_geometry.height <= f32::EPSILON;
                geometry = keynote_shape_geometry(shape, part)?;
                style = keynote_shape_style_reference(shape, part)?;
                shape_geometry = parsed_shape_geometry;
            }
            (field @ (2 | 4), 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote text storage reference",
                )?;
                let reference = parse_iwa_reference(reference)?;
                if field == 2 || storage.is_none() {
                    storage = Some(reference);
                }
            }
            (1 | 2 | 4, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote text-shape field has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok((
        geometry,
        storage,
        style,
        shape_geometry,
        auto_width,
        auto_height,
    ))
}

fn keynote_shape_style_reference(bytes: &[u8], part: &str) -> Result<Option<u64>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote shape-style key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (2, 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote shape-style reference",
                )?;
                return Ok(Some(parse_iwa_reference(reference)?));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(None)
}

fn keynote_shape_geometry(bytes: &[u8], part: &str) -> Result<KeynoteGeometry, Diagnostic> {
    let mut cursor = 0_usize;
    let mut geometry = KeynoteGeometry::default();
    let mut path_source = None;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote shape key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let drawable =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote shape drawable")?;
                geometry = keynote_drawable_geometry(drawable, part)?;
            }
            (3, 2) => {
                path_source = Some(read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote shape path source",
                )?);
            }
            (1 | 3, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote shape drawable has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    if (geometry.width <= f32::EPSILON || geometry.height <= f32::EPSILON)
        && let Some(path_source) = path_source
        && let Some(bezier) = keynote_nested_message(path_source, 5, part)?
        && let Some(size) = keynote_nested_message(bezier, 2, part)?
    {
        let (width, height) = iwork_float_pair(size, part)?;
        if geometry.width <= f32::EPSILON {
            geometry.width = width;
        }
        if geometry.height <= f32::EPSILON {
            geometry.height = height;
        }
    }
    Ok(geometry)
}

fn keynote_path_source(
    bytes: &[u8],
    width: f32,
    height: f32,
    part: &str,
) -> Result<Option<Geometry>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote path-source key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (3, 2) => {
                let point =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote point path")?;
                return keynote_point_path(point, width, height, part).map(Some);
            }
            (4, 2) => {
                let scalar =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote scalar path")?;
                return keynote_scalar_path(scalar, width, height, part).map(Some);
            }
            (5, 2) => {
                let bezier =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote bezier path")?;
                return keynote_bezier_path(bezier, width, height, part);
            }
            (7, 2) => {
                let connection = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote connection-line path",
                )?;
                return keynote_connection_path(connection, width, height, part);
            }
            (8, 2) => {
                let editable = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote editable bezier path",
                )?;
                return keynote_editable_bezier_path(editable, width, height, part);
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(None)
}

fn keynote_editable_bezier_path(
    bytes: &[u8],
    width: f32,
    height: f32,
    part: &str,
) -> Result<Option<Geometry>, Diagnostic> {
    let (natural_width, natural_height) = keynote_nested_message(bytes, 2, part)?
        .map(|size| iwork_float_pair(size, part))
        .transpose()?
        .unwrap_or((width, height));
    let mut commands = Vec::new();
    for subpath in numbers_bytes(bytes, 1, part)? {
        let mut nodes = Vec::new();
        for node in numbers_bytes(subpath, 1, part)? {
            if nodes.len() + commands.len() >= Limits::HARD_MAX.max_document_objects {
                return Err(iwa_error(
                    part,
                    "editable bezier path contains too many nodes",
                ));
            }
            let mut points = [(0.0, 0.0); 3];
            for (index, point) in points.iter_mut().enumerate() {
                let Some(value) = keynote_nested_message(node, index as u64 + 1, part)? else {
                    return Err(iwa_error(
                        part,
                        "editable bezier node is missing a control point",
                    ));
                };
                *point = iwork_float_pair(value, part)?;
            }
            nodes.push(points);
        }
        let Some(first) = nodes.first() else { continue };
        commands.push(PathCommand::MoveTo {
            x: first[1].0,
            y: first[1].1,
        });
        let closed = keynote_varint_field(subpath, 2, part)? == Some(1);
        let ends = nodes.iter().skip(1).chain(closed.then_some(first));
        for (start, end) in nodes.iter().zip(ends) {
            commands.push(if start[2] == start[1] && end[0] == end[1] {
                PathCommand::LineTo {
                    x: end[1].0,
                    y: end[1].1,
                }
            } else {
                PathCommand::BezierCurveTo {
                    cp1x: start[2].0,
                    cp1y: start[2].1,
                    cp2x: end[0].0,
                    cp2y: end[0].1,
                    x: end[1].0,
                    y: end[1].1,
                }
            });
        }
        if closed {
            commands.push(PathCommand::ClosePath);
        }
    }
    let scale_x = if natural_width > f32::EPSILON {
        width / natural_width
    } else {
        1.0
    };
    let scale_y = if natural_height > f32::EPSILON {
        height / natural_height
    } else {
        1.0
    };
    Ok((!commands.is_empty()).then(|| Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands: keynote_scale_path(commands, scale_x, scale_y),
    }))
}

fn keynote_connection_path(
    bytes: &[u8],
    width: f32,
    height: f32,
    part: &str,
) -> Result<Option<Geometry>, Diagnostic> {
    let Some(bezier) = keynote_nested_message(bytes, 1, part)? else {
        return Ok(None);
    };
    let Some(path) = keynote_nested_message(bezier, 3, part)? else {
        return Ok(None);
    };
    let (natural_width, natural_height) = keynote_nested_message(bezier, 2, part)?
        .map(|size| iwork_float_pair(size, part))
        .transpose()?
        .unwrap_or((width, height));
    let scale_x = if natural_width > f32::EPSILON {
        width / natural_width
    } else {
        1.0
    };
    let scale_y = if natural_height > f32::EPSILON {
        height / natural_height
    } else {
        1.0
    };
    let points = keynote_path(path, scale_x, scale_y, part)?
        .into_iter()
        .filter_map(|command| match command {
            PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => Some((x, y)),
            _ => None,
        })
        .collect::<Vec<_>>();
    let [(from_x, from_y), (middle_x, middle_y), (to_x, to_y)] = points.as_slice() else {
        return Ok(None);
    };
    let commands = if keynote_varint_field(bytes, 2, part)? == Some(1) {
        vec![
            PathCommand::MoveTo {
                x: *from_x,
                y: *from_y,
            },
            PathCommand::LineTo {
                x: *from_x,
                y: *middle_y,
            },
            PathCommand::LineTo {
                x: *to_x,
                y: *middle_y,
            },
            PathCommand::LineTo { x: *to_x, y: *to_y },
        ]
    } else {
        let direction_x = 0.2 * (to_x - from_x);
        let direction_y = 0.2 * (to_y - from_y);
        vec![
            PathCommand::MoveTo {
                x: *from_x,
                y: *from_y,
            },
            PathCommand::QuadraticCurveTo {
                cpx: middle_x - direction_x,
                cpy: middle_y - direction_y,
                x: *middle_x,
                y: *middle_y,
            },
            PathCommand::QuadraticCurveTo {
                cpx: middle_x + direction_x,
                cpy: middle_y + direction_y,
                x: *to_x,
                y: *to_y,
            },
        ]
    };
    Ok(Some(Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    }))
}

fn keynote_scalar_path(
    bytes: &[u8],
    width: f32,
    height: f32,
    part: &str,
) -> Result<Geometry, Diagnostic> {
    let mut cursor = 0_usize;
    let mut kind = 0_u64;
    let mut scalar = 0.0_f32;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote scalar-path key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote scalar-path type: {message}"))
                })?;
                cursor += consumed;
                kind = value;
            }
            (2, 5) => scalar = keynote_fixed32(bytes, &mut cursor, part, "path scalar")?,
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(match kind {
        0 => {
            let radius = scalar.abs().min(width.min(height) / 2.0);
            Geometry::RoundedRectangle {
                radius_x: radius,
                radius_y: radius,
            }
        }
        1 => keynote_regular_polygon(scalar, width, height),
        _ => Geometry::Rectangle,
    })
}

fn keynote_regular_polygon(sides: f32, width: f32, height: f32) -> Geometry {
    let sides = (sides.round() as usize).clamp(3, 100);
    let mut commands = (0..sides)
        .map(|index| {
            let angle =
                -std::f32::consts::FRAC_PI_2 + std::f32::consts::TAU * index as f32 / sides as f32;
            (angle.cos(), angle.sin())
        })
        .enumerate()
        .map(|(index, (x, y))| {
            let x = (x + 1.0) * width / 2.0;
            let y = (y + 1.0) * height / 2.0;
            if index == 0 {
                PathCommand::MoveTo { x, y }
            } else {
                PathCommand::LineTo { x, y }
            }
        })
        .collect::<Vec<_>>();
    commands.push(PathCommand::ClosePath);
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    }
}

fn keynote_bezier_path(
    bytes: &[u8],
    width: f32,
    height: f32,
    part: &str,
) -> Result<Option<Geometry>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut natural_size = None;
    let mut path = None;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote bezier-path key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (2, 2) => {
                let size = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote path natural size",
                )?;
                natural_size = Some(iwork_float_pair(size, part)?);
            }
            (3, 2) => {
                path = Some(read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote path elements",
                )?);
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    let Some(path) = path else {
        return Ok(None);
    };
    let (natural_width, natural_height) = natural_size.unwrap_or((width, height));
    let raw_commands = keynote_path(path, 1.0, 1.0, part)?;
    let (coordinate_width, coordinate_height) = keynote_path_coordinate_extent(&raw_commands);
    let scale_x = if coordinate_width > f32::EPSILON {
        natural_width / coordinate_width
    } else {
        1.0
    };
    let scale_y = if coordinate_height > f32::EPSILON {
        natural_height / coordinate_height
    } else {
        1.0
    };
    let commands = keynote_scale_path(raw_commands, scale_x, scale_y);
    Ok((!commands.is_empty()).then_some(Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    }))
}

fn keynote_path_coordinate_extent(commands: &[PathCommand]) -> (f32, f32) {
    let mut extent = (0.0_f32, 0.0_f32);
    let mut include = |x: f32, y: f32| {
        extent.0 = extent.0.max(x);
        extent.1 = extent.1.max(y);
    };
    for command in commands {
        match command {
            PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => include(*x, *y),
            PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => {
                include(*cpx, *cpy);
                include(*x, *y);
            }
            PathCommand::BezierCurveTo {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                include(*cp1x, *cp1y);
                include(*cp2x, *cp2y);
                include(*x, *y);
            }
            PathCommand::ClosePath => {}
        }
    }
    extent
}

fn keynote_scale_path(
    mut commands: Vec<PathCommand>,
    scale_x: f32,
    scale_y: f32,
) -> Vec<PathCommand> {
    for command in &mut commands {
        match command {
            PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => {
                *x *= scale_x;
                *y *= scale_y;
            }
            PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => {
                *cpx *= scale_x;
                *cpy *= scale_y;
                *x *= scale_x;
                *y *= scale_y;
            }
            PathCommand::BezierCurveTo {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                *cp1x *= scale_x;
                *cp1y *= scale_y;
                *cp2x *= scale_x;
                *cp2y *= scale_y;
                *x *= scale_x;
                *y *= scale_y;
            }
            PathCommand::ClosePath => {}
        }
    }
    commands
}

fn keynote_path(
    bytes: &[u8],
    scale_x: f32,
    scale_y: f32,
    part: &str,
) -> Result<Vec<PathCommand>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut commands = Vec::new();
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote path key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let element =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote path element")?;
                if let Some(command) = keynote_path_element(element, scale_x, scale_y, part)? {
                    commands.push(command);
                }
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(commands)
}

fn keynote_path_element(
    bytes: &[u8],
    scale_x: f32,
    scale_y: f32,
    part: &str,
) -> Result<Option<PathCommand>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut kind = None;
    let mut points = Vec::new();
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote path-element key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(
                        part,
                        format!("invalid Keynote path-element type: {message}"),
                    )
                })?;
                cursor += consumed;
                kind = Some(value);
            }
            (2, 2) => {
                let point =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote path point")?;
                let (x, y) = iwork_float_pair(point, part)?;
                points.push((x * scale_x, y * scale_y));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(match (kind, points.as_slice()) {
        (Some(1), [(x, y), ..]) => Some(PathCommand::MoveTo { x: *x, y: *y }),
        (Some(2), [(x, y), ..]) => Some(PathCommand::LineTo { x: *x, y: *y }),
        (Some(3), [(cpx, cpy), (x, y), ..]) => Some(PathCommand::QuadraticCurveTo {
            cpx: *cpx,
            cpy: *cpy,
            x: *x,
            y: *y,
        }),
        (Some(4), [(cp1x, cp1y), (cp2x, cp2y), (x, y), ..]) => Some(PathCommand::BezierCurveTo {
            cp1x: *cp1x,
            cp1y: *cp1y,
            cp2x: *cp2x,
            cp2y: *cp2y,
            x: *x,
            y: *y,
        }),
        (Some(5), _) => Some(PathCommand::ClosePath),
        _ => None,
    })
}

fn keynote_drawable_geometry(bytes: &[u8], part: &str) -> Result<KeynoteGeometry, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote drawable key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let geometry = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote drawable geometry",
                )?;
                return keynote_geometry(geometry, part);
            }
            (1, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote drawable geometry has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(KeynoteGeometry::default())
}

fn keynote_geometry(bytes: &[u8], part: &str) -> Result<KeynoteGeometry, Diagnostic> {
    let mut cursor = 0_usize;
    let mut geometry = KeynoteGeometry::default();
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(part, "Keynote geometry contains too many fields"));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote geometry key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (field @ (1 | 2), 2) => {
                let pair =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote geometry pair")?;
                let (first, second) = iwork_float_pair(pair, part)?;
                if field == 1 {
                    geometry.x = first;
                    geometry.y = second;
                } else {
                    geometry.width = first.max(0.0);
                    geometry.height = second.max(0.0);
                }
            }
            (1 | 2, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote geometry field has an invalid protobuf wire type",
                ));
            }
            (4, 5) => {
                geometry.rotation_degrees =
                    keynote_fixed32(bytes, &mut cursor, part, "geometry rotation")?;
            }
            (3, 0) => {
                let (flags, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote geometry flags: {message}"))
                })?;
                cursor += consumed;
                geometry.resize_flags = Some(flags);
                geometry.flip_horizontal = flags & 4 != 0;
                geometry.flip_vertical = flags & 8 != 0;
            }
            (4, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote geometry rotation has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(geometry)
}

fn iwork_float_pair(bytes: &[u8], part: &str) -> Result<(f32, f32), Diagnostic> {
    let mut cursor = 0_usize;
    let mut first = None;
    let mut second = None;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid iWork float-pair key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (field @ (1 | 2), 5) => {
                let end = cursor
                    .checked_add(4)
                    .ok_or_else(|| iwa_error(part, "iWork float-pair offset overflows"))?;
                let encoded: [u8; 4] = bytes
                    .get(cursor..end)
                    .ok_or_else(|| iwa_error(part, "iWork float pair is truncated"))?
                    .try_into()
                    .map_err(|_| iwa_error(part, "iWork float pair is invalid"))?;
                let value = f32::from_le_bytes(encoded);
                if !value.is_finite() {
                    return Err(iwa_error(part, "iWork float pair is not finite"));
                }
                if field == 1 {
                    first = Some(value);
                } else {
                    second = Some(value);
                }
                cursor = end;
            }
            (1 | 2, _) => {
                return Err(iwa_error(
                    part,
                    "iWork float-pair field has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok((first.unwrap_or(0.0), second.unwrap_or(0.0)))
}

fn keynote_storage_text(
    bytes: &[u8],
    part: &str,
    limits: Limits,
) -> Result<KeynoteTextStorage, Diagnostic> {
    let mut cursor = 0_usize;
    let mut storage = KeynoteTextStorage::default();
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                part,
                "Keynote text storage contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote text-storage key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (3, 2) => {
                let encoded = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote text-storage string",
                )?;
                let value = std::str::from_utf8(encoded)
                    .map_err(|_| iwa_error(part, "Keynote text-storage string is not UTF-8"))?;
                storage.text.try_reserve(value.len()).map_err(|_| {
                    Diagnostic::fatal(
                        DiagnosticCode::AllocationFailed,
                        Phase::Parse,
                        None,
                        "unable to allocate Keynote text",
                    )
                    .in_part(part)
                })?;
                storage.text.push_str(value);
            }
            (6, 2) => {
                let table =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Pages paragraph data")?;
                for entry in numbers_bytes(table, 1, part)? {
                    let character_index = numbers_varint(entry, 1, part)?
                        .ok_or_else(|| {
                            iwa_error(part, "iWork paragraph data has no character index")
                        })?
                        .try_into()
                        .map_err(|_| iwa_error(part, "iWork paragraph data index is too large"))?;
                    let list_level = numbers_varint(entry, 2, part)?
                        .unwrap_or(0)
                        .try_into()
                        .map_err(|_| iwa_error(part, "iWork list level is too large"))?;
                    if storage
                        .paragraph_data
                        .last()
                        .is_some_and(|change| change.character_index >= character_index)
                    {
                        return Err(iwa_error(
                            part,
                            "iWork paragraph data indexes are not strictly increasing",
                        ));
                    }
                    storage.paragraph_data.push(KeynoteParagraphDataChange {
                        character_index,
                        list_level,
                    });
                }
            }
            (field @ (5 | 7 | 8 | 9 | 17 | 28), 2) => {
                let table = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote text-style table",
                )?;
                let styles = keynote_attribute_references(table, part, limits)?;
                match field {
                    5 => storage.paragraph_styles = styles,
                    7 => storage.list_styles = styles,
                    8 => storage.character_styles = styles,
                    9 => storage.attachments = styles,
                    17 => storage.sections = styles,
                    28 => storage.drop_cap_styles = styles,
                    _ => unreachable!(),
                }
            }
            (3 | 5 | 7 | 8 | 9 | 17, _) => {
                return Err(iwa_error(
                    part,
                    "Keynote text-storage field has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(storage)
}

fn keynote_attribute_references(
    bytes: &[u8],
    part: &str,
    limits: Limits,
) -> Result<Vec<KeynoteStyleChange>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut changes = Vec::new();
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        // Repeated text-style references scale with document content, not metadata fields.
        if fields > limits.max_relationship_edges {
            return Err(iwa_error(
                part,
                "iWork attribute table exceeds the configured reference limit",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote attribute-table key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let entry = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote attribute-table entry",
                )?;
                let mut entry_cursor = 0_usize;
                let mut character_index = None;
                let mut identifier = None;
                while entry_cursor < entry.len() {
                    let (entry_key, consumed) =
                        read_varint(&entry[entry_cursor..]).map_err(|message| {
                            iwa_error(
                                part,
                                format!("invalid Keynote attribute entry key: {message}"),
                            )
                        })?;
                    entry_cursor += consumed;
                    match (entry_key >> 3, entry_key & 7) {
                        (1, 0) => {
                            let (value, consumed) =
                                read_varint(&entry[entry_cursor..]).map_err(|message| {
                                    iwa_error(
                                        part,
                                        format!(
                                            "invalid Keynote attribute character index: {message}"
                                        ),
                                    )
                                })?;
                            entry_cursor += consumed;
                            character_index = Some(usize::try_from(value).map_err(|_| {
                                iwa_error(part, "Keynote attribute character index is too large")
                            })?);
                        }
                        (2, 2) => {
                            let reference = read_iwa_length_delimited(
                                entry,
                                &mut entry_cursor,
                                part,
                                "Keynote attribute style reference",
                            )?;
                            identifier = Some(parse_iwa_reference(reference)?);
                        }
                        (_, wire) => skip_protobuf_value(entry, &mut entry_cursor, wire, part)?,
                    }
                }
                let character_index = character_index.ok_or_else(|| {
                    iwa_error(part, "Keynote attribute entry has no character index")
                })?;
                if changes.last().is_some_and(|change: &KeynoteStyleChange| {
                    change.character_index >= character_index
                }) {
                    return Err(iwa_error(
                        part,
                        "Keynote attribute character indexes are not strictly increasing",
                    ));
                }
                changes.try_reserve(1).map_err(|_| {
                    Diagnostic::fatal(
                        DiagnosticCode::AllocationFailed,
                        Phase::Parse,
                        None,
                        "unable to allocate iWork attribute references",
                    )
                    .in_part(part)
                })?;
                changes.push(KeynoteStyleChange {
                    character_index,
                    identifier,
                });
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(changes)
}

fn keynote_text_runs(
    storage: &KeynoteTextStorage,
    archives: &[IwaArchive],
    part: &str,
    fallback_style: &KeynoteTextStyle,
    limits: Limits,
) -> Result<Vec<TextRun>, Diagnostic> {
    let text_length = storage.text.encode_utf16().count();
    let mut boundaries = vec![0, text_length];
    boundaries.extend(
        storage
            .paragraph_styles
            .iter()
            .chain(&storage.list_styles)
            .chain(&storage.character_styles)
            .map(|change| change.character_index),
    );
    boundaries.sort_unstable();
    boundaries.dedup();

    if boundaries.iter().any(|boundary| *boundary > text_length) {
        return Err(iwa_error(
            part,
            "Keynote text style starts beyond the stored text",
        ));
    }

    let mut character_style = None;
    let mut list_counters = HashMap::<(u64, usize), usize>::new();
    let mut runs: Vec<TextRun> = Vec::new();
    for pair in boundaries.windows(2) {
        let start = pair[0];
        let end = pair[1];
        let paragraph_style = iwork_inherited_style_at(&storage.paragraph_styles, start);
        let list_style = iwork_inherited_style_at(&storage.list_styles, start);
        update_keynote_style_at(&storage.character_styles, start, &mut character_style);
        if start == end {
            continue;
        }
        let text = iwork_text_with_fields(storage, archives, part, start, end, limits)?;
        let start = keynote_utf16_byte_offset(&storage.text, start)
            .ok_or_else(|| iwa_error(part, "Keynote text style splits a UTF-16 surrogate pair"))?;
        let paragraph_start = start == 0 || storage.text[..start].ends_with('\n');
        if text.is_empty() {
            continue;
        }
        let mut style = paragraph_style
            .map(|identifier| keynote_text_style(archives, identifier, part))
            .transpose()?
            .unwrap_or_else(|| fallback_style.clone());
        style.list_level = iwork_paragraph_list_level_at(&storage.paragraph_data, pair[0]);
        if let Some(identifier) = character_style {
            keynote_apply_text_style(archives, identifier, part, &mut style, &mut Vec::new())?;
        }
        let list_marker = if paragraph_start && !text.starts_with('\n') {
            list_style
                .map(|identifier| {
                    keynote_list_marker(
                        archives,
                        identifier,
                        style.list_level,
                        part,
                        &mut Vec::new(),
                    )
                    .map(|marker| (identifier, marker))
                })
                .transpose()?
                .and_then(|(identifier, marker)| marker.map(|marker| (identifier, marker)))
        } else {
            None
        };
        if let Some((identifier, marker)) = list_marker {
            let sequence = list_counters
                .entry((identifier, marker.level))
                .and_modify(|value| *value = value.saturating_add(1))
                .or_insert(1);
            let marker_text = keynote_list_marker_text(&marker.kind, *sequence);
            let marker_style = keynote_list_marker_style(&style, &marker);
            if marker.level != 0 {
                let mut indent_run = keynote_text_run(" ".to_owned(), &marker_style);
                indent_run.letter_spacing += marker.indent - marker_style.font_size * 0.278;
                runs.push(indent_run);
            }
            let marker_run = keynote_text_run(marker_text.clone(), &marker_style);
            if let Some(previous) = runs.last_mut()
                && previous.same_style(&marker_run)
            {
                previous.text.push_str(&marker_run.text);
            } else {
                runs.push(marker_run);
            }
            runs.push(keynote_text_run("\t".to_owned(), &style));
        }
        let run = keynote_text_run(text, &style);
        if let Some(previous) = runs.last_mut()
            && previous.same_style(&run)
        {
            previous.text.push_str(&run.text);
        } else {
            runs.push(run);
        }
    }
    Ok(keynote_platform_font_runs(runs))
}

fn iwork_text_with_fields(
    storage: &KeynoteTextStorage,
    archives: &[IwaArchive],
    part: &str,
    start: usize,
    end: usize,
    limits: Limits,
) -> Result<String, Diagnostic> {
    iwork_text_with_dynamic_fields(storage, archives, part, start, end, limits, None)
}

#[allow(clippy::too_many_arguments)]
fn iwork_text_with_dynamic_fields(
    storage: &KeynoteTextStorage,
    archives: &[IwaArchive],
    part: &str,
    start: usize,
    end: usize,
    limits: Limits,
    page_number: Option<&str>,
) -> Result<String, Diagnostic> {
    let start_byte = keynote_utf16_byte_offset(&storage.text, start)
        .ok_or_else(|| iwa_error(part, "iWork field range splits a UTF-16 surrogate pair"))?;
    let end_byte = keynote_utf16_byte_offset(&storage.text, end)
        .filter(|end_byte| *end_byte >= start_byte)
        .ok_or_else(|| iwa_error(part, "iWork field range is invalid"))?;
    let mut text = String::new();
    let mut position = start;
    for character in storage.text[start_byte..end_byte].chars() {
        let mut encoded = [0; 4];
        let value = if character == '\u{fffc}' {
            if let Some(identifier) = storage
                .attachments
                .iter()
                .find(|change| change.character_index == position)
                .and_then(|change| change.identifier)
                && page_number.is_some()
                && archive_message(archives, identifier, IWORK_NUMBER_ATTACHMENT_TYPE, part)?
                    .is_some()
            {
                page_number.unwrap()
            } else if let Some(identifier) = storage
                .attachments
                .iter()
                .find(|change| change.character_index == position)
                .and_then(|change| change.identifier)
                && let Some(field) =
                    archive_message(archives, identifier, IWORK_TEXT_FIELD_TYPE, part)?
                && let Some(value) = numbers_first_bytes(&field.payload, 2, part)?
            {
                std::str::from_utf8(value)
                    .map_err(|_| iwa_error(part, "iWork text field is not UTF-8"))?
            } else {
                ""
            }
        } else {
            character.encode_utf8(&mut encoded)
        };
        if text.len().saturating_add(value.len()) > limits.max_entry_uncompressed_bytes {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ZipEntryTooLarge,
                Phase::Parse,
                None,
                "expanded iWork field text exceeds the configured byte limit",
            )
            .in_part(part));
        }
        text.try_reserve(value.len()).map_err(|_| {
            Diagnostic::fatal(
                DiagnosticCode::AllocationFailed,
                Phase::Parse,
                None,
                "unable to allocate expanded iWork field text",
            )
            .in_part(part)
        })?;
        text.push_str(value);
        position += character.len_utf16();
    }
    Ok(normalize_keynote_text(&text))
}

fn pages_text_runs_for_range(
    storage: &KeynoteTextStorage,
    archives: &[IwaArchive],
    part: &str,
    range_start: usize,
    range_end: usize,
    limits: Limits,
) -> Result<Vec<TextRun>, Diagnostic> {
    pages_text_runs_for_range_with_page_number(
        storage,
        archives,
        part,
        range_start,
        range_end,
        limits,
        None,
    )
}

fn pages_drop_cap(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
) -> Result<crate::model::TextDropCap, Diagnostic> {
    let unsupported = || {
        Diagnostic::warning(
            DiagnosticCode::UnsupportedFeature,
            Phase::Parse,
            Fidelity::Approximate,
            "PAGES_DROP_CAP_UNSUPPORTED: using ordinary text for an invalid or non-text drop cap",
        )
        .in_part(part)
    };
    let mut chain = Vec::new();
    let mut current = Some(identifier);
    while let Some(id) = current {
        if chain.len() >= 64 || chain.iter().any(|(seen, _)| *seen == id) {
            return Err(unsupported());
        }
        let Some(message) = archive_message(archives, id, PAGES_DROP_CAP_STYLE_TYPE, part)? else {
            break;
        };
        chain.push((id, message));
        current = numbers_first_bytes(&message.payload, 1, part)?
            .map(|value| keynote_style_parent(value, part))
            .transpose()?
            .flatten();
    }
    if chain.is_empty() {
        return Err(unsupported());
    }
    let mut cap = crate::model::TextDropCap {
        characters: 1,
        lines: 3,
        raised_lines: 0,
        padding: 0.0,
        outdent: 0.0,
    };
    for (_, message) in chain.into_iter().rev() {
        let Some(properties) = numbers_first_bytes(&message.payload, 12, part)? else {
            continue;
        };
        let Some(bytes) = numbers_first_bytes(properties, 1, part)? else {
            continue;
        };
        if numbers_varint(bytes, 1, part)?.unwrap_or(0) != 0
            || numbers_varint(bytes, 6, part)?.unwrap_or(0) != 0
            || numbers_varint(bytes, 7, part)?.unwrap_or(0) != 0
        {
            return Err(unsupported());
        }
        for (field, target) in [
            (2, &mut cap.lines),
            (3, &mut cap.raised_lines),
            (10, &mut cap.characters),
        ] {
            if let Some(value) = numbers_varint(bytes, field, part)? {
                *target = value.try_into().map_err(|_| unsupported())?;
            }
        }
        let mut cursor = 0;
        while cursor < bytes.len() {
            let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|_| unsupported())?;
            cursor += consumed;
            // Character scale and corner radius belong to background shapes, rejected above.
            let target = match key >> 3 {
                4 | 11 => Some(&mut cap.outdent),
                5 | 12 => Some(&mut cap.padding),
                _ => None,
            };
            if let Some(target) = target {
                match key & 7 {
                    1 => {
                        *target = numbers_le_f64(bytes, cursor, part)? as f32;
                        cursor += 8;
                    }
                    5 => {
                        *target = keynote_fixed32(bytes, &mut cursor, part, "drop cap metric")?;
                    }
                    wire => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
                }
            } else {
                skip_protobuf_value(bytes, &mut cursor, key & 7, part)?;
            }
        }
    }
    if !(1..=32).contains(&cap.characters)
        || !(1..=32).contains(&cap.lines)
        || cap.raised_lines > 32
        || !cap.padding.is_finite()
        || !(0.0..=1024.0).contains(&cap.padding)
        || !cap.outdent.is_finite()
        || cap.outdent.abs() > 1024.0
    {
        return Err(unsupported());
    }
    cap.padding *= IWORK_POINT_TO_CSS_PIXEL;
    cap.outdent *= IWORK_POINT_TO_CSS_PIXEL;
    Ok(cap)
}

#[allow(clippy::too_many_arguments)]
fn pages_text_runs_for_range_with_page_number(
    storage: &KeynoteTextStorage,
    archives: &[IwaArchive],
    part: &str,
    range_start: usize,
    range_end: usize,
    limits: Limits,
    page_number: Option<&str>,
) -> Result<Vec<TextRun>, Diagnostic> {
    let text_length = storage.text.encode_utf16().count();
    if range_start > range_end || range_end > text_length {
        return Err(iwa_error(
            part,
            "Pages saved page range exceeds the body text",
        ));
    }
    let mut boundaries = vec![range_start, range_end];
    let mut caps = Vec::new();
    for change in &storage.drop_cap_styles {
        let Some(id) = change.identifier else {
            continue;
        };
        if change.character_index < range_start || change.character_index >= range_end {
            continue;
        }
        let Ok(cap) = pages_drop_cap(archives, id, part) else {
            continue;
        };
        let Some(byte) = keynote_utf16_byte_offset(&storage.text, change.character_index) else {
            continue;
        };
        let count = storage.text[byte..]
            .chars()
            .take(cap.characters as usize)
            .take_while(|c| !matches!(c, '\n' | '\u{2028}'))
            .map(char::len_utf16)
            .sum::<usize>();
        let end = (change.character_index + count).min(range_end);
        boundaries.extend([change.character_index, end]);
        caps.push((change.character_index, end, id));
    }
    boundaries.extend(
        storage
            .paragraph_styles
            .iter()
            .chain(&storage.list_styles)
            .chain(&storage.character_styles)
            .map(|change| change.character_index)
            .filter(|index| range_start < *index && *index < range_end),
    );
    boundaries.sort_unstable();
    boundaries.dedup();

    let fallback_style = KeynoteTextStyle {
        font_family: "Helvetica".to_owned(),
        font_size: 10.5,
        ..KeynoteTextStyle::default()
    };
    let mut list_counters = HashMap::<(u64, usize), usize>::new();
    let mut runs: Vec<TextRun> = Vec::new();
    for pair in boundaries.windows(2) {
        let start = pair[0];
        let end = pair[1];
        if start == end {
            continue;
        }
        let start_byte = keynote_utf16_byte_offset(&storage.text, start)
            .ok_or_else(|| iwa_error(part, "Pages style splits a UTF-16 surrogate pair"))?;
        let paragraph_style = iwork_inherited_style_at(&storage.paragraph_styles, start);
        let list_style = iwork_inherited_style_at(&storage.list_styles, start);
        let character_style = pages_style_at(&storage.character_styles, start);
        let mut style = paragraph_style
            .map(|identifier| keynote_text_style(archives, identifier, part))
            .transpose()?
            .unwrap_or_else(|| fallback_style.clone());
        style.list_level = iwork_paragraph_list_level_at(&storage.paragraph_data, start);
        if let Some(identifier) = character_style {
            keynote_apply_text_style(archives, identifier, part, &mut style, &mut Vec::new())?;
        }
        if let Some((_, _, id)) = caps
            .iter()
            .find(|(cap_start, cap_end, _)| *cap_start <= start && start < *cap_end)
        {
            keynote_apply_text_style(archives, *id, part, &mut style, &mut Vec::new())?;
        }
        let paragraph_start = start == 0
            || storage.text[..start_byte].ends_with('\n')
            || storage.text[..start_byte].ends_with('\u{2028}');
        let text = iwork_text_with_dynamic_fields(
            storage,
            archives,
            part,
            start,
            end,
            limits,
            page_number,
        )?;
        if text.is_empty() {
            continue;
        }
        if paragraph_start
            && !text.starts_with('\n')
            && let Some((identifier, marker)) = list_style
                .map(|identifier| {
                    keynote_list_marker(
                        archives,
                        identifier,
                        style.list_level,
                        part,
                        &mut Vec::new(),
                    )
                    .map(|marker| (identifier, marker))
                })
                .transpose()?
                .and_then(|(identifier, marker)| marker.map(|marker| (identifier, marker)))
        {
            let sequence = list_counters
                .entry((identifier, marker.level))
                .and_modify(|value| *value = value.saturating_add(1))
                .or_insert(1);
            let marker_style = keynote_list_marker_style(&style, &marker);
            if marker.indent > 0.0 {
                let mut indent_run = keynote_text_run(" ".to_owned(), &style);
                indent_run.letter_spacing += marker.indent - style.font_size * 0.278;
                pages_append_text_run(&mut runs, indent_run);
            }
            pages_append_text_run(
                &mut runs,
                keynote_text_run(
                    keynote_list_marker_text(&marker.kind, *sequence),
                    &marker_style,
                ),
            );
            pages_append_text_run(&mut runs, keynote_text_run("\t".to_owned(), &style));
        }
        pages_append_text_run(&mut runs, keynote_text_run(text, &style));
    }
    Ok(keynote_platform_font_runs(runs))
}

fn pages_append_text_run(runs: &mut Vec<TextRun>, run: TextRun) {
    if let Some(previous) = runs.last_mut()
        && previous.same_style(&run)
    {
        previous.text.push_str(&run.text);
    } else {
        runs.push(run);
    }
}

fn pages_style_at(changes: &[KeynoteStyleChange], character_index: usize) -> Option<u64> {
    changes
        .partition_point(|change| change.character_index <= character_index)
        .checked_sub(1)
        .and_then(|index| changes[index].identifier)
}

fn iwork_inherited_style_at(changes: &[KeynoteStyleChange], character_index: usize) -> Option<u64> {
    // A nil paragraph/list entry continues the preceding style in both Pages and Keynote.
    changes[..changes.partition_point(|change| change.character_index <= character_index)]
        .iter()
        .rev()
        .find_map(|change| change.identifier)
}

fn iwork_paragraph_list_level_at(
    changes: &[KeynoteParagraphDataChange],
    character_index: usize,
) -> usize {
    changes
        .partition_point(|change| change.character_index <= character_index)
        .checked_sub(1)
        .and_then(|index| changes.get(index))
        .map_or(0, |change| change.list_level)
}

#[allow(clippy::too_many_arguments)]
fn pages_paragraph_layouts(
    storage: &KeynoteTextStorage,
    runs: &[TextRun],
    archives: &[IwaArchive],
    part: &str,
    range_start: usize,
    range_end: usize,
    messages: &NumbersMessageSpace,
    limits: Limits,
    view_scale: f32,
    font_metrics: &FontMetricTable,
    floating_text: bool,
) -> Result<Vec<TextParagraphLayout>, Diagnostic> {
    let start_byte = keynote_utf16_byte_offset(&storage.text, range_start)
        .ok_or_else(|| iwa_error(part, "Pages page start splits a UTF-16 surrogate pair"))?;
    let end_byte = keynote_utf16_byte_offset(&storage.text, range_end)
        .ok_or_else(|| iwa_error(part, "Pages page end splits a UTF-16 surrogate pair"))?;
    let mut layouts = Vec::new();
    let natural_height = |family: &str, italic: bool, bold: bool, size: f32| {
        let height = size
            * font_metrics
                .line_height_em(family, italic, bold, 1.0)
                .unwrap_or(1.2);
        // Native floating text uses NSFont's point-rounded default line box;
        // word-processing body layout retains fractional CoreText metrics.
        if floating_text {
            font_metrics
                .point_rounded_line_height(family, italic, bold, size)
                .unwrap_or_else(|| height.round().max(1.0))
        } else {
            height
        }
    };
    let mut run_heights = vec![0.0_f32];
    for run in runs {
        for segment in run.text.split_inclusive(['\n', '\u{2028}']) {
            if !segment.is_empty() {
                let height = run_heights.last_mut().unwrap();
                *height = height.max(natural_height(
                    &run.font_family,
                    run.italic,
                    run.bold,
                    run.font_size,
                ));
            }
            if segment.ends_with(['\n', '\u{2028}']) {
                run_heights.push(0.0);
            }
        }
    }
    let mut utf16_position = range_start;
    let mut soft_continuation = storage.text[..start_byte].ends_with('\u{2028}');
    for paragraph in storage.text[start_byte..end_byte].split_inclusive(['\n', '\u{2028}']) {
        let paragraph_end = utf16_position.saturating_add(paragraph.encode_utf16().count());
        let identifier = iwork_inherited_style_at(&storage.paragraph_styles, utf16_position);
        let mut style = identifier
            .map(|identifier| keynote_text_style(archives, identifier, part))
            .transpose()?
            .unwrap_or_else(|| KeynoteTextStyle {
                font_family: "Helvetica".to_owned(),
                font_size: 10.5,
                ..KeynoteTextStyle::default()
            });
        style.list_level = iwork_paragraph_list_level_at(&storage.paragraph_data, utf16_position);
        let natural_line_height = natural_height(
            &style.font_family,
            style.italic,
            style.bold,
            style.font_size,
        )
        .max(run_heights.get(layouts.len()).copied().unwrap_or(0.0));
        // Pages Lines spacing multiplies the font's natural line box, including
        // values below one; it is not an em-size minimum as in DrawingML.
        let mut line_height = style
            .line_height_multiple
            .map_or(natural_line_height, |multiple| {
                if multiple > 4.0 {
                    multiple
                } else {
                    natural_line_height * multiple
                }
            });
        let mut space_after = style.space_after;
        if let Some(attachment) = storage
            .attachments
            .iter()
            .find(|attachment| {
                utf16_position <= attachment.character_index
                    && attachment.character_index < paragraph_end
            })
            .and_then(|attachment| attachment.identifier)
            && let Some(drawable) = pages_attachment_drawable(archives, attachment, part)?
        {
            if let Some(table) = archive_message(archives, drawable, IWORK_TABLE_TYPE, part)? {
                line_height = pages_table_geometry(&table.payload, part)?
                    .height
                    .max(line_height);
                if let Some(model_identifier) = numbers_references(&table.payload, 2, part)?.first()
                    && let Some((model_part, model)) =
                        messages.message(*model_identifier, NUMBERS_TABLE_MODEL_TYPE)?
                {
                    let (rows, _, cells) =
                        numbers_table_cells(messages, model, model_part, limits)?;
                    let mut heights =
                        pages_table_row_heights(messages, model, model_part, rows, line_height)?;
                    let header_rows =
                        numbers_varint(&model.payload, 9, model_part)?.unwrap_or(0) as usize;
                    for height in heights.iter_mut().skip(header_rows) {
                        *height *= view_scale;
                    }
                    if iwork_table_image_row_heights(archives, &cells, &mut heights, limits)? {
                        line_height = line_height.max(heights.iter().sum());
                    }
                }
            } else if let Some(image) = archive_message(archives, drawable, IWORK_IMAGE_TYPE, part)?
            {
                let image_height = keynote_image(&image.payload, part)?.geometry.height;
                if paragraph
                    .trim_end()
                    .strip_suffix('\u{fffc}')
                    .is_some_and(|text| text.chars().any(|c| !c.is_whitespace() && c != '\u{fffc}'))
                {
                    // ponytail: reserve a terminal image once as a block;
                    // contour wrapping needs per-line exclusions, not tall text lines.
                    space_after += image_height;
                } else {
                    line_height = image_height.max(line_height);
                }
            } else if archive_message(archives, drawable, IWORK_GROUP_TYPE, part)?.is_some() {
                // Pages keeps a short anchor line when a grouped floating object is
                // continued on the following saved page.
                line_height = line_height.max(54.0);
            }
        }
        let list_offset = iwork_inherited_style_at(&storage.list_styles, utf16_position)
            .map(|identifier| {
                keynote_list_marker(
                    archives,
                    identifier,
                    style.list_level,
                    part,
                    &mut Vec::new(),
                )
            })
            .transpose()?
            .flatten()
            .map(|marker| style.margin_left + marker.indent + marker.text_indent * style.font_size);
        let (margin_left, first_line_indent) = style.paragraph_indents(list_offset);
        let space_before = super::collapsed_paragraph_space_before(
            style.space_before,
            layouts
                .last()
                .map_or(0.0, |layout: &TextParagraphLayout| layout.space_after),
        );
        layouts.push(TextParagraphLayout {
            align: style.align,
            margin_left,
            margin_right: style.margin_right,
            first_line_indent: if layouts.is_empty()
                && range_start > 0
                && !storage.text[..start_byte].ends_with(['\n', '\u{2028}'])
            {
                0.0
            } else {
                first_line_indent
            },
            default_tab_stop: style.default_tab_stop,
            line_height,
            space_before: space_before
                + if style.border_positions.is_some_and(|v| v & 1 != 0) {
                    style.border_padding()
                } else {
                    0.0
                },
            space_after: space_after
                + if style.border_positions.is_some_and(|v| v & 2 != 0) {
                    style.border_padding()
                } else {
                    0.0
                },
            latin_line_break: true,
            hanging_punctuation: false,
            rule_above: style.rule_above(),
            rule_below: style.horizontal_border(2),
            drop_cap: storage
                .drop_cap_styles
                .iter()
                .find(|change| change.character_index == utf16_position)
                .and_then(|change| change.identifier)
                .and_then(|id| pages_drop_cap(archives, id, part).ok()),
        });
        // A normalized soft break needs a line-layout entry, not another
        // paragraph's indentation, spacing, or borders.
        let layout = layouts.last_mut().unwrap();
        if soft_continuation {
            layout.first_line_indent = 0.0;
            layout.space_before = 0.0;
            layout.rule_above = None;
        }
        soft_continuation = paragraph.ends_with('\u{2028}');
        if soft_continuation {
            layout.space_after = 0.0;
            layout.rule_below = None;
        }
        utf16_position = paragraph_end;
    }
    if layouts.is_empty() {
        layouts.push(keynote_text_paragraph_layout(
            TextAlign::Start,
            12.6,
            0.0,
            None,
        ));
    }
    Ok(layouts)
}

fn keynote_platform_font_runs(runs: Vec<TextRun>) -> Vec<TextRun> {
    let mut result: Vec<TextRun> = Vec::new();
    for run in runs {
        let needs_platform_fallback = matches!(
            run.font_family.as_str(),
            "ArialMT" | "Calibri" | "Courier" | "PingFangSC-Regular"
        ) || run.font_family.starts_with("Helvetica");
        if !needs_platform_fallback {
            result.push(run);
            continue;
        }
        for character in run.text.chars() {
            let mut segment = run.clone();
            segment.text = character.to_string();
            if keynote_cjk_character(character) {
                segment.font_family = if run.bold {
                    segment.bold = false;
                    "PingFangSC-Semibold"
                } else {
                    "PingFangSC-Regular"
                }
                .to_owned();
            } else if matches!(run.font_family.as_str(), "Calibri" | "PingFangSC-Regular")
                && character != '\n'
            {
                segment.font_family = "Helvetica".to_owned();
            }
            if let Some(previous) = result.last_mut()
                && previous.same_style(&segment)
            {
                previous.text.push(character);
            } else {
                result.push(segment);
            }
        }
    }
    result
}

fn keynote_cjk_character(character: char) -> bool {
    matches!(
        character as u32,
        0x2e80..=0x9fff | 0xf900..=0xfaff | 0xff00..=0xffef | 0x20000..=0x3134f
    )
}

fn keynote_leading_space_indent(runs: &mut [TextRun]) -> f32 {
    const HELVETICA_SPACE_ADVANCE: f32 = 0.278;
    let Some(run) = runs.first_mut() else {
        return 0.0;
    };
    let space_count = run
        .text
        .chars()
        .take_while(|character| *character == ' ')
        .count();
    if space_count == 0 || run.font_family != "Helvetica" {
        return 0.0;
    }
    let space_advance = run.font_size * HELVETICA_SPACE_ADVANCE;
    run.letter_spacing -= space_advance;
    space_count as f32 * space_advance
}

fn update_keynote_style_at(
    changes: &[KeynoteStyleChange],
    character_index: usize,
    current: &mut Option<u64>,
) {
    if let Ok(index) =
        changes.binary_search_by_key(&character_index, |change| change.character_index)
    {
        *current = changes[index].identifier;
    }
}

fn keynote_utf16_byte_offset(text: &str, target: usize) -> Option<usize> {
    let mut utf16_offset = 0_usize;
    for (byte_offset, character) in text.char_indices() {
        if utf16_offset == target {
            return Some(byte_offset);
        }
        utf16_offset = utf16_offset.checked_add(character.len_utf16())?;
        if utf16_offset > target {
            return None;
        }
    }
    (utf16_offset == target).then_some(text.len())
}

fn normalize_keynote_text(text: &str) -> String {
    text.replace('\u{2028}', "\n").replace('\u{fffc}', "")
}

fn iwork_visible_table_text(text: &str) -> String {
    text.chars()
        .filter(|character| {
            *character != '\u{fffc}'
                && (matches!(character, '\n' | '\t') || !character.is_control())
        })
        .collect()
}

fn keynote_semantic_text(runs: &[TextRun]) -> String {
    runs.iter()
        .map(|run| {
            super::normalize_symbol_font_character(&run.text, Some(run.font_family.as_str()))
        })
        .collect()
}

fn keynote_text_run(text: String, style: &KeynoteTextStyle) -> TextRun {
    TextRun {
        paint: None,
        east_asian_line_breaks: true,
        text: if style.all_caps {
            text.to_uppercase()
        } else {
            text
        },
        font_family: style.font_family.clone(),
        font_size: style.font_size,
        color: style.color,
        bold: style.bold,
        italic: style.italic,
        underline: style.underline,
        strikethrough: style.strikethrough,
        highlight: style.highlight,
        baseline_shift: style.baseline_shift,
        letter_spacing: style.tracking * style.font_size,
        horizontal_scale: 1.0,
    }
}

fn keynote_list_marker_style(
    style: &KeynoteTextStyle,
    marker: &KeynoteListMarker,
) -> KeynoteTextStyle {
    let mut marker_style = style.clone();
    marker_style.font_family = marker
        .font_family
        .clone()
        .unwrap_or_else(|| "ArialMT".to_owned());
    marker_style.font_size *= marker.scale;
    if let Some(color) = marker.color {
        marker_style.color = color;
    }
    marker_style.bold = false;
    marker_style.underline = false;
    marker_style.strikethrough = false;
    marker_style
}

fn keynote_text_style(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
) -> Result<KeynoteTextStyle, Diagnostic> {
    let mut style = KeynoteTextStyle::default();
    keynote_apply_text_style(archives, identifier, part, &mut style, &mut Vec::new())?;
    Ok(style)
}

fn keynote_list_marker(
    archives: &[IwaArchive],
    identifier: u64,
    level: usize,
    part: &str,
    path: &mut Vec<u64>,
) -> Result<Option<KeynoteListMarker>, Diagnostic> {
    if path.len() >= 64 || path.contains(&identifier) {
        return Ok(None);
    }
    let Some(style) = archive_message(archives, identifier, IWORK_LIST_STYLE_TYPE, part)? else {
        return Ok(None);
    };
    path.push(identifier);
    let parent = keynote_nested_message(&style.payload, 1, part)?
        .map(|base| keynote_style_parent(base, part))
        .transpose()?
        .flatten();
    let mut cursor = 0_usize;
    let mut label_types = Vec::new();
    let mut number_types = Vec::new();
    let mut labels = Vec::new();
    let mut text_indents = Vec::new();
    let mut indents = Vec::new();
    let mut scales = Vec::new();
    let mut color = None;
    let mut font_family = None;
    while cursor < style.payload.len() {
        let (key, consumed) = read_varint(&style.payload[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote list-style key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (11, 0) => {
                let (value, consumed) =
                    read_varint(&style.payload[cursor..]).map_err(|message| {
                        iwa_error(part, format!("invalid Keynote list label type: {message}"))
                    })?;
                cursor += consumed;
                label_types.push(value);
            }
            (12, 5) => {
                text_indents.push(keynote_fixed32(
                    &style.payload,
                    &mut cursor,
                    part,
                    "Keynote list text indent",
                )?);
            }
            (13, 5) => {
                indents.push(keynote_fixed32(
                    &style.payload,
                    &mut cursor,
                    part,
                    "Keynote list indent",
                )?);
            }
            (14, 2) => {
                let geometry = read_iwa_length_delimited(
                    &style.payload,
                    &mut cursor,
                    part,
                    "Keynote list label geometry",
                )?;
                scales.push(keynote_fixed32_field(geometry, 1, part)?.unwrap_or(1.0));
            }
            (15, 0) => {
                let (value, consumed) =
                    read_varint(&style.payload[cursor..]).map_err(|message| {
                        iwa_error(part, format!("invalid Keynote list number type: {message}"))
                    })?;
                cursor += consumed;
                number_types.push(value);
            }
            (16, 2) => {
                let encoded = read_iwa_length_delimited(
                    &style.payload,
                    &mut cursor,
                    part,
                    "Keynote list label",
                )?;
                if let Ok(value) = std::str::from_utf8(encoded) {
                    labels.push(value.to_owned());
                }
            }
            (21, 2) => {
                let encoded = read_iwa_length_delimited(
                    &style.payload,
                    &mut cursor,
                    part,
                    "Keynote list font color",
                )?;
                color = Some(keynote_color(encoded, part)?);
            }
            (23, 2) => {
                let encoded = read_iwa_length_delimited(
                    &style.payload,
                    &mut cursor,
                    part,
                    "Keynote list font family",
                )?;
                font_family = std::str::from_utf8(encoded)
                    .ok()
                    .filter(|value| !value.is_empty())
                    .map(|value| keynote_font_family(value).to_owned());
            }
            (_, wire) => skip_protobuf_value(&style.payload, &mut cursor, wire, part)?,
        }
    }
    let marker = match label_types.get(level).copied() {
        Some(1 | 2) => Some(KeynoteListMarker {
            kind: KeynoteListMarkerKind::Text(
                labels
                    .get(level)
                    .filter(|value| !value.is_empty())
                    .cloned()
                    .unwrap_or_else(|| "•".to_owned()),
            ),
            level,
            font_family,
            scale: scales.get(level).copied().unwrap_or(1.0),
            color,
            indent: indents.get(level).copied().unwrap_or(level as f32 * 36.0),
            text_indent: text_indents.get(level).copied().unwrap_or(1.0),
        }),
        Some(3) => Some(KeynoteListMarker {
            kind: KeynoteListMarkerKind::Number(number_types.get(level).copied().unwrap_or(0)),
            level,
            font_family,
            scale: scales.get(level).copied().unwrap_or(1.0),
            color,
            indent: indents.get(level).copied().unwrap_or(level as f32 * 36.0),
            text_indent: text_indents.get(level).copied().unwrap_or(1.0),
        }),
        Some(0) => None,
        _ => {
            let mut marker = parent
                .map(|parent| keynote_list_marker(archives, parent, level, part, path))
                .transpose()?
                .flatten();
            if let Some(marker) = marker.as_mut() {
                if let Some(label) = labels.get(level).filter(|value| !value.is_empty())
                    && let KeynoteListMarkerKind::Text(text) = &mut marker.kind
                {
                    text.clone_from(label);
                }
                if let Some(indent) = text_indents.get(level) {
                    marker.text_indent = *indent;
                }
                if let Some(indent) = indents.get(level) {
                    marker.indent = *indent;
                }
                if font_family.is_some() {
                    marker.font_family.clone_from(&font_family);
                }
                if let Some(scale) = scales.get(level) {
                    marker.scale = *scale;
                }
                if color.is_some() {
                    marker.color = color;
                }
            }
            marker
        }
    };
    path.pop();
    Ok(marker)
}

fn keynote_list_marker_text(kind: &KeynoteListMarkerKind, sequence: usize) -> String {
    let KeynoteListMarkerKind::Number(number_type) = kind else {
        return match kind {
            KeynoteListMarkerKind::Text(text) => text.clone(),
            KeynoteListMarkerKind::Number(_) => unreachable!(),
        };
    };
    let decimal = sequence.to_string();
    let roman_upper = keynote_roman_numeral(sequence);
    let roman_lower = roman_upper.to_ascii_lowercase();
    let alpha_upper = keynote_alpha_numeral(sequence);
    let alpha_lower = alpha_upper.to_ascii_lowercase();
    match *number_type {
        0 => format!("{decimal}."),
        1 => format!("({decimal})"),
        2 => format!("{decimal})"),
        3 => format!("{roman_upper}."),
        4 => format!("({roman_upper})"),
        5 => format!("{roman_upper})"),
        6 => format!("{roman_lower}."),
        7 => format!("({roman_lower})"),
        8 => format!("{roman_lower})"),
        9 => format!("{alpha_upper}."),
        10 => format!("({alpha_upper})"),
        11 => format!("{alpha_upper})"),
        12 => format!("{alpha_lower}."),
        13 => format!("({alpha_lower})"),
        14 => format!("{alpha_lower})"),
        48 => char::from_u32(0x2460 + u32::try_from(sequence.saturating_sub(1)).unwrap_or(0))
            .filter(|_| sequence <= 20)
            .map_or(decimal, |character| character.to_string()),
        _ => format!("{decimal}."),
    }
}

fn keynote_roman_numeral(mut value: usize) -> String {
    if value == 0 || value > 3_999 {
        return value.to_string();
    }
    let mut result = String::new();
    for (amount, digits) in [
        (1_000, "M"),
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
        while value >= amount {
            result.push_str(digits);
            value -= amount;
        }
    }
    result
}

fn keynote_alpha_numeral(mut value: usize) -> String {
    if value == 0 {
        return value.to_string();
    }
    let mut result = Vec::new();
    while value > 0 {
        value -= 1;
        result.push((b'A' + u8::try_from(value % 26).unwrap_or(0)) as char);
        value /= 26;
    }
    result.into_iter().rev().collect()
}

fn keynote_apply_text_style(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
    result: &mut KeynoteTextStyle,
    path: &mut Vec<u64>,
) -> Result<(), Diagnostic> {
    if path.len() >= 64 || path.contains(&identifier) {
        return Ok(());
    }
    let paragraph_style = archive_message(archives, identifier, IWORK_PARAGRAPH_STYLE_TYPE, part)?;
    let toc_style = archive_message(archives, identifier, IWORK_TOC_PARAGRAPH_STYLE_TYPE, part)?;
    if paragraph_style.is_some() && toc_style.is_some() {
        return Err(iwa_error(
            part,
            "iWork style resolves to multiple paragraph style kinds",
        ));
    }
    let paragraph_style = paragraph_style.or(toc_style);
    let character_style =
        archive_message(archives, identifier, IWORK_CHARACTER_STYLE_TYPE, part)?.or(
            archive_message(archives, identifier, PAGES_DROP_CAP_STYLE_TYPE, part)?,
        );
    let style = match (paragraph_style, character_style) {
        (Some(_), Some(_)) => {
            return Err(iwa_error(
                part,
                "Keynote style identifier resolves to multiple style kinds",
            ));
        }
        (Some(style), None) | (None, Some(style)) => style,
        (None, None) => return Ok(()),
    };
    let payload = if style.message_type == IWORK_TOC_PARAGRAPH_STYLE_TYPE {
        numbers_first_bytes(&style.payload, 1, part)?
            .ok_or_else(|| iwa_error(part, "iWork TOC style has no paragraph style"))?
    } else {
        &style.payload
    };
    path.push(identifier);
    let mut cursor = 0_usize;
    let mut parent = None;
    let mut character = None;
    let mut paragraph = None;
    while cursor < payload.len() {
        let (key, consumed) = read_varint(&payload[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote paragraph-style key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let super_style =
                    read_iwa_length_delimited(payload, &mut cursor, part, "Keynote base style")?;
                parent = keynote_style_parent(super_style, part)?;
            }
            (11, 2) => {
                character = Some(read_iwa_length_delimited(
                    payload,
                    &mut cursor,
                    part,
                    "Keynote character style",
                )?);
            }
            (12, 2) if style.message_type != PAGES_DROP_CAP_STYLE_TYPE => {
                paragraph = Some(read_iwa_length_delimited(
                    payload,
                    &mut cursor,
                    part,
                    "Keynote paragraph properties",
                )?);
            }
            (_, wire) => skip_protobuf_value(payload, &mut cursor, wire, part)?,
        }
    }
    if let Some(parent) = parent {
        keynote_apply_text_style(archives, parent, part, result, path)?;
    }
    if let Some(character) = character {
        keynote_character_style(character, part, result)?;
    }
    if let Some(paragraph) = paragraph {
        keynote_paragraph_properties(paragraph, part, result)?;
    }
    path.pop();
    Ok(())
}

fn keynote_style_parent(bytes: &[u8], part: &str) -> Result<Option<u64>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote base-style key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (3, 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote parent-style reference",
                )?;
                return Ok(Some(parse_iwa_reference(reference)?));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(None)
}

fn keynote_character_style(
    bytes: &[u8],
    part: &str,
    style: &mut KeynoteTextStyle,
) -> Result<(), Diagnostic> {
    let mut cursor = 0_usize;
    let mut clear_bold = false;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote character-style key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (field @ (1 | 2), 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote text-style flag: {message}"))
                })?;
                cursor += consumed;
                if field == 1 {
                    style.bold = value != 0;
                    clear_bold = value == 0;
                } else {
                    style.italic = value != 0;
                }
            }
            (3, 5) => {
                style.font_size =
                    keynote_fixed32(bytes, &mut cursor, part, "font size")?.clamp(1.0, 1_024.0);
            }
            (5, 2) => {
                let encoded =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote font family")?;
                let value = std::str::from_utf8(encoded)
                    .map_err(|_| iwa_error(part, "Keynote font family is not UTF-8"))?;
                if !value.is_empty() {
                    style.font_family = keynote_font_family(value).to_owned();
                }
            }
            (7, 2) => {
                let color =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote text color")?;
                style.color = keynote_color(color, part)?;
            }
            (field @ 10..=13, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote text decoration: {message}"))
                })?;
                cursor += consumed;
                match field {
                    10 => {
                        style.baseline_shift = match value {
                            1 => style.font_size * 0.33,
                            2 => style.font_size * -0.2,
                            _ => 0.0,
                        };
                    }
                    11 => style.underline = value != 0,
                    12 => style.strikethrough = value != 0,
                    13 => style.all_caps = value == 1,
                    _ => unreachable!(),
                }
            }
            (field @ (14 | 27), 5) => {
                let value = keynote_fixed32(
                    bytes,
                    &mut cursor,
                    part,
                    if field == 14 {
                        "baseline shift"
                    } else {
                        "tracking"
                    },
                )?
                .clamp(-1_024.0, 1_024.0);
                if field == 14 {
                    style.baseline_shift = value;
                } else {
                    style.tracking = value;
                }
            }
            (25, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(
                        part,
                        format!("invalid Keynote background-color flag: {message}"),
                    )
                })?;
                cursor += consumed;
                if value != 0 {
                    style.highlight = 0;
                }
            }
            (26, 2) => {
                let color = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote text background color",
                )?;
                style.highlight = keynote_color(color, part)?;
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    // An explicit trait override wins over the PostScript face's stale weight.
    // Absent bold flags must still preserve an authored -Bold face.
    if clear_bold && let Some(family) = style.font_family.strip_suffix("-Bold") {
        style.font_family = family.to_owned();
    }
    Ok(())
}

fn keynote_font_family(value: &str) -> &str {
    match value {
        "KaiTi" => "PingFangSC-Regular",
        "Wingdings-Regular" => "Wingdings",
        "SymbolMT" => "Symbol",
        "Superclarendon-Regular" => "Superclarendon",
        _ => value,
    }
}

fn keynote_paragraph_properties(
    bytes: &[u8],
    part: &str,
    style: &mut KeynoteTextStyle,
) -> Result<(), Diagnostic> {
    let mut cursor = 0_usize;
    let mut border_positions = None;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote paragraph-properties key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (45, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(
                        part,
                        format!("invalid paragraph border positions: {message}"),
                    )
                })?;
                cursor += consumed;
                border_positions = Some(value);
            }
            (field @ (9 | 26), 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(
                        part,
                        format!("invalid paragraph pagination flag: {message}"),
                    )
                })?;
                cursor += consumed;
                if field == 9 {
                    style.keep_lines_together = value != 0;
                } else {
                    style.widow_control = value != 0;
                }
            }
            (15, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid paragraph border: {message}"))
                })?;
                cursor += consumed;
                style.rule_enabled = value == 1;
                style.border_positions = None;
            }
            (17, 2) => {
                let value =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "paragraph rule offset")?;
                style.rule_offset = iwork_float_pair(value, part)?;
            }
            (18, 5) => {
                style.rule_width =
                    keynote_fixed32(bytes, &mut cursor, part, "paragraph rule width")?;
            }
            (32, 2) => {
                let value =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "paragraph stroke")?;
                style.rule_stroke = Some(keynote_stroke(value, part)?);
            }
            (25, 2) => {
                let list = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "iWork paragraph tab stops",
                )?;
                style.tab_stops.clear();
                for tab in numbers_bytes(list, 1, part)? {
                    let Some(position) = keynote_fixed32_field(tab, 1, part)? else {
                        continue;
                    };
                    let align = match numbers_varint(tab, 2, part)?.unwrap_or(0) {
                        1 => TextAlign::Center,
                        2 => TextAlign::End,
                        _ => TextAlign::Start,
                    };
                    let leader = match numbers_first_bytes(tab, 3, part)? {
                        Some(b".") => TextTabLeader::Dot,
                        Some(b"-") => TextTabLeader::Hyphen,
                        Some(b"_") => TextTabLeader::Underscore,
                        _ => TextTabLeader::None,
                    };
                    style.tab_stops.push(TextTabStop {
                        position,
                        align,
                        leader,
                    });
                }
            }
            (1, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote text alignment: {message}"))
                })?;
                cursor += consumed;
                style.align = match value {
                    1 => TextAlign::End,
                    2 => TextAlign::Center,
                    3 => TextAlign::Justify,
                    _ => TextAlign::Start,
                };
                style.natural_alignment = value == 4;
            }
            (13, 2) => {
                let line_spacing = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote paragraph line spacing",
                )?;
                if let Some(value) = keynote_fixed32_field(line_spacing, 2, part)? {
                    let relative = keynote_varint_field(line_spacing, 1, part)?.unwrap_or(0) == 0;
                    style.line_height_multiple =
                        Some(value.clamp(0.1, if relative { 10.0 } else { 1_024.0 }));
                }
            }
            (field @ (4 | 7 | 11 | 19 | 20 | 21), 5) => {
                let value = keynote_fixed32(bytes, &mut cursor, part, "Keynote paragraph metric")?
                    .clamp(-10_000.0, 10_000.0);
                match field {
                    4 => style.default_tab_stop = value.max(1.0),
                    7 => style.first_line_indent = value,
                    11 => style.margin_left = value,
                    19 => style.margin_right = value,
                    20 => style.space_after = value.max(0.0),
                    21 => style.space_before = value.max(0.0),
                    _ => unreachable!(),
                }
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    if let Some(positions) = border_positions {
        style.border_positions = Some(positions);
    }
    Ok(())
}

fn keynote_drawable_style(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
    width: f32,
    height: f32,
) -> Result<KeynoteDrawableStyle, Diagnostic> {
    let mut style = KeynoteDrawableStyle::default();
    keynote_apply_drawable_style(
        archives,
        identifier,
        part,
        width,
        height,
        &mut style,
        &mut Vec::new(),
    )?;
    Ok(style)
}

fn keynote_text_layout_style(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
) -> Result<Option<TextLayout>, Diagnostic> {
    let mut layout = TextLayout::default();
    keynote_apply_text_layout_style(archives, identifier, part, &mut layout, &mut Vec::new())
        .map(|found| found.then_some(layout))
}

fn keynote_apply_text_layout_style(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
    result: &mut TextLayout,
    path: &mut Vec<u64>,
) -> Result<bool, Diagnostic> {
    if path.len() >= 64 || path.contains(&identifier) {
        return Ok(false);
    }
    let Some(style) = archive_message(archives, identifier, IWORK_TEXT_SHAPE_STYLE_TYPE, part)?
    else {
        return Ok(false);
    };
    path.push(identifier);
    let tsd_style = keynote_nested_message(&style.payload, 1, part)?;
    let parent = tsd_style
        .map(|tsd_style| keynote_nested_message(tsd_style, 1, part))
        .transpose()?
        .flatten()
        .map(|base| keynote_style_parent(base, part))
        .transpose()?
        .flatten();
    if let Some(parent) = parent {
        keynote_apply_text_layout_style(archives, parent, part, result, path)?;
    }
    if let Some(properties) = keynote_nested_message(&style.payload, 11, part)? {
        keynote_text_layout_properties(properties, part, result)?;
    }
    path.pop();
    Ok(true)
}

fn keynote_text_layout_properties(
    bytes: &[u8],
    part: &str,
    layout: &mut TextLayout,
) -> Result<(), Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote text-layout key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (field @ (1 | 2 | 5 | 8 | 11), 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(
                        part,
                        format!("invalid Keynote text-layout value: {message}"),
                    )
                })?;
                cursor += consumed;
                match field {
                    1 => {
                        layout.auto_fit = if value == 0 {
                            TextAutoFit::None
                        } else {
                            TextAutoFit::Shrink
                        };
                    }
                    2 => {
                        layout.vertical_align = match value {
                            1 | 3 => TextVerticalAlign::Center,
                            2 => TextVerticalAlign::Bottom,
                            _ => TextVerticalAlign::Top,
                        };
                    }
                    5 if value != 0 => {
                        layout.inset_left = 4.0;
                        layout.inset_top = 4.0;
                        layout.inset_right = 4.0;
                        layout.inset_bottom = 4.0;
                    }
                    8 | 11 if value != 0 => {
                        layout.orientation = TextOrientation::VerticalRl;
                    }
                    _ => {}
                }
            }
            (4, 2) => {
                let columns =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote text columns")?;
                keynote_text_columns(columns, part, layout)?;
            }
            (6, 2) => {
                let padding =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote text padding")?;
                keynote_text_padding(padding, part, layout)?;
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(())
}

fn keynote_text_columns(
    bytes: &[u8],
    part: &str,
    layout: &mut TextLayout,
) -> Result<(), Diagnostic> {
    let Some(equal_columns) = keynote_nested_message(bytes, 1, part)? else {
        return Ok(());
    };
    if let Some(count) = keynote_varint_field(equal_columns, 1, part)? {
        layout.column_count = u32::try_from(count).unwrap_or(64).clamp(1, 64);
    }
    if let Some(gap) = keynote_fixed32_field(equal_columns, 2, part)? {
        layout.column_spacing = gap.max(0.0);
    }
    Ok(())
}

fn keynote_text_padding(
    bytes: &[u8],
    part: &str,
    layout: &mut TextLayout,
) -> Result<(), Diagnostic> {
    layout.inset_left = 0.0;
    layout.inset_top = 0.0;
    layout.inset_right = 0.0;
    layout.inset_bottom = 0.0;
    for (field, target) in [
        (1, &mut layout.inset_left),
        (2, &mut layout.inset_top),
        (3, &mut layout.inset_right),
        (4, &mut layout.inset_bottom),
    ] {
        if let Some(value) = keynote_fixed32_field(bytes, field, part)? {
            *target = value.max(0.0);
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn keynote_apply_drawable_style(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
    width: f32,
    height: f32,
    result: &mut KeynoteDrawableStyle,
    path: &mut Vec<u64>,
) -> Result<(), Diagnostic> {
    if path.len() >= 64 || path.contains(&identifier) {
        return Ok(());
    }
    let text_style = archive_message(archives, identifier, IWORK_TEXT_SHAPE_STYLE_TYPE, part)?;
    let shape_style = archive_message(archives, identifier, IWORK_SHAPE_STYLE_TYPE, part)?;
    let payload = match (text_style, shape_style) {
        (Some(_), Some(_)) => {
            return Err(iwa_error(
                part,
                "Keynote drawable style resolves to multiple style kinds",
            ));
        }
        (Some(style), None) => keynote_nested_message(&style.payload, 1, part)?,
        (None, Some(style)) => Some(style.payload.as_slice()),
        (None, None) => None,
    };
    let Some(payload) = payload else {
        return Ok(());
    };
    path.push(identifier);
    keynote_apply_tsd_shape_style(archives, payload, part, width, height, result, path)?;
    path.pop();
    Ok(())
}

fn keynote_apply_tsd_shape_style(
    archives: &[IwaArchive],
    bytes: &[u8],
    part: &str,
    width: f32,
    height: f32,
    result: &mut KeynoteDrawableStyle,
    path: &mut Vec<u64>,
) -> Result<(), Diagnostic> {
    let mut cursor = 0_usize;
    let mut parent = None;
    let mut properties = None;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote drawable-style key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let base = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote drawable base style",
                )?;
                parent = keynote_style_parent(base, part)?;
            }
            (11, 2) => {
                properties = Some(read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote drawable-style properties",
                )?);
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    if let Some(parent) = parent {
        keynote_apply_drawable_style(archives, parent, part, width, height, result, path)?;
    }
    if let Some(properties) = properties {
        keynote_shape_style_properties(properties, part, width, height, result)?;
    }
    Ok(())
}

fn keynote_media_style(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
) -> Result<KeynoteDrawableStyle, Diagnostic> {
    let mut result = KeynoteDrawableStyle::default();
    keynote_apply_media_style(archives, identifier, part, &mut result, &mut Vec::new())?;
    Ok(result)
}

fn keynote_apply_media_style(
    archives: &[IwaArchive],
    identifier: u64,
    part: &str,
    result: &mut KeynoteDrawableStyle,
    path: &mut Vec<u64>,
) -> Result<(), Diagnostic> {
    if path.len() >= 64 || path.contains(&identifier) {
        return Ok(());
    }
    let Some(style) = archive_message(archives, identifier, IWORK_MEDIA_STYLE_TYPE, part)? else {
        return Ok(());
    };
    path.push(identifier);
    let mut cursor = 0_usize;
    let mut parent = None;
    let mut properties = None;
    while cursor < style.payload.len() {
        let (key, consumed) = read_varint(&style.payload[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote media-style key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let base = read_iwa_length_delimited(
                    &style.payload,
                    &mut cursor,
                    part,
                    "Keynote media base style",
                )?;
                parent = keynote_style_parent(base, part)?;
            }
            (11, 2) => {
                properties = Some(read_iwa_length_delimited(
                    &style.payload,
                    &mut cursor,
                    part,
                    "Keynote media-style properties",
                )?);
            }
            (_, wire) => skip_protobuf_value(&style.payload, &mut cursor, wire, part)?,
        }
    }
    if let Some(parent) = parent {
        keynote_apply_media_style(archives, parent, part, result, path)?;
    }
    if let Some(properties) = properties {
        keynote_media_style_properties(properties, part, result)?;
    }
    path.pop();
    Ok(())
}

fn keynote_nested_message<'a>(
    bytes: &'a [u8],
    field: u64,
    part: &str,
) -> Result<Option<&'a [u8]>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote nested-message key: {message}"),
            )
        })?;
        cursor += consumed;
        if key >> 3 == field && key & 7 == 2 {
            return read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote nested message")
                .map(Some);
        }
        skip_protobuf_value(bytes, &mut cursor, key & 7, part)?;
    }
    Ok(None)
}

fn keynote_shape_style_properties(
    bytes: &[u8],
    part: &str,
    width: f32,
    height: f32,
    style: &mut KeynoteDrawableStyle,
) -> Result<(), Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote shape properties key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let fill =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote shape fill")?;
                style.image_fill = keynote_fill_image_properties(fill, part)?;
                style.fill = keynote_fill(fill, part, width, height)?;
            }
            (2, 2) => {
                let stroke =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote shape stroke")?;
                let (paint, width, stroke_style) = keynote_stroke(stroke, part)?;
                style.stroke = paint;
                style.stroke_width = width;
                style.stroke_style = stroke_style;
            }
            (3, 5) => {
                style.opacity =
                    keynote_fixed32(bytes, &mut cursor, part, "shape opacity")?.clamp(0.0, 1.0);
            }
            (4, 2) => {
                let shadow =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote shape shadow")?;
                style.shadow = keynote_shadow(shadow, part)?;
            }
            (5, 2) => {
                let reflection = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote shape reflection",
                )?;
                style.reflection = keynote_reflection(reflection, part)?;
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(())
}

fn keynote_fill_image_properties(
    bytes: &[u8],
    part: &str,
) -> Result<Option<(u64, bool, Option<(f32, f32)>)>, Diagnostic> {
    let Some(image_fill) = keynote_nested_message(bytes, 3, part)? else {
        return Ok(None);
    };
    let mut identifier = None;
    for field in [6, 1, 7, 5] {
        if let Some(reference) = keynote_reference_field(image_fill, field, part)? {
            identifier = Some(reference);
            break;
        }
    }
    let Some(identifier) = identifier else {
        return Ok(None);
    };
    let tile = keynote_varint_field(image_fill, 2, part)? == Some(2);
    let tile_size = if tile {
        if let Some(size) = keynote_nested_message(image_fill, 4, part)? {
            match (
                keynote_fixed32_field(size, 1, part)?,
                keynote_fixed32_field(size, 2, part)?,
            ) {
                (Some(width), Some(height))
                    if width.is_finite() && height.is_finite() && width > 0.0 && height > 0.0 =>
                {
                    Some((width, height))
                }
                _ => None,
            }
        } else {
            None
        }
    } else {
        None
    };
    Ok(Some((identifier, tile, tile_size)))
}

fn keynote_media_style_properties(
    bytes: &[u8],
    part: &str,
    style: &mut KeynoteDrawableStyle,
) -> Result<(), Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote media properties key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let stroke =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote media stroke")?;
                let (paint, width, stroke_style) = keynote_stroke(stroke, part)?;
                style.stroke = paint;
                style.stroke_width = width;
                style.stroke_style = stroke_style;
            }
            (2, 5) => {
                style.opacity =
                    keynote_fixed32(bytes, &mut cursor, part, "media opacity")?.clamp(0.0, 1.0);
            }
            (3, 2) => {
                let shadow =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote media shadow")?;
                style.shadow = keynote_shadow(shadow, part)?;
            }
            (4, 2) => {
                let reflection = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote media reflection",
                )?;
                style.reflection = keynote_reflection(reflection, part)?;
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(())
}

fn keynote_fill(bytes: &[u8], part: &str, width: f32, height: f32) -> Result<Paint, Diagnostic> {
    let mut cursor = 0_usize;
    let mut paint = Paint::None;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote fill key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let color =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote fill color")?;
                paint = Paint::Solid(keynote_color(color, part)?);
            }
            (2, 2) => {
                let gradient =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote gradient fill")?;
                paint = keynote_gradient(gradient, part, width, height)?;
            }
            (3, 2) => {
                let image =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "iWork image fill")?;
                // ImageFillArchive.referencecolor is authored independently of tint.
                // The packaged image replaces this approximation when available.
                if let Some(color) = keynote_nested_message(image, 9, part)? {
                    paint = Paint::Solid(keynote_color(color, part)?);
                }
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(paint)
}

fn keynote_gradient(
    bytes: &[u8],
    part: &str,
    width: f32,
    height: f32,
) -> Result<Paint, Diagnostic> {
    let mut cursor = 0_usize;
    let mut kind = 0_u64;
    let mut stops = Vec::new();
    let mut opacity = 1.0_f32;
    let mut angle = 0.0_f32;
    let mut advanced = false;
    let mut transform = None;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote gradient key: {message}"))
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote gradient type: {message}"))
                })?;
                cursor += consumed;
                kind = value;
            }
            (2, 2) => {
                let stop =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote gradient stop")?;
                if let Some(stop) = keynote_gradient_stop(stop, part)? {
                    stops.push(stop);
                }
            }
            (3, 5) => {
                opacity =
                    keynote_fixed32(bytes, &mut cursor, part, "gradient opacity")?.clamp(0.0, 1.0);
            }
            (4, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(
                        part,
                        format!("invalid Keynote advanced-gradient flag: {message}"),
                    )
                })?;
                cursor += consumed;
                advanced = value != 0;
            }
            (5, 2) => {
                let angle_message =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote gradient angle")?;
                angle = keynote_fixed32_field(angle_message, 2, part)?.unwrap_or(0.0);
            }
            (6, 2) => {
                let encoded = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote transformed gradient",
                )?;
                transform = keynote_transform_gradient(encoded, width, height, part)?;
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    stops.sort_by(|left, right| left.offset.total_cmp(&right.offset));
    for stop in &mut stops {
        stop.color = keynote_color_opacity(stop.color, opacity);
    }
    if stops.len() < 2 {
        return Ok(stops
            .first()
            .map_or(Paint::None, |stop| Paint::Solid(stop.color)));
    }
    if let Some((start_x, start_y, end_x, end_y)) = transform {
        if kind == 1 {
            let radius = (end_x - start_x).hypot(end_y - start_y);
            return Ok(Paint::RadialGradient {
                x0: start_x,
                y0: start_y,
                r0: 0.0,
                x1: start_x,
                y1: start_y,
                r1: radius.max(f32::EPSILON),
                stops,
            });
        }
        return Ok(Paint::LinearGradient {
            x0: start_x,
            y0: start_y,
            x1: end_x,
            y1: end_y,
            stops,
        });
    }
    if kind == 1 && (advanced || stops.len() == 2) {
        let center_x = if angle.cos() < 0.0 { width } else { 0.0 };
        let center_y = if angle.sin() > 0.0 { height } else { 0.0 };
        Ok(Paint::RadialGradient {
            x0: center_x,
            y0: center_y,
            r0: 0.0,
            x1: center_x,
            y1: center_y,
            r1: (width / 2.0).hypot(height / 2.0),
            stops,
        })
    } else if kind == 1 || kind == 2 {
        let radius = (width / 2.0).hypot(height / 2.0);
        Ok(Paint::RadialGradient {
            x0: width / 2.0,
            y0: height / 2.0,
            r0: 0.0,
            x1: width / 2.0,
            y1: height / 2.0,
            r1: radius,
            stops,
        })
    } else {
        let dx = angle.cos() * width / 2.0;
        let dy = -angle.sin() * height / 2.0;
        Ok(Paint::LinearGradient {
            x0: width / 2.0 - dx,
            y0: height / 2.0 - dy,
            x1: width / 2.0 + dx,
            y1: height / 2.0 + dy,
            stops,
        })
    }
}

fn keynote_transform_gradient(
    bytes: &[u8],
    width: f32,
    height: f32,
    part: &str,
) -> Result<Option<(f32, f32, f32, f32)>, Diagnostic> {
    let start = keynote_nested_message(bytes, 1, part)?
        .map(|point| iwork_float_pair(point, part))
        .transpose()?;
    let end = keynote_nested_message(bytes, 2, part)?
        .map(|point| iwork_float_pair(point, part))
        .transpose()?;
    let natural_size = keynote_nested_message(bytes, 3, part)?
        .map(|size| iwork_float_pair(size, part))
        .transpose()?;
    let (Some((start_x, start_y)), Some((end_x, end_y))) = (start, end) else {
        return Ok(None);
    };
    let (natural_width, natural_height) = natural_size.unwrap_or((width, height));
    let scale_x = if natural_width.abs() > f32::EPSILON {
        width / natural_width
    } else {
        1.0
    };
    let scale_y = if natural_height.abs() > f32::EPSILON {
        height / natural_height
    } else {
        1.0
    };
    Ok(Some((
        start_x * scale_x,
        start_y * scale_y,
        end_x * scale_x,
        end_y * scale_y,
    )))
}

fn keynote_gradient_stop(bytes: &[u8], part: &str) -> Result<Option<GradientStop>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut color = None;
    let mut offset = 0.0_f32;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote gradient-stop key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let encoded = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    part,
                    "Keynote gradient-stop color",
                )?;
                color = Some(keynote_color(encoded, part)?);
            }
            (2, 5) => {
                offset = keynote_fixed32(bytes, &mut cursor, part, "gradient-stop offset")?
                    .clamp(0.0, 1.0);
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    Ok(color.map(|color| GradientStop { offset, color }))
}

fn keynote_stroke(bytes: &[u8], part: &str) -> Result<(Paint, f32, StrokeStyle), Diagnostic> {
    let mut cursor = 0_usize;
    let mut color = None;
    let mut width = 1.0_f32;
    let mut empty = false;
    let mut style = StrokeStyle::default();
    let mut dash_units = Vec::new();
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote stroke key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                let encoded =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote stroke color")?;
                color = Some(keynote_color(encoded, part)?);
            }
            (2, 5) => {
                width =
                    keynote_fixed32(bytes, &mut cursor, part, "stroke width")?.clamp(0.0, 1_024.0);
            }
            (3, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote line cap: {message}"))
                })?;
                cursor += consumed;
                style.cap = match value {
                    1 => LineCap::Round,
                    2 => LineCap::Square,
                    _ => LineCap::Flat,
                };
            }
            (4, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote line join: {message}"))
                })?;
                cursor += consumed;
                style.join = match value {
                    1 => LineJoin::Round,
                    2 => LineJoin::Bevel,
                    _ => LineJoin::Miter,
                };
            }
            (5, 5) => {
                style.miter_limit =
                    keynote_fixed32(bytes, &mut cursor, part, "stroke miter limit")?.max(0.01);
            }
            (6, 2) => {
                let pattern =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote stroke pattern")?;
                empty = keynote_varint_field(pattern, 1, part)? == Some(2);
                let mut pattern_cursor = 0_usize;
                while pattern_cursor < pattern.len() {
                    let (pattern_key, consumed) =
                        read_varint(&pattern[pattern_cursor..]).map_err(|message| {
                            iwa_error(part, format!("invalid Keynote stroke pattern: {message}"))
                        })?;
                    pattern_cursor += consumed;
                    if pattern_key >> 3 == 4 && pattern_key & 7 == 5 {
                        let value =
                            keynote_fixed32(pattern, &mut pattern_cursor, part, "stroke dash")?;
                        if value > f32::EPSILON && dash_units.len() < 256 {
                            dash_units.push(value);
                        }
                    } else {
                        skip_protobuf_value(pattern, &mut pattern_cursor, pattern_key & 7, part)?;
                    }
                }
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    style.dash = dash_units.into_iter().map(|value| value * width).collect();
    Ok(if empty {
        (Paint::None, 0.0, StrokeStyle::default())
    } else {
        (color.map_or(Paint::None, Paint::Solid), width, style)
    })
}

fn keynote_shadow(bytes: &[u8], part: &str) -> Result<Option<Shadow>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut color = 0x0000_00ff;
    let mut angle = 315.0_f32;
    let mut offset = 5.0_f32;
    let mut blur = 1.0_f32;
    let mut opacity = 1.0_f32;
    let mut enabled = true;
    let mut has_properties = false;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote shadow key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 2) => {
                has_properties = true;
                let encoded =
                    read_iwa_length_delimited(bytes, &mut cursor, part, "Keynote shadow color")?;
                color = keynote_color(encoded, part)?;
            }
            (field @ (2 | 3 | 5), 5) => {
                has_properties = true;
                let value = keynote_fixed32(bytes, &mut cursor, part, "shadow property")?;
                match field {
                    2 => angle = value,
                    3 => offset = value,
                    5 => opacity = value.clamp(0.0, 1.0),
                    _ => unreachable!(),
                }
            }
            (4, 0) => {
                has_properties = true;
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote shadow radius: {message}"))
                })?;
                cursor += consumed;
                blur = value.min(1_024) as f32;
            }
            (6, 0) => {
                has_properties = true;
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(part, format!("invalid Keynote shadow flag: {message}"))
                })?;
                cursor += consumed;
                enabled = value != 0;
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    if !enabled || !has_properties {
        return Ok(None);
    }
    let radians = angle.to_radians();
    Ok(Some(Shadow {
        color: keynote_color_opacity(color, opacity),
        blur,
        offset_x: offset * radians.cos(),
        offset_y: -offset * radians.sin(),
    }))
}

fn keynote_reflection(bytes: &[u8], part: &str) -> Result<Option<Reflection>, Diagnostic> {
    Ok(keynote_fixed32_field(bytes, 1, part)?
        .map(|opacity| opacity.clamp(0.0, 1.0))
        .filter(|opacity| *opacity > f32::EPSILON)
        .map(|opacity| Reflection {
            start_opacity: opacity,
            end_opacity: opacity,
            start_position: 0.0,
            end_position: 1.0,
            direction_degrees: 90.0,
            blur: 0.0,
            distance: 0.0,
            scale_x: 1.0,
            scale_y: 1.0,
        }))
}

fn keynote_fixed32_field(bytes: &[u8], field: u64, part: &str) -> Result<Option<f32>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                part,
                format!("invalid Keynote fixed32-field key: {message}"),
            )
        })?;
        cursor += consumed;
        if key >> 3 == field && key & 7 == 5 {
            return keynote_fixed32(bytes, &mut cursor, part, "fixed32 field").map(Some);
        }
        skip_protobuf_value(bytes, &mut cursor, key & 7, part)?;
    }
    Ok(None)
}

fn keynote_varint_field(bytes: &[u8], field: u64, part: &str) -> Result<Option<u64>, Diagnostic> {
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(part, format!("invalid Keynote varint-field key: {message}"))
        })?;
        cursor += consumed;
        if key >> 3 == field && key & 7 == 0 {
            let (value, _) = read_varint(&bytes[cursor..]).map_err(|message| {
                iwa_error(part, format!("invalid Keynote varint field: {message}"))
            })?;
            return Ok(Some(value));
        }
        skip_protobuf_value(bytes, &mut cursor, key & 7, part)?;
    }
    Ok(None)
}

fn keynote_color_opacity(color: u32, opacity: f32) -> u32 {
    let alpha = ((color & 0xff) as f32 * opacity.clamp(0.0, 1.0)).round() as u32;
    (color & 0xffff_ff00) | alpha
}

fn keynote_wrap_style(
    style: &KeynoteDrawableStyle,
    clip: Option<Geometry>,
    mut visual: Visual,
) -> Visual {
    if !style.stroke_style.dash.is_empty()
        || style.stroke_style.cap != LineCap::Flat
        || style.stroke_style.join != LineJoin::Miter
    {
        visual = Visual::StrokeStyle {
            style: style.stroke_style.clone(),
            visual: Box::new(visual),
        };
    }
    if style.reflection.is_some() {
        visual = Visual::AdvancedEffect {
            outer_shadow: None,
            inner_shadow: None,
            glow: None,
            reflection: style.reflection,
            soft_edge: None,
            three_d: None,
            visual: Box::new(visual),
        };
    }
    if style.shadow.is_some() || clip.is_some() {
        visual = Visual::Effect {
            shadow: style.shadow,
            clip,
            visual: Box::new(visual),
        };
    }
    if style.opacity < 1.0 {
        visual = Visual::Layer {
            transform: AffineTransform::IDENTITY,
            opacity: style.opacity,
            blend_mode: crate::model::BlendMode::Normal,
            visual: Box::new(visual),
        };
    }
    visual
}

fn keynote_color(bytes: &[u8], part: &str) -> Result<u32, Diagnostic> {
    let mut cursor = 0_usize;
    let mut red = 0.0_f32;
    let mut green = 0.0_f32;
    let mut blue = 0.0_f32;
    let mut alpha = 1.0_f32;
    let mut white = None;
    while cursor < bytes.len() {
        let (key, consumed) = read_varint(&bytes[cursor..])
            .map_err(|message| iwa_error(part, format!("invalid Keynote color key: {message}")))?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (field @ (3 | 4 | 5 | 6 | 11), 5) => {
                let value = keynote_fixed32(bytes, &mut cursor, part, "color channel")?;
                match field {
                    3 => red = value,
                    4 => green = value,
                    5 => blue = value,
                    6 => alpha = value,
                    11 => white = Some(value),
                    _ => unreachable!(),
                }
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, part)?,
        }
    }
    if let Some(white) = white {
        red = white;
        green = white;
        blue = white;
    }
    let channel = |value: f32| -> u32 { (value.clamp(0.0, 1.0) * 255.0).round() as u32 };
    Ok((channel(red) << 24) | (channel(green) << 16) | (channel(blue) << 8) | channel(alpha))
}

fn keynote_fixed32(
    bytes: &[u8],
    cursor: &mut usize,
    part: &str,
    label: &str,
) -> Result<f32, Diagnostic> {
    let end = cursor
        .checked_add(4)
        .ok_or_else(|| iwa_error(part, format!("Keynote {label} offset overflows")))?;
    let encoded: [u8; 4] = bytes
        .get(*cursor..end)
        .ok_or_else(|| iwa_error(part, format!("Keynote {label} is truncated")))?
        .try_into()
        .map_err(|_| iwa_error(part, format!("Keynote {label} is invalid")))?;
    *cursor = end;
    let value = f32::from_le_bytes(encoded);
    if !value.is_finite() {
        return Err(iwa_error(part, format!("Keynote {label} is not finite")));
    }
    Ok(value)
}

fn archive_message<'a>(
    archives: &'a [IwaArchive],
    identifier: u64,
    message_type: u64,
    part: &str,
) -> Result<Option<&'a IwaMessage>, Diagnostic> {
    let mut found = None;
    for archive in archives
        .iter()
        .filter(|archive| archive.identifier == identifier)
    {
        for message in archive
            .messages
            .iter()
            .filter(|message| message.message_type == message_type)
        {
            if found.replace(message).is_some() {
                return Err(iwa_error(
                    part,
                    "IWA object identifier resolves to multiple matching messages",
                ));
            }
        }
    }
    Ok(found)
}

fn keynote_shape_payload<'a>(
    archives: &'a [IwaArchive],
    identifier: u64,
    part: &str,
) -> Result<Option<&'a [u8]>, Diagnostic> {
    if let Some(shape) = archive_message(archives, identifier, IWORK_SHAPE_TYPE, part)? {
        return Ok(Some(&shape.payload));
    }
    let Some(connection) = archive_message(archives, identifier, IWORK_CONNECTION_LINE_TYPE, part)?
    else {
        return Ok(None);
    };
    keynote_nested_message(&connection.payload, 1, part)
}

fn keynote_document_show_reference(bytes: &[u8]) -> Result<Option<u64>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut show = None;
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                DOCUMENT_COMPONENT,
                "Keynote document contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                DOCUMENT_COMPONENT,
                format!("invalid Keynote document field key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (2, 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    DOCUMENT_COMPONENT,
                    "Keynote show reference",
                )?;
                let identifier = parse_iwa_reference(reference)?;
                if show.replace(identifier).is_some() {
                    return Err(iwa_error(
                        DOCUMENT_COMPONENT,
                        "Keynote document repeats its show reference",
                    ));
                }
            }
            (2, _) => {
                return Err(iwa_error(
                    DOCUMENT_COMPONENT,
                    "Keynote show reference has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, DOCUMENT_COMPONENT)?,
        }
    }
    Ok(show)
}

fn keynote_show(bytes: &[u8]) -> Result<KeynoteShow, Diagnostic> {
    let mut cursor = 0_usize;
    let mut slide_nodes = Vec::new();
    let mut size = None;
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                DOCUMENT_COMPONENT,
                "Keynote show contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                DOCUMENT_COMPONENT,
                format!("invalid Keynote show field key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (3, 2) => {
                let tree = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    DOCUMENT_COMPONENT,
                    "Keynote slide tree",
                )?;
                if !slide_nodes.is_empty() {
                    return Err(iwa_error(
                        DOCUMENT_COMPONENT,
                        "Keynote show repeats its slide tree",
                    ));
                }
                slide_nodes = keynote_slide_tree(tree)?;
            }
            (4, 2) => {
                let encoded = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    DOCUMENT_COMPONENT,
                    "Keynote show size",
                )?;
                if size.replace(iwork_size(encoded)?).is_some() {
                    return Err(iwa_error(
                        DOCUMENT_COMPONENT,
                        "Keynote show repeats its size",
                    ));
                }
            }
            (3 | 4, _) => {
                return Err(iwa_error(
                    DOCUMENT_COMPONENT,
                    "Keynote show field has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, DOCUMENT_COMPONENT)?,
        }
    }
    Ok(KeynoteShow { slide_nodes, size })
}

fn keynote_slide_tree(bytes: &[u8]) -> Result<Vec<u64>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut slides = Vec::new();
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                DOCUMENT_COMPONENT,
                "Keynote slide tree contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                DOCUMENT_COMPONENT,
                format!("invalid Keynote slide tree field key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (2, 2) => {
                let reference = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    DOCUMENT_COMPONENT,
                    "Keynote slide-node reference",
                )?;
                slides.push(parse_iwa_reference(reference)?);
            }
            (2, _) => {
                return Err(iwa_error(
                    DOCUMENT_COMPONENT,
                    "Keynote slide-node reference has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, DOCUMENT_COMPONENT)?,
        }
    }
    Ok(slides)
}

fn iwork_size(bytes: &[u8]) -> Result<(f32, f32), Diagnostic> {
    let mut cursor = 0_usize;
    let mut width = None;
    let mut height = None;
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                DOCUMENT_COMPONENT,
                "iWork size contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                DOCUMENT_COMPONENT,
                format!("invalid iWork size field key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (field @ (1 | 2), 5) => {
                let end = cursor
                    .checked_add(4)
                    .ok_or_else(|| iwa_error(DOCUMENT_COMPONENT, "iWork size offset overflows"))?;
                let encoded: [u8; 4] = bytes
                    .get(cursor..end)
                    .ok_or_else(|| iwa_error(DOCUMENT_COMPONENT, "iWork size is truncated"))?
                    .try_into()
                    .map_err(|_| iwa_error(DOCUMENT_COMPONENT, "iWork size is invalid"))?;
                let value = f32::from_le_bytes(encoded);
                if !value.is_finite() || value <= 0.0 {
                    return Err(iwa_error(DOCUMENT_COMPONENT, "iWork size is invalid"));
                }
                if field == 1 {
                    width = Some(value);
                } else {
                    height = Some(value);
                }
                cursor = end;
            }
            (1 | 2, _) => {
                return Err(iwa_error(
                    DOCUMENT_COMPONENT,
                    "iWork size field has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, DOCUMENT_COMPONENT)?,
        }
    }
    Ok((
        width.ok_or_else(|| iwa_error(DOCUMENT_COMPONENT, "iWork size has no width"))?,
        height.ok_or_else(|| iwa_error(DOCUMENT_COMPONENT, "iWork size has no height"))?,
    ))
}

fn iwork_data_files(
    package: &Package<'_>,
    limits: Limits,
) -> Result<Vec<(u64, String)>, Diagnostic> {
    if !package.has_part(METADATA_COMPONENT) {
        return Ok(Vec::new());
    }
    let bytes = package.required_part(METADATA_COMPONENT)?;
    let archives = parse_iwa_archives(package, METADATA_COMPONENT, &bytes, limits)?;
    let Some(metadata) = archives.iter().find_map(|archive| {
        archive
            .messages
            .iter()
            .find(|message| message.message_type == IWORK_PACKAGE_METADATA_TYPE)
    }) else {
        return Ok(Vec::new());
    };
    parse_iwork_package_metadata(&metadata.payload)
}

fn parse_iwork_package_metadata(bytes: &[u8]) -> Result<Vec<(u64, String)>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut files = Vec::new();
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                METADATA_COMPONENT,
                "iWork package metadata contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                METADATA_COMPONENT,
                format!("invalid iWork package metadata field key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (4, 2) => {
                let data = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    METADATA_COMPONENT,
                    "iWork data metadata",
                )?;
                if let Some((identifier, part)) = parse_iwork_data_info(data)? {
                    if files.iter().any(|(current, _)| *current == identifier) {
                        return Err(iwa_error(
                            METADATA_COMPONENT,
                            "iWork package metadata repeats a data identifier",
                        ));
                    }
                    files.push((identifier, part));
                }
            }
            (4, _) => {
                return Err(iwa_error(
                    METADATA_COMPONENT,
                    "iWork data metadata has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, METADATA_COMPONENT)?,
        }
    }
    Ok(files)
}

fn parse_iwork_data_info(bytes: &[u8]) -> Result<Option<(u64, String)>, Diagnostic> {
    let mut cursor = 0_usize;
    let mut identifier = None;
    let mut file_name = None;
    let mut fields = 0_usize;
    while cursor < bytes.len() {
        fields += 1;
        if fields > MAX_IWA_PROTOBUF_FIELDS {
            return Err(iwa_error(
                METADATA_COMPONENT,
                "iWork data metadata contains too many fields",
            ));
        }
        let (key, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
            iwa_error(
                METADATA_COMPONENT,
                format!("invalid iWork data metadata field key: {message}"),
            )
        })?;
        cursor += consumed;
        match (key >> 3, key & 7) {
            (1, 0) => {
                let (value, consumed) = read_varint(&bytes[cursor..]).map_err(|message| {
                    iwa_error(
                        METADATA_COMPONENT,
                        format!("invalid iWork data identifier: {message}"),
                    )
                })?;
                if value == 0 || identifier.replace(value).is_some() {
                    return Err(iwa_error(
                        METADATA_COMPONENT,
                        "iWork data identifier is zero or repeated",
                    ));
                }
                cursor += consumed;
            }
            (4, 2) => {
                let encoded = read_iwa_length_delimited(
                    bytes,
                    &mut cursor,
                    METADATA_COMPONENT,
                    "iWork data file name",
                )?;
                let value = std::str::from_utf8(encoded).map_err(|_| {
                    iwa_error(METADATA_COMPONENT, "iWork data file name is not UTF-8")
                })?;
                if file_name.replace(value.to_owned()).is_some() {
                    return Err(iwa_error(
                        METADATA_COMPONENT,
                        "iWork data metadata repeats its file name",
                    ));
                }
            }
            (1 | 4, _) => {
                return Err(iwa_error(
                    METADATA_COMPONENT,
                    "iWork data metadata field has an invalid protobuf wire type",
                ));
            }
            (_, wire) => skip_protobuf_value(bytes, &mut cursor, wire, METADATA_COMPONENT)?,
        }
    }
    let Some(identifier) = identifier else {
        return Err(iwa_error(
            METADATA_COMPONENT,
            "iWork data metadata has no identifier",
        ));
    };
    let Some(file_name) = file_name else {
        return Ok(None);
    };
    if file_name.is_empty()
        || file_name.contains(['/', '\\', '\0'])
        || file_name == "."
        || file_name == ".."
    {
        return Ok(None);
    }
    Ok(Some((identifier, format!("Data/{file_name}"))))
}

fn iwork_slide_preview_part(part: &str) -> bool {
    let Some(name) = part.strip_prefix("Data/st-") else {
        return false;
    };
    let name = name.to_ascii_lowercase();
    name.ends_with(".jpg") || name.ends_with(".jpeg") || name.ends_with(".png")
}

fn preview_properties(
    part: &str,
    bytes: &[u8],
    limits: Limits,
) -> Result<(&'static str, f32, f32), Diagnostic> {
    iwork_image_properties(part, bytes, limits)
}

fn iwork_image_properties(
    part: &str,
    bytes: &[u8],
    limits: Limits,
) -> Result<(&'static str, f32, f32), Diagnostic> {
    let lower = part.to_ascii_lowercase();
    if !matches!(
        lower.rsplit_once('.').map(|(_, extension)| extension),
        Some("jpg" | "jpeg" | "png" | "tif" | "tiff")
    ) {
        return Err(preview_error(part, "iWork image format is unsupported"));
    }
    // Keynote can retain the old filename extension after replacing an image.
    // Its native renderer selects the decoder from the bytes, so do the same
    // for the bounded set of image extensions accepted above.
    let media_type = if bytes.starts_with(b"\xff\xd8\xff") {
        "image/jpeg"
    } else if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        "image/png"
    } else if bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*") {
        "image/tiff"
    } else {
        return Err(preview_error(part, "iWork image structure is invalid"));
    };
    let (width, height) = match media_type {
        "image/png" => png_dimensions(bytes, limits.max_zip_entries),
        "image/jpeg" => jpeg_dimensions(bytes, limits.max_zip_entries),
        "image/tiff" => tiff_dimensions(bytes),
        _ => None,
    }
    .ok_or_else(|| preview_error(part, "iWork image structure is invalid"))?;
    let pixels = u64::from(width)
        .checked_mul(u64::from(height))
        .ok_or_else(|| image_limit_error(part, "iWork image pixel count overflow"))?;
    if pixels > limits.max_render_pixels as u64 {
        return Err(image_limit_error(
            part,
            "iWork image exceeds the configured render-pixel limit",
        ));
    }
    Ok((media_type, width as f32, height as f32))
}

fn tiff_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    let little_endian = if bytes.starts_with(b"II*\0") {
        true
    } else if bytes.starts_with(b"MM\0*") {
        false
    } else {
        return None;
    };
    let read_u16 = |offset: usize| {
        let raw: [u8; 2] = bytes.get(offset..offset.checked_add(2)?)?.try_into().ok()?;
        Some(if little_endian {
            u16::from_le_bytes(raw)
        } else {
            u16::from_be_bytes(raw)
        })
    };
    let read_u32 = |offset: usize| {
        let raw: [u8; 4] = bytes.get(offset..offset.checked_add(4)?)?.try_into().ok()?;
        Some(if little_endian {
            u32::from_le_bytes(raw)
        } else {
            u32::from_be_bytes(raw)
        })
    };
    let directory = usize::try_from(read_u32(4)?).ok()?;
    let entries = usize::from(read_u16(directory)?);
    let mut width = None;
    let mut height = None;
    for index in 0..entries {
        let entry = directory
            .checked_add(2)?
            .checked_add(index.checked_mul(12)?)?;
        let tag = read_u16(entry)?;
        if tag != 256 && tag != 257 {
            continue;
        }
        let field_type = read_u16(entry.checked_add(2)?)?;
        let count = read_u32(entry.checked_add(4)?)?;
        if count != 1 {
            return None;
        }
        let value = match field_type {
            3 => u32::from(read_u16(entry.checked_add(8)?)?),
            4 => read_u32(entry.checked_add(8)?)?,
            _ => return None,
        };
        if tag == 256 {
            width = Some(value);
        } else {
            height = Some(value);
        }
    }
    Some((width?, height?))
}

fn png_dimensions(bytes: &[u8], max_chunks: usize) -> Option<(u32, u32)> {
    if !bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        return None;
    }

    let mut cursor = 8_usize;
    let mut chunk_count = 0_usize;
    let mut dimensions = None;
    let mut color_type = 0_u8;
    let mut bit_depth = 0_u8;
    let mut seen_palette = false;
    let mut seen_idat = false;
    let mut idat_ended = false;
    let mut idat_bytes = 0_usize;
    while cursor < bytes.len() {
        chunk_count = chunk_count.checked_add(1)?;
        if chunk_count > max_chunks {
            return None;
        }
        let length_end = cursor.checked_add(4)?;
        let length = usize::try_from(u32::from_be_bytes(
            bytes.get(cursor..length_end)?.try_into().ok()?,
        ))
        .ok()?;
        let kind_end = length_end.checked_add(4)?;
        let kind: &[u8; 4] = bytes.get(length_end..kind_end)?.try_into().ok()?;
        if !kind.iter().all(u8::is_ascii_alphabetic) || !kind[2].is_ascii_uppercase() {
            return None;
        }
        let data_end = kind_end.checked_add(length)?;
        let crc_end = data_end.checked_add(4)?;
        let data = bytes.get(kind_end..data_end)?;
        let expected_crc = u32::from_be_bytes(bytes.get(data_end..crc_end)?.try_into().ok()?);
        if crate::zip::crc32(bytes.get(length_end..data_end)?) != expected_crc {
            return None;
        }

        match kind {
            b"IHDR" => {
                if chunk_count != 1 || length != 13 || dimensions.is_some() {
                    return None;
                }
                let width = u32::from_be_bytes(data.get(0..4)?.try_into().ok()?);
                let height = u32::from_be_bytes(data.get(4..8)?.try_into().ok()?);
                bit_depth = data[8];
                color_type = data[9];
                if width == 0
                    || height == 0
                    || !valid_png_color_depth(color_type, bit_depth)
                    || data[10] != 0
                    || data[11] != 0
                    || data[12] > 1
                {
                    return None;
                }
                dimensions = Some((width, height));
            }
            b"PLTE" => {
                if dimensions.is_none()
                    || seen_palette
                    || seen_idat
                    || matches!(color_type, 0 | 4)
                    || !(3..=768).contains(&length)
                    || length % 3 != 0
                    || (color_type == 3 && length / 3 > (1_usize << usize::from(bit_depth)))
                {
                    return None;
                }
                seen_palette = true;
            }
            b"IDAT" => {
                if dimensions.is_none() || idat_ended || (color_type == 3 && !seen_palette) {
                    return None;
                }
                seen_idat = true;
                idat_bytes = idat_bytes.checked_add(length)?;
            }
            b"IEND" => {
                if dimensions.is_none()
                    || length != 0
                    || !seen_idat
                    || idat_bytes == 0
                    || crc_end != bytes.len()
                {
                    return None;
                }
                return dimensions;
            }
            _ => {
                if dimensions.is_none() || kind[0].is_ascii_uppercase() {
                    return None;
                }
                if seen_idat {
                    idat_ended = true;
                }
            }
        }
        if kind != b"IDAT" && seen_idat {
            idat_ended = true;
        }
        cursor = crc_end;
    }
    None
}

fn valid_png_color_depth(color_type: u8, bit_depth: u8) -> bool {
    match color_type {
        0 => matches!(bit_depth, 1 | 2 | 4 | 8 | 16),
        2 => matches!(bit_depth, 8 | 16),
        3 => matches!(bit_depth, 1 | 2 | 4 | 8),
        4 | 6 => matches!(bit_depth, 8 | 16),
        _ => false,
    }
}

fn jpeg_dimensions(bytes: &[u8], max_markers: usize) -> Option<(u32, u32)> {
    if !bytes.starts_with(b"\xff\xd8") {
        return None;
    }

    let mut cursor = 2_usize;
    let mut marker_count = 0_usize;
    let mut in_scan = false;
    let mut dimensions = None;
    let mut frame_components = [0_u8; 4];
    let mut frame_component_count = 0_usize;
    let mut seen_scan = false;
    loop {
        let marker = if in_scan {
            next_jpeg_scan_marker(bytes, &mut cursor)?
        } else {
            next_jpeg_marker(bytes, &mut cursor)?
        };
        marker_count = marker_count.checked_add(1)?;
        if marker_count > max_markers {
            return None;
        }

        match marker {
            0xd9 => {
                return (seen_scan && dimensions.is_some() && cursor == bytes.len())
                    .then_some(dimensions?);
            }
            0xd8 | 0x00 | 0xff => return None,
            0xd0..=0xd7 if in_scan => continue,
            0xd0..=0xd7 => return None,
            0x01 => in_scan = false,
            0xda => {
                dimensions?;
                let segment = jpeg_segment(bytes, &mut cursor)?;
                let component_count = usize::from(*segment.first()?);
                if !(1..=4).contains(&component_count) || segment.len() != 4 + component_count * 2 {
                    return None;
                }
                let mut scan_components = [0_u8; 4];
                for index in 0..component_count {
                    let component = segment[1 + index * 2];
                    if !frame_components[..frame_component_count].contains(&component)
                        || scan_components[..index].contains(&component)
                    {
                        return None;
                    }
                    scan_components[index] = component;
                }
                seen_scan = true;
                in_scan = true;
            }
            marker if is_jpeg_start_of_frame(marker) => {
                if dimensions.is_some() || in_scan {
                    return None;
                }
                let segment = jpeg_segment(bytes, &mut cursor)?;
                if segment.len() < 9 {
                    return None;
                }
                let component_count = usize::from(segment[5]);
                if !(1..=4).contains(&component_count) || segment.len() != 6 + component_count * 3 {
                    return None;
                }
                let height = u32::from(u16::from_be_bytes(segment.get(1..3)?.try_into().ok()?));
                let width = u32::from(u16::from_be_bytes(segment.get(3..5)?.try_into().ok()?));
                if width == 0 || height == 0 || segment[0] == 0 {
                    return None;
                }
                for index in 0..component_count {
                    let offset = 6 + index * 3;
                    let component = segment[offset];
                    let sampling = segment[offset + 1];
                    if frame_components[..index].contains(&component)
                        || sampling >> 4 == 0
                        || sampling & 0x0f == 0
                        || segment[offset + 2] > 3
                    {
                        return None;
                    }
                    frame_components[index] = component;
                }
                frame_component_count = component_count;
                dimensions = Some((width, height));
            }
            marker if marker < 0xc0 => return None,
            _ => {
                in_scan = false;
                jpeg_segment(bytes, &mut cursor)?;
            }
        }
    }
}

fn next_jpeg_marker(bytes: &[u8], cursor: &mut usize) -> Option<u8> {
    if *bytes.get(*cursor)? != 0xff {
        return None;
    }
    while bytes.get(*cursor) == Some(&0xff) {
        *cursor = cursor.checked_add(1)?;
    }
    let marker = *bytes.get(*cursor)?;
    *cursor = cursor.checked_add(1)?;
    (marker != 0 && marker != 0xff).then_some(marker)
}

fn next_jpeg_scan_marker(bytes: &[u8], cursor: &mut usize) -> Option<u8> {
    while *cursor < bytes.len() {
        if bytes[*cursor] != 0xff {
            *cursor = cursor.checked_add(1)?;
            continue;
        }
        *cursor = cursor.checked_add(1)?;
        while bytes.get(*cursor) == Some(&0xff) {
            *cursor = cursor.checked_add(1)?;
        }
        let marker = *bytes.get(*cursor)?;
        *cursor = cursor.checked_add(1)?;
        if marker == 0 {
            continue;
        }
        return Some(marker);
    }
    None
}

fn jpeg_segment<'a>(bytes: &'a [u8], cursor: &mut usize) -> Option<&'a [u8]> {
    let length_end = cursor.checked_add(2)?;
    let length = usize::from(u16::from_be_bytes(
        bytes.get(*cursor..length_end)?.try_into().ok()?,
    ));
    if length < 2 {
        return None;
    }
    let segment_end = cursor.checked_add(length)?;
    let segment = bytes.get(length_end..segment_end)?;
    *cursor = segment_end;
    Some(segment)
}

fn is_jpeg_start_of_frame(marker: u8) -> bool {
    matches!(
        marker,
        0xc0 | 0xc1 | 0xc2 | 0xc3 | 0xc5 | 0xc6 | 0xc7 | 0xc9 | 0xca | 0xcb | 0xcd | 0xce | 0xcf
    )
}

fn format_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(
        DiagnosticCode::FormatInvalid,
        Phase::Identify,
        None,
        message,
    )
    .in_part(part)
}

fn preview_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message).in_part(part)
}

fn iwa_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message).in_part(part)
}

fn image_limit_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(
        DiagnosticCode::ImageDimensionLimit,
        Phase::Render,
        None,
        message,
    )
    .in_part(part)
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;

    use super::{DOCUMENT_COMPONENT, detect_and_parse, push_keynote_native_object, read_varint};
    use crate::diagnostic::{DiagnosticCode, Phase, Severity};
    use crate::limits::Limits;
    use crate::model::{
        DocumentFormat, DocumentKind, Geometry, ImageCrop, LineAlignment, MappingQuality,
        ObjectKind, Paint, PathCommand, Rect, SourceLocator, TextAlign, TextHorizontalOverflow,
        TextLayout, TextRun, TextVerticalAlign, UnitKind, Visual,
    };

    const PROPERTIES: &[u8] =
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><plist version=\"1.0\"><dict/></plist>";
    const BINARY_PROPERTIES: &[u8] = &[
        0x62, 0x70, 0x6c, 0x69, 0x73, 0x74, 0x30, 0x30, 0xd0, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x09,
    ];

    #[test]
    fn applies_keynote_geometry_rotation_around_the_shape_center() {
        let mut position = protobuf_fixed32(1, 10.0);
        position.extend_from_slice(&protobuf_fixed32(2, 20.0));
        let mut size = protobuf_fixed32(1, 30.0);
        size.extend_from_slice(&protobuf_fixed32(2, 40.0));
        let mut bytes = protobuf_message(1, &position);
        bytes.extend_from_slice(&protobuf_message(2, &size));
        bytes.extend_from_slice(&protobuf_fixed32(4, 90.0));

        let geometry = super::keynote_geometry(&bytes, "Index/Slide.iwa").unwrap();
        assert_eq!(geometry.rotation_degrees, 90.0);
        let visual = super::keynote_wrap_geometry_transform(
            geometry,
            Rect {
                x: 10.0,
                y: 20.0,
                width: 30.0,
                height: 40.0,
            },
            Visual::None,
        );
        let Visual::Layer { transform, .. } = visual else {
            panic!("rotated Keynote geometry should produce an affine layer");
        };
        assert!(transform.a.abs() < 0.0001);
        assert!((transform.b + 1.0).abs() < 0.0001);
        assert!((transform.c - 1.0).abs() < 0.0001);
        assert!(transform.d.abs() < 0.0001);
        assert!((transform.e + 15.0).abs() < 0.0001);
        assert!((transform.f - 65.0).abs() < 0.0001);
    }

    #[test]
    fn applies_keynote_geometry_horizontal_flip() {
        let mut position = protobuf_fixed32(1, 10.0);
        position.extend_from_slice(&protobuf_fixed32(2, 20.0));
        let mut size = protobuf_fixed32(1, 30.0);
        size.extend_from_slice(&protobuf_fixed32(2, 40.0));
        let mut bytes = protobuf_message(1, &position);
        bytes.extend_from_slice(&protobuf_message(2, &size));
        bytes.extend_from_slice(&protobuf_varint(3, 7));

        let geometry = super::keynote_geometry(&bytes, "Index/Slide.iwa").unwrap();
        assert!(geometry.flip_horizontal);
        assert!(!geometry.flip_vertical);
        let Visual::Layer { transform, .. } = super::keynote_wrap_geometry_transform(
            geometry,
            Rect {
                x: 10.0,
                y: 20.0,
                width: 30.0,
                height: 40.0,
            },
            Visual::None,
        ) else {
            panic!("flipped Keynote geometry should produce an affine layer");
        };
        assert_eq!(transform.a, -1.0);
        assert_eq!(transform.d, 1.0);
        assert_eq!(transform.e, 50.0);
    }

    #[test]
    fn scales_keynote_bezier_coordinates_to_the_authored_shape_size() {
        let point = |x, y| {
            let mut bytes = protobuf_fixed32(1, x);
            bytes.extend_from_slice(&protobuf_fixed32(2, y));
            bytes
        };
        let element = |kind, x, y| {
            let mut bytes = protobuf_varint(1, kind);
            bytes.extend_from_slice(&protobuf_message(2, &point(x, y)));
            protobuf_message(1, &bytes)
        };
        let mut path = element(1, 0.0, 0.0);
        path.extend_from_slice(&element(2, 400.0, 0.0));
        path.extend_from_slice(&element(2, 400.0, 200.0));
        path.extend_from_slice(&element(2, 0.0, 200.0));
        path.extend_from_slice(&protobuf_message(1, &protobuf_varint(1, 5)));
        let mut size = protobuf_fixed32(1, 100.0);
        size.extend_from_slice(&protobuf_fixed32(2, 50.0));
        let mut bezier = protobuf_message(2, &size);
        bezier.extend_from_slice(&protobuf_message(3, &path));

        let geometry = super::keynote_bezier_path(&bezier, 100.0, 50.0, "Index/Slide.iwa")
            .unwrap()
            .unwrap();
        let Geometry::Path { commands, .. } = geometry else {
            panic!("Keynote bezier path should remain a native vector path");
        };
        assert_eq!(
            super::keynote_path_coordinate_extent(&commands),
            (100.0, 50.0)
        );
    }

    #[test]
    fn renders_keynote_connection_lines_as_native_paths() {
        let point = |x, y| {
            let mut bytes = protobuf_fixed32(1, x);
            bytes.extend_from_slice(&protobuf_fixed32(2, y));
            bytes
        };
        let element = |kind, x, y| {
            let mut bytes = protobuf_varint(1, kind);
            bytes.extend_from_slice(&protobuf_message(2, &point(x, y)));
            protobuf_message(1, &bytes)
        };
        let mut path = element(1, 0.0, 0.0);
        path.extend_from_slice(&element(2, 25.0, 50.0));
        path.extend_from_slice(&element(2, 100.0, 100.0));
        let mut size = protobuf_fixed32(1, 100.0);
        size.extend_from_slice(&protobuf_fixed32(2, 100.0));
        let mut bezier = protobuf_message(2, &size);
        bezier.extend_from_slice(&protobuf_message(3, &path));
        let mut connection_source = protobuf_message(1, &bezier);
        connection_source.extend_from_slice(&protobuf_varint(2, 1));
        let path_source = protobuf_message(7, &connection_source);
        let mut shape = protobuf_message(1, &keynote_drawable(10.0, 20.0, 100.0, 100.0));
        shape.extend_from_slice(&protobuf_message(3, &path_source));
        let connection = protobuf_message(1, &shape);

        let archives = [super::IwaArchive {
            identifier: 201,
            messages: vec![super::IwaMessage {
                message_type: 3_009,
                payload: connection,
                data_references: Vec::new(),
            }],
        }];
        let shape = super::keynote_shape_payload(&archives, 201, "Index/Slide.iwa")
            .unwrap()
            .unwrap();
        let (_, _, Geometry::Path { commands, .. }) =
            super::keynote_shape(shape, "Index/Slide.iwa").unwrap()
        else {
            panic!("Keynote connection line should remain a native vector path");
        };
        assert_eq!(
            commands,
            vec![
                crate::model::PathCommand::MoveTo { x: 0.0, y: 0.0 },
                crate::model::PathCommand::LineTo { x: 0.0, y: 50.0 },
                crate::model::PathCommand::LineTo { x: 100.0, y: 50.0 },
                crate::model::PathCommand::LineTo { x: 100.0, y: 100.0 },
            ]
        );
    }

    #[test]
    fn renders_numbers_cells_from_native_tile_storage_without_a_preview() {
        let bytes = native_numbers_package();
        let document = detect_and_parse(&bytes, Limits::default())
            .expect("native Numbers package should parse")
            .expect("native Numbers package should be detected");

        assert_eq!(document.format, Some(DocumentFormat::Numbers));
        assert_eq!(document.kind, Some(DocumentKind::Spreadsheet));
        assert_eq!(document.units.len(), 1);
        assert_eq!(document.units[0].name, "Sheet1");
        assert_eq!(
            (document.units[0].rows, document.units[0].columns),
            (257, 2)
        );
        let unit = &document.units[0];
        assert!((unit.row_axis.default_size - 18.8).abs() < 0.001);
        assert_eq!(unit.row_axis.spans[0].start, 1);
        assert!((unit.row_axis.spans[0].size - 20.0).abs() < 0.001);
        assert!((unit.column_axis.default_size - 37.333_332).abs() < 0.001);
        assert_eq!(unit.column_axis.spans[0].start, 1);
        assert!((unit.column_axis.spans[0].size - 101.333_336).abs() < 0.001);
        assert!((unit.width - 138.666_67).abs() < 0.001);
        assert!((unit.height - 5_158.8).abs() < 0.01);
        assert_eq!(
            document
                .objects
                .iter()
                .filter_map(|object| object.text.as_deref())
                .collect::<Vec<_>>(),
            ["Name", "Alice", "42", "Tail"]
        );
        assert_eq!(
            document
                .objects
                .iter()
                .map(|object| object.stable_id.as_str())
                .collect::<Vec<_>>(),
            [
                "numbers:0:A1",
                "numbers:0:A3",
                "numbers:0:B3",
                "numbers:0:A257",
                "numbers:0:border:horizontal:2:0:1",
                "numbers:0:border:horizontal:2:1:1",
            ]
        );
        assert!(
            document
                .objects
                .iter()
                .filter(|object| object.kind == ObjectKind::Cell)
                .all(|object| matches!(
                    object.source.locator,
                    SourceLocator::Iwork { kind: "cell", .. }
                ) && !matches!(object.visual, Visual::Image { .. }))
        );
        let first = &document.objects[0];
        assert_eq!((first.bounds.x, first.bounds.y), (0.0, 0.0));
        assert!((first.bounds.width - 138.666_67).abs() < 0.001);
        assert!((first.bounds.height - 18.8).abs() < 0.001);
        let Visual::TextLayout { layout, visual } = &first.visual else {
            panic!("Numbers cells should have an explicit text layout");
        };
        assert!(!layout.wrap);
        let tail = document
            .objects
            .iter()
            .find(|object| object.stable_id == "numbers:0:A257")
            .unwrap();
        let Visual::TextLayout { layout, .. } = &tail.visual else {
            panic!("Numbers cells should have an explicit text layout");
        };
        assert!(!layout.wrap);
        assert_eq!(layout.horizontal_overflow, TextHorizontalOverflow::Overflow);
        let Visual::RichText { fill, runs, .. } = visual.as_ref() else {
            panic!("styled Numbers cell should remain rich text");
        };
        assert_eq!(*fill, Paint::Solid(0x70ad_47ff));
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].font_family, "Courier");
        assert!((runs[0].font_size - 16.0).abs() < 0.001);
        assert_eq!(runs[0].color, 0x0066_ffff);
        assert!(runs[0].bold);
        assert!(runs[0].italic);
        let third_row = document
            .objects
            .iter()
            .find(|object| object.stable_id == "numbers:0:A3")
            .unwrap();
        assert!((third_row.bounds.y - 38.8).abs() < 0.001);
        assert!((third_row.bounds.height - 40.0).abs() < 0.001);
        let border = document
            .objects
            .iter()
            .find(|object| object.stable_id == "numbers:0:border:horizontal:2:0:1")
            .expect("Numbers stroke-sidecar borders should be rendered");
        assert_eq!(border.kind, ObjectKind::Shape);
        assert_eq!(border.bounds.x, 0.0);
        assert!((border.bounds.y - 38.8).abs() < 0.001);
        assert!((border.bounds.width - 37.333_332).abs() < 0.001);
        assert_eq!(border.bounds.height, 0.0);
        assert!(matches!(border.visual, Visual::PaintedShape {
            geometry: Geometry::Line,
            fill: Paint::None,
            stroke: Paint::Solid(0x0000_00ff),
            stroke_width
        } if (stroke_width - 1.333_333_4).abs() < 0.001));
        let dashed_border = document
            .objects
            .iter()
            .find(|object| object.stable_id == "numbers:0:border:horizontal:2:1:1")
            .expect("Numbers patterned stroke-sidecar borders should be rendered");
        let Visual::StrokeStyle { style, visual } = &dashed_border.visual else {
            panic!("Numbers patterned borders should retain their stroke style");
        };
        assert_eq!(style.dash, [2.666_666_7, 2.666_666_7]);
        assert!(matches!(
            visual.as_ref(),
            Visual::PaintedShape {
                geometry: Geometry::Line,
                stroke: Paint::Solid(0x0000_00ff),
                ..
            }
        ));
        assert!(
            document
                .diagnostics
                .iter()
                .all(|diagnostic| !diagnostic.message.contains("PREVIEW_ONLY"))
        );
        crate::protocol::encode(&document).expect("native Numbers document should encode");
    }

    #[test]
    fn renders_original_numbers_chart_from_cross_component_reference() {
        let bytes = include_bytes!("../../tests/fixtures/chart-original.numbers");
        let document = detect_and_parse(bytes, Limits::default()).unwrap().unwrap();
        let chart = document
            .objects
            .iter()
            .find(|object| {
                matches!(
                    object.source.locator,
                    SourceLocator::Iwork { kind: "chart", .. }
                )
            })
            .expect("original Numbers chart 905361 must not be skipped");
        assert_eq!(chart.source.part, "Index/CalculationEngine-905377.iwa");
        assert_eq!(chart.unit_index, 0);
        assert_eq!(
            document
                .objects
                .iter()
                .filter(|object| object.kind == ObjectKind::Cell)
                .count(),
            4
        );
        let unit = &document.units[0];
        assert!(unit.width >= chart.bounds.x + chart.bounds.width);
        assert!(unit.height >= chart.bounds.y + chart.bounds.height);
        let bars = document
            .objects
            .iter()
            .filter(|object| {
                object.parent_numeric_id == Some(chart.numeric_id)
                    && matches!(
                        object.visual,
                        Visual::PaintedShape {
                            geometry: Geometry::Rectangle,
                            fill: Paint::Solid(_),
                            ..
                        }
                    )
            })
            .collect::<Vec<_>>();
        assert_eq!(bars.len(), 2);
        assert!((bars[0].bounds.height / bars[1].bounds.height - 10.0 / 58.0).abs() < 0.0001);
        // Native reference: geometry is the plot frame and gap=219% of a bar.
        let plot_width = 329.476_96 * super::IWORK_POINT_TO_CSS_PIXEL;
        assert!((bars[0].bounds.width - plot_width / 2.0 / 3.19).abs() < 0.01);
        let Visual::StrokeStyle { visual, .. } = &chart.visual else {
            panic!("chart background style missing")
        };
        assert!(matches!(
            visual.as_ref(),
            Visual::PaintedShape {
                fill: Paint::Solid(0xffff_ffff),
                stroke: Paint::Solid(0xd9d9_d9ff),
                ..
            }
        ));
        let grid = document.objects.iter().find(|object| object.parent_numeric_id == Some(chart.numeric_id)
            && matches!(&object.visual, Visual::StrokeStyle { visual, .. } if matches!(visual.as_ref(),
                Visual::PaintedShape { stroke: Paint::Solid(0xd9d9_d9ff), .. }))).expect("native gray grid lines missing");
        assert!((grid.bounds.width - plot_width).abs() < 0.01);
        assert!((grid.bounds.x - 137.023_04 * super::IWORK_POINT_TO_CSS_PIXEL).abs() < 0.01);
        for object in document
            .objects
            .iter()
            .filter(|object| object.parent_numeric_id == Some(chart.numeric_id))
        {
            if let Visual::TextLayout { layout, visual } = &object.visual {
                assert!(!layout.wrap);
                assert_eq!(layout.inset_left + layout.inset_right, 0.0);
                let Visual::RichText { runs, .. } = visual.as_ref() else {
                    panic!("chart label must contain native text")
                };
                for run in runs {
                    assert!(
                        !matches!(run.text.as_str(), "10" | "58"),
                        "authored value labels are disabled"
                    );
                    assert_eq!(run.color, 0x5959_59ff);
                    assert!(!run.bold);
                    assert_eq!(run.font_family, "Calibri");
                    if run.text == "Chart Title" {
                        assert!(
                            (run.font_size - 14.0 * super::IWORK_POINT_TO_CSS_PIXEL).abs() < 0.001
                        );
                    }
                }
            }
        }
        assert!(
            !document
                .objects
                .iter()
                .any(|object| matches!(object.visual, Visual::Image { .. }))
        );
        crate::protocol::encode(&document).unwrap();
    }

    #[test]
    fn numbers_charts_preserve_sheet_ownership_and_isolate_invalid_payloads() {
        let limits = Limits::default();
        let original = super::Package::open_iwork(
            include_bytes!("../../tests/fixtures/chart-original.numbers"),
            limits,
        )
        .unwrap();
        let part = "Index/CalculationEngine-905377.iwa";
        let archives = super::parse_iwa_archives(
            &original,
            part,
            &original.required_part(part).unwrap(),
            limits,
        )
        .unwrap();
        let chart = super::archive_message(&archives, 905_361, super::IWORK_CHART_TYPE, part)
            .unwrap()
            .unwrap();
        let native = native_numbers_package();
        let package = super::Package::open_iwork(&native, limits).unwrap();
        for (chart_only, invalid) in [(false, false), (false, true), (true, false)] {
            let mut sheet = protobuf_message(1, b"Sheet1");
            // A floating chart before the table must not rename or duplicate the sheet.
            sheet.extend(protobuf_message(2, &protobuf_reference(905_361)));
            if !chart_only {
                sheet.extend(protobuf_message(2, &protobuf_reference(10)));
            }
            let mut root = iwa_numbers_document();
            root.extend(iwa_stream(&[iwa_archive(2, 2, &sheet, &[])]));
            let chart_part = iwa_stream(&[iwa_archive(
                905_361,
                super::IWORK_CHART_TYPE,
                if invalid { &[0x0f] } else { &chart.payload },
                &[],
            )]);
            let mut parts = package
                .entry_names()
                .map(|name| {
                    (
                        name.to_owned(),
                        if name == "Index/Document.iwa" {
                            root.clone()
                        } else {
                            package.required_part(name).unwrap().to_vec()
                        },
                    )
                })
                .collect::<Vec<_>>();
            parts.push(("Index/Charts.iwa".to_owned(), chart_part));
            let bytes = stored_zip(
                &parts
                    .iter()
                    .map(|(name, bytes)| (name.as_str(), bytes.as_slice()))
                    .collect::<Vec<_>>(),
            );
            let document = detect_and_parse(&bytes, limits).unwrap().unwrap();
            assert_eq!(document.units.len(), 1);
            assert_eq!(document.units[0].name, "Sheet1");
            assert_eq!(
                document
                    .objects
                    .iter()
                    .filter(|object| matches!(
                        object.source.locator,
                        SourceLocator::Iwork { kind: "chart", .. }
                    ))
                    .count(),
                usize::from(!invalid)
            );
            assert_eq!(
                document
                    .objects
                    .iter()
                    .filter(|object| object.kind == ObjectKind::Cell)
                    .count(),
                if chart_only { 0 } else { 4 }
            );
            assert_eq!(
                document
                    .diagnostics
                    .iter()
                    .any(|diagnostic| diagnostic.message.starts_with("NUMBERS_CHART_OMITTED:")),
                invalid
            );
            assert!(document.objects.iter().all(|object| object.unit_index == 0));
            for object in &document.objects {
                assert!(document.units[0].width >= object.bounds.x + object.bounds.width);
                assert!(document.units[0].height >= object.bounds.y + object.bounds.height);
            }
            crate::protocol::encode(&document).unwrap();
            if chart_only {
                let error = detect_and_parse(
                    &bytes,
                    Limits {
                        max_document_objects: 4,
                        ..limits
                    },
                )
                .unwrap_err();
                assert_eq!(error.code, DiagnosticCode::ObjectLimit);
            }
        }
    }

    #[test]
    fn accepts_document_sized_numbers_repeated_fields() {
        let mut payload = Vec::new();
        for index in 0..2_048 {
            payload.extend_from_slice(&protobuf_message(3, &protobuf_varint(1, index)));
        }

        assert_eq!(
            super::numbers_bytes(&payload, 3, "Index/Tables/DataList.iwa")
                .unwrap()
                .len(),
            2_048
        );
    }

    #[test]
    fn identifies_pages_numbers_and_keynote_from_native_contents() {
        let cases = [
            (
                10_000,
                None,
                DocumentFormat::Pages,
                DocumentKind::Text,
                UnitKind::Page,
            ),
            (
                1,
                Some("Index/Slide-42.iwa"),
                DocumentFormat::Keynote,
                DocumentKind::Presentation,
                UnitKind::Slide,
            ),
        ];

        for (message_type, marker, format, kind, unit_kind) in cases {
            let bytes = iwork_package(message_type, marker, "preview.jpg", &jpeg(3, 2));
            let document = detect_and_parse(&bytes, Limits::default())
                .expect("native iWork package should parse")
                .expect("native iWork package should be detected");

            assert_eq!(document.format, Some(format));
            assert_eq!(document.kind, Some(kind));
            assert_eq!(document.units.len(), 1);
            assert_eq!(document.units[0].kind, unit_kind);
            assert_eq!(
                (document.units[0].width, document.units[0].height),
                (3.0, 2.0)
            );
            assert_eq!((document.units[0].rows, document.units[0].columns), (0, 0));
            assert_eq!(document.objects.len(), 1);
            assert!(matches!(
                &document.objects[0].source.locator,
                SourceLocator::Iwork {
                    kind: "preview",
                    component,
                } if component == "preview.jpg"
            ));
            assert!(matches!(
                &document.objects[0].visual,
                Visual::Image { media_type, .. } if media_type == "image/jpeg"
            ));
            assert_eq!(document.diagnostics.len(), 1);
            assert!(
                document.diagnostics[0]
                    .message
                    .starts_with("IWORK_PREVIEW_ONLY:")
            );
            assert_eq!(document.diagnostics[0].severity, Severity::Warning);
            crate::protocol::encode(&document).expect("iWork preview document should encode");
        }
    }

    #[test]
    fn clips_shared_iwork_table_borders_and_preserves_unreadable_sidecar_diagnostics() {
        let mut border = super::NumbersTableBorder {
            horizontal: false,
            boundary: 1,
            origin: 0,
            length: 4,
            stroke: super::NumbersStroke {
                color: 0x0000_00ff,
                width: 0.5,
                style: super::StrokeStyle::default(),
            },
        };
        let columns = [0.0, 80.0, 160.0];
        let rows = [0.0, 20.0, 40.0];
        let (bounds, visual) =
            super::iwork_table_border_visual(&border, &columns, &rows, 2).unwrap();
        assert_eq!(
            bounds,
            Rect {
                x: 80.0,
                y: 0.0,
                width: 0.0,
                height: 40.0
            }
        );
        assert!(matches!(
            visual,
            Visual::PaintedShape {
                stroke_width: 0.5,
                ..
            }
        ));
        border.horizontal = true;
        border.length = 2;
        assert!(super::iwork_table_border_visual(&border, &columns, &rows, 2).is_none());
        border.boundary = 2;
        let (bounds, _) = super::iwork_table_border_visual(&border, &columns, &rows, 2).unwrap();
        assert_eq!(
            bounds,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 160.0,
                height: 0.0
            }
        );

        let model = super::IwaMessage {
            message_type: super::NUMBERS_TABLE_MODEL_TYPE,
            payload: protobuf_message(49, &protobuf_reference(999)),
            data_references: Vec::new(),
        };
        let messages = super::NumbersMessageSpace {
            archives: Vec::new(),
        };
        let mut diagnostics = Vec::new();
        let borders =
            super::iwork_table_borders(&messages, &model, "table.iwa", 4, 2, &mut diagnostics);
        assert!(borders.unwrap().is_empty());
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(diagnostics[0].severity, Severity::Warning);
        assert!(
            diagnostics[0]
                .message
                .contains("stroke sidecar 999 is missing")
        );
    }

    #[test]
    fn preserves_supplied_pages_table_styles() {
        let parsed = detect_and_parse(
            include_bytes!("../../tests/fixtures/pages-text-alignment.pages"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let cells = parsed
            .objects
            .iter()
            .filter(|object| {
                object.unit_index == 17 && object.kind == ObjectKind::Cell && object.text.is_some()
            })
            .collect::<Vec<_>>();
        assert_eq!(cells.len(), 15);
        for (index, cell) in cells.iter().enumerate() {
            let Visual::RichText { align, runs, .. } = &cell.visual else {
                panic!("table cell text");
            };
            assert_eq!(
                *align,
                TextAlign::Start,
                "cell {index} must preserve authored left alignment"
            );
            assert_eq!(runs[0].bold, index == 0, "only the red first cell is bold");
            assert!((runs[0].font_size - 16.0).abs() < 0.001);
            assert_eq!(
                runs[0].color,
                if index == 0 { 0xff00_00ff } else { 0x0000_00ff }
            );
        }
        let table = parsed
            .objects
            .iter()
            .find(|object| object.unit_index == 17 && object.kind == ObjectKind::Table)
            .unwrap();
        let Visual::StrokeStyle { visual, .. } = &table.visual else {
            panic!("table background");
        };
        assert!(
            matches!(
                visual.as_ref(),
                Visual::PaintedShape {
                    stroke: Paint::None,
                    stroke_width: 0.0,
                    ..
                }
            ),
            "no invented thick outer frame"
        );
        let lines = parsed
            .objects
            .iter()
            .filter(|object| object.unit_index == 17 && object.kind == ObjectKind::Shape)
            .collect::<Vec<_>>();
        assert_eq!(
            lines.len(),
            10,
            "four horizontal and six vertical authored borders"
        );
        for line in lines {
            assert!(
                matches!(
                    &line.source.locator,
                    SourceLocator::Iwork {
                        kind: "table-grid",
                        ..
                    }
                ),
                "border locator must use the existing Pages/Keynote protocol contract"
            );

            let visual = if let Visual::StrokeStyle { visual, .. } = &line.visual {
                visual.as_ref()
            } else {
                &line.visual
            };
            let Visual::PaintedShape { stroke_width, .. } = visual else {
                panic!("border stroke");
            };
            assert!(
                (*stroke_width - 0.5 * super::IWORK_POINT_TO_CSS_PIXEL).abs() < 0.001,
                "border must remain 0.5pt"
            );
        }
    }

    #[test]
    fn preserves_supplied_pages_paragraph_alignment() {
        let parsed = detect_and_parse(
            include_bytes!("../../tests/fixtures/pages-text-alignment.pages"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        assert_eq!(parsed.units.len(), 18);
        let second_page = parsed
            .objects
            .iter()
            .find(|object| object.unit_index == 1)
            .unwrap();
        let Visual::TextLayout {
            layout: continuation,
            ..
        } = &second_page.visual
        else {
            panic!("Pages continuation layout");
        };
        assert_eq!(
            continuation.paragraphs[0].first_line_indent, 0.0,
            "a paragraph continuing on page 2 has no first-line indent"
        );
        let Visual::TextLayout { layout, visual } = &parsed.objects[0].visual else {
            panic!("Pages body layout");
        };
        assert_eq!(layout.paragraphs[0].align, TextAlign::Center);
        assert_eq!(layout.horizontal_overflow, TextHorizontalOverflow::Overflow);
        assert_eq!(
            layout.vertical_overflow,
            super::TextVerticalOverflow::Overflow
        );
        assert_eq!(layout.paragraphs[2].align, TextAlign::Start);
        assert_eq!(
            layout.paragraphs[3].align,
            TextAlign::Start,
            "both dates are left aligned in native Pages"
        );
        assert_eq!(layout.paragraphs[7].align, TextAlign::Justify);
        assert!(
            (layout.paragraphs[7].first_line_indent - 32.0 * super::IWORK_POINT_TO_CSS_PIXEL).abs()
                < 0.001
        );
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("Pages body runs");
        };
        assert!(
            runs.iter().all(|run| run.letter_spacing == 0.0),
            "native Pages uses zero tracking"
        );
        assert!(runs.iter().any(|run| run.text.contains("更新时间") && run.font_family == "PingFangSC-Regular"));
        assert!(
            runs.iter()
                .any(|run| run.text.contains("2023") && run.font_family == "Courier")
        );
    }

    #[test]
    fn preserves_supplied_pages_paragraph_rules_and_list_levels() {
        let document = detect_and_parse(
            include_bytes!("../../tests/fixtures/pages-note-list-rule.pages"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let body = document
            .objects
            .iter()
            .find(|object| {
                object
                    .text
                    .as_deref()
                    .is_some_and(|text| text.contains("Title\nSubject\n"))
            })
            .expect("Pages note body");
        let Visual::TextLayout { layout, visual } = &body.visual else {
            panic!("Pages note body layout");
        };
        assert_eq!(layout.paragraphs.len(), 10);
        let rule = layout.paragraphs[3]
            .rule_above
            .expect("Subject paragraph must retain its authored top rule");
        assert_eq!(rule.color, 0x5151_51ff);
        assert!((rule.stroke_width - 0.5 * super::IWORK_POINT_TO_CSS_PIXEL).abs() < 0.001);
        assert_eq!(
            layout.paragraphs[4].margin_left,
            12.0 * super::IWORK_POINT_TO_CSS_PIXEL
        );
        assert_eq!(
            layout.paragraphs[6].margin_left,
            24.0 * super::IWORK_POINT_TO_CSS_PIXEL
        );
        assert_eq!(
            layout.paragraphs[7].margin_left,
            24.0 * super::IWORK_POINT_TO_CSS_PIXEL
        );
        assert_eq!(
            layout.paragraphs[8].margin_left,
            12.0 * super::IWORK_POINT_TO_CSS_PIXEL
        );
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("Pages note rich text");
        };
        assert_eq!(
            runs.iter().filter(|run| run.text.starts_with('\t')).count(),
            6,
            "list bodies must use the authored hanging-indent anchor"
        );
        assert!(runs.iter().any(|run| {
            run.text == " "
                && (run.letter_spacing - (12.0 - 12.0 * 0.278) * super::IWORK_POINT_TO_CSS_PIXEL)
                    .abs()
                    < 0.001
        }));
        assert_eq!(
            runs.iter()
                .filter(|run| run.font_family == "ArialMT")
                .map(|run| run.text.as_str())
                .collect::<Vec<_>>(),
            vec!["-", "-", "•", "•", "-", "-"]
        );
    }

    #[test]
    fn preserves_supplied_keynote_first_line_indent() {
        let parsed = detect_and_parse(
            include_bytes!("../../tests/fixtures/keynote-first-line-indent.key"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let body = parsed
            .objects
            .iter()
            .find(|object| {
                object.unit_index == 1
                    && object
                        .text
                        .as_deref()
                        .is_some_and(|text| text.contains("PowerPoint is a very commonly"))
            })
            .expect("Introduction body");
        let Visual::TextLayout { layout, visual } = &body.visual else {
            panic!("Introduction paragraph layout");
        };
        // The opening tab inherits a 72pt interval from paragraph style 152.
        assert_eq!(layout.paragraphs[0].default_tab_stop, 96.0);
        assert_eq!(layout.paragraphs[0].margin_left, 0.0);
        assert_eq!(layout.paragraphs[0].first_line_indent, 0.0);
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("Introduction text");
        };
        assert_eq!(runs[0].text, "\t");
    }

    #[cfg(feature = "pdf-formats")]
    #[test]
    fn suppresses_supplied_keynote_master_image_placeholder() {
        let parsed = detect_and_parse(
            include_bytes!("../../tests/fixtures/keynote-first-line-indent.key"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        assert!(
            !parsed
                .objects
                .iter()
                .any(|object| object.stable_id.starts_with("iwork-15-657-pdf-")),
            "master sample image must not appear behind the authored video image"
        );
        let video = parsed
            .objects
            .iter()
            .find(|object| object.stable_id.ends_with("Index/Slide-1485-2.iwa-1517"))
            .expect("authored video image");
        assert!(matches!(video.visual, Visual::Image { .. }));
        assert_eq!(video.bounds.x, 576.0);
    }

    #[test]
    fn preserves_supplied_keynote_research_paragraph_spacing() {
        let parsed = detect_and_parse(
            include_bytes!("../../tests/fixtures/keynote-first-line-indent.key"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let label = parsed
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Research\n&\nDevelopment"))
            .expect("three-line diamond label");
        let Visual::TextLayout { layout, .. } = &label.visual else {
            panic!("diamond label layout");
        };
        assert_eq!(layout.paragraphs.len(), 3);
        // Paragraph style 1458 authors 10.8pt before each paragraph; frame top suppresses it.
        for (paragraph, before) in layout.paragraphs.iter().zip([0.0, 14.4, 14.4]) {
            assert!(
                (paragraph.space_before - before).abs() < 0.001,
                "expected paragraph spacing {before}, got {}",
                paragraph.space_before
            );
        }
        assert_eq!(label.bounds.y, 496.0, "preserve the authored text-box top");
    }

    #[test]
    fn preserves_supplied_keynote_cross_shape() {
        let parsed = detect_and_parse(
            include_bytes!("../../tests/fixtures/keynote-first-line-indent.key"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let cross = parsed
            .objects
            .iter()
            .find(|object| object.stable_id.ends_with("Index/Slide-1370-2.iwa-1906"))
            .expect("red cross on AutoShapes slide");
        let Visual::StrokeStyle { visual, .. } = &cross.visual else {
            panic!("cross stroke style");
        };
        let Visual::PaintedShape {
            geometry: Geometry::Path { commands, .. },
            fill,
            ..
        } = visual.as_ref()
        else {
            panic!("native cross path: {:?}", cross.visual);
        };
        assert_eq!(*fill, Paint::Solid(0xff00_00ff));
        assert_eq!(
            commands.len(),
            13,
            "cross has twelve corners, not an arrow's seven"
        );
        for command in commands.iter().take(12) {
            let (PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y }) = command else {
                panic!("cross corner");
            };
            assert!([0.0, 18.0, 54.0, 72.0].contains(x));
            assert!([0.0, 18.0, 54.0, 72.0].contains(y));
        }
    }

    #[test]
    fn keynote_cross_scales_authored_arm_widths_independently() {
        let mut path = protobuf_varint(1, 200);
        let mut point = protobuf_fixed32(1, 60.0);
        point.extend_from_slice(&protobuf_fixed32(2, 20.0));
        path.extend_from_slice(&protobuf_message(2, &point));
        let mut size = protobuf_fixed32(1, 80.0);
        size.extend_from_slice(&protobuf_fixed32(2, 40.0));
        path.extend_from_slice(&protobuf_message(3, &size));
        assert_eq!(
            super::keynote_point_path(&path, 120.0, 80.0, "Index/Slide.iwa").unwrap(),
            crate::format::cross_geometry(120.0, 80.0, 15.0, 105.0, 20.0, 60.0),
        );
    }

    #[test]
    fn renders_saved_pages_pagination_instead_of_a_single_preview() {
        let mut root = protobuf_message(4, &protobuf_reference(2));
        for (field, value) in [
            (30, 595.0),
            (31, 842.0),
            (32, 56.7),
            (33, 56.7),
            (34, 0.0),
            (35, 56.7),
            (36, 36.0),
            (37, 36.0),
        ] {
            root.extend_from_slice(&protobuf_fixed32(field, value));
        }
        let body_text = b"First page\nSecond page";
        let body = protobuf_message(3, body_text);
        let document = iwa_stream(&[
            iwa_archive(1, 10_000, &root, &[]),
            iwa_archive(2, 2_001, &body, &[]),
        ]);

        let target_hint =
            |location: u64, length: u64, origin: (f64, f64), size: (f64, f64), columns: u64| {
                let mut range = protobuf_varint(1, location);
                range.extend_from_slice(&protobuf_varint(2, length));
                let mut target = protobuf_message(3, &range);
                target.extend_from_slice(&protobuf_message(5, &range));
                target.extend_from_slice(&protobuf_varint(6, columns));
                let mut encoded_origin = protobuf_fixed64(1, origin.0);
                encoded_origin.extend_from_slice(&protobuf_fixed64(2, origin.1));
                target.extend_from_slice(&protobuf_message(8, &encoded_origin));
                let mut encoded_size = protobuf_fixed64(1, size.0);
                encoded_size.extend_from_slice(&protobuf_fixed64(2, size.1));
                target.extend_from_slice(&protobuf_message(9, &encoded_size));
                protobuf_message(2, &target)
            };
        let mut first_page = target_hint(0, 6, (0.0, 10.0), (200.0, 100.0), 1);
        first_page.extend_from_slice(&target_hint(6, 5, (0.0, 110.0), (432.0, 639.3), 2));
        let second_page = target_hint(11, 11, (0.0, 0.0), (432.0, 749.3), 1);
        let mut section = protobuf_message(1, &first_page);
        section.extend_from_slice(&protobuf_message(1, &second_page));
        let mut layout_state = protobuf_message(5, &section);
        layout_state.extend_from_slice(&protobuf_varint(6, 22));
        let view_state = iwa_stream(&[iwa_archive(3, 10_131, &layout_state, &[])]);
        let preview = jpeg(3, 2);
        let package = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/ViewState-3.iwa", &view_state),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);

        let parsed = detect_and_parse(&package, Limits::default())
            .expect("saved Pages pagination should parse")
            .expect("Pages package should be detected");
        assert_eq!(parsed.units.len(), 2);
        assert_eq!(
            parsed
                .units
                .iter()
                .map(|unit| (unit.kind, unit.width, unit.height))
                .collect::<Vec<_>>(),
            vec![
                (UnitKind::Page, 793.3334, 1122.6667),
                (UnitKind::Page, 793.3334, 1122.6667),
            ]
        );
        assert_eq!(
            parsed
                .objects
                .iter()
                .filter_map(|object| object.text.as_deref())
                .collect::<Vec<_>>(),
            vec!["First ", "page\n", "Second page"]
        );
        let text_objects = parsed
            .objects
            .iter()
            .filter(|object| object.kind == ObjectKind::TextBox)
            .collect::<Vec<_>>();
        assert!((text_objects[0].bounds.y - 61.333334).abs() < 0.001);
        assert!((text_objects[0].bounds.width - 266.6667).abs() < 0.001);
        let Visual::TextLayout { layout, .. } = &text_objects[1].visual else {
            panic!("saved Pages target should remain a text layout");
        };
        assert_eq!(layout.column_count, 2);
        assert!((layout.column_spacing - 48.0).abs() < 0.001);
        assert!(
            parsed.objects.iter().all(|object| !matches!(
                object.source.locator,
                SourceLocator::Iwork {
                    kind: "preview",
                    ..
                }
            )),
            "native Pages pages must never be covered by a raster preview"
        );
        assert!(
            parsed.diagnostics[0]
                .message
                .starts_with("PAGES_NATIVE_STATIC:")
        );
        crate::protocol::encode(&parsed).expect("native Pages document should encode");
    }

    #[test]
    fn shares_bounded_iwork_text_fields_without_shifting_utf16_ranges() {
        let storage = super::KeynoteTextStorage {
            text: "😀\u{fffc}X".to_owned(),
            attachments: vec![super::KeynoteStyleChange {
                character_index: 2,
                identifier: Some(42),
            }],
            ..Default::default()
        };
        let archives = [super::IwaArchive {
            identifier: 42,
            messages: vec![super::IwaMessage {
                message_type: super::IWORK_TEXT_FIELD_TYPE,
                payload: protobuf_message(2, b"12"),
                data_references: Vec::new(),
            }],
        }];
        let limits = Limits::default();
        assert_eq!(
            super::iwork_text_with_fields(&storage, &archives, DOCUMENT_COMPONENT, 2, 4, limits)
                .unwrap(),
            "12X"
        );
        assert!(
            super::iwork_text_with_fields(&storage, &archives, DOCUMENT_COMPONENT, 1, 4, limits)
                .is_err()
        );
        let tiny = Limits {
            max_entry_uncompressed_bytes: 1,
            ..limits
        };
        assert_eq!(
            super::iwork_text_with_fields(&storage, &archives, DOCUMENT_COMPONENT, 2, 3, tiny)
                .unwrap_err()
                .code,
            DiagnosticCode::ZipEntryTooLarge
        );
        assert_eq!(
            super::iwork_text_with_fields(
                &storage,
                &archives,
                DOCUMENT_COMPONENT,
                2,
                4,
                Limits {
                    max_entry_uncompressed_bytes: 2,
                    ..limits
                },
            )
            .unwrap_err()
            .code,
            DiagnosticCode::ZipEntryTooLarge
        );
        let pages =
            super::pages_text_runs_for_range(&storage, &archives, DOCUMENT_COMPONENT, 0, 4, limits)
                .unwrap();
        let keynote = super::keynote_text_runs(
            &storage,
            &archives,
            DOCUMENT_COMPONENT,
            &super::KeynoteTextStyle::default(),
            limits,
        )
        .unwrap();
        for runs in [pages, keynote] {
            assert_eq!(
                runs.iter().map(|run| run.text.as_str()).collect::<String>(),
                "😀12X"
            );
        }
    }

    #[test]
    fn renders_supplied_pages_toc_across_saved_page_ranges() {
        let parsed = detect_and_parse(
            include_bytes!("../../tests/fixtures/word-original.pages"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let text = |page| {
            parsed
                .objects
                .iter()
                .filter(|object| object.unit_index == page)
                .filter_map(|object| object.text.as_deref())
                .collect::<String>()
        };
        assert!(
            text(2).contains("词汇表\t5"),
            "first TOC page must render entries and saved page numbers"
        );
        assert!(
            text(3).contains("地表水\t22"),
            "TOC continuation must render on its saved page"
        );
        assert!(
            !text(2).contains("地表水\t22"),
            "continuation must not be duplicated on page 3"
        );
        assert_eq!(parsed.units.len(), 31);
        for object in parsed.objects.iter().filter(|object| {
            [2, 3].contains(&object.unit_index) && object.stable_id.contains("-toc-")
        }) {
            let Visual::TextLayout { layout, .. } = &object.visual else {
                panic!("TOC entry needs native text layout")
            };
            assert!(!layout.tab_stops.is_empty());
            assert!(
                layout
                    .tab_stops
                    .iter()
                    .all(|tab| tab.align == TextAlign::End)
            );
            assert!(
                object.bounds.y + object.bounds.height
                    < parsed.units[object.unit_index as usize].height,
                "every TOC entry must be inside its saved page"
            );
        }
    }

    #[test]
    fn renders_supplied_pages_first_page_cell_images() {
        let package = crate::package::Package::open_iwork(
            include_bytes!("../../tests/fixtures/word-original.pages"),
            Limits::default(),
        )
        .unwrap();
        let parsed = detect_and_parse(
            include_bytes!("../../tests/fixtures/word-original.pages"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let images = parsed
            .objects
            .iter()
            .filter(|object| object.unit_index == 0 && object.kind == ObjectKind::Image)
            .collect::<Vec<_>>();
        assert_eq!(
            images.len(),
            2,
            "both authored table-cell logos must render"
        );
        let table = parsed
            .objects
            .iter()
            .find(|object| object.unit_index == 0 && object.kind == ObjectKind::Table)
            .unwrap();
        let body = parsed
            .objects
            .iter()
            .find(|object| {
                object.unit_index == 0
                    && object
                        .text
                        .as_deref()
                        .is_some_and(|text| text.contains("控制危险废物"))
            })
            .unwrap();
        let Visual::TextLayout { layout, .. } = &body.visual else {
            panic!("body must retain paragraph layout")
        };
        assert!(
            layout.paragraphs[0].line_height >= table.bounds.height - 0.01,
            "body text must start below the complete inline table"
        );
        for (image, part) in images
            .iter()
            .zip(["Data/image1-31.png", "Data/image2-33.jpeg"])
        {
            let Visual::Image { bytes, .. } = &image.visual else {
                panic!("cell logo must use native image data")
            };
            assert_eq!(
                bytes.as_slice(),
                package.required_part(part).unwrap().as_ref(),
                "missing {part}"
            );
            assert!(image.bounds.width > 100.0 && image.bounds.height > 50.0);
            assert!(image.bounds.y + image.bounds.height <= table.bounds.y + table.bounds.height);
            assert!(matches!(
                image.source.locator,
                SourceLocator::Iwork {
                    kind: "inline-image",
                    ..
                }
            ));
        }
    }

    #[test]
    fn preserves_iwork_rich_cell_storage_in_both_tile_encodings() {
        let rich = [Some(super::IworkCellText {
            text: "\u{fffc}".to_owned(),
            storage: Some(42),
        })];
        for version in [4, 5] {
            let mut bytes = vec![0; 16];
            bytes[0] = version;
            bytes[1] = 9;
            if version == 4 {
                bytes[4..8].copy_from_slice(&0x0200_u32.to_le_bytes());
            } else {
                bytes[8..12].copy_from_slice(&0x0010_u32.to_le_bytes());
            }
            let value = super::numbers_cell_text(&bytes, &[], &rich, "tile")
                .unwrap()
                .unwrap();
            assert_eq!(value.storage, Some(42));
            assert_eq!(value.text, "\u{fffc}");
        }
        assert_eq!(super::iwork_visible_table_text("A\u{fffc}B"), "AB");
    }

    #[test]
    fn isolates_corrupt_supplied_pages_cell_image() {
        let package = crate::package::Package::open_iwork(
            include_bytes!("../../tests/fixtures/word-original.pages"),
            Limits::default(),
        )
        .unwrap();
        let parts = package
            .entry_names()
            .map(|name| {
                let bytes = if name == "Data/image1-31.png" {
                    b"broken image".to_vec()
                } else {
                    package.required_part(name).unwrap().into_vec()
                };
                (name, bytes)
            })
            .collect::<Vec<_>>();
        let bytes = stored_zip(
            &parts
                .iter()
                .map(|(name, bytes)| (*name, bytes.as_slice()))
                .collect::<Vec<_>>(),
        );
        let parsed = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();
        assert_eq!(parsed.units.len(), 31);
        assert_eq!(
            parsed
                .objects
                .iter()
                .filter(|object| object.unit_index == 0 && object.kind == ObjectKind::Image)
                .count(),
            1
        );
        assert!(parsed.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .message
                .starts_with("IWORK_CELL_ATTACHMENT_OMITTED:")
        }));
    }

    #[test]
    fn opens_supplied_pages_with_many_character_style_changes() {
        let parsed = detect_and_parse(
            include_bytes!("../../tests/fixtures/word-original.pages"),
            Limits::default(),
        )
        .expect("ordinary Pages character-style runs must not hit the metadata field limit")
        .expect("Pages package should be detected");
        assert_eq!(parsed.format, Some(DocumentFormat::Pages));
        assert!(parsed.units.len() > 1);
        assert!(
            parsed
                .objects
                .iter()
                .filter_map(|object| object.text.as_deref())
                .map(str::len)
                .sum::<usize>()
                > 30_000
        );
        assert!(parsed.objects.iter().all(|object| !matches!(
            object.source.locator,
            SourceLocator::Iwork {
                kind: "preview",
                ..
            }
        )));
        crate::protocol::encode(&parsed).expect("native Pages document should encode");
    }

    #[test]
    fn reads_iwork_attribute_tables_beyond_metadata_field_limit() {
        let styles = (0..1_965).map(|index| (index, 300)).collect::<Vec<_>>();
        let table = keynote_style_table(&styles);
        let changes =
            super::keynote_attribute_references(&table, DOCUMENT_COMPONENT, Limits::default())
                .expect("repeated style changes are content, not distinct metadata fields");
        assert_eq!(changes.len(), styles.len());
        assert_eq!(changes.last().unwrap().character_index, 1_964);
        assert!(changes.iter().all(|change| change.identifier == Some(300)));
    }

    #[test]
    fn bounds_iwork_attribute_references_and_rejects_malformed_entries() {
        let limits = Limits {
            max_relationship_edges: 2,
            ..Limits::default()
        };
        let table = keynote_style_table(&[(0, 300), (1, 301)]);
        assert_eq!(
            super::keynote_attribute_references(&table, DOCUMENT_COMPONENT, limits)
                .unwrap()
                .len(),
            2
        );
        let excessive = keynote_style_table(&[(0, 300), (1, 301), (2, 302)]);
        let error = super::keynote_attribute_references(&excessive, DOCUMENT_COMPONENT, limits)
            .err()
            .expect("configured reference budget must remain enforced");
        assert!(error.message.contains("configured reference limit"));
        for invalid in [
            keynote_style_table(&[(1, 300), (0, 301)]),
            keynote_style_table(&[(0, 300), (0, 301)]),
            table[..table.len() - 1].to_vec(),
        ] {
            assert!(
                super::keynote_attribute_references(&invalid, DOCUMENT_COMPONENT, limits).is_err()
            );
        }
        for field in [5, 7, 8, 9, 17] {
            let mut storage = protobuf_message(3, b"abc");
            storage.extend_from_slice(&protobuf_message(field, &excessive));
            assert!(super::keynote_storage_text(&storage, DOCUMENT_COMPONENT, limits).is_err());
        }
    }

    #[test]
    fn reads_pages_section_changes_from_text_storage() {
        let mut payload = protobuf_message(3, b"First section\nSecond section");
        payload.extend_from_slice(&protobuf_message(
            17,
            &keynote_style_table(&[(0, 40), (14, 41)]),
        ));

        let storage = super::keynote_storage_text(&payload, DOCUMENT_COMPONENT, Limits::default())
            .expect("Pages section table should parse");
        assert_eq!(
            storage
                .sections
                .iter()
                .map(|change| (change.character_index, change.identifier))
                .collect::<Vec<_>>(),
            [(0, Some(40)), (14, Some(41))]
        );
    }

    #[test]
    fn reads_pages_view_scale_from_saved_view_state() {
        assert_eq!(
            super::pages_view_scale(&protobuf_fixed32(15, 1.25), "Index/ViewState.iwa").unwrap(),
            1.25
        );
        assert_eq!(
            super::pages_view_scale(&[], "Index/ViewState.iwa").unwrap(),
            1.0
        );
    }

    #[test]
    fn inherits_iwork_nil_paragraph_styles_without_resetting_to_the_first() {
        let changes = [
            super::KeynoteStyleChange {
                character_index: 0,
                identifier: Some(10),
            },
            super::KeynoteStyleChange {
                character_index: 5,
                identifier: Some(20),
            },
            super::KeynoteStyleChange {
                character_index: 8,
                identifier: None,
            },
        ];

        assert_eq!(super::iwork_inherited_style_at(&changes, 6), Some(20));
        assert_eq!(super::iwork_inherited_style_at(&changes, 8), Some(20));
        assert_eq!(super::pages_style_at(&changes, 8), None);
    }

    #[test]
    fn measures_pages_text_with_each_paragraphs_own_line_height() {
        let storage = super::KeynoteTextStorage {
            text: "first\nsecond\n".to_owned(),
            ..Default::default()
        };
        let runs = [
            super::keynote_text_run(
                "first\n".to_owned(),
                &super::KeynoteTextStyle {
                    font_size: 36.0,
                    ..Default::default()
                },
            ),
            super::keynote_text_run(
                "second\n".to_owned(),
                &super::KeynoteTextStyle {
                    font_size: 10.0,
                    ..Default::default()
                },
            ),
        ];
        let paragraphs = super::pages_paragraph_layouts(
            &storage,
            &runs,
            &[],
            "Index/Document.iwa",
            0,
            storage.text.encode_utf16().count(),
            &super::NumbersMessageSpace {
                archives: Vec::new(),
            },
            Limits::default(),
            1.0,
            &super::FontMetricTable::default(),
            false,
        )
        .unwrap();

        assert!((paragraphs[0].line_height - 43.2).abs() < 0.001);
        assert!((paragraphs[1].line_height - 12.6).abs() < 0.001);
        assert!(
            (super::keynote_text_layout_height(&runs, &paragraphs, 1_000.0) - 55.8).abs() < 0.001
        );
    }

    #[test]
    fn renders_pages_table_grid_from_linked_model_dimensions_and_cells() {
        let mut model_payload = protobuf_varint(6, 4);
        model_payload.extend_from_slice(&protobuf_varint(7, 2));
        model_payload.extend_from_slice(&protobuf_varint(9, 1));
        model_payload.extend_from_slice(&protobuf_varint(10, 1));
        model_payload.extend_from_slice(&protobuf_message(3, &protobuf_reference(501)));
        let model = super::IwaMessage {
            message_type: 6_001,
            payload: model_payload,
            data_references: Vec::new(),
        };
        let mut table_properties = protobuf_varint(42, 1);
        table_properties.extend_from_slice(&protobuf_varint(43, 1));
        let table_style = super::IwaArchive {
            identifier: 501,
            messages: vec![super::IwaMessage {
                message_type: 6_003,
                payload: protobuf_message(11, &table_properties),
                data_references: Vec::new(),
            }],
        };
        let cells = [
            (0, 0, Some("Label 1")),
            (0, 1, Some("Value 1")),
            (1, 0, Some("2-")),
            (1, 1, Some("Value 2")),
            (2, 0, None),
            (2, 1, Some("Value 3")),
            (3, 0, Some("Label 4")),
            (3, 1, Some("\u{7f}\u{008d}\u{008e}\u{008f}")),
        ]
        .map(|(row, column, text)| super::NumbersCell {
            row,
            column,
            text: text.map(str::to_owned),
            text_storage: None,
            span: false,
            part: "tile".to_owned(),
            style_index: None,
            text_style_index: None,
        });
        let mut objects = Vec::new();
        super::pages_push_table_grid(
            &mut objects,
            Limits::default(),
            0,
            100,
            Rect {
                x: 20.0,
                y: 30.0,
                width: 300.0,
                height: 80.0,
            },
            &model,
            DOCUMENT_COMPONENT,
            4,
            &[100.0, 200.0],
            &[20.0; 4],
            &cells,
            std::slice::from_ref(&table_style),
            &[],
            &[],
            None,
            &[],
            None,
            0,
        )
        .expect("Pages table grid should render");

        let table = objects
            .iter()
            .find(|object| object.kind == ObjectKind::Table)
            .expect("Pages table object");
        let Visual::StrokeStyle { style, visual } = &table.visual else {
            panic!("Pages table border should be inset");
        };
        assert_eq!(style.alignment, LineAlignment::Inset);
        let Visual::PaintedShape {
            stroke,
            stroke_width,
            ..
        } = visual.as_ref()
        else {
            panic!("Pages table should paint its outer border");
        };
        assert_eq!(*stroke, Paint::Solid(0x0000_00ff));
        assert_eq!(*stroke_width, 2.0);
        assert_eq!(
            objects
                .iter()
                .filter(|object| object.kind == ObjectKind::Cell)
                .count(),
            7
        );
        assert_eq!(
            objects
                .iter()
                .filter_map(|object| object.text.as_deref())
                .collect::<Vec<_>>(),
            cells
                .iter()
                .filter_map(|cell| cell.text.as_deref())
                .collect::<Vec<_>>()
        );
        let first = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Label 1"))
            .expect("first header cell");
        let second = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Value 1"))
            .expect("first value cell");
        assert_eq!((first.bounds.width, second.bounds.width), (92.0, 192.0));
        let Visual::RichText { align, runs, .. } = &first.visual else {
            panic!("Pages cell should use rich text");
        };
        assert_eq!(*align, TextAlign::Center);
        assert!(runs[0].bold);
        let Visual::RichText { align, runs, .. } = &second.visual else {
            panic!("Pages cell should use rich text");
        };
        assert_eq!(*align, TextAlign::Center);
        assert!(runs[0].bold);

        for identifier in [101, 200] {
            let line = objects
                .iter()
                .find(|object| object.stable_id == format!("iwork-pages-0-table-grid-{identifier}"))
                .expect("header divider");
            let Visual::PaintedShape {
                stroke,
                stroke_width,
                ..
            } = &line.visual
            else {
                panic!("header divider should be painted");
            };
            assert_eq!(*stroke, Paint::Solid(0x0000_00ff));
            assert_eq!(*stroke_width, 2.0);
        }
        assert!(
            objects
                .iter()
                .all(|object| { object.stable_id != "iwork-pages-0-table-grid-201" })
        );
        let group_label = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("2-"))
            .expect("left group label");
        assert_eq!(
            (group_label.bounds.y, group_label.bounds.height),
            (50.0, 40.0)
        );
        let Visual::TextLayout { layout, .. } = &group_label.visual else {
            panic!("left group label should use a centered layout");
        };
        assert_eq!(layout.vertical_align, TextVerticalAlign::Center);
        assert!(!layout.wrap);

        let control_text = "\u{7f}\u{008d}\u{008e}\u{008f}";
        let control_cell = objects
            .iter()
            .find(|object| object.text.as_deref() == Some(control_text))
            .expect("control-character cell");
        let Visual::RichText { runs, .. } = &control_cell.visual else {
            panic!("control-character cell should use rich text");
        };
        assert_eq!(runs[0].text, "");
    }

    #[test]
    fn expands_pages_pdf_attachments_into_native_vector_objects() {
        let pdf = b"%PDF-1.7
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 120 80] >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /Resources << >> /Contents 4 0 R >> endobj
4 0 obj << /Length 25 >> stream
0 0 1 rg 10 20 40 30 re f
endstream endobj
trailer << /Root 1 0 R >>
%%EOF";
        let mut objects = Vec::new();
        let mut embedded_fonts = Vec::new();
        assert!(
            super::iwork_push_pdf_objects(
                "Data/chart.pdf",
                pdf,
                42,
                0,
                Rect {
                    x: 10.0,
                    y: 20.0,
                    width: 120.0,
                    height: 80.0,
                },
                50,
                "chart",
                1.0,
                false,
                Limits::default(),
                &mut objects,
                &mut embedded_fonts,
            )
            .expect("Pages vector PDF should parse")
        );
        assert!(!objects.is_empty());
        assert!(objects.iter().all(|object| {
            object.kind != ObjectKind::Image
                && matches!(
                    object.source.locator,
                    SourceLocator::Iwork { kind: "chart", .. }
                )
                && matches!(object.visual, Visual::Layer { .. })
        }));
    }

    #[test]
    fn namespaces_same_named_pdf_font_subsets_per_iwork_asset() {
        let authored = "SubsetFont PDF 9 0";
        let first = super::iwork_pdf_font_family("Data/first.pdf", 42, authored);
        let second = super::iwork_pdf_font_family("Data/second.pdf", 43, authored);

        assert_ne!(first, second);
        assert!(first.ends_with(authored));
        assert!(second.ends_with(authored));

        let mut visual = Visual::TextLayout {
            layout: TextLayout::default(),
            visual: Box::new(Visual::RichText {
                geometry: Geometry::Rectangle,
                fill: Paint::None,
                stroke: Paint::None,
                stroke_width: 0.0,
                align: TextAlign::Start,
                line_height: 12.0,
                runs: vec![TextRun {
                    paint: None,
                    east_asian_line_breaks: true,
                    text: "Device".to_owned(),
                    font_family: authored.to_owned(),
                    font_size: 12.0,
                    color: u32::MAX,
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
        };
        super::rename_iwork_pdf_visual_fonts(
            &mut visual,
            &HashMap::from([(authored.to_owned(), first.clone())]),
        );
        let Visual::TextLayout { visual, .. } = visual else {
            panic!("PDF text should retain its text layout");
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("PDF text should retain its rich text");
        };
        assert_eq!(runs[0].font_family, first);
    }

    #[test]
    fn rejects_non_native_numbers_and_identifies_cross_version_keynote_markers() {
        let numbers = iwa_numbers_document_with_sheets(3);
        let preview = jpeg(3, 2);
        let numbers_package = stored_zip(&[
            ("Index/Document.iwa", &numbers),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);
        let error = detect_and_parse(&numbers_package, Limits::default()).unwrap_err();
        assert!(
            error
                .message
                .contains("Numbers sheet reference 2 is missing")
        );

        for marker in [
            "Index/TemplateSlide-42.iwa",
            "Index/MasterSlide-1.iwa",
            "Index/Slide1.iwa",
        ] {
            let keynote = iwa_keynote_document();
            let marker_iwa = iwa_document(6);
            let package = stored_zip(&[
                ("Index/Document.iwa", &keynote),
                (marker, &marker_iwa),
                ("Metadata/Properties.plist", PROPERTIES),
                ("preview.jpg", &preview),
            ]);
            let document = detect_and_parse(&package, Limits::default())
                .unwrap()
                .unwrap();
            assert_eq!(document.format, Some(DocumentFormat::Keynote));
        }
    }

    #[test]
    fn exposes_embedded_keynote_slide_previews_in_native_order() {
        let document = keynote_document_with_slide_previews();
        let metadata = iwa_stream(&[iwa_archive(
            2,
            11_006,
            &package_metadata(&[(201, "st-slide-two.png"), (202, "st-slide-three.png")]),
            &[],
        )]);
        let marker = iwa_document(5);
        let root_preview = png(8, 5);
        let second_preview = png(4, 3);
        let third_preview = png(6, 4);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-101.iwa", &marker),
            ("Index/Metadata.iwa", &metadata),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.png", &root_preview),
            ("Data/st-slide-two.png", &second_preview),
            ("Data/st-slide-three.png", &third_preview),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();
        assert_eq!(document.format, Some(DocumentFormat::Keynote));
        assert_eq!(document.units.len(), 3);
        assert_eq!(document.objects.len(), 3);
        assert!(document.units.iter().all(|unit| {
            (unit.width - 21.333_334).abs() < 0.001 && (unit.height - 12.0).abs() < 0.001
        }));
        assert_eq!(
            document
                .objects
                .iter()
                .map(|object| object.source.part.as_str())
                .collect::<Vec<_>>(),
            [
                "preview.png",
                "Data/st-slide-two.png",
                "Data/st-slide-three.png",
            ]
        );
        assert!(document.objects.iter().all(|object| matches!(
            object.source.locator,
            SourceLocator::Iwork {
                kind: "slide-preview",
                ..
            }
        )));
        assert!(
            document.diagnostics[0]
                .message
                .starts_with("IWORK_KEYNOTE_PREVIEW_ONLY:")
        );
        crate::protocol::encode(&document).expect("multi-slide Keynote preview should encode");
    }

    #[test]
    fn keynote_native_object_ids_are_unique_across_slides() {
        let mut objects = Vec::new();
        for unit_index in [0, 1] {
            push_keynote_native_object(
                &mut objects,
                Limits::default(),
                42,
                unit_index,
                "Index/DocumentStylesheet.iwa",
                "slide-background",
                ObjectKind::Shape,
                Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 16.0,
                    height: 9.0,
                },
                i32::MIN,
                None,
                MappingQuality::Exact,
                Visual::Shape {
                    geometry: Geometry::Rectangle,
                    fill: 0xff,
                    stroke: 0,
                    stroke_width: 0.0,
                },
            )
            .unwrap();
        }

        assert_eq!(objects[0].numeric_id, 1);
        assert_eq!(objects[1].numeric_id, 2);
        assert_ne!(objects[0].stable_id, objects[1].stable_id);
    }

    #[test]
    fn renders_keynote_text_placeholders_as_native_objects_instead_of_upscaled_previews() {
        let document = keynote_document_with_native_slide(200);
        let legacy_document = keynote_document_with_legacy_slide();
        let slide = keynote_slide_with_text_placeholder("Sharp native title");
        let preview = jpeg(1_024, 768);
        for (document, slide_part) in [
            (&document, "Index/Slide-200.iwa"),
            (&legacy_document, "Index/Slide1.iwa"),
        ] {
            let bytes = stored_zip(&[
                ("Index/Document.iwa", &document),
                (slide_part, &slide),
                ("Metadata/Properties.plist", PROPERTIES),
                ("preview.jpg", &preview),
            ]);

            let document = detect_and_parse(&bytes, Limits::default())
                .unwrap()
                .unwrap();

            assert_eq!(document.objects.len(), 1);
            assert_eq!(document.objects[0].kind, ObjectKind::TextBox);
            assert_eq!(
                document.objects[0].text.as_deref(),
                Some("Sharp native title")
            );
            assert!(matches!(
                document.objects[0].source.locator,
                SourceLocator::Iwork {
                    kind: "text-box",
                    ..
                }
            ));
            assert!(document.diagnostics.iter().any(|diagnostic| {
                diagnostic
                    .message
                    .starts_with("IWORK_KEYNOTE_NATIVE_STATIC_SUBSET:")
            }));
        }
    }

    #[test]
    fn exposes_embedded_keynote_media_with_its_authored_poster() {
        const MP3: &[u8] = b"ID3\x04\0\0";
        let document = keynote_document_with_native_slide(200);
        let mut slide = protobuf_message(7, &protobuf_reference(201));
        slide.extend_from_slice(&protobuf_message(42, &protobuf_reference(201)));
        let media = protobuf_message(1, &keynote_drawable(80.0, 90.0, 480.0, 270.0));
        let slide = iwa_stream(&[
            iwa_archive(200, 5, &slide, &[201]),
            iwa_archive(201, super::IWORK_MEDIA_TYPE, &media, &[501, 502]),
        ]);
        let metadata = iwa_stream(&[iwa_archive(
            1,
            super::IWORK_PACKAGE_METADATA_TYPE,
            &package_metadata(&[(501, "clip.mp3"), (502, "poster.jpg")]),
            &[],
        )]);
        let preview = jpeg(1_024, 768);
        let poster = jpeg(480, 270);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-200.iwa", &slide),
            ("Index/Metadata.iwa", &metadata),
            ("Metadata/Properties.plist", PROPERTIES),
            ("Data/clip.mp3", MP3),
            ("Data/poster.jpg", &poster),
            ("preview.jpg", &preview),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();
        assert!(document.objects.iter().any(|object| matches!(
            &object.visual,
            Visual::Media {
                kind: crate::model::MediaKind::Audio,
                media_type,
                bytes,
                poster,
            } if media_type == "audio/mpeg"
                && bytes == MP3
                && matches!(poster.as_ref(), Visual::Image { media_type, .. } if media_type == "image/jpeg")
        )));
    }

    #[test]
    fn preserves_keynote_character_styles_in_native_placeholders() {
        let document = keynote_document_with_native_slide(200);
        let slide = keynote_slide_with_rich_text_placeholder();
        let stylesheet = keynote_rich_text_stylesheet();
        let preview = jpeg(1_024, 768);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-200.iwa", &slide),
            ("Index/DocumentStylesheet.iwa", &stylesheet),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();

        let Visual::RichText { align, runs, .. } = &document.objects[0].visual else {
            panic!("Keynote styled text should use native rich-text runs");
        };
        assert_eq!(*align, TextAlign::Center);
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Base ");
        assert_eq!(runs[0].font_family, "Helvetica");
        assert_eq!(runs[0].font_size, 32.0);
        assert_eq!(runs[0].color, 0xff00_00ff);
        assert!(!runs[0].bold);
        assert!(!runs[0].italic);
        assert_eq!(runs[1].text, "Accent");
        assert_eq!(runs[1].font_family, "Courier");
        assert!((runs[1].font_size - 42.666_668).abs() < 0.001);
        assert_eq!(runs[1].color, 0x0066_ffff);
        assert!(runs[1].bold);
        assert!(runs[1].italic);
        assert!(runs[1].underline);
        assert!(runs[1].strikethrough);
        assert_eq!(runs[1].highlight, 0xffff_00ff);
        assert_eq!(runs[1].baseline_shift, 4.0);
        assert_eq!(runs[1].letter_spacing, 2.0);
    }

    #[test]
    fn normalizes_keynote_postscript_symbol_runs_before_browser_rendering() {
        let document = keynote_document_with_native_slide(200);
        let mut slide_payload = protobuf_message(7, &protobuf_reference(201));
        slide_payload.extend_from_slice(&protobuf_message(42, &protobuf_reference(201)));
        let shape = protobuf_message(1, &keynote_drawable(100.0, 120.0, 824.0, 260.0));
        let mut shape_info = protobuf_message(1, &shape);
        shape_info.extend_from_slice(&protobuf_message(4, &protobuf_reference(202)));
        let placeholder = protobuf_message(1, &shape_info);
        let mut storage = protobuf_message(3, b"p");
        storage.extend_from_slice(&protobuf_message(5, &keynote_style_table(&[(0, 300)])));
        let slide = iwa_stream(&[
            iwa_archive(200, 5, &slide_payload, &[201]),
            iwa_archive(201, 7, &placeholder, &[202]),
            iwa_archive(202, 2_001, &storage, &[]),
        ]);
        let paragraph = protobuf_message(
            11,
            &keynote_character_properties("Wingdings-Regular", 18.0, 0x0000_00ff, false, false),
        );
        let stylesheet = iwa_stream(&[iwa_archive(300, 2_022, &paragraph, &[])]);
        let preview = jpeg(1_024, 768);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-200.iwa", &slide),
            ("Index/DocumentStylesheet.iwa", &stylesheet),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();
        assert_eq!(document.objects[0].text.as_deref(), Some("□"));
        let Visual::RichText { runs, .. } = &document.objects[0].visual else {
            panic!("Keynote symbol text should remain a native rich-text run");
        };
        assert_eq!(runs[0].text, "p");
        assert_eq!(runs[0].font_family, "Wingdings");
        assert!(!runs[0].bold);
    }

    #[test]
    fn reads_iwork_tracking_in_em_without_treating_kerning_as_tracking() {
        let mut style = super::KeynoteTextStyle::default();
        super::keynote_character_style(&protobuf_fixed32(15, 2.0), DOCUMENT_COMPONENT, &mut style)
            .unwrap();
        assert_eq!(
            super::keynote_text_run("AB".to_owned(), &style).letter_spacing,
            0.0
        );
        let mut properties = protobuf_fixed32(27, 0.125);
        properties.extend_from_slice(&protobuf_fixed32(3, 16.0));
        super::keynote_character_style(&properties, DOCUMENT_COMPONENT, &mut style).unwrap();
        assert_eq!(
            super::keynote_text_run("AB".to_owned(), &style).letter_spacing,
            2.0
        );
        super::keynote_character_style(&protobuf_fixed32(3, 32.0), DOCUMENT_COMPONENT, &mut style)
            .unwrap();
        assert_eq!(
            super::keynote_text_run("AB".to_owned(), &style).letter_spacing,
            4.0
        );
    }

    #[test]
    fn normalizes_keynote_font_aliases_and_paragraph_line_spacing() {
        assert_eq!(super::keynote_font_family("KaiTi"), "PingFangSC-Regular");
        assert_eq!(super::keynote_font_family("Wingdings-Regular"), "Wingdings");
        assert_eq!(super::keynote_font_family("SymbolMT"), "Symbol");
        assert_eq!(super::keynote_font_family("Helvetica"), "Helvetica");

        let mut properties = protobuf_varint(1, 2);
        properties.extend_from_slice(&protobuf_varint(27, u64::from(u32::MAX)));
        properties.extend_from_slice(&protobuf_message(13, &protobuf_fixed32(2, 1.5)));
        let mut style = super::KeynoteTextStyle::default();
        super::keynote_paragraph_properties(&properties, super::STYLESHEET_COMPONENT, &mut style)
            .unwrap();

        assert_eq!(style.align, TextAlign::Center);
        assert_eq!(style.list_level, 0);
        assert_eq!(style.line_height_multiple, Some(1.5));

        let mut exact_spacing = protobuf_varint(1, 2);
        exact_spacing.extend_from_slice(&protobuf_fixed32(2, 14.0));
        let exact_properties = protobuf_message(13, &exact_spacing);
        let mut exact_style = super::KeynoteTextStyle::default();
        super::keynote_paragraph_properties(
            &exact_properties,
            super::STYLESHEET_COMPONENT,
            &mut exact_style,
        )
        .unwrap();
        assert_eq!(exact_style.line_height_multiple, Some(14.0));

        let authored_layouts = super::keynote_text_paragraph_layouts(
            &[TextRun {
                paint: None,
                east_asian_line_breaks: true,
                text: "Large\n".to_owned(),
                font_family: "Arial".to_owned(),
                font_size: 124.0,
                color: 0,
                bold: false,
                italic: false,
                underline: false,
                strikethrough: false,
                highlight: 0,
                baseline_shift: 0.0,
                letter_spacing: 0.0,
                horizontal_scale: 1.0,
            }],
            TextAlign::Start,
            100.0,
            true,
            &[],
        );
        assert_eq!(authored_layouts[0].line_height, 100.0);

        let mut chinese = super::KeynoteTextStyle {
            font_family: "PingFangSC-Regular".to_owned(),
            font_size: 48.0,
            ..Default::default()
        };
        let fallback_runs = super::keynote_platform_font_runs(vec![super::keynote_text_run(
            "中20文".to_owned(),
            &chinese,
        )]);
        assert_eq!(
            fallback_runs
                .iter()
                .map(|run| (run.text.as_str(), run.font_family.as_str()))
                .collect::<Vec<_>>(),
            [
                ("中", "PingFangSC-Regular"),
                ("20", "Helvetica"),
                ("文", "PingFangSC-Regular"),
            ]
        );
        chinese.font_family = "Calibri".to_owned();
        let keynote_runs = super::keynote_platform_font_runs(vec![super::keynote_text_run(
            "密码 AES？".to_owned(),
            &chinese,
        )]);
        assert_eq!(
            keynote_runs
                .iter()
                .map(|run| (run.text.as_str(), run.font_family.as_str()))
                .collect::<Vec<_>>(),
            [
                ("密码", "PingFangSC-Regular"),
                (" AES", "Helvetica"),
                ("？", "PingFangSC-Regular"),
            ]
        );
        chinese.bold = true;
        let semibold = super::keynote_platform_font_runs(vec![super::keynote_text_run(
            "中文".to_owned(),
            &chinese,
        )]);
        assert_eq!(semibold[0].font_family, "PingFangSC-Semibold");
        assert!(!semibold[0].bold);
        let helvetica_cjk = super::keynote_platform_font_runs(vec![super::keynote_text_run(
            "标题 Bitcoin".to_owned(),
            &super::KeynoteTextStyle {
                font_family: "HelveticaNeue-Bold".to_owned(),
                bold: true,
                ..Default::default()
            },
        )]);
        assert_eq!(helvetica_cjk[0].font_family, "PingFangSC-Semibold");
        assert!(!helvetica_cjk[0].bold);
        assert_eq!(helvetica_cjk[1].font_family, "HelveticaNeue-Bold");
        let mut indented = vec![super::keynote_text_run(
            "    ".to_owned(),
            &super::KeynoteTextStyle {
                font_family: "Helvetica".to_owned(),
                font_size: 48.0,
                ..Default::default()
            },
        )];
        let indent = super::keynote_leading_space_indent(&mut indented);
        assert!((indent - 53.376).abs() < 0.001);
        assert!((indented[0].letter_spacing + 13.344).abs() < 0.001);
        chinese.font_family = "PingFangSC-Regular".to_owned();
        chinese.bold = false;
        let chinese_run = super::keynote_text_run("中文\n".to_owned(), &chinese);
        chinese.font_family = "Times-Roman".to_owned();
        let latin_run = super::keynote_text_run("English".to_owned(), &chinese);
        let layouts = super::keynote_text_paragraph_layouts(
            &[chinese_run, latin_run],
            TextAlign::Center,
            67.2,
            false,
            &[],
        );
        assert_eq!(layouts.len(), 2);
        assert!((layouts[0].line_height - 67.2).abs() < 0.001);
        assert!((layouts[1].line_height - 57.6).abs() < 0.001);
        assert_eq!(layouts[1].space_before, 3.0);

        let authored_layouts = super::keynote_text_paragraph_layouts(
            &[
                super::keynote_text_run("first\n".to_owned(), &chinese),
                super::keynote_text_run("second".to_owned(), &chinese),
            ],
            TextAlign::Start,
            67.2,
            true,
            &[],
        );
        assert_eq!(authored_layouts.len(), 2);
        assert!(
            authored_layouts
                .iter()
                .all(|layout| (layout.line_height - 67.2).abs() < 0.001)
        );
    }

    #[test]
    fn reads_iwork_packed_cell_and_text_style_indices() {
        let mut empty = vec![5, 0, 0, 0, 0, 0, 0, 0];
        empty.extend_from_slice(&0x20_u32.to_le_bytes());
        empty.extend_from_slice(&10_u32.to_le_bytes());
        assert_eq!(
            super::numbers_cell_style_indexes(&empty, "Index/Tables/Tile.iwa").unwrap(),
            (Some(10), None)
        );

        let mut text = vec![5, 3, 0, 0, 0, 0, 0, 0];
        text.extend_from_slice(&0x21068_u32.to_le_bytes());
        text.extend_from_slice(&11_u32.to_le_bytes());
        text.extend_from_slice(&8_u32.to_le_bytes());
        text.extend_from_slice(&9_u32.to_le_bytes());
        assert_eq!(
            super::numbers_cell_style_indexes(&text, "Index/Tables/Tile.iwa").unwrap(),
            (Some(8), Some(9))
        );
    }

    #[test]
    fn renders_keynote_regular_polygon_paths_at_authored_bounds() {
        let Geometry::Path { commands, .. } = super::keynote_regular_polygon(3.0, 300.0, 100.0)
        else {
            panic!("regular polygon should produce a native path");
        };
        let [
            crate::model::PathCommand::MoveTo { x: top_x, y: top_y },
            crate::model::PathCommand::LineTo {
                x: right_x,
                y: right_y,
            },
            crate::model::PathCommand::LineTo {
                x: left_x,
                y: left_y,
            },
            crate::model::PathCommand::ClosePath,
        ] = commands.as_slice()
        else {
            panic!("triangle should contain three vertices and a close command");
        };
        for (actual, expected) in [
            (*top_x, 150.0),
            (*top_y, 0.0),
            (*right_x, 279.903_8),
            (*right_y, 75.0),
            (*left_x, 20.096_19),
            (*left_y, 75.0),
        ] {
            assert!((actual - expected).abs() < 0.001);
        }

        let star = protobuf_varint(1, 100);
        let Geometry::Path { commands, .. } =
            super::keynote_point_path(&star, 300.0, 300.0, "Index/Slide.iwa").unwrap()
        else {
            panic!("five-point star should produce a native path");
        };
        assert_eq!(commands.len(), 11);

        let mut burst = protobuf_varint(1, 100);
        let mut parameters = protobuf_fixed32(1, 16.0);
        parameters.extend_from_slice(&protobuf_fixed32(2, 0.75));
        burst.extend_from_slice(&protobuf_message(2, &parameters));
        let Geometry::Path { commands, .. } =
            super::keynote_point_path(&burst, 84.0, 72.0, "Index/Slide.iwa").unwrap()
        else {
            panic!("Keynote burst should produce a native path");
        };
        assert_eq!(commands.len(), 33);
    }

    #[test]
    fn aligns_small_centered_group_text_to_its_single_shape() {
        let shape_bounds = Rect {
            x: 100.0,
            y: 100.0,
            width: 120.0,
            height: 100.0,
        };
        let mut objects = Vec::new();
        push_keynote_native_object(
            &mut objects,
            Limits::default(),
            1,
            0,
            DOCUMENT_COMPONENT,
            "shape",
            ObjectKind::Shape,
            shape_bounds,
            1,
            None,
            MappingQuality::Exact,
            Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: Paint::Solid(0x4472_c4ff),
                stroke: Paint::None,
                stroke_width: 0.0,
            },
        )
        .unwrap();
        push_keynote_native_object(
            &mut objects,
            Limits::default(),
            2,
            0,
            DOCUMENT_COMPONENT,
            "text-box",
            ObjectKind::TextBox,
            Rect {
                x: 120.0,
                y: 172.500_02,
                width: 50.0,
                height: 55.0,
            },
            2,
            Some("文字内容测试".to_owned()),
            MappingQuality::Exact,
            Visual::TextLayout {
                layout: crate::model::TextLayout {
                    vertical_align: TextVerticalAlign::Center,
                    ..crate::model::TextLayout::default()
                },
                visual: Box::new(Visual::RichText {
                    geometry: Geometry::Rectangle,
                    fill: Paint::None,
                    stroke: Paint::None,
                    stroke_width: 0.0,
                    align: TextAlign::Center,
                    line_height: 18.0,
                    runs: Vec::new(),
                }),
            },
        )
        .unwrap();

        super::keynote_align_group_text_to_shapes(&mut objects, 0);
        assert_eq!(
            objects[1].bounds,
            Rect {
                x: 120.0,
                y: shape_bounds.y,
                width: 50.0,
                height: shape_bounds.height,
            }
        );

        let first_bounds = Rect {
            x: 120.0,
            y: 172.500_02,
            width: 50.0,
            height: 55.0,
        };
        objects[1].bounds = first_bounds;
        let mut second = objects[1].clone();
        second.numeric_id = 3;
        second.stable_id.push_str("-second");
        second.bounds.y = 165.0;
        let second_bounds = second.bounds;
        objects.push(second);

        super::keynote_align_group_text_to_shapes(&mut objects, 0);
        assert_eq!(objects[1].bounds, first_bounds);
        assert_eq!(objects[2].bounds, second_bounds);
    }

    #[test]
    fn keynote_auto_width_ignores_the_derived_shape_width() {
        let style = super::KeynoteTextStyle {
            font_size: 100.0,
            align: TextAlign::Center,
            ..Default::default()
        };
        let runs = vec![super::keynote_text_run(
            "Company Introduction".to_owned(),
            &style,
        )];
        let bounds = super::keynote_text_bounds(
            super::KeynoteGeometry {
                x: 1_280.0,
                y: 66.5,
                width: 96.984,
                height: 63.656,
                rotation_degrees: 0.0,
                ..Default::default()
            },
            (0.0, 0.0),
            (2_560.0, 1_440.0),
            &runs,
            TextAlign::Center,
            120.0,
            0.0,
            true,
        );

        assert!(bounds.width > 900.0, "{bounds:?}");
        assert!((bounds.x + bounds.width / 2.0 - 1_280.0).abs() < 0.001);
    }

    #[test]
    fn keynote_path_derived_auto_height_keeps_its_authored_top_coordinate() {
        let style = super::KeynoteTextStyle {
            font_size: 18.0,
            ..Default::default()
        };
        let runs = vec![super::keynote_text_run("First\nSecond".to_owned(), &style)];
        let bounds = super::keynote_text_bounds(
            super::KeynoteGeometry {
                x: 351.6,
                y: 366.5,
                width: 252.5,
                height: 138.1,
                rotation_degrees: 0.0,
                ..Default::default()
            },
            (0.0, 0.0),
            (960.0, 540.0),
            &runs,
            TextAlign::Start,
            25.2,
            0.0,
            false,
        );

        assert!((bounds.y - 366.5).abs() < 0.001, "{bounds:?}");
    }

    #[test]
    fn inherits_and_overrides_keynote_table_cell_fills() {
        let cell_style = |identifier, base, color| {
            let mut payload = Vec::new();
            if let Some(base) = base {
                payload.extend_from_slice(&protobuf_message(
                    1,
                    &protobuf_message(3, &protobuf_reference(base)),
                ));
            }
            if let Some(color) = color {
                let fill = protobuf_message(1, &keynote_color_message(color));
                payload.extend_from_slice(&protobuf_message(11, &protobuf_message(1, &fill)));
            }
            super::IwaArchive {
                identifier,
                messages: vec![super::IwaMessage {
                    message_type: 6_004,
                    payload,
                    data_references: Vec::new(),
                }],
            }
        };
        let archives = [
            cell_style(245, None, Some(0x70ad_47ff)),
            cell_style(1_990, Some(245), None),
            cell_style(1_993, Some(245), Some(0xed7d_31ff)),
        ];
        assert_eq!(
            super::keynote_table_cell_fill(
                &archives,
                1_990,
                "Index/DocumentStylesheet.iwa",
                100.0,
                20.0,
                &mut Vec::new(),
            )
            .unwrap(),
            Paint::Solid(0x70ad_47ff)
        );
        assert_eq!(
            super::keynote_table_cell_fill(
                &archives,
                1_993,
                "Index/DocumentStylesheet.iwa",
                100.0,
                20.0,
                &mut Vec::new(),
            )
            .unwrap(),
            Paint::Solid(0xed7d_31ff)
        );

        let banded_fill = protobuf_message(1, &keynote_color_message(0xedf0_ebff));
        let mut table_properties = protobuf_varint(1, 1);
        table_properties.extend_from_slice(&protobuf_message(2, &banded_fill));
        let table_style = super::IwaArchive {
            identifier: 501,
            messages: vec![super::IwaMessage {
                message_type: 6_003,
                payload: protobuf_message(11, &table_properties),
                data_references: Vec::new(),
            }],
        };
        let mut model_payload = protobuf_message(18, &protobuf_reference(245));
        model_payload.extend_from_slice(&protobuf_message(3, &protobuf_reference(501)));
        let model = super::IwaMessage {
            message_type: super::NUMBERS_TABLE_MODEL_TYPE,
            payload: model_payload,
            data_references: Vec::new(),
        };
        let mut banded_archives = archives.to_vec();
        banded_archives.push(table_style);
        assert_eq!(
            super::keynote_table_banded_fills(
                &banded_archives,
                &model,
                "Index/DocumentStylesheet.iwa",
                100.0,
                20.0,
            )
            .unwrap(),
            Some((Paint::Solid(0x70ad_47ff), Paint::Solid(0xedf0_ebff),))
        );
    }

    #[test]
    fn preserves_keynote_stroke_dash_cap_and_join() {
        let mut stroke = protobuf_message(1, &keynote_color_message(0x1122_33ff));
        stroke.extend_from_slice(&protobuf_fixed32(2, 3.0));
        stroke.extend_from_slice(&protobuf_varint(3, 1));
        stroke.extend_from_slice(&protobuf_varint(4, 2));
        let mut pattern = protobuf_varint(1, 0);
        pattern.extend_from_slice(&protobuf_fixed32(4, 8.0));
        pattern.extend_from_slice(&protobuf_fixed32(4, 4.0));
        stroke.extend_from_slice(&protobuf_message(6, &pattern));

        let (paint, width, style) =
            super::keynote_stroke(&stroke, "Index/DocumentStylesheet.iwa").unwrap();
        assert_eq!(paint, Paint::Solid(0x1122_33ff));
        assert_eq!(width, 3.0);
        assert_eq!(style.cap, crate::model::LineCap::Round);
        assert_eq!(style.join, crate::model::LineJoin::Bevel);
        assert_eq!(style.dash, [24.0, 12.0]);
    }

    #[test]
    fn resolves_keynote_text_styles_from_theme_stylesheet() {
        let document = keynote_document_with_native_slide(200);
        let slide = keynote_slide_with_rich_text_placeholder();
        let stylesheet = keynote_rich_text_stylesheet();
        let preview = jpeg(1_024, 768);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-200.iwa", &slide),
            ("Index/ThemeStylesheet.iwa", &stylesheet),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();

        let Visual::RichText { runs, .. } = &document.objects[0].visual else {
            panic!("Keynote theme text should resolve to native rich-text runs");
        };
        assert_eq!(runs[0].font_family, "Helvetica");
        assert_eq!(runs[1].font_family, "Courier");
        assert_eq!(runs[1].color, 0x0066_ffff);
    }

    #[test]
    fn preserves_slide_images_and_nonplaceholder_master_images() {
        let limits = Limits::default();
        let bytes = stored_zip(&[("Data/image.png", &png(32, 32))]);
        let package = super::Package::open_iwork(&bytes, limits).unwrap();
        for render_placeholder_content in [false, true] {
            for flags in [None, Some(0), Some(1), Some(2), Some(3)] {
                let mut payload = protobuf_message(1, &keynote_drawable(10.0, 20.0, 32.0, 32.0));
                payload.extend(protobuf_message(11, &protobuf_reference(100)));
                if let Some(flags) = flags {
                    payload.extend(protobuf_varint(7, flags));
                }
                let archives = vec![super::IwaArchive {
                    identifier: 1,
                    messages: vec![super::IwaMessage {
                        message_type: super::IWORK_IMAGE_TYPE,
                        payload,
                        data_references: vec![100],
                    }],
                }];
                let mut objects = Vec::new();
                let mut path = Vec::new();
                let omitted = super::keynote_native_drawable(
                    &mut Vec::new(),
                    &package,
                    &archives,
                    &[],
                    &super::NumbersMessageSpace {
                        archives: Vec::new(),
                    },
                    &[(100, "Data/image.png".to_owned())],
                    1,
                    0,
                    "Index/Slide.iwa",
                    limits,
                    &mut objects,
                    &mut path,
                    960.0,
                    720.0,
                    0.0,
                    0.0,
                    0,
                    render_placeholder_content,
                    &mut Vec::new(),
                    1.0,
                    None,
                )
                .unwrap();
                let visible = render_placeholder_content || flags.unwrap_or(0) & 1 == 0;
                assert_eq!(
                    objects.len(),
                    usize::from(visible),
                    "slide={render_placeholder_content}, flags={flags:?}"
                );
                assert_eq!(
                    omitted, 0,
                    "intentional master placeholder suppression is not a parse failure"
                );
                assert!(path.is_empty());
            }
        }
    }

    #[test]
    fn suppresses_master_placeholder_sample_text() {
        let document = keynote_document_with_native_slide(200);
        let mut slide_payload = protobuf_message(17, &protobuf_reference(300));
        slide_payload.extend_from_slice(&protobuf_message(7, &protobuf_reference(201)));
        slide_payload.extend_from_slice(&protobuf_message(42, &protobuf_reference(201)));
        let slide_shape = protobuf_message(1, &keynote_drawable(100.0, 80.0, 824.0, 120.0));
        let mut slide_shape_info = protobuf_message(1, &slide_shape);
        slide_shape_info.extend_from_slice(&protobuf_message(4, &protobuf_reference(202)));
        let slide = iwa_stream(&[
            iwa_archive(200, 5, &slide_payload, &[201]),
            iwa_archive(201, 7, &protobuf_message(1, &slide_shape_info), &[202]),
            iwa_archive(202, 2_001, &protobuf_message(3, b"Actual title"), &[]),
        ]);

        let mut master_payload = protobuf_message(7, &protobuf_reference(301));
        master_payload.extend_from_slice(&protobuf_message(42, &protobuf_reference(301)));
        let master_shape = protobuf_message(1, &keynote_drawable(100.0, 80.0, 824.0, 120.0));
        let mut master_shape_info = protobuf_message(1, &master_shape);
        master_shape_info.extend_from_slice(&protobuf_message(4, &protobuf_reference(302)));
        let master = iwa_stream(&[
            iwa_archive(300, 5, &master_payload, &[301]),
            iwa_archive(301, 7, &protobuf_message(1, &master_shape_info), &[302]),
            iwa_archive(302, 2_001, &protobuf_message(3, b"Slide Title"), &[]),
        ]);
        let preview = jpeg(1_024, 768);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-200.iwa", &slide),
            ("Index/MasterSlide-300.iwa", &master),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);

        let parsed = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();
        assert_eq!(parsed.objects.len(), 1);
        assert_eq!(parsed.objects[0].text.as_deref(), Some("Actual title"));
    }

    #[test]
    fn resolves_keynote_text_styles_from_identifier_suffixed_stylesheet() {
        let document = keynote_document_with_native_slide(200);
        let slide = keynote_slide_with_rich_text_placeholder();
        let stylesheet = keynote_rich_text_stylesheet();
        let preview = jpeg(1_024, 768);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-200.iwa", &slide),
            ("Index/DocumentStylesheet-4426.iwa", &stylesheet),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();

        let Visual::RichText { align, runs, .. } = &document.objects[0].visual else {
            panic!("Keynote suffixed stylesheet should resolve to native rich-text runs");
        };
        assert_eq!(*align, TextAlign::Center);
        assert_eq!(runs[0].font_family, "Helvetica");
        assert_eq!(runs[0].font_size, 32.0);
        assert_eq!(runs[1].font_family, "Courier");
        assert!((runs[1].font_size - 42.666_668).abs() < 0.001);
    }

    #[test]
    fn inherits_keynote_text_box_padding_and_vertical_alignment() {
        let document = keynote_document_with_native_slide(200);
        let mut slide_payload = protobuf_message(7, &protobuf_reference(201));
        slide_payload.extend_from_slice(&protobuf_message(42, &protobuf_reference(201)));
        let mut position = protobuf_fixed32(1, 100.0);
        position.extend_from_slice(&protobuf_fixed32(2, 80.0));
        let mut size = protobuf_fixed32(1, 824.0);
        size.extend_from_slice(&protobuf_fixed32(2, 360.0));
        let mut geometry = protobuf_message(1, &position);
        geometry.extend_from_slice(&protobuf_message(2, &size));
        geometry.extend_from_slice(&protobuf_varint(3, 1));
        let mut shape = protobuf_message(1, &protobuf_message(1, &geometry));
        shape.extend_from_slice(&protobuf_message(2, &protobuf_reference(401)));
        let mut shape_info = protobuf_message(1, &shape);
        shape_info.extend_from_slice(&protobuf_message(4, &protobuf_reference(202)));
        let placeholder = protobuf_message(1, &shape_info);
        let storage = protobuf_message(3, b"Bottom aligned title");
        let slide = iwa_stream(&[
            iwa_archive(200, 5, &slide_payload, &[201]),
            iwa_archive(201, 7, &placeholder, &[202]),
            iwa_archive(202, 2_001, &storage, &[]),
        ]);

        let mut padding = protobuf_fixed32(1, 12.0);
        padding.extend_from_slice(&protobuf_fixed32(2, 16.0));
        padding.extend_from_slice(&protobuf_fixed32(3, 20.0));
        padding.extend_from_slice(&protobuf_fixed32(4, 24.0));
        let parent_properties = protobuf_message(6, &padding);
        let mut parent_style = protobuf_message(1, &[]);
        parent_style.extend_from_slice(&protobuf_message(11, &parent_properties));

        let base = protobuf_message(3, &protobuf_reference(400));
        let tsd_style = protobuf_message(1, &base);
        let mut child_style = protobuf_message(1, &tsd_style);
        child_style.extend_from_slice(&protobuf_message(11, &protobuf_varint(2, 2)));
        let stylesheet = iwa_stream(&[
            iwa_archive(400, 2_025, &parent_style, &[]),
            iwa_archive(401, 2_025, &child_style, &[]),
        ]);
        let preview = jpeg(1_024, 768);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-200.iwa", &slide),
            ("Index/DocumentStylesheet-4426.iwa", &stylesheet),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();
        let Visual::TextLayout { layout, visual } = &document.objects[0].visual else {
            panic!("Keynote text-box style should use the shared text-layout model");
        };
        assert_eq!(layout.vertical_align, TextVerticalAlign::Bottom);
        assert_eq!(document.objects[0].bounds.y, 0.0);
        assert_eq!(
            (
                layout.inset_left,
                layout.inset_top,
                layout.inset_right,
                layout.inset_bottom,
            ),
            (16.0, 21.333_334, 26.666_668, 32.0)
        );
        assert!(matches!(
            visual.as_ref(),
            Visual::RichText {
                line_height: 72.0,
                ..
            }
        ));
    }

    #[test]
    fn renders_keynote_theme_gradient_shapes_with_inherited_effects() {
        let document = keynote_document_with_native_slide(200);
        let mut slide = protobuf_message(7, &protobuf_reference(201));
        slide.extend_from_slice(&protobuf_message(42, &protobuf_reference(201)));
        let mut shape = protobuf_message(1, &keynote_drawable(80.0, 90.0, 480.0, 260.0));
        shape.extend_from_slice(&protobuf_message(2, &protobuf_reference(301)));
        let slide = iwa_stream(&[
            iwa_archive(200, 5, &slide, &[201]),
            iwa_archive(201, 3_004, &shape, &[]),
        ]);

        let mut gradient = protobuf_varint(1, 0);
        gradient.extend_from_slice(&protobuf_message(
            2,
            &keynote_gradient_stop_message(0xff33_66ff, 0.0),
        ));
        gradient.extend_from_slice(&protobuf_message(
            2,
            &keynote_gradient_stop_message(0x6633_ffff, 1.0),
        ));
        let fill = protobuf_message(2, &gradient);
        let parent_properties = protobuf_message(1, &fill);
        let parent_style = protobuf_message(11, &parent_properties);

        let mut stroke = protobuf_message(1, &keynote_color_message(0x1122_33ff));
        stroke.extend_from_slice(&protobuf_fixed32(2, 3.0));
        let mut shadow = protobuf_message(1, &keynote_color_message(0x0000_00aa));
        shadow.extend_from_slice(&protobuf_fixed32(2, 315.0));
        shadow.extend_from_slice(&protobuf_fixed32(3, 8.0));
        let mut child_properties = protobuf_message(2, &stroke);
        child_properties.extend_from_slice(&protobuf_fixed32(3, 0.75));
        child_properties.extend_from_slice(&protobuf_message(4, &shadow));
        child_properties.extend_from_slice(&protobuf_message(5, &protobuf_fixed32(1, 0.4)));
        let base = protobuf_message(3, &protobuf_reference(300));
        let mut child_style = protobuf_message(1, &base);
        child_style.extend_from_slice(&protobuf_message(11, &child_properties));
        let stylesheet = iwa_stream(&[
            iwa_archive(300, 3_015, &parent_style, &[]),
            iwa_archive(301, 3_015, &child_style, &[]),
        ]);
        let preview = jpeg(1_024, 768);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-200.iwa", &slide),
            ("Index/ThemeStylesheet.iwa", &stylesheet),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();
        let Visual::Layer {
            opacity,
            visual: layer,
            ..
        } = &document.objects[0].visual
        else {
            panic!("Keynote shape opacity should be preserved as a layer");
        };
        assert_eq!(*opacity, 0.75);
        let Visual::Effect {
            shadow: Some(_),
            visual: effect,
            ..
        } = layer.as_ref()
        else {
            panic!("Keynote theme shadow should be preserved");
        };
        let Visual::AdvancedEffect {
            reflection: Some(reflection),
            visual,
            ..
        } = effect.as_ref()
        else {
            panic!("Keynote theme reflection should be preserved");
        };
        assert_eq!(reflection.start_opacity, 0.4);
        assert_eq!(reflection.end_opacity, 0.4);
        assert!(matches!(
            visual.as_ref(),
            Visual::PaintedShape {
                fill: Paint::LinearGradient { stops, .. },
                stroke: Paint::Solid(0x1122_33ff),
                stroke_width: 4.0,
                ..
            } if stops.len() == 2
        ));
    }

    #[test]
    fn decodes_keynote_advanced_corner_and_focus_gradients() {
        let mut corner = protobuf_varint(1, 1);
        corner.extend_from_slice(&protobuf_varint(4, 1));
        corner.extend_from_slice(&protobuf_message(
            2,
            &keynote_gradient_stop_message(0x0020_60ff, 0.0),
        ));
        corner.extend_from_slice(&protobuf_message(
            2,
            &keynote_gradient_stop_message(0x7030_a0ff, 1.0),
        ));
        corner.extend_from_slice(&protobuf_message(
            5,
            &protobuf_fixed32(2, std::f32::consts::FRAC_PI_4 * 3.0),
        ));
        assert!(matches!(
            super::keynote_gradient(&corner, "Index/DocumentStylesheet.iwa", 960.0, 540.0).unwrap(),
            Paint::RadialGradient {
                x0: 960.0,
                y0: 540.0,
                r1,
                ..
            } if (r1 - 550.7268).abs() < 0.001
        ));

        let mut focus = corner;
        focus.extend_from_slice(&protobuf_varint(4, 0));
        focus.extend_from_slice(&protobuf_message(
            2,
            &keynote_gradient_stop_message(0xffff_ffff, 0.35),
        ));
        assert!(matches!(
            super::keynote_gradient(&focus, "Index/DocumentStylesheet.iwa", 960.0, 540.0).unwrap(),
            Paint::RadialGradient {
                x0: 480.0,
                y0: 270.0,
                ..
            }
        ));
    }

    #[test]
    fn maps_keynote_transformed_radial_gradients_from_their_natural_size() {
        let mut gradient = protobuf_varint(1, 1);
        gradient.extend_from_slice(&protobuf_message(
            2,
            &keynote_gradient_stop_message(0xffff_ffff, 0.0),
        ));
        gradient.extend_from_slice(&protobuf_message(
            2,
            &keynote_gradient_stop_message(0x70ad_47ff, 1.0),
        ));
        let mut transform = protobuf_message(
            1,
            &[protobuf_fixed32(1, 480.0), protobuf_fixed32(2, 270.0)].concat(),
        );
        transform.extend_from_slice(&protobuf_message(
            2,
            &[protobuf_fixed32(1, 900.0), protobuf_fixed32(2, 270.0)].concat(),
        ));
        transform.extend_from_slice(&protobuf_message(
            3,
            &[protobuf_fixed32(1, 960.0), protobuf_fixed32(2, 540.0)].concat(),
        ));
        gradient.extend_from_slice(&protobuf_message(6, &transform));
        assert!(matches!(
            super::keynote_gradient(&gradient, "Index/DocumentStylesheet.iwa", 1280.0, 720.0)
                .unwrap(),
            Paint::RadialGradient {
                x0: 640.0,
                y0: 360.0,
                r1: 560.0,
                ..
            }
        ));
    }

    #[test]
    fn decodes_keynote_tiled_image_fill_properties() {
        let mut fill_size = protobuf_fixed32(1, 8.0);
        fill_size.extend_from_slice(&protobuf_fixed32(2, 8.0));
        let mut image_fill = protobuf_varint(2, 2);
        image_fill.extend_from_slice(&protobuf_message(4, &fill_size));
        image_fill.extend_from_slice(&protobuf_message(6, &protobuf_varint(1, 11)));
        let fill = protobuf_message(3, &image_fill);

        assert_eq!(
            super::keynote_fill_image_properties(&fill, "Index/DocumentStylesheet.iwa").unwrap(),
            Some((11, true, Some((8.0, 8.0))))
        );
    }

    #[test]
    fn ignores_an_empty_keynote_reflection_override() {
        assert_eq!(
            super::keynote_reflection(&[], "Index/ThemeStylesheet.iwa").unwrap(),
            None
        );
    }

    #[test]
    fn ignores_an_empty_keynote_shadow_override() {
        assert_eq!(
            super::keynote_shadow(&[], "Index/ThemeStylesheet.iwa").unwrap(),
            None
        );
    }

    #[test]
    fn reads_keynote_tiff_dimensions_and_list_depth() {
        let mut tiff = b"II*\0\x08\0\0\0\x02\0".to_vec();
        tiff.extend_from_slice(&[0x00, 0x01, 0x04, 0x00, 0x01, 0, 0, 0]);
        tiff.extend_from_slice(&640_u32.to_le_bytes());
        tiff.extend_from_slice(&[0x01, 0x01, 0x03, 0x00, 0x01, 0, 0, 0]);
        tiff.extend_from_slice(&480_u16.to_le_bytes());
        tiff.extend_from_slice(&[0, 0]);
        assert_eq!(super::tiff_dimensions(&tiff), Some((640, 480)));

        let mut payload = protobuf_varint(11, 2);
        payload.extend_from_slice(&protobuf_varint(11, 2));
        payload.extend_from_slice(&protobuf_fixed32(12, 1.5));
        payload.extend_from_slice(&protobuf_fixed32(12, 1.125));
        payload.extend_from_slice(&protobuf_message(14, &protobuf_fixed32(1, 0.6)));
        payload.extend_from_slice(&protobuf_message(21, &keynote_color_message(0x3333_ccff)));
        payload.extend_from_slice(&protobuf_message(16, "•".as_bytes()));
        payload.extend_from_slice(&protobuf_message(16, "◦".as_bytes()));
        let mut archives = vec![super::IwaArchive {
            identifier: 42,
            messages: vec![super::IwaMessage {
                message_type: super::IWORK_LIST_STYLE_TYPE,
                payload,
                data_references: Vec::new(),
            }],
        }];
        assert_eq!(
            super::keynote_list_marker(
                &archives,
                42,
                1,
                "Index/DocumentStylesheet.iwa",
                &mut Vec::new(),
            )
            .unwrap(),
            Some(super::KeynoteListMarker {
                kind: super::KeynoteListMarkerKind::Text("◦".to_owned()),
                level: 1,
                font_family: None,
                scale: 1.0,
                color: Some(0x3333_ccff),
                indent: 36.0,
                text_indent: 1.125,
            })
        );

        let mut parent = protobuf_varint(11, 2);
        parent.extend_from_slice(&protobuf_varint(11, 2));
        parent.extend_from_slice(&protobuf_message(16, "•".as_bytes()));
        parent.extend_from_slice(&protobuf_message(16, "•".as_bytes()));
        let base = protobuf_message(3, &protobuf_reference(41));
        let mut child = protobuf_message(1, &base);
        child.extend_from_slice(&protobuf_message(16, "•".as_bytes()));
        child.extend_from_slice(&protobuf_message(16, "–".as_bytes()));
        child.extend_from_slice(&protobuf_fixed32(12, 1.0));
        child.extend_from_slice(&protobuf_fixed32(12, 0.8));
        let inherited = vec![
            super::IwaArchive {
                identifier: 41,
                messages: vec![super::IwaMessage {
                    message_type: super::IWORK_LIST_STYLE_TYPE,
                    payload: parent,
                    data_references: Vec::new(),
                }],
            },
            super::IwaArchive {
                identifier: 42,
                messages: vec![super::IwaMessage {
                    message_type: super::IWORK_LIST_STYLE_TYPE,
                    payload: child,
                    data_references: Vec::new(),
                }],
            },
        ];
        assert_eq!(
            super::keynote_list_marker(
                &inherited,
                42,
                1,
                "Index/DocumentStylesheet.iwa",
                &mut Vec::new(),
            )
            .unwrap(),
            Some(super::KeynoteListMarker {
                kind: super::KeynoteListMarkerKind::Text("–".to_owned()),
                level: 1,
                font_family: None,
                scale: 1.0,
                color: None,
                indent: 36.0,
                text_indent: 0.8,
            })
        );
        assert_eq!(
            super::keynote_list_marker_text(&super::KeynoteListMarkerKind::Number(0), 2),
            "2."
        );
        assert_eq!(
            super::keynote_list_marker_text(&super::KeynoteListMarkerKind::Number(3), 4),
            "IV."
        );
        assert_eq!(
            super::keynote_list_marker_text(&super::KeynoteListMarkerKind::Number(12), 2),
            "b."
        );
        assert_eq!(
            super::keynote_list_marker_text(&super::KeynoteListMarkerKind::Number(48), 2),
            "②"
        );

        let storage = super::KeynoteTextStorage {
            text: "Move".to_owned(),
            list_styles: vec![super::KeynoteStyleChange {
                character_index: 0,
                identifier: Some(42),
            }],
            ..Default::default()
        };
        let fallback_style = super::KeynoteTextStyle {
            font_family: "SegoeUI".to_owned(),
            font_size: 13.0,
            ..Default::default()
        };
        let runs = super::keynote_text_runs(
            &storage,
            &archives,
            "Index/DocumentStylesheet.iwa",
            &fallback_style,
            Limits::default(),
        )
        .unwrap();
        assert_eq!(runs[0].text, "•");
        assert!((runs[0].font_size - 7.8).abs() < 0.001);
        assert_eq!(runs[0].color, 0x3333_ccff);
        assert_eq!(runs[1].text, "\tMove");
        let pages_runs = super::pages_text_runs_for_range(
            &storage,
            &archives,
            "Index/DocumentStylesheet.iwa",
            0,
            4,
            Limits::default(),
        )
        .unwrap();
        assert_eq!(pages_runs[0].text, "•");
        assert_eq!(pages_runs[0].color, 0x3333_ccff);
        assert_eq!(pages_runs[1].text, "\tMove");
        assert!((pages_runs[0].font_size - pages_runs[1].font_size * 0.6).abs() < 0.001);
        let body_offsets = super::keynote_paragraph_indents(
            &storage,
            &archives,
            "Index/DocumentStylesheet.iwa",
            &fallback_style,
        )
        .unwrap();
        let layouts = super::keynote_text_paragraph_layouts(
            &runs,
            TextAlign::Start,
            15.6,
            false,
            &body_offsets,
        );
        assert!((layouts[0].margin_left - 19.5).abs() < 0.001);
        assert!((layouts[0].first_line_indent + 19.5).abs() < 0.001);
        assert_eq!(layouts[0].default_tab_stop, 36.0);

        let paragraph_properties = protobuf_fixed32(4, 72.0);
        archives.push(super::IwaArchive {
            identifier: 43,
            messages: vec![super::IwaMessage {
                message_type: super::IWORK_PARAGRAPH_STYLE_TYPE,
                payload: protobuf_message(12, &paragraph_properties),
                data_references: Vec::new(),
            }],
        });
        let inherited_storage = super::KeynoteTextStorage {
            text: "One\nTwo".to_owned(),
            paragraph_data: vec![super::KeynoteParagraphDataChange {
                character_index: 0,
                list_level: 1,
            }],
            paragraph_styles: vec![
                super::KeynoteStyleChange {
                    character_index: 0,
                    identifier: Some(43),
                },
                super::KeynoteStyleChange {
                    character_index: 4,
                    identifier: None,
                },
            ],
            list_styles: vec![super::KeynoteStyleChange {
                character_index: 0,
                identifier: Some(42),
            }],
            ..Default::default()
        };
        let inherited_offsets = super::keynote_paragraph_indents(
            &inherited_storage,
            &archives,
            "Index/DocumentStylesheet.iwa",
            &fallback_style,
        )
        .unwrap();
        assert_eq!(
            inherited_offsets
                .iter()
                .map(|layout| layout.margin_left)
                .collect::<Vec<_>>(),
            vec![67.5, 67.5]
        );
        assert!(
            inherited_offsets
                .iter()
                .all(|layout| layout.default_tab_stop == 72.0)
        );
    }

    #[test]
    fn applies_keynote_image_masks_as_crop_and_clip() {
        let document = keynote_document_with_native_slide(200);
        let mut slide = protobuf_message(7, &protobuf_reference(201));
        slide.extend_from_slice(&protobuf_message(42, &protobuf_reference(201)));
        let mut image = protobuf_message(1, &keynote_drawable(50.0, 70.0, 400.0, 200.0));
        image.extend_from_slice(&protobuf_message(5, &protobuf_reference(202)));
        image.extend_from_slice(&protobuf_message(11, &protobuf_reference(500)));
        let scalar = {
            let mut scalar = protobuf_varint(1, 0);
            scalar.extend_from_slice(&protobuf_fixed32(2, 18.0));
            scalar
        };
        let mut mask = protobuf_message(1, &keynote_drawable(150.0, 70.0, 200.0, 200.0));
        mask.extend_from_slice(&protobuf_message(2, &protobuf_message(4, &scalar)));
        let slide = iwa_stream(&[
            iwa_archive(200, 5, &slide, &[201]),
            iwa_archive(201, 3_005, &image, &[500]),
            iwa_archive(202, 3_006, &mask, &[]),
        ]);
        let metadata = iwa_stream(&[iwa_archive(
            2,
            11_006,
            &package_metadata(&[(500, "masked.png")]),
            &[],
        )]);
        let preview = jpeg(1_024, 768);
        let image_bytes = png(400, 200);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-200.iwa", &slide),
            ("Index/Metadata.iwa", &metadata),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
            ("Data/masked.png", &image_bytes),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();
        assert_eq!(document.objects[0].bounds.x, 200.0);
        assert!((document.objects[0].bounds.width - 266.666_7).abs() < 0.001);
        let Visual::Effect {
            clip: Some(Geometry::RoundedRectangle { radius_x, .. }),
            visual,
            ..
        } = &document.objects[0].visual
        else {
            panic!("Keynote image mask should be preserved as a clip");
        };
        assert_eq!(*radius_x, 24.0);
        assert!(matches!(
            visual.as_ref(),
            Visual::Image {
                crop: ImageCrop {
                    left: 0.25,
                    right: 0.25,
                    top: 0.0,
                    bottom: 0.0,
                },
                ..
            }
        ));
    }

    #[test]
    fn expands_keynote_pdf_images_into_native_vector_objects() {
        let document = keynote_document_with_native_slide(200);
        let mut slide = protobuf_message(7, &protobuf_reference(201));
        slide.extend_from_slice(&protobuf_message(42, &protobuf_reference(201)));
        let mut image = protobuf_message(1, &keynote_drawable(64.0, 96.0, 480.0, 320.0));
        image.extend_from_slice(&protobuf_message(11, &protobuf_reference(500)));
        let slide = iwa_stream(&[
            iwa_archive(200, 5, &slide, &[201]),
            iwa_archive(201, super::IWORK_IMAGE_TYPE, &image, &[500]),
        ]);
        let metadata = iwa_stream(&[iwa_archive(
            2,
            super::IWORK_PACKAGE_METADATA_TYPE,
            &package_metadata(&[(500, "slide.pdf")]),
            &[],
        )]);
        let pdf = b"%PDF-1.7
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 120 80] >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /Resources << >> /Contents 4 0 R >> endobj
4 0 obj << /Length 25 >> stream
0 0 1 rg 10 20 40 30 re f
endstream endobj
trailer << /Root 1 0 R >>
%%EOF";
        let preview = jpeg(1_024, 768);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-200.iwa", &slide),
            ("Index/Metadata.iwa", &metadata),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
            ("Data/slide.pdf", pdf),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();
        assert!(!document.objects.is_empty());
        assert!(document.objects.iter().all(|object| {
            matches!(
                object.source.locator,
                SourceLocator::Iwork {
                    kind: "image",
                    ref component,
                } if component.starts_with("embedded-pdf-archive-")
            ) && matches!(object.visual, Visual::Layer { .. })
        }));
        assert!(!document.diagnostics.iter().any(|diagnostic| {
            diagnostic
                .message
                .starts_with("IWORK_KEYNOTE_DRAWABLES_UNSUPPORTED:")
        }));
    }

    #[test]
    fn renders_keynote_cached_chart_grid_as_native_shapes() {
        let document = keynote_document_with_native_slide(200);
        let mut slide = protobuf_message(7, &protobuf_reference(201));
        slide.extend_from_slice(&protobuf_message(42, &protobuf_reference(201)));
        let mut grid = Vec::new();
        for series in ["2024 Study", "2023 Study"] {
            grid.extend_from_slice(&protobuf_message(1, series.as_bytes()));
        }
        for category in ["One", "Two", "Three"] {
            grid.extend_from_slice(&protobuf_message(2, category.as_bytes()));
        }
        for values in [[2.0, 6.0, 4.0], [5.0, 3.0, 7.0]] {
            let mut row = Vec::new();
            for value in values {
                row.extend_from_slice(&protobuf_message(1, &protobuf_fixed64(1, value)));
            }
            grid.extend_from_slice(&protobuf_message(3, &row));
        }
        let mut chart_archive = protobuf_varint(1, 1);
        chart_archive.extend_from_slice(&protobuf_varint(5, 1));
        chart_archive.extend_from_slice(&protobuf_message(7, &grid));
        chart_archive.extend_from_slice(&protobuf_varint(24, 1));
        // Value-label visibility is a series non-style property, not ChartArchive field 24.
        let labels = protobuf_message(10_000, &protobuf_varint(39, 1));
        let mut nonstyles = protobuf_varint(1, 2);
        for (index, identifier) in [(0, 202), (1, 203)] {
            nonstyles.extend(protobuf_message(
                2,
                &[
                    protobuf_varint(1, index),
                    protobuf_message(2, &protobuf_reference(identifier)),
                ]
                .concat(),
            ));
        }
        chart_archive.extend(protobuf_message(19, &nonstyles));
        let mut chart = protobuf_message(1, &keynote_drawable(120.0, 100.0, 700.0, 420.0));
        chart.extend_from_slice(&protobuf_message(10_000, &chart_archive));
        let slide = iwa_stream(&[
            iwa_archive(200, 5, &slide, &[201]),
            iwa_archive(201, 5_021, &chart, &[]),
            iwa_archive(202, super::IWORK_CHART_SERIES_NONSTYLE_TYPE, &labels, &[]),
            iwa_archive(203, super::IWORK_CHART_SERIES_NONSTYLE_TYPE, &labels, &[]),
        ]);
        let preview = jpeg(1_024, 768);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &document),
            ("Index/Slide-200.iwa", &slide),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();
        assert_eq!(document.objects.len(), 27);
        assert_eq!(document.objects[0].kind, ObjectKind::Group);
        assert_eq!(
            document
                .objects
                .iter()
                .filter(|object| object.kind == ObjectKind::Shape)
                .count(),
            26
        );
        assert!(document.objects[1..].iter().all(|object| matches!(
            object.source.locator,
            SourceLocator::Iwork {
                kind: "chart-series",
                ..
            }
        )));
        for text in ["Three", "2024 Study"] {
            assert!(document.objects.iter().any(|object| matches!(
                &object.visual,
                Visual::TextLayout { layout, visual } if !layout.wrap
                    && matches!(visual.as_ref(), Visual::RichText { runs, .. }
                        if runs.iter().any(|run| run.text == text))
            )));
        }
        assert_eq!(
            super::keynote_chart_range(&[vec![0.29, 0.29, 0.24]], false),
            (0.0, 0.3)
        );
        crate::protocol::encode(&document).expect("native Keynote chart should encode");
    }

    #[test]
    fn renders_keynote_one_column_pie_grid_as_multiple_slices() {
        let mut grid = Vec::new();
        for category in ["Approve", "Disapprove", "Neutral"] {
            grid.extend_from_slice(&protobuf_message(1, category.as_bytes()));
        }
        for value in [0.37, 0.59, 0.04] {
            let row = protobuf_message(1, &protobuf_fixed64(1, value));
            grid.extend_from_slice(&protobuf_message(3, &row));
        }
        let mut archive = protobuf_varint(1, 16);
        archive.extend_from_slice(&protobuf_varint(5, 1));
        archive.extend_from_slice(&protobuf_message(7, &grid));
        archive.extend_from_slice(&protobuf_message(10, &protobuf_reference(77)));
        let mut styles = protobuf_varint(1, 3);
        for (index, identifier) in [80, 81, 82].into_iter().enumerate() {
            let mut entry = protobuf_varint(1, index as u64);
            entry.extend_from_slice(&protobuf_message(2, &protobuf_reference(identifier)));
            styles.extend_from_slice(&protobuf_message(2, &entry));
        }
        archive.extend_from_slice(&protobuf_message(18, &styles));
        let mut payload = protobuf_message(1, &keynote_drawable(0.0, 0.0, 300.0, 300.0));
        payload.extend_from_slice(&protobuf_message(10_000, &archive));
        let mut title_properties = protobuf_varint(21, 1);
        title_properties.extend_from_slice(&protobuf_message(23, b"Presidential Approval Rating"));
        let title = protobuf_message(10_000, &title_properties);
        let object_archives = vec![super::IwaArchive {
            identifier: 77,
            messages: vec![super::IwaMessage {
                message_type: super::IWORK_CHART_TITLE_TYPE,
                payload: title,
                data_references: Vec::new(),
            }],
        }];
        let style_archives = [0x6666_66ff, 0xb3b3_b3ff, 0x8080_80ff]
            .into_iter()
            .enumerate()
            .map(|(index, color)| super::IwaArchive {
                identifier: 80 + index as u64,
                messages: vec![super::IwaMessage {
                    message_type: super::IWORK_CHART_SERIES_STYLE_TYPE,
                    payload: protobuf_message(
                        10_000,
                        &protobuf_message(10, &protobuf_message(1, &keynote_color_message(color))),
                    ),
                    data_references: Vec::new(),
                }],
            })
            .collect::<Vec<_>>();
        let chart = super::keynote_chart(
            &payload,
            &object_archives,
            &style_archives,
            "Index/Slide.iwa",
        )
        .unwrap()
        .unwrap();
        let mut objects = Vec::new();
        super::keynote_push_chart(
            chart,
            42,
            0,
            "Index/Slide.iwa",
            Limits::default(),
            &mut objects,
            0.0,
            0.0,
            0,
        )
        .unwrap();

        assert_eq!(objects.len(), 8);
        for (index, color) in [0x6666_66ff, 0xb3b3_b3ff, 0x8080_80ff]
            .into_iter()
            .enumerate()
        {
            assert!(matches!(
                objects[2 + index * 2].visual,
                Visual::PaintedShape {
                    fill: Paint::Solid(actual),
                    ..
                } if actual == color
            ));
        }
        let hidden = protobuf_message(
            10_000,
            &[protobuf_varint(21, 0), protobuf_message(23, b"Chart Title")].concat(),
        );
        let hidden_archives = vec![super::IwaArchive {
            identifier: 78,
            messages: vec![super::IwaMessage {
                message_type: super::IWORK_CHART_TITLE_TYPE,
                payload: hidden,
                data_references: Vec::new(),
            }],
        }];
        assert_eq!(
            super::keynote_chart_title(&hidden_archives, 78, "Index/Slide.iwa").unwrap(),
            None
        );
    }

    #[test]
    fn accepts_a_verified_root_png_preview() {
        let bytes = iwork_package(10_000, None, "preview.png", &png(7, 5));
        let document = detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();

        assert_eq!(
            (document.units[0].width, document.units[0].height),
            (7.0, 5.0)
        );
        assert!(matches!(
            &document.objects[0].visual,
            Visual::Image { media_type, .. } if media_type == "image/png"
        ));
    }

    #[test]
    fn decodes_replaced_keynote_images_by_signature() {
        let jpeg = jpeg(7, 5);
        assert_eq!(
            super::iwork_image_properties("Data/replaced.png", &jpeg, Limits::default()).unwrap(),
            ("image/jpeg", 7.0, 5.0)
        );
    }

    #[test]
    fn keynote_images_never_use_thumbnail_references() {
        let mut image = Vec::new();
        for (field, identifier) in [
            (11, 110),
            (12, 120),
            (13, 130),
            (15, 150),
            (16, 160),
            (17, 170),
        ] {
            image.extend_from_slice(&protobuf_message(field, &protobuf_reference(identifier)));
        }
        assert_eq!(
            super::keynote_image(&image, "Index/Slide.iwa")
                .unwrap()
                .data_identifiers,
            [150, 170, 110, 130]
        );
    }

    #[test]
    fn validates_xml_and_binary_metadata_plists_structurally() {
        for properties in [PROPERTIES, BINARY_PROPERTIES] {
            let bytes = iwork_package_with_properties(properties);
            assert!(
                detect_and_parse(&bytes, Limits::default())
                    .unwrap()
                    .is_some()
            );
        }

        let canonical = b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!DOCTYPE plist PUBLIC \"-//Apple//DTD PLIST 1.0//EN\" \"http://www.apple.com/DTDs/PropertyList-1.0.dtd\">\n<plist version=\"1.0\"><dict/></plist>";
        assert!(
            detect_and_parse(&iwork_package_with_properties(canonical), Limits::default(),)
                .unwrap()
                .is_some()
        );

        for properties in [
            b"bplist00not-a-property-list".as_slice(),
            b"<plistevil></plist>".as_slice(),
            b"<plist><dict><key>missing-value</key></dict></plist>".as_slice(),
            b"<!DOCTYPE plist [<!ENTITY x SYSTEM 'file:///etc/passwd'>]><plist version=\"1.0\"><dict/></plist>".as_slice(),
            cyclic_binary_plist().as_slice(),
        ] {
            let bytes = iwork_package_with_properties(properties);
            let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
            assert_eq!(error.code, DiagnosticCode::FormatInvalid);
            assert_eq!(
                error.location.part.as_deref(),
                Some("Metadata/Properties.plist")
            );
        }
    }

    #[test]
    fn never_treats_arbitrary_media_or_a_filename_claim_as_iwork() {
        let ordinary_zip = stored_zip(&[("Data/first.jpg", &jpeg(4, 3))]);
        assert!(
            detect_and_parse(&ordinary_zip, Limits::default())
                .unwrap()
                .is_none()
        );

        let incomplete = stored_zip(&[
            ("Metadata/Properties.plist", PROPERTIES),
            ("Data/first.jpg", &jpeg(4, 3)),
        ]);
        assert!(
            detect_and_parse(&incomplete, Limits::default())
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn rejects_iwork_without_a_root_preview_instead_of_guessing_an_asset() {
        let iwa = iwa_document(10_000);
        let asset = jpeg(4, 3);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &iwa),
            ("Metadata/Properties.plist", PROPERTIES),
            ("Data/theme.jpg", &asset),
        ]);

        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::UnsupportedFeature);
        assert!(error.message.starts_with("IWORK_PREVIEW_MISSING:"));
    }

    #[test]
    fn rejects_spoofed_preview_bytes_and_excessive_dimensions() {
        let spoofed = iwork_package(10_000, None, "preview.jpg", b"not a jpeg");
        let error = detect_and_parse(&spoofed, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::FormatInvalid);
        assert_eq!(error.phase, Phase::Parse);
        assert_eq!(error.location.part.as_deref(), Some("preview.jpg"));

        let oversized = iwork_package(10_000, None, "preview.png", &png(10_000, 10_000));
        let error = detect_and_parse(&oversized, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::ImageDimensionLimit);
    }

    #[test]
    fn rejects_malformed_or_contradictory_native_iwa_identity() {
        let malformed = stored_zip(&[
            ("Index/Document.iwa", b"not iwa"),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &jpeg(3, 2)),
        ]);
        let error = detect_and_parse(&malformed, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::FormatInvalid);
        assert_eq!(error.location.part.as_deref(), Some("Index/Document.iwa"));

        let contradictory = iwork_package(
            10_000,
            Some("Index/MasterSlide-9.iwa"),
            "preview.jpg",
            &jpeg(3, 2),
        );
        let error = detect_and_parse(&contradictory, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::FormatInvalid);
    }

    #[test]
    fn enforces_the_nested_iwa_expansion_budget_before_decompression() {
        let oversized_declaration = [0, 2, 0, 0, 0x80, 0x08];
        let preview = jpeg(3, 2);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &oversized_declaration),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);
        let limits = Limits {
            max_entry_uncompressed_bytes: 256,
            max_xml_bytes: 256,
            ..Limits::default()
        };

        let error = detect_and_parse(&bytes, limits).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::ZipEntryTooLarge);
        assert_eq!(error.location.part.as_deref(), Some("Index/Document.iwa"));
    }

    #[test]
    fn validates_every_iwa_snappy_chunk() {
        let mut iwa = iwa_document(10_000);
        append_snappy_frame(&mut iwa, &[1, 1, 0]);
        let bytes = iwork_package_with_iwa(&iwa, "preview.jpg", &jpeg(3, 2));

        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::FormatInvalid);
        assert_eq!(error.phase, Phase::Parse);
        assert_eq!(error.location.part.as_deref(), Some("Index/Document.iwa"));
        assert!(error.message.contains("copy offset"));
    }

    #[test]
    fn shares_zip_extraction_and_nested_snappy_expansion_budgets() {
        let mut iwa = iwa_document(10_000);
        append_snappy_frame(&mut iwa, &[120, 0, 0, 254, 1, 0, 218, 1, 0]);
        let bytes = iwork_package_with_iwa(&iwa, "preview.jpg", &jpeg(3, 2));
        let limits = Limits {
            max_entry_uncompressed_bytes: 160,
            max_total_uncompressed_bytes: 160,
            max_xml_bytes: 160,
            ..Limits::default()
        };

        let error = detect_and_parse(&bytes, limits).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::ZipTotalSizeLimit);
        assert_eq!(error.phase, Phase::Parse);
    }

    #[test]
    fn rejects_message_info_field_bombs() {
        let iwa = iwa_document_with_message_info_fields(1_025);
        let bytes = iwork_package_with_iwa(&iwa, "preview.jpg", &jpeg(3, 2));

        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::FormatInvalid);
        assert!(
            error
                .message
                .contains("MessageInfo contains too many fields")
        );
    }

    #[test]
    fn rejects_multiple_root_message_info_entries() {
        let iwa = iwa_document_with_duplicate_message_info();
        let bytes = iwork_package_with_iwa(&iwa, "preview.jpg", &jpeg(3, 2));

        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::FormatInvalid);
        assert!(error.message.contains("exactly one MessageInfo"));
    }

    #[test]
    fn does_not_trust_an_unvalidated_keynote_marker_path() {
        let numbers = iwa_numbers_document();
        let preview = jpeg(3, 2);
        let bytes = stored_zip(&[
            ("Index/Document.iwa", &numbers),
            ("Index/Slide-42.iwa", b"not-iwa"),
            ("Metadata/Properties.plist", PROPERTIES),
            ("preview.jpg", &preview),
        ]);

        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::FormatInvalid);
        assert_eq!(error.location.part.as_deref(), Some("Index/Slide-42.iwa"));
    }

    #[test]
    fn enforces_iwa_cumulative_output_and_operation_budgets() {
        let mut cumulative_iwa = iwa_document(10_000);
        append_snappy_frame(&mut cumulative_iwa, &[120, 0, 0, 254, 1, 0, 218, 1, 0]);
        let bytes = iwork_package_with_iwa(&cumulative_iwa, "preview.jpg", &jpeg(3, 2));
        let first_output = iwa_first_snappy_output(&cumulative_iwa);
        let shared_cap = PROPERTIES.len() + cumulative_iwa.len() + first_output + 119;
        let limits = Limits {
            max_entry_uncompressed_bytes: shared_cap,
            max_total_uncompressed_bytes: shared_cap,
            max_xml_bytes: PROPERTIES.len(),
            ..Limits::default()
        };
        let error = detect_and_parse(&bytes, limits).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::ZipTotalSizeLimit);
        assert_eq!(error.phase, Phase::Parse);

        let mut operation_iwa = iwa_document(10_000);
        let first_chunk_operations = iwa_snappy_operations(&operation_iwa);
        append_snappy_frame(&mut operation_iwa, &[1, 0, 0]);
        let bytes = iwork_package_with_iwa(&operation_iwa, "preview.jpg", &jpeg(3, 2));
        let limits = Limits {
            max_deflate_operations: first_chunk_operations + 3,
            ..Limits::default()
        };
        let error = detect_and_parse(&bytes, limits).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::ZipDeflateInvalid);
        assert_eq!(error.phase, Phase::Parse);
        assert!(error.message.contains("operation budget"));
    }

    #[test]
    fn shares_deflate_and_snappy_operation_budgets() {
        let iwa = iwa_document(10_000);
        let compressed_iwa = raw_deflate_stored(&iwa);
        let (_, deflate_operations) = crate::deflate::decompress_with_limit(
            &compressed_iwa,
            iwa.len(),
            Some(iwa.len()),
            usize::MAX,
        )
        .unwrap();
        let snappy_operations = iwa_snappy_operations(&iwa);
        let preview = jpeg(3, 2);
        let bytes = zip_with_deflated_entry(
            &[
                ("Index/Document.iwa", iwa.as_slice()),
                ("Metadata/Properties.plist", PROPERTIES),
                ("preview.jpg", preview.as_slice()),
            ],
            "Index/Document.iwa",
            &compressed_iwa,
        );

        let limits = Limits {
            max_deflate_operations: deflate_operations + snappy_operations - 1,
            ..Limits::default()
        };
        let error = detect_and_parse(&bytes, limits).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::ZipDeflateInvalid);
        assert_eq!(error.location.part.as_deref(), Some("Index/Document.iwa"));
        assert!(error.message.contains("Snappy operation budget"));

        let limits = Limits {
            max_deflate_operations: deflate_operations + snappy_operations,
            ..Limits::default()
        };
        assert!(detect_and_parse(&bytes, limits).unwrap().is_some());
    }

    #[test]
    fn validates_complete_png_and_jpeg_preview_structures() {
        let mut missing_iend = png(3, 2);
        missing_iend.truncate(missing_iend.len() - 12);
        let bytes = iwork_package(10_000, None, "preview.png", &missing_iend);
        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::FormatInvalid);
        assert_eq!(error.phase, Phase::Parse);

        let mut bad_later_chunk = png(3, 2);
        let idat = bad_later_chunk
            .windows(4)
            .position(|window| window == b"IDAT")
            .unwrap();
        bad_later_chunk[idat + 4] ^= 1;
        let bytes = iwork_package(10_000, None, "preview.png", &bad_later_chunk);
        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::FormatInvalid);
        assert_eq!(error.phase, Phase::Parse);

        let mut missing_eoi = jpeg(3, 2);
        missing_eoi.truncate(missing_eoi.len() - 2);
        let bytes = iwork_package(10_000, None, "preview.jpg", &missing_eoi);
        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::FormatInvalid);
        assert_eq!(error.phase, Phase::Parse);
    }

    #[test]
    fn reports_a_confirmed_unknown_iwork_message_type_as_unsupported() {
        let bytes = iwork_package(999, None, "preview.jpg", &jpeg(3, 2));

        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::UnsupportedFeature);
        assert_eq!(error.phase, Phase::Parse);
        assert!(error.message.starts_with("IWORK_MESSAGE_TYPE_UNSUPPORTED:"));
        assert!(error.message.contains("message type"));
    }

    #[test]
    fn rejects_a_zipped_iwork_directory_package_with_actionable_guidance() {
        let preview = jpeg(3, 2);
        let bytes = stored_zip(&[
            ("Report.pages/Index.zip", b"nested native components"),
            ("Report.pages/Metadata/Properties.plist", PROPERTIES),
            ("Report.pages/preview.jpg", &preview),
        ]);

        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::UnsupportedFeature);
        assert!(
            error
                .message
                .starts_with("IWORK_DIRECTORY_PACKAGE_UNSUPPORTED:")
        );
    }

    #[test]
    fn propagates_the_container_encryption_rejection() {
        let mut bytes = iwork_package(10_000, None, "preview.jpg", &jpeg(3, 2));
        set_first_entry_encrypted(&mut bytes);

        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::ZipEncrypted);
    }

    #[test]
    fn reports_native_iwork_password_markers_as_encrypted() {
        let iwa = iwa_document(10_000);
        let preview = jpeg(3, 2);
        for marker in [".iwpv2", ".iwph"] {
            let bytes = stored_zip(&[
                ("Index/Document.iwa", &iwa),
                ("Metadata/Properties.plist", PROPERTIES),
                ("preview.jpg", &preview),
                (marker, b"password metadata"),
            ]);

            let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
            assert_eq!(error.code, DiagnosticCode::ZipEncrypted);
            assert_eq!(error.phase, Phase::Security);
            assert_eq!(error.location.part.as_deref(), Some(marker));
            assert!(
                error
                    .message
                    .starts_with("IWORK_PASSWORD_PROTECTED_UNSUPPORTED:")
            );
        }
    }

    #[test]
    fn ignores_iwork_password_marker_names_in_non_iwork_zip_files() {
        let bytes = stored_zip(&[
            (".iwpv2", b"unrelated application metadata"),
            ("notes.txt", b"not an iWork package"),
        ]);

        assert!(
            detect_and_parse(&bytes, Limits::default())
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn validates_iwork_identity_before_password_marker_names() {
        let bytes = stored_zip(&[
            ("Index/Document.iwa", b"not an IWA archive"),
            ("Metadata/Properties.plist", PROPERTIES),
            (".iwpv2", b"spoofed password metadata"),
        ]);

        let error = detect_and_parse(&bytes, Limits::default()).unwrap_err();
        assert_ne!(error.code, DiagnosticCode::ZipEncrypted);
        assert_eq!(error.location.part.as_deref(), Some(DOCUMENT_COMPONENT));
    }

    #[test]
    fn top_level_dispatch_reaches_iwork_before_the_strict_native_zip_probe() {
        let preview = jpeg(3, 2);
        let bytes = iwork_package(10_000, None, "preview.jpg", &preview);
        let bytes = with_redundant_local_zip64_on_last_entry(bytes, preview.len() as u64);

        let document = crate::format::detect_and_parse(&bytes, Limits::default())
            .expect("top-level dispatch should accept Apple's bounded ZIP quirk")
            .expect("top-level dispatch should identify iWork");
        assert_eq!(document.format, Some(DocumentFormat::Pages));
    }

    #[test]
    fn varint_reader_rejects_truncation_and_overflow() {
        assert!(read_varint(&[0x80]).is_err());
        assert!(read_varint(&[0xff; 10]).is_err());
        assert_eq!(read_varint(&[0x90, 0x4e]), Ok((10_000, 2)));
    }

    fn iwork_package(
        message_type: u64,
        marker: Option<&str>,
        preview_part: &str,
        preview: &[u8],
    ) -> Vec<u8> {
        let iwa = if message_type == 1 && marker.is_some() {
            iwa_keynote_document()
        } else if message_type == 1 {
            iwa_numbers_document()
        } else {
            iwa_document(message_type)
        };
        let marker_iwa = iwa_document(5);
        let mut entries = vec![
            ("Index/Document.iwa", iwa.as_slice()),
            ("Metadata/Properties.plist", PROPERTIES),
            (preview_part, preview),
        ];
        if let Some(marker) = marker {
            entries.push((marker, marker_iwa.as_slice()));
        }
        stored_zip(&entries)
    }

    fn iwork_package_with_properties(properties: &[u8]) -> Vec<u8> {
        let iwa = iwa_document(10_000);
        let preview = jpeg(3, 2);
        stored_zip(&[
            ("Index/Document.iwa", &iwa),
            ("Metadata/Properties.plist", properties),
            ("preview.jpg", &preview),
        ])
    }

    fn iwork_package_with_iwa(iwa: &[u8], preview_part: &str, preview: &[u8]) -> Vec<u8> {
        stored_zip(&[
            ("Index/Document.iwa", iwa),
            ("Metadata/Properties.plist", PROPERTIES),
            (preview_part, preview),
        ])
    }

    fn iwa_document(message_type: u64) -> Vec<u8> {
        iwa_document_with_payload(message_type, &[0x08, 0x01], 0)
    }

    fn iwa_numbers_document() -> Vec<u8> {
        iwa_numbers_document_with_sheets(1)
    }

    fn iwa_numbers_document_with_sheets(sheet_count: usize) -> Vec<u8> {
        let mut payload = Vec::new();
        for index in 0..sheet_count {
            payload.extend_from_slice(&[0x0a, 0x02, 0x08]);
            payload.push(u8::try_from(index + 2).unwrap());
        }
        payload.extend_from_slice(&[
            0x22, 0x02, 0x08, 0x20, // document stylesheet
            0x2a, 0x02, 0x08, 0x21, // annotation storage
            0x32, 0x02, 0x08, 0x22, // calculation engine
            0x42, 0x02, 0x08, 0x01, // required TSA super archive
        ]);
        iwa_document_with_payload(1, &payload, 0)
    }

    fn native_numbers_package() -> Vec<u8> {
        let mut document = Vec::new();
        document.extend_from_slice(&protobuf_message(1, &protobuf_reference(2)));
        document.extend_from_slice(&protobuf_message(4, &protobuf_reference(32)));
        document.extend_from_slice(&protobuf_message(5, &protobuf_reference(33)));
        document.extend_from_slice(&protobuf_message(6, &protobuf_reference(34)));
        document.extend_from_slice(&protobuf_message(8, &protobuf_reference(1)));
        let mut sheet = protobuf_message(1, b"Sheet1");
        sheet.extend_from_slice(&protobuf_message(2, &protobuf_reference(10)));
        let document_iwa = iwa_stream(&[
            iwa_archive(1, 1, &document, &[]),
            iwa_archive(2, 2, &sheet, &[]),
        ]);

        let table_info = protobuf_message(2, &protobuf_reference(11));
        let mut tile_entry = protobuf_varint(1, 0);
        tile_entry.extend_from_slice(&protobuf_message(2, &protobuf_reference(12)));
        let mut tile_storage = protobuf_message(1, &tile_entry);
        let mut second_tile_entry = protobuf_varint(1, 1);
        second_tile_entry.extend_from_slice(&protobuf_message(2, &protobuf_reference(14)));
        tile_storage.extend_from_slice(&protobuf_message(1, &second_tile_entry));
        let mut store = protobuf_message(3, &tile_storage);
        store.extend_from_slice(&protobuf_message(4, &protobuf_reference(13)));
        store.extend_from_slice(&protobuf_message(5, &protobuf_reference(15)));
        store.extend_from_slice(&protobuf_message(
            1,
            &protobuf_message(2, &protobuf_reference(16)),
        ));
        store.extend_from_slice(&protobuf_message(2, &protobuf_reference(17)));
        let mut table_model = protobuf_message(4, &store);
        table_model.extend_from_slice(&protobuf_varint(6, 257));
        table_model.extend_from_slice(&protobuf_varint(7, 2));
        table_model.extend_from_slice(&protobuf_message(49, &protobuf_reference(18)));
        let mut column_range = protobuf_varint(1, 0);
        column_range.extend_from_slice(&protobuf_varint(2, 1));
        let row_range = protobuf_varint(1, 0);
        let mut tract = protobuf_message(3, &column_range);
        tract.extend_from_slice(&protobuf_message(4, &row_range));
        let mut node = protobuf_varint(1, 67);
        node.extend_from_slice(&protobuf_message(40, &tract));
        let formula = protobuf_message(1, &protobuf_message(1, &node));
        let mut pair = protobuf_varint(1, 1);
        pair.extend_from_slice(&protobuf_message(2, &formula));
        let mut formula_store = protobuf_varint(2, 2);
        formula_store.extend_from_slice(&protobuf_message(3, &pair));
        table_model.extend_from_slice(&protobuf_message(47, &protobuf_message(2, &formula_store)));

        let string_cell = |index: u32| {
            let mut cell = vec![5, 3, 0, 0, 0, 0, 0, 0];
            let styled = index == 0;
            cell.extend_from_slice(&(if styled { 0x68_u32 } else { 8 }).to_le_bytes());
            cell.extend_from_slice(&index.to_le_bytes());
            if styled {
                cell.extend_from_slice(&1_u32.to_le_bytes());
                cell.extend_from_slice(&2_u32.to_le_bytes());
            }
            cell
        };
        let mut tile = protobuf_varint(4, 2);
        tile.extend_from_slice(&protobuf_varint(7, 1));
        for (row, indices) in [(0_u64, [0_u32, 1_u32]), (2, [2_u32, 3_u32])] {
            let mut storage = Vec::new();
            for index in indices {
                if row == 0 && index == 1 {
                    let mut span = vec![5, 1, 0, 0, 0, 0, 0, 0];
                    span.extend_from_slice(&0_u32.to_le_bytes());
                    storage.extend_from_slice(&span);
                } else {
                    storage.extend_from_slice(&string_cell(index));
                }
            }
            let mut row_info = protobuf_varint(1, row);
            row_info.extend_from_slice(&protobuf_varint(2, 2));
            row_info.extend_from_slice(&protobuf_message(6, &storage));
            row_info.extend_from_slice(&protobuf_message(
                7,
                if row == 0 {
                    &[0, 0, 24, 0]
                } else {
                    &[0, 0, 16, 0]
                },
            ));
            tile.extend_from_slice(&protobuf_message(5, &row_info));
        }

        let mut second_tile = protobuf_varint(4, 1);
        second_tile.extend_from_slice(&protobuf_varint(7, 1));
        let mut second_row = protobuf_varint(1, 0);
        second_row.extend_from_slice(&protobuf_varint(2, 1));
        second_row.extend_from_slice(&protobuf_message(6, &string_cell(4)));
        second_row.extend_from_slice(&protobuf_message(7, &[0, 0]));
        second_tile.extend_from_slice(&protobuf_message(5, &second_row));

        let mut strings = protobuf_varint(1, 1);
        for (index, value) in ["Name", "Age", "Alice", "42", "Tail"]
            .into_iter()
            .enumerate()
        {
            let mut entry = protobuf_varint(1, index as u64);
            entry.extend_from_slice(&protobuf_message(3, value.as_bytes()));
            strings.extend_from_slice(&protobuf_message(3, &entry));
        }
        let style_entry = |index, reference| {
            let mut entry = protobuf_varint(1, index);
            entry.extend_from_slice(&protobuf_message(4, &protobuf_reference(reference)));
            protobuf_message(3, &entry)
        };
        let mut styles = protobuf_varint(1, 1);
        styles.extend_from_slice(&style_entry(1, 300));
        styles.extend_from_slice(&style_entry(2, 301));
        let header = |sizes: &[(u64, f32)]| {
            let mut payload = Vec::new();
            for (index, size) in sizes {
                let mut entry = protobuf_varint(1, *index);
                entry.extend_from_slice(&protobuf_fixed32(2, *size));
                payload.extend_from_slice(&protobuf_message(2, &entry));
            }
            payload
        };
        let calculation_engine = iwa_stream(&[
            iwa_archive(10, 6_000, &table_info, &[]),
            iwa_archive(11, 6_001, &table_model, &[18]),
            iwa_archive(12, 6_002, &tile, &[]),
            iwa_archive(13, 6_005, &strings, &[]),
            iwa_archive(14, 6_002, &second_tile, &[]),
            iwa_archive(15, 6_005, &styles, &[300, 301]),
            iwa_archive(16, 6_006, &header(&[(0, 14.1), (2, 30.0)]), &[]),
            iwa_archive(17, 6_006, &header(&[(0, 28.0), (1, 76.0)]), &[]),
            iwa_archive(18, 6_305, &numbers_stroke_sidecar(19), &[19]),
            iwa_archive(19, 6_306, &numbers_stroke_layer(), &[]),
        ]);
        let fill = protobuf_message(1, &keynote_color_message(0x70ad_47ff));
        let mut cell_properties = protobuf_message(1, &fill);
        cell_properties.extend_from_slice(&protobuf_varint(3, 0));
        let cell_style = protobuf_message(11, &cell_properties);
        let text_style = protobuf_message(
            11,
            &keynote_character_properties("Courier", 12.0, 0x0066_ffff, true, true),
        );
        let stylesheet = iwa_stream(&[
            iwa_archive(300, 6_004, &cell_style, &[]),
            iwa_archive(301, 2_021, &text_style, &[]),
        ]);
        stored_zip(&[
            ("Index/Document.iwa", &document_iwa),
            ("Index/CalculationEngine.iwa", &calculation_engine),
            ("Index/DocumentStylesheet.iwa", &stylesheet),
            ("Metadata/Properties.plist", PROPERTIES),
        ])
    }

    fn numbers_stroke_sidecar(layer: u64) -> Vec<u8> {
        let mut sidecar = protobuf_varint(1, 1);
        sidecar.extend_from_slice(&protobuf_varint(2, 2));
        sidecar.extend_from_slice(&protobuf_varint(3, 257));
        sidecar.extend_from_slice(&protobuf_message(6, &protobuf_reference(layer)));
        sidecar
    }

    fn numbers_stroke_layer() -> Vec<u8> {
        let pattern = protobuf_varint(1, 1);
        let mut stroke = protobuf_message(1, &keynote_color_message(0x0000_00ff));
        stroke.extend_from_slice(&protobuf_fixed32(2, 1.0));
        stroke.extend_from_slice(&protobuf_message(6, &pattern));
        let mut run = protobuf_varint(1, 0);
        run.extend_from_slice(&protobuf_varint(2, 1));
        run.extend_from_slice(&protobuf_message(3, &stroke));
        run.extend_from_slice(&protobuf_varint(4, 1));
        let mut layer = protobuf_varint(1, 2);
        layer.extend_from_slice(&protobuf_message(2, &run));
        let mut dash_pattern = protobuf_varint(1, 0);
        dash_pattern.extend_from_slice(&protobuf_varint(3, 2));
        dash_pattern.extend_from_slice(&protobuf_fixed32(4, 2.0));
        dash_pattern.extend_from_slice(&protobuf_fixed32(4, 2.0));
        let mut dash_stroke = protobuf_message(1, &keynote_color_message(0x0000_00ff));
        dash_stroke.extend_from_slice(&protobuf_fixed32(2, 1.0));
        dash_stroke.extend_from_slice(&protobuf_message(6, &dash_pattern));
        let mut dash_run = protobuf_varint(1, 1);
        dash_run.extend_from_slice(&protobuf_varint(2, 1));
        dash_run.extend_from_slice(&protobuf_message(3, &dash_stroke));
        dash_run.extend_from_slice(&protobuf_varint(4, 1));
        layer.extend_from_slice(&protobuf_message(2, &dash_run));
        layer
    }

    fn iwa_keynote_document() -> Vec<u8> {
        iwa_document_with_payload(
            1,
            &[
                0x12, 0x02, 0x08, 0x02, // show reference
                0x1a, 0x02, 0x08, 0x01, // required TSA super archive
            ],
            0,
        )
    }

    fn keynote_document_with_slide_previews() -> Vec<u8> {
        let mut document = Vec::new();
        document.extend_from_slice(&protobuf_message(2, &protobuf_reference(100)));
        document.extend_from_slice(&protobuf_message(3, &[0x08, 0x01]));

        let mut slide_tree = Vec::new();
        for identifier in [101, 102, 103] {
            slide_tree.extend_from_slice(&protobuf_message(2, &protobuf_reference(identifier)));
        }
        let mut size = vec![0x0d];
        size.extend_from_slice(&16.0_f32.to_le_bytes());
        size.push(0x15);
        size.extend_from_slice(&9.0_f32.to_le_bytes());
        let mut show = Vec::new();
        show.extend_from_slice(&protobuf_message(3, &slide_tree));
        show.extend_from_slice(&protobuf_message(4, &size));

        let first_slide = Vec::new();
        let second_slide = Vec::new();
        let third_slide = Vec::new();
        iwa_stream(&[
            iwa_archive(1, 1, &document, &[]),
            iwa_archive(100, 2, &show, &[]),
            iwa_archive(101, 4, &first_slide, &[]),
            iwa_archive(102, 4, &second_slide, &[201]),
            iwa_archive(103, 4, &third_slide, &[202]),
        ])
    }

    fn keynote_document_with_native_slide(slide_identifier: u64) -> Vec<u8> {
        let mut document = Vec::new();
        document.extend_from_slice(&protobuf_message(2, &protobuf_reference(100)));
        document.extend_from_slice(&protobuf_message(3, &[0x08, 0x01]));

        let slide_tree = protobuf_message(2, &protobuf_reference(101));
        let mut size = vec![0x0d];
        size.extend_from_slice(&1_024.0_f32.to_le_bytes());
        size.push(0x15);
        size.extend_from_slice(&768.0_f32.to_le_bytes());
        let mut show = Vec::new();
        show.extend_from_slice(&protobuf_message(3, &slide_tree));
        show.extend_from_slice(&protobuf_message(4, &size));
        let slide_node = protobuf_message(2, &protobuf_reference(slide_identifier));

        iwa_stream(&[
            iwa_archive(1, 1, &document, &[]),
            iwa_archive(100, 2, &show, &[]),
            iwa_archive(101, 4, &slide_node, &[]),
        ])
    }

    fn keynote_document_with_legacy_slide() -> Vec<u8> {
        let mut document = protobuf_message(2, &protobuf_reference(100));
        document.extend_from_slice(&protobuf_message(3, &[0x08, 0x01]));
        let mut size = vec![0x0d];
        size.extend_from_slice(&1_024.0_f32.to_le_bytes());
        size.push(0x15);
        size.extend_from_slice(&768.0_f32.to_le_bytes());
        let show = protobuf_message(4, &size);
        iwa_stream(&[
            iwa_archive(1, 1, &document, &[]),
            iwa_archive(100, 2, &show, &[]),
        ])
    }

    fn keynote_slide_with_text_placeholder(text: &str) -> Vec<u8> {
        let mut slide = Vec::new();
        slide.extend_from_slice(&protobuf_message(7, &protobuf_reference(201)));
        slide.extend_from_slice(&protobuf_message(42, &protobuf_reference(201)));

        let mut position = vec![0x0d];
        position.extend_from_slice(&100.0_f32.to_le_bytes());
        position.push(0x15);
        position.extend_from_slice(&120.0_f32.to_le_bytes());
        let mut size = vec![0x0d];
        size.extend_from_slice(&824.0_f32.to_le_bytes());
        size.push(0x15);
        size.extend_from_slice(&260.0_f32.to_le_bytes());
        let mut geometry = Vec::new();
        geometry.extend_from_slice(&protobuf_message(1, &position));
        geometry.extend_from_slice(&protobuf_message(2, &size));
        let drawable = protobuf_message(1, &geometry);
        let shape = protobuf_message(1, &drawable);
        let mut shape_info = protobuf_message(1, &shape);
        shape_info.extend_from_slice(&protobuf_message(4, &protobuf_reference(202)));
        let placeholder = protobuf_message(1, &shape_info);
        let storage = protobuf_message(3, text.as_bytes());

        iwa_stream(&[
            iwa_archive(200, 5, &slide, &[201]),
            iwa_archive(201, 7, &placeholder, &[202]),
            iwa_archive(202, 2_001, &storage, &[]),
        ])
    }

    fn keynote_slide_with_rich_text_placeholder() -> Vec<u8> {
        let mut slide = Vec::new();
        slide.extend_from_slice(&protobuf_message(7, &protobuf_reference(201)));
        slide.extend_from_slice(&protobuf_message(42, &protobuf_reference(201)));

        let mut position = vec![0x0d];
        position.extend_from_slice(&100.0_f32.to_le_bytes());
        position.push(0x15);
        position.extend_from_slice(&120.0_f32.to_le_bytes());
        let mut size = vec![0x0d];
        size.extend_from_slice(&824.0_f32.to_le_bytes());
        size.push(0x15);
        size.extend_from_slice(&260.0_f32.to_le_bytes());
        let mut geometry = Vec::new();
        geometry.extend_from_slice(&protobuf_message(1, &position));
        geometry.extend_from_slice(&protobuf_message(2, &size));
        let drawable = protobuf_message(1, &geometry);
        let shape = protobuf_message(1, &drawable);
        let mut shape_info = protobuf_message(1, &shape);
        shape_info.extend_from_slice(&protobuf_message(4, &protobuf_reference(202)));
        let placeholder = protobuf_message(1, &shape_info);

        let mut storage = protobuf_message(3, b"Base Accent");
        storage.extend_from_slice(&protobuf_message(5, &keynote_style_table(&[(0, 300)])));
        storage.extend_from_slice(&protobuf_message(8, &keynote_style_table(&[(5, 301)])));

        iwa_stream(&[
            iwa_archive(200, 5, &slide, &[201]),
            iwa_archive(201, 7, &placeholder, &[202]),
            iwa_archive(202, 2_001, &storage, &[]),
        ])
    }

    fn keynote_rich_text_stylesheet() -> Vec<u8> {
        let mut paragraph = Vec::new();
        paragraph.extend_from_slice(&protobuf_message(
            11,
            &keynote_character_properties("Helvetica", 24.0, 0xff00_00ff, false, false),
        ));
        paragraph.extend_from_slice(&protobuf_message(12, &[0x08, 0x02]));

        let mut accent = keynote_character_properties("Courier", 32.0, 0x0066_ffff, true, true);
        accent.extend_from_slice(&[0x58, 0x01, 0x60, 0x01, 0x75]);
        accent.extend_from_slice(&3.0_f32.to_le_bytes());
        accent.extend_from_slice(&protobuf_fixed32(27, 1.5 / 32.0));
        accent.extend_from_slice(&protobuf_message(26, &keynote_color_message(0xffff_00ff)));
        let character = protobuf_message(11, &accent);
        iwa_stream(&[
            iwa_archive(300, 2_022, &paragraph, &[]),
            iwa_archive(301, 2_021, &character, &[]),
        ])
    }

    fn keynote_style_table(styles: &[(u64, u64)]) -> Vec<u8> {
        let mut table = Vec::new();
        for (character_index, style_identifier) in styles {
            let mut entry = Vec::new();
            push_varint(&mut entry, 1 << 3);
            push_varint(&mut entry, *character_index);
            entry.extend_from_slice(&protobuf_message(2, &protobuf_reference(*style_identifier)));
            table.extend_from_slice(&protobuf_message(1, &entry));
        }
        table
    }

    fn keynote_character_properties(
        font_family: &str,
        font_size: f32,
        color: u32,
        bold: bool,
        italic: bool,
    ) -> Vec<u8> {
        let mut properties = vec![0x08, u8::from(bold), 0x10, u8::from(italic), 0x1d];
        properties.extend_from_slice(&font_size.to_le_bytes());
        properties.extend_from_slice(&protobuf_message(5, font_family.as_bytes()));

        properties.extend_from_slice(&protobuf_message(7, &keynote_color_message(color)));
        properties
    }

    fn keynote_color_message(color: u32) -> Vec<u8> {
        let mut encoded_color = vec![0x1d];
        encoded_color.extend_from_slice(&((color >> 24) as f32 / 255.0).to_le_bytes());
        encoded_color.push(0x25);
        encoded_color.extend_from_slice(&(((color >> 16) & 0xff) as f32 / 255.0).to_le_bytes());
        encoded_color.push(0x2d);
        encoded_color.extend_from_slice(&(((color >> 8) & 0xff) as f32 / 255.0).to_le_bytes());
        encoded_color.push(0x35);
        encoded_color.extend_from_slice(&((color & 0xff) as f32 / 255.0).to_le_bytes());
        encoded_color
    }

    fn keynote_drawable(x: f32, y: f32, width: f32, height: f32) -> Vec<u8> {
        let mut position = protobuf_fixed32(1, x);
        position.extend_from_slice(&protobuf_fixed32(2, y));
        let mut size = protobuf_fixed32(1, width);
        size.extend_from_slice(&protobuf_fixed32(2, height));
        let mut geometry = protobuf_message(1, &position);
        geometry.extend_from_slice(&protobuf_message(2, &size));
        protobuf_message(1, &geometry)
    }

    fn keynote_gradient_stop_message(color: u32, offset: f32) -> Vec<u8> {
        let mut stop = protobuf_message(1, &keynote_color_message(color));
        stop.extend_from_slice(&protobuf_fixed32(2, offset));
        stop
    }

    fn package_metadata(data: &[(u64, &str)]) -> Vec<u8> {
        let mut metadata = Vec::new();
        for (identifier, file_name) in data {
            let mut info = vec![0x08];
            push_varint(&mut info, *identifier);
            info.extend_from_slice(&protobuf_message(4, file_name.as_bytes()));
            metadata.extend_from_slice(&protobuf_message(4, &info));
        }
        metadata
    }

    fn protobuf_reference(identifier: u64) -> Vec<u8> {
        let mut reference = vec![0x08];
        push_varint(&mut reference, identifier);
        reference
    }

    fn protobuf_message(field: u64, payload: &[u8]) -> Vec<u8> {
        let mut message = Vec::new();
        push_varint(&mut message, field << 3 | 2);
        push_varint(&mut message, payload.len() as u64);
        message.extend_from_slice(payload);
        message
    }

    fn protobuf_varint(field: u64, value: u64) -> Vec<u8> {
        let mut message = Vec::new();
        push_varint(&mut message, field << 3);
        push_varint(&mut message, value);
        message
    }

    fn protobuf_fixed32(field: u64, value: f32) -> Vec<u8> {
        let mut message = Vec::new();
        push_varint(&mut message, field << 3 | 5);
        message.extend_from_slice(&value.to_le_bytes());
        message
    }

    fn protobuf_fixed64(field: u64, value: f64) -> Vec<u8> {
        let mut message = Vec::new();
        push_varint(&mut message, field << 3 | 1);
        message.extend_from_slice(&value.to_le_bytes());
        message
    }

    fn iwa_archive(
        identifier: u64,
        message_type: u64,
        payload: &[u8],
        data_references: &[u64],
    ) -> Vec<u8> {
        let mut message_info = vec![0x08];
        push_varint(&mut message_info, message_type);
        message_info.push(0x18);
        push_varint(&mut message_info, payload.len() as u64);
        if !data_references.is_empty() {
            let mut packed = Vec::new();
            for identifier in data_references {
                push_varint(&mut packed, *identifier);
            }
            message_info.extend_from_slice(&protobuf_message(6, &packed));
        }

        let mut archive_info = vec![0x08];
        push_varint(&mut archive_info, identifier);
        archive_info.extend_from_slice(&protobuf_message(2, &message_info));

        let mut archive = Vec::new();
        push_varint(&mut archive, archive_info.len() as u64);
        archive.extend_from_slice(&archive_info);
        archive.extend_from_slice(payload);
        archive
    }

    fn iwa_stream(archives: &[Vec<u8>]) -> Vec<u8> {
        let mut uncompressed = Vec::new();
        for archive in archives {
            uncompressed.extend_from_slice(archive);
        }
        snappy_literal_frame(&uncompressed)
    }

    fn iwa_document_with_message_info_fields(extra_fields: usize) -> Vec<u8> {
        iwa_document_with_payload(10_000, &[0x08, 0x01], extra_fields)
    }

    fn iwa_document_with_duplicate_message_info() -> Vec<u8> {
        let payload = [0x08, 0x01];
        let mut message_info = vec![0x08];
        push_varint(&mut message_info, 10_000);
        message_info.push(0x18);
        push_varint(&mut message_info, payload.len() as u64);

        let mut archive_info = vec![0x08, 0x01];
        for _ in 0..2 {
            archive_info.push(0x12);
            push_varint(&mut archive_info, message_info.len() as u64);
            archive_info.extend_from_slice(&message_info);
        }
        iwa_frame(&archive_info, &payload)
    }

    fn iwa_document_with_payload(
        message_type: u64,
        payload: &[u8],
        extra_message_info_fields: usize,
    ) -> Vec<u8> {
        let mut message_info = vec![0x08];
        push_varint(&mut message_info, message_type);
        message_info.push(0x18);
        push_varint(&mut message_info, payload.len() as u64);
        for _ in 0..extra_message_info_fields {
            message_info.extend_from_slice(&[0x20, 0x00]);
        }

        let mut archive_info = vec![0x08, 0x01, 0x12];
        push_varint(&mut archive_info, message_info.len() as u64);
        archive_info.extend_from_slice(&message_info);

        iwa_frame(&archive_info, payload)
    }

    fn iwa_frame(archive_info: &[u8], payload: &[u8]) -> Vec<u8> {
        let mut uncompressed = Vec::new();
        push_varint(&mut uncompressed, archive_info.len() as u64);
        uncompressed.extend_from_slice(archive_info);
        uncompressed.extend_from_slice(payload);

        snappy_literal_frame(&uncompressed)
    }

    fn snappy_literal_frame(uncompressed: &[u8]) -> Vec<u8> {
        let mut snappy = Vec::new();
        push_varint(&mut snappy, uncompressed.len() as u64);
        if uncompressed.len() <= 60 {
            snappy.push(((uncompressed.len() - 1) << 2) as u8);
        } else {
            let encoded_length = uncompressed.len() - 1;
            let length_bytes = if encoded_length <= u8::MAX as usize {
                1
            } else if encoded_length <= u16::MAX as usize {
                2
            } else {
                4
            };
            snappy.push(((59 + length_bytes) << 2) as u8);
            snappy.extend_from_slice(&encoded_length.to_le_bytes()[..length_bytes]);
        }
        snappy.extend_from_slice(uncompressed);

        let mut iwa = vec![0];
        let length = snappy.len() as u32;
        iwa.extend_from_slice(&length.to_le_bytes()[..3]);
        iwa.extend_from_slice(&snappy);
        iwa
    }

    fn cyclic_binary_plist() -> Vec<u8> {
        let mut bytes = b"bplist00".to_vec();
        bytes.extend_from_slice(&[0xd1, 0x01, 0x02]);
        bytes.extend_from_slice(&[0x51, b'k']);
        bytes.extend_from_slice(&[0xa1, 0x02]);
        bytes.extend_from_slice(&[8, 11, 13]);
        bytes.extend_from_slice(&[0; 6]);
        bytes.extend_from_slice(&[1, 1]);
        bytes.extend_from_slice(&3_u64.to_be_bytes());
        bytes.extend_from_slice(&0_u64.to_be_bytes());
        bytes.extend_from_slice(&15_u64.to_be_bytes());
        bytes
    }

    fn append_snappy_frame(iwa: &mut Vec<u8>, snappy: &[u8]) {
        assert!(snappy.len() <= 0x00ff_ffff);
        iwa.push(0);
        iwa.extend_from_slice(&(snappy.len() as u32).to_le_bytes()[..3]);
        iwa.extend_from_slice(snappy);
    }

    fn iwa_snappy_operations(iwa: &[u8]) -> usize {
        let compressed_length =
            usize::from(iwa[1]) | (usize::from(iwa[2]) << 8) | (usize::from(iwa[3]) << 16);
        let chunk = &iwa[4..4 + compressed_length];
        let (expanded_length, _) = read_varint(chunk).unwrap();
        compressed_length + usize::try_from(expanded_length).unwrap()
    }

    fn iwa_first_snappy_output(iwa: &[u8]) -> usize {
        let compressed_length =
            usize::from(iwa[1]) | (usize::from(iwa[2]) << 8) | (usize::from(iwa[3]) << 16);
        let chunk = &iwa[4..4 + compressed_length];
        usize::try_from(read_varint(chunk).unwrap().0).unwrap()
    }

    fn push_varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let byte = (value & 0x7f) as u8;
            value >>= 7;
            if value == 0 {
                output.push(byte);
                return;
            }
            output.push(byte | 0x80);
        }
    }

    fn stored_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
        zip_with_deflated_entry(entries, "", &[])
    }

    fn zip_with_deflated_entry(
        entries: &[(&str, &[u8])],
        deflated_name: &str,
        deflated: &[u8],
    ) -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut central = Vec::new();
        for (name, content) in entries {
            let compressed = if *name == deflated_name {
                deflated
            } else {
                content
            };
            let method = if *name == deflated_name { 8 } else { 0 };
            let name = name.as_bytes();
            let offset = u32::try_from(bytes.len()).unwrap();
            let crc = crate::zip::crc32(content);

            push_u32(&mut bytes, 0x0403_4b50);
            push_u16(&mut bytes, 20);
            push_u16(&mut bytes, 1 << 11);
            push_u16(&mut bytes, method);
            push_u16(&mut bytes, 0);
            push_u16(&mut bytes, 0);
            push_u32(&mut bytes, crc);
            push_u32(&mut bytes, compressed.len() as u32);
            push_u32(&mut bytes, content.len() as u32);
            push_u16(&mut bytes, name.len() as u16);
            push_u16(&mut bytes, 0);
            bytes.extend_from_slice(name);
            bytes.extend_from_slice(compressed);

            push_u32(&mut central, 0x0201_4b50);
            push_u16(&mut central, 20);
            push_u16(&mut central, 20);
            push_u16(&mut central, 1 << 11);
            push_u16(&mut central, method);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u32(&mut central, crc);
            push_u32(&mut central, compressed.len() as u32);
            push_u32(&mut central, content.len() as u32);
            push_u16(&mut central, name.len() as u16);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u32(&mut central, 0);
            push_u32(&mut central, offset);
            central.extend_from_slice(name);
        }

        let central_offset = bytes.len() as u32;
        let central_size = central.len() as u32;
        bytes.extend_from_slice(&central);
        push_u32(&mut bytes, 0x0605_4b50);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, entries.len() as u16);
        push_u16(&mut bytes, entries.len() as u16);
        push_u32(&mut bytes, central_size);
        push_u32(&mut bytes, central_offset);
        push_u16(&mut bytes, 0);
        bytes
    }

    fn raw_deflate_stored(data: &[u8]) -> Vec<u8> {
        let length = u16::try_from(data.len()).unwrap();
        let mut output = vec![1];
        output.extend_from_slice(&length.to_le_bytes());
        output.extend_from_slice(&(!length).to_le_bytes());
        output.extend_from_slice(data);
        output
    }

    fn push_u16(output: &mut Vec<u8>, value: u16) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn push_u32(output: &mut Vec<u8>, value: u32) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn jpeg(width: u16, height: u16) -> Vec<u8> {
        let mut bytes = vec![0xff, 0xd8, 0xff, 0xc0, 0x00, 0x0b, 8];
        bytes.extend_from_slice(&height.to_be_bytes());
        bytes.extend_from_slice(&width.to_be_bytes());
        bytes.extend_from_slice(&[
            1, 1, 0x11, 0, 0xff, 0xda, 0, 8, 1, 1, 0, 0, 63, 0, 0, 0xff, 0xd9,
        ]);
        bytes
    }

    fn png(width: u32, height: u32) -> Vec<u8> {
        let mut bytes = b"\x89PNG\r\n\x1a\n".to_vec();
        let mut ihdr = Vec::with_capacity(13);
        ihdr.extend_from_slice(&width.to_be_bytes());
        ihdr.extend_from_slice(&height.to_be_bytes());
        ihdr.extend_from_slice(&[8, 6, 0, 0, 0]);
        push_png_chunk(&mut bytes, b"IHDR", &ihdr);
        push_png_chunk(
            &mut bytes,
            b"IDAT",
            &[0x78, 0x01, 0x01, 0, 0, 0xff, 0xff, 0, 0, 0, 1],
        );
        push_png_chunk(&mut bytes, b"IEND", &[]);
        bytes
    }

    fn push_png_chunk(output: &mut Vec<u8>, kind: &[u8; 4], data: &[u8]) {
        output.extend_from_slice(&(data.len() as u32).to_be_bytes());
        let crc_start = output.len();
        output.extend_from_slice(kind);
        output.extend_from_slice(data);
        output.extend_from_slice(&crate::zip::crc32(&output[crc_start..]).to_be_bytes());
    }

    fn set_first_entry_encrypted(bytes: &mut [u8]) {
        let local = bytes
            .windows(4)
            .position(|window| window == 0x0403_4b50_u32.to_le_bytes())
            .unwrap();
        let flags = u16::from_le_bytes([bytes[local + 6], bytes[local + 7]]) | 1;
        bytes[local + 6..local + 8].copy_from_slice(&flags.to_le_bytes());

        let central = bytes
            .windows(4)
            .position(|window| window == 0x0201_4b50_u32.to_le_bytes())
            .unwrap();
        let flags = u16::from_le_bytes([bytes[central + 8], bytes[central + 9]]) | 1;
        bytes[central + 8..central + 10].copy_from_slice(&flags.to_le_bytes());
    }

    fn with_redundant_local_zip64_on_last_entry(mut bytes: Vec<u8>, size: u64) -> Vec<u8> {
        let old_eocd = bytes.len() - 22;
        let central_offset =
            u32::from_le_bytes(bytes[old_eocd + 16..old_eocd + 20].try_into().unwrap()) as usize;
        let mut cursor = central_offset;
        let mut last_local_offset = None;
        while cursor < old_eocd {
            assert_eq!(&bytes[cursor..cursor + 4], &0x0201_4b50_u32.to_le_bytes());
            last_local_offset = Some(u32::from_le_bytes(
                bytes[cursor + 42..cursor + 46].try_into().unwrap(),
            ) as usize);
            let name = usize::from(u16::from_le_bytes(
                bytes[cursor + 28..cursor + 30].try_into().unwrap(),
            ));
            let extra = usize::from(u16::from_le_bytes(
                bytes[cursor + 30..cursor + 32].try_into().unwrap(),
            ));
            let comment = usize::from(u16::from_le_bytes(
                bytes[cursor + 32..cursor + 34].try_into().unwrap(),
            ));
            cursor += 46 + name + extra + comment;
        }
        let local = last_local_offset.unwrap();
        let name_length = usize::from(u16::from_le_bytes(
            bytes[local + 26..local + 28].try_into().unwrap(),
        ));
        bytes[local + 28..local + 30].copy_from_slice(&20_u16.to_le_bytes());
        let mut extra = vec![0x01, 0x00, 0x10, 0x00];
        extra.extend_from_slice(&size.to_le_bytes());
        extra.extend_from_slice(&size.to_le_bytes());
        bytes.splice(local + 30 + name_length..local + 30 + name_length, extra);

        let new_eocd = old_eocd + 20;
        bytes[new_eocd + 16..new_eocd + 20]
            .copy_from_slice(&((central_offset + 20) as u32).to_le_bytes());
        bytes
    }

    #[test]
    fn opens_pages_with_an_infinite_saved_origin() {
        let document = detect_and_parse(
            include_bytes!("../../tests/fixtures/newletter.pages"),
            Limits::default(),
        )
        .expect("Pages infinite origin sentinel should not be fatal")
        .expect("Pages package should be detected");

        assert_eq!(document.format, Some(crate::model::DocumentFormat::Pages));
        assert_eq!(document.units.len(), 2);
        assert!(document.objects.len() >= 100);
        assert!(document.objects.iter().any(|object| {
            object
                .text
                .as_deref()
                .is_some_and(|text| text.contains("TODAY’S NEWS"))
        }));
    }

    #[test]
    fn preserves_report_drop_cap() {
        let document = detect_and_parse(
            include_bytes!("../../tests/fixtures/report.pages"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let body = document
            .objects
            .iter()
            .find(|o| {
                o.text
                    .as_deref()
                    .is_some_and(|t| t.contains("To get started"))
            })
            .unwrap();
        let Visual::TextLayout { layout, visual } = &body.visual else {
            panic!("missing layout")
        };
        let cap = layout.paragraphs[5]
            .drop_cap
            .expect("native first body paragraph has a drop cap");
        assert_eq!((cap.characters, cap.lines, cap.raised_lines), (1, 3, 0));
        assert_eq!((cap.padding, cap.outdent), (0.0, 0.0));
        assert!(
            layout
                .paragraphs
                .iter()
                .enumerate()
                .all(|(i, p)| i == 5 || p.drop_cap.is_none()),
            "the anchored drop-cap entry must not leak to subsequent paragraphs"
        );
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("missing text")
        };
        let initial = runs
            .iter()
            .find(|r| r.text == "T")
            .expect("drop-cap style isolates T without losing it");
        assert!(initial.bold);
        assert_eq!(initial.font_family, "HelveticaNeue-Bold");
        assert_eq!(
            runs.iter().map(|r| r.text.as_str()).collect::<String>(),
            body.text.as_deref().unwrap()
        );
    }

    #[test]
    fn preserves_report_header_on_every_saved_page() {
        let bytes = include_bytes!("../../tests/fixtures/report.pages");
        let document = detect_and_parse(bytes, Limits::default())
            .expect("Report.pages should parse")
            .expect("Pages package should be detected");

        for unit_index in 0..2 {
            let page_number = (unit_index + 1).to_string();
            let texts = document
                .objects
                .iter()
                .filter(|object| object.unit_index == unit_index)
                .filter_map(|object| object.text.as_deref())
                .collect::<Vec<_>>();
            assert!(
                document.objects.iter().any(|object| {
                    object.unit_index == unit_index
                        && object.text.as_deref() == Some("Simple Home Styling")
                }),
                "page {unit_index}: {texts:?}"
            );
            assert!(texts.contains(&page_number.as_str()));
        }
        let body = document
            .objects
            .iter()
            .find(|object| {
                object.unit_index == 1
                    && object
                        .text
                        .as_deref()
                        .is_some_and(|text| text.contains("You can use Pages for both"))
            })
            .unwrap();
        let Visual::TextLayout { layout, .. } = &body.visual else {
            panic!("missing body layout")
        };
        // Creator Studio: Heading uses 0.9 Lines; Body uses 1.1 Lines.
        assert!((layout.paragraphs[1].line_height - 34.56).abs() < 0.01);
        assert!((layout.paragraphs[2].line_height - 21.12).abs() < 0.01);
        let first_page = document
            .objects
            .iter()
            .find(|object| {
                object.unit_index == 0
                    && object
                        .text
                        .as_deref()
                        .is_some_and(|text| text.contains("It’s easy to edit text"))
            })
            .unwrap();
        let Visual::TextLayout { layout, .. } = &first_page.visual else {
            panic!("missing first page layout")
        };
        assert!(
            (layout.paragraphs[6].line_height - 21.12).abs() < 0.01,
            "the paragraph-anchored image must not inflate every body line: {}",
            layout.paragraphs[6].line_height,
        );
        assert!(
            layout.paragraphs[6].space_after > 400.0,
            "the following paragraph must still leave room for the anchored image"
        );
    }

    #[test]
    fn preserves_editable_bezier_handles_and_subpath_closure() {
        let point = |x, y| [protobuf_fixed32(1, x), protobuf_fixed32(2, y)].concat();
        let node = |incoming: (f32, f32), anchor: (f32, f32), outgoing: (f32, f32)| {
            [
                protobuf_message(1, &point(incoming.0, incoming.1)),
                protobuf_message(2, &point(anchor.0, anchor.1)),
                protobuf_message(3, &point(outgoing.0, outgoing.1)),
                protobuf_varint(4, 2),
            ]
            .concat()
        };
        let nodes = [
            protobuf_message(1, &node((0.0, 2.0), (0.0, 0.0), (2.0, 0.0))),
            protobuf_message(1, &node((8.0, 0.0), (10.0, 10.0), (10.0, 8.0))),
        ]
        .concat();
        for closed in [false, true] {
            let path = [
                protobuf_message(
                    1,
                    &[nodes.clone(), protobuf_varint(2, u64::from(closed))].concat(),
                ),
                protobuf_message(2, &point(10.0, 10.0)),
            ]
            .concat();
            let Geometry::Path { commands, .. } = super::keynote_path_source(
                &protobuf_message(8, &path),
                20.0,
                30.0,
                DOCUMENT_COMPONENT,
            )
            .unwrap()
            .unwrap() else {
                panic!("missing editable path")
            };
            assert_eq!(
                commands[1],
                crate::model::PathCommand::BezierCurveTo {
                    cp1x: 4.0,
                    cp1y: 0.0,
                    cp2x: 16.0,
                    cp2y: 0.0,
                    x: 20.0,
                    y: 30.0,
                }
            );
            assert_eq!(commands.len(), if closed { 4 } else { 2 });
            if closed {
                assert_eq!(
                    commands[2],
                    crate::model::PathCommand::BezierCurveTo {
                        cp1x: 20.0,
                        cp1y: 24.0,
                        cp2x: 0.0,
                        cp2y: 6.0,
                        x: 0.0,
                        y: 0.0,
                    }
                );
                assert_eq!(commands[3], crate::model::PathCommand::ClosePath);
            }
        }
        let invalid = protobuf_message(1, &protobuf_message(1, &[]));
        assert!(
            super::keynote_editable_bezier_path(&invalid, 10.0, 10.0, DOCUMENT_COMPONENT).is_err()
        );
    }

    #[test]
    fn preserves_untitled_brochure_map_outline() {
        let bytes = include_bytes!("../../tests/fixtures/untitled.pages");
        let document = detect_and_parse(bytes, Limits::default()).unwrap().unwrap();
        let map = document
            .objects
            .iter()
            .find(|o| o.unit_index == 0 && o.stable_id.ends_with("-58312"))
            .unwrap();
        let Visual::PaintedShape {
            geometry: Geometry::Path { commands, .. },
            ..
        } = &map.visual
        else {
            panic!("the Hawaii map must retain its island outlines, not a rectangle");
        };
        assert!(commands.len() > 30);
    }

    #[test]
    fn preserves_newletter2_tabs_and_allocates_same_page_flow() {
        let bytes = include_bytes!("../../tests/fixtures/newletter2.pages");
        let document = detect_and_parse(bytes, Limits::default()).unwrap().unwrap();
        let header = document
            .objects
            .iter()
            .find(|o| {
                o.unit_index == 1
                    && o.text
                        .as_deref()
                        .is_some_and(|t| t.starts_with("ISSUE 1\t"))
            })
            .unwrap();
        let Visual::TextLayout { layout, .. } = &header.visual else {
            panic!("missing header layout")
        };
        assert!(
            layout
                .tab_stops
                .iter()
                .any(|tab| tab.align == crate::model::TextAlign::End && tab.position > 600.0)
        );
        assert_eq!(
            document
                .objects
                .iter()
                .filter(|o| o.unit_index == 1
                    && o.text.as_deref().is_some_and(
                        |t| t.contains("This newsletter template uses linked text boxes")
                    ))
                .count(),
            1,
            "an unavailable per-frame range must not duplicate the complete story in the next frame"
        );
        assert!(
            document
                .diagnostics
                .iter()
                .any(|d| d.message.contains("PAGES_TEXT_FLOW_APPROXIMATE"))
        );
        let body = document
            .objects
            .iter()
            .find(|o| {
                o.unit_index == 1 && o.text.as_deref().is_some_and(|t| t.contains("pull quote"))
            })
            .unwrap();
        assert!(
            body.bounds.x > 400.0,
            "the pull quote must continue in the right frame"
        );
        let frames: Vec<_> = document
            .objects
            .iter()
            .filter(|o| {
                o.unit_index == 1
                    && (o.stable_id.ends_with("-67485") || o.stable_id.ends_with("-67361"))
            })
            .collect();
        assert_eq!(frames.len(), 2);
        assert_eq!(
            frames
                .iter()
                .map(|o| o.text.as_deref().unwrap().encode_utf16().count())
                .sum::<usize>(),
            1087,
            "linked frames must retain the complete story exactly once"
        );
        let Visual::TextLayout { visual, .. } = &body.visual else {
            panic!("missing body layout")
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("missing body runs")
        };
        let quote = runs.iter().find(|r| r.text.contains("pull quote")).unwrap();
        assert_eq!(quote.font_family, "Superclarendon");
        assert!(quote.bold);
        assert_eq!(quote.font_size, 24.0);
        let Visual::TextLayout { layout, .. } = &body.visual else {
            unreachable!()
        };
        assert_eq!(
            layout.paragraphs[0].margin_left + layout.paragraphs[0].first_line_indent,
            0.0,
            "the native quote first line starts at the frame edge, not at the continuation indent"
        );
    }

    #[test]
    fn preserves_newletter4_pull_quote_position() {
        let bytes = include_bytes!("../../tests/fixtures/newletter4.pages");
        let document = detect_and_parse(bytes, Limits::default()).unwrap().unwrap();
        let quote = document
            .objects
            .iter()
            .find(|o| {
                o.unit_index == 1 && o.text.as_deref().is_some_and(|t| t.contains("pull quote"))
            })
            .unwrap();
        assert!(
            (quote.bounds.y - 273.12 * 4.0 / 3.0).abs() < 0.01,
            "preserve the native quote position in points, got {:?}",
            quote.bounds
        );
        let body = document
            .objects
            .iter()
            .find(|o| {
                o.unit_index == 1
                    && o.text
                        .as_deref()
                        .is_some_and(|t| t.starts_with("can add additional"))
            })
            .unwrap();
        let Visual::TextLayout { layout, .. } = &body.visual else {
            panic!("missing flow layout")
        };
        assert!(
            layout
                .wrap_regions
                .iter()
                .any(|r| (r.x + body.bounds.x + layout.inset_left - quote.bounds.x).abs() < 0.01),
            "the quote must exclude body text at its saved position"
        );
        assert!(
            layout.wrap_regions.iter().any(|r| r.width >= 300.0),
            "the staircase must also exclude body text"
        );
    }

    #[test]
    fn preserves_newletter3_linked_text_page_ranges() {
        let bytes = include_bytes!("../../tests/fixtures/newletter3.pages");
        let document = detect_and_parse(bytes, Limits::default()).unwrap().unwrap();
        let first = document
            .objects
            .iter()
            .find(|o| {
                o.unit_index == 0
                    && o.text
                        .as_deref()
                        .is_some_and(|t| t.contains("To get started"))
            })
            .unwrap();
        assert_eq!(
            first.text.as_deref().unwrap().encode_utf16().count(),
            485,
            "the first linked frame must use its saved page range, not the entire story"
        );
        let second = document
            .objects
            .iter()
            .find(|o| {
                o.unit_index == 1
                    && o.text
                        .as_deref()
                        .is_some_and(|t| t.contains("You can use Pages"))
            })
            .unwrap();
        assert_eq!(second.text.as_deref().unwrap().encode_utf16().count(), 843);
        assert!(!second.text.as_deref().unwrap().starts_with("Main Heading"));
        let Visual::TextLayout { layout, .. } = &second.visual else {
            panic!("missing continuation layout")
        };
        assert_eq!(
            layout.paragraphs[1].space_before, 12.0,
            "24px before the heading must overlap the previous paragraph's 12px after"
        );
        assert_eq!(
            layout.vertical_overflow,
            crate::model::TextVerticalOverflow::Overflow,
            "a validated saved range must not lose lines to approximate font-box clipping"
        );
        assert!(
            !document
                .diagnostics
                .iter()
                .any(|d| d.message.contains("PAGES_TEXT_FLOW_APPROXIMATE"))
        );

        let target = protobuf_message(3, &[protobuf_varint(1, 0), protobuf_varint(2, 1)].concat());
        for flow_hint in [Vec::new(), protobuf_message(15, &[0xff])] {
            let page = [
                protobuf_message(2, &target),
                protobuf_message(14, &protobuf_reference(7)),
                flow_hint,
            ]
            .concat();
            let page = super::pages_parse_page_hint(&page, DOCUMENT_COMPONENT, &[])
                .unwrap()
                .unwrap();
            assert_eq!(
                page.targets.len(),
                1,
                "a damaged flow hint must not discard ordinary body text"
            );
            assert!(page.flow_targets.is_empty());
            assert_eq!(page.flow_diagnostics.len(), 1);
        }
    }

    #[test]
    fn preserves_newletter3_page_layout_headers() {
        let bytes = include_bytes!("../../tests/fixtures/newletter3.pages");
        let document = detect_and_parse(bytes, Limits::default()).unwrap().unwrap();
        let caption = document.objects.iter().find(|o| {
            o.unit_index == 1
                && o.text
                    .as_deref()
                    .is_some_and(|t| t.contains("replace this image caption"))
        });
        assert!(
            caption.is_some(),
            "a missing image resource must not erase its stored caption"
        );
        let title = document
            .objects
            .iter()
            .find(|o| o.text.as_deref() == Some("NEWSLETTER"))
            .unwrap();
        let Visual::TextLayout { layout, .. } = &title.visual else {
            panic!("missing title layout")
        };
        assert!(layout.paragraphs[0].rule_above.is_some());
        assert!(layout.paragraphs[0].rule_below.is_some());
        let mut border = super::KeynoteTextStyle {
            border_positions: Some(3),
            ..Default::default()
        };
        super::keynote_paragraph_properties(
            &protobuf_varint(15, 0),
            DOCUMENT_COMPONENT,
            &mut border,
        )
        .unwrap();
        assert!(
            border.rule_above().is_none(),
            "an explicit legacy reset must clear an inherited modern border"
        );
        super::keynote_paragraph_properties(
            &[protobuf_varint(45, 3), protobuf_varint(15, 0)].concat(),
            DOCUMENT_COMPONENT,
            &mut border,
        )
        .unwrap();
        assert_eq!(
            border.border_positions,
            Some(3),
            "the modern field wins regardless of wire order"
        );
        for page in 0..2 {
            assert!(
                document.objects.iter().any(|o| o.unit_index == page
                    && o.text
                        .as_deref()
                        .is_some_and(|t| t.contains("SEPTEMBER 2, 2026"))),
                "page-layout documents must retain their document section headers"
            );
            let issue = document
                .objects
                .iter()
                .find(|o| {
                    o.unit_index == page && o.text.as_deref().is_some_and(|t| t.contains("ISSUE 1"))
                })
                .unwrap();
            let Visual::TextLayout { layout, .. } = &issue.visual else {
                panic!("missing header layout")
            };
            assert_eq!(
                layout.paragraphs[0].align,
                crate::model::TextAlign::End,
                "the right header slot must resolve natural alignment to the right edge"
            );
            let number = (page + 1).to_string();
            let footer = document
                .objects
                .iter()
                .find(|o| o.unit_index == page && o.text.as_deref() == Some(number.as_str()))
                .unwrap();
            let Visual::TextLayout { layout, .. } = &footer.visual else {
                panic!("missing footer layout")
            };
            assert_eq!(layout.paragraphs[0].align, crate::model::TextAlign::End);
        }
        let mut style = super::KeynoteTextStyle::default();
        for value in [4, 0, 1, 2, 3] {
            super::keynote_paragraph_properties(
                &protobuf_varint(1, value),
                DOCUMENT_COMPONENT,
                &mut style,
            )
            .unwrap();
            assert_eq!(
                style.natural_alignment,
                value == 4,
                "explicit paragraph alignment must override the header slot default"
            );
        }
    }

    #[test]
    fn preserves_newletter3_body_regular_weight() {
        let bytes = include_bytes!("../../tests/fixtures/newletter3.pages");
        let document = detect_and_parse(bytes, Limits::default()).unwrap().unwrap();
        let body = document
            .objects
            .iter()
            .find(|o| {
                o.unit_index == 0
                    && o.text
                        .as_deref()
                        .is_some_and(|t| t.contains("To get started"))
            })
            .unwrap();
        let Visual::TextLayout { visual, .. } = &body.visual else {
            panic!("missing body layout")
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("missing body runs")
        };
        let run = runs
            .iter()
            .find(|r| r.text.contains("To get started"))
            .unwrap();
        assert!(
            !run.bold && !run.font_family.contains("Bold"),
            "Georgia Regular must not be bold"
        );
        assert_eq!(run.font_family, "Georgia");
        for flag in [None, Some(1), Some(0)] {
            let mut style = super::KeynoteTextStyle::default();
            let properties = [
                flag.map(|value| protobuf_varint(1, value))
                    .unwrap_or_default(),
                protobuf_message(5, b"Georgia-Bold"),
            ]
            .concat();
            super::keynote_character_style(&properties, DOCUMENT_COMPONENT, &mut style).unwrap();
            assert_eq!(
                style.font_family,
                if flag == Some(0) {
                    "Georgia"
                } else {
                    "Georgia-Bold"
                }
            );
        }
    }

    #[test]
    fn preserves_untitled_authored_all_caps() {
        let bytes = include_bytes!("../../tests/fixtures/untitled.pages");
        let document = detect_and_parse(bytes, Limits::default()).unwrap().unwrap();
        let title = document
            .objects
            .iter()
            .find(|o| o.unit_index == 0 && o.stable_id.ends_with("-58714"))
            .unwrap();
        assert_eq!(title.text.as_deref().unwrap().trim(), "TITLE");
        let mut style = super::KeynoteTextStyle::default();
        super::keynote_character_style(&protobuf_varint(13, 1), DOCUMENT_COMPONENT, &mut style)
            .unwrap();
        assert_eq!(
            super::keynote_text_run("Straße\n中文".to_owned(), &style).text,
            "STRASSE\n中文"
        );
        super::keynote_character_style(&protobuf_varint(13, 0), DOCUMENT_COMPONENT, &mut style)
            .unwrap();
        assert_eq!(
            super::keynote_text_run("Title".to_owned(), &style).text,
            "Title"
        );
    }

    #[test]
    fn preserves_untitled_missing_texture_reference_color() {
        let bytes = include_bytes!("../../tests/fixtures/untitled.pages");
        let document = detect_and_parse(bytes, Limits::default()).unwrap().unwrap();
        for id in [58407, 58635, 58532] {
            let background = document
                .objects
                .iter()
                .find(|o| o.unit_index == 0 && o.stable_id.ends_with(&format!("-{id}")))
                .expect("missing texture must retain the authored reference color");
            assert!(matches!(
                &background.visual,
                Visual::PaintedShape {
                    fill: Paint::Solid(0xede7dcff),
                    ..
                }
            ));
        }
        assert!(document.diagnostics.iter().any(|d| {
            d.message
                .contains("IWORK_IMAGE_FILL_UNAVAILABLE: image data 10")
        }));
        let package = crate::package::Package::open_iwork(bytes, Limits::default()).unwrap();
        let data_files = super::iwork_data_files(&package, Limits::default()).unwrap();
        let mut style = super::KeynoteDrawableStyle {
            fill: Paint::Solid(0xede7dcff),
            image_fill: Some((14, false, None)),
            ..super::KeynoteDrawableStyle::default()
        };
        let mut diagnostics = Vec::new();
        super::keynote_materialize_image_fill(
            &mut diagnostics,
            &package,
            &data_files,
            &mut style,
            Limits::default(),
            DOCUMENT_COMPONENT,
        )
        .unwrap();
        assert!(
            matches!(style.fill, Paint::Image { .. }),
            "packaged image takes precedence over reference color"
        );
        assert!(diagnostics.is_empty());
    }
}
