//! Native ODP presentation parsing.

use super::optional_xml_attribute as optional_attribute;
use std::collections::{HashMap, HashSet};

use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
use crate::font_metrics::FontMetricTable;
#[cfg(feature = "legacy-office-formats")]
use crate::limits::Limits;
use crate::model::{
    AffineTransform, Document, DocumentFormat, DocumentKind, FillRule, Geometry, GradientStop,
    ImageCrop, LineCap, MappingQuality, Object, ObjectKind, Paint, PathCommand, PathFillMode,
    PathLayer, Rect, Shadow, SlideMetadata, SourceLocator, SourceRef, StrokeStyle, TextAlign,
    TextAutoFit, TextLayout, TextOrientation, TextParagraphLayout, TextRun, TextVerticalAlign,
    Unit, UnitKind, Visual,
};
use crate::package::Package;
use crate::xml::{XmlAttribute, XmlEvent, decode_xml_text, parse_xml};

use super::embedded_media::{
    EmbeddedMediaError, embedded_media_type, embedded_media_type_from_mime,
};
use super::local_name;
use super::odf_math::{OdfMath, parse_odf_math};
use super::presentation_image::{
    ImageCacheEntry, OdfImageTarget, OfficeImageError, clone_image_bytes, office_image_media_type,
    office_image_media_type_from_mime, office_image_media_type_from_signature,
    reserve_materialized_image_bytes, resolve_odf_image_target,
};

const STYLES_PART: &str = "styles.xml";
const CONTENT_PART: &str = "content.xml";
const MANIFEST_PART: &str = "META-INF/manifest.xml";
const EMPTY_TABLE_TEMPLATE: &str = "{00000000-0000-0000-0000-000000000000}";
const CSS_PIXELS_PER_INCH: f32 = 96.0;

trait OdpSource {
    fn limits(&self) -> crate::limits::Limits;
    fn has_part(&self, name: &str) -> bool;
    fn part(&self, name: &str) -> Result<Option<crate::package::PartBytes>, Diagnostic>;
    fn required_part(&self, name: &str) -> Result<crate::package::PartBytes, Diagnostic>;
    fn is_flat(&self) -> bool {
        false
    }
}

impl OdpSource for Package<'_> {
    fn limits(&self) -> crate::limits::Limits {
        self.limits()
    }

    fn has_part(&self, name: &str) -> bool {
        self.has_part(name)
    }

    fn part(&self, name: &str) -> Result<Option<crate::package::PartBytes>, Diagnostic> {
        self.part(name)
    }

    fn required_part(&self, name: &str) -> Result<crate::package::PartBytes, Diagnostic> {
        self.required_part(name)
    }
}

#[cfg(feature = "odf-formats")]
struct FlatOdpSource {
    bytes: std::sync::Arc<[u8]>,
    limits: crate::limits::Limits,
}

#[cfg(feature = "odf-formats")]
impl OdpSource for FlatOdpSource {
    fn limits(&self) -> crate::limits::Limits {
        self.limits
    }

    fn has_part(&self, _name: &str) -> bool {
        false
    }

    fn part(&self, name: &str) -> Result<Option<crate::package::PartBytes>, Diagnostic> {
        Ok(matches!(name, STYLES_PART | CONTENT_PART)
            .then(|| crate::package::PartBytes::Shared(self.bytes.clone())))
    }

    fn required_part(&self, name: &str) -> Result<crate::package::PartBytes, Diagnostic> {
        self.part(name)?
            .ok_or_else(|| format_error(name, format!("required flat ODF part is missing: {name}")))
    }

    fn is_flat(&self) -> bool {
        true
    }
}

fn style_parts(source: &dyn OdpSource) -> &'static [&'static str] {
    if source.is_flat() || !source.has_part(STYLES_PART) {
        &[CONTENT_PART]
    } else {
        &[STYLES_PART, CONTENT_PART]
    }
}

#[derive(Clone, Copy, Debug)]
struct PageSize {
    width: f32,
    height: f32,
}

#[derive(Clone, Debug)]
enum OdfPaint {
    None,
    Solid(u32),
    Gradient(String),
    Bitmap(String),
}

#[derive(Clone, Copy, Debug)]
enum OdfLineHeight {
    Multiple(f32),
    Exact(f32),
}

#[derive(Clone, Copy, Debug)]
enum OdfListFontSize {
    Relative(f32),
    Exact(f32),
}

#[derive(Clone, Copy, Debug)]
enum OdfDashLength {
    Absolute(f32),
    StrokeWidthMultiple(f32),
}

#[derive(Clone, Copy, Debug)]
struct OdfImageClip {
    left: f32,
    top: f32,
    right: f32,
    bottom: f32,
}

#[derive(Clone, Debug)]
struct GraphicStyle {
    fill: OdfPaint,
    stroke: OdfPaint,
    stroke_width: f32,
    opacity: f32,
    font_family: String,
    font_size: f32,
    color: u32,
    bold: bool,
    italic: bool,
    align: TextAlign,
    orientation: TextOrientation,
    vertical_align: TextVerticalAlign,
    auto_fit: TextAutoFit,
    wrap: bool,
    columns: (u32, f32),
    inset_left: f32,
    inset_right: f32,
    inset_top: f32,
    inset_bottom: f32,
    marker_start: bool,
    marker_end: bool,
    stroke_dash: Option<String>,
    line_cap: Option<LineCap>,
    border_style: Option<StrokeStyle>,
    shadow: Option<Shadow>,
    margin_left: f32,
    margin_right: f32,
    first_line_indent: f32,
    default_tab_stop: f32,
    paragraph_spacing: f32,
    line_height: Option<OdfLineHeight>,
    image_clip: Option<OdfImageClip>,
    fill_image_tile: bool,
    fill_image_width: Option<f32>,
    fill_image_height: Option<f32>,
}

impl Default for GraphicStyle {
    fn default() -> Self {
        Self {
            fill: OdfPaint::None,
            stroke: OdfPaint::None,
            stroke_width: 0.0,
            opacity: 1.0,
            font_family: "Arial".to_owned(),
            font_size: 18.0 * CSS_PIXELS_PER_INCH / 72.0,
            color: 0x0000_00ff,
            bold: false,
            italic: false,
            align: TextAlign::Start,
            orientation: TextOrientation::Horizontal,
            vertical_align: TextVerticalAlign::Top,
            auto_fit: TextAutoFit::None,
            wrap: true,
            columns: (1, 0.0),
            inset_left: 4.0,
            inset_right: 4.0,
            inset_top: 4.0,
            inset_bottom: 4.0,
            marker_start: false,
            marker_end: false,
            stroke_dash: None,
            line_cap: None,
            border_style: None,
            shadow: None,
            margin_left: 0.0,
            margin_right: 0.0,
            first_line_indent: 0.0,
            default_tab_stop: 36.0,
            paragraph_spacing: 0.0,
            line_height: None,
            image_clip: None,
            fill_image_tile: true,
            fill_image_width: None,
            fill_image_height: None,
        }
    }
}

#[derive(Clone, Debug, Default)]
struct GraphicStylePatch {
    fill: Option<OdfPaint>,
    fill_color: Option<u32>,
    stroke: Option<OdfPaint>,
    stroke_width: Option<f32>,
    opacity: Option<f32>,
    font_family: Option<String>,
    font_size: Option<f32>,
    color: Option<u32>,
    bold: Option<bool>,
    italic: Option<bool>,
    align: Option<TextAlign>,
    orientation: Option<TextOrientation>,
    vertical_align: Option<TextVerticalAlign>,
    auto_fit: Option<TextAutoFit>,
    wrap: Option<bool>,
    columns: Option<(u32, f32)>,
    inset_left: Option<f32>,
    inset_right: Option<f32>,
    inset_top: Option<f32>,
    inset_bottom: Option<f32>,
    marker_start: Option<bool>,
    marker_end: Option<bool>,
    stroke_dash: Option<Option<String>>,
    line_cap: Option<LineCap>,
    border_style: Option<Option<StrokeStyle>>,
    shadow: Option<Option<Shadow>>,
    margin_left: Option<f32>,
    margin_right: Option<f32>,
    first_line_indent: Option<f32>,
    default_tab_stop: Option<f32>,
    paragraph_spacing: Option<f32>,
    line_height: Option<OdfLineHeight>,
    image_clip: Option<OdfImageClip>,
    fill_image_tile: Option<bool>,
    fill_image_width: Option<f32>,
    fill_image_height: Option<f32>,
}

impl GraphicStylePatch {
    fn apply_to(&self, style: &mut GraphicStyle) {
        if let Some(fill) = self.fill.as_ref() {
            style.fill = fill.clone();
        }
        if let Some(fill_color) = self.fill_color
            && let OdfPaint::Solid(color) = &mut style.fill
        {
            *color = fill_color;
        }
        if let Some(stroke) = self.stroke.as_ref() {
            style.stroke = stroke.clone();
        }
        if let Some(stroke_width) = self.stroke_width {
            style.stroke_width = stroke_width;
        }
        if let Some(opacity) = self.opacity {
            style.opacity = opacity;
        }
        if let Some(font_family) = self.font_family.as_ref() {
            style.font_family = font_family.clone();
        }
        if let Some(font_size) = self.font_size {
            style.font_size = font_size;
        }
        if let Some(color) = self.color {
            style.color = color;
        }
        if let Some(bold) = self.bold {
            style.bold = bold;
        }
        if let Some(italic) = self.italic {
            style.italic = italic;
        }
        if let Some(align) = self.align {
            style.align = align;
        }
        if let Some(orientation) = self.orientation {
            style.orientation = orientation;
        }
        if let Some(vertical_align) = self.vertical_align {
            style.vertical_align = vertical_align;
        }
        if let Some(auto_fit) = self.auto_fit {
            style.auto_fit = auto_fit;
        }
        if let Some(columns) = self.columns {
            style.columns = columns;
        }
        if let Some(wrap) = self.wrap {
            style.wrap = wrap;
        }
        if let Some(inset_left) = self.inset_left {
            style.inset_left = inset_left;
        }
        if let Some(inset_right) = self.inset_right {
            style.inset_right = inset_right;
        }
        if let Some(inset_top) = self.inset_top {
            style.inset_top = inset_top;
        }
        if let Some(inset_bottom) = self.inset_bottom {
            style.inset_bottom = inset_bottom;
        }
        if let Some(marker_start) = self.marker_start {
            style.marker_start = marker_start;
        }
        if let Some(marker_end) = self.marker_end {
            style.marker_end = marker_end;
        }
        if let Some(stroke_dash) = self.stroke_dash.as_ref() {
            style.stroke_dash.clone_from(stroke_dash);
        }
        if let Some(line_cap) = self.line_cap {
            style.line_cap = Some(line_cap);
        }
        if let Some(border_style) = self.border_style.as_ref() {
            style.border_style.clone_from(border_style);
        }
        if let Some(shadow) = self.shadow.as_ref() {
            style.shadow = *shadow;
        }
        if let Some(margin_left) = self.margin_left {
            style.margin_left = margin_left;
        }
        if let Some(margin_right) = self.margin_right {
            style.margin_right = margin_right;
        }
        if let Some(first_line_indent) = self.first_line_indent {
            style.first_line_indent = first_line_indent;
        }
        if let Some(default_tab_stop) = self.default_tab_stop {
            style.default_tab_stop = default_tab_stop;
        }
        if let Some(paragraph_spacing) = self.paragraph_spacing {
            style.paragraph_spacing = paragraph_spacing;
        }
        if let Some(line_height) = self.line_height {
            style.line_height = Some(line_height);
        }
        if let Some(image_clip) = self.image_clip {
            style.image_clip = Some(image_clip);
        }
        if let Some(fill_image_tile) = self.fill_image_tile {
            style.fill_image_tile = fill_image_tile;
        }
        if let Some(fill_image_width) = self.fill_image_width {
            style.fill_image_width = Some(fill_image_width);
        }
        if let Some(fill_image_height) = self.fill_image_height {
            style.fill_image_height = Some(fill_image_height);
        }
        if !matches!(style.stroke, OdfPaint::None) && style.stroke_width <= 0.0 {
            style.stroke_width = 1.0;
        }
    }

    fn merge(&mut self, patch: Self) {
        if patch.fill.is_some() {
            self.fill = patch.fill;
        }
        if patch.fill_color.is_some() {
            self.fill_color = patch.fill_color;
        }
        if patch.stroke.is_some() {
            self.stroke = patch.stroke;
        }
        if patch.stroke_width.is_some() {
            self.stroke_width = patch.stroke_width;
        }
        if patch.opacity.is_some() {
            self.opacity = patch.opacity;
        }
        if patch.font_family.is_some() {
            self.font_family = patch.font_family;
        }
        if patch.font_size.is_some() {
            self.font_size = patch.font_size;
        }
        if patch.color.is_some() {
            self.color = patch.color;
        }
        if patch.bold.is_some() {
            self.bold = patch.bold;
        }
        if patch.italic.is_some() {
            self.italic = patch.italic;
        }
        if patch.align.is_some() {
            self.align = patch.align;
        }
        if patch.vertical_align.is_some() {
            self.vertical_align = patch.vertical_align;
        }
        if patch.auto_fit.is_some() {
            self.auto_fit = patch.auto_fit;
        }
        if patch.wrap.is_some() {
            self.wrap = patch.wrap;
        }
        if patch.inset_left.is_some() {
            self.inset_left = patch.inset_left;
        }
        if patch.inset_right.is_some() {
            self.inset_right = patch.inset_right;
        }
        if patch.inset_top.is_some() {
            self.inset_top = patch.inset_top;
        }
        if patch.inset_bottom.is_some() {
            self.inset_bottom = patch.inset_bottom;
        }
        if patch.marker_start.is_some() {
            self.marker_start = patch.marker_start;
        }
        if patch.marker_end.is_some() {
            self.marker_end = patch.marker_end;
        }
        if patch.stroke_dash.is_some() {
            self.stroke_dash = patch.stroke_dash;
        }
        if patch.line_cap.is_some() {
            self.line_cap = patch.line_cap;
        }
        if patch.border_style.is_some() {
            self.border_style = patch.border_style;
        }
        if patch.shadow.is_some() {
            self.shadow = patch.shadow;
        }
        if patch.margin_left.is_some() {
            self.margin_left = patch.margin_left;
        }
        if patch.margin_right.is_some() {
            self.margin_right = patch.margin_right;
        }
        if patch.first_line_indent.is_some() {
            self.first_line_indent = patch.first_line_indent;
        }
        if patch.default_tab_stop.is_some() {
            self.default_tab_stop = patch.default_tab_stop;
        }
        if patch.paragraph_spacing.is_some() {
            self.paragraph_spacing = patch.paragraph_spacing;
        }
        if patch.line_height.is_some() {
            self.line_height = patch.line_height;
        }
        if patch.image_clip.is_some() {
            self.image_clip = patch.image_clip;
        }
        if patch.fill_image_tile.is_some() {
            self.fill_image_tile = patch.fill_image_tile;
        }
        if patch.fill_image_width.is_some() {
            self.fill_image_width = patch.fill_image_width;
        }
        if patch.fill_image_height.is_some() {
            self.fill_image_height = patch.fill_image_height;
        }
    }
}

#[derive(Clone, Debug)]
struct StyleDefinition {
    family: String,
    parent: Option<String>,
    patch: GraphicStylePatch,
    part: String,
}

#[derive(Debug)]
struct PendingStyle {
    depth: usize,
    name: Option<String>,
    family: String,
    parent: Option<String>,
    patch: GraphicStylePatch,
}

#[derive(Clone, Copy, Debug)]
struct GradientDefinition {
    start: u32,
    end: u32,
    angle_degrees: f32,
    radial: bool,
    rectangular: bool,
    center_x: f32,
    center_y: f32,
}

#[derive(Clone, Debug)]
enum OdfListLabel {
    Bullet(String),
    Image {
        href: String,
        width: f32,
        height: f32,
    },
    Number {
        format: String,
        prefix: String,
        suffix: String,
        start_value: Option<u32>,
    },
}

#[derive(Clone, Debug)]
struct OdfListLevelStyle {
    label: OdfListLabel,
    space_before: f32,
    min_label_width: f32,
    font_family: Option<String>,
    label_style: OdfListLabelStyle,
}

#[derive(Clone, Copy, Debug, Default)]
struct OdfListLabelStyle {
    color: Option<u32>,
    font_size: Option<OdfListFontSize>,
}

#[derive(Clone, Debug)]
struct OdfTextRun {
    text: String,
    style_name: Option<String>,
    label_style: Option<OdfListLabelStyle>,
}

#[derive(Clone, Debug)]
struct OdfListParagraphLayout {
    margin_left: f32,
    first_line_indent: f32,
    image: Option<(String, f32, f32)>,
    outline_style: Option<String>,
}

#[derive(Clone, Debug, Default)]
struct OdfTableTemplate {
    body: Option<String>,
    first_row: Option<String>,
    last_row: Option<String>,
    first_column: Option<String>,
    last_column: Option<String>,
    even_rows: Option<String>,
    odd_rows: Option<String>,
    even_columns: Option<String>,
    odd_columns: Option<String>,
}

#[derive(Default)]
struct StyleCatalog {
    font_faces: HashMap<String, String>,
    graphics: HashMap<String, GraphicStyle>,
    style_patches: HashMap<String, GraphicStylePatch>,
    outline_prefixes: HashMap<String, String>,
    gradients: HashMap<String, GradientDefinition>,
    opacity_gradients: HashMap<String, GradientDefinition>,
    fill_images: HashMap<String, OdfImageData>,
    diagnostics: Vec<Diagnostic>,
    stroke_dashes: HashMap<String, (Vec<OdfDashLength>, LineCap)>,
    list_styles: HashMap<String, HashMap<u32, OdfListLevelStyle>>,
    table_column_widths: HashMap<String, f32>,
    table_row_heights: HashMap<String, f32>,
    table_templates: HashMap<String, OdfTableTemplate>,
    masters: HashMap<String, Vec<MasterTemplate>>,
    master_backgrounds: HashMap<String, String>,
    invalid_master_shapes: usize,
}

#[derive(Clone, Debug)]
struct MasterTemplate {
    element_id: Option<String>,
    bounds: Rect,
    transform: AffineTransform,
    geometry: Geometry,
    text_path: Option<String>,
    style_name: Option<String>,
    text: String,
    text_runs: Vec<OdfTextRun>,
    paragraph_style_names: Vec<Option<String>>,
    image: Option<OdfImageData>,
    path: String,
}

fn odf_named_shape_geometry(name: Option<&str>, bounds: Rect) -> Option<Geometry> {
    let name = name?.to_lowercase();
    if ["oval", "ellipse", "circle", "楕円", "圆", "圓"]
        .iter()
        .any(|candidate| name.contains(candidate))
    {
        Some(Geometry::Ellipse)
    } else if [
        "rounded rectangle",
        "round rectangle",
        "roundrect",
        "rounded corners",
        "corners rounded",
        "角を丸く",
        "alternate process",
    ]
    .iter()
    .any(|candidate| name.contains(candidate))
    {
        Some(Geometry::RoundedRectangle {
            radius_x: bounds.width.min(bounds.height) * 0.12,
            radius_y: bounds.width.min(bounds.height) * 0.12,
        })
    } else {
        None
    }
}

#[derive(Clone, Debug)]
struct OdfImageData {
    media_type: String,
    bytes: Vec<u8>,
    target: String,
}

pub fn parse(package: &Package<'_>) -> Result<Document, Diagnostic> {
    parse_with_font_metrics(package, &FontMetricTable::default())
}

pub fn parse_with_font_metrics(
    package: &Package<'_>,
    font_metrics: &FontMetricTable,
) -> Result<Document, Diagnostic> {
    parse_source(package, font_metrics)
}

#[cfg(feature = "odf-formats")]
pub(super) fn parse_flat(
    bytes: &[u8],
    limits: crate::limits::Limits,
    font_metrics: &FontMetricTable,
) -> Result<Document, Diagnostic> {
    let source = FlatOdpSource {
        bytes: bytes.to_vec().into(),
        limits,
    };
    let mut document = parse_source(&source, font_metrics)?;
    document
        .diagnostics
        .extend(super::security::inspect_flat_odf(bytes, limits)?);
    Ok(document)
}

fn parse_source(
    package: &dyn OdpSource,
    font_metrics: &FontMetricTable,
) -> Result<Document, Diagnostic> {
    let page_sizes = parse_page_sizes(package)?;
    let styles = parse_style_catalog(package)?;
    let content = parse_content_with_font_metrics(package, &page_sizes, &styles, font_metrics)?;
    if content.units.is_empty() {
        return Err(format_error(
            CONTENT_PART,
            "presentation contains no slides",
        ));
    }

    Ok(Document {
        fatal: false,
        format: Some(DocumentFormat::Odp),
        kind: Some(DocumentKind::Presentation),
        units: content.units,
        outline: Vec::new(),
        objects: content.objects,
        embedded_fonts: Vec::new(),
        font_alternate_names: Vec::new(),
        diagnostics: content.diagnostics,
    })
}

fn parse_page_sizes(package: &dyn OdpSource) -> Result<HashMap<String, PageSize>, Diagnostic> {
    let bytes = package.required_part(style_parts(package)[0])?;
    let mut current_layout = None;
    let mut layouts = HashMap::new();
    let mut masters = Vec::new();

    parse_xml(&bytes, package.limits(), |event| {
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
                    let layout_name = current_layout.as_ref().ok_or_else(|| {
                        format_error(STYLES_PART, "page layout properties have no parent layout")
                    })?;
                    let width = required_length(&attributes, "page-width", STYLES_PART)?;
                    let height = required_length(&attributes, "page-height", STYLES_PART)?;
                    let size = PageSize { width, height };
                    if !valid_page_size(size) {
                        // The content parser reports the missing usable layout and supplies
                        // the normal page-size fallback without losing other pages.
                        return Ok(());
                    }
                    if layouts.insert(layout_name.clone(), size).is_some() {
                        return Err(format_error(
                            STYLES_PART,
                            format!("duplicate page layout properties: {layout_name}"),
                        ));
                    }
                }
                "master-page" => {
                    let name = required_attribute(&attributes, "name", STYLES_PART)?;
                    let layout = required_attribute(&attributes, "page-layout-name", STYLES_PART)?;
                    masters.push((name, layout));
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

    let mut page_sizes = HashMap::new();
    for (master, layout) in masters {
        let Some(size) = layouts.get(&layout).copied() else {
            continue;
        };
        if page_sizes.insert(master.clone(), size).is_some() {
            return Err(format_error(
                STYLES_PART,
                format!("duplicate master page: {master}"),
            ));
        }
    }
    Ok(page_sizes)
}

fn parse_style_catalog(package: &dyn OdpSource) -> Result<StyleCatalog, Diagnostic> {
    let mut catalog = StyleCatalog::default();
    let mut definitions = HashMap::new();
    let mut defaults = HashMap::new();
    let manifest_media_types = parse_manifest_media_types(package)?;
    for &part in style_parts(package) {
        let bytes = package.required_part(part)?;
        parse_style_part(
            &bytes,
            package,
            part,
            &mut catalog,
            &mut definitions,
            &mut defaults,
            &manifest_media_types,
        )?;
    }
    catalog.list_styles = parse_list_styles(package)?;
    let (column_widths, row_heights) = parse_table_style_dimensions(package)?;
    catalog.table_column_widths = column_widths;
    catalog.table_row_heights = row_heights;
    catalog.table_templates = parse_table_templates(package)?;
    let mut resolved = HashMap::new();
    let mut resolved_patches = HashMap::new();
    for name in definitions.keys() {
        resolve_graphic_style(
            name,
            &definitions,
            &defaults,
            &mut resolved,
            &mut HashSet::new(),
        )?;
        resolve_graphic_style_patch(
            name,
            &definitions,
            &mut resolved_patches,
            &mut HashSet::new(),
        )?;
    }
    for name in definitions.keys() {
        let mut ancestor = name.as_str();
        for _ in 0..definitions.len() {
            if let Some(prefix) = ancestor.strip_suffix("outline1") {
                catalog
                    .outline_prefixes
                    .insert(name.clone(), format!("{prefix}outline"));
                break;
            }
            let Some(parent) = definitions.get(ancestor).and_then(|d| d.parent.as_deref()) else {
                break;
            };
            ancestor = parent;
        }
    }
    catalog.graphics = resolved;
    catalog.style_patches = resolved_patches;
    let (masters, invalid_master_shapes) = parse_master_templates(package)?;
    catalog.masters = masters;
    catalog.invalid_master_shapes = invalid_master_shapes;
    Ok(catalog)
}

fn resolve_graphic_style(
    name: &str,
    definitions: &HashMap<String, StyleDefinition>,
    defaults: &HashMap<String, GraphicStylePatch>,
    resolved: &mut HashMap<String, GraphicStyle>,
    visiting: &mut HashSet<String>,
) -> Result<GraphicStyle, Diagnostic> {
    if let Some(style) = resolved.get(name) {
        return Ok(style.clone());
    }
    let definition = definitions
        .get(name)
        .cloned()
        .ok_or_else(|| format_error(CONTENT_PART, format!("unknown ODP style {name}")))?;
    if !visiting.insert(name.to_owned()) {
        return Err(format_error(
            &definition.part,
            format!("ODP style inheritance contains a cycle at {name}"),
        ));
    }
    let mut style = if let Some(parent) = definition
        .parent
        .as_deref()
        .filter(|parent| definitions.contains_key(*parent))
    {
        resolve_graphic_style(parent, definitions, defaults, resolved, visiting)?
    } else {
        let mut style = GraphicStyle::default();
        if let Some(default) = defaults.get(&definition.family) {
            default.apply_to(&mut style);
        }
        style
    };
    definition.patch.apply_to(&mut style);
    visiting.remove(name);
    resolved.insert(name.to_owned(), style.clone());
    Ok(style)
}

fn resolve_graphic_style_patch(
    name: &str,
    definitions: &HashMap<String, StyleDefinition>,
    resolved: &mut HashMap<String, GraphicStylePatch>,
    visiting: &mut HashSet<String>,
) -> Result<GraphicStylePatch, Diagnostic> {
    if let Some(patch) = resolved.get(name) {
        return Ok(patch.clone());
    }
    let definition = definitions
        .get(name)
        .cloned()
        .ok_or_else(|| format_error(CONTENT_PART, format!("unknown ODP style {name}")))?;
    if !visiting.insert(name.to_owned()) {
        return Err(format_error(
            &definition.part,
            format!("ODP style inheritance contains a cycle at {name}"),
        ));
    }
    let mut patch = if let Some(parent) = definition
        .parent
        .as_deref()
        .filter(|parent| definitions.contains_key(*parent))
    {
        resolve_graphic_style_patch(parent, definitions, resolved, visiting)?
    } else {
        GraphicStylePatch::default()
    };
    patch.merge(definition.patch);
    visiting.remove(name);
    resolved.insert(name.to_owned(), patch.clone());
    Ok(patch)
}

fn parse_master_templates(
    package: &dyn OdpSource,
) -> Result<(HashMap<String, Vec<MasterTemplate>>, usize), Diagnostic> {
    #[derive(Debug)]
    struct MasterState {
        depth: usize,
        index: usize,
        name: String,
        templates: Vec<MasterTemplate>,
    }
    #[derive(Debug)]
    struct TemplateState {
        depth: usize,
        template: MasterTemplate,
        enhanced_geometry: Option<OdfEnhancedGeometryState>,
        paragraph_depth: Option<usize>,
        paragraph_count: u32,
        span_depth: Option<usize>,
        span_style_name: Option<String>,
    }

    let bytes = package.required_part(style_parts(package)[0])?;
    let manifest_media_types = parse_manifest_media_types(package)?;
    let mut depth = 0_usize;
    let mut master_count = 0_usize;
    let mut master: Option<MasterState> = None;
    let mut template: Option<TemplateState> = None;
    let mut notes_depth = None;
    let mut ignored_template_depth = None;
    let mut invalid_master_shapes = 0_usize;
    let mut masters = HashMap::new();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                if ignored_template_depth.is_some() {
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }
                let local = local_name(name);
                if local == "master-page" && master.is_none() {
                    let name = required_attribute(&attributes, "name", STYLES_PART)?;
                    master_count = master_count.saturating_add(1);
                    if empty {
                        masters.insert(name, Vec::new());
                    } else {
                        master = Some(MasterState {
                            depth,
                            index: master_count,
                            name,
                            templates: Vec::new(),
                        });
                    }
                } else if local == "notes" && master.is_some() {
                    notes_depth = (!empty).then_some(depth);
                } else if notes_depth.is_none()
                    && master
                    .as_ref()
                    .is_some_and(|master| depth > master.depth)
                    && template.is_none()
                    && matches!(
                        local,
                        "frame" | "rect" | "ellipse" | "line" | "custom-shape"
                    )
                    && optional_attribute(&attributes, "class", STYLES_PART)?.is_none()
                {
                    let bounds = if local == "line" {
                        let x1 = optional_length(&attributes, "x1", STYLES_PART)?.unwrap_or(0.0);
                        let y1 = optional_length(&attributes, "y1", STYLES_PART)?.unwrap_or(0.0);
                        let x2 = optional_length(&attributes, "x2", STYLES_PART)?.unwrap_or(x1);
                        let y2 = optional_length(&attributes, "y2", STYLES_PART)?.unwrap_or(y1);
                        Rect {
                            x: x1.min(x2),
                            y: y1.min(y2),
                            width: (x2 - x1).abs(),
                            height: (y2 - y1).abs(),
                        }
                    } else {
                        Rect {
                            x: optional_length(&attributes, "x", STYLES_PART)?.unwrap_or(0.0),
                            y: optional_length(&attributes, "y", STYLES_PART)?.unwrap_or(0.0),
                            width: optional_length(&attributes, "width", STYLES_PART)?
                                .unwrap_or(0.0),
                            height: optional_length(&attributes, "height", STYLES_PART)?
                                .unwrap_or(0.0),
                        }
                    };
                    if !bounds.is_valid() {
                        invalid_master_shapes = invalid_master_shapes.saturating_add(1);
                        ignored_template_depth = (!empty).then_some(depth);
                        if !empty {
                            depth = depth.saturating_add(1);
                        }
                        return Ok(());
                    }
                    let index = master
                        .as_ref()
                        .map(|master| master.templates.len())
                        .unwrap_or(0);
                    let master_index = master.as_ref().map(|master| master.index).unwrap_or(1);
                    let current = TemplateState {
                        depth,
                        template: MasterTemplate {
                            element_id: optional_attribute(&attributes, "id", STYLES_PART)?,
                            bounds,
                            transform: optional_attribute(
                                &attributes,
                                "transform",
                                STYLES_PART,
                            )?
                            .map(|value| parse_odf_master_transform(&value))
                            .transpose()?
                            .unwrap_or(AffineTransform::IDENTITY),
                            geometry: match local {
                                "ellipse" => Geometry::Ellipse,
                                "line" => Geometry::Line,
                                "custom-shape" => odf_named_shape_geometry(
                                    optional_attribute(&attributes, "name", STYLES_PART)?.as_deref(),
                                    bounds,
                                )
                                .unwrap_or(Geometry::Rectangle),
                                _ => Geometry::Rectangle,
                            },
                            text_path: None,
                            style_name: optional_attribute(
                                &attributes,
                                "style-name",
                                STYLES_PART,
                            )?,
                            text: String::new(),
                            text_runs: Vec::new(),
                            paragraph_style_names: Vec::new(),
                            image: None,
                            path: format!("/office:document-styles/office:master-styles/style:master-page[{master_index}]/draw:{local}[{}]", index + 1),
                        },
                        enhanced_geometry: None,
                        paragraph_depth: None,
                        paragraph_count: 0,
                        span_depth: None,
                        span_style_name: None,
                    };
                    if empty {
                        master
                            .as_mut()
                            .ok_or_else(|| format_error(STYLES_PART, "master parser state is missing"))?
                            .templates
                            .push(current.template);
                    } else {
                        template = Some(current);
                    }
                } else if let Some(current) = template.as_mut() {
                    match local {
                        "image" if current.template.image.is_none() => {
                            if let Some(href) =
                                optional_attribute(&attributes, "href", STYLES_PART)?
                            {
                                let target = resolve_odf_image_target(
                                    &href,
                                    package.limits().max_zip_path_bytes,
                                )
                                .map_err(|message| format_error(STYLES_PART, message))?;
                                if let OdfImageTarget::Embedded(target) = target
                                    && let Some(bytes) = package.part(&target)?
                                {
                                    let declared_media_type = optional_attribute(
                                        &attributes,
                                        "mime-type",
                                        STYLES_PART,
                                    )?
                                    .or_else(|| manifest_media_types.get(&target).cloned());
                                    let media_type = identify_odf_image(
                                        &target,
                                        declared_media_type.as_deref(),
                                        &bytes,
                                    );
                                    if let Ok(media_type) = media_type {
                                        current.template.image = Some(OdfImageData {
                                            media_type: media_type.to_owned(),
                                            bytes: bytes.into_vec(),
                                            target,
                                        });
                                    }
                                }
                            }
                        }
                        "enhanced-geometry" => {
                            current.template.text_path = odf_text_path_mode(&attributes, STYLES_PART)?;
                            if let Some(geometry) = odf_enhanced_geometry_kind(
                                &attributes,
                                current.template.bounds,
                                STYLES_PART,
                            )? {
                                current.template.geometry = geometry;
                            }
                            if let Some(state) =
                                begin_odf_enhanced_geometry(depth, &attributes, STYLES_PART)?
                            {
                                if empty {
                                    if let Some(geometry) = parse_odf_enhanced_path_geometry(
                                        &state,
                                        current.template.bounds,
                                        FillRule::NonZero,
                                        &mut current.template.transform,
                                    ) {
                                        current.template.geometry = geometry;
                                    }
                                } else if current.enhanced_geometry.replace(state).is_some() {
                                    return Err(format_error(
                                        STYLES_PART,
                                        "nested ODP master enhanced geometry is invalid",
                                    ));
                                }
                            }
                        }
                        "equation" => {
                            if let Some(geometry) = current.enhanced_geometry.as_mut()
                                && let (Some(name), Some(formula)) = (
                                    optional_attribute(&attributes, "name", STYLES_PART)?,
                                    optional_attribute(&attributes, "formula", STYLES_PART)?,
                                )
                            {
                                geometry.equations.push((name, formula));
                            }
                        }
                        "p" => {
                            if current.paragraph_count != 0 {
                                append_odf_paragraph_break(
                                    &mut current.template.text,
                                    &mut current.template.text_runs,
                                );
                            }
                            current.paragraph_count += 1;
                            current.template.paragraph_style_names.push(optional_attribute(
                                &attributes,
                                "style-name",
                                STYLES_PART,
                            )?);
                            current.paragraph_depth = (!empty).then_some(depth);
                        }
                        "span" => {
                            if current.span_depth.is_some() {
                                return Err(format_error(
                                    STYLES_PART,
                                    "nested ODP master text spans are invalid",
                                ));
                            }
                            if !empty {
                                current.span_depth = Some(depth);
                                current.span_style_name = optional_attribute(
                                    &attributes,
                                    "style-name",
                                    STYLES_PART,
                                )?;
                            }
                        }
                        "s" | "tab" | "line-break"
                            if current.paragraph_depth.is_some() =>
                        {
                            let text = odf_control_text(
                                local,
                                &attributes,
                                package.limits().max_xml_bytes,
                                STYLES_PART,
                            )?;
                            let style_name = current_odf_text_style(
                                current.span_depth,
                                &current.span_style_name,
                                &current.template.paragraph_style_names,
                            );
                            append_odf_text(
                                &text,
                                &mut current.template.text,
                                &mut current.template.text_runs,
                                style_name.as_deref(),
                                None,
                            );
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
                if let Some(start) = ignored_template_depth {
                    if start == depth {
                        ignored_template_depth = None;
                    }
                    return Ok(());
                }
                let local = local_name(name);
                if local == "notes" && notes_depth == Some(depth) {
                    notes_depth = None;
                }
                if local == "enhanced-geometry"
                    && let Some(current) = template.as_mut()
                    && current
                        .enhanced_geometry
                        .as_ref()
                        .is_some_and(|geometry| geometry.depth == depth)
                {
                    let geometry = current.enhanced_geometry.take().ok_or_else(|| {
                        format_error(
                            STYLES_PART,
                            "master enhanced geometry parser state ended unexpectedly",
                        )
                    })?;
                    if let Some(geometry) = parse_odf_enhanced_path_geometry(
                        &geometry,
                        current.template.bounds,
                        FillRule::NonZero,
                        &mut current.template.transform,
                    ) {
                        current.template.geometry = geometry;
                    }
                }
                if local == "p"
                    && let Some(current) = template.as_mut()
                    && current.paragraph_depth == Some(depth)
                {
                    current.paragraph_depth = None;
                }
                if local == "span"
                    && let Some(current) = template.as_mut()
                    && current.span_depth == Some(depth)
                {
                    current.span_depth = None;
                    current.span_style_name = None;
                }
                if matches!(
                    local,
                    "frame" | "rect" | "ellipse" | "line" | "custom-shape"
                )
                    && template
                        .as_ref()
                        .is_some_and(|template| template.depth == depth)
                {
                    let current = template.take().ok_or_else(|| {
                        format_error(STYLES_PART, "master template parser state ended unexpectedly")
                    })?;
                    master
                        .as_mut()
                        .ok_or_else(|| format_error(STYLES_PART, "master parser state is missing"))?
                        .templates
                        .push(current.template);
                }
                if local == "master-page"
                    && master.as_ref().is_some_and(|master| master.depth == depth)
                {
                    let current = master.take().ok_or_else(|| {
                        format_error(STYLES_PART, "master parser state ended unexpectedly")
                    })?;
                    masters.insert(current.name, current.templates);
                }
            }
            XmlEvent::Text(text) => {
                if let Some(current) = template
                    .as_mut()
                    .filter(|template| template.paragraph_depth.is_some())
                {
                    let Some(text) = decode_odf_text_node(text, STYLES_PART)? else {
                        return Ok(());
                    };
                    let style_name = current_odf_text_style(
                        current.span_depth,
                        &current.span_style_name,
                        &current.template.paragraph_style_names,
                    );
                    append_odf_text(
                        &text,
                        &mut current.template.text,
                        &mut current.template.text_runs,
                        style_name.as_deref(),
                        None,
                    );
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(current) = template
                    .as_mut()
                    .filter(|template| template.paragraph_depth.is_some())
                {
                    let style_name = current_odf_text_style(
                        current.span_depth,
                        &current.span_style_name,
                        &current.template.paragraph_style_names,
                    );
                    append_odf_text(
                        text,
                        &mut current.template.text,
                        &mut current.template.text_runs,
                        style_name.as_deref(),
                        None,
                    );
                }
            }
        }
        Ok(())
    })
    .map(|_| (masters, invalid_master_shapes))
    .map_err(|error| with_part(error, STYLES_PART))
}

fn parse_style_part(
    bytes: &[u8],
    package: &dyn OdpSource,
    part: &str,
    catalog: &mut StyleCatalog,
    definitions: &mut HashMap<String, StyleDefinition>,
    defaults: &mut HashMap<String, GraphicStylePatch>,
    manifest_media_types: &HashMap<String, String>,
) -> Result<(), Diagnostic> {
    let mut depth = 0_usize;
    let mut current: Option<PendingStyle> = None;
    let mut fill_image = None;
    let mut binary_data: Option<String> = None;
    let mut binary_data_depth = None;
    let mut inline_bytes = 0;
    parse_xml(bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                match local_name(name) {
                    "master-page" => {
                        if let Some(style) = optional_attribute(&attributes, "style-name", part)? {
                            catalog
                                .master_backgrounds
                                .insert(required_attribute(&attributes, "name", part)?, style);
                        }
                    }
                    "font-face" => {
                        if let Some(name) = optional_attribute(&attributes, "name", part)? {
                            let family = optional_attribute(&attributes, "font-family", part)?
                                .unwrap_or_else(|| name.clone());
                            if let Some(family) = normalize_odf_font_family(&family) {
                                catalog.font_faces.insert(name, family);
                            }
                        }
                    }
                    "style" | "default-style" if current.is_none() => {
                        let family =
                            optional_attribute(&attributes, "family", part)?.unwrap_or_default();
                        if matches!(
                            family.as_str(),
                            "graphic"
                                | "presentation"
                                | "table-cell"
                                | "drawing-page"
                                | "paragraph"
                                | "text"
                        ) {
                            current = Some(PendingStyle {
                                depth,
                                name: (local_name(name) == "style")
                                    .then(|| required_attribute(&attributes, "name", part))
                                    .transpose()?,
                                family,
                                parent: optional_attribute(&attributes, "parent-style-name", part)?,
                                patch: GraphicStylePatch::default(),
                            });
                        }
                    }
                    "graphic-properties" | "drawing-page-properties" => {
                        if let Some(style) = current.as_mut() {
                            apply_graphic_properties_patch(&mut style.patch, &attributes, part)?;
                        }
                    }
                    "columns"
                        if current.as_ref().is_some_and(|style| {
                            matches!(style.family.as_str(), "graphic" | "presentation")
                        }) =>
                    {
                        let count = optional_attribute(&attributes, "column-count", part)?
                            .and_then(|value| value.parse::<u32>().ok())
                            .unwrap_or(1)
                            .clamp(1, 64);
                        let gap = optional_length(&attributes, "column-gap", part)?
                            .unwrap_or(0.0)
                            .max(0.0);
                        // Drawing frames flow through equal-width columns. ODT sections
                        // have separate balancing and unequal-column semantics.
                        current.as_mut().unwrap().patch.columns = Some((count, gap));
                    }
                    "table-cell-properties" => {
                        if let Some(style) = current.as_mut() {
                            if let Some(color) =
                                optional_attribute(&attributes, "background-color", part)?
                                    .and_then(|color| parse_odf_color(&color))
                            {
                                style.patch.fill = Some(OdfPaint::Solid(color));
                            }
                            for name in [
                                "border",
                                "border-left",
                                "border-right",
                                "border-top",
                                "border-bottom",
                            ] {
                                let Some(border) = optional_attribute(&attributes, name, part)?
                                else {
                                    continue;
                                };
                                if border.eq_ignore_ascii_case("none") {
                                    style.patch.stroke = Some(OdfPaint::None);
                                    style.patch.stroke_width = Some(0.0);
                                    style.patch.border_style = Some(None);
                                } else if let Some((width, color, stroke_style)) =
                                    parse_odf_border(&border)
                                {
                                    style.patch.stroke = Some(OdfPaint::Solid(color));
                                    style.patch.stroke_width = Some(width);
                                    style.patch.border_style = Some(Some(stroke_style));
                                }
                                break;
                            }
                            if let Some(vertical_align) =
                                optional_attribute(&attributes, "vertical-align", part)?
                            {
                                style.patch.vertical_align = Some(match vertical_align.as_str() {
                                    "middle" | "center" => TextVerticalAlign::Center,
                                    "bottom" => TextVerticalAlign::Bottom,
                                    _ => TextVerticalAlign::Top,
                                });
                            }
                            if let Some(padding) = optional_length(&attributes, "padding", part)? {
                                style.patch.inset_left = Some(padding);
                                style.patch.inset_right = Some(padding);
                                style.patch.inset_top = Some(padding);
                                style.patch.inset_bottom = Some(padding);
                            }
                            style.patch.inset_left =
                                optional_length(&attributes, "padding-left", part)?
                                    .or(style.patch.inset_left);
                            style.patch.inset_right =
                                optional_length(&attributes, "padding-right", part)?
                                    .or(style.patch.inset_right);
                            style.patch.inset_top =
                                optional_length(&attributes, "padding-top", part)?
                                    .or(style.patch.inset_top);
                            style.patch.inset_bottom =
                                optional_length(&attributes, "padding-bottom", part)?
                                    .or(style.patch.inset_bottom);
                        }
                    }
                    "text-properties" => {
                        if let Some(style) = current.as_mut() {
                            // Resolve declared face aliases before inheriting a parent's family.
                            // ODP retains its family-list normalization; ODT has different list handling.
                            let font_name = optional_attribute(&attributes, "font-name", part)?;
                            let family = font_name
                                .as_ref()
                                .and_then(|name| catalog.font_faces.get(name).cloned())
                                .or(font_name)
                                .or(optional_attribute(&attributes, "font-family", part)?)
                                .or(optional_attribute(&attributes, "font-family-asian", part)?)
                                .or(optional_attribute(
                                    &attributes,
                                    "font-family-complex",
                                    part,
                                )?);
                            if let Some(family) =
                                family.as_deref().and_then(normalize_odf_font_family)
                            {
                                style.patch.font_family = Some(family);
                            }
                            if let Some(size) = optional_attribute(&attributes, "font-size", part)?
                                && let Some(size) = parse_length(&size)
                            {
                                style.patch.font_size = Some(size);
                            }
                            if let Some(color) = optional_attribute(&attributes, "color", part)?
                                .and_then(|color| parse_odf_color(&color))
                            {
                                style.patch.color = Some(color);
                            }
                            if let Some(weight) =
                                optional_attribute(&attributes, "font-weight", part)?
                            {
                                style.patch.bold = Some(weight.eq_ignore_ascii_case("bold"));
                            }
                            if let Some(value) =
                                optional_attribute(&attributes, "font-style", part)?
                            {
                                style.patch.italic = Some(value.eq_ignore_ascii_case("italic"));
                            }
                        }
                    }
                    "paragraph-properties" => {
                        if let Some(style) = current.as_mut() {
                            if let Some(writing_mode) =
                                optional_attribute(&attributes, "writing-mode", part)?
                            {
                                style.patch.orientation = match writing_mode.as_str() {
                                    "tb-rl" | "vertical-rl" => Some(TextOrientation::VerticalRl),
                                    "tb-lr" | "vertical-lr" => Some(TextOrientation::VerticalLr),
                                    "sideways-rl" => Some(TextOrientation::Rotated90),
                                    "sideways-lr" => Some(TextOrientation::Rotated270),
                                    "lr" | "lr-tb" | "rl" | "rl-tb" => {
                                        Some(TextOrientation::Horizontal)
                                    }
                                    _ => None,
                                };
                            }
                            if let Some(align) =
                                optional_attribute(&attributes, "text-align", part)?
                            {
                                style.patch.align = Some(super::odf_text_align(&align));
                            }
                            style.patch.margin_left =
                                optional_length(&attributes, "margin-left", part)?;
                            style.patch.margin_right =
                                optional_length(&attributes, "margin-right", part)?;
                            style.patch.first_line_indent =
                                optional_length(&attributes, "text-indent", part)?;
                            style.patch.default_tab_stop =
                                optional_length(&attributes, "tab-stop-distance", part)?;
                            let margin_top = optional_length(&attributes, "margin-top", part)?;
                            let margin_bottom =
                                optional_length(&attributes, "margin-bottom", part)?;
                            if margin_top.is_some() || margin_bottom.is_some() {
                                style.patch.paragraph_spacing =
                                    Some(margin_top.unwrap_or(0.0) + margin_bottom.unwrap_or(0.0));
                            }
                            if let Some(value) =
                                optional_attribute(&attributes, "line-height", part)?
                            {
                                style.patch.line_height = parse_odf_line_height(&value);
                            }
                        }
                    }
                    "gradient" => {
                        let name = required_attribute(&attributes, "name", part)?;
                        let gradient_style = optional_attribute(&attributes, "style", part)?;
                        let radial = gradient_style.as_deref().is_some_and(|style| {
                            matches!(style, "radial" | "ellipsoid" | "square" | "rectangular")
                        });
                        let rectangular = gradient_style.as_deref() == Some("rectangular");
                        let start = optional_attribute(&attributes, "start-color", part)?
                            .and_then(|color| parse_odf_color(&color))
                            .unwrap_or(0xffff_ffff);
                        let end = optional_attribute(&attributes, "end-color", part)?
                            .and_then(|color| parse_odf_color(&color))
                            .unwrap_or(0x0000_00ff);
                        let angle_degrees = optional_attribute(&attributes, "angle", part)?
                            .and_then(|value| value.parse::<f32>().ok())
                            .map(|value| value / 10.0)
                            .unwrap_or(0.0);
                        let center_x = optional_attribute(&attributes, "cx", part)?
                            .and_then(|value| parse_percentage(&value))
                            .unwrap_or(0.5);
                        let center_y = optional_attribute(&attributes, "cy", part)?
                            .and_then(|value| parse_percentage(&value))
                            .unwrap_or(0.5);
                        catalog.gradients.insert(
                            name,
                            GradientDefinition {
                                start,
                                end,
                                angle_degrees,
                                radial,
                                rectangular,
                                center_x,
                                center_y,
                            },
                        );
                    }
                    "opacity" => {
                        let (name, gradient) = parse_odf_opacity_gradient(&attributes, part)?;
                        catalog.opacity_gradients.insert(name, gradient);
                    }
                    "stroke-dash" => {
                        let name = required_attribute(&attributes, "name", part)?;
                        let dots1 = optional_attribute(&attributes, "dots1", part)?
                            .and_then(|value| value.parse::<usize>().ok())
                            .unwrap_or(1)
                            .min(64);
                        let dots2 = optional_attribute(&attributes, "dots2", part)?
                            .and_then(|value| value.parse::<usize>().ok())
                            .unwrap_or(0)
                            .min(64);
                        let dots1_length = if dots1 == 0 {
                            None
                        } else {
                            optional_dash_length(&attributes, "dots1-length", part)?
                        };
                        let dots2_length = if dots2 == 0 {
                            None
                        } else {
                            optional_dash_length(&attributes, "dots2-length", part)?
                        };
                        let distance = optional_dash_length(&attributes, "distance", part)?;
                        let mut dash = Vec::with_capacity((dots1 + dots2) * 2);
                        if let (Some(dots1_length), Some(distance)) = (dots1_length, distance) {
                            for _ in 0..dots1 {
                                dash.extend([dots1_length, distance]);
                            }
                        }
                        if let (Some(dots2_length), Some(distance)) = (dots2_length, distance) {
                            for _ in 0..dots2 {
                                dash.extend([dots2_length, distance]);
                            }
                        }
                        if !dash.is_empty() {
                            let cap =
                                match optional_attribute(&attributes, "style", part)?.as_deref() {
                                    Some("round") => LineCap::Round,
                                    _ => LineCap::Flat,
                                };
                            catalog.stroke_dashes.insert(name, (dash, cap));
                        }
                    }
                    "binary-data" if fill_image.is_some() => {
                        binary_data = Some(String::new());
                        binary_data_depth = (!empty).then_some(depth);
                    }
                    "fill-image" => {
                        let image_name = required_attribute(&attributes, "name", part)?;
                        if !empty {
                            fill_image = Some(image_name.clone());
                        }
                        if let Some(href) = optional_attribute(&attributes, "href", part)? {
                            let target = resolve_odf_image_target(
                                &href,
                                package.limits().max_zip_path_bytes,
                            )
                            .map_err(|message| format_error(part, message))?;
                            if let OdfImageTarget::Embedded(target) = target
                                && let Some(bytes) = package.part(&target)?
                            {
                                let declared_media_type =
                                    optional_attribute(&attributes, "mime-type", part)?
                                        .or_else(|| manifest_media_types.get(&target).cloned());
                                let media_type = identify_odf_image(
                                    &target,
                                    declared_media_type.as_deref(),
                                    &bytes,
                                );
                                if let Ok(media_type) = media_type {
                                    catalog.fill_images.insert(
                                        image_name,
                                        OdfImageData {
                                            media_type: media_type.to_owned(),
                                            bytes: bytes.into_vec(),
                                            target,
                                        },
                                    );
                                }
                            }
                        }
                    }
                    _ => {}
                }
                if empty
                    && matches!(local_name(name), "style" | "default-style")
                    && let Some(style) = current.take()
                {
                    store_pending_style(style, part, definitions, defaults);
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if binary_data_depth == Some(depth) && local_name(name) == "binary-data" {
                    binary_data_depth = None;
                }
                if local_name(name) == "fill-image" {
                    if let (Some(image_name), Some(encoded)) =
                        (fill_image.take(), binary_data.take())
                    {
                        match decode_inline_base64(&encoded) {
                            Ok(bytes) => {
                                reserve_materialized_image_bytes(
                                    &mut inline_bytes,
                                    bytes.len(),
                                    package.limits().max_total_uncompressed_bytes,
                                    part,
                                )?;
                                match identify_odf_image(part, None, &bytes) {
                                    Ok(media_type) => {
                                        catalog.fill_images.insert(
                                            image_name,
                                            OdfImageData {
                                                media_type: media_type.to_owned(),
                                                bytes,
                                                target: part.to_owned(),
                                            },
                                        );
                                    }
                                    Err(error) => catalog
                                        .diagnostics
                                        .push(unsupported_image_diagnostic(part, error)),
                                }
                            }
                            Err(error) if error.code == DiagnosticCode::AllocationFailed => {
                                return Err(error);
                            }
                            Err(_) => catalog.diagnostics.push(
                                Diagnostic::warning(
                                    DiagnosticCode::UnsupportedFeature,
                                    Phase::Parse,
                                    Fidelity::Omitted,
                                    "invalid inline fill image was omitted",
                                )
                                .in_part(part),
                            ),
                        }
                    }
                }
                if matches!(local_name(name), "style" | "default-style")
                    && current.as_ref().is_some_and(|style| style.depth == depth)
                {
                    let style = current.take().ok_or_else(|| {
                        format_error(part, "style parser state ended unexpectedly")
                    })?;
                    store_pending_style(style, part, definitions, defaults);
                }
            }
            XmlEvent::Text(text) if binary_data_depth.is_some() => {
                append_inline_base64(binary_data.as_mut().unwrap(), &decode_xml_text(text)?)?;
            }
            XmlEvent::Cdata(text) if binary_data_depth.is_some() => {
                append_inline_base64(binary_data.as_mut().unwrap(), text)?;
            }
            _ => {}
        }
        Ok(())
    })
    .map(|_| ())
    .map_err(|error| with_part(error, part))
}

fn parse_list_styles(
    package: &dyn OdpSource,
) -> Result<HashMap<String, HashMap<u32, OdfListLevelStyle>>, Diagnostic> {
    #[derive(Debug)]
    struct PendingListStyle {
        depth: usize,
        name: String,
        levels: HashMap<u32, OdfListLevelStyle>,
    }

    #[derive(Debug)]
    struct PendingListLevel {
        depth: usize,
        level: u32,
        style: OdfListLevelStyle,
    }

    let mut styles = HashMap::new();
    for &part in style_parts(package) {
        let bytes = package.required_part(part)?;
        let mut depth = 0_usize;
        let mut current_style: Option<PendingListStyle> = None;
        let mut current_level: Option<PendingListLevel> = None;
        parse_xml(&bytes, package.limits(), |event| {
            match event {
                XmlEvent::StartElement {
                    name,
                    attributes,
                    empty,
                } => {
                    let local = local_name(name);
                    match local {
                        "list-style" if current_style.is_none() => {
                            let name = required_attribute(&attributes, "name", part)?;
                            if empty {
                                styles.insert(name, HashMap::new());
                            } else {
                                current_style = Some(PendingListStyle {
                                    depth,
                                    name,
                                    levels: HashMap::new(),
                                });
                            }
                        }
                        "list-level-style-bullet"
                        | "list-level-style-number"
                        | "list-level-style-image"
                            if current_style.is_some() && current_level.is_none() =>
                        {
                            let level = required_attribute(&attributes, "level", part)?
                                .parse::<u32>()
                                .ok()
                                .filter(|level| *level != 0)
                                .ok_or_else(|| {
                                    format_error(
                                        part,
                                        "ODF list level is not a supported positive integer",
                                    )
                                })?;
                            let label = if local == "list-level-style-bullet" {
                                OdfListLabel::Bullet(
                                    optional_attribute(&attributes, "bullet-char", part)?
                                        .unwrap_or_default(),
                                )
                            } else if local == "list-level-style-image" {
                                OdfListLabel::Image {
                                    href: optional_attribute(&attributes, "href", part)?
                                        .unwrap_or_default(),
                                    width: 0.0,
                                    height: 0.0,
                                }
                            } else {
                                OdfListLabel::Number {
                                    format: optional_attribute(&attributes, "num-format", part)?
                                        .unwrap_or_else(|| "1".to_owned()),
                                    prefix: optional_attribute(&attributes, "num-prefix", part)?
                                        .unwrap_or_default(),
                                    suffix: optional_attribute(&attributes, "num-suffix", part)?
                                        .unwrap_or_default(),
                                    start_value: optional_attribute(
                                        &attributes,
                                        "start-value",
                                        part,
                                    )?
                                    .map(|value| {
                                        value.parse::<u32>().map_err(|_| {
                                            format_error(
                                                part,
                                                "ODF list start value exceeds the supported range",
                                            )
                                        })
                                    })
                                    .transpose()?,
                                }
                            };
                            let pending = PendingListLevel {
                                depth,
                                level,
                                style: OdfListLevelStyle {
                                    label,
                                    space_before: 0.0,
                                    min_label_width: 0.0,
                                    font_family: None,
                                    label_style: OdfListLabelStyle::default(),
                                },
                            };
                            if empty {
                                current_style
                                    .as_mut()
                                    .ok_or_else(|| {
                                        format_error(part, "ODF list style parser state is missing")
                                    })?
                                    .levels
                                    .insert(level, pending.style);
                            } else {
                                current_level = Some(pending);
                            }
                        }
                        "list-level-properties" => {
                            if let Some(level) = current_level.as_mut() {
                                if let OdfListLabel::Image { width, height, .. } =
                                    &mut level.style.label
                                {
                                    *width =
                                        optional_length(&attributes, "width", part)?.unwrap_or(0.0);
                                    *height = optional_length(&attributes, "height", part)?
                                        .unwrap_or(0.0);
                                }
                                level.style.space_before =
                                    optional_length(&attributes, "space-before", part)?
                                        .unwrap_or(0.0);
                                level.style.min_label_width =
                                    optional_length(&attributes, "min-label-width", part)?
                                        .unwrap_or(0.0);
                            }
                        }
                        "text-properties" => {
                            if let Some(level) = current_level.as_mut() {
                                level.style.font_family =
                                    optional_attribute(&attributes, "font-family", part)?
                                        .or(optional_attribute(
                                            &attributes,
                                            "font-family-asian",
                                            part,
                                        )?)
                                        .or(optional_attribute(
                                            &attributes,
                                            "font-family-complex",
                                            part,
                                        )?)
                                        .as_deref()
                                        .and_then(normalize_odf_font_family);
                                level.style.label_style.color =
                                    optional_attribute(&attributes, "color", part)?
                                        .as_deref()
                                        .and_then(parse_odf_color);
                                level.style.label_style.font_size =
                                    optional_attribute(&attributes, "font-size", part)?
                                        .as_deref()
                                        .and_then(parse_odf_list_font_size);
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
                    if matches!(
                        local,
                        "list-level-style-bullet"
                            | "list-level-style-number"
                            | "list-level-style-image"
                    ) && current_level
                        .as_ref()
                        .is_some_and(|level| level.depth == depth)
                    {
                        let level = current_level.take().ok_or_else(|| {
                            format_error(part, "ODF list level parser state ended unexpectedly")
                        })?;
                        current_style
                            .as_mut()
                            .ok_or_else(|| {
                                format_error(part, "ODF list style parser state is missing")
                            })?
                            .levels
                            .insert(level.level, level.style);
                    } else if local == "list-style"
                        && current_style
                            .as_ref()
                            .is_some_and(|style| style.depth == depth)
                    {
                        let style = current_style.take().ok_or_else(|| {
                            format_error(part, "ODF list style parser state ended unexpectedly")
                        })?;
                        styles.insert(style.name, style.levels);
                    }
                }
                _ => {}
            }
            Ok(())
        })
        .map_err(|error| with_part(error, part))?;
    }
    Ok(styles)
}

type TableStyleDimensions = (HashMap<String, f32>, HashMap<String, f32>);

fn parse_table_style_dimensions(
    package: &dyn OdpSource,
) -> Result<TableStyleDimensions, Diagnostic> {
    #[derive(Debug)]
    struct PendingTableStyle {
        depth: usize,
        name: String,
        family: String,
    }

    let mut column_widths = HashMap::new();
    let mut row_heights = HashMap::new();
    for &part in style_parts(package) {
        let bytes = package.required_part(part)?;
        let mut depth = 0_usize;
        let mut current: Option<PendingTableStyle> = None;
        parse_xml(&bytes, package.limits(), |event| {
            match event {
                XmlEvent::StartElement {
                    name,
                    attributes,
                    empty,
                } => {
                    let local = local_name(name);
                    match local {
                        "style" if current.is_none() => {
                            let family = optional_attribute(&attributes, "family", part)?
                                .unwrap_or_default();
                            if matches!(family.as_str(), "table-column" | "table-row") && !empty {
                                current = Some(PendingTableStyle {
                                    depth,
                                    name: required_attribute(&attributes, "name", part)?,
                                    family,
                                });
                            }
                        }
                        "table-column-properties"
                            if current
                                .as_ref()
                                .is_some_and(|style| style.family == "table-column") =>
                        {
                            if let Some(width) = optional_length(&attributes, "column-width", part)?
                            {
                                let name = current
                                    .as_ref()
                                    .map(|style| style.name.clone())
                                    .ok_or_else(|| {
                                        format_error(
                                            part,
                                            "ODF table-column style parser state is missing",
                                        )
                                    })?;
                                column_widths.insert(name, width);
                            }
                        }
                        "table-row-properties"
                            if current
                                .as_ref()
                                .is_some_and(|style| style.family == "table-row") =>
                        {
                            if let Some(height) = optional_length(&attributes, "row-height", part)?
                            {
                                let name = current
                                    .as_ref()
                                    .map(|style| style.name.clone())
                                    .ok_or_else(|| {
                                        format_error(
                                            part,
                                            "ODF table-row style parser state is missing",
                                        )
                                    })?;
                                row_heights.insert(name, height);
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
                    if local_name(name) == "style"
                        && current.as_ref().is_some_and(|style| style.depth == depth)
                    {
                        current = None;
                    }
                }
                _ => {}
            }
            Ok(())
        })
        .map_err(|error| with_part(error, part))?;
    }
    Ok((column_widths, row_heights))
}

fn parse_table_templates(
    package: &dyn OdpSource,
) -> Result<HashMap<String, OdfTableTemplate>, Diagnostic> {
    #[derive(Debug)]
    struct PendingTableTemplate {
        depth: usize,
        name: String,
        template: OdfTableTemplate,
    }

    let mut templates = HashMap::new();
    for &part in style_parts(package) {
        let bytes = package.required_part(part)?;
        let mut depth = 0_usize;
        let mut current: Option<PendingTableTemplate> = None;
        parse_xml(&bytes, package.limits(), |event| {
            match event {
                XmlEvent::StartElement {
                    name,
                    attributes,
                    empty,
                } => {
                    let local = local_name(name);
                    if local == "table-template" && current.is_none() {
                        let pending = PendingTableTemplate {
                            depth,
                            name: optional_attribute(&attributes, "name", part)?
                                .or(optional_attribute(&attributes, "style-name", part)?)
                                .ok_or_else(|| {
                                    format_error(part, "ODF table template has no name")
                                })?,
                            template: OdfTableTemplate::default(),
                        };
                        if empty {
                            templates.insert(pending.name, pending.template);
                        } else {
                            current = Some(pending);
                        }
                    } else if let Some(pending) = current.as_mut() {
                        let style_name = optional_attribute(&attributes, "style-name", part)?;
                        match local {
                            "body" => pending.template.body = style_name,
                            "first-row" => pending.template.first_row = style_name,
                            "last-row" => pending.template.last_row = style_name,
                            "first-column" => pending.template.first_column = style_name,
                            "last-column" => pending.template.last_column = style_name,
                            "even-rows" => pending.template.even_rows = style_name,
                            "odd-rows" => pending.template.odd_rows = style_name,
                            "even-columns" => pending.template.even_columns = style_name,
                            "odd-columns" => pending.template.odd_columns = style_name,
                            _ => {}
                        }
                    }
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                }
                XmlEvent::EndElement { name } => {
                    depth = depth.saturating_sub(1);
                    if local_name(name) == "table-template"
                        && current
                            .as_ref()
                            .is_some_and(|template| template.depth == depth)
                    {
                        let pending = current.take().ok_or_else(|| {
                            format_error(part, "ODF table template parser state ended unexpectedly")
                        })?;
                        templates.insert(pending.name, pending.template);
                    }
                }
                _ => {}
            }
            Ok(())
        })
        .map_err(|error| with_part(error, part))?;
    }
    Ok(templates)
}

fn store_pending_style(
    style: PendingStyle,
    part: &str,
    definitions: &mut HashMap<String, StyleDefinition>,
    defaults: &mut HashMap<String, GraphicStylePatch>,
) {
    if let Some(name) = style.name {
        definitions.insert(
            name,
            StyleDefinition {
                family: style.family,
                parent: style.parent,
                patch: style.patch,
                part: part.to_owned(),
            },
        );
    } else {
        defaults.entry(style.family).or_default().merge(style.patch);
    }
}

fn normalize_odf_font_family(value: &str) -> Option<String> {
    let value = super::odf_primary_font_family(value);
    match value.to_ascii_lowercase().as_str() {
        "+mn-lt" | "+mn-cs" | "+mj-lt" | "+mj-cs" => Some("Aptos".to_owned()),
        "+mn-ea" | "+mj-ea" => Some("sans-serif".to_owned()),
        _ => (!value.is_empty() && !value.starts_with('+')).then(|| value.to_owned()),
    }
}

fn parse_odf_line_height(value: &str) -> Option<OdfLineHeight> {
    let value = value.trim();
    if value.eq_ignore_ascii_case("normal") {
        return None;
    }
    if let Some(percent) = value.strip_suffix('%') {
        return percent
            .parse::<f32>()
            .ok()
            .map(|percent| percent / 100.0)
            .filter(|multiple| multiple.is_finite() && *multiple > 0.0)
            .map(OdfLineHeight::Multiple);
    }
    parse_length(value)
        .filter(|height| *height > 0.0)
        .map(OdfLineHeight::Exact)
}

fn parse_odf_list_font_size(value: &str) -> Option<OdfListFontSize> {
    let value = value.trim();
    if let Some(percent) = value.strip_suffix('%') {
        return percent
            .parse::<f32>()
            .ok()
            .map(|percent| percent / 100.0)
            .filter(|relative| relative.is_finite() && *relative > 0.0)
            .map(OdfListFontSize::Relative);
    }
    parse_length(value)
        .filter(|size| *size > 0.0)
        .map(OdfListFontSize::Exact)
}

fn apply_graphic_properties_patch(
    style: &mut GraphicStylePatch,
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<(), Diagnostic> {
    let fill_kind = optional_attribute(attributes, "fill", part)?;
    let fill_color = optional_attribute(attributes, "fill-color", part)?
        .and_then(|color| parse_odf_color(&color));
    let gradient = optional_attribute(attributes, "fill-gradient-name", part)?;
    let bitmap = optional_attribute(attributes, "fill-image-name", part)?;
    if fill_kind.is_some() {
        style.fill = Some(match fill_kind.as_deref() {
            Some("none") => OdfPaint::None,
            Some("gradient") => gradient.map_or(OdfPaint::None, OdfPaint::Gradient),
            Some("bitmap") => bitmap.map_or(OdfPaint::None, OdfPaint::Bitmap),
            Some("solid") => fill_color.map_or(OdfPaint::None, OdfPaint::Solid),
            _ => fill_color.map_or(OdfPaint::None, OdfPaint::Solid),
        });
    }
    style.fill_color = fill_color;
    if let Some(repeat) = optional_attribute(attributes, "repeat", part)? {
        style.fill_image_tile = Some(repeat != "stretch" && repeat != "no-repeat");
    }
    if let Some(width) = optional_attribute(attributes, "fill-image-width", part)?
        && let Some(width) = parse_length(&width).filter(|width| *width > 0.0)
    {
        style.fill_image_width = Some(width);
    }
    if let Some(height) = optional_attribute(attributes, "fill-image-height", part)?
        && let Some(height) = parse_length(&height).filter(|height| *height > 0.0)
    {
        style.fill_image_height = Some(height);
    }
    let stroke_kind = optional_attribute(attributes, "stroke", part)?;
    let stroke_color = optional_attribute(attributes, "stroke-color", part)?
        .and_then(|color| parse_odf_color(&color));
    if stroke_kind.is_some() || stroke_color.is_some() {
        style.stroke = Some(match stroke_kind.as_deref() {
            Some("none") => OdfPaint::None,
            Some(_) => stroke_color.map_or(OdfPaint::None, OdfPaint::Solid),
            None => stroke_color.map_or(OdfPaint::None, OdfPaint::Solid),
        });
    }
    if let Some(stroke_kind) = stroke_kind.as_deref() {
        style.stroke_dash = Some(if stroke_kind == "dash" {
            optional_attribute(attributes, "stroke-dash", part)?
        } else {
            None
        });
    }
    if let Some(width) = optional_attribute(attributes, "stroke-width", part)?
        && let Some(width) = parse_length(&width)
    {
        style.stroke_width = Some(width);
    }
    if let Some(cap) = optional_attribute(attributes, "stroke-linecap", part)? {
        style.line_cap = match cap.as_str() {
            "butt" => Some(LineCap::Flat),
            "round" => Some(LineCap::Round),
            "square" => Some(LineCap::Square),
            _ => None,
        };
    }
    if let Some(opacity) = optional_attribute(attributes, "opacity", part)? {
        style.opacity = parse_percentage(&opacity);
    }
    if let Some(vertical_align) = optional_attribute(attributes, "textarea-vertical-align", part)? {
        style.vertical_align = Some(match vertical_align.as_str() {
            "middle" | "center" => TextVerticalAlign::Center,
            "bottom" => TextVerticalAlign::Bottom,
            _ => TextVerticalAlign::Top,
        });
    }
    if let Some(horizontal_align) =
        optional_attribute(attributes, "textarea-horizontal-align", part)?
    {
        style.align = Some(match horizontal_align.as_str() {
            "center" => TextAlign::Center,
            "end" | "right" => TextAlign::End,
            _ => TextAlign::Start,
        });
    }
    if let Some(wrap) = optional_attribute(attributes, "wrap-option", part)? {
        style.wrap = Some(wrap != "no-wrap");
    }
    if let Some(clip) = optional_attribute(attributes, "clip", part)? {
        style.image_clip = parse_odf_image_clip(&clip);
    }
    if let Some(shrink) = optional_attribute(attributes, "shrink-to-fit", part)? {
        style.auto_fit = Some(if matches!(shrink.as_str(), "true" | "1") {
            TextAutoFit::Shrink
        } else {
            TextAutoFit::None
        });
    }
    if let Some(fit) = optional_attribute(attributes, "fit-to-size", part)? {
        style.auto_fit = match fit.as_str() {
            "true" | "all" => Some(TextAutoFit::FitFrame),
            "shrink-to-fit" => Some(TextAutoFit::Shrink),
            _ => style.auto_fit,
        };
    }
    if let Some(padding) = optional_attribute(attributes, "padding", part)?
        .as_deref()
        .and_then(parse_length)
    {
        style.inset_left = Some(padding);
        style.inset_right = Some(padding);
        style.inset_top = Some(padding);
        style.inset_bottom = Some(padding);
    }
    if let Some(padding) = optional_attribute(attributes, "padding-left", part)?
        .as_deref()
        .and_then(parse_length)
    {
        style.inset_left = Some(padding);
    }
    if let Some(padding) = optional_attribute(attributes, "padding-right", part)?
        .as_deref()
        .and_then(parse_length)
    {
        style.inset_right = Some(padding);
    }
    if let Some(padding) = optional_attribute(attributes, "padding-top", part)?
        .as_deref()
        .and_then(parse_length)
    {
        style.inset_top = Some(padding);
    }
    if let Some(padding) = optional_attribute(attributes, "padding-bottom", part)?
        .as_deref()
        .and_then(parse_length)
    {
        style.inset_bottom = Some(padding);
    }
    if let Some(marker) = optional_attribute(attributes, "marker-start", part)? {
        style.marker_start = Some(!marker.is_empty() && marker != "none");
    }
    if let Some(marker) = optional_attribute(attributes, "marker-end", part)? {
        style.marker_end = Some(!marker.is_empty() && marker != "none");
    }
    if let Some(shadow) = optional_attribute(attributes, "shadow", part)? {
        if matches!(shadow.as_str(), "visible" | "true" | "1") {
            let opacity = optional_attribute(attributes, "shadow-opacity", part)?
                .as_deref()
                .and_then(parse_percentage)
                .unwrap_or(1.0);
            let mut color = optional_attribute(attributes, "shadow-color", part)?
                .and_then(|color| parse_odf_color(&color))
                .unwrap_or(0x0000_00ff);
            color = (color & 0xffff_ff00) | (opacity * 255.0).round() as u32;
            let offset_x = optional_attribute(attributes, "shadow-offset-x", part)?
                .as_deref()
                .and_then(parse_length)
                .unwrap_or(0.0);
            let offset_y = optional_attribute(attributes, "shadow-offset-y", part)?
                .as_deref()
                .and_then(parse_length)
                .unwrap_or(0.0);
            let blur = optional_attribute(attributes, "shadow-blur", part)?
                .as_deref()
                .and_then(parse_length)
                .unwrap_or(0.0);
            style.shadow = Some(Some(Shadow {
                color,
                blur,
                offset_x,
                offset_y,
            }));
        } else {
            style.shadow = Some(None);
        }
    }
    Ok(())
}

fn parse_odf_image_clip(value: &str) -> Option<OdfImageClip> {
    let value = value.trim();
    let inner = value
        .strip_prefix("rect(")
        .and_then(|value| value.strip_suffix(')'))?;
    let values = inner
        .split(|character: char| character == ',' || character.is_ascii_whitespace())
        .filter(|value| !value.is_empty())
        .collect::<Vec<_>>();
    let [top, right, bottom, left] = values.as_slice() else {
        return None;
    };
    let length = |value: &str| {
        if value.eq_ignore_ascii_case("auto") {
            Some(0.0)
        } else {
            parse_length(value)
        }
    };
    Some(OdfImageClip {
        left: length(left)?,
        top: length(top)?,
        right: length(right)?,
        bottom: length(bottom)?,
    })
}

fn normalized_odf_image_crop(clip: Option<OdfImageClip>, bounds: Rect) -> ImageCrop {
    let Some(clip) = clip else {
        return ImageCrop::default();
    };
    let right_is_extent = clip.right > bounds.width;
    let bottom_is_extent = clip.bottom > bounds.height;
    let source_width = if right_is_extent {
        clip.right
    } else {
        bounds.width + clip.left + clip.right
    };
    let source_height = if bottom_is_extent {
        clip.bottom
    } else {
        bounds.height + clip.top + clip.bottom
    };
    if source_width <= f32::EPSILON || source_height <= f32::EPSILON {
        return ImageCrop::default();
    }
    let crop = ImageCrop {
        left: clip.left / source_width,
        top: clip.top / source_height,
        right: if right_is_extent {
            (source_width - bounds.width - clip.left).max(0.0) / source_width
        } else {
            clip.right / source_width
        },
        bottom: if bottom_is_extent {
            (source_height - bounds.height - clip.top).max(0.0) / source_height
        } else {
            clip.bottom / source_height
        },
    };
    if crop.is_valid() {
        crop
    } else {
        ImageCrop::default()
    }
}

fn parse_odf_color(value: &str) -> Option<u32> {
    let value = value.trim();
    let hex = value.strip_prefix('#')?;
    match hex.len() {
        6 => u32::from_str_radix(hex, 16)
            .ok()
            .map(|rgb| (rgb << 8) | 0xff),
        3 => {
            let mut expanded = String::with_capacity(6);
            for character in hex.chars() {
                expanded.push(character);
                expanded.push(character);
            }
            u32::from_str_radix(&expanded, 16)
                .ok()
                .map(|rgb| (rgb << 8) | 0xff)
        }
        _ => None,
    }
}

fn parse_percentage(value: &str) -> Option<f32> {
    value
        .trim()
        .strip_suffix('%')
        .and_then(|number| number.parse::<f32>().ok())
        .map(|number| (number / 100.0).clamp(0.0, 1.0))
}

fn parse_odf_border(value: &str) -> Option<(f32, u32, StrokeStyle)> {
    let mut width = None;
    let mut color = None;
    let mut pattern = "solid";
    for token in value.split_whitespace() {
        width = width.or_else(|| parse_length(token));
        color = color.or_else(|| parse_odf_color(token));
        if matches!(
            token,
            "solid"
                | "dotted"
                | "dashed"
                | "fine-dashed"
                | "dash-dot"
                | "dash-dot-dot"
                | "double"
                | "double-thin"
        ) {
            pattern = token;
        }
    }
    let width = width?.max(0.5);
    Some((
        width,
        color?,
        super::odf_border_stroke_style(pattern, width),
    ))
}

#[derive(Debug)]
struct PageState {
    depth: usize,
    unit_index: u32,
    frame_index: u32,
}

#[derive(Debug, Default)]
struct OdfListTextState {
    stack: Vec<ActiveOdfList>,
    numbering_counters: Vec<u32>,
    numbering_formats: Vec<String>,
    pending_prefix: Option<(String, OdfListLabelStyle)>,
    pending_counter_restore: Option<(usize, u32)>,
}

#[derive(Debug)]
struct ActiveOdfList {
    depth: usize,
    style_name: Option<String>,
    reset_numbering: bool,
    continue_numbering: bool,
    item_depth: Option<usize>,
    item_count: u32,
    item_start_value: Option<u32>,
    item_has_paragraph: bool,
}

#[derive(Debug)]
struct FrameState {
    auto_height: bool,
    depth: usize,
    frame_index: u32,
    parent_numeric_id: Option<u32>,
    element_id: Option<String>,
    bounds: Rect,
    style_name: Option<String>,
    text_style_name: Option<String>,
    transform: AffineTransform,
    has_text_box: bool,
    text_box_depth: Option<usize>,
    paragraph_depth: Option<usize>,
    paragraph_count: u32,
    text: String,
    text_runs: Vec<OdfTextRun>,
    paragraph_style_names: Vec<Option<String>>,
    paragraph_list_layouts: Vec<Option<OdfListParagraphLayout>>,
    list_text: OdfListTextState,
    span_depth: Option<usize>,
    span_style_name: Option<String>,
    image_element_count: u32,
    images: Vec<ImageRepresentation>,
    active_image: Option<ActiveImage>,
    media: Option<MediaRepresentation>,
    table: Option<OdpTableState>,
    chart: Option<OdpBasicChart>,
    math: Option<OdfMath>,
    placeholder: Option<&'static str>,
}

struct ChartFrame {
    frame_index: u32,
    parent_numeric_id: Option<u32>,
    element_id: Option<String>,
    bounds: Rect,
    transform: AffineTransform,
}

impl From<&FrameState> for ChartFrame {
    fn from(frame: &FrameState) -> Self {
        Self {
            frame_index: frame.frame_index,
            parent_numeric_id: frame.parent_numeric_id,
            element_id: frame.element_id.clone(),
            bounds: frame.bounds,
            transform: frame.transform,
        }
    }
}

#[derive(Debug)]
struct MediaRepresentation {
    href: String,
    declared_media_type: Option<String>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum OdpChartKind {
    Bar,
    Line,
    Scatter,
    Pie,
}

#[derive(Debug)]
pub(super) struct OdpBasicChart {
    pub(super) kind: OdpChartKind,
    pub(super) title: Option<String>,
    pub(super) border_visible: bool,
    pub(super) categories: Vec<String>,
    pub(super) series: Vec<OdpChartSeries>,
    pub(super) show_legend: bool,
    pub(super) stacked: bool,
    pub(super) category_axis_at_end: bool,
    pub(super) value_axis_at_end: bool,
    pub(super) plot_area: Option<Rect>,
    pub(super) value_axis_visible: bool,
    pub(super) wall: Option<OdpChartWall>,
    pub(super) source_part: String,
}

#[derive(Clone, Copy, Debug)]
pub(super) struct OdpChartWall {
    color: u32,
    opacity_gradient: Option<GradientDefinition>,
}

#[derive(Debug)]
pub(super) struct OdpChartSeries {
    pub(super) label: Option<String>,
    pub(super) color: Option<u32>,
    pub(super) values: Vec<f32>,
    pub(super) domains: Vec<f32>,
    pub(super) custom_labels: Vec<Option<String>>,
    pub(super) point_explosions: Vec<f32>,
    pub(super) data_labels: Option<crate::format::odf_chart::OdfChartDataLabels>,
}

#[derive(Debug)]
struct OdpChartSeriesSpec {
    values_range: String,
    label_cell_address: Option<String>,
    color: Option<u32>,
    custom_labels: Vec<Option<String>>,
    domain_range: Option<String>,
    point_explosions: Vec<f32>,
    data_labels: Option<crate::format::odf_chart::OdfChartDataLabels>,
}

#[derive(Clone, Debug)]
struct OdpChartCell {
    value: Option<f32>,
    text: String,
    paragraph_depth: Option<usize>,
}

#[derive(Debug)]
struct OdpChartRow {
    depth: usize,
    cells: Vec<OdpChartCell>,
}

#[derive(Debug)]
struct OdpChartParseState {
    depth: usize,
    kind: OdpChartKind,
    title: String,
    title_depth: Option<usize>,
    title_paragraph_depth: Option<usize>,
    border_visible: bool,
    categories_range: Option<String>,
    series: Vec<OdpChartSeriesSpec>,
    show_legend: bool,
    stacked: bool,
    category_axis_at_end: bool,
    value_axis_at_end: bool,
    chart_size: Option<PageSize>,
    plot_area: Option<Rect>,
    value_axis_visible: bool,
    wall: Option<OdpChartWall>,
    rows: Vec<Vec<OdpChartCell>>,
    row: Option<OdpChartRow>,
    cell: Option<(usize, OdpChartCell)>,
}

impl OdpChartParseState {
    fn start_paragraph(&mut self, depth: usize, empty: bool) {
        if let Some((_, cell)) = self.cell.as_mut() {
            cell.paragraph_depth = (!empty).then_some(depth);
        } else if self.title_depth.is_some() {
            if !self.title.is_empty() {
                self.title.push('\n');
            }
            self.title_paragraph_depth = (!empty).then_some(depth);
        }
    }

    fn end_paragraph(&mut self, depth: usize) {
        if let Some((_, cell)) = self.cell.as_mut()
            && cell.paragraph_depth == Some(depth)
        {
            cell.paragraph_depth = None;
        } else if self.title_paragraph_depth == Some(depth) {
            self.title_paragraph_depth = None;
        }
    }

    fn push_text(&mut self, text: &str) {
        if let Some((_, cell)) = self
            .cell
            .as_mut()
            .filter(|(_, cell)| cell.paragraph_depth.is_some())
        {
            cell.text.push_str(text);
        } else if self.title_paragraph_depth.is_some() {
            self.title.push_str(text);
        }
    }
}

#[derive(Debug)]
struct GroupState {
    depth: usize,
    numeric_id: u32,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum OdfConnectorKind {
    Line,
    Standard,
}

#[derive(Clone, Debug)]
pub(super) struct OdfConnectorAttachment {
    pub(super) shape_id: String,
    pub(super) glue_point: Option<u32>,
}

#[derive(Clone, Debug)]
pub(super) struct OdfConnector {
    pub(super) kind: OdfConnectorKind,
    pub(super) start: (f32, f32),
    pub(super) end: (f32, f32),
    pub(super) attachments: [Option<OdfConnectorAttachment>; 2],
}

#[derive(Clone, Copy, Debug)]
pub(super) struct OdfShapeAnchor {
    pub(super) bounds: Rect,
    pub(super) transform: AffineTransform,
}

#[derive(Debug)]
struct DrawShapeState {
    depth: usize,
    parent_numeric_id: Option<u32>,
    element_id: Option<String>,
    path: String,
    bounds: Rect,
    geometry: Geometry,
    connector: Option<OdfConnector>,
    geometry_fallback: bool,
    enhanced_geometry: Option<OdfEnhancedGeometryState>,
    text_path: Option<String>,
    text_area: Option<Rect>,
    style_name: Option<String>,
    transform: AffineTransform,
    paragraph_depth: Option<usize>,
    paragraph_count: u32,
    text: String,
    text_runs: Vec<OdfTextRun>,
    paragraph_style_names: Vec<Option<String>>,
    paragraph_list_layouts: Vec<Option<OdfListParagraphLayout>>,
    list_text: OdfListTextState,
    span_depth: Option<usize>,
    span_style_name: Option<String>,
}

#[derive(Debug)]
struct OdfEnhancedGeometryState {
    depth: usize,
    view_box: (f32, f32, f32, f32),
    path_stretchpoint_x: Option<f32>,
    path_stretchpoint_y: Option<f32>,
    path: String,
    text_areas: Option<String>,
    modifiers: Vec<f32>,
    equations: Vec<(String, String)>,
    mirror_horizontal: bool,
    mirror_vertical: bool,
}

#[derive(Debug)]
struct OdpTableState {
    depth: usize,
    parent_numeric_id: Option<u32>,
    element_id: Option<String>,
    bounds: Rect,
    transform: AffineTransform,
    column_widths: Vec<f32>,
    template_name: Option<String>,
    use_banding_columns: bool,
    use_banding_rows: bool,
    use_first_column: bool,
    use_first_row: bool,
    use_last_column: bool,
    use_last_row: bool,
    rows: Vec<OdpTableRowState>,
    row: Option<OdpTableRowState>,
    cell: Option<OdpTableCellState>,
}

#[derive(Debug)]
struct OdpTableRowState {
    depth: usize,
    minimum_height: f32,
    cells: Vec<OdpTableCellState>,
}

#[derive(Debug)]
struct OdpTableCellState {
    depth: usize,
    element_id: Option<String>,
    style_name: Option<String>,
    column_span: usize,
    row_span: usize,
    covered: bool,
    paragraph_depth: Option<usize>,
    paragraph_count: u32,
    text: String,
    text_runs: Vec<OdfTextRun>,
    paragraph_style_names: Vec<Option<String>>,
    paragraph_list_layouts: Vec<Option<OdfListParagraphLayout>>,
    list_text: OdfListTextState,
    span_depth: Option<usize>,
    span_style_name: Option<String>,
}

#[derive(Debug)]
struct ImageRepresentation {
    href: Option<String>,
    declared_media_type: Option<String>,
    inline_base64: Option<String>,
}

#[derive(Debug)]
struct ActiveImage {
    depth: usize,
    binary_data_depth: Option<usize>,
    representation: ImageRepresentation,
}

struct ImageResources {
    manifest_media_types: HashMap<String, String>,
    cache: HashMap<String, ImageCacheEntry>,
    materialized_bytes: usize,
}

struct ParsedContent {
    units: Vec<Unit>,
    objects: Vec<Object>,
    diagnostics: Vec<Diagnostic>,
}

fn push_image_representation(
    frame: &mut FrameState,
    representation: ImageRepresentation,
    image_reference_count: &mut usize,
    image_reference_limit: usize,
) -> Result<(), Diagnostic> {
    if representation.inline_base64.is_none() && representation.href.is_some() {
        *image_reference_count = image_reference_count
            .checked_add(1)
            .ok_or_else(|| relationship_limit_error(CONTENT_PART))?;
        if *image_reference_count > image_reference_limit {
            return Err(relationship_limit_error(CONTENT_PART));
        }
    }
    frame.images.push(representation);
    Ok(())
}

fn parse_odf_shape_anchors(
    bytes: &[u8],
    package: &dyn OdpSource,
) -> Result<HashMap<String, OdfShapeAnchor>, Diagnostic> {
    let mut anchors = HashMap::new();
    parse_xml(bytes, package.limits(), |event| {
        let XmlEvent::StartElement {
            name, attributes, ..
        } = event
        else {
            return Ok(());
        };
        let local = local_name(name);
        if !matches!(
            local,
            "frame"
                | "rect"
                | "ellipse"
                | "custom-shape"
                | "caption"
                | "path"
                | "polygon"
                | "polyline"
        ) {
            return Ok(());
        }
        let Some(element_id) = optional_attribute(&attributes, "id", CONTENT_PART)? else {
            return Ok(());
        };
        let bounds = Rect {
            x: optional_length(&attributes, "x", CONTENT_PART)?.unwrap_or(0.0),
            y: optional_length(&attributes, "y", CONTENT_PART)?.unwrap_or(0.0),
            width: optional_length(&attributes, "width", CONTENT_PART)?.unwrap_or(0.0),
            height: optional_length(&attributes, "height", CONTENT_PART)?.unwrap_or(0.0),
        };
        if !bounds.is_valid() {
            return Ok(());
        }
        let transform = optional_attribute(&attributes, "transform", CONTENT_PART)?
            .map(|value| parse_odf_transform(&value))
            .transpose()?
            .unwrap_or(AffineTransform::IDENTITY);
        anchors.insert(element_id, OdfShapeAnchor { bounds, transform });
        Ok(())
    })
    .map(|_| anchors)
    .map_err(|error| with_part(error, CONTENT_PART))
}

#[cfg(all(test, feature = "odf-formats"))]
fn parse_content(
    package: &dyn OdpSource,
    page_sizes: &HashMap<String, PageSize>,
    styles: &StyleCatalog,
) -> Result<ParsedContent, Diagnostic> {
    parse_content_with_font_metrics(package, page_sizes, styles, &FontMetricTable::default())
}

fn parse_content_with_font_metrics(
    package: &dyn OdpSource,
    page_sizes: &HashMap<String, PageSize>,
    styles: &StyleCatalog,
    font_metrics: &FontMetricTable,
) -> Result<ParsedContent, Diagnostic> {
    let mut image_resources = ImageResources {
        manifest_media_types: parse_manifest_media_types(package)?,
        cache: HashMap::new(),
        materialized_bytes: 0,
    };
    let bytes = package.required_part(CONTENT_PART)?;
    let chart_styles = parse_odp_chart_styles(&bytes, package, CONTENT_PART)?;
    let shape_anchors = parse_odf_shape_anchors(&bytes, package)?;
    let mut depth = 0_usize;
    let mut page = None;
    let mut frame = None;
    let mut shape = None;
    let mut table = None;
    let mut chart: Option<OdpChartParseState> = None;
    let mut notes_depth = None;
    let mut groups: Vec<GroupState> = Vec::new();
    let mut units = Vec::new();
    let mut objects = Vec::new();
    let mut diagnostics = Vec::new();
    diagnostics.extend(styles.diagnostics.iter().cloned());
    if styles.invalid_master_shapes != 0 {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Omitted,
                "invalid ODP master shapes were omitted",
            )
            .in_part(STYLES_PART)
            .with_detail("count", styles.invalid_master_shapes.to_string()),
        );
    }
    let mut image_reference_count = 0_usize;
    let mut unsupported_chart_reported = false;
    let mut unsupported_geometry_reported = false;

    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if notes_depth.is_some() {
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }
                if local == "notes" && page.is_some() {
                    if !empty {
                        notes_depth = Some(depth);
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }
                if matches!(
                    name,
                    "draw:custom-shape"
                        | "draw:rect"
                        | "draw:ellipse"
                        | "draw:line"
                        | "draw:path"
                        | "draw:polygon"
                        | "draw:polyline"
                        | "draw:connector"
                        | "draw:caption"
                ) && page.is_none() && !unsupported_geometry_reported
                {
                    diagnostics.push(
                        Diagnostic::warning(
                            DiagnosticCode::UnsupportedFeature,
                            Phase::Render,
                            Fidelity::Omitted,
                            "ODP shape geometry objects are not rendered",
                        )
                        .in_part(CONTENT_PART),
                    );
                    unsupported_geometry_reported = true;
                }
                match local {
                    "page" => {
                        if page.is_some() {
                            return Err(format_error(CONTENT_PART, "nested slides are invalid"));
                        }
                        let master = optional_attribute(&attributes, "master-page-name", CONTENT_PART)?
                            .unwrap_or_default();
                        let size = page_sizes.get(&master).copied().unwrap_or_else(|| {
                            diagnostics.push(Diagnostic::warning(
                                DiagnosticCode::ApproximateLayout,
                                Phase::Layout,
                                Fidelity::Approximate,
                                format!("slide has no usable master page {master:?}; using default page size 28 x 21 cm"),
                            ).in_part(CONTENT_PART));
                            PageSize {
                                width: 28.0 / 2.54 * CSS_PIXELS_PER_INCH,
                                height: 21.0 / 2.54 * CSS_PIXELS_PER_INCH,
                            }
                        });
                        let unit_index = u32::try_from(units.len()).map_err(|_| {
                            format_error(CONTENT_PART, "slide count exceeds supported range")
                        })?;
                        let unit_name = optional_attribute(&attributes, "name", CONTENT_PART)?
                            .filter(|name| !name.is_empty())
                            .unwrap_or_else(|| format!("Slide {}", unit_index + 1));
                        let source_id = optional_attribute(&attributes, "id", CONTENT_PART)?
                            .filter(|id| !id.is_empty());
                        let hidden = optional_attribute(
                            &attributes,
                            "visibility",
                            CONTENT_PART,
                        )?
                        .is_some_and(|visibility| visibility == "hidden");
                        units.push(Unit {
                            kind: UnitKind::Slide,
                            index: unit_index,
                            id: format!("unit:{unit_index}"),
                            name: unit_name,
                            width: size.width,
                            height: size.height,
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
                            slide: Some(SlideMetadata {
                                source_id,
                                source_part: Some(CONTENT_PART.to_owned()),
                                speaker_notes: None,
                                speaker_notes_part: None,
                                speaker_note_paragraphs: Vec::new(),
                                number: unit_index + 1,
                                hidden,
                            }),
                        });
                        let page_style = optional_attribute(&attributes, "style-name", CONTENT_PART)?;
                        let background_style = page_style.as_ref().filter(|name| {
                            styles.style_patches.get(*name).is_some_and(|patch| patch.fill.is_some())
                        }).or_else(|| styles.master_backgrounds.get(&master));
                        if let Some(style_name) = background_style {
                            push_odp_background(
                                unit_index,
                                size,
                                &style_name,
                                styles,
                                OdpBackgroundSink {
                                    objects: &mut objects,
                                    object_limit: package.limits().max_document_objects,
                                    materialized_image_bytes: &mut image_resources
                                        .materialized_bytes,
                                    materialized_image_limit: package
                                        .limits()
                                        .max_total_uncompressed_bytes,
                                },
                            )?;
                        }
                        if let Some(templates) = styles.masters.get(&master) {
                            for template in templates {
                                push_master_template(
                                    template,
                                    unit_index,
                                    &mut objects,
                                    &mut diagnostics,
                                    styles,
                                    package.limits().max_document_objects,
                                    &mut image_resources.materialized_bytes,
                                    package.limits().max_total_uncompressed_bytes,
                                )?;
                            }
                        }
                        if !empty {
                            page = Some(PageState {
                                depth,
                                unit_index,
                                frame_index: 0,
                            });
                        }
                    }
                    "g" if page.is_some() => {
                        if objects.len() >= package.limits().max_document_objects {
                            return Err(object_limit_error());
                        }
                        let numeric_id = u32::try_from(objects.len()).map_err(|_| {
                            format_error(CONTENT_PART, "object count exceeds supported range")
                        })?;
                        let parent_numeric_id = groups.last().map(|group| group.numeric_id);
                        let element_id = optional_attribute(&attributes, "id", CONTENT_PART)?;
                        let transform = optional_attribute(
                            &attributes,
                            "transform",
                            CONTENT_PART,
                        )?
                        .map(|value| parse_odf_transform(&value))
                        .transpose()?
                        .unwrap_or(AffineTransform::IDENTITY);
                        let unit_index = page.as_ref().map(|page| page.unit_index).unwrap_or(0);
                        objects.push(Object {
                            numeric_id,
                            parent_numeric_id,
                            stable_id: format!("object:{numeric_id}"),
                            parent_stable_id: parent_numeric_id
                                .map(|parent| format!("object:{parent}")),
                            kind: ObjectKind::Group,
                            unit_index,
                            bounds: Rect::default(),
                            z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
                            text: None,
                            source: SourceRef {
                                part: CONTENT_PART.to_owned(),
                                mapping: if element_id.is_some() {
                                    MappingQuality::Exact
                                } else {
                                    MappingQuality::Derived
                                },
                                locator: SourceLocator::OdpElement {
                                    element_id,
                                    path: format!(
                                        "/office:document-content/office:body/office:presentation/draw:page[{}]/draw:g[{}]",
                                        unit_index + 1,
                                        numeric_id + 1
                                    ),
                                    row: None,
                                    column: None,
                                },
                            },
                            visual: Visual::Layer {
                                transform,
                                opacity: 1.0,
                                blend_mode: crate::model::BlendMode::Normal,
                                visual: Box::new(Visual::None),
                            },
                        });
                        if !empty {
                            groups.push(GroupState { depth, numeric_id });
                        }
                    }
                    "rect" | "ellipse" | "line" | "custom-shape" | "caption"
                    | "path" | "polygon" | "polyline" | "connector"
                        if page.is_some() && shape.is_none() =>
                    {
                        let page_state = page.as_ref().ok_or_else(|| {
                            format_error(CONTENT_PART, "shape has no parent slide")
                        })?;
                        let current = create_draw_shape(
                            local,
                            depth,
                            page_state.unit_index,
                            groups.last().map(|group| group.numeric_id),
                            &attributes,
                            objects.len(),
                        )?;
                        if empty {
                            push_draw_shape(
                                current,
                                page_state.unit_index,
                                &mut objects,
                                &mut diagnostics,
                                styles,
                                &shape_anchors,
                                package.limits().max_document_objects,
                                &mut image_resources.materialized_bytes,
                                package.limits().max_total_uncompressed_bytes,
                            )?;
                        } else {
                            shape = Some(current);
                        }
                    }
                    "frame" if page.is_some() && frame.is_none() => {
                        let page_state = page.as_mut().ok_or_else(|| {
                            format_error(CONTENT_PART, "frame has no parent slide")
                        })?;
                        let frame_index = page_state.frame_index;
                        page_state.frame_index =
                            page_state.frame_index.checked_add(1).ok_or_else(|| {
                                format_error(CONTENT_PART, "frame count exceeds supported range")
                            })?;
                        let transform_value =
                            optional_attribute(&attributes, "transform", CONTENT_PART)?;
                        let transform = transform_value
                            .as_deref()
                            .map(parse_odf_transform)
                            .transpose()?
                            .unwrap_or(AffineTransform::IDENTITY);
                        let x = optional_length(&attributes, "x", CONTENT_PART)?;
                        let y = optional_length(&attributes, "y", CONTENT_PART)?;
                        let (x, y) = match (x, y) {
                            (Some(x), Some(y)) => (x, y),
                            (None, None)
                                if transform_value
                                    .as_deref()
                                    .is_some_and(|value| !value.trim().is_empty()) =>
                            {
                                (0.0, 0.0)
                            }
                            (None, _) => {
                                return Err(format_error(CONTENT_PART, "element is missing x"));
                            }
                            (_, None) => {
                                return Err(format_error(CONTENT_PART, "element is missing y"));
                            }
                        };
                        let height = optional_length(&attributes, "height", CONTENT_PART)?;
                        let bounds = Rect {
                            x,
                            y,
                            width: required_length(&attributes, "width", CONTENT_PART)?,
                            height: height.unwrap_or(1.0),
                        };
                        if !bounds.is_valid() {
                            return Err(format_error(CONTENT_PART, "frame bounds are invalid"));
                        }
                        if !empty {
                            frame = Some(FrameState {
                                auto_height: height.is_none(),
                                depth,
                                frame_index,
                                parent_numeric_id: groups.last().map(|group| group.numeric_id),
                                element_id: optional_attribute(&attributes, "id", CONTENT_PART)?,
                                bounds,
                                style_name: optional_attribute(
                                    &attributes,
                                    "style-name",
                                    CONTENT_PART,
                                )?,
                                text_style_name: optional_attribute(&attributes, "text-style-name", CONTENT_PART)?,
                                transform,
                                has_text_box: false,
                                text_box_depth: None,
                                paragraph_depth: None,
                                paragraph_count: 0,
                                text: String::new(),
                                text_runs: Vec::new(),
                                paragraph_style_names: Vec::new(),
                                paragraph_list_layouts: Vec::new(),
                                list_text: OdfListTextState::default(),
                                span_depth: None,
                                span_style_name: None,
                                image_element_count: 0,
                                images: Vec::new(),
                                active_image: None,
                                media: None,
                                table: None,
                                chart: None,
                                math: None,
                                placeholder: None,
                            });
                        }
                    }
                    "chart" if frame.is_some() && chart.is_none() => {
                        let kind = optional_attribute(&attributes, "class", CONTENT_PART)?
                            .and_then(|value| match local_name(&value) {
                                "bar" | "column" => Some(OdpChartKind::Bar),
                                "line" => Some(OdpChartKind::Line),
                                "scatter" => Some(OdpChartKind::Scatter),
                                "circle" | "pie" => Some(OdpChartKind::Pie),
                                _ => None,
                            });
                        if let Some(kind) = kind
                            && !empty
                        {
                            let border_visible = optional_attribute(
                                &attributes,
                                "style-name",
                                CONTENT_PART,
                            )?
                            .as_deref()
                            .and_then(|name| chart_styles.get(name))
                            .and_then(|style| style.border_visible)
                            .unwrap_or(true);
                            chart = Some(OdpChartParseState {
                                depth,
                                kind,
                                title: String::new(),
                                title_depth: None,
                                title_paragraph_depth: None,
                                border_visible,
                                categories_range: None,
                                series: Vec::new(),
                                show_legend: false,
                                stacked: false,
                                category_axis_at_end: false,
                                value_axis_at_end: false,
                                chart_size: match (
                                    optional_length(&attributes, "width", CONTENT_PART)?,
                                    optional_length(&attributes, "height", CONTENT_PART)?,
                                ) {
                                    (Some(width), Some(height)) => Some(PageSize { width, height }),
                                    _ => None,
                                },
                                plot_area: None,
                                value_axis_visible: true,
                                wall: None,
                                rows: Vec::new(),
                                row: None,
                                cell: None,
                            });
                        } else if let Some(current) = frame.as_mut() {
                            current.placeholder = Some("Chart");
                            if !unsupported_chart_reported {
                                diagnostics.push(
                                    Diagnostic::warning(
                                        DiagnosticCode::UnsupportedFeature,
                                        Phase::Render,
                                        Fidelity::Approximate,
                                        "ODP chart content has no usable cached series and uses a source-mapped static placeholder",
                                    )
                                    .in_part(CONTENT_PART),
                                );
                                unsupported_chart_reported = true;
                            }
                        }
                    }
                    "categories" if chart.is_some() => {
                        if let Some(cell_range) =
                            optional_attribute(&attributes, "cell-range-address", CONTENT_PART)?
                            && let Some(current) = chart.as_mut()
                        {
                            current.categories_range = Some(cell_range);
                        }
                    }
                    "plot-area" if chart.is_some() => {
                        if let Some(style_name) =
                            optional_attribute(&attributes, "style-name", CONTENT_PART)?
                            && let Some(stacked) = chart_styles
                                .get(&style_name)
                                .and_then(|style| style.stacked)
                        {
                            chart.as_mut().expect("checked above").stacked = stacked;
                        }
                    }
                    "coordinate-region" if chart.is_some() => {
                        if let (Some(x), Some(y), Some(width), Some(height)) = (
                            optional_length(&attributes, "x", CONTENT_PART)?,
                            optional_length(&attributes, "y", CONTENT_PART)?,
                            optional_length(&attributes, "width", CONTENT_PART)?,
                            optional_length(&attributes, "height", CONTENT_PART)?,
                        ) {
                            chart.as_mut().expect("checked above").plot_area =
                                Some(Rect { x, y, width, height });
                        }
                    }
                    "title"
                        if chart
                            .as_ref()
                            .is_some_and(|chart| depth == chart.depth.saturating_add(1)) =>
                    {
                        chart.as_mut().expect("checked above").title_depth =
                            (!empty).then_some(depth);
                    }
                    "legend" if chart.is_some() => {
                        chart.as_mut().expect("checked above").show_legend = true;
                    }
                    "axis" if chart.is_some() => {
                        let dimension = optional_attribute(
                            &attributes,
                            "dimension",
                            CONTENT_PART,
                        )?;
                        let style = optional_attribute(
                            &attributes,
                            "style-name",
                            CONTENT_PART,
                        )?
                        .as_deref()
                        .and_then(|name| chart_styles.get(name));
                        let current = chart.as_mut().expect("checked above");
                        if dimension
                            .as_deref()
                            .is_some_and(|dimension| local_name(dimension) == "y")
                        {
                            if style.and_then(|style| style.visible) == Some(false) {
                                current.value_axis_visible = false;
                            }
                            current.value_axis_at_end =
                                style.and_then(|style| style.axis_at_end).unwrap_or(false);
                        } else if dimension
                            .as_deref()
                            .is_some_and(|dimension| local_name(dimension) == "x")
                        {
                            current.category_axis_at_end =
                                style.and_then(|style| style.axis_at_end).unwrap_or(false);
                        }
                    }
                    "wall" if chart.is_some() => {
                        if let Some(style_name) =
                            optional_attribute(&attributes, "style-name", CONTENT_PART)?
                            && let Some(style) = chart_styles.get(&style_name)
                            && let Some(color) = style.color
                        {
                            chart.as_mut().expect("checked above").wall = Some(OdpChartWall {
                                color,
                                opacity_gradient: style
                                    .opacity_gradient
                                    .as_deref()
                                    .and_then(|name| styles.opacity_gradients.get(name))
                                    .copied(),
                            });
                        }
                    }
                    "series" if chart.is_some() => {
                        if let Some(values_range) = optional_attribute(
                            &attributes,
                            "values-cell-range-address",
                            CONTENT_PART,
                        )? && let Some(current) = chart.as_mut()
                        {
                            let style_name =
                                optional_attribute(&attributes, "style-name", CONTENT_PART)?;
                            let chart_style = style_name
                                .as_deref()
                                .and_then(|name| chart_styles.get(name));
                            current.series.push(OdpChartSeriesSpec {
                                values_range,
                                label_cell_address: optional_attribute(
                                    &attributes,
                                    "label-cell-address",
                                    CONTENT_PART,
                                )?,
                                color: chart_style.and_then(|style| style.color),
                                custom_labels: Vec::new(),
                                domain_range: None,
                                point_explosions: Vec::new(),
                                data_labels: chart_style.map(|style| {
                                    style.chart_props.data_labels(
                                        11.0 * 96.0 / 72.0,
                                        "Arial",
                                        0x0000_00ff,
                                        false,
                                    )
                                }),
                            });
                        }
                    }
                    "domain" if chart.is_some() => {
                        if let Some(range) =
                            optional_attribute(&attributes, "cell-range-address", CONTENT_PART)?
                            && let Some(series) = chart
                                .as_mut()
                                .and_then(|chart| chart.series.last_mut())
                        {
                            series.domain_range = Some(range);
                        }
                    }
                    "data-point" if chart.is_some() => {
                        let repeat = optional_attribute(
                            &attributes,
                            "repeated",
                            CONTENT_PART,
                        )?
                        .and_then(|value| value.parse::<usize>().ok())
                        .unwrap_or(1)
                        .max(1);
                        let label = optional_attribute(
                            &attributes,
                            "custom-label-field",
                            CONTENT_PART,
                        )?
                        .filter(|label| !label.is_empty());
                        let explosion = optional_attribute(
                            &attributes,
                            "style-name",
                            CONTENT_PART,
                        )?
                        .and_then(|name| chart_styles.get(&name))
                        .and_then(|style| style.chart_props.pie_offset)
                        .unwrap_or(0.0);
                        if let Some(series) = chart
                            .as_mut()
                            .and_then(|chart| chart.series.last_mut())
                        {
                            series.custom_labels.extend(std::iter::repeat_n(label, repeat));
                            series
                                .point_explosions
                                .extend(std::iter::repeat_n(explosion, repeat));
                        }
                    }
                    "table" if chart.is_some() => {}
                    "table-row" if chart.is_some() => {
                        let current = chart.as_mut().ok_or_else(|| {
                            format_error(CONTENT_PART, "chart row has no parent chart")
                        })?;
                        if current.row.is_some() {
                            return Err(format_error(
                                CONTENT_PART,
                                "nested chart data rows are invalid",
                            ));
                        }
                        current.row = Some(OdpChartRow {
                            depth,
                            cells: Vec::new(),
                        });
                    }
                    "table-cell" | "covered-table-cell"
                        if chart.as_ref().is_some_and(|chart| chart.row.is_some()) =>
                    {
                        let current = chart.as_mut().ok_or_else(|| {
                            format_error(CONTENT_PART, "chart cell has no parent chart")
                        })?;
                        if current.cell.is_some() {
                            return Err(format_error(
                                CONTENT_PART,
                                "nested chart data cells are invalid",
                            ));
                        }
                        let cell = OdpChartCell {
                            value: optional_attribute(&attributes, "value", CONTENT_PART)?
                                .and_then(|value| value.parse::<f32>().ok())
                                .filter(|value| value.is_finite()),
                            text: optional_attribute(
                                &attributes,
                                "string-value",
                                CONTENT_PART,
                            )?
                            .unwrap_or_default(),
                            paragraph_depth: None,
                        };
                        let repeat = optional_attribute(
                            &attributes,
                            "number-columns-repeated",
                            CONTENT_PART,
                        )?
                        .and_then(|value| value.parse::<usize>().ok())
                        .unwrap_or(1)
                        .max(1);
                        if empty {
                            let row = current.row.as_mut().ok_or_else(|| {
                                format_error(CONTENT_PART, "chart cell has no row")
                            })?;
                            row.cells.extend(std::iter::repeat_n(cell, repeat));
                        } else {
                            current.cell = Some((depth, cell));
                        }
                    }
                    "table" if frame.is_some() && table.is_none() && chart.is_none() => {
                        let current = frame.as_ref().ok_or_else(|| {
                            format_error(CONTENT_PART, "table has no parent frame")
                        })?;
                        table = Some(OdpTableState {
                            depth,
                            parent_numeric_id: current.parent_numeric_id,
                            element_id: optional_attribute(&attributes, "id", CONTENT_PART)?,
                            bounds: current.bounds,
                            transform: current.transform,
                            column_widths: Vec::new(),
                            template_name: optional_attribute(
                                &attributes,
                                "template-name",
                                CONTENT_PART,
                            )?
                            .filter(|name| !name.trim().is_empty()),
                            use_banding_columns: odf_true_attribute(
                                &attributes,
                                "use-banding-columns-styles",
                                CONTENT_PART,
                            )?,
                            use_banding_rows: odf_true_attribute(
                                &attributes,
                                "use-banding-rows-styles",
                                CONTENT_PART,
                            )?,
                            use_first_column: odf_true_attribute(
                                &attributes,
                                "use-first-column-styles",
                                CONTENT_PART,
                            )?,
                            use_first_row: odf_true_attribute(
                                &attributes,
                                "use-first-row-styles",
                                CONTENT_PART,
                            )?,
                            use_last_column: odf_true_attribute(
                                &attributes,
                                "use-last-column-styles",
                                CONTENT_PART,
                            )?,
                            use_last_row: odf_true_attribute(
                                &attributes,
                                "use-last-row-styles",
                                CONTENT_PART,
                            )?,
                            rows: Vec::new(),
                            row: None,
                            cell: None,
                        });
                    }
                    "table-column" if table.is_some() && chart.is_none() => {
                        let style_name =
                            optional_attribute(&attributes, "style-name", CONTENT_PART)?;
                        let width = style_name
                            .as_ref()
                            .and_then(|name| styles.table_column_widths.get(name))
                            .copied()
                            .unwrap_or(0.0);
                        let repeat = optional_attribute(
                            &attributes,
                            "number-columns-repeated",
                            CONTENT_PART,
                        )?
                        .and_then(|value| value.parse::<usize>().ok())
                        .unwrap_or(1)
                        .max(1);
                        let current = table.as_mut().ok_or_else(|| {
                            format_error(CONTENT_PART, "column has no parent table")
                        })?;
                        current
                            .column_widths
                            .extend(std::iter::repeat_n(width, repeat));
                    }
                    "table-row" if table.is_some() => {
                        let current = table.as_mut().ok_or_else(|| {
                            format_error(CONTENT_PART, "row has no parent table")
                        })?;
                        if current.row.is_some() {
                            return Err(format_error(CONTENT_PART, "nested table rows are invalid"));
                        }
                        current.row = Some(OdpTableRowState {
                            depth,
                            minimum_height: optional_attribute(
                                &attributes,
                                "style-name",
                                CONTENT_PART,
                            )?
                            .as_ref()
                            .and_then(|name| styles.table_row_heights.get(name))
                            .copied()
                            .unwrap_or(0.0),
                            cells: Vec::new(),
                        });
                    }
                    "table-cell" | "covered-table-cell"
                        if table.as_ref().is_some_and(|table| table.row.is_some()) =>
                    {
                        let current = table.as_mut().ok_or_else(|| {
                            format_error(CONTENT_PART, "cell has no parent table")
                        })?;
                        if current.cell.is_some() {
                            return Err(format_error(CONTENT_PART, "nested table cells are invalid"));
                        }
                        current.cell = Some(OdpTableCellState {
                            depth,
                            element_id: optional_attribute(&attributes, "id", CONTENT_PART)?,
                            style_name: optional_attribute(
                                &attributes,
                                "style-name",
                                CONTENT_PART,
                            )?,
                            column_span: optional_attribute(
                                &attributes,
                                "number-columns-spanned",
                                CONTENT_PART,
                            )?
                            .and_then(|value| value.parse::<usize>().ok())
                            .unwrap_or(1)
                            .max(1),
                            row_span: optional_attribute(
                                &attributes,
                                "number-rows-spanned",
                                CONTENT_PART,
                            )?
                            .and_then(|value| value.parse::<usize>().ok())
                            .unwrap_or(1)
                            .max(1),
                            covered: local == "covered-table-cell",
                            paragraph_depth: None,
                            paragraph_count: 0,
                            text: String::new(),
                            text_runs: Vec::new(),
                            paragraph_style_names: Vec::new(),
                            paragraph_list_layouts: Vec::new(),
                            list_text: OdfListTextState::default(),
                            span_depth: None,
                            span_style_name: None,
                        });
                        if empty {
                            let cell = current.cell.take().ok_or_else(|| {
                                format_error(
                                    CONTENT_PART,
                                    "table-cell parser state ended unexpectedly",
                                )
                            })?;
                            current
                                .row
                                .as_mut()
                                .ok_or_else(|| {
                                    format_error(CONTENT_PART, "table cell has no row")
                                })?
                                .cells
                                .push(cell);
                        }
                    }
                    "image" => {
                        if let Some(current) = frame.as_mut() {
                            if current.active_image.is_some() {
                                return Err(format_error(
                                    CONTENT_PART,
                                    "nested draw:image elements are invalid",
                                ));
                            }
                            current.image_element_count =
                                current.image_element_count.checked_add(1).ok_or_else(|| {
                                    format_error(
                                        CONTENT_PART,
                                        "image count exceeds supported range",
                                    )
                                })?;
                            let representation = ImageRepresentation {
                                href: optional_attribute(&attributes, "href", CONTENT_PART)?,
                                declared_media_type: optional_attribute(
                                    &attributes,
                                    "mime-type",
                                    CONTENT_PART,
                                )?
                                .filter(|media_type| !media_type.trim().is_empty()),
                                inline_base64: None,
                            };
                            if empty {
                                push_image_representation(
                                    current,
                                    representation,
                                    &mut image_reference_count,
                                    package.limits().max_relationship_edges,
                                )?;
                            } else {
                                current.active_image = Some(ActiveImage {
                                    depth,
                                    binary_data_depth: None,
                                    representation,
                                });
                            }
                        }
                    }
                    "plugin" => {
                        if let Some(current) = frame.as_mut() {
                            let href = optional_attribute(&attributes, "href", CONTENT_PART)?
                                .ok_or_else(|| {
                                    format_error(CONTENT_PART, "draw:plugin is missing xlink:href")
                                })?;
                            image_reference_count = image_reference_count
                                .checked_add(1)
                                .ok_or_else(|| relationship_limit_error(CONTENT_PART))?;
                            if image_reference_count > package.limits().max_relationship_edges {
                                return Err(relationship_limit_error(CONTENT_PART));
                            }
                            current.media = Some(MediaRepresentation {
                                href,
                                declared_media_type: optional_attribute(
                                    &attributes,
                                    "mime-type",
                                    CONTENT_PART,
                                )?
                                .filter(|media_type| !media_type.trim().is_empty()),
                            });
                        }
                    }
                    "binary-data" => {
                        if let Some(image) = frame
                            .as_mut()
                            .and_then(|current| current.active_image.as_mut())
                        {
                            if image.representation.inline_base64.is_some() {
                                return Err(format_error(
                                    CONTENT_PART,
                                    "draw:image contains multiple office:binary-data elements",
                                ));
                            }
                            image.representation.inline_base64 = Some(String::new());
                            image.binary_data_depth = (!empty).then_some(depth);
                        }
                    }
                    "text-box" => {
                        if let Some(current) = frame.as_mut() {
                            current.has_text_box = true;
                            current.text_box_depth = (!empty).then_some(depth);
                        }
                    }
                    "list" => {
                        let style_name =
                            optional_attribute(&attributes, "style-name", CONTENT_PART)?;
                        let reset_numbering = optional_attribute(
                            &attributes,
                            "continue-numbering",
                            CONTENT_PART,
                        )?
                        .is_some_and(|value| matches!(value.as_str(), "false" | "0"));
                        if !empty {
                            let active = ActiveOdfList {
                                depth,
                                style_name,
                                reset_numbering,
                                continue_numbering: optional_attribute(&attributes, "continue-numbering", CONTENT_PART)?.is_some_and(|v| matches!(v.as_str(), "true" | "1")),
                                item_depth: None,
                                item_count: 0,
                                item_start_value: None,
                                item_has_paragraph: false,
                            };
                            if let Some(current) = table
                                .as_mut()
                                .and_then(|table| table.cell.as_mut())
                            {
                                current.list_text.stack.push(active);
                            } else if let Some(current) = shape.as_mut() {
                                current.list_text.stack.push(active);
                            } else if let Some(current) = frame
                                .as_mut()
                                .filter(|current| current.text_box_depth.is_some())
                            {
                                current.list_text.stack.push(active);
                            }
                        }
                    }
                    "list-item" => {
                        if !empty {
                            let start_value =
                                optional_attribute(&attributes, "start-value", CONTENT_PART)?
                                    .map(|value| {
                                        value.parse::<u32>().map_err(|_| {
                                            format_error(
                                                CONTENT_PART,
                                                "ODF list item start value exceeds the supported range",
                                            )
                                        })
                                    })
                                    .transpose()?;
                            let active = if let Some(current) = table
                                .as_mut()
                                .and_then(|table| table.cell.as_mut())
                            {
                                current.list_text.stack.last_mut()
                            } else if let Some(current) = shape.as_mut() {
                                current.list_text.stack.last_mut()
                            } else {
                                frame
                                    .as_mut()
                                    .filter(|current| current.text_box_depth.is_some())
                                    .and_then(|current| current.list_text.stack.last_mut())
                            };
                            if let Some(active) = active {
                                active.item_depth = Some(depth);
                                active.item_count = active.item_count.saturating_add(1);
                                active.item_start_value = start_value;
                                active.item_has_paragraph = false;
                            }
                        }
                    }
                    "p" => {
                        if let Some(current) = chart.as_mut() {
                            current.start_paragraph(depth, empty);
                        } else if let Some(current) = table
                            .as_mut()
                            .and_then(|table| table.cell.as_mut())
                        {
                            if current.paragraph_count != 0 {
                                append_odf_paragraph_break(
                                    &mut current.text,
                                    &mut current.text_runs,
                                );
                            }
                            current.paragraph_count += 1;
                            current.paragraph_style_names.push(optional_attribute(
                                &attributes,
                                "style-name",
                                CONTENT_PART,
                            )?);
                            current.paragraph_list_layouts.push(begin_odf_list_paragraph(
                                &mut current.list_text,
                                styles,
                                None,
                            ));
                            current.paragraph_depth = (!empty).then_some(depth);
                            if empty {
                                discard_odf_pending_prefix(&mut current.list_text);
                            }
                        } else if let Some(current) = shape.as_mut() {
                            if current.paragraph_count != 0 {
                                append_odf_paragraph_break(
                                    &mut current.text,
                                    &mut current.text_runs,
                                );
                            }
                            current.paragraph_count += 1;
                            current.paragraph_style_names.push(optional_attribute(
                                &attributes,
                                "style-name",
                                CONTENT_PART,
                            )?);
                            current.paragraph_list_layouts.push(begin_odf_list_paragraph(
                                &mut current.list_text,
                                styles,
                                None,
                            ));
                            current.paragraph_depth = (!empty).then_some(depth);
                            if empty {
                                discard_odf_pending_prefix(&mut current.list_text);
                            }
                        } else if let Some(current) = frame
                            .as_mut()
                            .filter(|current| current.text_box_depth.is_some())
                        {
                            if current.paragraph_count != 0 {
                                append_odf_paragraph_break(
                                    &mut current.text,
                                    &mut current.text_runs,
                                );
                            }
                            current.paragraph_count += 1;
                            current.paragraph_style_names.push(optional_attribute(
                                &attributes,
                                "style-name",
                                CONTENT_PART,
                            )?.or_else(|| current.text_style_name.clone()));
                            current.paragraph_list_layouts.push(begin_odf_list_paragraph(
                                &mut current.list_text,
                                styles,
                                current.style_name.as_deref(),
                            ));
                            current.paragraph_depth = (!empty).then_some(depth);
                            if empty {
                                discard_odf_pending_prefix(&mut current.list_text);
                            }
                        }
                    }
                    "span" => {
                        if let Some(current) = table
                            .as_mut()
                            .and_then(|table| table.cell.as_mut())
                        {
                            if current.span_depth.is_some() {
                                return Err(format_error(
                                    CONTENT_PART,
                                    "nested ODP table text spans are invalid",
                                ));
                            }
                            if !empty {
                                current.span_depth = Some(depth);
                                current.span_style_name = optional_attribute(
                                    &attributes,
                                    "style-name",
                                    CONTENT_PART,
                                )?;
                            }
                        } else if let Some(current) = shape.as_mut() {
                            if current.span_depth.is_some() {
                                return Err(format_error(
                                    CONTENT_PART,
                                    "nested ODP text spans are invalid",
                                ));
                            }
                            if !empty {
                                current.span_depth = Some(depth);
                                current.span_style_name = optional_attribute(
                                    &attributes,
                                    "style-name",
                                    CONTENT_PART,
                                )?;
                            }
                        } else if let Some(current) = frame
                            .as_mut()
                            .filter(|frame| frame.text_box_depth.is_some())
                        {
                            if current.span_depth.is_some() {
                                return Err(format_error(
                                    CONTENT_PART,
                                    "nested ODP text spans are invalid",
                                ));
                            }
                            if !empty {
                                current.span_depth = Some(depth);
                                current.span_style_name = optional_attribute(
                                    &attributes,
                                    "style-name",
                                    CONTENT_PART,
                                )?;
                            }
                        }
                    }
                    "s" | "tab" | "line-break" => {
                        let text = odf_control_text(
                            local,
                            &attributes,
                            package.limits().max_xml_bytes,
                            CONTENT_PART,
                        )?;
                        if let Some(current) = chart
                            .as_mut()
                            .and_then(|chart| chart.cell.as_mut())
                            .map(|(_, cell)| cell)
                            .filter(|cell| cell.paragraph_depth.is_some())
                        {
                            current.text.push_str(&text);
                        } else if let Some(current) = table
                            .as_mut()
                            .and_then(|table| table.cell.as_mut())
                            .filter(|cell| cell.paragraph_depth.is_some())
                        {
                            let style_name = current_odf_text_style(
                                current.span_depth,
                                &current.span_style_name,
                                &current.paragraph_style_names,
                            );
                            append_odf_text_with_list_prefix(
                                &text,
                                &mut current.text,
                                &mut current.text_runs,
                                &mut current.list_text,
                                style_name.as_deref(),
                            );
                        } else if let Some(current) = shape
                            .as_mut()
                            .filter(|shape| shape.paragraph_depth.is_some())
                        {
                            let style_name = current_odf_text_style(
                                current.span_depth,
                                &current.span_style_name,
                                &current.paragraph_style_names,
                            );
                            append_odf_text_with_list_prefix(
                                &text,
                                &mut current.text,
                                &mut current.text_runs,
                                &mut current.list_text,
                                style_name.as_deref(),
                            );
                        } else if let Some(current) = frame
                            .as_mut()
                            .filter(|frame| frame.paragraph_depth.is_some())
                        {
                            let style_name = current_odf_text_style(
                                current.span_depth,
                                &current.span_style_name,
                                &current.paragraph_style_names,
                            );
                            append_odf_text_with_list_prefix(
                                &text,
                                &mut current.text,
                                &mut current.text_runs,
                                &mut current.list_text,
                                style_name.as_deref(),
                            );
                        }
                    }
                    "object" => {
                        let embedded_part = optional_attribute(&attributes, "href", CONTENT_PART)?
                            .and_then(|href| {
                                match resolve_odf_image_target(
                                    &href,
                                    package.limits().max_zip_path_bytes,
                                ) {
                                    Ok(OdfImageTarget::Embedded(target)) => Some(if target
                                        .ends_with(".xml")
                                    {
                                        target
                                    } else {
                                        format!("{target}/content.xml")
                                    }),
                                    Ok(OdfImageTarget::External) | Err(_) => None,
                                }
                            });
                        let parsed_math = if let Some(part) = embedded_part
                            .as_deref()
                            .filter(|part| package.has_part(part))
                        {
                            parse_odf_math(&package.required_part(part)?, package.limits(), part)?
                        } else {
                            None
                        };
                        let parsed_chart = if parsed_math.is_none() {
                            if let Some(part) = embedded_part
                                .as_deref()
                                .filter(|part| package.has_part(part))
                            {
                                parse_odp_chart_part_from_source(package, part)?
                            } else {
                                None
                            }
                        } else {
                            None
                        };
                        if let Some(current) = frame.as_mut() {
                            if let Some(chart) = parsed_chart {
                                current.chart = Some(chart);
                                current.placeholder = None;
                            } else if let Some(math) = parsed_math {
                                if math.text.is_empty() {
                                    diagnostics.push(
                                        Diagnostic::warning(
                                            DiagnosticCode::UnsupportedFeature,
                                            Phase::Render,
                                            Fidelity::Omitted,
                                            "ODP embedded MathML has no readable semantic content and was omitted",
                                        )
                                        .in_part(&math.source_part),
                                    );
                                } else {
                                    if math.approximate {
                                        diagnostics.push(
                                            Diagnostic::warning(
                                                DiagnosticCode::UnsupportedFeature,
                                                Phase::Render,
                                                Fidelity::Approximate,
                                                "complex ODP MathML was reduced to readable linear math text",
                                            )
                                            .in_part(&math.source_part),
                                        );
                                    }
                                    current.math = Some(math);
                                }
                                current.placeholder = None;
                            } else {
                                current.placeholder = Some("Embedded object");
                                if !unsupported_chart_reported {
                                    diagnostics.push(
                                        Diagnostic::warning(
                                            DiagnosticCode::UnsupportedFeature,
                                            Phase::Render,
                                            Fidelity::Approximate,
                                            "ODP embedded object has no supported cached bar, line, or pie chart data and uses a source-mapped static placeholder",
                                        )
                                        .in_part(
                                            embedded_part.as_deref().unwrap_or(CONTENT_PART),
                                        ),
                                    );
                                    unsupported_chart_reported = true;
                                }
                            }
                        }
                    }
                    "enhanced-geometry" => {
                        if let Some(current) = shape.as_mut() {
                            current.text_path = odf_text_path_mode(&attributes, CONTENT_PART)?;
                            if let Some(geometry) = odf_enhanced_geometry_kind(
                                &attributes,
                                current.bounds,
                                CONTENT_PART,
                            )? {
                                current.geometry = geometry;
                                current.geometry_fallback = false;
                            }
                            if let Some(state) =
                                begin_odf_enhanced_geometry(depth, &attributes, CONTENT_PART)?
                            {
                                if empty {
                                    current.text_area =
                                        parse_odf_enhanced_text_area(&state, current.bounds);
                                    if let Some(geometry) = parse_odf_enhanced_path_geometry(
                                        &state,
                                        current.bounds,
                                        FillRule::NonZero,
                                        &mut current.transform,
                                    ) {
                                        current.geometry = geometry;
                                        current.geometry_fallback = false;
                                    }
                                } else if current.enhanced_geometry.replace(state).is_some() {
                                    return Err(format_error(
                                        CONTENT_PART,
                                        "nested ODP enhanced geometry is invalid",
                                    ));
                                }
                            }
                        }
                    }
                    "equation" => {
                        if let Some(geometry) = shape
                            .as_mut()
                            .and_then(|shape| shape.enhanced_geometry.as_mut())
                            && let (Some(name), Some(formula)) = (
                                optional_attribute(&attributes, "name", CONTENT_PART)?,
                                optional_attribute(&attributes, "formula", CONTENT_PART)?,
                            )
                        {
                            geometry.equations.push((name, formula));
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
                if notes_depth.is_some() {
                    if local == "notes" && notes_depth == Some(depth) {
                        notes_depth = None;
                    }
                    return Ok(());
                }
                if local == "enhanced-geometry"
                    && let Some(current) = shape.as_mut()
                    && current
                        .enhanced_geometry
                        .as_ref()
                        .is_some_and(|geometry| geometry.depth == depth)
                {
                    let geometry = current.enhanced_geometry.take().ok_or_else(|| {
                        format_error(
                            CONTENT_PART,
                            "enhanced geometry parser state ended unexpectedly",
                        )
                    })?;
                    current.text_area =
                        parse_odf_enhanced_text_area(&geometry, current.bounds);
                    if let Some(geometry) = parse_odf_enhanced_path_geometry(
                        &geometry,
                        current.bounds,
                        FillRule::NonZero,
                        &mut current.transform,
                    ) {
                        current.geometry = geometry;
                        current.geometry_fallback = false;
                    }
                } else if local == "binary-data" {
                    if let Some(image) = frame
                        .as_mut()
                        .and_then(|current| current.active_image.as_mut())
                        && image.binary_data_depth == Some(depth)
                    {
                        image.binary_data_depth = None;
                    }
                } else if local == "image" {
                    if let Some(current) = frame.as_mut()
                        && current
                            .active_image
                            .as_ref()
                            .is_some_and(|image| image.depth == depth)
                    {
                        let image = current.active_image.take().ok_or_else(|| {
                            format_error(CONTENT_PART, "image parser state ended unexpectedly")
                        })?;
                        push_image_representation(
                            current,
                            image.representation,
                            &mut image_reference_count,
                            package.limits().max_relationship_edges,
                        )?;
                    }
                } else if local == "p" {
                    if let Some(current) = chart.as_mut() {
                        current.end_paragraph(depth);
                    } else if let Some(current) = table
                        .as_mut()
                        .and_then(|table| table.cell.as_mut())
                        && current.paragraph_depth == Some(depth)
                    {
                        discard_odf_pending_prefix(&mut current.list_text);
                        current.paragraph_depth = None;
                    } else if let Some(current) = shape.as_mut()
                        && current.paragraph_depth == Some(depth)
                    {
                        discard_odf_pending_prefix(&mut current.list_text);
                        current.paragraph_depth = None;
                    } else if let Some(current) = frame.as_mut()
                        && current.paragraph_depth == Some(depth)
                    {
                        discard_odf_pending_prefix(&mut current.list_text);
                        current.paragraph_depth = None;
                    }
                } else if local == "list-item" {
                    let active = if let Some(current) = table
                        .as_mut()
                        .and_then(|table| table.cell.as_mut())
                    {
                        current.list_text.stack.last_mut()
                    } else if let Some(current) = shape.as_mut() {
                        current.list_text.stack.last_mut()
                    } else {
                        frame
                            .as_mut()
                            .filter(|current| current.text_box_depth.is_some())
                            .and_then(|current| current.list_text.stack.last_mut())
                    };
                    if let Some(active) = active
                        && active.item_depth == Some(depth)
                    {
                        active.item_depth = None;
                        active.item_start_value = None;
                        active.item_has_paragraph = false;
                    }
                } else if local == "list" {
                    if let Some(current) = table
                        .as_mut()
                        .and_then(|table| table.cell.as_mut())
                    {
                        if current
                            .list_text
                            .stack
                            .last()
                            .is_some_and(|active| active.depth == depth)
                        {
                            current.list_text.stack.pop();
                        }
                    } else if let Some(current) = shape.as_mut() {
                        if current
                            .list_text
                            .stack
                            .last()
                            .is_some_and(|active| active.depth == depth)
                        {
                            current.list_text.stack.pop();
                        }
                    } else if let Some(current) = frame
                        .as_mut()
                        .filter(|current| current.text_box_depth.is_some())
                        && current
                            .list_text
                            .stack
                            .last()
                            .is_some_and(|active| active.depth == depth)
                    {
                        current.list_text.stack.pop();
                    }
                } else if local == "title"
                    && chart
                        .as_ref()
                        .is_some_and(|chart| chart.title_depth == Some(depth))
                {
                    chart.as_mut().expect("checked above").title_depth = None;
                } else if matches!(local, "table-cell" | "covered-table-cell")
                    && chart
                        .as_ref()
                        .and_then(|chart| chart.cell.as_ref())
                        .is_some_and(|(start, _)| *start == depth)
                {
                    let current = chart.as_mut().ok_or_else(|| {
                        format_error(CONTENT_PART, "chart parser state ended unexpectedly")
                    })?;
                    let (_, cell) = current.cell.take().ok_or_else(|| {
                        format_error(CONTENT_PART, "chart-cell parser state ended unexpectedly")
                    })?;
                    current
                        .row
                        .as_mut()
                        .ok_or_else(|| format_error(CONTENT_PART, "chart cell has no row"))?
                        .cells
                        .push(cell);
                } else if local == "table-row"
                    && chart
                        .as_ref()
                        .and_then(|chart| chart.row.as_ref())
                        .is_some_and(|row| row.depth == depth)
                {
                    let current = chart.as_mut().ok_or_else(|| {
                        format_error(CONTENT_PART, "chart parser state ended unexpectedly")
                    })?;
                    let row = current.row.take().ok_or_else(|| {
                        format_error(CONTENT_PART, "chart-row parser state ended unexpectedly")
                    })?;
                    current.rows.push(row.cells);
                } else if local == "chart"
                    && chart.as_ref().is_some_and(|chart| chart.depth == depth)
                {
                    let current = chart.take().ok_or_else(|| {
                        format_error(CONTENT_PART, "chart parser state ended unexpectedly")
                    })?;
                    if let Some(parsed) = finish_odp_chart_state(current, CONTENT_PART) {
                        if let Some(frame) = frame.as_mut() {
                            frame.chart = Some(parsed);
                            frame.placeholder = None;
                        }
                    } else if let Some(frame) = frame.as_mut() {
                        frame.placeholder = Some("Chart");
                        if !unsupported_chart_reported {
                            diagnostics.push(
                                Diagnostic::warning(
                                    DiagnosticCode::UnsupportedFeature,
                                    Phase::Render,
                                    Fidelity::Approximate,
                                    "ODP chart content has no usable cached series and uses a source-mapped static placeholder",
                                )
                                .in_part(CONTENT_PART),
                            );
                            unsupported_chart_reported = true;
                        }
                    }
                } else if local == "span" {
                    if let Some(current) = table
                        .as_mut()
                        .and_then(|table| table.cell.as_mut())
                        .filter(|cell| cell.span_depth == Some(depth))
                    {
                        current.span_depth = None;
                        current.span_style_name = None;
                    } else if let Some(current) = shape
                        .as_mut()
                        .filter(|shape| shape.span_depth == Some(depth))
                    {
                        current.span_depth = None;
                        current.span_style_name = None;
                    } else if let Some(current) = frame
                        .as_mut()
                        .filter(|frame| frame.span_depth == Some(depth))
                    {
                        current.span_depth = None;
                        current.span_style_name = None;
                    }
                } else if matches!(local, "table-cell" | "covered-table-cell")
                    && table
                        .as_ref()
                        .and_then(|table| table.cell.as_ref())
                        .is_some_and(|cell| cell.depth == depth)
                {
                    let current = table.as_mut().ok_or_else(|| {
                        format_error(CONTENT_PART, "table parser state ended unexpectedly")
                    })?;
                    let cell = current.cell.take().ok_or_else(|| {
                        format_error(CONTENT_PART, "table-cell parser state ended unexpectedly")
                    })?;
                    current
                        .row
                        .as_mut()
                        .ok_or_else(|| format_error(CONTENT_PART, "table cell has no row"))?
                        .cells
                        .push(cell);
                } else if local == "table-row"
                    && table
                        .as_ref()
                        .and_then(|table| table.row.as_ref())
                        .is_some_and(|row| row.depth == depth)
                {
                    let current = table.as_mut().ok_or_else(|| {
                        format_error(CONTENT_PART, "table parser state ended unexpectedly")
                    })?;
                    let row = current.row.take().ok_or_else(|| {
                        format_error(CONTENT_PART, "table-row parser state ended unexpectedly")
                    })?;
                    current.rows.push(row);
                } else if local == "table"
                    && table.as_ref().is_some_and(|table| table.depth == depth)
                {
                    let current = table.take().ok_or_else(|| {
                        format_error(CONTENT_PART, "table parser state ended unexpectedly")
                    })?;
                    let parent = frame.as_mut().ok_or_else(|| {
                        format_error(CONTENT_PART, "table has no parent frame")
                    })?;
                    if parent.table.replace(current).is_some() {
                        return Err(format_error(
                            CONTENT_PART,
                            "frame contains multiple table representations",
                        ));
                    }
                } else if local == "text-box" {
                    if let Some(current) = frame.as_mut()
                        && current.text_box_depth == Some(depth)
                    {
                        current.text_box_depth = None;
                    }
                } else if local == "frame"
                    && frame.as_ref().is_some_and(|current| current.depth == depth)
                {
                    let mut current = frame.take().ok_or_else(|| {
                        format_error(CONTENT_PART, "frame parser state ended unexpectedly")
                    })?;
                    let unit_index = page
                        .as_ref()
                        .map(|page| page.unit_index)
                        .ok_or_else(|| format_error(CONTENT_PART, "frame has no parent slide"))?;
                    let has_supported_native_table =
                        current.table.is_some() && !current.has_text_box;
                    let prefers_cached_table_image = has_supported_native_table
                        && current.image_element_count != 0
                        && current
                            .table
                            .as_ref()
                            .and_then(|table| table.template_name.as_deref())
                            == Some(EMPTY_TABLE_TEMPLATE);
                    let image_object_count = objects.len();
                    if current.image_element_count != 0
                        && current.chart.is_none()
                        && (!has_supported_native_table || prefers_cached_table_image)
                    {
                        push_image_frame(
                            &current,
                            unit_index,
                            package,
                            styles,
                            &mut objects,
                            &mut diagnostics,
                            &mut image_resources,
                        )?;
                    }
                    let has_media = current.media.is_some();
                    if has_media {
                        attach_media_frame(
                            &current,
                            unit_index,
                            package,
                            &mut objects,
                            &mut diagnostics,
                            &mut image_resources,
                            image_object_count,
                        )?;
                    }
                    let rendered_cached_image = objects.len() > image_object_count;
                    let rendered_cached_table_image =
                        prefers_cached_table_image && rendered_cached_image;
                    if has_media {
                        // The media object, including any draw:image poster, was emitted above.
                    } else if current.has_text_box {
                        push_text_box(
                            current,
                            unit_index,
                            &mut objects,
                            &mut diagnostics,
                            styles,
                            package.limits().max_document_objects,
                            font_metrics,
                            package,
                            &mut image_resources,
                        )?;
                    } else if let Some(table) = current.table.take() {
                        push_odp_table(
                            table,
                            unit_index,
                            &mut objects,
                            styles,
                            font_metrics,
                            rendered_cached_table_image,
                            package.limits().max_document_objects,
                        )?;
                    } else if rendered_cached_image {
                        // The cached image was already emitted as the fallback representation.
                    } else if let Some(chart) = current.chart.as_ref() {
                        push_odp_chart(
                            &ChartFrame::from(&current),
                            chart,
                            unit_index,
                            &mut objects,
                            package.limits().max_document_objects,
                        )?;
                    } else if let Some(math) = current.math.as_ref() {
                        push_odp_math(
                            &current,
                            math,
                            unit_index,
                            &mut objects,
                            package.limits().max_document_objects,
                        )?;
                    } else if let Some(label) = current.placeholder {
                        push_odp_placeholder(
                            &current,
                            label,
                            unit_index,
                            &mut objects,
                            package.limits().max_document_objects,
                        )?;
                    }
                } else if matches!(
                    local,
                    "rect"
                        | "ellipse"
                        | "line"
                        | "custom-shape"
                        | "caption"
                        | "path"
                        | "polygon"
                        | "polyline"
                        | "connector"
                ) && shape.as_ref().is_some_and(|shape| shape.depth == depth)
                {
                    let current = shape.take().ok_or_else(|| {
                        format_error(CONTENT_PART, "shape parser state ended unexpectedly")
                    })?;
                    let unit_index = page
                        .as_ref()
                        .map(|page| page.unit_index)
                        .ok_or_else(|| format_error(CONTENT_PART, "shape has no parent slide"))?;
                    push_draw_shape(
                        current,
                        unit_index,
                        &mut objects,
                        &mut diagnostics,
                        styles,
                        &shape_anchors,
                        package.limits().max_document_objects,
                        &mut image_resources.materialized_bytes,
                        package.limits().max_total_uncompressed_bytes,
                    )?;
                } else if local == "g"
                    && groups.last().is_some_and(|group| group.depth == depth)
                {
                    groups.pop();
                } else if local == "page"
                    && page.as_ref().is_some_and(|current| current.depth == depth)
                {
                    page = None;
                }
            }
            XmlEvent::Text(text) => {
                if notes_depth.is_some() {
                    return Ok(());
                }
                if let Some(encoded) = frame
                    .as_mut()
                    .and_then(|current| current.active_image.as_mut())
                    .filter(|image| image.binary_data_depth.is_some())
                    .and_then(|image| image.representation.inline_base64.as_mut())
                {
                    let text =
                        decode_xml_text(text).map_err(|error| with_part(error, CONTENT_PART))?;
                    append_inline_base64(encoded, &text)?;
                } else if let Some(current) = chart.as_mut() {
                    if let Some(text) = decode_odf_text_node(text, CONTENT_PART)? {
                        current.push_text(&text);
                    }
                } else if let Some(current) = table
                    .as_mut()
                    .and_then(|table| table.cell.as_mut())
                    .filter(|cell| cell.paragraph_depth.is_some())
                {
                    let Some(text) = decode_odf_text_node(text, CONTENT_PART)? else {
                        return Ok(());
                    };
                    let style_name = current_odf_text_style(
                        current.span_depth,
                        &current.span_style_name,
                        &current.paragraph_style_names,
                    );
                    append_odf_text_with_list_prefix(
                        &text,
                        &mut current.text,
                        &mut current.text_runs,
                        &mut current.list_text,
                        style_name.as_deref(),
                    );
                } else if let Some(current) = shape
                    .as_mut()
                    .filter(|shape| shape.paragraph_depth.is_some())
                {
                    let Some(text) = decode_odf_text_node(text, CONTENT_PART)? else {
                        return Ok(());
                    };
                    let style_name = current_odf_text_style(
                        current.span_depth,
                        &current.span_style_name,
                        &current.paragraph_style_names,
                    );
                    append_odf_text_with_list_prefix(
                        &text,
                        &mut current.text,
                        &mut current.text_runs,
                        &mut current.list_text,
                        style_name.as_deref(),
                    );
                } else if let Some(current) = frame
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    let Some(text) = decode_odf_text_node(text, CONTENT_PART)? else {
                        return Ok(());
                    };
                    let style_name = current_odf_text_style(
                        current.span_depth,
                        &current.span_style_name,
                        &current.paragraph_style_names,
                    );
                    append_odf_text_with_list_prefix(
                        &text,
                        &mut current.text,
                        &mut current.text_runs,
                        &mut current.list_text,
                        style_name.as_deref(),
                    );
                }
            }
            XmlEvent::Cdata(text) => {
                if notes_depth.is_some() {
                    return Ok(());
                }
                if let Some(encoded) = frame
                    .as_mut()
                    .and_then(|current| current.active_image.as_mut())
                    .filter(|image| image.binary_data_depth.is_some())
                    .and_then(|image| image.representation.inline_base64.as_mut())
                {
                    append_inline_base64(encoded, text)?;
                } else if let Some(current) = chart.as_mut() {
                    current.push_text(text);
                } else if let Some(current) = table
                    .as_mut()
                    .and_then(|table| table.cell.as_mut())
                    .filter(|cell| cell.paragraph_depth.is_some())
                {
                    let style_name = current_odf_text_style(
                        current.span_depth,
                        &current.span_style_name,
                        &current.paragraph_style_names,
                    );
                    append_odf_text_with_list_prefix(
                        text,
                        &mut current.text,
                        &mut current.text_runs,
                        &mut current.list_text,
                        style_name.as_deref(),
                    );
                } else if let Some(current) = shape
                    .as_mut()
                    .filter(|shape| shape.paragraph_depth.is_some())
                {
                    let style_name = current_odf_text_style(
                        current.span_depth,
                        &current.span_style_name,
                        &current.paragraph_style_names,
                    );
                    append_odf_text_with_list_prefix(
                        text,
                        &mut current.text,
                        &mut current.text_runs,
                        &mut current.list_text,
                        style_name.as_deref(),
                    );
                } else if let Some(current) = frame
                    .as_mut()
                    .filter(|current| current.paragraph_depth.is_some())
                {
                    let style_name = current_odf_text_style(
                        current.span_depth,
                        &current.span_style_name,
                        &current.paragraph_style_names,
                    );
                    append_odf_text_with_list_prefix(
                        text,
                        &mut current.text,
                        &mut current.text_runs,
                        &mut current.list_text,
                        style_name.as_deref(),
                    );
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, CONTENT_PART))?;

    Ok(ParsedContent {
        units,
        objects,
        diagnostics,
    })
}

#[derive(Clone, Copy, Debug)]
struct OdpCellRange {
    start_row: usize,
    start_column: usize,
    end_row: usize,
    end_column: usize,
}

#[derive(Clone, Debug, Default)]
struct OdpChartStyle {
    color: Option<u32>,
    border_visible: Option<bool>,
    visible: Option<bool>,
    stacked: Option<bool>,
    axis_at_end: Option<bool>,
    opacity_gradient: Option<String>,
    chart_props: crate::format::odf_chart::OdfChartStyleProps,
}

fn parse_odp_chart_styles(
    bytes: &[u8],
    package: &dyn OdpSource,
    part: &str,
) -> Result<HashMap<String, OdpChartStyle>, Diagnostic> {
    #[derive(Debug)]
    struct PendingChartStyle {
        depth: usize,
        name: String,
        properties: OdpChartStyle,
    }

    let mut depth = 0_usize;
    let mut current: Option<PendingChartStyle> = None;
    let mut styles = HashMap::new();
    parse_xml(bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                match local_name(name) {
                    "style"
                        if current.is_none()
                            && optional_attribute(&attributes, "family", part)?.as_deref()
                                == Some("chart") =>
                    {
                        current = Some(PendingChartStyle {
                            depth,
                            name: required_attribute(&attributes, "name", part)?,
                            properties: OdpChartStyle::default(),
                        });
                    }
                    "graphic-properties" => {
                        if let Some(style) = current.as_mut() {
                            style.properties.border_visible =
                                optional_attribute(&attributes, "stroke", part)?
                                    .map(|stroke| stroke != "none")
                                    .or(style.properties.border_visible);
                            if optional_attribute(&attributes, "fill", part)?.as_deref()
                                == Some("none")
                            {
                                style.properties.color = None;
                            } else if let Some(color) =
                                optional_attribute(&attributes, "fill-color", part)?
                                    .and_then(|color| parse_odf_color(&color))
                            {
                                style.properties.color = Some(color);
                            }
                            style.properties.opacity_gradient =
                                optional_attribute(&attributes, "opacity-name", part)?
                                    .or(style.properties.opacity_gradient.take());
                        }
                    }
                    "chart-properties" => {
                        if let Some(style) = current.as_mut() {
                            style.properties.visible =
                                optional_attribute(&attributes, "visible", part)?
                                    .map(|visible| visible != "false");
                            style.properties.stacked =
                                optional_attribute(&attributes, "stacked", part)?
                                    .map(|stacked| matches!(stacked.as_str(), "true" | "1"));
                            style.properties.axis_at_end =
                                optional_attribute(&attributes, "axis-position", part)?
                                    .map(|position| position == "end");
                            if let Some(number) =
                                optional_attribute(&attributes, "data-label-number", part)?
                            {
                                style.properties.chart_props.data_label_number =
                                    crate::format::odf_chart::parse_data_label_number(&number);
                            }
                            if let Some(text) =
                                optional_attribute(&attributes, "data-label-text", part)?
                            {
                                style.properties.chart_props.data_label_text =
                                    Some(matches!(text.as_str(), "true" | "1"));
                            }
                            if let Some(symbol) =
                                optional_attribute(&attributes, "data-label-symbol", part)?
                            {
                                style.properties.chart_props.data_label_symbol =
                                    Some(matches!(symbol.as_str(), "true" | "1"));
                            }
                            if let Some(position) =
                                optional_attribute(&attributes, "label-position", part)?
                            {
                                style.properties.chart_props.label_position =
                                    crate::format::odf_chart::parse_label_position(&position);
                            }
                            if let Some(offset) =
                                optional_attribute(&attributes, "pie-offset", part)?
                            {
                                style.properties.chart_props.pie_offset =
                                    crate::format::odf_chart::parse_percent_fraction(&offset);
                            }
                            if optional_attribute(&attributes, "solid-type", part)?.as_deref()
                                == Some("cuboid")
                            {
                                style.properties.chart_props.solid_type_cuboid = true;
                            }
                        }
                    }
                    _ => {}
                }
                if empty && local_name(name) == "style" {
                    current = None;
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if local_name(name) == "style"
                    && current.as_ref().is_some_and(|style| style.depth == depth)
                    && let Some(style) = current.take()
                {
                    styles.insert(style.name, style.properties);
                }
            }
            _ => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(styles)
}

fn parse_odf_opacity_gradient(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<(String, GradientDefinition), Diagnostic> {
    let alpha = |value: Option<String>, default| {
        let transparency = value
            .as_deref()
            .and_then(parse_percentage)
            .unwrap_or(default)
            .clamp(0.0, 1.0);
        ((255.0 * (1.0 - transparency)).round() as u32).min(255)
    };
    let gradient_style = optional_attribute(attributes, "style", part)?;
    let radial = gradient_style
        .as_deref()
        .is_some_and(|style| matches!(style, "radial" | "ellipsoid" | "square" | "rectangular"));
    Ok((
        required_attribute(attributes, "name", part)?,
        GradientDefinition {
            start: alpha(optional_attribute(attributes, "start", part)?, 0.0),
            end: alpha(optional_attribute(attributes, "end", part)?, 1.0),
            angle_degrees: optional_attribute(attributes, "angle", part)?
                .and_then(|value| {
                    value.strip_suffix("deg").map_or_else(
                        || value.parse::<f32>().ok().map(|angle| angle / 10.0),
                        |angle| angle.parse::<f32>().ok(),
                    )
                })
                .unwrap_or(0.0),
            radial,
            rectangular: gradient_style.as_deref() == Some("rectangular"),
            center_x: optional_attribute(attributes, "cx", part)?
                .and_then(|value| parse_percentage(&value))
                .unwrap_or(0.5),
            center_y: optional_attribute(attributes, "cy", part)?
                .and_then(|value| parse_percentage(&value))
                .unwrap_or(0.5),
        },
    ))
}

fn parse_odf_opacity_gradients(
    bytes: &[u8],
    package: &dyn OdpSource,
    part: &str,
    gradients: &mut HashMap<String, GradientDefinition>,
) -> Result<(), Diagnostic> {
    parse_xml(bytes, package.limits(), |event| {
        if let XmlEvent::StartElement {
            name, attributes, ..
        } = event
            && local_name(name) == "opacity"
        {
            let (name, gradient) = parse_odf_opacity_gradient(&attributes, part)?;
            gradients.insert(name, gradient);
        }
        Ok(())
    })
    .map(|_| ())
    .map_err(|error| with_part(error, part))
}

pub(super) fn parse_odp_chart_part(
    package: &Package<'_>,
    part: &str,
) -> Result<Option<OdpBasicChart>, Diagnostic> {
    parse_odp_chart_part_from_source(package, part)
}

#[cfg(feature = "legacy-office-formats")]
pub(super) fn render_embedded_odf_chart(
    bytes: &[u8],
    bounds: Rect,
    limits: Limits,
) -> Result<Option<Vec<Object>>, Diagnostic> {
    let package = Package::open(bytes, limits)?;
    let Some(chart) = parse_odp_chart_part(&package, CONTENT_PART)? else {
        return Ok(None);
    };
    let frame = ChartFrame {
        frame_index: 0,
        parent_numeric_id: None,
        element_id: None,
        bounds,
        transform: AffineTransform::IDENTITY,
    };
    let mut objects = Vec::new();
    push_odp_chart(&frame, &chart, 0, &mut objects, limits.max_document_objects)?;
    Ok(Some(objects))
}

fn parse_odp_chart_part_from_source(
    package: &dyn OdpSource,
    part: &str,
) -> Result<Option<OdpBasicChart>, Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut opacity_gradients = HashMap::new();
    parse_odf_opacity_gradients(&bytes, package, part, &mut opacity_gradients)?;
    if let Some(styles_part) = part
        .strip_suffix("content.xml")
        .map(|root| format!("{root}styles.xml"))
        && let Some(styles_bytes) = package.part(&styles_part)?
    {
        parse_odf_opacity_gradients(&styles_bytes, package, &styles_part, &mut opacity_gradients)?;
    }
    let chart_styles = parse_odp_chart_styles(&bytes, package, part)?;
    let mut depth = 0_usize;
    let mut chart: Option<OdpChartParseState> = None;
    let mut parsed = None;
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                match local {
                    "chart" if chart.is_none() && parsed.is_none() => {
                        let kind =
                            optional_attribute(&attributes, "class", part)?.and_then(|value| {
                                match local_name(&value) {
                                    "bar" | "column" => Some(OdpChartKind::Bar),
                                    "line" => Some(OdpChartKind::Line),
                                    "scatter" => Some(OdpChartKind::Scatter),
                                    "circle" | "pie" => Some(OdpChartKind::Pie),
                                    _ => None,
                                }
                            });
                        if let Some(kind) = kind
                            && !empty
                        {
                            let border_visible =
                                optional_attribute(&attributes, "style-name", part)?
                                    .as_deref()
                                    .and_then(|name| chart_styles.get(name))
                                    .and_then(|style| style.border_visible)
                                    .unwrap_or(true);
                            chart = Some(OdpChartParseState {
                                depth,
                                kind,
                                title: String::new(),
                                title_depth: None,
                                title_paragraph_depth: None,
                                border_visible,
                                categories_range: None,
                                series: Vec::new(),
                                show_legend: false,
                                stacked: false,
                                category_axis_at_end: false,
                                value_axis_at_end: false,
                                chart_size: match (
                                    optional_length(&attributes, "width", part)?,
                                    optional_length(&attributes, "height", part)?,
                                ) {
                                    (Some(width), Some(height)) => Some(PageSize { width, height }),
                                    _ => None,
                                },
                                plot_area: None,
                                value_axis_visible: true,
                                wall: None,
                                rows: Vec::new(),
                                row: None,
                                cell: None,
                            });
                        }
                    }
                    "categories" if chart.is_some() => {
                        if let Some(cell_range) =
                            optional_attribute(&attributes, "cell-range-address", part)?
                            && let Some(current) = chart.as_mut()
                        {
                            current.categories_range = Some(cell_range);
                        }
                    }
                    "plot-area" if chart.is_some() => {
                        if let Some(style_name) =
                            optional_attribute(&attributes, "style-name", part)?
                            && let Some(stacked) = chart_styles
                                .get(&style_name)
                                .and_then(|style| style.stacked)
                        {
                            chart.as_mut().expect("checked above").stacked = stacked;
                        }
                    }
                    "coordinate-region" if chart.is_some() => {
                        if let (Some(x), Some(y), Some(width), Some(height)) = (
                            optional_length(&attributes, "x", part)?,
                            optional_length(&attributes, "y", part)?,
                            optional_length(&attributes, "width", part)?,
                            optional_length(&attributes, "height", part)?,
                        ) {
                            chart.as_mut().expect("checked above").plot_area = Some(Rect {
                                x,
                                y,
                                width,
                                height,
                            });
                        }
                    }
                    "title"
                        if chart
                            .as_ref()
                            .is_some_and(|chart| depth == chart.depth.saturating_add(1)) =>
                    {
                        chart.as_mut().expect("checked above").title_depth =
                            (!empty).then_some(depth);
                    }
                    "legend" if chart.is_some() => {
                        chart.as_mut().expect("checked above").show_legend = true;
                    }
                    "axis" if chart.is_some() => {
                        let dimension = optional_attribute(&attributes, "dimension", part)?;
                        let style = optional_attribute(&attributes, "style-name", part)?
                            .as_deref()
                            .and_then(|name| chart_styles.get(name));
                        let current = chart.as_mut().expect("checked above");
                        if dimension
                            .as_deref()
                            .is_some_and(|dimension| local_name(dimension) == "y")
                        {
                            if style.and_then(|style| style.visible) == Some(false) {
                                current.value_axis_visible = false;
                            }
                            current.value_axis_at_end =
                                style.and_then(|style| style.axis_at_end).unwrap_or(false);
                        } else if dimension
                            .as_deref()
                            .is_some_and(|dimension| local_name(dimension) == "x")
                        {
                            current.category_axis_at_end =
                                style.and_then(|style| style.axis_at_end).unwrap_or(false);
                        }
                    }
                    "wall" if chart.is_some() => {
                        if let Some(style_name) =
                            optional_attribute(&attributes, "style-name", part)?
                            && let Some(style) = chart_styles.get(&style_name)
                            && let Some(color) = style.color
                        {
                            chart.as_mut().expect("checked above").wall = Some(OdpChartWall {
                                color,
                                opacity_gradient: style
                                    .opacity_gradient
                                    .as_deref()
                                    .and_then(|name| opacity_gradients.get(name))
                                    .copied(),
                            });
                        }
                    }
                    "series" if chart.is_some() => {
                        if let Some(values_range) =
                            optional_attribute(&attributes, "values-cell-range-address", part)?
                            && let Some(current) = chart.as_mut()
                        {
                            let style_name = optional_attribute(&attributes, "style-name", part)?;
                            let chart_style = style_name
                                .as_deref()
                                .and_then(|name| chart_styles.get(name));
                            current.series.push(OdpChartSeriesSpec {
                                values_range,
                                label_cell_address: optional_attribute(
                                    &attributes,
                                    "label-cell-address",
                                    part,
                                )?,
                                color: chart_style.and_then(|style| style.color),
                                custom_labels: Vec::new(),
                                domain_range: None,
                                point_explosions: Vec::new(),
                                data_labels: chart_style.map(|style| {
                                    style.chart_props.data_labels(
                                        11.0 * 96.0 / 72.0,
                                        "Arial",
                                        0x0000_00ff,
                                        false,
                                    )
                                }),
                            });
                        }
                    }
                    "domain" if chart.is_some() => {
                        if let Some(range) =
                            optional_attribute(&attributes, "cell-range-address", part)?
                            && let Some(series) =
                                chart.as_mut().and_then(|chart| chart.series.last_mut())
                        {
                            series.domain_range = Some(range);
                        }
                    }
                    "data-point" if chart.is_some() => {
                        let repeat = optional_attribute(&attributes, "repeated", part)?
                            .and_then(|value| value.parse::<usize>().ok())
                            .unwrap_or(1)
                            .max(1);
                        let label = optional_attribute(&attributes, "custom-label-field", part)?
                            .filter(|label| !label.is_empty());
                        let explosion = optional_attribute(&attributes, "style-name", part)?
                            .and_then(|name| chart_styles.get(&name))
                            .and_then(|style| style.chart_props.pie_offset)
                            .unwrap_or(0.0);
                        if let Some(series) =
                            chart.as_mut().and_then(|chart| chart.series.last_mut())
                        {
                            series
                                .custom_labels
                                .extend(std::iter::repeat_n(label, repeat));
                            series
                                .point_explosions
                                .extend(std::iter::repeat_n(explosion, repeat));
                        }
                    }
                    "table-row" if chart.is_some() => {
                        let current = chart.as_mut().ok_or_else(|| {
                            format_error(part, "embedded chart row has no parent chart")
                        })?;
                        if current.row.is_some() {
                            return Err(format_error(
                                part,
                                "nested embedded chart data rows are invalid",
                            ));
                        }
                        current.row = Some(OdpChartRow {
                            depth,
                            cells: Vec::new(),
                        });
                    }
                    "table-cell" | "covered-table-cell"
                        if chart.as_ref().is_some_and(|chart| chart.row.is_some()) =>
                    {
                        let current = chart.as_mut().ok_or_else(|| {
                            format_error(part, "embedded chart cell has no parent chart")
                        })?;
                        if current.cell.is_some() {
                            return Err(format_error(
                                part,
                                "nested embedded chart data cells are invalid",
                            ));
                        }
                        let cell = OdpChartCell {
                            value: optional_attribute(&attributes, "value", part)?
                                .and_then(|value| value.parse::<f32>().ok())
                                .filter(|value| value.is_finite()),
                            text: optional_attribute(&attributes, "string-value", part)?
                                .unwrap_or_default(),
                            paragraph_depth: None,
                        };
                        let repeat =
                            optional_attribute(&attributes, "number-columns-repeated", part)?
                                .and_then(|value| value.parse::<usize>().ok())
                                .unwrap_or(1)
                                .max(1);
                        if empty {
                            current
                                .row
                                .as_mut()
                                .ok_or_else(|| {
                                    format_error(part, "embedded chart cell has no row")
                                })?
                                .cells
                                .extend(std::iter::repeat_n(cell, repeat));
                        } else {
                            current.cell = Some((depth, cell));
                        }
                    }
                    "p" => {
                        if let Some(current) = chart.as_mut() {
                            current.start_paragraph(depth, empty);
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
                if local == "p" {
                    if let Some(current) = chart.as_mut() {
                        current.end_paragraph(depth);
                    }
                } else if local == "title"
                    && chart
                        .as_ref()
                        .is_some_and(|chart| chart.title_depth == Some(depth))
                {
                    chart.as_mut().expect("checked above").title_depth = None;
                } else if matches!(local, "table-cell" | "covered-table-cell")
                    && chart
                        .as_ref()
                        .and_then(|chart| chart.cell.as_ref())
                        .is_some_and(|(start, _)| *start == depth)
                {
                    let current = chart.as_mut().ok_or_else(|| {
                        format_error(part, "embedded chart parser state ended unexpectedly")
                    })?;
                    let (_, cell) = current.cell.take().ok_or_else(|| {
                        format_error(part, "embedded chart cell ended unexpectedly")
                    })?;
                    current
                        .row
                        .as_mut()
                        .ok_or_else(|| format_error(part, "embedded chart cell has no row"))?
                        .cells
                        .push(cell);
                } else if local == "table-row"
                    && chart
                        .as_ref()
                        .and_then(|chart| chart.row.as_ref())
                        .is_some_and(|row| row.depth == depth)
                {
                    let current = chart.as_mut().ok_or_else(|| {
                        format_error(part, "embedded chart parser state ended unexpectedly")
                    })?;
                    let row = current.row.take().ok_or_else(|| {
                        format_error(part, "embedded chart row ended unexpectedly")
                    })?;
                    current.rows.push(row.cells);
                } else if local == "chart"
                    && chart.as_ref().is_some_and(|chart| chart.depth == depth)
                {
                    let current = chart.take().ok_or_else(|| {
                        format_error(part, "embedded chart parser state ended unexpectedly")
                    })?;
                    parsed = finish_odp_chart_state(current, part);
                }
            }
            XmlEvent::Text(text) => {
                if let Some(current) = chart.as_mut()
                    && let Some(text) = decode_odf_text_node(text, part)?
                {
                    current.push_text(&text);
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(current) = chart.as_mut() {
                    current.push_text(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(parsed)
}

fn finish_odp_chart_state(state: OdpChartParseState, source_part: &str) -> Option<OdpBasicChart> {
    let plot_area = state
        .plot_area
        .zip(state.chart_size)
        .and_then(|(plot, size)| {
            (size.width > 0.0 && size.height > 0.0).then_some(Rect {
                x: plot.x / size.width,
                y: plot.y / size.height,
                width: plot.width / size.width,
                height: plot.height / size.height,
            })
        });
    let categories = state
        .categories_range
        .as_deref()
        .and_then(parse_odf_cell_range)
        .map(|range| {
            odp_chart_cells(&state.rows, range)
                .into_iter()
                .map(|cell| odp_chart_cell_text(cell).unwrap_or_default())
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let mut series = Vec::new();
    for spec in state.series {
        let Some(range) = parse_odf_cell_range(&spec.values_range) else {
            continue;
        };
        let values = odp_chart_cells(&state.rows, range)
            .into_iter()
            .filter_map(|cell| cell.value.or_else(|| cell.text.trim().parse::<f32>().ok()))
            .filter(|value| value.is_finite())
            .collect::<Vec<_>>();
        if values.is_empty() {
            continue;
        }
        let label = spec
            .label_cell_address
            .as_deref()
            .and_then(parse_odf_cell_range)
            .and_then(|range| odp_chart_cells(&state.rows, range).first().copied())
            .and_then(odp_chart_cell_text);
        let domains = spec
            .domain_range
            .as_deref()
            .and_then(parse_odf_cell_range)
            .map(|range| {
                odp_chart_cells(&state.rows, range)
                    .into_iter()
                    .filter_map(|cell| cell.value.or_else(|| cell.text.trim().parse::<f32>().ok()))
                    .filter(|value| value.is_finite())
                    .collect()
            })
            .unwrap_or_default();
        series.push(OdpChartSeries {
            label,
            color: spec.color,
            values,
            domains,
            custom_labels: spec.custom_labels,
            point_explosions: spec.point_explosions,
            data_labels: spec.data_labels,
        });
    }
    (!series.is_empty()).then_some(OdpBasicChart {
        kind: state.kind,
        title: (!state.title.trim().is_empty()).then(|| state.title.trim().to_owned()),
        border_visible: state.border_visible,
        categories,
        series,
        show_legend: state.show_legend,
        stacked: state.stacked,
        category_axis_at_end: state.category_axis_at_end,
        value_axis_at_end: state.value_axis_at_end,
        plot_area,
        value_axis_visible: state.value_axis_visible,
        wall: state.wall,
        source_part: source_part.to_owned(),
    })
}

fn odp_chart_cell_text(cell: &OdpChartCell) -> Option<String> {
    let text = cell.text.trim();
    if !text.is_empty() {
        return Some(text.to_owned());
    }
    cell.value.map(|value| value.to_string())
}

fn odp_chart_cells(rows: &[Vec<OdpChartCell>], range: OdpCellRange) -> Vec<&OdpChartCell> {
    let mut cells = Vec::new();
    for row_index in range.start_row..=range.end_row {
        let Some(row) = rows.get(row_index) else {
            continue;
        };
        for column_index in range.start_column..=range.end_column {
            if let Some(cell) = row.get(column_index) {
                cells.push(cell);
            }
        }
    }
    cells
}

fn parse_odf_cell_range(value: &str) -> Option<OdpCellRange> {
    let (start, end) = value.split_once(':').unwrap_or((value, value));
    let (start_row, start_column) = parse_odf_cell_address(start)?;
    let (end_row, end_column) = parse_odf_cell_address(end)?;
    Some(OdpCellRange {
        start_row: start_row.min(end_row),
        start_column: start_column.min(end_column),
        end_row: start_row.max(end_row),
        end_column: start_column.max(end_column),
    })
}

fn parse_odf_cell_address(value: &str) -> Option<(usize, usize)> {
    let address = value.rsplit_once('.').map_or(value, |(_, address)| address);
    let address = address.replace('$', "");
    let split = address
        .char_indices()
        .find(|(_, character)| character.is_ascii_digit())?
        .0;
    let (column, row) = address.split_at(split);
    let mut column_index = 0_usize;
    for character in column.bytes() {
        if !character.is_ascii_alphabetic() {
            return None;
        }
        column_index = column_index
            .checked_mul(26)?
            .checked_add(usize::from(character.to_ascii_uppercase() - b'A' + 1))?;
    }
    let row_index = row.parse::<usize>().ok()?.checked_sub(1)?;
    Some((row_index, column_index.checked_sub(1)?))
}

fn parse_manifest_media_types(
    package: &dyn OdpSource,
) -> Result<HashMap<String, String>, Diagnostic> {
    let Some(bytes) = package.part(MANIFEST_PART)? else {
        return Ok(HashMap::new());
    };
    let mut media_types = HashMap::new();
    parse_xml(&bytes, package.limits(), |event| {
        let XmlEvent::StartElement {
            name, attributes, ..
        } = event
        else {
            return Ok(());
        };
        if local_name(name) != "file-entry" {
            return Ok(());
        }
        let Some(path) = optional_attribute(&attributes, "full-path", MANIFEST_PART)? else {
            return Ok(());
        };
        let Some(media_type) = optional_attribute(&attributes, "media-type", MANIFEST_PART)?
            .filter(|media_type| !media_type.trim().is_empty())
        else {
            return Ok(());
        };
        if path == "/" || path.ends_with('/') {
            return Ok(());
        }
        let target = match resolve_odf_image_target(&path, package.limits().max_zip_path_bytes)
            .map_err(|message| format_error(MANIFEST_PART, message))?
        {
            OdfImageTarget::Embedded(target) => target,
            OdfImageTarget::External => {
                return Err(format_error(
                    MANIFEST_PART,
                    "manifest file paths must refer to embedded package parts",
                ));
            }
        };
        if media_types.insert(target.clone(), media_type).is_some() {
            return Err(format_error(
                MANIFEST_PART,
                format!("duplicate manifest file entry: {target}"),
            ));
        }
        if media_types.len() > package.limits().max_relationship_edges {
            return Err(relationship_limit_error(MANIFEST_PART));
        }
        Ok(())
    })
    .map_err(|error| with_part(error, MANIFEST_PART))?;
    Ok(media_types)
}

fn tokenize_svg_path(value: &str) -> Option<Vec<String>> {
    let bytes = value.as_bytes();
    let mut tokens = Vec::new();
    let mut index = 0_usize;
    while index < bytes.len() {
        if bytes[index].is_ascii_whitespace() || bytes[index] == b',' {
            index += 1;
            continue;
        }
        if bytes[index].is_ascii_alphabetic() {
            tokens.push(value.get(index..index + 1)?.to_owned());
            index += 1;
            continue;
        }
        let start = index;
        if matches!(bytes[index], b'+' | b'-') {
            index += 1;
        }
        let integer_start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        let mut has_digits = index > integer_start;
        if index < bytes.len() && bytes[index] == b'.' {
            index += 1;
            let fraction_start = index;
            while index < bytes.len() && bytes[index].is_ascii_digit() {
                index += 1;
            }
            has_digits |= index > fraction_start;
        }
        if !has_digits {
            return None;
        }
        if index < bytes.len() && matches!(bytes[index], b'e' | b'E') {
            index += 1;
            if index < bytes.len() && matches!(bytes[index], b'+' | b'-') {
                index += 1;
            }
            let exponent_start = index;
            while index < bytes.len() && bytes[index].is_ascii_digit() {
                index += 1;
            }
            if index == exponent_start {
                return None;
            }
        }
        let token = value.get(start..index)?;
        token
            .parse::<f32>()
            .ok()
            .filter(|value| value.is_finite())?;
        tokens.push(token.to_owned());
    }
    Some(tokens)
}

fn svg_path_command(token: &str) -> Option<char> {
    let mut characters = token.chars();
    let command = characters.next()?;
    (characters.next().is_none() && command.is_ascii_alphabetic()).then_some(command)
}

fn take_svg_path_number(tokens: &[String], index: &mut usize) -> Option<f32> {
    let token = tokens.get(*index)?;
    if svg_path_command(token).is_some() {
        return None;
    }
    *index += 1;
    token.parse::<f32>().ok().filter(|value| value.is_finite())
}

fn parse_svg_view_box(value: &str) -> Option<(f32, f32, f32, f32)> {
    let tokens = tokenize_svg_path(value)?;
    if tokens.len() != 4 || tokens.iter().any(|token| svg_path_command(token).is_some()) {
        return None;
    }
    let values = tokens
        .iter()
        .map(|token| token.parse::<f32>().ok())
        .collect::<Option<Vec<_>>>()?;
    (values[2] > 0.0 && values[3] > 0.0).then_some((values[0], values[1], values[2], values[3]))
}

fn scale_svg_path_commands(
    commands: Vec<PathCommand>,
    view_box: (f32, f32, f32, f32),
    bounds: Rect,
) -> Vec<PathCommand> {
    let (min_x, min_y, view_width, view_height) = view_box;
    let scale_x = bounds.width / view_width;
    let scale_y = bounds.height / view_height;
    let point = |x: f32, y: f32| ((x - min_x) * scale_x, (y - min_y) * scale_y);
    commands
        .into_iter()
        .map(|command| match command {
            PathCommand::MoveTo { x, y } => {
                let (x, y) = point(x, y);
                PathCommand::MoveTo { x, y }
            }
            PathCommand::LineTo { x, y } => {
                let (x, y) = point(x, y);
                PathCommand::LineTo { x, y }
            }
            PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => {
                let (cpx, cpy) = point(cpx, cpy);
                let (x, y) = point(x, y);
                PathCommand::QuadraticCurveTo { cpx, cpy, x, y }
            }
            PathCommand::BezierCurveTo {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                let (cp1x, cp1y) = point(cp1x, cp1y);
                let (cp2x, cp2y) = point(cp2x, cp2y);
                let (x, y) = point(x, y);
                PathCommand::BezierCurveTo {
                    cp1x,
                    cp1y,
                    cp2x,
                    cp2y,
                    x,
                    y,
                }
            }
            PathCommand::ClosePath => PathCommand::ClosePath,
        })
        .collect()
}

fn parse_svg_path_geometry(
    data: &str,
    view_box: (f32, f32, f32, f32),
    bounds: Rect,
    fill_rule: FillRule,
) -> Option<Geometry> {
    let tokens = tokenize_svg_path(data)?;
    let mut index = 0_usize;
    let mut command = None;
    let mut current = (0.0_f32, 0.0_f32);
    let mut subpath_start = current;
    let mut cubic_control = None;
    let mut quadratic_control = None;
    let mut commands = Vec::new();
    while index < tokens.len() {
        if let Some(next) = tokens.get(index).and_then(|token| svg_path_command(token)) {
            command = Some(next);
            index += 1;
        }
        let current_command = command?;
        let relative = current_command.is_ascii_lowercase();
        let absolute_point = |x: f32, y: f32, current: (f32, f32)| {
            if relative {
                (current.0 + x, current.1 + y)
            } else {
                (x, y)
            }
        };
        match current_command.to_ascii_uppercase() {
            'M' => {
                let point = absolute_point(
                    take_svg_path_number(&tokens, &mut index)?,
                    take_svg_path_number(&tokens, &mut index)?,
                    current,
                );
                commands.push(PathCommand::MoveTo {
                    x: point.0,
                    y: point.1,
                });
                current = point;
                subpath_start = point;
                cubic_control = None;
                quadratic_control = None;
                command = Some(if relative { 'l' } else { 'L' });
            }
            'L' => {
                let point = absolute_point(
                    take_svg_path_number(&tokens, &mut index)?,
                    take_svg_path_number(&tokens, &mut index)?,
                    current,
                );
                commands.push(PathCommand::LineTo {
                    x: point.0,
                    y: point.1,
                });
                current = point;
                cubic_control = None;
                quadratic_control = None;
            }
            'H' => {
                let x = take_svg_path_number(&tokens, &mut index)?;
                current.0 = if relative { current.0 + x } else { x };
                commands.push(PathCommand::LineTo {
                    x: current.0,
                    y: current.1,
                });
                cubic_control = None;
                quadratic_control = None;
            }
            'V' => {
                let y = take_svg_path_number(&tokens, &mut index)?;
                current.1 = if relative { current.1 + y } else { y };
                commands.push(PathCommand::LineTo {
                    x: current.0,
                    y: current.1,
                });
                cubic_control = None;
                quadratic_control = None;
            }
            'C' => {
                let cp1 = absolute_point(
                    take_svg_path_number(&tokens, &mut index)?,
                    take_svg_path_number(&tokens, &mut index)?,
                    current,
                );
                let cp2 = absolute_point(
                    take_svg_path_number(&tokens, &mut index)?,
                    take_svg_path_number(&tokens, &mut index)?,
                    current,
                );
                let point = absolute_point(
                    take_svg_path_number(&tokens, &mut index)?,
                    take_svg_path_number(&tokens, &mut index)?,
                    current,
                );
                commands.push(PathCommand::BezierCurveTo {
                    cp1x: cp1.0,
                    cp1y: cp1.1,
                    cp2x: cp2.0,
                    cp2y: cp2.1,
                    x: point.0,
                    y: point.1,
                });
                current = point;
                cubic_control = Some(cp2);
                quadratic_control = None;
            }
            'S' => {
                let cp1 = cubic_control
                    .map(|control: (f32, f32)| {
                        (2.0 * current.0 - control.0, 2.0 * current.1 - control.1)
                    })
                    .unwrap_or(current);
                let cp2 = absolute_point(
                    take_svg_path_number(&tokens, &mut index)?,
                    take_svg_path_number(&tokens, &mut index)?,
                    current,
                );
                let point = absolute_point(
                    take_svg_path_number(&tokens, &mut index)?,
                    take_svg_path_number(&tokens, &mut index)?,
                    current,
                );
                commands.push(PathCommand::BezierCurveTo {
                    cp1x: cp1.0,
                    cp1y: cp1.1,
                    cp2x: cp2.0,
                    cp2y: cp2.1,
                    x: point.0,
                    y: point.1,
                });
                current = point;
                cubic_control = Some(cp2);
                quadratic_control = None;
            }
            'Q' => {
                let control = absolute_point(
                    take_svg_path_number(&tokens, &mut index)?,
                    take_svg_path_number(&tokens, &mut index)?,
                    current,
                );
                let point = absolute_point(
                    take_svg_path_number(&tokens, &mut index)?,
                    take_svg_path_number(&tokens, &mut index)?,
                    current,
                );
                commands.push(PathCommand::QuadraticCurveTo {
                    cpx: control.0,
                    cpy: control.1,
                    x: point.0,
                    y: point.1,
                });
                current = point;
                cubic_control = None;
                quadratic_control = Some(control);
            }
            'T' => {
                let control = quadratic_control
                    .map(|control: (f32, f32)| {
                        (2.0 * current.0 - control.0, 2.0 * current.1 - control.1)
                    })
                    .unwrap_or(current);
                let point = absolute_point(
                    take_svg_path_number(&tokens, &mut index)?,
                    take_svg_path_number(&tokens, &mut index)?,
                    current,
                );
                commands.push(PathCommand::QuadraticCurveTo {
                    cpx: control.0,
                    cpy: control.1,
                    x: point.0,
                    y: point.1,
                });
                current = point;
                cubic_control = None;
                quadratic_control = Some(control);
            }
            'Z' => {
                commands.push(PathCommand::ClosePath);
                current = subpath_start;
                cubic_control = None;
                quadratic_control = None;
                command = None;
            }
            _ => return None,
        }
    }
    (commands.len() >= 2).then(|| Geometry::Path {
        fill_rule,
        commands: scale_svg_path_commands(commands, view_box, bounds),
    })
}

fn parse_svg_points_geometry(
    points: &str,
    view_box: (f32, f32, f32, f32),
    bounds: Rect,
    close: bool,
    fill_rule: FillRule,
) -> Option<Geometry> {
    let tokens = tokenize_svg_path(points)?;
    if tokens.len() < 4
        || !tokens.len().is_multiple_of(2)
        || tokens.iter().any(|token| svg_path_command(token).is_some())
    {
        return None;
    }
    let mut commands = Vec::with_capacity(tokens.len() / 2 + usize::from(close));
    for (index, pair) in tokens.chunks_exact(2).enumerate() {
        let x = pair[0].parse::<f32>().ok()?;
        let y = pair[1].parse::<f32>().ok()?;
        commands.push(if index == 0 {
            PathCommand::MoveTo { x, y }
        } else {
            PathCommand::LineTo { x, y }
        });
    }
    if close {
        commands.push(PathCommand::ClosePath);
    }
    Some(Geometry::Path {
        fill_rule,
        commands: scale_svg_path_commands(commands, view_box, bounds),
    })
}

fn odf_enhanced_geometry_kind(
    attributes: &[XmlAttribute<'_>],
    bounds: Rect,
    part: &str,
) -> Result<Option<Geometry>, Diagnostic> {
    Ok(
        optional_attribute(attributes, "type", part)?.and_then(|kind| match kind.as_str() {
            "ellipse" | "circle" => Some(Geometry::Ellipse),
            "round-rect" | "roundRect" => {
                let radius = bounds.width.min(bounds.height) * 0.12;
                Some(Geometry::RoundedRectangle {
                    radius_x: radius,
                    radius_y: radius,
                })
            }
            _ => None,
        }),
    )
}

fn begin_odf_enhanced_geometry(
    depth: usize,
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<Option<OdfEnhancedGeometryState>, Diagnostic> {
    let Some(path) = optional_attribute(attributes, "enhanced-path", part)? else {
        return Ok(None);
    };
    let Some(view_box) = optional_attribute(attributes, "viewBox", part)?
        .as_deref()
        .and_then(parse_svg_view_box)
    else {
        return Ok(None);
    };
    let mut modifiers = Vec::new();
    if let Some(value) = optional_attribute(attributes, "modifiers", part)? {
        for token in value.split_ascii_whitespace() {
            let Ok(value) = token.parse::<f32>() else {
                return Ok(None);
            };
            if !value.is_finite() {
                return Ok(None);
            }
            modifiers.push(value);
        }
    }
    Ok(Some(OdfEnhancedGeometryState {
        depth,
        view_box,
        path_stretchpoint_x: optional_attribute(attributes, "path-stretchpoint-x", part)?
            .and_then(|value| value.parse::<f32>().ok())
            .filter(|value| value.is_finite()),
        path_stretchpoint_y: optional_attribute(attributes, "path-stretchpoint-y", part)?
            .and_then(|value| value.parse::<f32>().ok())
            .filter(|value| value.is_finite()),
        path,
        text_areas: optional_attribute(attributes, "text-areas", part)?,
        modifiers,
        equations: Vec::new(),
        mirror_horizontal: optional_attribute(attributes, "mirror-horizontal", part)?
            .is_some_and(|value| matches!(value.as_str(), "true" | "1")),
        mirror_vertical: optional_attribute(attributes, "mirror-vertical", part)?
            .is_some_and(|value| matches!(value.as_str(), "true" | "1")),
    }))
}

struct OdfFormulaEvaluator {
    formulas: HashMap<String, String>,
    modifiers: Vec<f32>,
    view_box: (f32, f32, f32, f32),
    values: HashMap<String, f32>,
    visiting: HashSet<String>,
}

impl OdfFormulaEvaluator {
    fn new(state: &OdfEnhancedGeometryState, view_box: (f32, f32, f32, f32)) -> Self {
        Self {
            formulas: state.equations.iter().cloned().collect(),
            modifiers: state.modifiers.clone(),
            view_box,
            values: HashMap::new(),
            visiting: HashSet::new(),
        }
    }

    fn evaluate_reference(&mut self, name: &str) -> Option<f32> {
        if let Some(value) = self.values.get(name) {
            return Some(*value);
        }
        let formula = self.formulas.get(name)?.clone();
        if !self.visiting.insert(name.to_owned()) {
            return None;
        }
        let result = OdfFormulaParser::new(&formula, self).parse();
        self.visiting.remove(name);
        let value = result?;
        if !value.is_finite() {
            return None;
        }
        self.values.insert(name.to_owned(), value);
        Some(value)
    }

    fn evaluate_path_value(&mut self, token: &str) -> Option<f32> {
        if let Some(name) = token.strip_prefix('?') {
            self.evaluate_reference(name)
        } else if let Some(index) = token.strip_prefix('$') {
            self.modifiers.get(index.parse::<usize>().ok()?).copied()
        } else {
            token.parse::<f32>().ok().filter(|value| value.is_finite())
        }
    }

    fn evaluate_builtin(&self, name: &str) -> Option<f32> {
        let (min_x, min_y, width, height) = self.view_box;
        match name {
            "left" => Some(min_x),
            "right" => Some(min_x + width),
            "top" => Some(min_y),
            "bottom" => Some(min_y + height),
            "pi" => Some(std::f32::consts::PI),
            _ => None,
        }
    }

    fn evaluate_modifier(&self, index: usize) -> Option<f32> {
        self.modifiers.get(index).copied()
    }
}

struct OdfFormulaParser<'a, 'b> {
    input: &'a str,
    index: usize,
    evaluator: &'b mut OdfFormulaEvaluator,
}

impl<'a, 'b> OdfFormulaParser<'a, 'b> {
    fn new(input: &'a str, evaluator: &'b mut OdfFormulaEvaluator) -> Self {
        Self {
            input,
            index: 0,
            evaluator,
        }
    }

    fn parse(mut self) -> Option<f32> {
        let value = self.parse_additive()?;
        self.skip_whitespace();
        (self.index == self.input.len() && value.is_finite()).then_some(value)
    }

    fn parse_additive(&mut self) -> Option<f32> {
        let mut value = self.parse_multiplicative()?;
        loop {
            self.skip_whitespace();
            let Some(operator) = self.peek_byte() else {
                break;
            };
            if !matches!(operator, b'+' | b'-') {
                break;
            }
            self.index += 1;
            let operand = self.parse_multiplicative()?;
            value = if operator == b'+' {
                value + operand
            } else {
                value - operand
            };
        }
        value.is_finite().then_some(value)
    }

    fn parse_multiplicative(&mut self) -> Option<f32> {
        let mut value = self.parse_unary()?;
        loop {
            self.skip_whitespace();
            let Some(operator) = self.peek_byte() else {
                break;
            };
            if !matches!(operator, b'*' | b'/') {
                break;
            }
            self.index += 1;
            let operand = self.parse_unary()?;
            value = if operator == b'*' {
                value * operand
            } else {
                if operand.abs() <= f32::EPSILON {
                    return None;
                }
                value / operand
            };
        }
        value.is_finite().then_some(value)
    }

    fn parse_unary(&mut self) -> Option<f32> {
        self.skip_whitespace();
        match self.peek_byte()? {
            b'+' => {
                self.index += 1;
                self.parse_unary()
            }
            b'-' => {
                self.index += 1;
                self.parse_unary().map(|value| -value)
            }
            _ => self.parse_primary(),
        }
    }

    fn parse_primary(&mut self) -> Option<f32> {
        self.skip_whitespace();
        match self.peek_byte()? {
            b'(' => {
                self.index += 1;
                let value = self.parse_additive()?;
                self.skip_whitespace();
                (self.take_byte()? == b')').then_some(value)
            }
            b'?' => {
                self.index += 1;
                let name = self.parse_identifier()?;
                self.evaluator.evaluate_reference(name)
            }
            b'$' => {
                self.index += 1;
                let index = self.parse_unsigned_integer()?;
                self.evaluator.evaluate_modifier(index)
            }
            byte if byte.is_ascii_digit() || byte == b'.' => self.parse_number(),
            byte if byte.is_ascii_alphabetic() => {
                let name = self.parse_identifier()?.to_owned();
                self.skip_whitespace();
                if self.peek_byte() != Some(b'(') {
                    return self.evaluator.evaluate_builtin(&name);
                }
                self.index += 1;
                let arguments = self.parse_arguments()?;
                match (name.as_str(), arguments.as_slice()) {
                    ("min", [left, right]) => Some(left.min(*right)),
                    ("max", [left, right]) => Some(left.max(*right)),
                    ("if", [condition, positive, non_positive]) => Some(if *condition > 0.0 {
                        *positive
                    } else {
                        *non_positive
                    }),
                    ("abs", [value]) => Some(value.abs()),
                    ("sqrt", [value]) if *value >= 0.0 => Some(value.sqrt()),
                    ("sin", [value]) => Some(value.sin()),
                    ("cos", [value]) => Some(value.cos()),
                    ("atan2", [y, x]) => Some(y.atan2(*x)),
                    _ => None,
                }
                .filter(|value| value.is_finite())
            }
            _ => None,
        }
    }

    fn parse_arguments(&mut self) -> Option<Vec<f32>> {
        let mut arguments = Vec::new();
        self.skip_whitespace();
        if self.peek_byte() == Some(b')') {
            self.index += 1;
            return Some(arguments);
        }
        loop {
            arguments.push(self.parse_additive()?);
            self.skip_whitespace();
            match self.take_byte()? {
                b',' => {}
                b')' => return Some(arguments),
                _ => return None,
            }
        }
    }

    fn parse_identifier(&mut self) -> Option<&'a str> {
        let start = self.index;
        while self
            .peek_byte()
            .is_some_and(|byte| byte.is_ascii_alphanumeric() || byte == b'_')
        {
            self.index += 1;
        }
        (self.index > start)
            .then(|| self.input.get(start..self.index))
            .flatten()
    }

    fn parse_unsigned_integer(&mut self) -> Option<usize> {
        let start = self.index;
        while self.peek_byte().is_some_and(|byte| byte.is_ascii_digit()) {
            self.index += 1;
        }
        self.input.get(start..self.index)?.parse::<usize>().ok()
    }

    fn parse_number(&mut self) -> Option<f32> {
        let start = self.index;
        while self.peek_byte().is_some_and(|byte| byte.is_ascii_digit()) {
            self.index += 1;
        }
        if self.peek_byte() == Some(b'.') {
            self.index += 1;
            while self.peek_byte().is_some_and(|byte| byte.is_ascii_digit()) {
                self.index += 1;
            }
        }
        if self
            .peek_byte()
            .is_some_and(|byte| matches!(byte, b'e' | b'E'))
        {
            self.index += 1;
            if self
                .peek_byte()
                .is_some_and(|byte| matches!(byte, b'+' | b'-'))
            {
                self.index += 1;
            }
            let exponent_start = self.index;
            while self.peek_byte().is_some_and(|byte| byte.is_ascii_digit()) {
                self.index += 1;
            }
            if self.index == exponent_start {
                return None;
            }
        }
        let value = self.input.get(start..self.index)?.parse::<f32>().ok()?;
        (self.index > start && value.is_finite()).then_some(value)
    }

    fn skip_whitespace(&mut self) {
        while self
            .peek_byte()
            .is_some_and(|byte| byte.is_ascii_whitespace())
        {
            self.index += 1;
        }
    }

    fn peek_byte(&self) -> Option<u8> {
        self.input.as_bytes().get(self.index).copied()
    }

    fn take_byte(&mut self) -> Option<u8> {
        let byte = self.peek_byte()?;
        self.index += 1;
        Some(byte)
    }
}

fn stretched_odf_enhanced_view_box(
    state: &OdfEnhancedGeometryState,
    bounds: Rect,
) -> (f32, f32, f32, f32) {
    let (min_x, min_y, mut width, mut height) = state.view_box;
    if width <= f32::EPSILON
        || height <= f32::EPSILON
        || bounds.width <= f32::EPSILON
        || bounds.height <= f32::EPSILON
    {
        return state.view_box;
    }
    let source_ratio = width / height;
    let target_ratio = bounds.width / bounds.height;
    let stretches_x = state
        .path_stretchpoint_x
        .is_some_and(|point| point >= min_x && point <= min_x + width);
    let stretches_y = state
        .path_stretchpoint_y
        .is_some_and(|point| point >= min_y && point <= min_y + height);
    if target_ratio > source_ratio && stretches_x {
        width = height * target_ratio;
    } else if target_ratio < source_ratio && stretches_y {
        height = width / target_ratio;
    }
    (min_x, min_y, width, height)
}

fn tokenize_odf_enhanced_path(value: &str) -> Option<Vec<String>> {
    let bytes = value.as_bytes();
    let mut tokens = Vec::new();
    let mut index = 0_usize;
    while index < bytes.len() {
        if bytes[index].is_ascii_whitespace() || bytes[index] == b',' {
            index += 1;
            continue;
        }
        if bytes[index].is_ascii_alphabetic() {
            tokens.push(value.get(index..index + 1)?.to_owned());
            index += 1;
            continue;
        }
        if matches!(bytes[index], b'?' | b'$') {
            let start = index;
            index += 1;
            let name_start = index;
            while index < bytes.len()
                && (bytes[index].is_ascii_alphanumeric() || matches!(bytes[index], b'_' | b'-'))
            {
                index += 1;
            }
            if index == name_start {
                return None;
            }
            tokens.push(value.get(start..index)?.to_owned());
            continue;
        }
        let start = index;
        if matches!(bytes[index], b'+' | b'-') {
            index += 1;
        }
        let integer_start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        let mut has_digits = index > integer_start;
        if index < bytes.len() && bytes[index] == b'.' {
            index += 1;
            let fraction_start = index;
            while index < bytes.len() && bytes[index].is_ascii_digit() {
                index += 1;
            }
            has_digits |= index > fraction_start;
        }
        if !has_digits {
            return None;
        }
        if index < bytes.len() && matches!(bytes[index], b'e' | b'E') {
            index += 1;
            if index < bytes.len() && matches!(bytes[index], b'+' | b'-') {
                index += 1;
            }
            let exponent_start = index;
            while index < bytes.len() && bytes[index].is_ascii_digit() {
                index += 1;
            }
            if index == exponent_start {
                return None;
            }
        }
        let token = value.get(start..index)?;
        token
            .parse::<f32>()
            .ok()
            .filter(|value| value.is_finite())?;
        tokens.push(token.to_owned());
    }
    Some(tokens)
}

fn take_odf_enhanced_path_value(
    tokens: &[String],
    index: &mut usize,
    evaluator: &mut OdfFormulaEvaluator,
) -> Option<f32> {
    let token = tokens.get(*index)?;
    if svg_path_command(token).is_some() {
        return None;
    }
    *index += 1;
    evaluator.evaluate_path_value(token)
}

fn append_odf_angle_ellipse(
    commands: &mut Vec<PathCommand>,
    center_x: f32,
    center_y: f32,
    radius_x: f32,
    radius_y: f32,
    start_degrees: f32,
    end_degrees: f32,
    move_to_start: bool,
) {
    let start = start_degrees.rem_euclid(360.0);
    let end = end_degrees.rem_euclid(360.0);
    // ODF angle-ellipse uses radial angles modulo one turn, not repeated revolutions.
    // Equal endpoints form a full ellipse only for an authored single 360-degree turn.
    if start == end && (end_degrees - start_degrees).abs() != 360.0 {
        return;
    }
    let start = start.to_radians();
    let end = end.to_radians();
    super::append_elliptical_arc(
        commands,
        [
            center_x - radius_x,
            center_y - radius_y,
            center_x + radius_x,
            center_y + radius_y,
        ],
        [
            center_x + start.cos(),
            center_y + start.sin(),
            center_x + end.cos(),
            center_y + end.sin(),
        ],
        true,
        move_to_start,
    );
}

fn parse_odf_enhanced_path_geometry(
    state: &OdfEnhancedGeometryState,
    bounds: Rect,
    fill_rule: FillRule,
    transform: &mut AffineTransform,
) -> Option<Geometry> {
    if state.mirror_horizontal != state.mirror_vertical {
        // A single-axis mirror changes the shape coordinate orientation. Conjugate
        // the linear transform around the frame center, preserving its position.
        transform.e += transform.c * bounds.height;
        transform.f += transform.b * bounds.width;
        transform.b = -transform.b;
        transform.c = -transform.c;
    }
    let normalized_path = state
        .path
        .split_ascii_whitespace()
        .collect::<Vec<_>>()
        .join(" ");
    if normalized_path == "M 0 0 L 21600 0 21600 21600 0 21600 Z N" {
        return Some(Geometry::Rectangle);
    }
    if normalized_path
        == "M ?f7 0 X 0 ?f8 L 0 ?f9 Y ?f7 21600 L ?f10 21600 X 21600 ?f9 L 21600 ?f8 Y ?f10 0 Z N"
    {
        let modifier = state.modifiers.first().copied()?.abs();
        let radius = (modifier / state.view_box.2 * bounds.width)
            .min(modifier / state.view_box.3 * bounds.height)
            .clamp(0.0, bounds.width.min(bounds.height) / 2.0);
        return Some(Geometry::RoundedRectangle {
            radius_x: radius,
            radius_y: radius,
        });
    }

    let view_box = stretched_odf_enhanced_view_box(state, bounds);
    let tokens = tokenize_odf_enhanced_path(&state.path)?;
    let mut evaluator = OdfFormulaEvaluator::new(state, view_box);
    let mut index = 0_usize;
    let mut commands = Vec::new();
    let mut layers = Vec::new();
    let mut fill = PathFillMode::Normal;
    let mut stroke = true;
    while index < tokens.len() {
        let command = svg_path_command(tokens.get(index)?)?.to_ascii_uppercase();
        index += 1;
        match command {
            'M' => {
                let mut first = true;
                while index < tokens.len() && svg_path_command(&tokens[index]).is_none() {
                    let x = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    let y = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    commands.push(if first {
                        first = false;
                        PathCommand::MoveTo { x, y }
                    } else {
                        PathCommand::LineTo { x, y }
                    });
                }
                if first {
                    return None;
                }
            }
            'L' => {
                let start = commands.len();
                while index < tokens.len() && svg_path_command(&tokens[index]).is_none() {
                    let x = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    let y = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    commands.push(PathCommand::LineTo { x, y });
                }
                if commands.len() == start {
                    return None;
                }
            }
            'C' => {
                let start = commands.len();
                while index < tokens.len() && svg_path_command(&tokens[index]).is_none() {
                    let cp1x = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    let cp1y = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    let cp2x = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    let cp2y = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    let x = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    let y = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    commands.push(PathCommand::BezierCurveTo {
                        cp1x,
                        cp1y,
                        cp2x,
                        cp2y,
                        x,
                        y,
                    });
                }
                if commands.len() == start {
                    return None;
                }
            }
            'Q' => {
                let start = commands.len();
                while index < tokens.len() && svg_path_command(&tokens[index]).is_none() {
                    let cpx = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    let cpy = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    let x = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    let y = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    commands.push(PathCommand::QuadraticCurveTo { cpx, cpy, x, y });
                }
                if commands.len() == start {
                    return None;
                }
            }
            'A' | 'B' | 'V' | 'W' => {
                let mut consumed = false;
                while index < tokens.len() && svg_path_command(&tokens[index]).is_none() {
                    let mut values = [0.0_f32; 8];
                    for value in &mut values {
                        *value = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    }
                    consumed = true;
                    super::append_elliptical_arc(
                        &mut commands,
                        [values[0], values[1], values[2], values[3]],
                        [values[4], values[5], values[6], values[7]],
                        matches!(command, 'V' | 'W'),
                        matches!(command, 'B' | 'V'),
                    );
                }
                if !consumed {
                    return None;
                }
            }
            'T' | 'U' => {
                let mut consumed = false;
                while index < tokens.len() && svg_path_command(&tokens[index]).is_none() {
                    let mut values = [0.0_f32; 6];
                    for value in &mut values {
                        *value = take_odf_enhanced_path_value(&tokens, &mut index, &mut evaluator)?;
                    }
                    consumed = true;
                    append_odf_angle_ellipse(
                        &mut commands,
                        values[0],
                        values[1],
                        values[2],
                        values[3],
                        values[4],
                        values[5],
                        command == 'U',
                    );
                }
                if !consumed {
                    return None;
                }
            }
            'Z' => commands.push(PathCommand::ClosePath),
            'N' => {
                if !commands.is_empty() {
                    layers.push(PathLayer {
                        fill_rule,
                        fill,
                        stroke,
                        commands: std::mem::take(&mut commands),
                    });
                }
                fill = PathFillMode::Normal;
                stroke = true;
            }
            'F' => fill = PathFillMode::None,
            'S' => stroke = false,
            _ => return None,
        }
    }
    if !commands.is_empty() {
        layers.push(PathLayer {
            fill_rule,
            fill,
            stroke,
            commands,
        });
    }
    if layers.is_empty() {
        return None;
    }
    for layer in &mut layers {
        layer.commands =
            scale_svg_path_commands(std::mem::take(&mut layer.commands), view_box, bounds);
        mirror_odf_path_commands(
            &mut layer.commands,
            bounds,
            state.mirror_horizontal,
            state.mirror_vertical,
        );
    }
    if layers.len() == 1
        && layers[0].fill == PathFillMode::Normal
        && layers[0].stroke
        && layers[0].commands.len() >= 2
    {
        let layer = layers.pop()?;
        Some(Geometry::Path {
            fill_rule: layer.fill_rule,
            commands: layer.commands,
        })
    } else {
        Some(Geometry::LayeredPath { layers })
    }
}

fn parse_odf_enhanced_text_area(state: &OdfEnhancedGeometryState, bounds: Rect) -> Option<Rect> {
    let text_areas = state.text_areas.as_deref()?;
    let view_box = stretched_odf_enhanced_view_box(state, bounds);
    let mut evaluator = OdfFormulaEvaluator::new(state, view_box);
    let mut tokens = text_areas.split_ascii_whitespace();
    let mut values = [0.0_f32; 4];
    for value in &mut values {
        *value = evaluator.evaluate_path_value(tokens.next()?)?;
    }
    if tokens.next().is_some() {
        return None;
    }
    if state.mirror_horizontal {
        values[0] = view_box.0 + view_box.2 - values[0];
        values[2] = view_box.0 + view_box.2 - values[2];
    }
    if state.mirror_vertical {
        values[1] = view_box.1 + view_box.3 - values[1];
        values[3] = view_box.1 + view_box.3 - values[3];
    }
    let left = values[0].min(values[2]);
    let right = values[0].max(values[2]);
    let top = values[1].min(values[3]);
    let bottom = values[1].max(values[3]);
    let area = Rect {
        x: (left - view_box.0) / view_box.2 * bounds.width,
        y: (top - view_box.1) / view_box.3 * bounds.height,
        width: (right - left) / view_box.2 * bounds.width,
        height: (bottom - top) / view_box.3 * bounds.height,
    };
    area.is_valid().then_some(area)
}

fn mirror_odf_path_commands(
    commands: &mut [PathCommand],
    bounds: Rect,
    mirror_horizontal: bool,
    mirror_vertical: bool,
) {
    if !mirror_horizontal && !mirror_vertical {
        return;
    }
    let point = |x: f32, y: f32| {
        (
            if mirror_horizontal {
                bounds.width - x
            } else {
                x
            },
            if mirror_vertical {
                bounds.height - y
            } else {
                y
            },
        )
    };
    for command in commands {
        match command {
            PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => {
                (*x, *y) = point(*x, *y);
            }
            PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => {
                (*cpx, *cpy) = point(*cpx, *cpy);
                (*x, *y) = point(*x, *y);
            }
            PathCommand::BezierCurveTo {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                (*cp1x, *cp1y) = point(*cp1x, *cp1y);
                (*cp2x, *cp2y) = point(*cp2x, *cp2y);
                (*x, *y) = point(*x, *y);
            }
            PathCommand::ClosePath => {}
        }
    }
}

fn create_draw_shape(
    local: &str,
    depth: usize,
    unit_index: u32,
    parent_numeric_id: Option<u32>,
    attributes: &[XmlAttribute<'_>],
    object_index: usize,
) -> Result<DrawShapeState, Diagnostic> {
    let element_id = optional_attribute(attributes, "id", CONTENT_PART)?;
    let connector = parse_odf_connector(local, attributes)?;
    let bounds = if let Some(connector) = connector.as_ref() {
        Rect {
            x: connector.start.0.min(connector.end.0),
            y: connector.start.1.min(connector.end.1),
            width: (connector.end.0 - connector.start.0).abs(),
            height: (connector.end.1 - connector.start.1).abs(),
        }
    } else {
        Rect {
            x: optional_length(attributes, "x", CONTENT_PART)?.unwrap_or(0.0),
            y: optional_length(attributes, "y", CONTENT_PART)?.unwrap_or(0.0),
            width: optional_length(attributes, "width", CONTENT_PART)?.unwrap_or(0.0),
            height: optional_length(attributes, "height", CONTENT_PART)?.unwrap_or(0.0),
        }
    };
    if !bounds.is_valid() {
        return Err(format_error(CONTENT_PART, "shape bounds are invalid"));
    }
    let fill_rule = if optional_attribute(attributes, "fill-rule", CONTENT_PART)?.as_deref()
        == Some("evenodd")
    {
        FillRule::EvenOdd
    } else {
        FillRule::NonZero
    };
    let view_box = optional_attribute(attributes, "viewBox", CONTENT_PART)?
        .as_deref()
        .and_then(parse_svg_view_box);
    let precise_geometry = match local {
        "path" => optional_attribute(attributes, "d", CONTENT_PART)?.and_then(|data| {
            view_box
                .and_then(|view_box| parse_svg_path_geometry(&data, view_box, bounds, fill_rule))
        }),
        "polygon" | "polyline" => {
            optional_attribute(attributes, "points", CONTENT_PART)?.and_then(|points| {
                view_box.and_then(|view_box| {
                    parse_svg_points_geometry(
                        &points,
                        view_box,
                        bounds,
                        local == "polygon",
                        fill_rule,
                    )
                })
            })
        }
        _ => None,
    };
    let named_geometry = if local == "custom-shape" {
        odf_named_shape_geometry(
            optional_attribute(attributes, "name", CONTENT_PART)?.as_deref(),
            bounds,
        )
    } else {
        None
    };
    let geometry_fallback = matches!(local, "custom-shape" | "path" | "polygon" | "polyline")
        && precise_geometry.is_none()
        && named_geometry.is_none();
    let geometry = precise_geometry.or(named_geometry).unwrap_or(match local {
        "ellipse" => Geometry::Ellipse,
        "line" | "connector" => Geometry::Line,
        _ => Geometry::Rectangle,
    });
    Ok(DrawShapeState {
        depth,
        parent_numeric_id,
        element_id,
        path: format!(
            "/office:document-content/office:body/office:presentation/draw:page[{}]/draw:{local}[{}]",
            unit_index + 1,
            object_index + 1
        ),
        bounds,
        geometry,
        connector,
        geometry_fallback,
        enhanced_geometry: None,
        text_path: None,
        text_area: None,
        style_name: optional_attribute(attributes, "style-name", CONTENT_PART)?,
        transform: optional_attribute(attributes, "transform", CONTENT_PART)?
            .map(|value| parse_odf_transform(&value))
            .transpose()?
            .unwrap_or(AffineTransform::IDENTITY),
        paragraph_depth: None,
        paragraph_count: 0,
        text: String::new(),
        text_runs: Vec::new(),
        paragraph_style_names: Vec::new(),
        paragraph_list_layouts: Vec::new(),
        list_text: OdfListTextState::default(),
        span_depth: None,
        span_style_name: None,
    })
}

pub(super) fn parse_odf_connector(
    local: &str,
    attributes: &[XmlAttribute<'_>],
) -> Result<Option<OdfConnector>, Diagnostic> {
    if !matches!(local, "line" | "connector") {
        return Ok(None);
    }
    let x1 = optional_length(attributes, "x1", CONTENT_PART)?.unwrap_or(0.0);
    let y1 = optional_length(attributes, "y1", CONTENT_PART)?.unwrap_or(0.0);
    let x2 = optional_length(attributes, "x2", CONTENT_PART)?.unwrap_or(x1);
    let y2 = optional_length(attributes, "y2", CONTENT_PART)?.unwrap_or(y1);
    let attachment = |shape_name: &str,
                      glue_point_name: &str|
     -> Result<Option<OdfConnectorAttachment>, Diagnostic> {
        let Some(shape_id) = optional_attribute(attributes, shape_name, CONTENT_PART)? else {
            return Ok(None);
        };
        let glue_point = optional_attribute(attributes, glue_point_name, CONTENT_PART)?
            .map(|value| {
                value.parse::<u32>().map_err(|_| {
                    format_error(
                        CONTENT_PART,
                        format!("attribute {glue_point_name} is not a valid glue point"),
                    )
                })
            })
            .transpose()?;
        Ok(Some(OdfConnectorAttachment {
            shape_id,
            glue_point,
        }))
    };
    Ok(Some(OdfConnector {
        kind: if local == "connector"
            && optional_attribute(attributes, "type", CONTENT_PART)?.as_deref() != Some("line")
        {
            OdfConnectorKind::Standard
        } else {
            OdfConnectorKind::Line
        },
        start: (x1, y1),
        end: (x2, y2),
        attachments: [
            attachment("start-shape", "start-glue-point")?,
            attachment("end-shape", "end-glue-point")?,
        ],
    }))
}

pub(super) struct OdfConnectorStyle<'a> {
    pub(super) stroke_width: f32,
    pub(super) dash: Option<&'a [f32]>,
    pub(super) marker_start: bool,
    pub(super) marker_end: bool,
}

pub(super) fn odf_connector_geometry(
    connector: &OdfConnector,
    bounds: Rect,
    transform: AffineTransform,
    shape_anchors: &HashMap<String, OdfShapeAnchor>,
    style: OdfConnectorStyle<'_>,
) -> Geometry {
    let OdfConnectorStyle {
        stroke_width,
        dash,
        marker_start,
        marker_end,
    } = style;
    let transform_point = |point: (f32, f32)| {
        (
            transform.a * point.0 + transform.c * point.1 + transform.e,
            transform.b * point.0 + transform.d * point.1 + transform.f,
        )
    };
    let global_start = transform_point(connector.start);
    let global_end = transform_point(connector.end);
    let mut global_directions = [None, None];
    let mut direction_distances = [f32::INFINITY, f32::INFINITY];
    for attachment in connector.attachments.iter().flatten() {
        let Some((glue, direction)) = odf_connector_attachment(attachment, shape_anchors) else {
            continue;
        };
        let start_distance = (glue.0 - global_start.0).hypot(glue.1 - global_start.1);
        let end_distance = (glue.0 - global_end.0).hypot(glue.1 - global_end.1);
        let (index, distance) = if start_distance <= end_distance {
            (0, start_distance)
        } else {
            (1, end_distance)
        };
        if distance < direction_distances[index] {
            global_directions[index] = Some(direction);
            direction_distances[index] = distance;
        }
    }
    let determinant = transform.a * transform.d - transform.b * transform.c;
    let local_direction = |direction: (f32, f32)| {
        if determinant.abs() <= f32::EPSILON {
            return direction;
        }
        let x = (transform.d * direction.0 - transform.c * direction.1) / determinant;
        let y = (-transform.b * direction.0 + transform.a * direction.1) / determinant;
        if x.abs() >= y.abs() {
            (x.signum(), 0.0)
        } else {
            (0.0, y.signum())
        }
    };
    let directions = global_directions.map(|direction| direction.map(local_direction));
    let start = (connector.start.0 - bounds.x, connector.start.1 - bounds.y);
    let end = (connector.end.0 - bounds.x, connector.end.1 - bounds.y);
    let mut points = match connector.kind {
        OdfConnectorKind::Line => vec![start, end],
        OdfConnectorKind::Standard => {
            odf_standard_connector_points(start, end, directions[0], directions[1])
        }
    };
    points.dedup_by(|left, right| left.0 == right.0 && left.1 == right.1);
    if points.len() < 2 {
        return Geometry::Line;
    }
    let mut commands = Vec::new();
    if let Some(dash) = dash.filter(|dash| !dash.is_empty()) {
        crate::model::append_dashed_polyline(&mut commands, &points, dash, 0.0);
    } else {
        commands.push(PathCommand::MoveTo {
            x: points[0].0,
            y: points[0].1,
        });
        commands.extend(
            points
                .iter()
                .skip(1)
                .map(|&(x, y)| PathCommand::LineTo { x, y }),
        );
    }
    if marker_start {
        commands.push(PathCommand::MoveTo {
            x: points[0].0,
            y: points[0].1,
        });
        append_odf_triangle_marker(&mut commands, points[0], points[1], stroke_width);
    }
    if marker_end {
        let last = points.len() - 1;
        commands.push(PathCommand::MoveTo {
            x: points[last].0,
            y: points[last].1,
        });
        append_odf_triangle_marker(&mut commands, points[last], points[last - 1], stroke_width);
    }
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    }
}

fn odf_connector_attachment(
    attachment: &OdfConnectorAttachment,
    shape_anchors: &HashMap<String, OdfShapeAnchor>,
) -> Option<((f32, f32), (f32, f32))> {
    let anchor = shape_anchors.get(&attachment.shape_id)?;
    let center_x = anchor.bounds.x + anchor.bounds.width / 2.0;
    let center_y = anchor.bounds.y + anchor.bounds.height / 2.0;
    let (point, direction) = match attachment.glue_point? {
        0 | 4 => ((center_x, anchor.bounds.y), (0.0, -1.0)),
        1 => (
            (anchor.bounds.x + anchor.bounds.width, center_y),
            (1.0, 0.0),
        ),
        2 | 6 => (
            (center_x, anchor.bounds.y + anchor.bounds.height),
            (0.0, 1.0),
        ),
        3 | 5 => ((anchor.bounds.x, center_y), (-1.0, 0.0)),
        7 => (
            (anchor.bounds.x + anchor.bounds.width, center_y),
            (1.0, 0.0),
        ),
        _ => return None,
    };
    let point = (
        anchor.transform.a * point.0 + anchor.transform.c * point.1 + anchor.transform.e,
        anchor.transform.b * point.0 + anchor.transform.d * point.1 + anchor.transform.f,
    );
    let direction = (
        anchor.transform.a * direction.0 + anchor.transform.c * direction.1,
        anchor.transform.b * direction.0 + anchor.transform.d * direction.1,
    );
    let direction = if direction.0.abs() >= direction.1.abs() {
        (direction.0.signum(), 0.0)
    } else {
        (0.0, direction.1.signum())
    };
    Some((point, direction))
}

fn odf_standard_connector_points(
    start: (f32, f32),
    end: (f32, f32),
    start_direction: Option<(f32, f32)>,
    end_direction: Option<(f32, f32)>,
) -> Vec<(f32, f32)> {
    const EPSILON: f32 = 0.5;
    let forward = |origin: (f32, f32), point: (f32, f32), direction: (f32, f32)| {
        (point.0 - origin.0) * direction.0 + (point.1 - origin.1) * direction.1 >= -EPSILON
    };
    if (start.0 - end.0).abs() <= EPSILON || (start.1 - end.1).abs() <= EPSILON {
        return vec![start, end];
    }
    if let (Some(start_direction), Some(end_direction)) = (start_direction, end_direction) {
        let dot = start_direction.0 * end_direction.0 + start_direction.1 * end_direction.1;
        if dot.abs() <= EPSILON {
            let corner = if start_direction.0 != 0.0 {
                (end.0, start.1)
            } else {
                (start.0, end.1)
            };
            if forward(start, corner, start_direction) && forward(end, corner, end_direction) {
                return vec![start, corner, end];
            }
        } else if dot <= -1.0 + EPSILON {
            if start_direction.0 != 0.0 {
                let middle_x = (start.0 + end.0) / 2.0;
                let first = (middle_x, start.1);
                let second = (middle_x, end.1);
                if forward(start, first, start_direction) && forward(end, second, end_direction) {
                    return vec![start, first, second, end];
                }
            } else {
                let middle_y = (start.1 + end.1) / 2.0;
                let first = (start.0, middle_y);
                let second = (end.0, middle_y);
                if forward(start, first, start_direction) && forward(end, second, end_direction) {
                    return vec![start, first, second, end];
                }
            }
        } else if start_direction == end_direction {
            const ESCAPE: f32 = 24.0;
            if start_direction.0 > 0.0 {
                let outside = start.0.max(end.0) + ESCAPE;
                return vec![start, (outside, start.1), (outside, end.1), end];
            }
            if start_direction.0 < 0.0 {
                let outside = start.0.min(end.0) - ESCAPE;
                return vec![start, (outside, start.1), (outside, end.1), end];
            }
            if start_direction.1 > 0.0 {
                let outside = start.1.max(end.1) + ESCAPE;
                return vec![start, (start.0, outside), (end.0, outside), end];
            }
            let outside = start.1.min(end.1) - ESCAPE;
            return vec![start, (start.0, outside), (end.0, outside), end];
        }
    }
    if let Some(direction) = start_direction {
        let (first, second) = if direction.0 != 0.0 {
            let middle_x = (start.0 + end.0) / 2.0;
            ((middle_x, start.1), (middle_x, end.1))
        } else {
            let middle_y = (start.1 + end.1) / 2.0;
            ((start.0, middle_y), (end.0, middle_y))
        };
        if forward(start, first, direction) {
            return vec![start, first, second, end];
        }
    }
    if let Some(direction) = end_direction {
        let (first, second) = if direction.0 != 0.0 {
            let middle_x = (start.0 + end.0) / 2.0;
            ((middle_x, start.1), (middle_x, end.1))
        } else {
            let middle_y = (start.1 + end.1) / 2.0;
            ((start.0, middle_y), (end.0, middle_y))
        };
        if forward(end, second, direction) {
            return vec![start, first, second, end];
        }
    }
    let middle_x = (start.0 + end.0) / 2.0;
    vec![start, (middle_x, start.1), (middle_x, end.1), end]
}

fn append_odf_triangle_marker(
    commands: &mut Vec<PathCommand>,
    tip: (f32, f32),
    adjacent: (f32, f32),
    stroke_width: f32,
) {
    let delta_x = adjacent.0 - tip.0;
    let delta_y = adjacent.1 - tip.1;
    let segment_length = delta_x.hypot(delta_y);
    if segment_length <= f32::EPSILON || stroke_width <= f32::EPSILON {
        return;
    }
    let direction_x = delta_x / segment_length;
    let direction_y = delta_y / segment_length;
    let marker_length = (stroke_width * 3.0).min(segment_length);
    let half_width = stroke_width * 1.5;
    let base_x = tip.0 + direction_x * marker_length;
    let base_y = tip.1 + direction_y * marker_length;
    let perpendicular_x = -direction_y * half_width;
    let perpendicular_y = direction_x * half_width;
    commands.extend([
        PathCommand::MoveTo { x: tip.0, y: tip.1 },
        PathCommand::LineTo {
            x: base_x + perpendicular_x,
            y: base_y + perpendicular_y,
        },
        PathCommand::LineTo {
            x: base_x - perpendicular_x,
            y: base_y - perpendicular_y,
        },
        PathCommand::ClosePath,
    ]);
}

fn odf_geometry_with_markers(
    geometry: Geometry,
    marker_start: bool,
    marker_end: bool,
    stroke_width: f32,
) -> Geometry {
    if (!marker_start && !marker_end) || stroke_width <= f32::EPSILON {
        return geometry;
    }
    let Geometry::Path {
        fill_rule,
        mut commands,
    } = geometry
    else {
        return geometry;
    };
    let mut cursor = None;
    let mut first_segment = None;
    let mut last_segment = None;
    for command in &commands {
        let (end, start_adjacent, end_adjacent) = match *command {
            PathCommand::MoveTo { x, y } => {
                cursor = Some((x, y));
                continue;
            }
            PathCommand::LineTo { x, y } => ((x, y), (x, y), cursor),
            PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => {
                ((x, y), (cpx, cpy), Some((cpx, cpy)))
            }
            PathCommand::BezierCurveTo {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => ((x, y), (cp1x, cp1y), Some((cp2x, cp2y))),
            PathCommand::ClosePath => continue,
        };
        if let Some(start) = cursor {
            first_segment.get_or_insert((start, start_adjacent));
            last_segment = Some((end, end_adjacent.unwrap_or(start)));
        }
        cursor = Some(end);
    }
    if marker_start && let Some((tip, adjacent)) = first_segment {
        append_odf_triangle_marker(&mut commands, tip, adjacent, stroke_width);
    }
    if marker_end && let Some((tip, adjacent)) = last_segment {
        append_odf_triangle_marker(&mut commands, tip, adjacent, stroke_width);
    }
    Geometry::Path {
        fill_rule,
        commands,
    }
}

// ODF Fontwork paints glyphs between envelope paths; these paths are not filled shapes.
// Preserve the existing geometry/protocol so page and master text share the renderer.
fn odf_text_path_mode(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<Option<String>, Diagnostic> {
    if optional_attribute(attributes, "text-path", part)?
        .is_some_and(|value| matches!(value.as_str(), "true" | "1"))
    {
        return Ok(Some(
            optional_attribute(attributes, "text-path-mode", part)?.unwrap_or_default(),
        ));
    }
    Ok(None)
}

fn odf_text_path_visual(
    mut visual: Visual,
    text_path: Option<&str>,
    diagnostics: &mut Vec<Diagnostic>,
    part: &str,
) -> Visual {
    if text_path.is_some()
        && let Visual::TextLayout {
            layout,
            visual: inner,
        } = &mut visual
        && let Visual::RichText {
            geometry,
            fill,
            stroke,
            stroke_width,
            ..
        } = inner.as_mut()
    {
        let supported = text_path == Some("shape")
            && matches!(geometry,
            Geometry::LayeredPath { layers } if layers.len() == 2 && layers.iter().all(|layer|
                matches!(layer.commands.as_slice(), [PathCommand::MoveTo { .. }, PathCommand::LineTo { .. } | PathCommand::BezierCurveTo { .. }])));
        layout.warp = supported.then(|| "text-envelope".to_owned());
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ApproximateLayout,
                Phase::Render,
                Fidelity::Approximate,
                if supported {
                    "ODF Fontwork is rasterized between its authored envelope curves"
                } else {
                    "unsupported ODF text path is rendered as ordinary text"
                },
            )
            .in_part(part),
        );
        layout.text_fill = !matches!(fill, Paint::None);
        layout.text_paint = std::mem::replace(fill, Paint::None);
        layout.text_stroke_paint = std::mem::replace(stroke, Paint::None);
        layout.text_stroke_width = *stroke_width;
        *stroke_width = 0.0;
    }
    visual
}

fn push_draw_shape(
    shape: DrawShapeState,
    unit_index: u32,
    objects: &mut Vec<Object>,
    diagnostics: &mut Vec<Diagnostic>,
    styles: &StyleCatalog,
    shape_anchors: &HashMap<String, OdfShapeAnchor>,
    object_limit: usize,
    materialized_image_bytes: &mut usize,
    materialized_image_limit: usize,
) -> Result<(), Diagnostic> {
    if objects.len() >= object_limit {
        return Err(object_limit_error());
    }
    if shape.geometry_fallback {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Approximate,
                "ODP custom/path geometry uses a bounded rectangle fallback",
            )
            .in_part(CONTENT_PART)
            .with_detail("path", shape.path.clone())
            .with_detail(
                "elementId",
                shape.element_id.as_deref().unwrap_or("unknown"),
            ),
        );
    }
    let style = shape
        .style_name
        .as_ref()
        .and_then(|name| styles.graphics.get(name))
        .cloned()
        .unwrap_or_default();
    let fill = resolve_odf_shape_fill(
        &style.fill,
        style.fill_image_tile,
        style.fill_image_width,
        style.fill_image_height,
        shape.bounds,
        styles,
        materialized_image_bytes,
        materialized_image_limit,
    )?;
    let (fill, layer_opacity) = odf_fill_opacity(fill, style.opacity);
    let stroke = resolve_odf_paint(&style.stroke, shape.bounds, styles);
    let resolved_dash = resolved_odf_stroke_dash(&style, styles);
    let materialized_dashes = shape.connector.is_some();
    let geometry = shape.connector.as_ref().map_or_else(
        || shape.geometry.clone(),
        |connector| {
            odf_connector_geometry(
                connector,
                shape.bounds,
                shape.transform,
                shape_anchors,
                OdfConnectorStyle {
                    stroke_width: style.stroke_width,
                    dash: resolved_dash.as_deref(),
                    marker_start: style.marker_start,
                    marker_end: style.marker_end,
                },
            )
        },
    );
    let geometry = if shape.connector.is_none() {
        odf_geometry_with_markers(
            geometry,
            style.marker_start,
            style.marker_end,
            style.stroke_width,
        )
    } else {
        geometry
    };
    let has_text = !shape.text.is_empty();
    let visual = if has_text {
        let runs = odf_text_runs(&shape.text, &shape.text_runs, &style, styles, &[]);
        let line_height = odf_text_line_height(&style, &shape.paragraph_style_names, &runs, styles);
        let mut layout = odf_text_layout(
            &style,
            &shape.paragraph_style_names,
            &shape.paragraph_list_layouts,
            &runs,
            line_height,
            true,
            styles,
        );
        if let Some(text_area) = shape.text_area {
            layout.inset_left = text_area.x.max(0.0);
            layout.inset_top = text_area.y.max(0.0);
            layout.inset_right = (shape.bounds.width - text_area.x - text_area.width).max(0.0);
            layout.inset_bottom = (shape.bounds.height - text_area.y - text_area.height).max(0.0);
        }
        let rich_text = Visual::RichText {
            geometry,
            fill,
            stroke,
            stroke_width: style.stroke_width,
            align: odf_text_align(&style, &shape.paragraph_style_names, styles),
            line_height,
            runs,
        };
        Visual::TextLayout {
            layout,
            visual: Box::new(rich_text),
        }
    } else {
        Visual::PaintedShape {
            geometry,
            fill,
            stroke,
            stroke_width: style.stroke_width,
        }
    };
    let visual = odf_text_path_visual(
        visual,
        shape.text_path.as_deref(),
        diagnostics,
        CONTENT_PART,
    );
    let visual = odf_stroke_style_visual(visual, &style, styles, materialized_dashes);
    let visual = effect_visual(visual, style.shadow);
    let visual = layer_visual(visual, shape.transform, layer_opacity);
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    let mapping = if shape.element_id.is_some() {
        MappingQuality::Exact
    } else {
        MappingQuality::Derived
    };
    let text_length = u32::try_from(shape.text.chars().count()).unwrap_or(u32::MAX);
    objects.push(Object {
        numeric_id,
        parent_numeric_id: shape.parent_numeric_id,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: shape
            .parent_numeric_id
            .map(|parent| format!("object:{parent}")),
        kind: if has_text {
            ObjectKind::TextBox
        } else {
            ObjectKind::Shape
        },
        unit_index,
        bounds: shape.bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: has_text.then_some(shape.text),
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping,
            locator: SourceLocator::OdpElement {
                element_id: shape.element_id,
                path: shape.path,
                row: None,
                column: None,
            },
        },
        visual,
    });
    if has_text && text_length == u32::MAX {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Approximate,
                "ODP shape text range exceeds the public mapping range",
            )
            .in_part(CONTENT_PART),
        );
    }
    Ok(())
}

struct OdpBackgroundSink<'a> {
    objects: &'a mut Vec<Object>,
    object_limit: usize,
    materialized_image_bytes: &'a mut usize,
    materialized_image_limit: usize,
}

fn push_odp_background(
    unit_index: u32,
    size: PageSize,
    style_name: &str,
    styles: &StyleCatalog,
    sink: OdpBackgroundSink<'_>,
) -> Result<(), Diagnostic> {
    let OdpBackgroundSink {
        objects,
        object_limit,
        materialized_image_bytes,
        materialized_image_limit,
    } = sink;
    let Some(style) = styles.graphics.get(style_name) else {
        return Ok(());
    };
    let bounds = Rect {
        x: 0.0,
        y: 0.0,
        width: size.width,
        height: size.height,
    };
    let (kind, visual) = match &style.fill {
        OdfPaint::Bitmap(_) if style.fill_image_tile => {
            let fill = resolve_odf_shape_fill(
                &style.fill,
                true,
                style.fill_image_width,
                style.fill_image_height,
                bounds,
                styles,
                materialized_image_bytes,
                materialized_image_limit,
            )?;
            (
                ObjectKind::Shape,
                Visual::PaintedShape {
                    geometry: Geometry::Rectangle,
                    fill,
                    stroke: Paint::None,
                    stroke_width: 0.0,
                },
            )
        }
        OdfPaint::Bitmap(name) => {
            let Some(image) = styles.fill_images.get(name) else {
                return Ok(());
            };
            reserve_materialized_image_bytes(
                materialized_image_bytes,
                image.bytes.len(),
                materialized_image_limit,
                &image.target,
            )?;
            (
                ObjectKind::Image,
                Visual::Image {
                    media_type: image.media_type.clone(),
                    bytes: clone_image_bytes(&image.bytes, &image.target)?,
                    crop: ImageCrop::default(),
                },
            )
        }
        _ => {
            let fill = resolve_odf_paint(&style.fill, bounds, styles);
            if matches!(fill, Paint::None) {
                return Ok(());
            }
            (
                ObjectKind::Shape,
                Visual::PaintedShape {
                    geometry: Geometry::Rectangle,
                    fill,
                    stroke: Paint::None,
                    stroke_width: 0.0,
                },
            )
        }
    };
    if objects.len() >= object_limit {
        return Err(object_limit_error());
    }
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: None,
        kind,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping: MappingQuality::Derived,
            locator: SourceLocator::OdpElement {
                element_id: None,
                path: format!(
                    "/office:document-content/office:body/office:presentation/draw:page[{}]",
                    unit_index + 1
                ),
                row: None,
                column: None,
            },
        },
        visual: layer_visual(
            odf_stroke_style_visual(visual, style, styles, false),
            AffineTransform::IDENTITY,
            style.opacity,
        ),
    });
    Ok(())
}

fn push_master_template(
    template: &MasterTemplate,
    unit_index: u32,
    objects: &mut Vec<Object>,
    diagnostics: &mut Vec<Diagnostic>,
    styles: &StyleCatalog,
    object_limit: usize,
    materialized_image_bytes: &mut usize,
    materialized_image_limit: usize,
) -> Result<(), Diagnostic> {
    if objects.len() >= object_limit {
        return Err(object_limit_error());
    }
    let style = template
        .style_name
        .as_ref()
        .and_then(|name| styles.graphics.get(name))
        .cloned()
        .unwrap_or_default();
    let has_text = !template.text.is_empty();
    let visual = if let Some(image) = template.image.as_ref() {
        reserve_materialized_image_bytes(
            materialized_image_bytes,
            image.bytes.len(),
            materialized_image_limit,
            &image.target,
        )?;
        Visual::Image {
            media_type: image.media_type.clone(),
            bytes: clone_image_bytes(&image.bytes, &image.target)?,
            crop: normalized_odf_image_crop(style.image_clip, template.bounds),
        }
    } else if has_text {
        let runs = odf_text_runs(&template.text, &template.text_runs, &style, styles, &[]);
        let line_height =
            odf_text_line_height(&style, &template.paragraph_style_names, &runs, styles);
        let layout = odf_text_layout(
            &style,
            &template.paragraph_style_names,
            &[],
            &runs,
            line_height,
            true,
            styles,
        );
        let rich_text = Visual::RichText {
            geometry: template.geometry.clone(),
            fill: resolve_odf_paint(&style.fill, template.bounds, styles),
            stroke: resolve_odf_paint(&style.stroke, template.bounds, styles),
            stroke_width: style.stroke_width,
            align: odf_text_align(&style, &template.paragraph_style_names, styles),
            line_height,
            runs,
        };
        Visual::TextLayout {
            layout,
            visual: Box::new(rich_text),
        }
    } else {
        Visual::PaintedShape {
            geometry: template.geometry.clone(),
            fill: resolve_odf_paint(&style.fill, template.bounds, styles),
            stroke: resolve_odf_paint(&style.stroke, template.bounds, styles),
            stroke_width: style.stroke_width,
        }
    };
    let visual = odf_text_path_visual(
        visual,
        template.text_path.as_deref(),
        diagnostics,
        STYLES_PART,
    );
    let visual = odf_stroke_style_visual(visual, &style, styles, false);
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(STYLES_PART, "object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: None,
        kind: if template.image.is_some() {
            ObjectKind::Image
        } else if has_text {
            ObjectKind::TextBox
        } else {
            ObjectKind::Shape
        },
        unit_index,
        bounds: template.bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: has_text.then_some(template.text.clone()),
        source: SourceRef {
            part: STYLES_PART.to_owned(),
            mapping: if template.element_id.is_some() {
                MappingQuality::Exact
            } else {
                MappingQuality::Derived
            },
            locator: SourceLocator::OdpElement {
                element_id: template.element_id.clone(),
                path: template.path.clone(),
                row: None,
                column: None,
            },
        },
        visual: layer_visual(visual, template.transform, style.opacity),
    });
    Ok(())
}

fn odf_style_overlay(
    base_style: &GraphicStyle,
    style_name: Option<&str>,
    styles: &StyleCatalog,
) -> GraphicStyle {
    let Some(style_name) = style_name else {
        return base_style.clone();
    };
    if let Some(patch) = styles.style_patches.get(style_name) {
        let mut style = base_style.clone();
        patch.apply_to(&mut style);
        style
    } else {
        styles
            .graphics
            .get(style_name)
            .cloned()
            .unwrap_or_else(|| base_style.clone())
    }
}

fn odp_table_template_cell_style_names<'a>(
    table: &OdpTableState,
    row_index: usize,
    column_index: usize,
    row_span: usize,
    column_span: usize,
    row_count: usize,
    column_count: usize,
    template: &'a OdfTableTemplate,
) -> [Option<&'a str>; 3] {
    let edge = if table.use_first_row && row_index == 0 {
        template.first_row.as_deref()
    } else if table.use_last_row && row_index.saturating_add(row_span) >= row_count {
        template.last_row.as_deref()
    } else if table.use_first_column && column_index == 0 {
        template.first_column.as_deref()
    } else if table.use_last_column && column_index.saturating_add(column_span) >= column_count {
        template.last_column.as_deref()
    } else {
        None
    };

    let row_style = if table.use_banding_rows {
        let band_index = row_index.saturating_sub(usize::from(table.use_first_row));
        if band_index.is_multiple_of(2) {
            template.odd_rows.as_deref()
        } else {
            template.even_rows.as_deref()
        }
    } else {
        None
    };
    let column_style = if table.use_banding_columns {
        let band_index = column_index.saturating_sub(usize::from(table.use_first_column));
        if band_index.is_multiple_of(2) {
            template.odd_columns.as_deref()
        } else {
            template.even_columns.as_deref()
        }
    } else {
        None
    };
    [row_style, column_style, edge]
}

fn odp_table_cell_style(
    table: &OdpTableState,
    row_index: usize,
    column_index: usize,
    row_span: usize,
    column_span: usize,
    row_count: usize,
    column_count: usize,
    direct_style_name: Option<&str>,
    styles: &StyleCatalog,
) -> GraphicStyle {
    let template = table
        .template_name
        .as_deref()
        .and_then(|name| styles.table_templates.get(name));
    let mut style = template
        .and_then(|template| template.body.as_deref())
        .and_then(|name| styles.graphics.get(name))
        .cloned()
        .unwrap_or_default();
    if let Some(template) = template {
        for style_name in odp_table_template_cell_style_names(
            table,
            row_index,
            column_index,
            row_span,
            column_span,
            row_count,
            column_count,
            template,
        )
        .into_iter()
        .flatten()
        {
            style = odf_style_overlay(&style, Some(style_name), styles);
        }
    }
    odf_style_overlay(&style, direct_style_name, styles)
}

fn push_odp_table(
    table: OdpTableState,
    unit_index: u32,
    objects: &mut Vec<Object>,
    styles: &StyleCatalog,
    font_metrics: &FontMetricTable,
    semantic_only: bool,
    object_limit: usize,
) -> Result<(), Diagnostic> {
    let cell_count = table.rows.iter().map(|row| row.cells.len()).sum::<usize>();
    if objects
        .len()
        .checked_add(cell_count + 1)
        .is_none_or(|required| required > object_limit)
    {
        return Err(object_limit_error());
    }
    let table_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    let mapping = if table.element_id.is_some() {
        MappingQuality::Exact
    } else {
        MappingQuality::Derived
    };
    objects.push(Object {
        numeric_id: table_id,
        parent_numeric_id: table.parent_numeric_id,
        stable_id: format!("object:{table_id}"),
        parent_stable_id: table
            .parent_numeric_id
            .map(|parent| format!("object:{parent}")),
        kind: ObjectKind::Table,
        unit_index,
        bounds: table.bounds,
        z: i32::try_from(table_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping,
            locator: SourceLocator::OdpElement {
                element_id: table.element_id.clone(),
                path: format!(
                    "/office:document-content/office:body/office:presentation/draw:page[{}]/table:table[1]",
                    unit_index + 1
                ),
                row: None,
                column: None,
            },
        },
        visual: layer_visual(Visual::None, table.transform, 1.0),
    });
    let row_count = table.rows.len().max(1);
    let column_count = table
        .column_widths
        .len()
        .max(
            table
                .rows
                .iter()
                .map(|row| row.cells.len())
                .max()
                .unwrap_or(0),
        )
        .max(1);
    let column_sizes =
        super::normalized_odf_table_sizes(&table.column_widths, column_count, table.bounds.width);
    let row_sizes = odp_table_row_sizes(
        &table,
        &column_sizes,
        table.bounds.height,
        styles,
        font_metrics,
    );
    for (row_index, row) in table.rows.iter().enumerate() {
        for (column_index, cell) in row.cells.iter().enumerate() {
            if cell.covered {
                continue;
            }
            let x = table.bounds.x + column_sizes.iter().take(column_index).sum::<f32>();
            let y = table.bounds.y + row_sizes.iter().take(row_index).sum::<f32>();
            let width = column_sizes
                .iter()
                .skip(column_index)
                .take(
                    cell.column_span
                        .min(column_count.saturating_sub(column_index))
                        .max(1),
                )
                .sum::<f32>();
            let height = row_sizes
                .iter()
                .skip(row_index)
                .take(
                    cell.row_span
                        .min(row_count.saturating_sub(row_index))
                        .max(1),
                )
                .sum::<f32>();
            let bounds = Rect {
                x,
                y,
                width,
                height,
            };
            let has_text = !cell.text.is_empty();
            let visual = if semantic_only {
                Visual::None
            } else {
                let style = odp_table_cell_style(
                    &table,
                    row_index,
                    column_index,
                    cell.row_span,
                    cell.column_span,
                    row_count,
                    column_count,
                    cell.style_name.as_deref(),
                    styles,
                );
                let fill = resolve_odf_paint(&style.fill, bounds, styles);
                let stroke = resolve_odf_paint(&style.stroke, bounds, styles);
                let stroke_width = if matches!(style.stroke, OdfPaint::None) {
                    0.0
                } else {
                    style.stroke_width
                };
                if has_text {
                    let runs = odf_text_runs(&cell.text, &cell.text_runs, &style, styles, &[]);
                    let line_height =
                        odf_text_line_height(&style, &cell.paragraph_style_names, &runs, styles);
                    let layout = odf_text_layout(
                        &style,
                        &cell.paragraph_style_names,
                        &cell.paragraph_list_layouts,
                        &runs,
                        line_height,
                        false,
                        styles,
                    );
                    Visual::TextLayout {
                        layout,
                        visual: Box::new(Visual::RichText {
                            geometry: Geometry::Rectangle,
                            fill,
                            stroke,
                            stroke_width,
                            align: odf_text_align(&style, &cell.paragraph_style_names, styles),
                            line_height,
                            runs,
                        }),
                    }
                } else {
                    Visual::PaintedShape {
                        geometry: Geometry::Rectangle,
                        fill,
                        stroke,
                        stroke_width,
                    }
                }
            };
            let visual = layer_visual(visual, table.transform, 1.0);
            let numeric_id = u32::try_from(objects.len())
                .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
            objects.push(Object {
                numeric_id,
                parent_numeric_id: Some(table_id),
                stable_id: format!("object:{numeric_id}"),
                parent_stable_id: Some(format!("object:{table_id}")),
                kind: ObjectKind::Cell,
                unit_index,
                bounds,
                z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
                text: has_text.then(|| cell.text.clone()),
                source: SourceRef {
                    part: CONTENT_PART.to_owned(),
                    mapping: if cell.element_id.is_some() {
                        MappingQuality::Exact
                    } else {
                        MappingQuality::Derived
                    },
                    locator: SourceLocator::OdpElement {
                        element_id: cell.element_id.clone(),
                        path: format!(
                            "/office:document-content/office:body/office:presentation/draw:page[{}]/table:table[1]/table:table-row[{}]/table:table-cell[{}]",
                            unit_index + 1,
                            row_index + 1,
                            column_index + 1
                        ),
                        row: Some(u32::try_from(row_index).unwrap_or(u32::MAX)),
                        column: Some(u32::try_from(column_index).unwrap_or(u32::MAX)),
                    },
                },
                visual,
            });
        }
    }
    Ok(())
}

fn odp_table_row_sizes(
    table: &OdpTableState,
    column_sizes: &[f32],
    available: f32,
    styles: &StyleCatalog,
    font_metrics: &FontMetricTable,
) -> Vec<f32> {
    let rows = &table.rows;
    if rows.is_empty() {
        return Vec::new();
    }
    let row_count = rows.len();
    let column_count = column_sizes.len();
    let minimums = rows
        .iter()
        .map(|row| row.minimum_height.max(1.0))
        .collect::<Vec<_>>();
    let mut sizes = minimums.clone();
    let mut spanning_constraints = Vec::new();

    // Resolve single-row cells first. A spanning cell constrains the sum of its
    // rows; assigning an equal share up front over-allocates tables when one of
    // those rows is already tall because of another cell.
    for (row_index, row) in rows.iter().enumerate() {
        for (column_index, cell) in row.cells.iter().enumerate() {
            if cell.covered {
                continue;
            }
            let style = odp_table_cell_style(
                table,
                row_index,
                column_index,
                cell.row_span,
                cell.column_span,
                row_count,
                column_count,
                cell.style_name.as_deref(),
                styles,
            );
            if cell.text.is_empty() {
                continue;
            }
            let width = column_sizes
                .iter()
                .skip(column_index)
                .take(cell.column_span.max(1))
                .sum::<f32>();
            let required =
                estimated_odp_table_cell_height(cell, width, &style, styles, font_metrics);
            let span = cell
                .row_span
                .min(rows.len().saturating_sub(row_index))
                .max(1);
            if span == 1 {
                sizes[row_index] = sizes[row_index].max(required);
            } else {
                spanning_constraints.push((row_index, span, required));
            }
        }
    }
    for (row_index, span, required) in spanning_constraints {
        let current = sizes.iter().skip(row_index).take(span).sum::<f32>();
        if required > current {
            let extra = (required - current) / span as f32;
            for size in sizes.iter_mut().skip(row_index).take(span) {
                *size += extra;
            }
        }
    }

    let mut total = sizes.iter().sum::<f32>();
    if !total.is_finite() || total <= f32::EPSILON {
        return vec![available / rows.len() as f32; rows.len()];
    }

    // Office keeps the authored table frame as the layout constraint. Intrinsic
    // text estimates must not make every automatic row grow beyond that frame.
    // Compress only the amount above authored minima, preserving explicit row
    // heights and span needs.
    let target = available;
    if available.is_finite() && target.is_finite() && total > target {
        let shrinkable = sizes
            .iter()
            .zip(&minimums)
            .map(|(size, minimum)| (size - minimum).max(0.0))
            .sum::<f32>();
        let reduction = (total - target).min(shrinkable);
        if reduction > 0.0 && shrinkable > f32::EPSILON {
            for (size, minimum) in sizes.iter_mut().zip(&minimums) {
                let share = (*size - *minimum).max(0.0) / shrinkable;
                *size = (*size - reduction * share).max(*minimum);
            }
            total = sizes.iter().sum::<f32>();
        }
    }
    if available.is_finite() && available > total {
        let stretch_rows = rows
            .iter()
            .enumerate()
            .filter_map(|(index, row)| {
                (row.minimum_height <= f32::EPSILON && row.cells.iter().all(|cell| cell.covered))
                    .then_some(index)
            })
            .collect::<Vec<_>>();
        if !stretch_rows.is_empty() {
            let extra = (available - total) / stretch_rows.len() as f32;
            for index in stretch_rows {
                sizes[index] += extra;
            }
        }
    }
    sizes
}

fn estimated_odp_table_cell_height(
    cell: &OdpTableCellState,
    width: f32,
    style: &GraphicStyle,
    styles: &StyleCatalog,
    font_metrics: &FontMetricTable,
) -> f32 {
    let runs = odf_text_runs(&cell.text, &cell.text_runs, style, styles, &[]);
    let line_height = odf_text_line_height(style, &cell.paragraph_style_names, &runs, styles);
    let layout = odf_text_layout(
        style,
        &cell.paragraph_style_names,
        &cell.paragraph_list_layouts,
        &runs,
        line_height,
        false,
        styles,
    );
    let font_size = runs
        .iter()
        .filter(|run| !run.text.is_empty() && run.text != "\n")
        .map(|run| run.font_size)
        .reduce(f32::max)
        .unwrap_or(style.font_size)
        .max(1.0);
    let resolved_line_height = if line_height > 0.0 {
        line_height
    } else {
        font_size * 1.2
    };
    let paragraphs = cell.text.split('\n').collect::<Vec<_>>();
    let lines = paragraphs
        .iter()
        .enumerate()
        .map(|(index, paragraph)| {
            let paragraph_layout = layout
                .paragraphs
                .get(index)
                .or_else(|| layout.paragraphs.last());
            let margin_left = paragraph_layout.map_or(0.0, |paragraph| paragraph.margin_left);
            let margin_right = paragraph_layout.map_or(0.0, |paragraph| paragraph.margin_right);
            let usable_width =
                (width - layout.inset_left - layout.inset_right - margin_left - margin_right)
                    .max(1.0);
            estimated_odp_wrapped_lines(paragraph, style, font_size, usable_width, font_metrics)
        })
        .sum::<f32>()
        .max(1.0);
    lines * resolved_line_height
        + layout.paragraph_spacing * paragraphs.len().saturating_sub(1) as f32
        + layout.inset_top
        + layout.inset_bottom
}

fn estimated_odp_wrapped_lines(
    text: &str,
    style: &GraphicStyle,
    font_size: f32,
    usable_width: f32,
    font_metrics: &FontMetricTable,
) -> f32 {
    let text = text.split_once('\t').map_or(text, |(_, content)| content);
    let mut line_count = 1.0_f32;
    let mut line_width = 0.0_f32;
    for word in text.split_whitespace() {
        let word_width = word
            .chars()
            .map(|character| {
                estimated_odp_character_width(character, style, font_size, font_metrics)
            })
            .sum::<f32>();
        let separator_width = if line_width > 0.0 {
            estimated_odp_character_width(' ', style, font_size, font_metrics)
        } else {
            0.0
        };
        if line_width > 0.0 && line_width + separator_width + word_width > usable_width {
            line_count += 1.0;
            line_width = 0.0;
        }
        if word_width > usable_width {
            let wrapped = (word_width / usable_width).ceil().max(1.0);
            line_count += wrapped - 1.0;
            line_width = word_width - usable_width * (wrapped - 1.0);
        } else {
            line_width += separator_width + word_width;
        }
    }
    line_count
}

fn estimated_odp_character_width(
    character: char,
    style: &GraphicStyle,
    font_size: f32,
    font_metrics: &FontMetricTable,
) -> f32 {
    let advance = font_metrics
        .advance_em_at_size(
            &style.font_family,
            style.italic,
            style.bold,
            character,
            font_size,
        )
        .map(|advance| advance * font_size)
        .unwrap_or_else(|| match character {
            ' ' => font_size * 0.28,
            character if character.is_ascii_punctuation() => font_size * 0.35,
            character if character.is_ascii() => font_size * 0.45,
            _ => font_size,
        });
    advance.max(0.0)
}

fn push_odp_chart(
    frame: &ChartFrame,
    chart: &OdpBasicChart,
    unit_index: u32,
    objects: &mut Vec<Object>,
    object_limit: usize,
) -> Result<(), Diagnostic> {
    let point_count = chart
        .series
        .iter()
        .map(|series| series.values.len())
        .sum::<usize>();
    let labeled_series_count = chart
        .series
        .iter()
        .filter(|series| series.label.is_some())
        .count();
    let custom_label_count = chart
        .series
        .iter()
        .flat_map(|series| &series.custom_labels)
        .filter(|label| label.is_some())
        .count();
    let title_count = usize::from(chart.title.is_some());
    let decoration_count = (if chart.kind == OdpChartKind::Pie {
        labeled_series_count
            .saturating_mul(2)
            .saturating_add(custom_label_count)
    } else {
        chart
            .categories
            .len()
            .saturating_add(24)
            .saturating_add(title_count)
            .saturating_add(labeled_series_count.saturating_mul(2))
    })
    .saturating_add(title_count * usize::from(chart.kind == OdpChartKind::Pie));
    if objects
        .len()
        .checked_add(point_count)
        .and_then(|required| required.checked_add(decoration_count))
        .and_then(|required| required.checked_add(3))
        .is_none_or(|required| required > object_limit)
    {
        return Err(object_limit_error());
    }
    let chart_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    let mapping = if frame.element_id.is_some() {
        MappingQuality::Exact
    } else {
        MappingQuality::Derived
    };
    objects.push(Object {
        numeric_id: chart_id,
        parent_numeric_id: frame.parent_numeric_id,
        stable_id: format!("object:{chart_id}"),
        parent_stable_id: frame
            .parent_numeric_id
            .map(|parent| format!("object:{parent}")),
        kind: ObjectKind::Group,
        unit_index,
        bounds: frame.bounds,
        z: i32::try_from(chart_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping,
            locator: SourceLocator::OdpElement {
                element_id: frame.element_id.clone(),
                path: format!(
                    "/office:document-content/office:body/office:presentation/draw:page[{}]/draw:frame[{}]/chart:chart[1]",
                    unit_index + 1,
                    frame.frame_index + 1
                ),
                row: None,
                column: None,
            },
        },
        visual: layer_visual(
            Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: Paint::Solid(0xffff_ffff),
                stroke: if chart.border_visible {
                    Paint::Solid(0xd1d5_dbff)
                } else {
                    Paint::None
                },
                stroke_width: if chart.border_visible { 1.0 } else { 0.0 },
            },
            frame.transform,
            1.0,
        ),
    });
    let plot = odp_chart_plot(chart, frame.bounds);
    if let Some(title) = chart.title.as_deref() {
        push_odp_chart_text(
            objects,
            chart_id,
            unit_index,
            &chart.source_part,
            Rect {
                x: frame.bounds.x + frame.bounds.width * 0.08,
                y: frame.bounds.y + frame.bounds.height * 0.02,
                width: frame.bounds.width * 0.84,
                height: frame.bounds.height * 0.12,
            },
            title,
            (frame.bounds.height * 0.055).clamp(14.0, 22.0),
            TextAlign::Center,
            frame,
            "chart:title[1]",
            None,
            None,
            MappingQuality::Exact,
        )?;
    }
    if let Some(wall) = chart.wall {
        let fill = odp_chart_wall_fill(wall, plot);
        push_odp_chart_shape(
            objects,
            chart_id,
            unit_index,
            &chart.source_part,
            plot,
            Geometry::Rectangle,
            fill,
            Paint::None,
            0.0,
            frame,
            Some("chart:plot-area[1]/chart:wall[1]"),
            None,
            None,
            MappingQuality::Exact,
        )?;
    }
    if chart.kind != OdpChartKind::Pie {
        for (index, axis_bounds) in odp_chart_axis_bounds(chart, plot).into_iter().enumerate() {
            if index == 0 && !chart.value_axis_visible {
                continue;
            }
            push_odp_chart_shape(
                objects,
                chart_id,
                unit_index,
                &chart.source_part,
                axis_bounds,
                Geometry::Line,
                Paint::None,
                Paint::Solid(0x6b72_80ff),
                1.0,
                frame,
                None,
                None,
                None,
                MappingQuality::Derived,
            )?;
        }
    }
    match chart.kind {
        OdpChartKind::Bar => {
            let category_count = chart
                .series
                .iter()
                .map(|series| series.values.len())
                .max()
                .unwrap_or(0);
            let (minimum, maximum) = odp_chart_value_range(chart);
            if category_count == 0 || maximum <= 0.0 {
                return Ok(());
            }
            let (axis_minimum, axis_maximum, tick_step) = odp_chart_axis(minimum, maximum);
            push_odp_cartesian_chart_decorations(
                objects,
                chart_id,
                unit_index,
                chart,
                frame,
                plot,
                category_count,
                axis_minimum,
                axis_maximum,
                tick_step,
            )?;
            for (series_index, category_index, bounds) in
                odp_chart_bar_bounds(chart, plot, axis_minimum, axis_maximum)
            {
                push_odp_chart_shape(
                    objects,
                    chart_id,
                    unit_index,
                    &chart.source_part,
                    bounds,
                    Geometry::Rectangle,
                    Paint::Solid(odp_chart_series_color(
                        &chart.series[series_index],
                        series_index,
                    )),
                    Paint::None,
                    0.0,
                    frame,
                    None,
                    Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                    Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                    MappingQuality::Exact,
                )?;
            }
        }
        OdpChartKind::Line | OdpChartKind::Scatter => {
            let category_count = chart
                .series
                .iter()
                .map(|series| series.values.len())
                .max()
                .unwrap_or(0);
            let (minimum, maximum) = chart
                .series
                .iter()
                .flat_map(|series| series.values.iter().copied())
                .filter(|value| value.is_finite() && *value >= 0.0)
                .fold(
                    (f32::INFINITY, f32::NEG_INFINITY),
                    |(minimum, maximum), value| (minimum.min(value), maximum.max(value)),
                );
            if category_count < 2 || maximum <= 0.0 {
                return Ok(());
            }
            let (axis_minimum, axis_maximum, tick_step) = odp_chart_axis(minimum, maximum);
            push_odp_cartesian_chart_decorations(
                objects,
                chart_id,
                unit_index,
                chart,
                frame,
                plot,
                category_count,
                axis_minimum,
                axis_maximum,
                tick_step,
            )?;
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
                    let (x0, y0, category_index) = segment[0];
                    let (x1, y1, _) = segment[1];
                    let bounds = Rect {
                        x: x0.min(x1),
                        y: y0.min(y1),
                        width: (x1 - x0).abs(),
                        height: (y1 - y0).abs(),
                    };
                    push_odp_chart_shape(
                        objects,
                        chart_id,
                        unit_index,
                        &chart.source_part,
                        bounds,
                        Geometry::Path {
                            fill_rule: FillRule::NonZero,
                            commands: super::odf_chart::line_segment_commands(
                                (x0 - bounds.x, y0 - bounds.y),
                                (x1 - bounds.x, y1 - bounds.y),
                                None,
                            ),
                        },
                        Paint::None,
                        Paint::Solid(odp_chart_series_color(series, series_index)),
                        2.0,
                        frame,
                        None,
                        Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                        Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                        MappingQuality::Exact,
                    )?;
                }
            }
        }
        OdpChartKind::Pie => {
            let Some(series) = chart.series.first() else {
                return Ok(());
            };
            let total = series
                .values
                .iter()
                .copied()
                .filter(|value| value.is_finite() && *value > 0.0)
                .sum::<f32>();
            if total <= 0.0 {
                return Ok(());
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
                let (bounds, geometry) = odp_pie_segment_geometry(
                    center_x + offset_x,
                    center_y + offset_y,
                    radius,
                    start_angle,
                    sweep,
                );
                push_odp_chart_shape(
                    objects,
                    chart_id,
                    unit_index,
                    &chart.source_part,
                    bounds,
                    geometry,
                    Paint::Solid(super::office_chart_palette_color(category_index)),
                    Paint::Solid(0xffff_ffff),
                    1.0,
                    frame,
                    None,
                    Some(0),
                    Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                    MappingQuality::Exact,
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
                let generated =
                    labels.and_then(|labels| labels.format_label(category, value, total));
                if let Some(label) = custom.or(generated.as_deref()) {
                    let font_size = labels
                        .map_or((frame.bounds.height * 0.035).clamp(10.0, 14.0), |labels| {
                            labels.font_size.max(8.0)
                        });
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
                    push_odp_chart_text(
                        objects,
                        chart_id,
                        unit_index,
                        &chart.source_part,
                        Rect {
                            x: label_x - label_width / 2.0,
                            y: label_y - label_height / 2.0,
                            width: label_width,
                            height: label_height,
                        },
                        label,
                        font_size,
                        TextAlign::Center,
                        frame,
                        &source_path,
                        Some(0),
                        Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                        MappingQuality::Exact,
                    )?;
                }
                start_angle += sweep;
            }
        }
    }
    Ok(())
}

pub(super) fn odp_pie_segment_geometry(
    center_x: f32,
    center_y: f32,
    radius: f32,
    start_angle: f32,
    sweep: f32,
) -> (Rect, Geometry) {
    let segment_count = (sweep.abs() / (std::f32::consts::PI / 18.0))
        .ceil()
        .max(2.0) as usize;
    let mut points = Vec::with_capacity(segment_count + 2);
    points.push((center_x, center_y));
    for index in 0..=segment_count {
        let angle = start_angle + sweep * index as f32 / segment_count as f32;
        points.push((
            center_x + radius * angle.cos(),
            center_y + radius * angle.sin(),
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
        x: min_x,
        y: min_y,
        width: max_x - min_x,
        height: max_y - min_y,
    };
    let commands = points
        .into_iter()
        .enumerate()
        .map(|(index, (x, y))| {
            if index == 0 {
                PathCommand::MoveTo {
                    x: x - bounds.x,
                    y: y - bounds.y,
                }
            } else {
                PathCommand::LineTo {
                    x: x - bounds.x,
                    y: y - bounds.y,
                }
            }
        })
        .chain(std::iter::once(PathCommand::ClosePath))
        .collect();
    (
        bounds,
        Geometry::Path {
            fill_rule: FillRule::NonZero,
            commands,
        },
    )
}

pub(super) use super::odf_chart::value_axis_range as odp_chart_axis;

pub(super) fn odp_chart_plot(chart: &OdpBasicChart, bounds: Rect) -> Rect {
    chart.plot_area.map_or_else(
        || {
            let has_legend =
                chart.show_legend && chart.series.iter().any(|series| series.label.is_some());
            let has_title = chart.title.is_some();
            Rect {
                x: bounds.x + bounds.width * 0.08,
                y: bounds.y + bounds.height * if has_title { 0.18 } else { 0.06 },
                width: bounds.width * if has_legend { 0.72 } else { 0.86 },
                height: bounds.height * if has_title { 0.66 } else { 0.78 },
            }
        },
        |plot| Rect {
            x: bounds.x + bounds.width * plot.x,
            y: bounds.y + bounds.height * plot.y,
            width: bounds.width * plot.width,
            height: bounds.height * plot.height,
        },
    )
}

pub(super) fn odp_chart_value_range(chart: &OdpBasicChart) -> (f32, f32) {
    if chart.stacked {
        let category_count = chart
            .series
            .iter()
            .map(|series| series.values.len())
            .max()
            .unwrap_or(0);
        let maximum = (0..category_count)
            .map(|category| {
                chart
                    .series
                    .iter()
                    .filter_map(|series| series.values.get(category))
                    .filter(|value| value.is_finite() && **value > 0.0)
                    .sum::<f32>()
            })
            .fold(0.0_f32, f32::max);
        return (0.0, maximum);
    }
    chart
        .series
        .iter()
        .flat_map(|series| series.values.iter().copied())
        .filter(|value| value.is_finite() && *value > 0.0)
        .fold(
            (f32::INFINITY, f32::NEG_INFINITY),
            |(minimum, maximum), value| (minimum.min(value), maximum.max(value)),
        )
}

pub(super) fn odp_chart_axis_bounds(chart: &OdpBasicChart, plot: Rect) -> [Rect; 2] {
    [
        Rect {
            x: if chart.value_axis_at_end {
                plot.x + plot.width
            } else {
                plot.x
            },
            y: plot.y,
            width: 0.01,
            height: plot.height,
        },
        Rect {
            x: plot.x,
            y: if chart.category_axis_at_end {
                plot.y
            } else {
                plot.y + plot.height
            },
            width: plot.width,
            height: 0.01,
        },
    ]
}

pub(super) fn odp_chart_category_label_bounds(
    chart: &OdpBasicChart,
    plot: Rect,
    category_count: usize,
    category_index: usize,
    font_size: f32,
    text_height: f32,
) -> Rect {
    let category_width = plot.width / category_count.max(1) as f32;
    Rect {
        x: plot.x + category_index as f32 * category_width,
        y: if chart.category_axis_at_end {
            plot.y - text_height - font_size * 0.35
        } else {
            plot.y + plot.height + font_size * 0.35
        },
        width: category_width,
        height: text_height,
    }
}

pub(super) fn odp_chart_value_label_bounds(
    chart: &OdpBasicChart,
    bounds: Rect,
    plot: Rect,
    font_size: f32,
    text_height: f32,
    y: f32,
) -> (Rect, TextAlign) {
    if chart.value_axis_at_end {
        let x = plot.x + plot.width + font_size * 0.4;
        (
            Rect {
                x,
                y: y - text_height / 2.0,
                width: (bounds.x + bounds.width - x).max(font_size),
                height: text_height,
            },
            TextAlign::Start,
        )
    } else {
        (
            Rect {
                x: bounds.x,
                y: y - text_height / 2.0,
                width: (plot.x - bounds.x - font_size * 0.4).max(font_size),
                height: text_height,
            },
            TextAlign::End,
        )
    }
}

pub(super) fn odp_chart_bar_bounds(
    chart: &OdpBasicChart,
    plot: Rect,
    axis_minimum: f32,
    axis_maximum: f32,
) -> Vec<(usize, usize, Rect)> {
    let category_count = chart
        .series
        .iter()
        .map(|series| series.values.len())
        .max()
        .unwrap_or(0);
    if category_count == 0 || axis_maximum <= axis_minimum {
        return Vec::new();
    }
    let category_width = plot.width / category_count as f32;
    let bar_width = category_width * 0.70
        / if chart.stacked {
            1.0
        } else {
            chart.series.len().max(1) as f32
        };
    let mut stacked = vec![0.0_f32; category_count];
    let mut bounds = Vec::new();
    for (series_index, series) in chart.series.iter().enumerate() {
        for (category_index, value) in series.values.iter().copied().enumerate() {
            if !value.is_finite() || value <= 0.0 {
                continue;
            }
            let start = if chart.stacked {
                stacked[category_index]
            } else {
                axis_minimum
            };
            let end = if chart.stacked { start + value } else { value };
            if chart.stacked {
                stacked[category_index] = end;
            }
            let y0 = plot.y + plot.height
                - plot.height * (start - axis_minimum) / (axis_maximum - axis_minimum);
            let y1 = plot.y + plot.height
                - plot.height * (end - axis_minimum) / (axis_maximum - axis_minimum);
            bounds.push((
                series_index,
                category_index,
                Rect {
                    x: plot.x
                        + category_index as f32 * category_width
                        + category_width * 0.15
                        + if chart.stacked {
                            0.0
                        } else {
                            series_index as f32 * bar_width
                        },
                    y: y0.min(y1),
                    width: bar_width,
                    height: (y1 - y0).abs(),
                },
            ));
        }
    }
    bounds
}

#[allow(clippy::too_many_arguments)]
fn push_odp_cartesian_chart_decorations(
    objects: &mut Vec<Object>,
    chart_id: u32,
    unit_index: u32,
    chart: &OdpBasicChart,
    frame: &ChartFrame,
    plot: Rect,
    category_count: usize,
    axis_minimum: f32,
    axis_maximum: f32,
    tick_step: f32,
) -> Result<(), Diagnostic> {
    let font_size = (frame.bounds.height * 0.035).clamp(10.0, 14.0);
    let text_height = font_size * 1.4;
    for (category_index, category) in chart.categories.iter().take(category_count).enumerate() {
        if category.is_empty() {
            continue;
        }
        push_odp_chart_text(
            objects,
            chart_id,
            unit_index,
            &chart.source_part,
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
            frame,
            "chart:plot-area[1]/chart:axis[1]/chart:categories[1]",
            None,
            Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
            MappingQuality::Exact,
        )?;
    }

    let tick_count = (((axis_maximum - axis_minimum) / tick_step).round() as usize).min(12);
    for tick_index in 0..=tick_count {
        let value = axis_minimum + tick_step * tick_index as f32;
        let y = plot.y + plot.height
            - plot.height * (value - axis_minimum) / (axis_maximum - axis_minimum);
        if tick_index != 0 {
            push_odp_chart_shape(
                objects,
                chart_id,
                unit_index,
                &chart.source_part,
                Rect {
                    x: plot.x,
                    y,
                    width: plot.width,
                    height: 0.01,
                },
                Geometry::Line,
                Paint::None,
                Paint::Solid(0xb3b3_b3ff),
                1.0,
                frame,
                None,
                None,
                Some(u32::try_from(tick_index).unwrap_or(u32::MAX)),
                MappingQuality::Derived,
            )?;
        }
        if chart.value_axis_visible {
            let (bounds, align) =
                odp_chart_value_label_bounds(chart, frame.bounds, plot, font_size, text_height, y);
            push_odp_chart_text(
                objects,
                chart_id,
                unit_index,
                &chart.source_part,
                bounds,
                &format_odp_chart_number(value),
                font_size,
                align,
                frame,
                "chart:plot-area[1]/chart:axis[2]",
                Some(u32::try_from(tick_index).unwrap_or(u32::MAX)),
                None,
                MappingQuality::Derived,
            )?;
        }
    }

    let legend = chart
        .series
        .iter()
        .enumerate()
        .filter(|_| chart.show_legend)
        .filter_map(|(index, series)| series.label.as_deref().map(|label| (index, label)))
        .collect::<Vec<_>>();
    let legend_row_height = font_size * 1.55;
    let legend_y = plot.y + (plot.height - legend.len() as f32 * legend_row_height) / 2.0;
    for (legend_index, (series_index, label)) in legend.into_iter().enumerate() {
        let swatch_size = font_size * 0.72;
        let row_y = legend_y + legend_index as f32 * legend_row_height;
        let swatch_x = plot.x + plot.width + frame.bounds.width * 0.03;
        push_odp_chart_shape(
            objects,
            chart_id,
            unit_index,
            &chart.source_part,
            Rect {
                x: swatch_x,
                y: row_y + (legend_row_height - swatch_size) / 2.0,
                width: swatch_size,
                height: swatch_size,
            },
            Geometry::Rectangle,
            Paint::Solid(odp_chart_series_color(
                &chart.series[series_index],
                series_index,
            )),
            Paint::None,
            0.0,
            frame,
            None,
            Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
            None,
            MappingQuality::Derived,
        )?;
        let text_x = swatch_x + swatch_size + font_size * 0.35;
        push_odp_chart_text(
            objects,
            chart_id,
            unit_index,
            &chart.source_part,
            Rect {
                x: text_x,
                y: row_y,
                width: (frame.bounds.x + frame.bounds.width - text_x).max(font_size),
                height: legend_row_height,
            },
            label,
            font_size,
            TextAlign::Start,
            frame,
            "chart:legend[1]",
            Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
            None,
            MappingQuality::Exact,
        )?;
    }
    Ok(())
}

pub(super) fn format_odp_chart_number(value: f32) -> String {
    if (value - value.round()).abs() < 0.000_01 {
        return format!("{value:.0}");
    }
    format!("{value:.6}")
        .trim_end_matches('0')
        .trim_end_matches('.')
        .to_owned()
}

#[allow(clippy::too_many_arguments)]
fn push_odp_chart_text(
    objects: &mut Vec<Object>,
    parent_numeric_id: u32,
    unit_index: u32,
    source_part: &str,
    bounds: Rect,
    text: &str,
    font_size: f32,
    align: TextAlign,
    frame: &ChartFrame,
    source_path: &str,
    row: Option<u32>,
    column: Option<u32>,
    mapping: MappingQuality,
) -> Result<(), Diagnostic> {
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: Some(parent_numeric_id),
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: Some(format!("object:{parent_numeric_id}")),
        kind: ObjectKind::TextBox,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: Some(text.to_owned()),
        source: SourceRef {
            part: source_part.to_owned(),
            mapping,
            locator: SourceLocator::OdpElement {
                element_id: frame.element_id.clone(),
                path: format!(
                    "/office:document-content/office:body/office:presentation/draw:page[{}]/draw:frame[{}]/chart:chart[1]/{source_path}",
                    unit_index + 1,
                    frame.frame_index + 1,
                ),
                row,
                column,
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
fn push_odp_chart_shape(
    objects: &mut Vec<Object>,
    parent_numeric_id: u32,
    unit_index: u32,
    source_part: &str,
    bounds: Rect,
    geometry: Geometry,
    fill: Paint,
    stroke: Paint,
    stroke_width: f32,
    frame: &ChartFrame,
    source_path: Option<&str>,
    row: Option<u32>,
    column: Option<u32>,
    mapping: MappingQuality,
) -> Result<(), Diagnostic> {
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: Some(parent_numeric_id),
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: Some(format!("object:{parent_numeric_id}")),
        kind: ObjectKind::Shape,
        unit_index,
        bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: source_part.to_owned(),
            mapping,
            locator: SourceLocator::OdpElement {
                element_id: frame.element_id.clone(),
                path: format!(
                    "/office:document-content/office:body/office:presentation/draw:page[{}]/draw:frame[{}]/chart:chart[1]/{}",
                    unit_index + 1,
                    frame.frame_index + 1,
                    source_path.map(str::to_owned).unwrap_or_else(|| format!(
                        "chart:series[{}]/chart:data-point[{}]",
                        row.unwrap_or(0) + 1,
                        column.unwrap_or(0) + 1
                    ))
                ),
                row,
                column,
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

pub(super) fn odp_chart_series_color(series: &OdpChartSeries, index: usize) -> u32 {
    series
        .color
        .unwrap_or_else(|| super::office_chart_palette_color(index))
}

fn push_odp_math(
    frame: &FrameState,
    math: &OdfMath,
    unit_index: u32,
    objects: &mut Vec<Object>,
    object_limit: usize,
) -> Result<(), Diagnostic> {
    if objects.len() >= object_limit {
        return Err(object_limit_error());
    }
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(&math.source_part, "object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: frame.parent_numeric_id,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: frame
            .parent_numeric_id
            .map(|parent| format!("object:{parent}")),
        kind: ObjectKind::TextBox,
        unit_index,
        bounds: frame.bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: Some(math.text.clone()),
        source: SourceRef {
            part: math.source_part.clone(),
            mapping: if frame.element_id.is_some() {
                MappingQuality::Exact
            } else {
                MappingQuality::Derived
            },
            locator: SourceLocator::OdpElement {
                element_id: frame.element_id.clone(),
                path: format!(
                    "/office:document-content/office:body/office:presentation/draw:page[{}]/draw:frame[{}]/draw:object[1]/math:math[1]",
                    unit_index + 1,
                    frame.frame_index + 1
                ),
                row: None,
                column: None,
            },
        },
        visual: layer_visual(
            Visual::RichText {
                geometry: Geometry::Rectangle,
                fill: Paint::None,
                stroke: Paint::None,
                stroke_width: 0.0,
                align: TextAlign::Center,
                line_height: 0.0,
                runs: vec![TextRun {
                    paint: None,
                    east_asian_line_breaks: true,
                    text: math.text.clone(),
                    font_family: "Cambria Math".to_owned(),
                    font_size: 18.0 * CSS_PIXELS_PER_INCH / 72.0,
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
            frame.transform,
            1.0,
        ),
    });
    Ok(())
}

fn push_odp_placeholder(
    frame: &FrameState,
    label: &str,
    unit_index: u32,
    objects: &mut Vec<Object>,
    object_limit: usize,
) -> Result<(), Diagnostic> {
    if objects.len() >= object_limit {
        return Err(object_limit_error());
    }
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    objects.push(Object {
        numeric_id,
        parent_numeric_id: frame.parent_numeric_id,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: frame
            .parent_numeric_id
            .map(|parent| format!("object:{parent}")),
        kind: ObjectKind::Unknown,
        unit_index,
        bounds: frame.bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: Some(label.to_owned()),
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping: if frame.element_id.is_some() {
                MappingQuality::Exact
            } else {
                MappingQuality::Derived
            },
            locator: SourceLocator::OdpElement {
                element_id: frame.element_id.clone(),
                path: format!(
                    "/office:document-content/office:body/office:presentation/draw:page[{}]/draw:frame[{}]/draw:object[1]",
                    unit_index + 1,
                    frame.frame_index + 1
                ),
                row: None,
                column: None,
            },
        },
        visual: layer_visual(
            Visual::RichText {
                geometry: Geometry::RoundedRectangle {
                    radius_x: 6.0,
                    radius_y: 6.0,
                },
                fill: Paint::Solid(0xf3f4_f6ff),
                stroke: Paint::Solid(0x9ca3_afff),
                stroke_width: 1.0,
                align: TextAlign::Center,
                line_height: 0.0,
                runs: vec![odf_text_run(
                    label.to_owned(),
                    &GraphicStyle {
                        bold: true,
                        ..GraphicStyle::default()
                    },
                )],
            },
            frame.transform,
            1.0,
        ),
    });
    Ok(())
}

fn resolve_odf_paint(paint: &OdfPaint, bounds: Rect, styles: &StyleCatalog) -> Paint {
    match paint {
        OdfPaint::None => Paint::None,
        OdfPaint::Solid(color) => Paint::Solid(*color),
        OdfPaint::Gradient(name) => styles
            .gradients
            .get(name)
            .map(|gradient| resolve_odf_gradient(*gradient, bounds, gradient.start, gradient.end))
            .unwrap_or(Paint::None),
        OdfPaint::Bitmap(_) => Paint::None,
    }
}

fn resolve_odf_gradient(gradient: GradientDefinition, bounds: Rect, start: u32, end: u32) -> Paint {
    let stops = vec![
        GradientStop {
            offset: 0.0,
            color: start,
        },
        GradientStop {
            offset: 1.0,
            color: end,
        },
    ];
    if gradient.rectangular {
        return Paint::RectGradient {
            center_x: bounds.width * (1.0 - gradient.center_x),
            center_y: bounds.height * gradient.center_y,
            stops: vec![
                GradientStop {
                    offset: 0.0,
                    color: end,
                },
                GradientStop {
                    offset: 1.0,
                    color: start,
                },
            ],
        };
    }
    if gradient.radial {
        let powerpoint_corner_focus = gradient.center_y <= f32::EPSILON
            && (gradient.center_x <= f32::EPSILON || gradient.center_x >= 1.0 - f32::EPSILON);
        let center_x = bounds.width
            * if powerpoint_corner_focus {
                1.0 - gradient.center_x
            } else {
                gradient.center_x
            };
        let center_y = bounds.height * gradient.center_y;
        let radius_x = center_x.max(bounds.width - center_x);
        let radius_y = center_y.max(bounds.height - center_y);
        return Paint::RadialGradient {
            x0: center_x,
            y0: center_y,
            r0: 0.0,
            x1: center_x,
            y1: center_y,
            r1: if powerpoint_corner_focus {
                bounds.width.min(bounds.height)
            } else {
                radius_x.hypot(radius_y)
            },
            stops: if powerpoint_corner_focus {
                vec![
                    GradientStop {
                        offset: 0.0,
                        color: end,
                    },
                    GradientStop {
                        offset: 1.0,
                        color: start,
                    },
                ]
            } else {
                stops
            },
        };
    }
    // ODF measures gradient angles counter-clockwise from the vertical axis.
    let radians = (90.0 - gradient.angle_degrees).to_radians();
    let dx = radians.cos() * bounds.width / 2.0;
    let dy = radians.sin() * bounds.height / 2.0;
    Paint::LinearGradient {
        x0: bounds.width / 2.0 - dx,
        y0: bounds.height / 2.0 - dy,
        x1: bounds.width / 2.0 + dx,
        y1: bounds.height / 2.0 + dy,
        stops,
    }
}

pub(super) fn odp_chart_wall_fill(wall: OdpChartWall, bounds: Rect) -> Paint {
    wall.opacity_gradient
        .map_or(Paint::Solid(wall.color), |gradient| {
            let color = |alpha: u32| {
                let base_alpha = wall.color & 0xff;
                (wall.color & 0xffff_ff00) | ((base_alpha * (alpha & 0xff) + 127) / 255)
            };
            resolve_odf_gradient(gradient, bounds, color(gradient.start), color(gradient.end))
        })
}

fn odf_fill_opacity(paint: Paint, opacity: f32) -> (Paint, f32) {
    let opacity = opacity.clamp(0.0, 1.0);
    if opacity == 1.0 {
        return (paint, 1.0);
    }
    let color = |value: u32| {
        let alpha = ((value & 0xff) as f32 * opacity).round() as u32;
        (value & 0xffff_ff00) | alpha
    };
    let paint = match paint {
        Paint::None => Paint::None,
        Paint::Solid(value) => Paint::Solid(color(value)),
        Paint::LinearGradient {
            x0,
            y0,
            x1,
            y1,
            mut stops,
        } => {
            for stop in &mut stops {
                stop.color = color(stop.color);
            }
            Paint::LinearGradient {
                x0,
                y0,
                x1,
                y1,
                stops,
            }
        }
        Paint::RadialGradient {
            x0,
            y0,
            r0,
            x1,
            y1,
            r1,
            mut stops,
        } => {
            for stop in &mut stops {
                stop.color = color(stop.color);
            }
            Paint::RadialGradient {
                x0,
                y0,
                r0,
                x1,
                y1,
                r1,
                stops,
            }
        }
        Paint::RectGradient {
            center_x,
            center_y,
            mut stops,
        } => {
            for stop in &mut stops {
                stop.color = color(stop.color);
            }
            Paint::RectGradient {
                center_x,
                center_y,
                stops,
            }
        }
        Paint::CircleGradient {
            x0,
            y0,
            r0,
            x1,
            y1,
            r1,
            mut stops,
        } => {
            for stop in &mut stops {
                stop.color = color(stop.color);
            }
            Paint::CircleGradient {
                x0,
                y0,
                r0,
                x1,
                y1,
                r1,
                stops,
            }
        }
        Paint::Pattern {
            preset,
            foreground,
            background,
        } => Paint::Pattern {
            preset,
            foreground: color(foreground),
            background: color(background),
        },
        paint @ (Paint::Image { .. }
        | Paint::Visual { .. }
        | Paint::XpsGradient { .. }
        | Paint::ShapeGradient { .. }
        | Paint::MappedGradient { .. }) => {
            return (paint, opacity);
        }
    };
    (paint, 1.0)
}

fn resolve_odf_shape_fill(
    paint: &OdfPaint,
    tile: bool,
    tile_width: Option<f32>,
    tile_height: Option<f32>,
    bounds: Rect,
    styles: &StyleCatalog,
    materialized_image_bytes: &mut usize,
    materialized_image_limit: usize,
) -> Result<Paint, Diagnostic> {
    let OdfPaint::Bitmap(name) = paint else {
        return Ok(resolve_odf_paint(paint, bounds, styles));
    };
    let Some(image) = styles.fill_images.get(name) else {
        return Ok(Paint::None);
    };
    reserve_materialized_image_bytes(
        materialized_image_bytes,
        image.bytes.len(),
        materialized_image_limit,
        &image.target,
    )?;
    Ok(Paint::Image {
        mapping: None,
        media_type: image.media_type.clone(),
        bytes: clone_image_bytes(&image.bytes, &image.target)?,
        crop: ImageCrop::default(),
        tile,
        tile_width,
        tile_height,
    })
}

fn append_odf_text(
    text: &str,
    plain_text: &mut String,
    runs: &mut Vec<OdfTextRun>,
    style_name: Option<&str>,
    label_style: Option<OdfListLabelStyle>,
) {
    if text.is_empty() {
        return;
    }
    plain_text.push_str(text);
    if let Some(run) = runs.last_mut()
        && run.style_name.as_deref() == style_name
        && run.label_style.is_none()
        && label_style.is_none()
    {
        run.text.push_str(text);
        return;
    }
    runs.push(OdfTextRun {
        text: text.to_owned(),
        style_name: style_name.map(str::to_owned),
        label_style,
    });
}

fn decode_odf_text_node(text: &str, part: &str) -> Result<Option<String>, Diagnostic> {
    let text = decode_xml_text(text).map_err(|error| with_part(error, part))?;
    let formatting_whitespace = text.trim().is_empty()
        && text
            .bytes()
            .any(|byte| matches!(byte, b'\t' | b'\n' | b'\r'));
    Ok((!formatting_whitespace).then(|| text.into_owned()))
}

fn append_odf_paragraph_break(plain_text: &mut String, runs: &mut Vec<OdfTextRun>) {
    let style_name = runs.last().and_then(|run| run.style_name.clone());
    append_odf_text("\n", plain_text, runs, style_name.as_deref(), None);
}

fn append_odf_text_with_list_prefix(
    text: &str,
    plain_text: &mut String,
    runs: &mut Vec<OdfTextRun>,
    list_text: &mut OdfListTextState,
    style_name: Option<&str>,
) {
    if text.is_empty() {
        return;
    }
    if let Some((prefix, label_style)) = list_text.pending_prefix.take() {
        append_odf_text(&prefix, plain_text, runs, style_name, Some(label_style));
        list_text.pending_counter_restore = None;
    }
    append_odf_text(text, plain_text, runs, style_name, None);
}

fn begin_odf_list_paragraph(
    list_text: &mut OdfListTextState,
    styles: &StyleCatalog,
    frame_style: Option<&str>,
) -> Option<OdfListParagraphLayout> {
    discard_odf_pending_prefix(list_text);
    let level = u32::try_from(list_text.stack.len()).ok()?;
    let level_style = list_text
        .stack
        .iter()
        .rev()
        .find_map(|list| list.style_name.as_ref())
        .and_then(|name| styles.list_styles.get(name))
        .and_then(|levels| levels.get(&level))
        .cloned();
    let Some(active) = list_text.stack.last_mut() else {
        list_text.numbering_counters.fill(0);
        return None;
    };
    let layout = level_style.as_ref().map(|style| OdfListParagraphLayout {
        outline_style: frame_style
            .and_then(|name| styles.outline_prefixes.get(name))
            .map(|prefix| format!("{prefix}{level}"))
            .filter(|name| styles.graphics.contains_key(name)),
        margin_left: style.space_before + style.min_label_width,
        first_line_indent: -style.min_label_width,
        image: if active.item_depth.is_some() && !active.item_has_paragraph {
            match &style.label {
                OdfListLabel::Image {
                    href,
                    width,
                    height,
                } => Some((href.clone(), *width, *height)),
                _ => None,
            }
        } else {
            None
        },
    });
    let emit_prefix = active.item_depth.is_some() && !active.item_has_paragraph;
    active.item_has_paragraph = true;
    if !emit_prefix {
        return layout;
    }
    let Some(level_style) = level_style else {
        return layout;
    };
    let counter_index = usize::try_from(level.saturating_sub(1)).unwrap_or(usize::MAX);
    if counter_index == usize::MAX {
        return layout;
    }
    if list_text.numbering_counters.len() <= counter_index {
        list_text
            .numbering_counters
            .resize(counter_index.saturating_add(1), 0);
    }
    let previous_counter = list_text.numbering_counters[counter_index];
    for counter in list_text
        .numbering_counters
        .iter_mut()
        .skip(counter_index.saturating_add(1))
    {
        *counter = 0;
    }
    let prefix_style = level_style.label_style;
    let prefix = match level_style.label {
        OdfListLabel::Image { .. } => {
            list_text.numbering_counters[counter_index] = 0;
            "\t".to_owned()
        }
        OdfListLabel::Bullet(character) => {
            list_text.numbering_counters[counter_index] = 0;
            format!(
                "{}\t",
                super::normalize_symbol_font_character(
                    &character,
                    level_style.font_family.as_deref(),
                )
            )
        }
        OdfListLabel::Number {
            format,
            prefix,
            suffix,
            start_value,
        } => {
            if format.is_empty() {
                list_text.numbering_counters[counter_index] = 0;
                active.reset_numbering = false;
                return layout;
            }
            if list_text.numbering_formats.len() <= counter_index {
                list_text
                    .numbering_formats
                    .resize(counter_index + 1, String::new());
            }
            if !active.continue_numbering && list_text.numbering_formats[counter_index] != format {
                active.reset_numbering = true;
            }
            list_text.numbering_formats[counter_index].clone_from(&format);
            let explicit_start = active
                .item_start_value
                .or_else(|| (active.item_count == 1).then_some(start_value).flatten());
            let value = explicit_start.unwrap_or_else(|| {
                if active.reset_numbering || previous_counter == 0 {
                    start_value.unwrap_or(1)
                } else {
                    previous_counter.saturating_add(1)
                }
            });
            list_text.numbering_counters[counter_index] = value;
            format!(
                "{prefix}{}{suffix}\t",
                format_odf_list_number(&format, value)
            )
        }
    };
    active.reset_numbering = false;
    list_text.pending_prefix = (!prefix.is_empty()).then_some((prefix, prefix_style));
    list_text.pending_counter_restore = Some((counter_index, previous_counter));
    layout
}

fn discard_odf_pending_prefix(list_text: &mut OdfListTextState) {
    list_text.pending_prefix = None;
    if let Some((index, value)) = list_text.pending_counter_restore.take()
        && let Some(counter) = list_text.numbering_counters.get_mut(index)
    {
        *counter = value;
    }
}

fn format_odf_list_number(format: &str, value: u32) -> String {
    if let Some(number) = super::zero_padded_number(value, format) {
        return number;
    }
    match format {
        "a" => super::alphabetic_number(value, false),
        "A" => super::alphabetic_number(value, true),
        "i" => super::roman_number(value, false),
        "I" => super::roman_number(value, true),
        "①, ②, ③, ..." => super::circled_number(value).unwrap_or_else(|| value.to_string()),
        "一, 二, 三, 四, ..." => super::chinese_number(value),
        _ => value.to_string(),
    }
}

fn current_odf_text_style(
    span_depth: Option<usize>,
    span_style_name: &Option<String>,
    paragraph_style_names: &[Option<String>],
) -> Option<String> {
    if span_depth.is_some() {
        span_style_name.clone()
    } else {
        paragraph_style_names.last().cloned().flatten()
    }
}

fn odf_control_text(
    local: &str,
    attributes: &[XmlAttribute<'_>],
    max_text_bytes: usize,
    part: &str,
) -> Result<String, Diagnostic> {
    match local {
        "tab" => Ok("\t".to_owned()),
        "line-break" => Ok("\n".to_owned()),
        "s" => {
            let count = optional_attribute(attributes, "c", part)?
                .map(|value| {
                    value.parse::<usize>().map_err(|_| {
                        format_error(part, "text:s count is not a supported positive integer")
                    })
                })
                .transpose()?
                .unwrap_or(1);
            if count == 0 || count > max_text_bytes {
                return Err(format_error(
                    part,
                    "text:s count exceeds the configured XML text limit",
                ));
            }
            let mut spaces = String::new();
            spaces.try_reserve_exact(count).map_err(|_| {
                Diagnostic::fatal(
                    DiagnosticCode::AllocationFailed,
                    Phase::Parse,
                    None,
                    "unable to allocate explicit ODF spaces",
                )
                .in_part(part)
            })?;
            spaces.extend(std::iter::repeat_n(' ', count));
            Ok(spaces)
        }
        _ => Ok(String::new()),
    }
}

fn odf_text_runs(
    text: &str,
    text_runs: &[OdfTextRun],
    base_style: &GraphicStyle,
    styles: &StyleCatalog,
    paragraphs: &[Option<OdfListParagraphLayout>],
) -> Vec<TextRun> {
    if text_runs.is_empty() {
        return vec![odf_text_run(text.to_owned(), base_style)];
    }
    let mut paragraph_index = 0;
    text_runs
        .iter()
        .map(|text_run| {
            let base = odf_style_overlay(
                base_style,
                paragraphs
                    .get(paragraph_index)
                    .and_then(Option::as_ref)
                    .and_then(|p| p.outline_style.as_deref()),
                styles,
            );
            let run_style = odf_style_overlay(&base, text_run.style_name.as_deref(), styles);
            paragraph_index += text_run.text.chars().filter(|ch| *ch == '\n').count();
            let mut run = odf_text_run(text_run.text.clone(), &run_style);
            if let Some(label_style) = text_run.label_style {
                if let Some(color) = label_style.color {
                    run.color = color;
                }
                if let Some(font_size) = label_style.font_size {
                    run.font_size = match font_size {
                        OdfListFontSize::Relative(relative) => run.font_size * relative,
                        OdfListFontSize::Exact(size) => size,
                    };
                }
            }
            run
        })
        .collect()
}

fn odf_text_align(
    base_style: &GraphicStyle,
    paragraph_style_names: &[Option<String>],
    styles: &StyleCatalog,
) -> TextAlign {
    paragraph_style_names
        .first()
        .and_then(|name| name.as_ref())
        .map_or(base_style.align, |name| {
            odf_style_overlay(base_style, Some(name), styles).align
        })
}

fn odf_text_layout(
    base_style: &GraphicStyle,
    paragraph_style_names: &[Option<String>],
    paragraph_list_layouts: &[Option<OdfListParagraphLayout>],
    runs: &[TextRun],
    line_height: f32,
    apply_center_line_offset: bool,
    styles: &StyleCatalog,
) -> TextLayout {
    let paragraph_spacing = paragraph_style_names
        .iter()
        .map(|style_name| {
            odf_style_overlay(base_style, style_name.as_deref(), styles).paragraph_spacing
        })
        .fold(base_style.paragraph_spacing, f32::max);
    let text_font_size = runs
        .iter()
        .filter(|run| !run.text.is_empty() && run.text != "\n")
        .map(|run| run.font_size)
        .reduce(f32::max)
        .unwrap_or(base_style.font_size)
        .max(1.0);
    let leading_above = if base_style.orientation != TextOrientation::Horizontal {
        0.0
    } else {
        match (apply_center_line_offset, base_style.vertical_align) {
            (false, _) => 0.0,
            (true, TextVerticalAlign::Top) => ((line_height - text_font_size) / 2.0).max(0.0),
            // PowerPoint anchors multi-paragraph ODF text around paragraph line origins.
            // The renderer centers complete line boxes, so compensate by one authored
            // line only for true multi-paragraph frames. Wrapped single paragraphs are
            // already centered correctly without this offset.
            (true, TextVerticalAlign::Center)
                if paragraph_style_names.len() > 1
                    && line_height > 0.0
                    && base_style.auto_fit == TextAutoFit::None =>
            {
                line_height
            }
            (true, TextVerticalAlign::Center | TextVerticalAlign::Bottom) => 0.0,
        }
    };
    TextLayout {
        column_count: base_style.columns.0,
        column_spacing: base_style.columns.1,
        orientation: base_style.orientation,
        auto_fit: base_style.auto_fit,
        vertical_align: base_style.vertical_align,
        inset_left: base_style.inset_left,
        inset_right: base_style.inset_right,
        inset_top: base_style.inset_top + leading_above,
        inset_bottom: base_style.inset_bottom,
        wrap: base_style.auto_fit != TextAutoFit::FitFrame && base_style.wrap,
        paragraph_spacing,
        paragraphs: paragraph_style_names
            .iter()
            .enumerate()
            .map(|(index, style_name)| {
                let list_layout = paragraph_list_layouts.get(index).and_then(Option::as_ref);
                let base = odf_style_overlay(
                    base_style,
                    list_layout.and_then(|p| p.outline_style.as_deref()),
                    styles,
                );
                let paragraph_style = odf_style_overlay(&base, style_name.as_deref(), styles);
                let paragraph_line_height = match paragraph_style.line_height {
                    Some(OdfLineHeight::Multiple(multiple)) => text_font_size * 1.2 * multiple,
                    Some(OdfLineHeight::Exact(height)) => height,
                    None if list_layout.is_some_and(|p| p.outline_style.is_some()) => {
                        paragraph_style.font_size * 1.2
                    }
                    None => 0.0,
                };
                TextParagraphLayout {
                    align: paragraph_style.align,
                    margin_left: if paragraph_style.margin_left != 0.0 {
                        paragraph_style.margin_left.max(0.0)
                    } else {
                        list_layout.map_or(0.0, |layout| layout.margin_left.max(0.0))
                    },
                    margin_right: paragraph_style.margin_right.max(0.0),
                    first_line_indent: if paragraph_style.first_line_indent != 0.0 {
                        paragraph_style.first_line_indent
                    } else {
                        list_layout.map_or(0.0, |layout| layout.first_line_indent)
                    },
                    default_tab_stop: paragraph_style.default_tab_stop.max(1.0),
                    line_height: paragraph_line_height,
                    space_before: 0.0,
                    space_after: paragraph_style.paragraph_spacing,
                    latin_line_break: true,
                    hanging_punctuation: false,
                    rule_above: None,
                    rule_below: None,
                    drop_cap: None,
                }
            })
            .collect(),
        ..TextLayout::default()
    }
}

fn odf_text_line_height(
    base_style: &GraphicStyle,
    paragraph_style_names: &[Option<String>],
    runs: &[TextRun],
    styles: &StyleCatalog,
) -> f32 {
    let style = paragraph_style_names
        .iter()
        .filter_map(|style_name| {
            let style_name = style_name.as_deref()?;
            styles
                .style_patches
                .get(style_name)
                .is_some_and(|patch| patch.line_height.is_some())
                .then(|| odf_style_overlay(base_style, Some(style_name), styles))
        })
        .next()
        .unwrap_or_else(|| base_style.clone());
    let text_font_size = runs
        .iter()
        .filter(|run| !run.text.is_empty() && run.text != "\n")
        .map(|run| run.font_size)
        .reduce(f32::max)
        .unwrap_or(style.font_size)
        .max(1.0);
    match style.line_height {
        Some(OdfLineHeight::Multiple(multiple)) => text_font_size * 1.2 * multiple,
        Some(OdfLineHeight::Exact(height)) => height,
        None => 0.0,
    }
}

fn odf_text_run(text: String, style: &GraphicStyle) -> TextRun {
    TextRun {
        paint: None,
        east_asian_line_breaks: true,
        text,
        font_family: style.font_family.clone(),
        font_size: style.font_size,
        color: style.color,
        bold: style.bold,
        italic: style.italic,
        underline: false,
        strikethrough: false,
        highlight: 0,
        baseline_shift: 0.0,
        letter_spacing: 0.0,
        horizontal_scale: 1.0,
    }
}

fn odf_stroke_style_visual(
    visual: Visual,
    graphic_style: &GraphicStyle,
    styles: &StyleCatalog,
    materialized_dashes: bool,
) -> Visual {
    let mut style = if let Some(style) = graphic_style.border_style.as_ref() {
        style.clone()
    } else {
        // ODF 1.3 20.171: an explicit linecap overrides draw:stroke-dash's draw:style.
        let dash_cap = graphic_style
            .stroke_dash
            .as_ref()
            .and_then(|name| styles.stroke_dashes.get(name))
            .map(|(_, cap)| *cap);
        StrokeStyle {
            cap: graphic_style.line_cap.or(dash_cap).unwrap_or(LineCap::Flat),
            dash: resolved_odf_stroke_dash(graphic_style, styles).unwrap_or_default(),
            ..StrokeStyle::default()
        }
    };
    if materialized_dashes {
        // Connectors already split their path around dash gaps and arrow markers.
        style.dash.clear();
    }
    if style == StrokeStyle::default() {
        return visual;
    }
    Visual::StrokeStyle {
        style,
        visual: Box::new(visual),
    }
}

fn resolved_odf_stroke_dash(
    graphic_style: &GraphicStyle,
    styles: &StyleCatalog,
) -> Option<Vec<f32>> {
    let (dash, _) = graphic_style
        .stroke_dash
        .as_deref()
        .and_then(|name| styles.stroke_dashes.get(name))?;
    let resolved = dash
        .iter()
        .map(|length| match length {
            OdfDashLength::Absolute(length) => *length,
            OdfDashLength::StrokeWidthMultiple(multiple) => graphic_style.stroke_width * multiple,
        })
        .collect::<Vec<_>>();
    (!resolved.is_empty()
        && resolved
            .iter()
            .all(|length| length.is_finite() && *length > 0.0))
    .then_some(resolved)
}

fn layer_visual(visual: Visual, transform: AffineTransform, opacity: f32) -> Visual {
    if transform == AffineTransform::IDENTITY && opacity == 1.0 {
        visual
    } else {
        Visual::Layer {
            transform,
            opacity,
            blend_mode: crate::model::BlendMode::Normal,
            visual: Box::new(visual),
        }
    }
}

fn effect_visual(visual: Visual, shadow: Option<Shadow>) -> Visual {
    match shadow {
        Some(shadow) => Visual::Effect {
            shadow: Some(shadow),
            clip: None,
            visual: Box::new(visual),
        },
        None => visual,
    }
}

fn optional_length(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<f32>, Diagnostic> {
    optional_attribute(attributes, name, part)?.map_or(Ok(None), |value| {
        parse_length(&value).map(Some).ok_or_else(|| {
            format_error(
                part,
                format!("attribute {name} is not a supported ODF length"),
            )
        })
    })
}

fn optional_dash_length(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<OdfDashLength>, Diagnostic> {
    optional_attribute(attributes, name, part)?.map_or(Ok(None), |value| {
        let parsed = value
            .trim()
            .strip_suffix('%')
            .and_then(|number| number.parse::<f32>().ok())
            .filter(|number| number.is_finite() && *number > 0.0)
            .map(|percent| OdfDashLength::StrokeWidthMultiple(percent / 100.0))
            .or_else(|| {
                parse_length(&value)
                    .filter(|length| *length > 0.0)
                    .map(OdfDashLength::Absolute)
            });
        parsed.map(Some).ok_or_else(|| {
            format_error(
                part,
                format!("attribute {name} is not a supported ODF dash length"),
            )
        })
    })
}

fn parse_odf_transform(value: &str) -> Result<AffineTransform, Diagnostic> {
    parse_odf_transform_with_rotation(value, true)
}

fn parse_odf_master_transform(value: &str) -> Result<AffineTransform, Diagnostic> {
    parse_odf_transform_with_rotation(value, false)
}

fn parse_odf_transform_with_rotation(
    value: &str,
    invert_rotation: bool,
) -> Result<AffineTransform, Diagnostic> {
    let mut remaining = value.trim();
    let mut transform = AffineTransform::IDENTITY;
    while !remaining.is_empty() {
        let Some(open) = remaining.find('(') else {
            return Err(format_error(CONTENT_PART, "draw:transform is malformed"));
        };
        let name = remaining[..open].trim();
        let Some(close_relative) = remaining[open + 1..].find(')') else {
            return Err(format_error(CONTENT_PART, "draw:transform is malformed"));
        };
        let close = open + 1 + close_relative;
        let arguments = remaining[open + 1..close]
            .split(|character: char| character == ',' || character.is_whitespace())
            .filter(|value| !value.is_empty())
            .collect::<Vec<_>>();
        let number = |index: usize| -> Result<f32, Diagnostic> {
            let value = arguments.get(index).ok_or_else(|| {
                format_error(CONTENT_PART, "draw:transform is missing an argument")
            })?;
            parse_length(value)
                .or_else(|| value.parse::<f32>().ok())
                .filter(|value| value.is_finite())
                .ok_or_else(|| format_error(CONTENT_PART, "draw:transform has an invalid argument"))
        };
        let next = match name {
            "translate" => AffineTransform {
                e: number(0)?,
                f: arguments
                    .get(1)
                    .map(|_| number(1))
                    .transpose()?
                    .unwrap_or(0.0),
                ..AffineTransform::IDENTITY
            },
            "scale" => {
                let x = number(0)?;
                let y = arguments
                    .get(1)
                    .map(|_| number(1))
                    .transpose()?
                    .unwrap_or(x);
                AffineTransform {
                    a: x,
                    d: y,
                    ..AffineTransform::IDENTITY
                }
            }
            "rotate" => {
                // ODF angles use a mathematical y-up coordinate system, while
                // Canvas and the scene model use y-down screen coordinates.
                let angle = number(0)? * if invert_rotation { -1.0 } else { 1.0 };
                AffineTransform {
                    a: angle.cos(),
                    b: angle.sin(),
                    c: -angle.sin(),
                    d: angle.cos(),
                    e: 0.0,
                    f: 0.0,
                }
            }
            "skewX" | "skewY" if arguments.len() == 1 => {
                // Like rotation, ODF shear angles use the mathematical y-up axes.
                let shear = (number(0)? * if invert_rotation { -1.0 } else { 1.0 }).tan();
                if !shear.is_finite() {
                    return Err(format_error(
                        CONTENT_PART,
                        "draw:transform shear is invalid",
                    ));
                }
                AffineTransform {
                    b: if name == "skewY" { shear } else { 0.0 },
                    c: if name == "skewX" { shear } else { 0.0 },
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
                    CONTENT_PART,
                    format!("unsupported draw:transform function {name}"),
                ));
            }
        };
        transform = next.concat(transform);
        remaining =
            remaining[close + 1..].trim_start_matches(|c: char| c.is_whitespace() || c == ',');
    }
    Ok(transform)
}

fn object_limit_error() -> Diagnostic {
    Diagnostic::fatal(
        DiagnosticCode::ObjectLimit,
        Phase::Parse,
        None,
        "document exceeds the configured object limit",
    )
    .in_part(CONTENT_PART)
}

fn append_inline_base64(encoded: &mut String, text: &str) -> Result<(), Diagnostic> {
    encoded.try_reserve(text.len()).map_err(|_| {
        Diagnostic::fatal(
            DiagnosticCode::AllocationFailed,
            Phase::Parse,
            None,
            "unable to allocate storage for inline image data",
        )
        .in_part(CONTENT_PART)
    })?;
    encoded.push_str(text);
    Ok(())
}

fn decode_inline_base64(encoded: &str) -> Result<Vec<u8>, Diagnostic> {
    let symbol_count = encoded
        .bytes()
        .filter(|byte| !byte.is_ascii_whitespace())
        .count();
    if !symbol_count.is_multiple_of(4) {
        return Err(invalid_inline_base64());
    }
    let capacity = symbol_count
        .checked_div(4)
        .and_then(|groups| groups.checked_mul(3))
        .ok_or_else(invalid_inline_base64)?;
    let mut decoded = Vec::new();
    decoded.try_reserve_exact(capacity).map_err(|_| {
        Diagnostic::fatal(
            DiagnosticCode::AllocationFailed,
            Phase::Parse,
            None,
            "unable to allocate decoded inline image data",
        )
        .in_part(CONTENT_PART)
    })?;

    let mut quartet = [0_u8; 4];
    let mut quartet_len = 0_usize;
    let mut finished = false;
    for byte in encoded.bytes() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        if finished {
            return Err(invalid_inline_base64());
        }
        quartet[quartet_len] = match byte {
            b'A'..=b'Z' => byte - b'A',
            b'a'..=b'z' => byte - b'a' + 26,
            b'0'..=b'9' => byte - b'0' + 52,
            b'+' => 62,
            b'/' => 63,
            b'=' => 64,
            _ => return Err(invalid_inline_base64()),
        };
        quartet_len += 1;
        if quartet_len != 4 {
            continue;
        }
        if quartet[0] == 64 || quartet[1] == 64 {
            return Err(invalid_inline_base64());
        }
        decoded.push((quartet[0] << 2) | (quartet[1] >> 4));
        match (quartet[2], quartet[3]) {
            (64, 64) if quartet[1] & 0x0f == 0 => finished = true,
            (third, 64) if third != 64 && third & 0x03 == 0 => {
                decoded.push((quartet[1] << 4) | (third >> 2));
                finished = true;
            }
            (third, fourth) if third != 64 && fourth != 64 => {
                decoded.push((quartet[1] << 4) | (third >> 2));
                decoded.push((third << 6) | fourth);
            }
            _ => return Err(invalid_inline_base64()),
        }
        quartet_len = 0;
    }
    if quartet_len != 0 {
        return Err(invalid_inline_base64());
    }
    Ok(decoded)
}

fn invalid_inline_base64() -> Diagnostic {
    format_error(
        CONTENT_PART,
        "office:binary-data is not valid canonical base64",
    )
}

fn image_cache_key(target: &str, declared_media_type: Option<&str>) -> String {
    declared_media_type.map_or_else(
        || target.to_owned(),
        |media_type| format!("{target}\0{}", media_type.trim().to_ascii_lowercase()),
    )
}

fn identify_odf_image(
    target: &str,
    declared_media_type: Option<&str>,
    bytes: &[u8],
) -> Result<&'static str, OfficeImageError> {
    declared_media_type
        .filter(|media_type| !media_type.trim().is_empty())
        .map_or_else(
            || office_image_media_type(target, bytes),
            |media_type| office_image_media_type_from_mime(media_type, bytes),
        )
        .or_else(|error| {
            if error == OfficeImageError::UnsupportedFormat {
                office_image_media_type_from_signature(bytes)
            } else {
                Err(error)
            }
        })
}

fn push_image_frame(
    frame: &FrameState,
    unit_index: u32,
    package: &dyn OdpSource,
    styles: &StyleCatalog,
    objects: &mut Vec<Object>,
    diagnostics: &mut Vec<Diagnostic>,
    image_resources: &mut ImageResources,
) -> Result<(), Diagnostic> {
    let ImageResources {
        manifest_media_types,
        cache: image_cache,
        materialized_bytes: materialized_image_bytes,
    } = image_resources;
    if frame.images.is_empty() {
        return Ok(());
    }

    let mut selected = None;
    let mut runtime_fallback = None;
    for image in &frame.images {
        if let Some(encoded) = image.inline_base64.as_deref() {
            let bytes = decode_inline_base64(encoded)?;
            let Some(declared_media_type) = image.declared_media_type.as_deref() else {
                diagnostics.push(unsupported_image_diagnostic(
                    CONTENT_PART,
                    OfficeImageError::UnsupportedFormat,
                ));
                continue;
            };
            let media_type = match office_image_media_type_from_mime(declared_media_type, &bytes) {
                Ok(media_type) => media_type.to_owned(),
                Err(error) => {
                    diagnostics.push(unsupported_image_diagnostic(CONTENT_PART, error));
                    continue;
                }
            };
            reserve_materialized_image_bytes(
                materialized_image_bytes,
                bytes.len(),
                package.limits().max_total_uncompressed_bytes,
                CONTENT_PART,
            )?;
            if selected.is_none() {
                selected = Some((media_type, bytes, None));
                continue;
            }
            runtime_fallback = Some((media_type, bytes));
            break;
        }

        let Some(href) = image.href.as_deref() else {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Parse,
                    Fidelity::Omitted,
                    "draw:image has neither inline data nor an embedded package reference",
                )
                .in_part(CONTENT_PART),
            );
            continue;
        };
        if href.is_empty() {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Parse,
                    Fidelity::Omitted,
                    "draw:image has an empty package reference",
                )
                .in_part(CONTENT_PART),
            );
            continue;
        }
        let target = resolve_odf_image_target(href, package.limits().max_zip_path_bytes)
            .map_err(|message| format_error(CONTENT_PART, message))?;
        let OdfImageTarget::Embedded(target) = target else {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::ExternalResourceBlocked,
                    Phase::Security,
                    Fidelity::Blocked,
                    "external ODF picture reference was blocked",
                )
                .in_part(CONTENT_PART),
            );
            continue;
        };
        let declared_media_type = image
            .declared_media_type
            .as_deref()
            .or_else(|| manifest_media_types.get(&target).map(String::as_str));
        let cache_key = image_cache_key(&target, declared_media_type);
        let materialized = match image_cache.get(&cache_key).copied() {
            Some(ImageCacheEntry::Unsupported(error)) => {
                diagnostics.push(unsupported_image_diagnostic(&target, error));
                continue;
            }
            Some(ImageCacheEntry::Object(index)) => {
                let Some((media_type, bytes)) = objects
                    .get(index)
                    .and_then(|object| odp_image_visual_data(&object.visual))
                else {
                    return Err(format_error(
                        CONTENT_PART,
                        "embedded image cache is inconsistent",
                    ));
                };
                reserve_materialized_image_bytes(
                    materialized_image_bytes,
                    bytes.len(),
                    package.limits().max_total_uncompressed_bytes,
                    &target,
                )?;
                (
                    media_type.to_owned(),
                    clone_image_bytes(bytes, &target)?,
                    None,
                )
            }
            None => {
                let Some(bytes) = package.part(&target)? else {
                    diagnostics.push(
                        Diagnostic::warning(
                            DiagnosticCode::UnsupportedFeature,
                            Phase::Container,
                            Fidelity::Omitted,
                            "embedded ODF image alternative is missing from the package",
                        )
                        .in_part(&target),
                    );
                    continue;
                };
                let media_type = match identify_odf_image(&target, declared_media_type, &bytes) {
                    Ok(media_type) => media_type.to_owned(),
                    Err(error) => {
                        image_cache.insert(cache_key, ImageCacheEntry::Unsupported(error));
                        diagnostics.push(unsupported_image_diagnostic(&target, error));
                        continue;
                    }
                };
                reserve_materialized_image_bytes(
                    materialized_image_bytes,
                    bytes.len(),
                    package.limits().max_total_uncompressed_bytes,
                    &target,
                )?;
                (media_type, bytes.into_vec(), Some(cache_key))
            }
        };
        if selected.is_none() {
            selected = Some(materialized);
        } else {
            runtime_fallback = Some((materialized.0, materialized.1));
            break;
        }
    }
    let Some((media_type, bytes, cache_target)) = selected else {
        return Ok(());
    };
    if objects.len() >= package.limits().max_document_objects {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "document exceeds the configured object limit",
        )
        .in_part(CONTENT_PART));
    }
    let object_index = objects.len();
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    let mapping = if frame.element_id.is_some() {
        MappingQuality::Exact
    } else {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Approximate,
                "picture frame has no source ID; a structural mapping was derived",
            )
            .in_part(CONTENT_PART),
        );
        MappingQuality::Derived
    };
    let path = format!(
        "/office:document-content/office:body/office:presentation/draw:page[{}]/draw:frame[{}]/draw:image[1]",
        unit_index + 1,
        frame.frame_index + 1,
    );
    let frame_style = frame
        .style_name
        .as_ref()
        .and_then(|name| styles.graphics.get(name))
        .cloned()
        .unwrap_or_default();
    let crop = normalized_odf_image_crop(frame_style.image_clip, frame.bounds);
    let visual = match runtime_fallback {
        Some((fallback_media_type, fallback_bytes)) => Visual::ImageWithFallback {
            media_type,
            bytes,
            fallback_media_type,
            fallback_bytes,
            crop,
        },
        None => Visual::Image {
            media_type,
            bytes,
            crop,
        },
    };
    let parent_numeric_id = frame.parent_numeric_id;
    objects.push(Object {
        numeric_id,
        parent_numeric_id,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
        kind: ObjectKind::Image,
        unit_index,
        bounds: frame.bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping,
            locator: SourceLocator::OdpElement {
                element_id: frame.element_id.clone(),
                path: path.clone(),
                row: None,
                column: None,
            },
        },
        visual: layer_visual(visual, frame.transform, 1.0),
    });
    if let Some(target) = cache_target {
        image_cache.insert(target, ImageCacheEntry::Object(object_index));
    }
    push_image_frame_stroke(
        frame,
        unit_index,
        &path,
        mapping,
        &frame_style,
        styles,
        objects,
        package.limits().max_document_objects,
    )?;
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn attach_media_frame(
    frame: &FrameState,
    unit_index: u32,
    package: &dyn OdpSource,
    objects: &mut Vec<Object>,
    diagnostics: &mut Vec<Diagnostic>,
    image_resources: &mut ImageResources,
    poster_index: usize,
) -> Result<(), Diagnostic> {
    let Some(media) = frame.media.as_ref() else {
        return Ok(());
    };
    let target = resolve_odf_image_target(&media.href, package.limits().max_zip_path_bytes)
        .map_err(|message| format_error(CONTENT_PART, message))?;
    let OdfImageTarget::Embedded(target) = target else {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ExternalResourceBlocked,
                Phase::Security,
                Fidelity::Blocked,
                "external ODF audio/video reference was blocked",
            )
            .in_part(CONTENT_PART),
        );
        return Ok(());
    };
    let Some(bytes) = package.part(&target)? else {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Container,
                Fidelity::Omitted,
                "embedded ODF audio/video package part is missing",
            )
            .in_part(&target),
        );
        return Ok(());
    };
    let declared_media_type = media.declared_media_type.as_deref().or_else(|| {
        image_resources
            .manifest_media_types
            .get(&target)
            .map(String::as_str)
    });
    let identified = declared_media_type.map_or_else(
        || embedded_media_type(&target, &bytes, None),
        |media_type| {
            embedded_media_type_from_mime(media_type, &bytes).or_else(|error| {
                if error == EmbeddedMediaError::UnsupportedFormat {
                    embedded_media_type(&target, &bytes, None)
                } else {
                    Err(error)
                }
            })
        },
    );
    let (kind, media_type) = match identified {
        Ok(identified) => identified,
        Err(error) => {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Parse,
                    Fidelity::Omitted,
                    match error {
                        EmbeddedMediaError::UnsupportedFormat => {
                            "embedded ODF audio/video format is outside the browser-playable allowlist"
                        }
                        EmbeddedMediaError::SignatureMismatch => {
                            "embedded ODF audio/video bytes do not match the declared media format"
                        }
                    },
                )
                .in_part(&target),
            );
            return Ok(());
        }
    };
    reserve_materialized_image_bytes(
        &mut image_resources.materialized_bytes,
        bytes.len(),
        package.limits().max_total_uncompressed_bytes,
        &target,
    )?;
    let bytes = bytes.into_vec();

    if let Some(poster) = objects.get_mut(poster_index)
        && poster.unit_index == unit_index
        && poster.kind == ObjectKind::Image
        && poster.bounds == frame.bounds
    {
        let visual = core::mem::replace(&mut poster.visual, Visual::None);
        poster.visual = Visual::Media {
            kind,
            media_type: media_type.to_owned(),
            bytes,
            poster: Box::new(visual),
        };
        return Ok(());
    }

    if objects.len() >= package.limits().max_document_objects {
        return Err(object_limit_error());
    }
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    let parent_numeric_id = frame.parent_numeric_id;
    let mapping = if frame.element_id.is_some() {
        MappingQuality::Exact
    } else {
        MappingQuality::Derived
    };
    let path = format!(
        "/office:document-content/office:body/office:presentation/draw:page[{}]/draw:frame[{}]/draw:plugin[1]",
        unit_index + 1,
        frame.frame_index + 1,
    );
    let visual = Visual::Media {
        kind,
        media_type: media_type.to_owned(),
        bytes,
        poster: Box::new(Visual::None),
    };
    objects.push(Object {
        numeric_id,
        parent_numeric_id,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
        kind: ObjectKind::Image,
        unit_index,
        bounds: frame.bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping,
            locator: SourceLocator::OdpElement {
                element_id: frame.element_id.clone(),
                path,
                row: None,
                column: None,
            },
        },
        visual: layer_visual(visual, frame.transform, 1.0),
    });
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn push_image_frame_stroke(
    frame: &FrameState,
    unit_index: u32,
    image_path: &str,
    mapping: MappingQuality,
    style: &GraphicStyle,
    styles: &StyleCatalog,
    objects: &mut Vec<Object>,
    object_limit: usize,
) -> Result<(), Diagnostic> {
    if style.stroke_width <= 0.0 {
        return Ok(());
    }
    let stroke = resolve_odf_paint(&style.stroke, frame.bounds, styles);
    if matches!(stroke, Paint::None) {
        return Ok(());
    }
    if objects.len() >= object_limit {
        return Err(object_limit_error());
    }
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    let parent_numeric_id = frame.parent_numeric_id;
    let visual = Visual::PaintedShape {
        geometry: Geometry::Rectangle,
        fill: Paint::None,
        stroke,
        stroke_width: style.stroke_width,
    };
    let visual = odf_stroke_style_visual(visual, style, styles, false);
    objects.push(Object {
        numeric_id,
        parent_numeric_id,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
        kind: ObjectKind::Shape,
        unit_index,
        bounds: frame.bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: None,
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping,
            locator: SourceLocator::OdpElement {
                element_id: frame.element_id.clone(),
                path: image_path
                    .strip_suffix("/draw:image[1]")
                    .unwrap_or(image_path)
                    .to_owned(),
                row: None,
                column: None,
            },
        },
        visual: layer_visual(visual, frame.transform, style.opacity),
    });
    Ok(())
}

fn odp_image_visual_data(visual: &Visual) -> Option<(&str, &[u8])> {
    match visual {
        Visual::Image {
            media_type, bytes, ..
        }
        | Visual::ImageWithFallback {
            media_type, bytes, ..
        } => Some((media_type, bytes)),
        Visual::Layer { visual, .. }
        | Visual::Effect { visual, .. }
        | Visual::TextLayout { visual, .. }
        | Visual::StrokeStyle { visual, .. }
        | Visual::AdvancedEffect { visual, .. }
        | Visual::ImageColorChange { visual, .. } => odp_image_visual_data(visual),
        Visual::Media { poster, .. } => odp_image_visual_data(poster),
        _ => None,
    }
}

fn unsupported_image_diagnostic(part: &str, error: OfficeImageError) -> Diagnostic {
    let (phase, fidelity, message) = match error {
        OfficeImageError::UnsupportedFormat => (
            Phase::Parse,
            Fidelity::Omitted,
            "picture format is outside the supported modern Office/Open XML image set",
        ),
        OfficeImageError::SignatureMismatch => (
            Phase::Parse,
            Fidelity::Omitted,
            "picture bytes do not match the declared image format",
        ),
        OfficeImageError::DisabledByOffice => (
            Phase::Security,
            Fidelity::Blocked,
            "DisabledByOffice: EPS and legacy PICT picture formats are disabled by current Microsoft Office",
        ),
    };
    Diagnostic::warning(DiagnosticCode::UnsupportedFeature, phase, fidelity, message).in_part(part)
}

fn relationship_limit_error(part: &str) -> Diagnostic {
    Diagnostic::fatal(
        DiagnosticCode::RelationshipLimit,
        Phase::Parse,
        None,
        "document exceeds the configured image-reference limit",
    )
    .in_part(part)
}

fn push_text_box(
    mut frame: FrameState,
    unit_index: u32,
    objects: &mut Vec<Object>,
    diagnostics: &mut Vec<Diagnostic>,
    styles: &StyleCatalog,
    object_limit: usize,
    font_metrics: &FontMetricTable,
    package: &dyn OdpSource,
    image_resources: &mut ImageResources,
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
    let numeric_id = u32::try_from(objects.len())
        .map_err(|_| format_error(CONTENT_PART, "object count exceeds supported range"))?;
    let mapping = if frame.element_id.is_some() {
        MappingQuality::Exact
    } else {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Approximate,
                "frame has no source ID; a structural mapping was derived",
            )
            .in_part(CONTENT_PART),
        );
        MappingQuality::Derived
    };
    let path = format!(
        "/office:document-content/office:body/office:presentation/draw:page[{}]/draw:frame[{}]",
        unit_index + 1,
        frame.frame_index + 1,
    );
    let style = frame
        .style_name
        .as_ref()
        .and_then(|name| styles.graphics.get(name))
        .cloned()
        .unwrap_or_default();
    let parent_numeric_id = frame.parent_numeric_id;
    let runs = odf_text_runs(
        &frame.text,
        &frame.text_runs,
        &style,
        styles,
        &frame.paragraph_list_layouts,
    );
    let text = std::mem::take(&mut frame.text);
    let line_height = odf_text_line_height(&style, &frame.paragraph_style_names, &runs, styles);
    let layout = odf_text_layout(
        &style,
        &frame.paragraph_style_names,
        &frame.paragraph_list_layouts,
        &runs,
        line_height,
        true,
        styles,
    );
    if frame.auto_height {
        let mut glyphs = Vec::new();
        let mut heights = Vec::new();
        let mut approximate = false;
        for run in &runs {
            let height = font_metrics.line_height_em(&run.font_family, run.italic, run.bold, 1.0);
            for character in run.text.chars() {
                let advance = font_metrics.advance_em_at_size(
                    &run.font_family,
                    run.italic,
                    run.bold,
                    character,
                    run.font_size,
                );
                approximate |= height.is_none() || advance.is_none();
                // ponytail: missing font metrics use em estimates; injected font metrics
                // improve sizing without introducing a separate ODF line-breaking algorithm.
                glyphs.push((
                    character,
                    advance.unwrap_or(0.6) * run.font_size * run.horizontal_scale
                        + run.letter_spacing,
                ));
                heights.push(if line_height > 0.0 {
                    line_height
                } else {
                    height.unwrap_or(1.2) * run.font_size
                });
            }
        }
        let width = (frame.bounds.width - layout.inset_left - layout.inset_right).max(1.0);
        let breaks = crate::text_layout::horizontal_line_breaks(
            &glyphs,
            width,
            width,
            layout.default_tab_stop,
            false,
        );
        let mut start = 0;
        let mut height = layout.inset_top + layout.inset_bottom;
        for end in breaks {
            height += heights[start..end].iter().copied().fold(1.0, f32::max);
            start = end;
        }
        height += layout
            .paragraphs
            .iter()
            .map(|p| p.space_before + p.space_after)
            .sum::<f32>();
        frame.bounds.height = height.max(1.0);
        if approximate {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::ApproximateLayout,
                    Phase::Layout,
                    Fidelity::Approximate,
                    "automatic ODF text-frame height uses fallback font metrics",
                )
                .in_part(CONTENT_PART),
            );
        }
    }
    let mut paragraph_glyphs = vec![Vec::new()];
    let mut paragraph_fonts = vec![style.font_size];
    for run in &runs {
        for ch in run.text.chars() {
            if ch == '\n' {
                paragraph_glyphs.push(Vec::new());
                paragraph_fonts.push(run.font_size);
            } else {
                let advance = font_metrics
                    .advance_em_at_size(&run.font_family, run.italic, run.bold, ch, run.font_size)
                    .unwrap_or(0.6);
                paragraph_glyphs.last_mut().unwrap().push((
                    ch,
                    advance * run.font_size * run.horizontal_scale + run.letter_spacing,
                ));
                let size = paragraph_fonts.last_mut().unwrap();
                *size = size.max(run.font_size);
            }
        }
    }
    let paragraph_heights: Vec<_> = layout
        .paragraphs
        .iter()
        .enumerate()
        .map(|(index, p)| {
            let height = if p.line_height > 0.0 {
                p.line_height
            } else if line_height > 0.0 {
                line_height
            } else {
                paragraph_fonts
                    .get(index)
                    .copied()
                    .unwrap_or(style.font_size)
                    * 1.2
            };
            let width = (frame.bounds.width
                - layout.inset_left
                - layout.inset_right
                - p.margin_left
                - p.margin_right)
                .max(1.0);
            let count = paragraph_glyphs.get(index).map_or(1, |glyphs| {
                crate::text_layout::horizontal_line_breaks(
                    glyphs,
                    (width - p.first_line_indent).max(1.0),
                    width,
                    p.default_tab_stop,
                    false,
                )
                .len()
                .max(1)
            });
            (height, height * count as f32)
        })
        .collect();
    let fill = resolve_odf_paint(&style.fill, frame.bounds, styles);
    let stroke = resolve_odf_paint(&style.stroke, frame.bounds, styles);
    let rich_text = Visual::RichText {
        geometry: Geometry::Rectangle,
        fill,
        stroke,
        stroke_width: style.stroke_width,
        align: odf_text_align(&style, &frame.paragraph_style_names, styles),
        line_height,
        runs,
    };
    let visual = Visual::TextLayout {
        layout: layout.clone(),
        visual: Box::new(rich_text),
    };
    objects.push(Object {
        numeric_id,
        parent_numeric_id,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
        kind: ObjectKind::TextBox,
        unit_index,
        bounds: frame.bounds,
        z: i32::try_from(numeric_id).unwrap_or(i32::MAX),
        text: Some(text),
        source: SourceRef {
            part: CONTENT_PART.to_owned(),
            mapping,
            locator: SourceLocator::OdpElement {
                element_id: frame.element_id.clone(),
                path,
                row: None,
                column: None,
            },
        },
        visual: layer_visual(
            odf_stroke_style_visual(visual, &style, styles, false),
            frame.transform,
            style.opacity,
        ),
    });
    // ODF image labels retain their authored rectangular size, unlike square PPTX bullets.
    // Reuse the normal image path for package resolution, resource limits and SVG handling.
    let bounds = frame.bounds;
    let mut paragraph_y = bounds.y + layout.inset_top;
    let mut placements = Vec::new();
    for (index, paragraph) in layout.paragraphs.iter().enumerate() {
        let (height, paragraph_height) = paragraph_heights[index];
        paragraph_y += paragraph.space_before;
        if let Some((href, width, image_height)) = frame
            .paragraph_list_layouts
            .get(index)
            .and_then(Option::as_ref)
            .and_then(|p| p.image.as_ref())
        {
            if *width > 0.0 && *image_height > 0.0 {
                placements.push((
                    href.clone(),
                    Rect {
                        x: bounds.x
                            + layout.inset_left
                            + paragraph.margin_left
                            + paragraph.first_line_indent,
                        y: paragraph_y + (height - image_height) / 2.0,
                        width: *width,
                        height: *image_height,
                    },
                ));
            }
        }
        paragraph_y += paragraph_height + paragraph.space_after;
    }
    let content_height = paragraph_y - bounds.y - layout.inset_top;
    let available_height = bounds.height - layout.inset_top - layout.inset_bottom;
    let offset = match layout.vertical_align {
        TextVerticalAlign::Top => 0.0,
        TextVerticalAlign::Center => ((available_height - content_height) / 2.0).max(0.0),
        TextVerticalAlign::Bottom => (available_height - content_height).max(0.0),
    };
    frame.style_name = None;
    for (href, mut bounds) in placements {
        bounds.y += offset;
        frame.bounds = bounds;
        frame.images = vec![ImageRepresentation {
            href: Some(href),
            declared_media_type: None,
            inline_base64: None,
        }];
        push_image_frame(
            &frame,
            unit_index,
            package,
            styles,
            objects,
            diagnostics,
            image_resources,
        )?;
    }
    Ok(())
}

fn required_length(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<f32, Diagnostic> {
    let value = required_attribute(attributes, name, part)?;
    parse_length(&value).ok_or_else(|| {
        format_error(
            part,
            format!("attribute {name} is not a supported ODF length"),
        )
    })
}

fn parse_length(value: &str) -> Option<f32> {
    let value = value.trim();
    let units = [
        ("in", CSS_PIXELS_PER_INCH),
        ("cm", CSS_PIXELS_PER_INCH / 2.54),
        ("mm", CSS_PIXELS_PER_INCH / 25.4),
        ("pt", CSS_PIXELS_PER_INCH / 72.0),
        ("pc", CSS_PIXELS_PER_INCH / 6.0),
        ("px", 1.0),
    ];
    units.iter().find_map(|(suffix, scale)| {
        value
            .strip_suffix(suffix)
            .and_then(|number| number.parse::<f32>().ok())
            .map(|number| number * scale)
            .filter(|number| number.is_finite())
    })
}

fn valid_page_size(size: PageSize) -> bool {
    size.width.is_finite() && size.height.is_finite() && size.width > 0.0 && size.height > 0.0
}

fn required_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<String, Diagnostic> {
    optional_attribute(attributes, name, part)?
        .ok_or_else(|| format_error(part, format!("element is missing {name}")))
}

fn odf_true_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<bool, Diagnostic> {
    Ok(optional_attribute(attributes, name, part)?
        .is_some_and(|value| matches!(value.trim(), "true" | "1")))
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

#[cfg(all(test, feature = "odf-formats"))]
mod tests {
    use super::{
        CONTENT_PART, CSS_PIXELS_PER_INCH, GraphicStyle, OdfImageClip, PageSize, StyleCatalog,
        normalized_odf_image_crop, odf_text_layout, odf_text_run, parse, parse_content, parse_flat,
        parse_length, parse_odf_master_transform, parse_odf_transform,
    };
    use crate::diagnostic::{DiagnosticCode, Fidelity, Phase};
    use crate::font_metrics::FontMetricTable;
    use crate::format::detect_and_parse_with_font_metrics;
    use crate::format::presentation_image::stored_zip;
    use crate::limits::Limits;
    use crate::model::{
        DocumentFormat, Geometry, MappingQuality, ObjectKind, Paint, PathCommand, PathFillMode,
        Rect, SourceLocator, TextAlign, TextAutoFit, TextOrientation, TextVerticalAlign, Visual,
    };
    use crate::package::Package;
    use std::collections::HashMap;

    fn single_advance_metric_table(family: &str, character: char, advance_em: f32) -> Vec<u8> {
        let family = family.as_bytes();
        let metrics_offset = 20 + 16 + family.len();
        let mut bytes = Vec::with_capacity(metrics_offset + 8);
        bytes.extend_from_slice(b"OVFM");
        bytes.extend_from_slice(&1_u16.to_le_bytes());
        bytes.extend_from_slice(&20_u16.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(&(metrics_offset as u32).to_le_bytes());
        bytes.extend_from_slice(&(family.len() as u16).to_le_bytes());
        bytes.push(0);
        bytes.push(4);
        bytes.extend_from_slice(&400_u16.to_le_bytes());
        bytes.extend_from_slice(&0_u16.to_le_bytes());
        bytes.extend_from_slice(&0_u32.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(family);
        bytes.extend_from_slice(&(character as u32).to_le_bytes());
        bytes.extend_from_slice(&advance_em.to_le_bytes());
        bytes
    }

    const PNG: &[u8] = b"\x89PNG\r\n\x1a\nimage";

    #[test]
    fn renders_legacy_math_through_presentation_adapter() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/math_OOo311.odt"),
            crate::limits::Limits::default(),
        )
        .unwrap();
        let math = package.required_part("Object 1/content.xml").unwrap();
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame svg:x="1in" svg:y="1in" svg:width="1in" svg:height="1in"><draw:object xlink:href="./Object 1"/></draw:frame>
          </draw:page></office:presentation></office:body></office:document-content>"#;
        let parsed =
            parse_content_fixture(&[(CONTENT_PART, content), ("Object 1/content.xml", &math)])
                .unwrap();
        let formula = parsed
            .objects
            .iter()
            .find(|o| o.text.as_deref() == Some("1+1"))
            .unwrap();
        assert_eq!(formula.source.part, "Object 1/content.xml");
    }

    #[test]
    fn supplied_fontwork_preserves_envelopes_in_pages_and_masters() {
        let source = include_str!("../../tests/fixtures/fit-to-size.fodp")
            .replace("style:name=\"Arial Black\"", "style:name=\"HeavyFace\"")
            .replace(
                "style:font-name=\"Arial Black\"",
                "style:font-name=\"HeavyFace\"",
            );
        let start = source.find("    <draw:custom-shape").unwrap();
        let end = source[start..].find("    <presentation:notes").unwrap() + start;
        let shapes = &source[start..end];
        let master = source.replace(shapes, "").replace(
            "</style:master-page>",
            &format!("{shapes}</style:master-page>"),
        );
        for content in [&source, &master] {
            let zipped = stored_zip(&[
                (
                    "mimetype",
                    b"application/vnd.oasis.opendocument.presentation",
                ),
                (CONTENT_PART, content.as_bytes()),
                (super::STYLES_PART, content.as_bytes()),
            ]);
            for bytes in [content.as_bytes(), zipped.as_slice()] {
                let document = detect_and_parse_with_font_metrics(
                    bytes,
                    Limits::default(),
                    &FontMetricTable::default(),
                )
                .unwrap()
                .unwrap();
                let fontwork: Vec<_> = document
                    .objects
                    .iter()
                    .filter(|o| o.text.as_deref().is_some_and(|t| t.contains("Fontwork")))
                    .collect();
                assert_eq!(fontwork.len(), 2);
                for object in fontwork {
                    let mut visual = &object.visual;
                    while let Visual::Layer { visual: inner, .. } = visual {
                        visual = inner;
                    }
                    let Visual::TextLayout { layout, visual } = visual else {
                        panic!("text layout")
                    };
                    assert_eq!(layout.warp.as_deref(), Some("text-envelope"));
                    let Visual::RichText {
                        geometry: Geometry::LayeredPath { layers },
                        fill: Paint::None,
                        runs,
                        ..
                    } = visual.as_ref()
                    else {
                        panic!("unpainted envelope")
                    };
                    assert_eq!(layers.len(), 2);
                    assert!(
                        runs.iter()
                            .filter(|r| !r.text.trim().is_empty())
                            .all(|r| r.font_family == "Arial Black")
                    );
                }
            }
        }
        let unsupported = source.replace(
            "draw:text-path-mode=\"shape\"",
            "draw:text-path-mode=\"path\"",
        );
        let document = parse_flat(
            unsupported.as_bytes(),
            Limits::default(),
            &FontMetricTable::default(),
        )
        .unwrap();
        assert!(
            document
                .diagnostics
                .iter()
                .any(|d| d.message.contains("unsupported ODF text path"))
        );
    }

    #[test]
    fn opens_masterless_presentations_in_flat_and_packaged_odf() {
        for source in [
            include_bytes!("../../tests/fixtures/empty.fodp").as_slice(),
            include_bytes!("../../tests/fixtures/draw-object-link.fodp").as_slice(),
            include_bytes!("../../tests/fixtures/draw-image-link.fodp").as_slice(),
            include_bytes!("../../tests/fixtures/linked_ole.fodp").as_slice(),
        ] {
            let zipped = stored_zip(&[
                (
                    "mimetype",
                    b"application/vnd.oasis.opendocument.presentation",
                ),
                (CONTENT_PART, source),
                (
                    "styles.xml",
                    b"<office:document-styles xmlns:office=\"office\"/>",
                ),
            ]);
            for bytes in [source, zipped.as_slice()] {
                let document = detect_and_parse_with_font_metrics(
                    bytes,
                    Limits::default(),
                    &FontMetricTable::default(),
                )
                .unwrap()
                .unwrap();
                assert!(!document.fatal);
                assert_eq!(document.format, Some(DocumentFormat::Odp));
                assert_eq!(document.units.len(), 1);
                assert!(document.units[0].width > 0.0 && document.units[0].height > 0.0);
                if source == include_bytes!("../../tests/fixtures/draw-object-link.fodp")
                    || source == include_bytes!("../../tests/fixtures/linked_ole.fodp")
                {
                    assert!(
                        document
                            .objects
                            .iter()
                            .any(|o| o.text.as_deref() == Some("Embedded object"))
                    );
                    assert!(
                        document
                            .diagnostics
                            .iter()
                            .any(|d| d.message.contains("static placeholder"))
                    );
                } else {
                    assert!(document.objects.is_empty());
                }
                assert!(
                    document
                        .diagnostics
                        .iter()
                        .any(|d| d.message.contains("default page size"))
                );
                if source == include_bytes!("../../tests/fixtures/draw-image-link.fodp") {
                    assert!(
                        document
                            .diagnostics
                            .iter()
                            .any(|d| d.code == DiagnosticCode::ExternalResourceBlocked)
                    );
                }
            }
        }
    }

    #[test]
    fn parses_oasis_3933_inline_shape_fill() {
        let document = parse_flat(
            include_bytes!("../../tests/fixtures/oasis-3933-fill-image.fodg"),
            Limits::default(),
            &FontMetricTable::default(),
        )
        .unwrap();
        assert_eq!(document.units.len(), 1);
        assert!(document.objects.iter().any(|object| matches!(
            &object.visual,
            Visual::PaintedShape { fill: Paint::Image { media_type, bytes, tile: true, .. }, .. }
                if media_type == "image/jpeg" && bytes.len() > 80_000
        )), "the real diamond must contain its embedded pebble image");
        let source = std::str::from_utf8(include_bytes!(
            "../../tests/fixtures/oasis-3933-fill-image.fodg"
        ))
        .unwrap();
        let broken = source.replacen("/9j/4AAQ", "!9j/4AAQ", 1);
        let degraded = parse_flat(
            broken.as_bytes(),
            Limits::default(),
            &FontMetricTable::default(),
        )
        .unwrap();
        assert_eq!(degraded.units.len(), 1);
        assert_eq!(degraded.objects.len(), document.objects.len());
        assert!(
            degraded
                .diagnostics
                .iter()
                .any(|d| d.message == "invalid inline fill image was omitted")
        );
    }

    #[test]
    fn parses_a_flat_odf_presentation_without_a_zip_container() {
        let bytes = br#"<?xml version="1.0" encoding="UTF-8"?>
          <office:document xmlns:office="office" xmlns:style="style" xmlns:fo="fo" xmlns:draw="draw"
            office:mimetype="application/vnd.oasis.opendocument.presentation">
            <office:scripts><office:script/></office:scripts>
            <office:automatic-styles>
              <style:page-layout style:name="PageLayout">
                <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
              </style:page-layout>
            </office:automatic-styles>
            <office:master-styles>
              <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
            </office:master-styles>
            <office:body><office:presentation>
              <draw:page draw:name="Slide 1" draw:master-page-name="Master"/>
            </office:presentation></office:body>
          </office:document>"#;

        let document = parse_flat(bytes, Limits::default(), &FontMetricTable::default())
            .expect("valid flat ODF presentation");

        assert_eq!(document.format, Some(DocumentFormat::Odp));
        assert_eq!(document.units.len(), 1);
        assert_eq!(document.units[0].width, 960.0);
        assert_eq!(document.units[0].height, 720.0);
        assert!(
            document
                .diagnostics
                .iter()
                .any(|diagnostic| diagnostic.code == DiagnosticCode::ActiveContentBlocked)
        );
    }

    #[test]
    fn converts_odf_lengths_to_css_pixels_at_96_dpi() {
        assert_eq!(parse_length("10in"), Some(960.0));
        assert_eq!(parse_length("7.5in"), Some(720.0));
        assert_eq!(parse_length("72pt"), Some(96.0));
        assert_eq!(parse_length("25.4mm"), Some(96.0));
        assert_eq!(parse_length("50%"), None);
    }

    #[test]
    fn parses_an_embedded_raster_picture_with_exact_source_mapping() {
        let parsed = parse_fixture(
            "Pictures/image1.png",
            Some(("Pictures/image1.png", PNG)),
            Limits::default(),
        )
        .expect("valid embedded picture");

        assert_eq!(parsed.objects.len(), 1);
        let picture = &parsed.objects[0];
        assert_eq!(picture.kind, ObjectKind::Image);
        assert_eq!(picture.source.mapping, MappingQuality::Exact);
        assert!(matches!(
            &picture.source.locator,
            SourceLocator::OdpElement {
                element_id: Some(id),
                path,
                ..
            } if id == "picture-1" && path.ends_with("draw:frame[1]/draw:image[1]")
        ));
        assert!(matches!(
            &picture.visual,
            Visual::Image {
                media_type, bytes, ..
            }
                if media_type == "image/png" && bytes == PNG
        ));
        assert!(parsed.diagnostics.is_empty());
    }

    #[test]
    fn applies_odf_picture_clip_margins_as_normalized_crop() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <style:style style:name="CroppedPicture" style:family="graphic">
              <style:graphic-properties fo:clip="rect(0.5in, 1in, 0.5in, 1in)"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame draw:id="picture" draw:style-name="CroppedPicture"
              svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in">
              <draw:image xlink:href="Pictures/image.png"/>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[
            ("styles.xml", styles),
            (CONTENT_PART, content),
            ("Pictures/image.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("cropped ODF picture");

        let Visual::Image { crop, .. } = &document.objects[0].visual else {
            panic!("picture must retain its normalized source crop");
        };
        assert_eq!(crop.left, 0.25);
        assert_eq!(crop.top, 0.25);
        assert_eq!(crop.right, 0.25);
        assert_eq!(crop.bottom, 0.25);
    }

    #[test]
    fn treats_oversized_odf_picture_clip_edges_as_source_extents() {
        let crop = normalized_odf_image_crop(
            Some(OdfImageClip {
                left: 0.0,
                top: 0.0,
                right: 0.0,
                bottom: 6.631_47 * CSS_PIXELS_PER_INCH,
            }),
            Rect {
                x: 0.0,
                y: 0.0,
                width: 5.625 * CSS_PIXELS_PER_INCH,
                height: 4.702_38 * CSS_PIXELS_PER_INCH,
            },
        );

        assert_eq!(crop.left, 0.0);
        assert_eq!(crop.top, 0.0);
        assert_eq!(crop.right, 0.0);
        assert!((crop.bottom - 0.290_899_3).abs() < 0.000_01);
    }

    #[test]
    fn renders_extensionless_odf_images_by_signature_without_implicit_crop() {
        let mut wmf = vec![0_u8; 18];
        wmf[..2].copy_from_slice(&1_u16.to_le_bytes());
        wmf[2..4].copy_from_slice(&9_u16.to_le_bytes());
        wmf[4..6].copy_from_slice(&0x0300_u16.to_le_bytes());
        wmf[6..10].copy_from_slice(&9_u32.to_le_bytes());
        wmf[12..14].copy_from_slice(&1_u16.to_le_bytes());
        let parsed = parse_fixture(
            "ObjectReplacements/Object 1",
            Some(("ObjectReplacements/Object 1", &wmf)),
            Limits::default(),
        )
        .expect("extensionless WMF picture");
        assert!(matches!(
            &parsed.objects[0].visual,
            Visual::Image { media_type, crop, .. }
                if media_type == "image/x-wmf" && *crop == Default::default()
        ));
    }

    #[test]
    fn preserves_authored_odf_picture_frame_stroke() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:fo="fo" xmlns:svg="svg">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <style:style style:name="FramedPicture" style:family="graphic">
              <style:graphic-properties draw:fill="none" draw:stroke="solid"
                svg:stroke-width="0.03125in" svg:stroke-color="#eeaf78"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame draw:id="picture" draw:style-name="FramedPicture"
              svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in">
              <draw:image xlink:href="Pictures/image.png"/>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[
            ("styles.xml", styles),
            (CONTENT_PART, content),
            ("Pictures/image.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("framed ODF picture");

        assert_eq!(document.objects.len(), 2);
        assert_eq!(document.objects[0].kind, ObjectKind::Image);
        assert_eq!(document.objects[1].kind, ObjectKind::Shape);
        assert!(matches!(
            &document.objects[1].visual,
            Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: Paint::None,
                stroke: Paint::Solid(0xeeaf_78ff),
                stroke_width,
            } if (*stroke_width - 3.0).abs() < 0.001
        ));
    }

    #[test]
    fn prefers_inline_binary_image_data_and_ignores_its_href() {
        let content = content_xml_with_image(
            r#"<draw:image draw:mime-type="image/png" xlink:href="../missing.png"><office:binary-data>
              iVBORw0KGgppbWFnZQ==
            </office:binary-data></draw:image>"#,
        );
        let parsed = parse_content_fixture(&[(CONTENT_PART, content.as_bytes())])
            .expect("inline image data must not consult its href");

        assert_eq!(parsed.objects.len(), 1);
        assert!(matches!(
            &parsed.objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/png" && bytes == PNG
        ));
        assert!(parsed.diagnostics.is_empty());
    }

    #[test]
    fn skips_an_empty_image_reference_without_rejecting_the_presentation() {
        let parsed = parse_fixture("", None, Limits::default())
            .expect("an empty picture reference is an isolated resource failure");

        assert!(parsed.objects.is_empty());
        assert_eq!(parsed.diagnostics.len(), 1);
        assert_eq!(
            parsed.diagnostics[0].code,
            DiagnosticCode::UnsupportedFeature
        );
        assert!(parsed.diagnostics[0].message.contains("empty"));
    }

    #[test]
    fn resolves_images_from_any_safe_relative_package_path() {
        let parsed = parse_fixture(
            "Assets/branding/image1.png",
            Some(("Assets/branding/image1.png", PNG)),
            Limits::default(),
        )
        .expect("ODF images are not restricted to Pictures/");

        assert_eq!(parsed.objects.len(), 1);
        assert!(matches!(
            &parsed.objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/png" && bytes == PNG
        ));
    }

    #[test]
    fn uses_draw_mime_type_for_extensionless_embedded_images() {
        let content = content_xml_with_image(
            r#"<draw:image draw:mime-type="image/png" xlink:href="Assets/image-data"/>"#,
        );
        let parsed = parse_content_fixture(&[
            (CONTENT_PART, content.as_bytes()),
            ("Assets/image-data", PNG),
        ])
        .expect("draw:mime-type declares an extensionless image");

        assert_eq!(parsed.objects.len(), 1);
        assert!(matches!(
            &parsed.objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/png" && bytes == PNG
        ));
    }

    #[test]
    fn uses_manifest_media_type_and_still_checks_the_signature() {
        let content = content_xml_with_image(r#"<draw:image xlink:href="Assets/image-data"/>"#);
        let manifest = br#"<manifest:manifest xmlns:manifest="manifest">
          <manifest:file-entry manifest:full-path="Assets/image-data" manifest:media-type="image/png"/>
        </manifest:manifest>"#;
        let parsed = parse_content_fixture(&[
            (CONTENT_PART, content.as_bytes()),
            ("META-INF/manifest.xml", manifest),
            ("Assets/image-data", b"GIF89anot-a-png"),
        ])
        .expect("signature mismatch is a non-fatal omitted image");

        assert!(parsed.objects.is_empty());
        assert_eq!(parsed.diagnostics.len(), 1);
        assert_eq!(
            parsed.diagnostics[0].code,
            DiagnosticCode::UnsupportedFeature
        );
        assert!(parsed.diagnostics[0].message.contains("declared"));
    }

    #[test]
    fn uses_manifest_media_type_for_extensionless_embedded_images() {
        let content = content_xml_with_image(r#"<draw:image xlink:href="Assets/image-data"/>"#);
        let manifest = br#"<manifest:manifest xmlns:manifest="manifest">
          <manifest:file-entry manifest:full-path="Assets/image-data" manifest:media-type="image/png"/>
        </manifest:manifest>"#;
        let parsed = parse_content_fixture(&[
            (CONTENT_PART, content.as_bytes()),
            ("META-INF/manifest.xml", manifest),
            ("Assets/image-data", PNG),
        ])
        .expect("manifest media type declares an extensionless image");

        assert_eq!(parsed.objects.len(), 1);
        assert!(matches!(
            &parsed.objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/png" && bytes == PNG
        ));
    }

    #[test]
    fn falls_back_to_the_first_recognizable_draw_image_representation() {
        let content = content_xml_with_image(
            r#"<draw:image xlink:href="Assets/unsupported.avif"/>
               <draw:image xlink:href="Assets/fallback.png"/>"#,
        );
        let parsed = parse_content_fixture(&[
            (CONTENT_PART, content.as_bytes()),
            ("Assets/unsupported.avif", b"unsupported"),
            ("Assets/fallback.png", PNG),
        ])
        .expect("a later recognizable draw:image is a parser fallback");

        assert_eq!(parsed.objects.len(), 1);
        assert!(matches!(
            &parsed.objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/png" && bytes == PNG
        ));
        assert_eq!(parsed.diagnostics.len(), 1);
        assert_eq!(
            parsed.diagnostics[0].code,
            DiagnosticCode::UnsupportedFeature
        );
    }

    #[test]
    fn falls_back_when_an_earlier_draw_image_part_is_missing() {
        let content = content_xml_with_image(
            r#"<draw:image xlink:href="Assets/missing.png"/>
               <draw:image xlink:href="Assets/fallback.png"/>"#,
        );
        let parsed = parse_content_fixture(&[
            (CONTENT_PART, content.as_bytes()),
            ("Assets/fallback.png", PNG),
        ])
        .expect("a missing alternative does not suppress a later embedded image");

        assert_eq!(parsed.objects.len(), 1);
        assert!(matches!(
            &parsed.objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/png" && bytes == PNG
        ));
        assert_eq!(parsed.diagnostics.len(), 1);
        assert!(parsed.diagnostics[0].message.contains("missing"));
    }

    #[test]
    fn preserves_the_second_recognizable_draw_image_for_runtime_fallback() {
        let content = content_xml_with_image(
            r#"<draw:image xlink:href="Assets/preferred.svg"/>
               <draw:image xlink:href="Assets/fallback.png"/>"#,
        );
        let parsed = parse_content_fixture(&[
            (CONTENT_PART, content.as_bytes()),
            ("Assets/preferred.svg", b"<svg/>"),
            ("Assets/fallback.png", PNG),
        ])
        .expect("two recognizable alternatives preserve a runtime fallback");

        assert_eq!(parsed.objects.len(), 1);
        assert!(matches!(
            &parsed.objects[0].visual,
            Visual::ImageWithFallback {
                media_type,
                bytes,
                fallback_media_type,
                fallback_bytes,
                ..
            } if media_type == "image/svg+xml"
                && bytes == b"<svg/>"
                && fallback_media_type == "image/png"
                && fallback_bytes == PNG
        ));
        assert!(parsed.diagnostics.is_empty());
    }

    #[test]
    fn blocks_external_and_traversing_picture_references_without_loading_them() {
        let parsed = parse_fixture("https://example.invalid/image.png", None, Limits::default())
            .expect("external image is blocked non-fatally");
        assert!(parsed.objects.is_empty());
        assert_eq!(
            parsed.diagnostics[0].code,
            DiagnosticCode::ExternalResourceBlocked
        );

        let parsed = parse_fixture("../Pictures/image1.png", None, Limits::default())
            .expect("relative external images are blocked without failing the document");
        assert!(parsed.objects.is_empty());
        assert_eq!(
            parsed.diagnostics[0].code,
            DiagnosticCode::ExternalResourceBlocked
        );
    }

    #[test]
    fn parses_an_embedded_svg_picture() {
        let parsed = parse_fixture(
            "Pictures/image1.svg",
            Some(("Pictures/image1.svg", b"<svg/>")),
            Limits::default(),
        )
        .expect("valid SVG picture");
        assert_eq!(parsed.objects.len(), 1);
        assert!(matches!(
            &parsed.objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/svg+xml" && bytes == b"<svg/>"
        ));
        assert!(parsed.diagnostics.is_empty());
    }

    #[test]
    fn renders_an_embedded_chart_when_its_cached_image_is_missing() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame svg:x="1in" svg:y="1in" svg:width="4in" svg:height="2in">
              <draw:object xlink:href="./Object 1"/>
              <draw:image xlink:href="./ObjectReplacements/Object 1"/>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let chart = br##"<office:document-content xmlns:office="office" xmlns:chart="chart" xmlns:draw="draw" xmlns:style="style" xmlns:table="table" xmlns:text="text">
          <office:automatic-styles>
            <style:style style:name="YellowSeries" style:family="chart">
              <style:graphic-properties draw:fill-color="#ffd320"/>
            </style:style>
            <style:style style:name="HiddenAxis" style:family="chart">
              <style:chart-properties chart:visible="false"/>
            </style:style>
          </office:automatic-styles>
          <office:body><office:chart><chart:chart chart:class="chart:bar">
            <chart:plot-area>
              <chart:axis chart:dimension="x">
                <chart:categories table:cell-range-address="local-table.$A$2:.$A$3"/>
              </chart:axis>
              <chart:axis chart:dimension="y" chart:style-name="HiddenAxis"/>
              <chart:series chart:values-cell-range-address="local-table.$B$2:.$B$3"
                chart:label-cell-address="local-table.$B$1" chart:style-name="YellowSeries"/>
              <table:table table:name="local-table">
                <table:table-row>
                  <table:table-cell/><table:table-cell office:value-type="string"><text:p>Column 1</text:p></table:table-cell>
                </table:table-row>
                <table:table-row>
                  <table:table-cell office:value-type="string"><text:p>Row 1</text:p></table:table-cell>
                  <table:table-cell office:value="2"/>
                </table:table-row>
                <table:table-row>
                  <table:table-cell office:value-type="string"><text:p>Row 2</text:p></table:table-cell>
                  <table:table-cell office:value="4"/>
                </table:table-row>
              </table:table>
            </chart:plot-area>
          </chart:chart></office:chart></office:body>
        </office:document-content>"##;
        let parsed =
            parse_content_fixture(&[(CONTENT_PART, content), ("Object 1/content.xml", chart)])
                .expect("a supported native chart must not depend on its missing cached preview");

        assert!(
            parsed
                .objects
                .iter()
                .any(|object| object.source.part == "Object 1/content.xml")
        );
        let chart_text = parsed
            .objects
            .iter()
            .filter(|object| object.source.part == "Object 1/content.xml")
            .filter_map(|object| object.text.as_deref())
            .collect::<Vec<_>>();
        assert!(chart_text.contains(&"Row 1"));
        assert!(chart_text.contains(&"Row 2"));
        assert!(!chart_text.contains(&"Column 1"));
        assert!(!chart_text.contains(&"0"));
        assert!(parsed.objects.iter().any(|object| {
            object.source.part == "Object 1/content.xml"
                && object.source.mapping == MappingQuality::Exact
                && matches!(
                    object.visual,
                    Visual::PaintedShape {
                        fill: Paint::Solid(0xffd3_20ff),
                        ..
                    }
                )
        }));
    }

    #[test]
    fn preserves_embedded_chart_stacking_plot_area_and_end_axes() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame svg:x="1in" svg:y="1in" svg:width="4in" svg:height="2in">
              <draw:object xlink:href="./Object 1"/>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let chart = br#"<office:document-content xmlns:office="office" xmlns:chart="chart" xmlns:chartooo="chartooo" xmlns:style="style" xmlns:svg="svg" xmlns:table="table">
          <office:automatic-styles>
            <style:style style:name="Stacked" style:family="chart"><style:chart-properties chart:stacked="true"/></style:style>
            <style:style style:name="End" style:family="chart"><style:chart-properties chart:axis-position="end"/></style:style>
          </office:automatic-styles>
          <office:body><office:chart><chart:chart chart:class="chart:bar" svg:width="400px" svg:height="200px">
            <chart:plot-area chart:style-name="Stacked">
              <chartooo:coordinate-region svg:x="100px" svg:y="40px" svg:width="240px" svg:height="120px"/>
              <chart:axis chart:dimension="x" chart:style-name="End"><chart:categories table:cell-range-address="local.$A$1"/></chart:axis>
              <chart:axis chart:dimension="y" chart:style-name="End"/>
              <chart:series chart:values-cell-range-address="local.$A$2"/>
              <chart:series chart:values-cell-range-address="local.$A$3"/>
              <table:table table:name="local">
                <table:table-row><table:table-cell office:string-value="Q1"/></table:table-row>
                <table:table-row><table:table-cell office:value="2"/></table:table-row>
                <table:table-row><table:table-cell office:value="3"/></table:table-row>
              </table:table>
            </chart:plot-area>
          </chart:chart></office:chart></office:body>
        </office:document-content>"#;
        let parsed =
            parse_content_fixture(&[(CONTENT_PART, content), ("Object 1/content.xml", chart)])
                .expect("stacked chart with end-positioned axes");
        let mut axes = parsed
            .objects
            .iter()
            .filter(|object| {
                object.source.part == "Object 1/content.xml"
                    && matches!(
                        object.visual,
                        Visual::PaintedShape {
                            stroke: Paint::Solid(0x6b72_80ff),
                            ..
                        }
                    )
            })
            .map(|object| object.bounds)
            .collect::<Vec<_>>();
        axes.sort_by(|left, right| left.width.total_cmp(&right.width));
        assert_eq!(axes.len(), 2);
        assert!((axes[0].x - 422.4).abs() < 0.01);
        assert!((axes[1].y - 134.4).abs() < 0.01);

        let bars = parsed
            .objects
            .iter()
            .filter(|object| {
                object.source.part == "Object 1/content.xml"
                    && matches!(
                        &object.source.locator,
                        SourceLocator::OdpElement {
                            path,
                            row: Some(_),
                            column: Some(0),
                            ..
                        } if path.contains("data-point")
                    )
            })
            .map(|object| object.bounds)
            .collect::<Vec<_>>();
        assert_eq!(bars.len(), 2);
        assert!((bars[0].x - bars[1].x).abs() < 0.01);
        assert!((bars[1].y + bars[1].height - bars[0].y).abs() < 0.01);
    }

    #[test]
    fn renders_embedded_chart_title_and_narrow_automatic_axis() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame svg:x="1in" svg:y="1in" svg:width="4in" svg:height="2.25in">
              <draw:object xlink:href="./Object 1"/>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let chart = br##"<office:document-content xmlns:office="office" xmlns:chart="chart" xmlns:draw="draw" xmlns:style="style" xmlns:table="table" xmlns:text="text">
          <office:automatic-styles>
            <style:style style:name="Chart" style:family="chart">
              <style:graphic-properties draw:stroke="none"/>
            </style:style>
            <style:style style:name="Wall" style:family="chart">
              <style:graphic-properties draw:fill="none" draw:fill-color="#e6e6e6"/>
            </style:style>
          </office:automatic-styles>
          <office:body><office:chart><chart:chart chart:class="chart:bar" chart:style-name="Chart">
            <chart:title><text:p>A lame chart</text:p></chart:title>
            <chart:plot-area>
              <chart:axis chart:dimension="x"><chart:categories table:cell-range-address="local.$B$1:.$D$1"/></chart:axis>
              <chart:axis chart:dimension="y"/>
              <chart:series chart:values-cell-range-address="local.$B$2:.$D$2" chart:label-cell-address="local.$A$2"/>
              <chart:wall chart:style-name="Wall"/>
              <table:table table:name="local">
                <table:table-row><table:table-cell/><table:table-cell><text:p>Oct, 13</text:p></table:table-cell><table:table-cell><text:p>Nov, 13</text:p></table:table-cell><table:table-cell><text:p>Dec, 13</text:p></table:table-cell></table:table-row>
                <table:table-row><table:table-cell><text:p>Size</text:p></table:table-cell><table:table-cell office:value="12"/><table:table-cell office:value="13"/><table:table-cell office:value="14"/></table:table-row>
              </table:table>
            </chart:plot-area>
          </chart:chart></office:chart></office:body>
        </office:document-content>"##;
        let parsed =
            parse_content_fixture(&[(CONTENT_PART, content), ("Object 1/content.xml", chart)])
                .expect("embedded ODP chart");
        let chart_objects = parsed
            .objects
            .iter()
            .filter(|object| object.source.part == "Object 1/content.xml")
            .collect::<Vec<_>>();
        let chart_text = chart_objects
            .iter()
            .filter_map(|object| object.text.as_deref())
            .collect::<Vec<_>>();

        assert!(chart_text.contains(&"A lame chart"));
        assert!(chart_text.contains(&"11"));
        assert!(chart_text.contains(&"11.5"));
        assert!(chart_text.contains(&"14.5"));
        assert!(!chart_text.contains(&"0"));
        assert!(parsed.objects.iter().any(|object| {
            object.source.part == CONTENT_PART
                && match &object.visual {
                    Visual::PaintedShape {
                        stroke: Paint::None,
                        ..
                    } => true,
                    Visual::Layer { visual, .. } => matches!(
                        visual.as_ref(),
                        Visual::PaintedShape {
                            stroke: Paint::None,
                            ..
                        }
                    ),
                    _ => false,
                }
        }));
        assert!(!chart_objects.iter().any(|object| {
            matches!(
                &object.source.locator,
                SourceLocator::OdpElement { path, .. } if path.ends_with("/chart:wall[1]")
            )
        }));
    }

    #[test]
    fn renders_an_odp_pie_chart_custom_data_label() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame svg:x="1in" svg:y="1in" svg:width="4in" svg:height="3in">
              <draw:object xlink:href="./Object 1"/>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let chart = br#"<office:document-content xmlns:office="office" xmlns:chart="chart" xmlns:loext="loext" xmlns:table="table" xmlns:text="text">
          <office:body><office:chart><chart:chart chart:class="chart:circle"><chart:plot-area>
            <chart:series chart:values-cell-range-address="local-table.$B$2:.$B$3">
              <chart:data-point/><chart:data-point loext:custom-label-field="Kiskacsa"/>
            </chart:series>
            <table:table table:name="local-table">
              <table:table-row><table:table-cell/><table:table-cell><text:p>Sales</text:p></table:table-cell></table:table-row>
              <table:table-row><table:table-cell><text:p>One</text:p></table:table-cell><table:table-cell office:value="1"/></table:table-row>
              <table:table-row><table:table-cell><text:p>Two</text:p></table:table-cell><table:table-cell office:value="2"/></table:table-row>
            </table:table>
          </chart:plot-area></chart:chart></office:chart></office:body>
        </office:document-content>"#;
        let parsed =
            parse_content_fixture(&[(CONTENT_PART, content), ("Object 1/content.xml", chart)])
                .expect("ODP pie chart with a custom data label");

        assert!(parsed.objects.iter().any(|object| {
            object.text.as_deref() == Some("Kiskacsa")
                && object.source.part == "Object 1/content.xml"
                && object.source.mapping == MappingQuality::Exact
        }));
    }

    #[test]
    fn renders_odp_pie_percentage_labels_and_explosion_from_chart_properties() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame svg:x="1in" svg:y="1in" svg:width="4in" svg:height="3in">
              <draw:object xlink:href="./Object 1"/>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let chart = br##"<office:document-content xmlns:office="office" xmlns:chart="chart" xmlns:style="style" xmlns:table="table" xmlns:text="text" xmlns:draw="draw">
          <office:automatic-styles>
            <style:style style:name="ch7" style:family="chart">
              <style:chart-properties chart:data-label-number="percentage" chart:data-label-text="false" chart:data-label-symbol="false" chart:label-position="outside"/>
            </style:style>
            <style:style style:name="ch10" style:family="chart">
              <style:chart-properties chart:pie-offset="29"/>
              <style:graphic-properties draw:fill-color="#ffd320"/>
            </style:style>
          </office:automatic-styles>
          <office:body><office:chart><chart:chart chart:class="chart:circle"><chart:plot-area>
            <chart:series chart:style-name="ch7" chart:values-cell-range-address="local-table.$B$2:.$B$3">
              <chart:data-point/><chart:data-point chart:style-name="ch10"/>
            </chart:series>
            <table:table table:name="local-table">
              <table:table-row><table:table-cell/><table:table-cell><text:p>y</text:p></table:table-cell></table:table-row>
              <table:table-row><table:table-cell><text:p>A</text:p></table:table-cell><table:table-cell office:value="5"/></table:table-row>
              <table:table-row><table:table-cell><text:p>B</text:p></table:table-cell><table:table-cell office:value="3"/></table:table-row>
            </table:table>
          </chart:plot-area></chart:chart></office:chart></office:body>
        </office:document-content>"##;
        let parsed =
            parse_content_fixture(&[(CONTENT_PART, content), ("Object 1/content.xml", chart)])
                .expect("ODP pie chart with shared data labels and explosion");

        let texts: Vec<_> = parsed
            .objects
            .iter()
            .filter_map(|object| object.text.as_deref())
            .collect();
        assert!(texts.contains(&"62.5%"), "labels {texts:?}");
        assert!(texts.contains(&"37.5%"), "labels {texts:?}");

        let slices: Vec<_> = parsed
            .objects
            .iter()
            .filter(|object| {
                matches!(
                    object.visual,
                    Visual::PaintedShape {
                        geometry: Geometry::Path { .. },
                        ..
                    }
                ) && object.source.part == "Object 1/content.xml"
            })
            .collect();
        assert_eq!(slices.len(), 2);
        let apex = |object: &crate::model::Object| match &object.visual {
            Visual::PaintedShape {
                geometry: Geometry::Path { commands, .. },
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
        let (ax, ay) = apex(&slices[0]);
        let (bx, by) = apex(&slices[1]);
        let delta = ((ax - bx).powi(2) + (ay - by).powi(2)).sqrt();
        assert!(delta > 1.0, "exploded apex moves away (delta={delta})");
    }

    #[test]
    fn renders_an_odp_chart_wall_transparency_gradient() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:xlink="xlink">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame svg:x="1in" svg:y="1in" svg:width="4in" svg:height="3in">
              <draw:object xlink:href="./Object 1"/>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let chart = br##"<office:document-content xmlns:office="office" xmlns:chart="chart" xmlns:draw="draw" xmlns:style="style" xmlns:table="table">
          <office:automatic-styles><style:style style:name="Wall" style:family="chart">
            <style:graphic-properties draw:fill="solid" draw:fill-color="#32cd32" draw:opacity-name="Fade"/>
          </style:style></office:automatic-styles>
          <office:body><office:chart><chart:chart chart:class="chart:bar"><chart:plot-area>
            <chart:series chart:values-cell-range-address="local-table.$A$1:.$A$1"/>
            <chart:wall chart:style-name="Wall"/>
            <table:table table:name="local-table"><table:table-row><table:table-cell office:value="1"/></table:table-row></table:table>
          </chart:plot-area></chart:chart></office:chart></office:body>
        </office:document-content>"##;
        let chart_styles = br#"<office:document-styles xmlns:office="office" xmlns:draw="draw">
          <office:styles><draw:opacity draw:name="Fade" draw:style="linear" draw:start="0%" draw:end="100%" draw:angle="0deg"/></office:styles>
        </office:document-styles>"#;
        let parsed = parse_content_fixture(&[
            (CONTENT_PART, content),
            ("Object 1/content.xml", chart),
            ("Object 1/styles.xml", chart_styles),
        ])
        .expect("ODP chart wall with a transparency gradient");

        assert!(parsed.objects.iter().any(|object| {
            object.source.part == "Object 1/content.xml"
                && matches!(
                    &object.source.locator,
                    SourceLocator::OdpElement { path, .. } if path.ends_with("/chart:wall[1]")
                )
                && matches!(
                    &object.visual,
                    Visual::PaintedShape {
                        fill: Paint::LinearGradient { stops, .. },
                        ..
                    } if stops == &vec![
                        crate::model::GradientStop { offset: 0.0, color: 0x32cd_32ff },
                        crate::model::GradientStop { offset: 1.0, color: 0x32cd_3200 },
                    ]
                )
        }));
    }

    #[test]
    fn explicitly_blocks_images_disabled_by_current_office() {
        let parsed = parse_fixture(
            "Pictures/image1.pict",
            Some(("Pictures/image1.pict", b"legacy PICT")),
            Limits::default(),
        )
        .expect("Office-disabled image is non-fatal");
        assert!(parsed.objects.is_empty());
        assert_eq!(parsed.diagnostics.len(), 1);
        assert_eq!(
            parsed.diagnostics[0].code,
            DiagnosticCode::UnsupportedFeature
        );
        assert_eq!(parsed.diagnostics[0].phase, Phase::Security);
        assert_eq!(parsed.diagnostics[0].fidelity, Fidelity::Blocked);
        assert!(
            parsed.diagnostics[0]
                .message
                .starts_with("DisabledByOffice:")
        );
    }

    #[test]
    fn reports_a_missing_embedded_picture_part() {
        let parsed = parse_fixture("Pictures/missing.png", None, Limits::default())
            .expect("missing package images are isolated");
        assert!(parsed.objects.is_empty());
        assert_eq!(parsed.diagnostics.len(), 1);
        assert_eq!(parsed.diagnostics[0].fidelity, Fidelity::Omitted);
        assert!(parsed.diagnostics[0].message.contains("missing"));
    }

    #[test]
    fn applies_the_object_limit_to_pictures() {
        let content = content_xml("Pictures/image1.png").replace(
            "</draw:page>",
            "<draw:frame draw:id=\"picture-2\" svg:x=\"2in\" svg:y=\"1in\" svg:width=\"1in\" svg:height=\"1in\"><draw:image xlink:href=\"Pictures/image1.png\"/></draw:frame></draw:page>",
        );
        let bytes = stored_zip(&[
            (CONTENT_PART, content.as_bytes()),
            ("Pictures/image1.png", PNG),
        ]);
        let package = Package::open(
            &bytes,
            Limits {
                max_document_objects: 1,
                ..Limits::default()
            },
        )
        .unwrap();
        let error = match parse_content(&package, &page_sizes(), &StyleCatalog::default()) {
            Ok(_) => panic!("second picture exceeds the object budget"),
            Err(error) => error,
        };
        assert_eq!(error.code, DiagnosticCode::ObjectLimit);
    }

    #[test]
    fn bounds_materialized_bytes_when_a_part_is_reused() {
        let content = content_xml("Pictures/image1.png").replace(
            "</draw:page>",
            "<draw:frame draw:id=\"picture-2\" svg:x=\"2in\" svg:y=\"1in\" svg:width=\"1in\" svg:height=\"1in\"><draw:image xlink:href=\"Pictures/image1.png\"/></draw:frame></draw:page>",
        );
        let mut image = vec![0_u8; 4_096];
        image[..PNG.len()].copy_from_slice(PNG);
        let bytes = stored_zip(&[
            (CONTENT_PART, content.as_bytes()),
            ("Pictures/image1.png", &image),
        ]);
        let package = Package::open(
            &bytes,
            Limits {
                max_entry_uncompressed_bytes: 6_000,
                max_total_uncompressed_bytes: 6_000,
                max_xml_bytes: 6_000,
                ..Limits::default()
            },
        )
        .unwrap();
        let error = match parse_content(&package, &page_sizes(), &StyleCatalog::default()) {
            Ok(_) => panic!("reusing a part must not bypass the materialized-byte budget"),
            Err(error) => error,
        };
        assert_eq!(error.code, DiagnosticCode::ZipTotalSizeLimit);
    }

    #[test]
    fn diagnoses_only_remaining_unsupported_presentation_objects() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:table="table" xmlns:chart="chart" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <table:table/><table:table/>
            <draw:frame svg:x="0in" svg:y="0in" svg:width="1in" svg:height="1in"><draw:object/></draw:frame>
            <draw:g/><draw:g/>
            <draw:custom-shape/><draw:path/>
            <draw:path svg:x="0in" svg:y="0in" svg:width="1in" svg:height="1in" svg:viewBox="0 0 1 1" svg:d="M0 0 L1 1"/>
            <draw:custom-shape/>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[(CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let parsed = parse_content(&package, &page_sizes(), &StyleCatalog::default())
            .expect("unsupported presentation objects are diagnosed non-fatally");

        let unsupported = parsed
            .diagnostics
            .iter()
            .filter(|diagnostic| diagnostic.code == DiagnosticCode::UnsupportedFeature)
            .collect::<Vec<_>>();
        assert_eq!(
            unsupported
                .iter()
                .filter(|diagnostic| diagnostic.message.contains("static placeholder"))
                .count(),
            1
        );
        assert_eq!(
            unsupported
                .iter()
                .filter(|diagnostic| diagnostic.message.contains("geometry"))
                .count(),
            3
        );
        assert!(unsupported.iter().all(|diagnostic| {
            diagnostic.message.contains("static placeholder")
                || diagnostic.message.contains("geometry")
        }));
    }

    #[test]
    fn positions_transform_only_frames_from_their_local_origin() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame draw:id="rotated-text" draw:transform="translate(2in 3in)" svg:width="1in" svg:height="0.5in">
              <draw:text-box><text:p>Transform only</text:p></draw:text-box>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let parsed = parse_content_fixture(&[(CONTENT_PART, content)])
            .expect("a transformed frame may use its local origin instead of svg:x and svg:y");

        assert_eq!(parsed.objects.len(), 1);
        let object = &parsed.objects[0];
        assert_eq!(object.bounds.x, 0.0);
        assert_eq!(object.bounds.y, 0.0);
        assert_eq!(object.bounds.width, 96.0);
        assert_eq!(object.bounds.height, 48.0);
        assert!(matches!(
            object.visual,
            Visual::Layer { transform, .. }
                if transform.e == 192.0 && transform.f == 288.0
        ));
    }

    #[test]
    fn supplied_odf_preset_shapes_fit_the_scene_protocol() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/odf-preset-shapes.odp"),
            Limits::default(),
        )
        .unwrap();
        let mut document = parse(&package).expect("preset shapes parse");
        let objects = std::mem::take(&mut document.objects);
        for object in objects {
            document.objects = vec![object];
            let result = crate::protocol::encode(&document);
            assert!(
                result.is_ok(),
                "each actual preset shape must cross the scene protocol: {:?}",
                result.err()
            );
        }
    }

    #[test]
    fn supplied_odf_invalid_page_size_keeps_a_usable_slide() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/odf-zero-page-size.odp"),
            Limits::default(),
        )
        .unwrap();
        let document =
            parse(&package).expect("an invalid page layout must use the normal fallback");
        assert!(!document.units.is_empty());
        assert!(!document.objects.is_empty());
        assert!(
            document
                .diagnostics
                .iter()
                .any(|d| d.code == DiagnosticCode::ApproximateLayout)
        );
    }

    #[cfg(feature = "odf-formats")]
    #[test]
    fn supplied_odf_auto_height_preserves_empty_paragraphs() {
        let document = super::parse_flat(
            include_bytes!("../../tests/fixtures/odf-auto-height.fodp"),
            Limits::default(),
            &FontMetricTable::default(),
        )
        .unwrap();
        let text = document
            .objects
            .iter()
            .find(|o| o.text.as_deref() == Some("a\n\nc"))
            .expect("all three paragraphs survive");
        assert!(text.bounds.height > 30.0);
        crate::protocol::encode(&document).unwrap();
    }

    #[test]
    fn supplied_odf_without_optional_styles_keeps_text() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/odf-no-styles.odp"),
            Limits::default(),
        )
        .unwrap();
        let document = parse(&package).expect("styles.xml is optional");
        assert_eq!(document.units.len(), 1);
        assert!(
            document
                .objects
                .iter()
                .any(|object| object.text.as_deref() == Some("Test"))
        );
    }

    #[test]
    fn supplied_odf_image_failures_preserve_other_content() {
        for bytes in [
            include_bytes!("../../tests/fixtures/odf-missing-image.odp").as_slice(),
            include_bytes!("../../tests/fixtures/odf-linked-graphic.odp").as_slice(),
        ] {
            let package = Package::open(bytes, Limits::default()).unwrap();
            let document = parse(&package).expect("an unavailable image is a local failure");
            assert!(!document.objects.is_empty());
            assert!(!document.diagnostics.is_empty());
        }
    }

    #[test]
    fn supplied_odf_single_axis_mirrors_preserve_native_rotation() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/odf-rotate-flip.odp"),
            Limits::default(),
        )
        .unwrap();
        let document = parse(&package).unwrap();
        for (object, positive) in document.objects.iter().zip([false, false, false]) {
            let Visual::Layer { transform, .. } = object.visual else {
                panic!("authored transform must survive")
            };
            assert_eq!(
                transform.b > 0.0,
                positive,
                "native arrow rotation for {}",
                object.stable_id
            );
        }
    }

    #[test]
    fn supplied_odf_skew_transforms_keep_shapes_renderable() {
        for bytes in [
            include_bytes!("../../tests/fixtures/odf-rotate-flip.odp").as_slice(),
            include_bytes!("../../tests/fixtures/odf-shear-flip.odg").as_slice(),
        ] {
            let package = Package::open(bytes, Limits::default()).unwrap();
            let document = parse(&package).expect("ODF skew is a supported affine transform");
            assert!(!document.objects.is_empty());
        }
    }

    #[test]
    fn applies_odf_transform_functions_in_document_order() {
        let transform =
            parse_odf_transform("translate(-3in -4in) rotate(-1.57079632679) translate(3in 4in)")
                .expect("rotation around a translated center");
        let x = transform.a * (4.0 * 96.0) + transform.c * (5.0 * 96.0) + transform.e;
        let y = transform.b * (4.0 * 96.0) + transform.d * (5.0 * 96.0) + transform.f;

        assert!((x - 2.0 * 96.0).abs() < 0.001);
        assert!((y - 5.0 * 96.0).abs() < 0.001);
    }

    #[test]
    fn maps_odf_rotation_to_the_canvas_coordinate_system() {
        let transform = parse_odf_transform("rotate(-1.57079632679)").expect("valid ODF rotation");
        let x = transform.a;
        let y = transform.b;

        assert!(x.abs() < 0.001);
        assert!((y - 1.0).abs() < 0.001);
    }

    #[test]
    fn preserves_master_shape_rotation_direction() {
        let transform = parse_odf_master_transform("rotate(-1.57079632679)")
            .expect("valid master shape rotation");

        assert!(transform.a.abs() < 0.001);
        assert!((transform.b + 1.0).abs() < 0.001);
    }

    #[test]
    fn renders_master_images_and_excludes_master_placeholders_and_notes() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:presentation="presentation" xmlns:svg="svg" xmlns:text="text" xmlns:xlink="xlink" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <style:style style:name="FooterText" style:family="text">
              <style:text-properties style:font-family="Segoe UI" fo:font-size="10pt" fo:color="#808080"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout">
              <draw:frame draw:id="banner" svg:x="0in" svg:y="0in" svg:width="10in" svg:height="1in">
                <draw:image xlink:href="Pictures/banner.png"/>
              </draw:frame>
              <draw:frame draw:id="title-placeholder" presentation:class="title" svg:x="0in" svg:y="0in" svg:width="10in" svg:height="1in">
                <draw:text-box><text:p>INSERT TITLE HERE</text:p></draw:text-box>
              </draw:frame>
              <draw:frame draw:id="footer" svg:x="0.2in" svg:y="7in" svg:width="4in" svg:height="0.3in">
                <draw:text-box><text:p><text:span text:style-name="FooterText">Footer</text:span></text:p></draw:text-box>
              </draw:frame>
              <presentation:notes>
                <draw:frame draw:id="notes-date" svg:x="0in" svg:y="0in" svg:width="1in" svg:height="0.3in">
                  <draw:text-box><text:p>2/2/2026</text:p></draw:text-box>
                </draw:frame>
              </presentation:notes>
            </style:master-page>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw">
          <office:body><office:presentation><draw:page draw:master-page-name="Master"/></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[
            ("styles.xml", styles),
            (CONTENT_PART, content),
            ("Pictures/banner.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("valid master background");

        assert_eq!(document.objects.len(), 2);
        assert_eq!(document.objects[0].kind, ObjectKind::Image);
        assert!(matches!(
            &document.objects[0].visual,
            Visual::Image {
                media_type,
                bytes,
                ..
            } if media_type == "image/png" && bytes == PNG
        ));
        assert_eq!(document.objects[1].text.as_deref(), Some("Footer"));
        assert!(
            document
                .objects
                .iter()
                .all(|object| object.text.as_deref() != Some("INSERT TITLE HERE")
                    && object.text.as_deref() != Some("2/2/2026"))
        );
    }

    #[test]
    fn renders_decorative_master_custom_shapes() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <style:style style:name="Panel" style:family="graphic">
              <style:graphic-properties draw:fill="solid" draw:fill-color="#ffffff" draw:stroke="none"/>
            </style:style>
            <style:style style:name="Marker" style:family="graphic">
              <style:graphic-properties draw:fill="solid" draw:fill-color="#ef7622" draw:stroke="none"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout">
              <draw:g draw:id="decorations">
              <draw:custom-shape draw:id="panel" draw:name="Rectangle 2" draw:style-name="Panel"
                svg:x="0in" svg:y="0in" svg:width="4in" svg:height="7.5in">
                <text:p/>
                <draw:enhanced-geometry draw:type="non-primitive" svg:viewBox="0 0 21600 21600"
                  draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 Z N"/>
              </draw:custom-shape>
              <draw:custom-shape draw:id="marker" draw:name="Oval 26" draw:style-name="Marker"
                svg:width="0.75in" svg:height="0.75in" draw:transform="translate(5in 1in)">
                <text:p/>
                <draw:enhanced-geometry draw:type="non-primitive" svg:viewBox="0 0 21600 21600"/>
              </draw:custom-shape>
              </draw:g>
            </style:master-page>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw">
          <office:body><office:presentation><draw:page draw:master-page-name="Master"/></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("decorative master shapes");

        assert_eq!(document.objects.len(), 2);
        assert!(matches!(
            &document.objects[0].visual,
            Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: Paint::Solid(0xffff_ffff),
                ..
            }
        ));
        assert!(matches!(
            &document.objects[1].visual,
            Visual::Layer {
                transform,
                visual,
                ..
            } if transform.e == 480.0
                && transform.f == 96.0
                && matches!(visual.as_ref(), Visual::PaintedShape {
                    geometry: Geometry::Ellipse,
                    fill: Paint::Solid(0xef76_22ff),
                    ..
                })
        ));
    }

    #[test]
    fn preserves_circular_corners_for_stretched_odf_round_rectangles() {
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:custom-shape draw:id="card" draw:name="Rounded Rectangle"
              svg:x="0.825in" svg:y="2.015in" svg:width="5.62in" svg:height="1.62in">
              <draw:enhanced-geometry draw:type="non-primitive"
                draw:path-stretchpoint-x="10800" draw:path-stretchpoint-y="10800"
                svg:viewBox="0 0 21600 21600" draw:modifiers="3600"
                draw:enhanced-path="M ?f7 0 X 0 ?f8 L 0 ?f9 Y ?f7 21600 L ?f10 21600 X 21600 ?f9 L 21600 ?f8 Y ?f10 0 Z N"/>
            </draw:custom-shape>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let parsed = parse_content_fixture(&[(CONTENT_PART, content)]).expect("rounded card");

        let Visual::PaintedShape { geometry, .. } = &parsed.objects[0].visual else {
            panic!("enhanced rounded card must remain a painted shape");
        };
        let Geometry::RoundedRectangle { radius_x, radius_y } = geometry else {
            panic!("enhanced rounded card must use rounded rectangle geometry");
        };
        let expected_radius = 1.62 * 96.0 / 6.0;
        assert!((radius_x - expected_radius).abs() < 0.001);
        assert!((radius_y - expected_radius).abs() < 0.001);
    }

    #[test]
    fn evaluates_odf_enhanced_path_equations_for_custom_shapes() {
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:custom-shape draw:id="arrow" draw:name="Localized Arrow"
              svg:x="1in" svg:y="1in" svg:width="4in" svg:height="2in">
              <draw:enhanced-geometry draw:type="non-primitive" draw:modifiers="50000"
                svg:viewBox="0 0 21600 21600"
                draw:enhanced-path="M ?f0 ?f2 L ?f11 ?f2 ?f1 ?f6 ?f11 ?f3 ?f0 ?f3 Z N">
                <draw:equation draw:name="f0" draw:formula="left"/>
                <draw:equation draw:name="f1" draw:formula="right"/>
                <draw:equation draw:name="f2" draw:formula="top"/>
                <draw:equation draw:name="f3" draw:formula="bottom"/>
                <draw:equation draw:name="f4" draw:formula="?f3 - ?f2"/>
                <draw:equation draw:name="f5" draw:formula="?f4 / 2"/>
                <draw:equation draw:name="f6" draw:formula="?f2 + ?f5"/>
                <draw:equation draw:name="f7" draw:formula="?f1 - ?f0"/>
                <draw:equation draw:name="f8" draw:formula="min(?f7, ?f4)"/>
                <draw:equation draw:name="f9" draw:formula="$0"/>
                <draw:equation draw:name="f11" draw:formula="?f1 - ?f10"/>
                <draw:equation draw:name="f10" draw:formula="?f8 * ?f9 / 100000"/>
              </draw:enhanced-geometry>
            </draw:custom-shape>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let parsed = parse_content_fixture(&[(CONTENT_PART, content)]).expect("enhanced arrow");

        let Visual::PaintedShape { geometry, .. } = &parsed.objects[0].visual else {
            panic!("enhanced custom shape must remain a painted shape");
        };
        let Geometry::Path { commands, .. } = geometry else {
            panic!("enhanced custom shape must use evaluated path geometry");
        };
        assert_eq!(
            commands,
            &[
                PathCommand::MoveTo { x: 0.0, y: 0.0 },
                PathCommand::LineTo { x: 192.0, y: 0.0 },
                PathCommand::LineTo { x: 384.0, y: 96.0 },
                PathCommand::LineTo { x: 192.0, y: 192.0 },
                PathCommand::LineTo { x: 0.0, y: 192.0 },
                PathCommand::ClosePath,
            ]
        );
        assert!(parsed.diagnostics.is_empty());
    }

    #[test]
    fn preserves_odf_stretch_point_shape_proportions() {
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:custom-shape draw:id="double-arrow"
              svg:x="0in" svg:y="0in" svg:width="4in" svg:height="1in">
              <draw:enhanced-geometry draw:type="non-primitive"
                draw:path-stretchpoint-x="21600" draw:path-stretchpoint-y="21600"
                draw:modifiers="50000 50000" svg:viewBox="0 0 21600 21600"
                draw:enhanced-path="M ?f0 ?f6 L ?f13 ?f2 ?f13 ?f16 ?f14 ?f16 ?f14 ?f2 ?f1 ?f6 ?f14 ?f3 ?f14 ?f17 ?f13 ?f17 ?f13 ?f3 Z N">
                <draw:equation draw:name="f0" draw:formula="left"/>
                <draw:equation draw:name="f1" draw:formula="right"/>
                <draw:equation draw:name="f2" draw:formula="top"/>
                <draw:equation draw:name="f3" draw:formula="bottom"/>
                <draw:equation draw:name="f4" draw:formula="?f3 - ?f2"/>
                <draw:equation draw:name="f5" draw:formula="?f4 / 2"/>
                <draw:equation draw:name="f6" draw:formula="?f2 + ?f5"/>
                <draw:equation draw:name="f7" draw:formula="?f1 - ?f0"/>
                <draw:equation draw:name="f10" draw:formula="min(?f7, ?f4)"/>
                <draw:equation draw:name="f13" draw:formula="?f10 * $1 / 100000"/>
                <draw:equation draw:name="f14" draw:formula="?f1 - ?f13"/>
                <draw:equation draw:name="f15" draw:formula="?f4 * $0 / 200000"/>
                <draw:equation draw:name="f16" draw:formula="?f6 - ?f15"/>
                <draw:equation draw:name="f17" draw:formula="?f6 + ?f15"/>
              </draw:enhanced-geometry>
            </draw:custom-shape>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let parsed = parse_content_fixture(&[(CONTENT_PART, content)]).expect("double arrow");

        let Visual::PaintedShape { geometry, .. } = &parsed.objects[0].visual else {
            panic!("stretch-point custom shape must remain a painted shape");
        };
        let Geometry::Path { commands, .. } = geometry else {
            panic!("stretch-point custom shape must use evaluated path geometry");
        };
        assert_eq!(
            commands,
            &[
                PathCommand::MoveTo { x: 0.0, y: 48.0 },
                PathCommand::LineTo { x: 48.0, y: 0.0 },
                PathCommand::LineTo { x: 48.0, y: 24.0 },
                PathCommand::LineTo { x: 336.0, y: 24.0 },
                PathCommand::LineTo { x: 336.0, y: 0.0 },
                PathCommand::LineTo { x: 384.0, y: 48.0 },
                PathCommand::LineTo { x: 336.0, y: 96.0 },
                PathCommand::LineTo { x: 336.0, y: 72.0 },
                PathCommand::LineTo { x: 48.0, y: 72.0 },
                PathCommand::LineTo { x: 48.0, y: 96.0 },
                PathCommand::ClosePath,
            ]
        );
    }

    #[test]
    fn applies_odf_enhanced_geometry_text_areas_to_shape_layout() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:custom-shape draw:id="text-area" svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in">
              <draw:text-box><text:p>Shape text</text:p></draw:text-box>
              <draw:enhanced-geometry svg:viewBox="0 0 100 100"
                draw:enhanced-path="M 0 0 L 100 0 100 100 0 100 Z"
                draw:text-areas="25 10 75 90"/>
            </draw:custom-shape>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let parsed = parse_content_fixture(&[(CONTENT_PART, content)]).expect("enhanced text area");

        let Visual::TextLayout { layout, .. } = &parsed.objects[0].visual else {
            panic!("custom shape text must retain layout");
        };
        assert!((layout.inset_left - 48.0).abs() < 0.001);
        assert!((layout.inset_right - 48.0).abs() < 0.001);
        assert!((layout.inset_top - 9.6).abs() < 0.001);
        assert!((layout.inset_bottom - 9.6).abs() < 0.001);
    }

    #[test]
    fn applies_odf_shape_opacity_to_fill_without_fading_text() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <style:style style:name="TransparentShape" style:family="graphic">
              <style:graphic-properties draw:fill="solid" draw:fill-color="#336699" draw:opacity="25%"/>
              <style:text-properties fo:color="#ffffff"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:rect draw:id="transparent" draw:style-name="TransparentShape"
              svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in">
              <draw:text-box><text:p>Opaque text</text:p></draw:text-box>
            </draw:rect>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let parsed = parse(&package).expect("transparent shape");

        let Visual::TextLayout { visual, .. } = &parsed.objects[0].visual else {
            panic!("fill opacity must not wrap the complete text shape");
        };
        let Visual::RichText { fill, runs, .. } = visual.as_ref() else {
            panic!("transparent shape must retain rich text");
        };
        assert_eq!(*fill, Paint::Solid(0x3366_9940));
        assert_eq!(runs[0].color, 0xffff_ffff);
    }

    #[test]
    fn renders_odf_counterclockwise_and_clockwise_elliptical_arcs() {
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:custom-shape draw:id="arcs" svg:x="1in" svg:y="1in"
              svg:width="1in" svg:height="1in">
              <draw:enhanced-geometry draw:type="non-primitive" svg:viewBox="0 0 100 100"
                draw:mirror-horizontal="true"
                draw:enhanced-path="M 100 50 A 0 0 100 100 100 50 50 0 W 0 0 100 100 50 0 100 50 N"/>
            </draw:custom-shape>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let parsed = parse_content_fixture(&[(CONTENT_PART, content)]).expect("enhanced arcs");

        let Visual::PaintedShape { geometry, .. } = &parsed.objects[0].visual else {
            panic!("enhanced arcs must remain a painted shape");
        };
        let Geometry::Path { commands, .. } = geometry else {
            panic!("enhanced arcs must use path geometry");
        };
        assert_eq!(
            commands
                .iter()
                .filter(|command| matches!(command, PathCommand::BezierCurveTo { .. }))
                .count(),
            2
        );
        assert!(matches!(
            commands.first(),
            Some(PathCommand::MoveTo { x, y })
                if x.abs() < 0.001 && (*y - 48.0).abs() < 0.001
        ));
        assert!(parsed.diagnostics.is_empty());
    }

    #[test]
    fn ignores_degenerate_odf_arc_branches_without_inserting_lines() {
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:custom-shape draw:id="conditional-arc" svg:x="1in" svg:y="1in"
              svg:width="1in" svg:height="1in">
              <draw:enhanced-geometry draw:type="non-primitive" svg:viewBox="0 0 100 100"
                draw:enhanced-path="M 50 50 A 50 0 50 100 50 50 100 50 W 0 0 100 100 50 50 100 50 N"/>
            </draw:custom-shape>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let parsed =
            parse_content_fixture(&[(CONTENT_PART, content)]).expect("conditional enhanced arc");

        let Visual::PaintedShape { geometry, .. } = &parsed.objects[0].visual else {
            panic!("conditional arc must remain a painted shape");
        };
        let Geometry::Path { commands, .. } = geometry else {
            panic!("conditional arc must use path geometry");
        };
        assert_eq!(
            commands
                .iter()
                .filter(|command| matches!(command, PathCommand::LineTo { .. }))
                .count(),
            1
        );
        assert!(parsed.diagnostics.is_empty());
    }

    #[test]
    fn preserves_odf_enhanced_path_fill_and_stroke_layers() {
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:custom-shape draw:id="layers" svg:x="1in" svg:y="1in"
              svg:width="1in" svg:height="1in">
              <draw:enhanced-geometry draw:type="non-primitive" svg:viewBox="0 0 100 100"
                draw:enhanced-path="S M 0 0 L 100 0 N F M 0 100 L 100 100 N"/>
            </draw:custom-shape>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let parsed = parse_content_fixture(&[(CONTENT_PART, content)]).expect("enhanced layers");

        let Visual::PaintedShape { geometry, .. } = &parsed.objects[0].visual else {
            panic!("enhanced layers must remain a painted shape");
        };
        let Geometry::LayeredPath { layers } = geometry else {
            panic!("enhanced fill and stroke modes must use layered path geometry");
        };
        assert_eq!(layers.len(), 2);
        assert_eq!(layers[0].fill, PathFillMode::Normal);
        assert!(!layers[0].stroke);
        assert_eq!(layers[1].fill, PathFillMode::None);
        assert!(layers[1].stroke);
    }

    #[test]
    fn renders_odf_angle_ellipse_smiley_paths() {
        let content = br##"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:automatic-styles>
            <style:style xmlns:style="style" style:name="smiley-style" style:family="graphic">
              <style:graphic-properties draw:stroke="solid" svg:stroke-width="0cm" draw:stroke-color="#3465a4" draw:fill="solid" draw:fill-color="#729fcf"/>
            </style:style>
          </office:automatic-styles>
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:custom-shape draw:id="smiley" draw:style-name="smiley-style" svg:x="1in" svg:y="1in"
              svg:width="2in" svg:height="2in">
              <draw:enhanced-geometry draw:type="smiley" svg:viewBox="0 0 21600 21600"
                draw:modifiers="18520"
                draw:enhanced-path="U 10800 10800 10800 10800 0 360 Z N U 7305 7515 1000 1865 0 360 Z N U 14295 7515 1000 1865 0 360 Z N M 4870 ?f1 C 8680 ?f2 12920 ?f2 16730 ?f1 F N">
                <draw:equation draw:name="f0" draw:formula="$0 -14510"/>
                <draw:equation draw:name="f1" draw:formula="18520-?f0"/>
                <draw:equation draw:name="f2" draw:formula="14510+?f0"/>
              </draw:enhanced-geometry>
            </draw:custom-shape>
          </draw:page></office:presentation></office:body>
        </office:document-content>"##;
        let styles_xml = br#"<office:document-styles xmlns:office="office"/>"#;
        let bytes = stored_zip(&[(CONTENT_PART, content), ("styles.xml", styles_xml)]);
        let package = Package::open(&bytes, Limits::default()).expect("ODF smiley package");
        let styles = super::parse_style_catalog(&package).expect("ODF smiley styles");
        let parsed = parse_content(&package, &page_sizes(), &styles).expect("ODF smiley");

        let Visual::PaintedShape {
            geometry,
            stroke_width,
            ..
        } = &parsed.objects[0].visual
        else {
            panic!("smiley must remain a painted shape");
        };
        let Geometry::LayeredPath { layers } = geometry else {
            panic!("smiley must use layered path geometry");
        };
        assert_eq!(layers.len(), 4);
        assert!(layers.iter().all(|layer| layer.stroke));
        assert_eq!(*stroke_width, 1.0);
        assert_eq!(
            layers
                .iter()
                .flat_map(|layer| &layer.commands)
                .filter(|command| matches!(command, PathCommand::BezierCurveTo { .. }))
                .count(),
            13
        );
        assert!(parsed.diagnostics.is_empty());
    }

    #[test]
    fn excludes_slide_notes_from_the_visual_object_tree() {
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:presentation="presentation" xmlns:svg="svg" xmlns:text="text">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame draw:id="visible" svg:x="1in" svg:y="1in" svg:width="3in" svg:height="1in">
              <draw:text-box><text:p>Visible slide text</text:p></draw:text-box>
            </draw:frame>
            <presentation:notes>
              <draw:frame draw:id="speaker-notes" svg:x="1in" svg:y="4in" svg:width="6in" svg:height="2in">
                <draw:text-box><text:p>Speaker notes must not render</text:p></draw:text-box>
              </draw:frame>
            </presentation:notes>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let parsed = parse_content_fixture(&[(CONTENT_PART, content)]).expect("valid slide notes");

        assert_eq!(parsed.objects.len(), 1);
        assert_eq!(
            parsed.objects[0].text.as_deref(),
            Some("Visible slide text")
        );
    }

    #[test]
    fn renders_a_stretched_bitmap_drawing_page_background() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:fo="fo" xmlns:xlink="xlink">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <draw:fill-image draw:name="Background" xlink:href="Pictures/background.png"/>
            <style:style style:name="SlideBackground" style:family="drawing-page">
              <style:drawing-page-properties draw:fill="bitmap" draw:fill-image-name="Background" style:repeat="stretch"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw">
          <office:body><office:presentation><draw:page draw:master-page-name="Master" draw:style-name="SlideBackground"/></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[
            ("styles.xml", styles),
            (CONTENT_PART, content),
            ("Pictures/background.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("bitmap slide background");

        assert_eq!(document.objects.len(), 1);
        let background = &document.objects[0];
        assert_eq!(background.kind, ObjectKind::Image);
        assert_eq!(background.bounds.width, 960.0);
        assert_eq!(background.bounds.height, 720.0);
        assert!(matches!(
            &background.visual,
            Visual::Image {
                media_type,
                bytes,
                ..
            } if media_type == "image/png" && bytes == PNG
        ));
    }

    #[test]
    fn renders_a_tiled_bitmap_drawing_page_background_at_authored_size() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:fo="fo" xmlns:xlink="xlink">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="28cm" fo:page-height="21cm"/>
            </style:page-layout>
            <draw:fill-image draw:name="Aqua" xlink:href="Pictures/aqua.png"/>
            <style:style style:name="SlideBackground" style:family="drawing-page">
              <style:drawing-page-properties draw:fill="bitmap" draw:fill-image-name="Aqua"
                draw:fill-image-width="0.483cm" draw:fill-image-height="0.483cm"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw">
          <office:body><office:presentation><draw:page draw:master-page-name="Master" draw:style-name="SlideBackground"/></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[
            ("styles.xml", styles),
            (CONTENT_PART, content),
            ("Pictures/aqua.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("tiled bitmap slide background");

        assert_eq!(document.objects.len(), 1);
        let Visual::PaintedShape { fill, .. } = &document.objects[0].visual else {
            panic!("a tiled slide background must remain an image paint");
        };
        let Paint::Image {
            media_type,
            bytes,
            tile,
            tile_width,
            tile_height,
            ..
        } = fill
        else {
            panic!("a tiled slide background must retain its embedded image");
        };
        assert_eq!(media_type, "image/png");
        assert_eq!(bytes, PNG);
        assert!(*tile);
        assert_eq!(*tile_width, parse_length("0.483cm"));
        assert_eq!(*tile_height, parse_length("0.483cm"));
    }

    #[test]
    fn preserves_inherited_no_fill_when_child_only_changes_fill_color() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:fo="fo">
          <office:styles>
            <style:style style:name="MasterTitle" style:family="presentation">
              <style:graphic-properties draw:fill="none" draw:stroke="none" draw:textarea-vertical-align="middle"/>
            </style:style>
          </office:styles>
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <style:style style:name="Title" style:family="presentation" style:parent-style-name="MasterTitle">
              <style:graphic-properties draw:fill-color="#ffffff"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame draw:id="title" draw:style-name="Title" svg:x="1in" svg:y="1in" svg:width="8in" svg:height="3in">
              <draw:text-box><text:p>Transparent title</text:p></draw:text-box>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("inherited transparent title");

        let Visual::TextLayout { visual, .. } = &document.objects[0].visual else {
            panic!("title must retain text layout");
        };
        let Visual::RichText { fill, .. } = visual.as_ref() else {
            panic!("title must retain rich text");
        };
        assert!(matches!(fill, Paint::None));
    }

    #[test]
    fn ignores_pretty_print_whitespace_around_odf_text_controls() {
        let styles =
            br#"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame draw:id="title" svg:x="1in" svg:y="1in" svg:width="8in" svg:height="3in">
              <draw:text-box>
                <text:p>
                  <text:span>First line </text:span>
                  <text:span>
                    <text:line-break/>
                  </text:span>
                  <text:span>
                    <text:line-break/>
                  </text:span>
                  <text:span>Second line</text:span>
                </text:p>
              </draw:text-box>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("pretty-printed ODF text");

        assert_eq!(
            document.objects[0].text.as_deref(),
            Some("First line \n\nSecond line")
        );
    }

    #[test]
    fn renders_embedded_bitmap_shape_fills() {
        let styles =
            br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:fo="fo" xmlns:xlink="xlink">
          <office:styles>
            <draw:fill-image draw:name="Texture" xlink:href="Pictures/texture.png"/>
            <style:style style:name="Textured" style:family="graphic">
              <style:graphic-properties draw:fill="bitmap" draw:fill-image-name="Texture"
                style:repeat="repeat" draw:fill-image-width="0.5in"
                draw:fill-image-height="0.25in" draw:stroke="none"/>
            </style:style>
            <style:style style:name="Stretched" style:family="graphic">
              <style:graphic-properties draw:fill="bitmap" draw:fill-image-name="Texture"
                style:repeat="stretch" draw:stroke="none"/>
            </style:style>
          </office:styles>
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:rect draw:id="texture" draw:style-name="Textured"
              svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in"/>
            <draw:rect draw:id="stretch" draw:style-name="Stretched"
              svg:x="4in" svg:y="1in" svg:width="2in" svg:height="1in"/>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[
            ("styles.xml", styles),
            (CONTENT_PART, content),
            ("Pictures/texture.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let document = parse(&package).expect("bitmap-filled ODP shape");

        let Visual::PaintedShape { fill, .. } = &document.objects[0].visual else {
            panic!("bitmap-filled shape must remain painted");
        };
        let Paint::Image {
            media_type,
            bytes,
            tile,
            tile_width,
            tile_height,
            ..
        } = fill
        else {
            panic!("bitmap shape fill must retain its embedded image");
        };
        assert_eq!(media_type, "image/png");
        assert_eq!(bytes, PNG);
        assert!(*tile);
        assert_eq!(*tile_width, Some(48.0));
        assert_eq!(*tile_height, Some(24.0));
        assert!(matches!(
            &document.objects[1].visual,
            Visual::PaintedShape {
                fill: Paint::Image { tile: false, .. },
                ..
            }
        ));
    }

    #[test]
    fn applies_shape_span_paragraph_and_text_box_styles() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <style:style style:name="Shape" style:family="graphic">
              <style:graphic-properties draw:fill="solid" draw:fill-color="#000080" draw:stroke="none"
                fo:padding-left="0.1in" fo:padding-right="0.2in" fo:padding-top="0.05in" fo:padding-bottom="0.15in"
                fo:wrap-option="no-wrap" draw:textarea-vertical-align="middle" draw:textarea-horizontal-align="center"
                style:shrink-to-fit="true"/>
            </style:style>
            <style:style style:name="Paragraph" style:family="paragraph">
              <style:paragraph-properties fo:text-align="center" fo:line-height="90%"/>
            </style:style>
            <style:style style:name="Span" style:family="text">
              <style:text-properties style:font-family="Segoe UI" fo:font-size="12pt" fo:color="#ffffff" fo:font-weight="bold"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:custom-shape draw:id="styled-shape" draw:style-name="Shape" svg:x="1in" svg:y="1in" svg:width="3in" svg:height="1in">
              <text:p text:style-name="Paragraph"><text:span text:style-name="Span">Fleet</text:span><text:s text:c="2"/><text:span text:style-name="Span">Server</text:span></text:p>
            </draw:custom-shape>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("styled shape");

        assert_eq!(document.objects.len(), 1);
        assert_eq!(document.objects[0].text.as_deref(), Some("Fleet  Server"));
        let Visual::TextLayout { layout, visual } = &document.objects[0].visual else {
            panic!("shape text must carry explicit text-box layout");
        };
        assert_eq!(layout.vertical_align, TextVerticalAlign::Center);
        assert_eq!(layout.auto_fit, TextAutoFit::Shrink);
        assert!(!layout.wrap);
        assert_eq!(layout.inset_left, 9.6);
        assert_eq!(layout.inset_right, 19.2);
        assert_eq!(layout.inset_top, 4.8);
        assert_eq!(layout.inset_bottom, 14.400001);
        assert_eq!(layout.paragraphs.len(), 1);
        assert_eq!(layout.paragraphs[0].align, TextAlign::Center);
        let Visual::RichText { align, runs, .. } = visual.as_ref() else {
            panic!("shape text must remain rich text");
        };
        assert_eq!(*align, TextAlign::Center);
        assert_eq!(
            runs.iter().map(|run| run.text.as_str()).collect::<String>(),
            "Fleet  Server"
        );
        assert_eq!(runs[0].font_family, "Segoe UI");
        assert_eq!(runs[0].font_size, 16.0);
        assert_eq!(runs[0].color, 0xffff_ffff);
        assert!(runs[0].bold);
    }

    #[test]
    fn preserves_odg_two_columns() {
        let bytes = include_bytes!("../../tests/fixtures/two_columns.odg");
        let package = Package::open(bytes, Limits::default()).unwrap();
        let document = parse(&package).unwrap();
        let object = document
            .objects
            .iter()
            .find(|object| {
                object
                    .text
                    .as_deref()
                    .is_some_and(|text| text.starts_with("Lorem ipsum"))
            })
            .unwrap();
        let Visual::TextLayout { layout, .. } = &object.visual else {
            panic!("drawing text must retain its layout");
        };
        assert_eq!(layout.column_count, 2);
        assert!((layout.column_spacing - 0.7 * CSS_PIXELS_PER_INCH / 2.54).abs() < 0.001);
        assert!(object.text.as_ref().unwrap().contains("mollis nunc."));

        // The same source style must survive inheritance, and an explicit
        // single-column child must reset both the inherited count and gap.
        let content =
            String::from_utf8(package.required_part(CONTENT_PART).unwrap().into_vec()).unwrap();
        let styles = package.required_part("styles.xml").unwrap();
        for (properties, count) in [
            ("", 2),
            (
                "<style:graphic-properties><style:columns fo:column-count=\"1\"/></style:graphic-properties>",
                1,
            ),
        ] {
            let content = content.replace("</office:automatic-styles>", &format!(
                "<style:style style:name=\"Child\" style:family=\"graphic\" style:parent-style-name=\"gr1\">{properties}</style:style></office:automatic-styles>"
            )).replace("draw:style-name=\"gr1\"", "draw:style-name=\"Child\"");
            let bytes = stored_zip(&[
                (CONTENT_PART, content.as_bytes()),
                ("styles.xml", styles.as_ref()),
            ]);
            let package = Package::open(&bytes, Limits::default()).unwrap();
            let document = parse(&package).unwrap();
            let Visual::TextLayout {
                layout: inherited, ..
            } = &document.objects[0].visual
            else {
                panic!("inherited drawing text layout");
            };
            assert_eq!(inherited.column_count, count);
            assert_eq!(
                inherited.column_spacing,
                if count == 1 {
                    0.0
                } else {
                    layout.column_spacing
                }
            );
        }
    }

    #[test]
    fn preserves_fit_to_frame_text_layout() {
        let styles = br#"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
          <office:automatic-styles><style:page-layout style:name="PageLayout">
            <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
          </style:page-layout></office:automatic-styles>
          <office:master-styles><style:master-page style:name="Master" style:page-layout-name="PageLayout"/></office:master-styles>
        </office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text" xmlns:fo="fo">
          <office:automatic-styles>
            <style:style style:name="FitFrame" style:family="graphic">
              <style:graphic-properties draw:fit-to-size="true"/>
            </style:style>
            <style:style style:name="Huge" style:family="text">
              <style:text-properties fo:font-size="999pt"/>
            </style:style>
          </office:automatic-styles>
          <office:body><office:drawing><draw:page draw:master-page-name="Master">
            <draw:frame draw:style-name="FitFrame" svg:x="0in" svg:y="0in" svg:width="4in" svg:height="0.2in">
              <draw:text-box><text:p><text:span text:style-name="Huge">Fit frame text</text:span></text:p></draw:text-box>
            </draw:frame>
          </draw:page></office:drawing></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("fit-to-frame drawing");

        assert_eq!(document.objects[0].text.as_deref(), Some("Fit frame text"));
        let Visual::TextLayout { layout, .. } = &document.objects[0].visual else {
            panic!("fit-to-frame text must carry text layout");
        };
        assert_eq!(layout.auto_fit, TextAutoFit::FitFrame);
        assert!(!layout.wrap);
        assert!(!layout.text_scale_to_fit);
    }

    #[test]
    fn renders_sideways_lr_text_boxes_as_rotated_270() {
        let styles = br#"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <style:style style:name="VerticalDate" style:family="graphic">
              <style:graphic-properties fo:padding="0in" draw:textarea-vertical-align="middle"/>
              <style:paragraph-properties style:writing-mode="sideways-lr"/>
            </style:style>
            <style:style style:name="HorizontalParagraph" style:family="paragraph">
              <style:paragraph-properties fo:line-height="100%" style:writing-mode="lr-tb"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame draw:id="date" draw:style-name="VerticalDate" svg:x="2in" svg:y="2in" svg:width="0.3in" svg:height="1in">
              <draw:text-box><text:p text:style-name="HorizontalParagraph">03/2026</text:p></draw:text-box>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("sideways ODF text box");

        let Visual::TextLayout { layout, .. } = &document.objects[0].visual else {
            panic!("sideways text must carry explicit text layout");
        };
        assert_eq!(layout.orientation, TextOrientation::Rotated270);
        assert_eq!(layout.vertical_align, TextVerticalAlign::Center);
        assert_eq!(layout.inset_top, 0.0);
    }

    #[test]
    fn maps_odf_gradient_axes_and_radial_centers_to_object_space() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <draw:gradient draw:name="Linear" draw:style="linear" draw:angle="2700"
              draw:start-color="#ff0000" draw:end-color="#0000ff"/>
            <draw:gradient draw:name="Radial" draw:style="radial" draw:cx="25%" draw:cy="75%"
              draw:start-color="#ffffff" draw:end-color="#000000"/>
            <draw:gradient draw:name="Vertical" draw:style="linear" draw:angle="0"
              draw:start-color="#ff0000" draw:end-color="#0000ff"/>
            <draw:gradient draw:name="PowerPointCorner" draw:style="radial" draw:cx="100%" draw:cy="0%"
              draw:start-color="#d6d6d6" draw:end-color="#002060"/>
            <draw:gradient draw:name="PowerPointRectCorner" draw:style="rectangular" draw:cx="0%" draw:cy="100%"
              draw:start-color="#7030a0" draw:end-color="#002060"/>
            <style:style style:name="LinearFill" style:family="graphic">
              <style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="Linear" draw:stroke="none"/>
            </style:style>
            <style:style style:name="RadialFill" style:family="graphic">
              <style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="Radial" draw:stroke="none"/>
            </style:style>
            <style:style style:name="VerticalFill" style:family="graphic">
              <style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="Vertical" draw:stroke="none"/>
            </style:style>
            <style:style style:name="PowerPointCornerFill" style:family="graphic">
              <style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="PowerPointCorner" draw:stroke="none"/>
            </style:style>
            <style:style style:name="PowerPointRectCornerFill" style:family="graphic">
              <style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="PowerPointRectCorner" draw:stroke="none"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:rect draw:style-name="LinearFill" svg:x="0in" svg:y="0in" svg:width="2in" svg:height="1in"/>
            <draw:rect draw:style-name="RadialFill" svg:x="0in" svg:y="1in" svg:width="2in" svg:height="1in"/>
            <draw:rect draw:style-name="VerticalFill" svg:x="0in" svg:y="2in" svg:width="2in" svg:height="1in"/>
            <draw:rect draw:style-name="PowerPointCornerFill" svg:x="0in" svg:y="3in" svg:width="2in" svg:height="1in"/>
            <draw:rect draw:style-name="PowerPointRectCornerFill" svg:x="0in" svg:y="4in" svg:width="2in" svg:height="1in"/>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("ODF gradients");

        let Visual::PaintedShape { fill, .. } = &document.objects[0].visual else {
            panic!("linear gradient shape");
        };
        let Paint::LinearGradient { x0, y0, x1, y1, .. } = fill else {
            panic!("linear ODF gradient");
        };
        assert!((*x0 - 192.0).abs() < 0.001);
        assert!((*y0 - 48.0).abs() < 0.001);
        assert!(x1.abs() < 0.001);
        assert!((*y1 - 48.0).abs() < 0.001);

        let Visual::PaintedShape { fill, .. } = &document.objects[1].visual else {
            panic!("radial gradient shape");
        };
        let Paint::RadialGradient {
            x0, y0, x1, y1, r1, ..
        } = fill
        else {
            panic!("radial ODF gradient");
        };
        assert!((*x0 - 48.0).abs() < 0.001);
        assert!((*y0 - 72.0).abs() < 0.001);
        assert!((*x1 - 48.0).abs() < 0.001);
        assert!((*y1 - 72.0).abs() < 0.001);
        assert!((*r1 - 160.996_89).abs() < 0.001);

        let Visual::PaintedShape { fill, .. } = &document.objects[2].visual else {
            panic!("vertical gradient shape");
        };
        let Paint::LinearGradient { x0, y0, x1, y1, .. } = fill else {
            panic!("vertical ODF gradient");
        };
        assert!((*x0 - 96.0).abs() < 0.001);
        assert!(y0.abs() < 0.001);
        assert!((*x1 - 96.0).abs() < 0.001);
        assert!((*y1 - 96.0).abs() < 0.001);

        let Visual::PaintedShape { fill, .. } = &document.objects[3].visual else {
            panic!("PowerPoint corner gradient shape");
        };
        let Paint::RadialGradient {
            x0, y0, r1, stops, ..
        } = fill
        else {
            panic!("PowerPoint corner radial gradient");
        };
        assert!(x0.abs() < 0.001);
        assert!(y0.abs() < 0.001);
        assert!((*r1 - 96.0).abs() < 0.001);
        assert_eq!(stops[0].color, 0x0020_60ff);
        assert_eq!(stops[1].color, 0xd6d6_d6ff);

        let Visual::PaintedShape { fill, .. } = &document.objects[4].visual else {
            panic!("PowerPoint corner rectangular gradient shape");
        };
        let Paint::RectGradient {
            center_x,
            center_y,
            stops,
        } = fill
        else {
            panic!("PowerPoint corner rectangular gradient");
        };
        assert!((*center_x - 192.0).abs() < 0.001);
        assert!((*center_y - 96.0).abs() < 0.001);
        assert_eq!(stops[0].color, 0x0020_60ff);
        assert_eq!(stops[1].color, 0x7030_a0ff);
    }

    #[test]
    fn inherits_nested_list_style_in_n828390_5() {
        let bytes = include_bytes!("../../tests/fixtures/n828390_5.odp");
        let package = Package::open(bytes, Limits::default()).unwrap();
        let document = parse(&package).unwrap();
        let object = document
            .objects
            .iter()
            .find(|object| {
                object
                    .text
                    .as_deref()
                    .is_some_and(|text| text.contains("1st"))
            })
            .unwrap();
        assert_eq!(object.text.as_deref(), Some("\n–\t1st\n–\t2nd\n–\t3rd"));
    }

    #[test]
    fn renders_odf_list_labels_and_paragraph_metrics() {
        let styles =
            br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br##"<office:document-content xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text" xmlns:fo="fo">
          <office:automatic-styles>
            <style:style style:name="Paragraph" style:family="paragraph">
              <style:paragraph-properties fo:text-align="justify" fo:line-height="200%" fo:margin-left="0.25in"
                fo:margin-right="0.125in" fo:text-indent="-0.25in"
                fo:margin-bottom="0.083333in" style:tab-stop-distance="1in"/>
              <style:text-properties fo:font-size="0.25in"/>
            </style:style>
            <text:list-style style:name="Bullet">
              <text:list-level-style-bullet text:level="1" text:bullet-char="n">
                <style:list-level-properties text:space-before="0in" text:min-label-width="0.25in"/>
                <style:text-properties fo:color="#3333cc" fo:font-family="Wingdings" fo:font-size="60%"/>
              </text:list-level-style-bullet>
            </text:list-style>
            <text:list-style style:name="Number">
              <text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix=".">
                <style:list-level-properties text:space-before="0in" text:min-label-width="0.25in"/>
              </text:list-level-style-number>
            </text:list-style>
          </office:automatic-styles>
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame draw:id="list" svg:x="1in" svg:y="1in" svg:width="8in" svg:height="5in">
              <draw:text-box>
                <text:list text:style-name="Bullet"><text:list-item>
                  <text:p text:style-name="Paragraph">First</text:p>
                </text:list-item></text:list>
                <text:list text:style-name="Number"><text:list-item>
                  <text:p text:style-name="Paragraph">Second</text:p>
                </text:list-item></text:list>
                <text:list text:style-name="Number"><text:list-item>
                  <text:p text:style-name="Paragraph">Third</text:p>
                </text:list-item></text:list>
              </draw:text-box>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"##;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("ODF lists with paragraph metrics");

        assert_eq!(
            document.objects[0].text.as_deref(),
            Some("◼\tFirst\n1.\tSecond\n2.\tThird")
        );
        let Visual::TextLayout { layout, visual } = &document.objects[0].visual else {
            panic!("list text must carry explicit text layout");
        };
        assert_eq!(layout.paragraph_spacing, 7.999968);
        assert!((layout.inset_top - 20.8).abs() < 0.001);
        assert_eq!(layout.paragraphs.len(), 3);
        assert_eq!(layout.paragraphs[0].align, TextAlign::Justify);
        assert_eq!(layout.paragraphs[0].margin_left, 24.0);
        assert_eq!(layout.paragraphs[0].margin_right, 12.0);
        assert_eq!(layout.paragraphs[0].first_line_indent, -24.0);
        assert_eq!(layout.paragraphs[0].default_tab_stop, 96.0);
        assert!((layout.paragraphs[0].line_height - 57.6).abs() < 0.001);
        assert_eq!(layout.paragraphs[0].space_before, 0.0);
        assert!((layout.paragraphs[0].space_after - 7.999_968).abs() < 0.001);
        let Visual::RichText {
            line_height, runs, ..
        } = visual.as_ref()
        else {
            panic!("list text must remain rich text");
        };
        assert!((*line_height - 57.6).abs() < 0.001);
        assert_eq!(
            runs.iter().map(|run| run.text.as_str()).collect::<String>(),
            "◼\tFirst\n1.\tSecond\n2.\tThird"
        );
        assert_eq!(runs[0].text, "◼\t");
        assert_eq!(runs[0].color, 0x3333_ccff);
        assert!((runs[0].font_size - 14.4).abs() < 0.001);
        assert_eq!(runs[1].text, "First\n");
        assert_eq!(runs[1].color, 0x0000_00ff);
        assert!((runs[1].font_size - 24.0).abs() < 0.001);
        assert_eq!(
            super::super::normalize_symbol_font_character("q", Some("Wingdings")),
            "❑"
        );
    }

    #[test]
    fn formats_extended_microsoft_odf_list_numbers() {
        assert_eq!(
            super::format_odf_list_number("0001, 0002, 0003, ...", 1),
            "0001"
        );
        assert_eq!(super::format_odf_list_number("①, ②, ③, ...", 2), "②");
        assert_eq!(
            super::format_odf_list_number("一, 二, 三, 四, ...", 12),
            "十二"
        );
        assert_eq!(
            super::parse_odf_border("1px dash-dot-dot #123456")
                .unwrap()
                .2
                .dash,
            [8.0, 4.0, 2.0, 4.0, 2.0, 4.0]
        );
    }

    #[test]
    fn centers_single_and_multi_paragraph_odf_text_with_distinct_line_anchors() {
        let style = GraphicStyle {
            vertical_align: TextVerticalAlign::Center,
            font_size: 10.0,
            ..GraphicStyle::default()
        };
        let runs = vec![odf_text_run("Text".to_owned(), &style)];
        let styles = StyleCatalog::default();

        let single = odf_text_layout(&style, &[None], &[None], &runs, 20.0, true, &styles);
        let multiple = odf_text_layout(
            &style,
            &[None, None],
            &[None, None],
            &runs,
            20.0,
            true,
            &styles,
        );

        assert_eq!(single.inset_top, 4.0);
        assert_eq!(multiple.inset_top, 24.0);
    }

    #[test]
    fn preserves_explicit_list_text_when_number_format_is_empty() {
        let styles =
            br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:text="text">
          <office:automatic-styles>
            <text:list-style style:name="NoGeneratedLabel">
              <text:list-level-style-number text:level="1" style:num-format="">
                <style:list-level-properties text:space-before="0.25in"/>
              </text:list-level-style-number>
            </text:list-style>
          </office:automatic-styles>
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame draw:id="list" svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in">
              <draw:text-box>
                <text:list text:style-name="NoGeneratedLabel"><text:list-item>
                  <text:p>1</text:p>
                </text:list-item></text:list>
              </draw:text-box>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("list with intentionally empty number format");

        assert_eq!(document.objects[0].text.as_deref(), Some("1"));
        let Visual::TextLayout { layout, .. } = &document.objects[0].visual else {
            panic!("list text must retain its paragraph layout");
        };
        assert_eq!(layout.paragraphs[0].margin_left, 24.0);
    }

    #[test]
    fn renders_standard_connectors_as_elbows_with_triangle_markers() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <draw:marker draw:name="StartArrow" svg:viewBox="0 0 20 30" svg:d="m10 0-10 30h20z"/>
            <draw:marker draw:name="EndArrow" svg:viewBox="0 0 20 30" svg:d="m10 0-10 30h20z"/>
            <style:style style:name="Connector" style:family="graphic">
              <style:graphic-properties draw:fill="none" draw:stroke="solid" svg:stroke-width="0.03125in"
                svg:stroke-color="#7f7f7f" draw:marker-start="StartArrow" draw:marker-end="EndArrow"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:connector draw:id="elbow" draw:type="standard" draw:style-name="Connector"
              svg:x1="1in" svg:y1="1in" svg:x2="2in" svg:y2="4in"/>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("standard connector");

        let Visual::PaintedShape { geometry, .. } = &document.objects[0].visual else {
            panic!("connector must remain a painted shape");
        };
        let Geometry::Path { commands, .. } = geometry else {
            panic!("standard connector must use routed path geometry");
        };
        assert_eq!(
            &commands[..4],
            &[
                PathCommand::MoveTo { x: 0.0, y: 0.0 },
                PathCommand::LineTo { x: 48.0, y: 0.0 },
                PathCommand::LineTo { x: 48.0, y: 288.0 },
                PathCommand::LineTo { x: 96.0, y: 288.0 },
            ]
        );
        assert_eq!(
            commands
                .iter()
                .filter(|command| matches!(command, PathCommand::ClosePath))
                .count(),
            2
        );
    }

    #[test]
    fn renders_markers_on_mirrored_custom_lines() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <draw:marker draw:name="Arrow" svg:viewBox="0 0 20 30" svg:d="m10 0-10 30h20z"/>
            <style:style style:name="ArrowLine" style:family="graphic">
              <style:graphic-properties draw:fill="none" draw:stroke="solid"
                svg:stroke-width="0.03in" svg:stroke-color="#333399" draw:marker-start="Arrow"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:custom-shape draw:style-name="ArrowLine"
              svg:x="1in" svg:y="1in" svg:width="1in" svg:height="1in">
              <draw:enhanced-geometry draw:type="non-primitive" svg:viewBox="0 0 21600 21600"
                draw:enhanced-path="M 0 0 L 21600 21600 N" draw:mirror-horizontal="true"/>
            </draw:custom-shape>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("custom arrow line");

        let Visual::PaintedShape { geometry, .. } = &document.objects[0].visual else {
            panic!("custom line must remain a painted shape");
        };
        let Geometry::Path { commands, .. } = geometry else {
            panic!("custom line must retain its path geometry");
        };
        assert_eq!(commands[0], PathCommand::MoveTo { x: 96.0, y: 0.0 });
        assert_eq!(
            commands
                .iter()
                .filter(|command| matches!(command, PathCommand::ClosePath))
                .count(),
            1
        );
    }

    #[test]
    fn routes_a_connector_to_a_shape_declared_later_in_the_page() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in" xmlns:fo="fo"/>
            </style:page-layout>
            <style:style style:name="Connector" style:family="graphic">
              <style:graphic-properties draw:fill="none" draw:stroke="solid"
                svg:stroke-width="0.02in" svg:stroke-color="#000000"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:connector draw:id="elbow" draw:type="standard" draw:style-name="Connector"
              draw:end-shape="target" draw:end-glue-point="0"
              svg:x1="5in" svg:y1="0in" svg:x2="1.5in" svg:y2="1in"/>
            <draw:rect draw:id="target"
              svg:x="1in" svg:y="1in" svg:width="1in" svg:height="1in"/>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("connector may target a later shape");

        let Visual::PaintedShape { geometry, .. } = &document.objects[0].visual else {
            panic!("connector must remain a painted shape");
        };
        let Geometry::Path { commands, .. } = geometry else {
            panic!("attached connector must use routed path geometry");
        };
        assert_eq!(
            &commands[..4],
            &[
                PathCommand::MoveTo { x: 336.0, y: 0.0 },
                PathCommand::LineTo { x: 336.0, y: 48.0 },
                PathCommand::LineTo { x: 0.0, y: 48.0 },
                PathCommand::LineTo { x: 0.0, y: 96.0 },
            ]
        );
    }

    #[test]
    fn expands_dashed_connectors_without_wrapping_the_open_path() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <draw:stroke-dash draw:name="Dots" draw:dots1="1"
              draw:dots1-length="0.05in" draw:distance="0.03in"/>
            <draw:marker draw:name="Arrow" svg:viewBox="0 0 20 30" svg:d="m10 0-10 30h20z"/>
            <style:style style:name="DashedConnector" style:family="graphic">
              <style:graphic-properties draw:fill="none" draw:stroke="dash"
                draw:stroke-dash="Dots" svg:stroke-width="0.02in"
                svg:stroke-color="#000000" draw:marker-end="Arrow"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:connector draw:id="dashed" draw:type="line" draw:style-name="DashedConnector"
              svg:x1="1in" svg:y1="1in" svg:x2="3in" svg:y2="1in"/>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("dashed connector");

        let Visual::PaintedShape { geometry, .. } = &document.objects[0].visual else {
            panic!("connector must not use a stroke-style wrapper");
        };
        let Geometry::Path { commands, .. } = geometry else {
            panic!("dashed connector must use explicit path segments");
        };
        assert!(
            commands
                .iter()
                .filter(|command| matches!(command, PathCommand::MoveTo { .. }))
                .count()
                > 2
        );
        assert_eq!(
            commands
                .iter()
                .filter(|command| matches!(command, PathCommand::ClosePath))
                .count(),
            1
        );
    }

    #[test]
    fn preserves_supplied_odf_dash_linecaps() {
        use crate::model::LineCap::{Flat, Round, Square};

        let bytes = include_bytes!("../../tests/fixtures/odf-dash-linecaps.odp");
        let package = Package::open(bytes, Limits::default()).unwrap();
        let document = parse(&package).expect("supplied dash and linecap combinations");
        assert_eq!(document.objects.len(), 8);
        for (object, cap) in document
            .objects
            .iter()
            .zip([Flat, Flat, Round, Round, Square, Square, Flat, Round])
        {
            let (style, visual) = match &object.visual {
                Visual::StrokeStyle { style, visual } => (style.clone(), visual.as_ref()),
                visual => (crate::model::StrokeStyle::default(), visual),
            };
            assert_eq!(style.cap, cap, "{}", object.stable_id);
            let Visual::PaintedShape {
                stroke_width,
                stroke,
                geometry: Geometry::Path { commands, .. },
                ..
            } = visual
            else {
                panic!("line must remain parsed vector geometry");
            };
            assert_eq!(*stroke, Paint::Solid(0x0000_ffff));
            assert!(style.dash.is_empty(), "dash gaps are already in the path");
            assert_eq!(
                commands.len(),
                10,
                "five dashes; no dash begins at the path endpoint"
            );
            let (PathCommand::MoveTo { x: start, .. }, PathCommand::LineTo { x: end, .. }) =
                (&commands[0], &commands[1])
            else {
                panic!("dash segment");
            };
            assert!((end - start - stroke_width * 3.0).abs() < 0.001);
            let Some(PathCommand::LineTo { x, .. }) = commands.last() else {
                panic!("last dash endpoint");
            };
            assert!((*x - stroke_width * 43.0).abs() < 0.001);
        }
    }

    #[test]
    fn inherits_odf_linecaps_on_solid_and_marked_lines() {
        use crate::model::LineCap::{Flat, Round, Square};

        let source = Package::open(
            include_bytes!("../../tests/fixtures/odf-dash-linecaps.odp"),
            Limits::default(),
        )
        .unwrap();
        let styles = source.required_part("styles.xml").unwrap();
        let content = String::from_utf8(source.required_part(CONTENT_PART).unwrap().into_vec())
            .unwrap()
            .replace(
                "style:parent-style-name=\"objectwithoutfill\"",
                "style:parent-style-name=\"RoundParent\"",
            )
            .replace(
                "<office:automatic-styles>",
                r#"<office:automatic-styles>
              <style:style style:name="RoundParent" style:family="graphic">
                <style:graphic-properties svg:stroke-linecap="round"/>
              </style:style>"#,
            );
        for marked in [false, true] {
            let content = if marked {
                content.replace(
                    "draw:stroke=\"dash\"",
                    "draw:stroke=\"dash\" draw:marker-end=\"Arrow\"",
                )
            } else {
                content.replace("draw:stroke=\"dash\"", "draw:stroke=\"solid\"")
            };
            let bytes = stored_zip(&[("styles.xml", &styles), (CONTENT_PART, content.as_bytes())]);
            let package = Package::open(&bytes, Limits::default()).unwrap();
            let document = parse(&package).unwrap();
            assert_eq!(document.objects.len(), 8);
            for (object, expected) in document
                .objects
                .iter()
                .zip([Flat, Flat, Round, Round, Square, Square, Round, Round])
            {
                let (cap, visual) = match &object.visual {
                    Visual::StrokeStyle { style, visual } => {
                        assert!(style.dash.is_empty(), "do not dash an expanded path twice");
                        (style.cap, visual.as_ref())
                    }
                    visual => (Flat, visual),
                };
                assert_eq!(cap, expected);
                let Visual::PaintedShape {
                    geometry: Geometry::Path { commands, .. },
                    ..
                } = visual
                else {
                    panic!("line must retain its path");
                };
                assert_eq!(
                    commands.iter().any(|c| matches!(c, PathCommand::ClosePath)),
                    marked
                );
            }
        }
    }

    #[test]
    fn resolves_percentage_dash_lengths_against_stroke_width() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <draw:stroke-dash draw:name="Relative" draw:dots1="1"
              draw:dots1-length="300%" draw:dots2="0" draw:dots2-length="0in"
              draw:distance="700%"/>
            <style:style style:name="Dashed" style:family="graphic">
              <style:graphic-properties draw:fill="none" draw:stroke="dash"
                draw:stroke-dash="Relative" svg:stroke-width="2px" svg:stroke-color="#000000"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:rect draw:id="dashed" draw:style-name="Dashed"
              svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in"/>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("percentage dash lengths");

        let Visual::StrokeStyle { style, .. } = &document.objects[0].visual else {
            panic!("dashed shape must carry stroke styling");
        };
        assert_eq!(style.dash, vec![6.0, 14.0]);
    }

    #[test]
    fn omits_invalid_master_shapes_with_a_diagnostic() {
        let styles = br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:draw="draw" xmlns:svg="svg" xmlns:fo="fo" xmlns:text="text">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout">
              <draw:custom-shape svg:x="100cm" svg:y="100cm" svg:width="10cm" svg:height="-2cm">
                <text:p>Corrupt master shape</text:p>
              </draw:custom-shape>
            </style:master-page>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw">
          <office:body><office:presentation><draw:page draw:master-page-name="Master"/></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("invalid master geometry must not poison the slide");

        assert!(document.objects.is_empty());
        assert!(document.diagnostics.iter().any(|diagnostic| {
            diagnostic.code == DiagnosticCode::UnsupportedFeature
                && diagnostic.message.contains("invalid ODP master shapes")
        }));
    }

    #[test]
    fn prefers_a_supported_native_table_over_its_cached_image_fallback() {
        let styles =
            br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo" xmlns:table="table" xmlns:text="text">
          <office:styles>
            <style:style style:name="TemplateBody" style:family="table-cell">
              <style:table-cell-properties fo:background-color="#f0f0f0"
                fo:border="0.01389in solid #a5a5a5"/>
              <style:text-properties fo:font-family="+mn-lt" fo:font-size="9pt"/>
            </style:style>
            <style:style style:name="TemplateOddRow" style:family="table-cell">
              <style:table-cell-properties fo:background-color="#e1e1e1"/>
            </style:style>
            <table:table-template table:name="{0505E3EF-67EA-436B-97B2-0124C06EBD24}">
              <table:body table:style-name="TemplateBody"/>
              <table:odd-rows table:style-name="TemplateOddRow"/>
            </table:table-template>
          </office:styles>
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <style:style style:name="Column1" style:family="table-column">
              <style:table-column-properties style:column-width="1in"/>
            </style:style>
            <style:style style:name="Column2" style:family="table-column">
              <style:table-column-properties style:column-width="3in"/>
            </style:style>
            <style:style style:name="HeaderRow" style:family="table-row">
              <style:table-row-properties style:row-height="0.3in"/>
            </style:style>
            <style:style style:name="BodyRow" style:family="table-row">
              <style:table-row-properties style:row-height="0.7in"/>
            </style:style>
            <style:style style:name="HeaderCell" style:family="table-cell">
              <style:table-cell-properties style:vertical-align="middle"
                fo:background-color="#19226d" fo:padding="0.01in"/>
            </style:style>
            <style:style style:name="BodyCell" style:family="table-cell">
              <style:table-cell-properties style:vertical-align="middle" fo:padding="0.02in"/>
            </style:style>
            <style:style style:name="Center" style:family="paragraph">
              <style:paragraph-properties fo:text-align="center" fo:line-height="150%"/>
            </style:style>
            <style:style style:name="HeaderText" style:family="text">
              <style:text-properties fo:font-size="10pt" fo:color="#ffffff"
                fo:font-weight="bold"/>
            </style:style>
            <style:style style:name="BodyText" style:family="text">
              <style:text-properties fo:font-size="8pt"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:table="table" xmlns:text="text" xmlns:xlink="xlink">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame draw:id="cached-table" svg:x="1in" svg:y="1in" svg:width="4in" svg:height="2in">
              <table:table table:template-name="{0505E3EF-67EA-436B-97B2-0124C06EBD24}"
                table:use-banding-rows-styles="true" table:use-first-row-styles="true">
                <table:table-column table:style-name="Column1"/>
                <table:table-column table:style-name="Column2"/>
                <table:table-row table:style-name="HeaderRow">
                  <table:table-cell table:style-name="HeaderCell">
                    <text:p text:style-name="Center"><text:span text:style-name="HeaderText">No.</text:span></text:p>
                  </table:table-cell>
                  <table:table-cell table:style-name="HeaderCell">
                    <text:p text:style-name="Center"><text:span text:style-name="HeaderText">Gating Item</text:span></text:p>
                  </table:table-cell>
                </table:table-row>
                <table:table-row table:style-name="BodyRow">
                  <table:table-cell table:style-name="BodyCell">
                    <text:p text:style-name="Center"><text:span text:style-name="BodyText">01</text:span></text:p>
                  </table:table-cell>
                  <table:table-cell table:style-name="BodyCell">
                    <text:p><text:span text:style-name="BodyText">Duplicate table</text:span></text:p>
                    <text:p><text:span text:style-name="BodyText">text</text:span></text:p>
                  </table:table-cell>
                </table:table-row>
              </table:table>
              <draw:image xlink:href="Pictures/table.png"/>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[
            ("styles.xml", styles),
            (CONTENT_PART, content),
            ("Pictures/table.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let document = parse(&package).expect("native table with cached fallback");

        assert_eq!(document.objects.len(), 5);
        assert_eq!(document.objects[0].kind, ObjectKind::Table);
        assert_eq!(document.objects[1].kind, ObjectKind::Cell);
        assert_eq!(document.objects[1].bounds.width, 96.0);
        assert_eq!(document.objects[2].bounds.x, 192.0);
        assert_eq!(document.objects[2].bounds.width, 288.0);
        assert!((document.objects[1].bounds.height - 28.8).abs() < 0.001);
        assert!((document.objects[3].bounds.y - 124.8).abs() < 0.001);
        assert!((document.objects[3].bounds.height - 67.2).abs() < 0.001);
        assert_eq!(
            document.objects[4].text.as_deref(),
            Some("Duplicate table\ntext")
        );
        let Visual::TextLayout {
            layout,
            visual: header,
        } = &document.objects[1].visual
        else {
            panic!("header cell must retain text layout");
        };
        assert_eq!(layout.vertical_align, TextVerticalAlign::Center);
        assert_eq!(layout.inset_left, 0.96);
        assert_eq!(layout.inset_top, 0.96);
        let Visual::RichText {
            fill,
            stroke,
            stroke_width,
            align,
            runs,
            ..
        } = header.as_ref()
        else {
            panic!("header cell must retain rich text");
        };
        assert_eq!(*fill, Paint::Solid(0x1922_6dff));
        assert_eq!(*stroke, Paint::Solid(0xa5a5_a5ff));
        assert!((*stroke_width - 1.333_44).abs() < 0.001);
        assert_eq!(*align, TextAlign::Center);
        assert_eq!(runs[0].font_family, "Aptos");
        assert!((runs[0].font_size - 13.333_333).abs() < 0.001);
        assert_eq!(runs[0].color, 0xffff_ffff);
        assert!(runs[0].bold);
        let Visual::TextLayout { visual: body, .. } = &document.objects[3].visual else {
            panic!("body cell must retain text layout");
        };
        let Visual::RichText { fill, runs, .. } = body.as_ref() else {
            panic!("body cell must retain rich text");
        };
        assert_eq!(*fill, Paint::Solid(0xe1e1_e1ff));
        assert_eq!(runs[0].font_family, "Aptos");
        assert!((runs[0].font_size - 10.666_667).abs() < 0.001);
        assert!(
            runs.iter()
                .all(|run| (run.font_size - 10.666_667).abs() < 0.001)
        );
    }

    #[test]
    fn combines_odp_row_and_column_banding_after_header_axes() {
        let styles =
            br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo" xmlns:table="table">
          <office:styles>
            <style:style style:name="Body" style:family="table-cell">
              <style:table-cell-properties fo:background-color="#ebf1e9"/>
            </style:style>
            <style:style style:name="OddRow" style:family="table-cell">
              <style:table-cell-properties fo:background-color="#d5e3cf"/>
            </style:style>
            <style:style style:name="EvenRow" style:family="table-cell">
              <style:table-cell-properties/>
            </style:style>
            <style:style style:name="OddColumn" style:family="table-cell">
              <style:table-cell-properties fo:background-color="#b4c6e7"/>
            </style:style>
            <style:style style:name="EvenColumn" style:family="table-cell">
              <style:table-cell-properties/>
            </style:style>
            <style:style style:name="Header" style:family="table-cell">
              <style:table-cell-properties fo:background-color="#70ad47"/>
            </style:style>
            <style:style style:name="LastColumn" style:family="table-cell">
              <style:text-properties fo:color="#ffffff"/>
            </style:style>
            <table:table-template text:style-name="Banded">
              <table:body table:style-name="Body"/>
              <table:odd-rows table:style-name="OddRow"/>
              <table:even-rows table:style-name="EvenRow"/>
              <table:odd-columns table:style-name="OddColumn"/>
              <table:even-columns table:style-name="EvenColumn"/>
              <table:first-row table:style-name="Header"/>
              <table:first-column table:style-name="Header"/>
              <table:last-column table:style-name="LastColumn"/>
            </table:table-template>
          </office:styles>
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content =
            br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:table="table" xmlns:text="text">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame svg:x="1in" svg:y="1in" svg:width="3in" svg:height="3in">
              <table:table table:template-name="Banded"
                table:use-banding-rows-styles="true"
                table:use-banding-columns-styles="true"
                table:use-first-row-styles="true"
                table:use-first-column-styles="true"
                table:use-last-column-styles="true">
                <table:table-column/><table:table-column/><table:table-column/>
                <table:table-row>
                  <table:table-cell><text:p>h0</text:p></table:table-cell>
                  <table:table-cell><text:p>h1</text:p></table:table-cell>
                  <table:table-cell><text:p>h2</text:p></table:table-cell>
                </table:table-row>
                <table:table-row>
                  <table:table-cell><text:p>r1c0</text:p></table:table-cell>
                  <table:table-cell><text:p>r1c1</text:p></table:table-cell>
                  <table:table-cell><text:p>r1c2</text:p></table:table-cell>
                </table:table-row>
                <table:table-row>
                  <table:table-cell><text:p>r2c0</text:p></table:table-cell>
                  <table:table-cell><text:p>r2c1</text:p></table:table-cell>
                  <table:table-cell><text:p>r2c2</text:p></table:table-cell>
                </table:table-row>
                <table:table-row>
                  <table:table-cell><text:p>r3c0</text:p></table:table-cell>
                  <table:table-cell table:number-columns-spanned="2"><text:p>r3span</text:p></table:table-cell>
                  <table:covered-table-cell/>
                </table:table-row>
              </table:table>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[("styles.xml", styles), (CONTENT_PART, content)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let document = parse(&package).expect("two-axis banded ODP table");
        let fill = |text: &str| {
            let cell = document
                .objects
                .iter()
                .find(|object| object.text.as_deref() == Some(text))
                .expect("named table cell");
            let Visual::TextLayout { visual, .. } = &cell.visual else {
                panic!("table cell must retain text layout");
            };
            let Visual::RichText { fill, .. } = visual.as_ref() else {
                panic!("table cell must retain rich text");
            };
            fill.clone()
        };

        assert_eq!(fill("r1c1"), Paint::Solid(0xb4c6_e7ff));
        assert_eq!(fill("r1c2"), Paint::Solid(0xd5e3_cfff));
        assert_eq!(fill("r2c1"), Paint::Solid(0xb4c6_e7ff));
        assert_eq!(fill("r2c2"), Paint::Solid(0xebf1_e9ff));
        let spanning = document
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("r3span"))
            .expect("spanning table cell");
        let Visual::TextLayout { visual, .. } = &spanning.visual else {
            panic!("spanning cell must retain text layout");
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("spanning cell must retain rich text");
        };
        assert_eq!(runs[0].color, 0xffff_ffff);
    }

    #[test]
    fn browser_font_metrics_drive_odp_table_row_layout() {
        let styles =
            br#"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <style:style style:name="NarrowColumn" style:family="table-column">
              <style:table-column-properties style:column-width="0.35in"/>
            </style:style>
            <style:style style:name="ShortRow" style:family="table-row">
              <style:table-row-properties style:row-height="0.1in"/>
            </style:style>
            <style:style style:name="Cell" style:family="table-cell">
              <style:table-cell-properties fo:padding="0in"/>
              <style:text-properties fo:font-family="Arial" fo:font-size="10pt"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:table="table" xmlns:text="text">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame svg:x="1in" svg:y="1in" svg:width="0.35in" svg:height="2in">
              <table:table>
                <table:table-column table:style-name="NarrowColumn"/>
                <table:table-row table:style-name="ShortRow">
                  <table:table-cell table:style-name="Cell"><text:p>WWWW</text:p></table:table-cell>
                </table:table-row>
              </table:table>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[
            (
                "mimetype",
                b"application/vnd.oasis.opendocument.presentation",
            ),
            ("styles.xml", styles),
            (CONTENT_PART, content),
        ]);
        let metric_bytes = single_advance_metric_table("Arial", 'W', 1.0);
        let metrics = FontMetricTable::decode(&metric_bytes, Limits::default()).unwrap();

        let document = detect_and_parse_with_font_metrics(&bytes, Limits::default(), &metrics)
            .expect("ODP with measured font metrics")
            .expect("recognized ODP");

        let cell = document
            .objects
            .iter()
            .find(|object| object.kind == ObjectKind::Cell)
            .expect("table cell");
        assert!(cell.bounds.height > 25.0, "height={}", cell.bounds.height);
    }

    #[test]
    fn automatic_odp_rows_are_constrained_by_the_table_content_box() {
        let styles =
            br#"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
            <style:style style:name="NarrowColumn" style:family="table-column">
              <style:table-column-properties style:column-width="0.35in"/>
            </style:style>
            <style:style style:name="ShortRow" style:family="table-row">
              <style:table-row-properties style:row-height="0.1in"/>
            </style:style>
            <style:style style:name="PaddedCell" style:family="table-cell">
              <style:table-cell-properties fo:padding-top="0.05in" fo:padding-bottom="0.05in"
                fo:padding-left="0in" fo:padding-right="0in"/>
              <style:text-properties fo:font-family="Arial" fo:font-size="10pt"/>
            </style:style>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"#;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:table="table" xmlns:text="text">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame svg:x="1in" svg:y="1in" svg:width="0.35in" svg:height="0.2in">
              <table:table>
                <table:table-column table:style-name="NarrowColumn"/>
                <table:table-row table:style-name="ShortRow">
                  <table:table-cell table:style-name="PaddedCell"><text:p>WWWW</text:p></table:table-cell>
                </table:table-row>
              </table:table>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[
            (
                "mimetype",
                b"application/vnd.oasis.opendocument.presentation",
            ),
            ("styles.xml", styles),
            (CONTENT_PART, content),
        ]);
        let metric_bytes = single_advance_metric_table("Arial", 'W', 1.0);
        let metrics = FontMetricTable::decode(&metric_bytes, Limits::default()).unwrap();

        let document = detect_and_parse_with_font_metrics(&bytes, Limits::default(), &metrics)
            .expect("ODP with constrained automatic row")
            .expect("recognized ODP");
        let cell = document
            .objects
            .iter()
            .find(|object| object.kind == ObjectKind::Cell)
            .expect("table cell");

        assert!(
            (cell.bounds.height - 19.2).abs() < 0.01,
            "height={}",
            cell.bounds.height
        );
    }

    #[test]
    fn uses_cached_visual_for_layout_tables_while_preserving_cell_semantics() {
        let styles =
            br##"<office:document-styles xmlns:office="office" xmlns:style="style" xmlns:fo="fo">
          <office:automatic-styles>
            <style:page-layout style:name="PageLayout">
              <style:page-layout-properties fo:page-width="10in" fo:page-height="7.5in"/>
            </style:page-layout>
          </office:automatic-styles>
          <office:master-styles>
            <style:master-page style:name="Master" style:page-layout-name="PageLayout"/>
          </office:master-styles>
        </office:document-styles>"##;
        let content = br#"<office:document-content xmlns:office="office" xmlns:draw="draw" xmlns:svg="svg" xmlns:table="table" xmlns:text="text" xmlns:xlink="xlink">
          <office:body><office:presentation><draw:page draw:master-page-name="Master">
            <draw:frame draw:id="layout-table" svg:x="1in" svg:y="1in" svg:width="2in" svg:height="1in">
              <table:table table:template-name="{00000000-0000-0000-0000-000000000000}">
                <table:table-column/>
                <table:table-row>
                  <table:table-cell><text:p>Semantic cell</text:p></table:table-cell>
                </table:table-row>
              </table:table>
              <draw:image xlink:href="Pictures/table.png"/>
            </draw:frame>
          </draw:page></office:presentation></office:body>
        </office:document-content>"#;
        let bytes = stored_zip(&[
            ("styles.xml", styles),
            (CONTENT_PART, content),
            ("Pictures/table.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();

        let document = parse(&package).expect("layout table with cached visual");

        assert_eq!(document.objects.len(), 3);
        assert_eq!(document.objects[0].kind, ObjectKind::Image);
        assert!(matches!(document.objects[0].visual, Visual::Image { .. }));
        assert_eq!(document.objects[1].kind, ObjectKind::Table);
        assert!(matches!(document.objects[1].visual, Visual::None));
        assert_eq!(document.objects[2].kind, ObjectKind::Cell);
        assert_eq!(document.objects[2].text.as_deref(), Some("Semantic cell"));
        assert!(matches!(document.objects[2].visual, Visual::None));
    }

    #[test]
    fn preserves_embedded_video_plugin_for_browser_playback() {
        const MP4: &[u8] = b"\0\0\0\x18ftypisom\0\0\0\0";
        let content = content_xml_with_image(
            r#"<draw:plugin xlink:href="Media/movie.mp4" draw:mime-type="video/mp4"/>"#,
        );
        let parsed =
            parse_content_fixture(&[(CONTENT_PART, content.as_bytes()), ("Media/movie.mp4", MP4)])
                .expect("embedded ODP video");

        assert_eq!(parsed.objects.len(), 1);
        assert!(matches!(
            &parsed.objects[0].visual,
            Visual::Media {
                kind: crate::model::MediaKind::Video,
                media_type,
                bytes,
                poster,
            } if media_type == "video/mp4"
                && bytes == MP4
                && matches!(poster.as_ref(), Visual::None)
        ));
        assert!(parsed.diagnostics.is_empty(), "{:#?}", parsed.diagnostics);
    }

    #[test]
    fn preserves_embedded_audio_plugin_for_browser_playback() {
        const MP3: &[u8] = b"ID3\x04\0\0";
        let content = content_xml_with_image(
            r#"<draw:plugin xlink:href="Media/sound.mp3" draw:mime-type="audio/mpeg"/>"#,
        );
        let parsed =
            parse_content_fixture(&[(CONTENT_PART, content.as_bytes()), ("Media/sound.mp3", MP3)])
                .expect("embedded ODP audio");

        assert!(matches!(
            &parsed.objects[0].visual,
            Visual::Media {
                kind: crate::model::MediaKind::Audio,
                media_type,
                bytes,
                ..
            } if media_type == "audio/mpeg" && bytes == MP3
        ));
        assert!(parsed.diagnostics.is_empty(), "{:#?}", parsed.diagnostics);
    }

    #[test]
    fn corpus_master_background_and_frame_alignment() {
        for (bytes, color) in [
            (
                include_bytes!("../../tests/fixtures/corpus-background.odp").as_slice(),
                0x729fcfff,
            ),
            (
                include_bytes!("../../tests/fixtures/corpus-invalidBuAutoNumEnumValue.odp")
                    .as_slice(),
                0x127622ff,
            ),
        ] {
            let package = Package::open(bytes, Limits::default()).unwrap();
            let document = parse(&package).unwrap();
            assert!(
                document
                    .objects
                    .iter()
                    .any(|object| matches!(&object.visual,
                Visual::PaintedShape { fill: Paint::Solid(value), .. } if *value == color)),
                "master background missing"
            );
        }
        let package = Package::open(
            include_bytes!("../../tests/fixtures/corpus-TestEmbeddedFonts_DejaVu.odp"),
            Limits::default(),
        )
        .unwrap();
        let document = parse(&package).unwrap();
        assert!(document.objects.iter().any(|object| matches!(&object.visual,
            Visual::TextLayout { layout, .. } if layout.paragraphs.first().is_some_and(|p| p.align == TextAlign::Center))), "frame paragraph alignment missing");
    }

    #[test]
    fn corpus_picture_list_markers_are_images() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/corpus-invalidBuAutoNumEnumValue.odp"),
            Limits::default(),
        )
        .unwrap();
        let document = parse(&package).unwrap();
        assert_eq!(
            document
                .objects
                .iter()
                .filter(|object| object.kind == ObjectKind::Image)
                .count(),
            2
        );
        let object = document
            .objects
            .iter()
            .find(|o| {
                o.text
                    .as_deref()
                    .is_some_and(|t| t.contains("Разные маркеры"))
            })
            .unwrap();
        let Visual::TextLayout { visual, .. } = &object.visual else {
            panic!("text layout")
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("rich text")
        };
        assert!(
            runs.iter()
                .filter(|r| r.text.contains("Разные"))
                .all(|r| (r.font_size - 32.0).abs() < 0.01),
            "outline level 3 uses 24pt"
        );
    }

    fn parse_fixture(
        href: &str,
        picture: Option<(&str, &[u8])>,
        limits: Limits,
    ) -> Result<super::ParsedContent, crate::diagnostic::Diagnostic> {
        let content = content_xml(href);
        let mut entries = vec![(CONTENT_PART, content.as_bytes())];
        if let Some(picture) = picture {
            entries.push(picture);
        }
        let bytes = stored_zip(&entries);
        let package = Package::open(&bytes, limits)?;
        parse_content(&package, &page_sizes(), &StyleCatalog::default())
    }

    fn parse_content_fixture(
        entries: &[(&str, &[u8])],
    ) -> Result<super::ParsedContent, crate::diagnostic::Diagnostic> {
        let bytes = stored_zip(entries);
        let package = Package::open(&bytes, Limits::default())?;
        parse_content(&package, &page_sizes(), &StyleCatalog::default())
    }

    fn page_sizes() -> HashMap<String, PageSize> {
        HashMap::from([(
            "Master".to_owned(),
            PageSize {
                width: 960.0,
                height: 720.0,
            },
        )])
    }

    fn content_xml(href: &str) -> String {
        content_xml_with_image(&format!(r#"<draw:image xlink:href="{href}"/>"#))
    }

    fn content_xml_with_image(image: &str) -> String {
        format!(
            "<office:document-content xmlns:office=\"office\" xmlns:draw=\"draw\" xmlns:svg=\"svg\" xmlns:xlink=\"xlink\"><office:body><office:presentation><draw:page draw:master-page-name=\"Master\"><draw:frame draw:id=\"picture-1\" svg:x=\"1in\" svg:y=\"2in\" svg:width=\"3in\" svg:height=\"4in\">{image}</draw:frame></draw:page></office:presentation></office:body></office:document-content>"
        )
    }
}
