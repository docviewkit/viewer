//! Native PPTX presentation and slide parsing.
use super::drawingml::drawingml_shape_transform as shape_transform;

#[path = "drawingml_presets.rs"]
mod drawingml_presets;

use std::cell::RefCell;
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

use super::omml::{OmmlCapture, OmmlNode};
use super::optional_xml_attribute as string_attribute;
use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
use crate::font_metrics::FontMetricTable;
use crate::model::{
    AffineTransform, Backdrop3D, Document, DocumentFormat, DocumentKind, FillRule, Geometry, Glow,
    GradientStop, ImageAdjustment, ImageCrop, LineAlignment, LineCap, LineCompound, LineJoin,
    MappingQuality, MediaKind, Object, ObjectKind, OuterShadow, Paint, PathCommand, PathFillMode,
    PathLayer, PptxAction, PptxObjectMetadata, Rect, Reflection, Shadow, SlideMetadata,
    SourceLocator, SourceRef, SpeakerNoteParagraph, StrokeStyle, TextAlign, TextAutoFit,
    TextDirection, TextEffect, TextHorizontalOverflow, TextLayout, TextOrientation,
    TextParagraphLayout, TextRun, TextVerticalAlign, TextVerticalOverflow, ThreeDStyle, Unit,
    UnitKind, Visual, VisualBrushChild, leading_transform,
};
use crate::package::{Package, Relationship};
use crate::xml::{XmlAttribute, XmlEvent, decode_xml_text};

fn parse_xml<F>(
    input: &[u8],
    limits: crate::limits::Limits,
    on_event: F,
) -> Result<crate::xml::XmlSummary, Diagnostic>
where
    F: for<'a> FnMut(XmlEvent<'a>) -> Result<(), Diagnostic>,
{
    // VML ink is read separately; linked OLE VML shapes still need their DrawingML fallback.
    crate::xml::parse_ooxml_excluding(input, limits, &["urn:schemas-microsoft-com:vml"], on_event)
}

use super::drawingml::{
    Chart as BasicChart, ChartKind as BasicChartKind, ChartLegendPosition, Diagram as PptxDiagram,
    DrawingMlDashPattern, DrawingMlPictureEffectsCapture, DrawingMlRelativeRectangle,
    DrawingMlThemeLineStyles, apply_color_transform as apply_drawingml_color_transform,
    chart_3d_series_axis_labels, chart_area_3d_axis_labels, chart_area_3d_faces,
    chart_area_geometry, chart_bar_3d_plane_bounds, chart_bar_3d_walls, chart_bar_cone_cap,
    chart_bar_cone_to_max_geometry, chart_bar_cylinder_cap, chart_bar_data_label,
    chart_bar_geometry, chart_bar_segment_bounds, chart_bubble_bounds, chart_bubble_paint,
    chart_data_table_layout, chart_line_geometry, chart_marker_geometry, chart_pie_label_layout,
    chart_pie_side_paint, chart_pie_slices, chart_series_line_geometries, chart_title_bounds,
    drawingml_dash_lengths, drawingml_fallback_character_width, drawingml_fill_reference_has_paint,
    drawingml_outer_shadow, drawingml_outer_shadow_is_identity, drawingml_relative_rectangle,
    drawingml_shadow_alignment, drawingml_stroke_style, drawingml_text_orientation,
    ellipse_arc_bezier_points, format_axis_value, linear_to_srgb, numeric_attribute,
    parse_chart_with_theme as parse_drawingml_chart, parse_diagram as parse_pptx_diagram,
    parse_drawingml_reflection, parse_rgb_color, parse_three_d_backdrop_point, parse_three_d_bevel,
    parse_three_d_camera, parse_three_d_light, parse_three_d_rotation, parse_three_d_shape,
    radar_geometry, signed_numeric_attribute,
};
use super::embedded_media::{
    EmbeddedMediaError, embedded_media_type, embedded_media_type_from_mime,
};
use super::presentation_image::{
    ImageCacheEntry, OfficeImageError, clone_image_bytes, office_image_media_type,
    office_image_media_type_from_mime, reserve_materialized_image_bytes,
};
use super::{
    ContentTypes, diamond_geometry, local_name, normalize_symbol_font_character, polygon_geometry,
    regular_polygon_geometry, star_geometry,
};
use drawingml_presets::PRESET_SHAPE_DEFINITIONS;

const EMU_PER_CSS_PIXEL: f32 = 9_525.0;
const POINTS_TO_CSS_PIXELS: f32 = 96.0 / 72.0;
const TEXT_LEVEL_COUNT: usize = 9;
const TABLE_STYLES_PART: &str = "ppt/tableStyles.xml";

#[derive(Clone, Debug)]
enum ParagraphSpacing {
    Percent(f32),
    Points(f32),
}

impl ParagraphSpacing {
    fn resolve(&self, font_size: f32) -> f32 {
        match self {
            Self::Percent(ratio) => font_size * ratio,
            Self::Points(value) => *value,
        }
    }

    fn resolve_for_layout(&self, font_size: Option<f32>) -> f32 {
        match (self, font_size) {
            // DrawingML percentage line spacing scales the font's natural line
            // box. Canvas exposes only the authored em size here, so use the
            // same 1.2-em natural line box as the renderer's unstyled text.
            (Self::Percent(ratio), Some(font_size)) => font_size * 1.2 * ratio,
            (Self::Percent(_), None) => 0.0,
            (Self::Points(value), _) => *value,
        }
    }
}

#[derive(Clone, Debug)]
enum ParagraphBullet {
    None,
    Character(String),
    AutoNumber { kind: String, start_at: u32 },
    Image(String),
}

#[derive(Clone, Debug)]
enum ParagraphBulletSize {
    Percent(f32),
    Points(f32),
}

impl ParagraphBulletSize {
    fn resolve(&self, text_font_size: f32) -> f32 {
        match self {
            Self::Percent(ratio) => text_font_size * ratio,
            Self::Points(value) => *value,
        }
    }
}

#[derive(Clone, Debug, Default)]
struct ParagraphStyle {
    font_paint: Option<Box<Paint>>,
    font_family: Option<String>,
    font_east_asian: Option<String>,
    font_complex_script: Option<String>,
    font_size: Option<f32>,
    font_color: Option<u32>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strikethrough: Option<bool>,
    baseline_ratio: Option<f32>,
    letter_spacing: Option<f32>,
    align: Option<TextAlign>,
    margin_left: Option<f32>,
    margin_right: Option<f32>,
    indent: Option<f32>,
    default_tab_stop: Option<f32>,
    line_spacing: Option<ParagraphSpacing>,
    space_before: Option<ParagraphSpacing>,
    space_after: Option<ParagraphSpacing>,
    latin_line_break: Option<bool>,
    hanging_punctuation: Option<bool>,
    east_asian_line_breaks: Option<bool>,
    capitalization: Option<TextCapitalization>,
    bullet: Option<ParagraphBullet>,
    bullet_font_family: Option<String>,
    bullet_size: Option<ParagraphBulletSize>,
    bullet_color: Option<u32>,
    text_effect: Option<TextEffect>,
}

impl ParagraphStyle {
    fn resolve_bullet_indentation(&mut self, direct: &Self) {
        if matches!(self.bullet, None | Some(ParagraphBullet::None)) {
            return;
        }
        let Some(indent) = self.indent else {
            return;
        };
        // A locally specified bullet/indent pair uses its own margin (zero when
        // omitted), rather than borrowing the master list's body anchor.
        let margin = if direct.indent.is_some()
            && !matches!(direct.bullet, None | Some(ParagraphBullet::None))
        {
            direct.margin_left.unwrap_or(0.0)
        } else {
            self.margin_left.unwrap_or(0.0)
        };
        // PowerPoint keeps the marker on the left for either sign and preserves
        // the full gap even when a hanging marker would cross the shape edge.
        self.margin_left = Some(margin.min(margin + indent).max(0.0) + indent.abs());
        self.indent = Some(-indent.abs());
    }

    fn apply(&mut self, overrides: &Self) {
        if overrides.font_family.is_some() {
            self.font_family.clone_from(&overrides.font_family);
        }
        if overrides.font_east_asian.is_some() {
            self.font_east_asian.clone_from(&overrides.font_east_asian);
        }
        if overrides.font_complex_script.is_some() {
            self.font_complex_script
                .clone_from(&overrides.font_complex_script);
        }
        if overrides.font_size.is_some() {
            self.font_size = overrides.font_size;
        }
        if overrides.font_color.is_some() {
            self.font_color = overrides.font_color;
            self.font_paint = None;
        }
        if overrides.font_paint.is_some() {
            self.font_paint.clone_from(&overrides.font_paint);
        }
        if overrides.bold.is_some() {
            self.bold = overrides.bold;
        }
        if overrides.italic.is_some() {
            self.italic = overrides.italic;
        }
        if overrides.underline.is_some() {
            self.underline = overrides.underline;
        }
        if overrides.strikethrough.is_some() {
            self.strikethrough = overrides.strikethrough;
        }
        if overrides.baseline_ratio.is_some() {
            self.baseline_ratio = overrides.baseline_ratio;
        }
        if overrides.letter_spacing.is_some() {
            self.letter_spacing = overrides.letter_spacing;
        }
        if overrides.align.is_some() {
            self.align = overrides.align;
        }
        if overrides.margin_left.is_some() {
            self.margin_left = overrides.margin_left;
        }
        if overrides.margin_right.is_some() {
            self.margin_right = overrides.margin_right;
        }
        if overrides.indent.is_some() {
            self.indent = overrides.indent;
        }
        if overrides.default_tab_stop.is_some() {
            self.default_tab_stop = overrides.default_tab_stop;
        }
        if overrides.line_spacing.is_some() {
            self.line_spacing.clone_from(&overrides.line_spacing);
        }
        if overrides.space_before.is_some() {
            self.space_before.clone_from(&overrides.space_before);
        }
        if overrides.space_after.is_some() {
            self.space_after.clone_from(&overrides.space_after);
        }
        if overrides.latin_line_break.is_some() {
            self.latin_line_break = overrides.latin_line_break;
        }
        if overrides.hanging_punctuation.is_some() {
            self.hanging_punctuation = overrides.hanging_punctuation;
        }
        if overrides.east_asian_line_breaks.is_some() {
            self.east_asian_line_breaks = overrides.east_asian_line_breaks;
        }
        if overrides.capitalization.is_some() {
            self.capitalization = overrides.capitalization;
        }
        if overrides.bullet.is_some() {
            self.bullet.clone_from(&overrides.bullet);
        }
        if overrides.bullet_font_family.is_some() {
            self.bullet_font_family
                .clone_from(&overrides.bullet_font_family);
        }
        if overrides.bullet_size.is_some() {
            self.bullet_size.clone_from(&overrides.bullet_size);
        }
        if overrides.bullet_color.is_some() {
            self.bullet_color = overrides.bullet_color;
        }
        if overrides.text_effect.is_some() {
            self.text_effect.clone_from(&overrides.text_effect);
        }
    }
}

#[derive(Clone, Debug)]
struct MasterTextStyles {
    title: Vec<ParagraphStyle>,
    body: Vec<ParagraphStyle>,
    other: Vec<ParagraphStyle>,
}

impl Default for MasterTextStyles {
    fn default() -> Self {
        Self {
            title: vec![ParagraphStyle::default(); TEXT_LEVEL_COUNT],
            body: vec![ParagraphStyle::default(); TEXT_LEVEL_COUNT],
            other: vec![ParagraphStyle::default(); TEXT_LEVEL_COUNT],
        }
    }
}

impl MasterTextStyles {
    fn for_placeholder(&self, placeholder_type: Option<&str>) -> &[ParagraphStyle] {
        match placeholder_type {
            Some("title" | "ctrTitle") => &self.title,
            Some("body" | "obj" | "subTitle") => &self.body,
            _ => &self.other,
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum ParagraphStyleTarget {
    Default,
    Level(usize),
    Current,
}

#[derive(Clone, Copy, Debug)]
enum ParagraphSpacingTarget {
    Line,
    Before,
    After,
}

#[derive(Debug)]
struct ParagraphStyleCapture {
    depth: usize,
    target: ParagraphStyleTarget,
    style: ParagraphStyle,
    spacing: Option<(usize, ParagraphSpacingTarget)>,
    bullet_color_depth: Option<usize>,
    text_effect_depth: Option<usize>,
}

pub(super) fn parse(
    package: &Package<'_>,
    main_part: &str,
    diagnostics: Vec<Diagnostic>,
    content_types: &ContentTypes,
    font_metrics: &FontMetricTable,
) -> Result<Document, Diagnostic> {
    parse_with_mode(
        package,
        main_part,
        diagnostics,
        content_types,
        font_metrics,
        PptxParseMode::Complete,
        0,
    )
    .map(|initial| initial.document)
}

pub(super) struct PptxInitial {
    pub(super) document: Document,
    pub(super) loaded_units: BTreeSet<u32>,
    pub(super) materialized_image_bytes: usize,
}

pub(super) fn parse_initial(
    package: &Package<'_>,
    main_part: &str,
    diagnostics: Vec<Diagnostic>,
    content_types: &ContentTypes,
    font_metrics: &FontMetricTable,
) -> Result<PptxInitial, Diagnostic> {
    parse_with_mode(
        package,
        main_part,
        diagnostics,
        content_types,
        font_metrics,
        PptxParseMode::Initial,
        0,
    )
}

pub(super) fn parse_unit(
    package: &Package<'_>,
    main_part: &str,
    content_types: &ContentTypes,
    font_metrics: &FontMetricTable,
    unit_index: u32,
    materialized_image_bytes: usize,
) -> Result<(Document, usize), Diagnostic> {
    parse_with_mode(
        package,
        main_part,
        Vec::new(),
        content_types,
        font_metrics,
        PptxParseMode::Unit(unit_index),
        materialized_image_bytes,
    )
    .map(|initial| (initial.document, initial.materialized_image_bytes))
}

#[derive(Clone, Copy)]
enum PptxParseMode {
    Complete,
    Initial,
    Unit(u32),
}

fn parse_with_mode(
    package: &Package<'_>,
    main_part: &str,
    diagnostics: Vec<Diagnostic>,
    content_types: &ContentTypes,
    font_metrics: &FontMetricTable,
    mode: PptxParseMode,
    materialized_image_bytes: usize,
) -> Result<PptxInitial, Diagnostic> {
    let presentation = package.required_part(main_part)?;
    let mut slide_relationship_ids = Vec::new();
    let mut slide_width = None;
    let mut slide_height = None;
    let mut first_slide_number = 1_u32;
    let mut found_presentation = false;
    parse_xml(&presentation, package.limits(), |event| {
        let XmlEvent::StartElement {
            name, attributes, ..
        } = event
        else {
            return Ok(());
        };
        match local_name(name) {
            "presentation" if !found_presentation => {
                first_slide_number = numeric_attribute(&attributes, "firstSlideNum", main_part)?
                    .map(u32::try_from)
                    .transpose()
                    .map_err(|_| format_error(main_part, "first slide number exceeds u32 range"))?
                    .unwrap_or(1);
                found_presentation = true;
            }
            "sldId" => {
                if let Some(id) = relationship_id(&attributes, main_part)? {
                    let source_id = string_attribute(&attributes, "id", main_part)?;
                    slide_relationship_ids.push((id, source_id));
                }
            }
            "sldSz" => {
                slide_width = numeric_attribute(&attributes, "cx", main_part)?;
                slide_height = numeric_attribute(&attributes, "cy", main_part)?;
            }
            _ => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, main_part))?;

    let slide_width = slide_width
        .ok_or_else(|| format_error(main_part, "presentation is missing p:sldSz/@cx"))?
        as f32
        / EMU_PER_CSS_PIXEL;
    let slide_height = slide_height
        .ok_or_else(|| format_error(main_part, "presentation is missing p:sldSz/@cy"))?
        as f32
        / EMU_PER_CSS_PIXEL;
    if !slide_width.is_finite()
        || !slide_height.is_finite()
        || slide_width <= 0.0
        || slide_height <= 0.0
    {
        return Err(format_error(
            main_part,
            "presentation slide size is invalid",
        ));
    }

    let relationships = package.relationships(Some(main_part))?;
    let mut diagnostics = diagnostics;
    let default_theme_part =
        presentation_fallback_theme_part(&relationships, main_part, &mut diagnostics);
    let mut theme_cache = HashMap::new();
    let default_theme = if let Some(theme_part) = default_theme_part {
        let theme = parse_pptx_theme(package, &theme_part)?;
        theme_cache.insert(theme_part, theme.clone());
        theme
    } else {
        PptxTheme::default()
    };
    let default_text_styles =
        parse_default_text_styles(&presentation, package.limits(), main_part, &default_theme)?;
    let relationship_map: HashMap<&str, &Relationship> = relationships
        .iter()
        .map(|relationship| (relationship.id.as_str(), relationship))
        .collect();
    let mut units = Vec::new();
    let mut state = PptxParseState {
        objects: Vec::new(),
        next_z: 0,
        diagnostics,
        image_cache: HashMap::new(),
        materialized_image_bytes,
        reported_unsupported_backgrounds: HashSet::new(),
        placeholders: HashMap::new(),
        placeholder_text_styles: HashMap::new(),
        placeholder_presets: HashMap::new(),
        pending_placeholder_objects: HashSet::new(),
        default_theme,
        default_text_styles,
        theme_cache,
    };
    let mut parsed_slide_parts = HashSet::new();
    let mut loaded_units = BTreeSet::new();
    for (index, (relationship_id, source_id)) in slide_relationship_ids.iter().enumerate() {
        let relationship = relationship_map
            .get(relationship_id.as_str())
            .ok_or_else(|| {
                format_error(
                    main_part,
                    format!("slide relationship {relationship_id} does not exist"),
                )
            })?;
        if relationship.external {
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::ExternalResourceBlocked,
                    Phase::Security,
                    Fidelity::Blocked,
                    "external slide relationship was blocked",
                )
                .in_part(main_part),
            );
            continue;
        }
        if !relationship.type_uri.ends_with("/slide") {
            return Err(format_error(
                main_part,
                format!("relationship {relationship_id} is not a slide"),
            ));
        }
        let slide_part = relationship.target.as_str();
        if !parsed_slide_parts.insert(slide_part) {
            return Err(format_error(
                main_part,
                format!("multiple slides reference the same package part {slide_part}"),
            ));
        }
        let slide_properties = presentation_part_properties(package, slide_part)?;
        let (speaker_notes, speaker_notes_part, speaker_note_paragraphs) =
            parse_speaker_notes(package, slide_part, &state.default_theme)?;
        let slide_number = first_slide_number.saturating_add(index as u32);
        units.push(Unit {
            kind: UnitKind::Slide,
            index: index as u32,
            id: format!("unit:{index}"),
            name: slide_properties
                .name
                .clone()
                .unwrap_or_else(|| format!("Slide {slide_number}")),
            width: slide_width,
            height: slide_height,
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
                source_id: source_id.clone(),
                source_part: Some(slide_part.to_owned()),
                speaker_notes,
                speaker_notes_part,
                speaker_note_paragraphs,
                number: slide_number,
                hidden: !slide_properties.shown,
            }),
        });
        let unit_index = index as u32;
        let materialize = match mode {
            PptxParseMode::Complete => true,
            PptxParseMode::Initial => {
                unit_index == 0 || presentation_part_contains_table(package, slide_part)?
            }
            PptxParseMode::Unit(target) => unit_index == target,
        };
        if materialize {
            loaded_units.insert(unit_index);
            parse_slide(
                package,
                SlideParseContext {
                    part: slide_part,
                    unit_index,
                    properties: slide_properties,
                    bounds: Rect {
                        x: 0.0,
                        y: 0.0,
                        width: slide_width,
                        height: slide_height,
                    },
                },
                &mut state,
                content_types,
                font_metrics,
            )?;
        }
    }
    if units.is_empty() {
        return Err(format_error(
            main_part,
            "presentation contains no readable slides",
        ));
    }
    let embedded_fonts = if matches!(mode, PptxParseMode::Unit(_)) {
        Vec::new()
    } else {
        super::embedded_font::extract_pptx(
            package,
            main_part,
            &presentation,
            &mut state.diagnostics,
        )
    };

    let materialized_image_bytes = state.materialized_image_bytes;
    Ok(PptxInitial {
        document: Document {
            fatal: false,
            format: Some(DocumentFormat::Pptx),
            kind: Some(DocumentKind::Presentation),
            units,
            outline: Vec::new(),
            objects: state.objects,
            embedded_fonts,
            font_alternate_names: Vec::new(),
            diagnostics: state.diagnostics,
        },
        loaded_units,
        materialized_image_bytes,
    })
}

fn presentation_part_contains_table(package: &Package<'_>, part: &str) -> Result<bool, Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut contains_table = false;
    parse_xml(&bytes, package.limits(), |event| {
        if let XmlEvent::StartElement { name, .. } = event
            && local_name(name) == "tbl"
        {
            contains_table = true;
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(contains_table)
}

#[derive(Clone, Copy, Debug)]
enum DrawingMlLineEndKind {
    Arrow,
    Stealth,
    Triangle,
    Diamond,
    Oval,
}

#[derive(Clone, Copy, Debug, Default)]
enum DrawingMlLineEndSize {
    Small,
    #[default]
    Medium,
    Large,
}

impl DrawingMlLineEndSize {
    fn from_attribute(value: Option<&str>) -> Self {
        match value {
            Some("sm") => Self::Small,
            Some("lg") => Self::Large,
            _ => Self::Medium,
        }
    }

    const fn stroke_multiplier(self) -> f32 {
        // PowerPoint renders sm/med/lg line ends at 2/3/5 stroke widths.
        match self {
            Self::Small => 2.0,
            Self::Medium => 3.0,
            Self::Large => 5.0,
        }
    }
}

#[derive(Clone, Copy, Debug)]
pub(super) struct DrawingMlLineEnd {
    kind: DrawingMlLineEndKind,
    width: DrawingMlLineEndSize,
    length: DrawingMlLineEndSize,
}

pub(super) fn drawingml_line_end(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<Option<DrawingMlLineEnd>, Diagnostic> {
    let kind = match string_attribute(attributes, "type", part)?.as_deref() {
        Some("arrow") => DrawingMlLineEndKind::Arrow,
        Some("stealth") => DrawingMlLineEndKind::Stealth,
        Some("triangle") => DrawingMlLineEndKind::Triangle,
        Some("diamond") => DrawingMlLineEndKind::Diamond,
        Some("oval") => DrawingMlLineEndKind::Oval,
        _ => return Ok(None),
    };
    let width = string_attribute(attributes, "w", part)?;
    let length = string_attribute(attributes, "len", part)?;
    Ok(Some(DrawingMlLineEnd {
        kind,
        width: DrawingMlLineEndSize::from_attribute(width.as_deref()),
        length: DrawingMlLineEndSize::from_attribute(length.as_deref()),
    }))
}

fn non_visual_metadata(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<PptxObjectMetadata, Diagnostic> {
    Ok(PptxObjectMetadata {
        name: string_attribute(attributes, "name", part)?.filter(|value| !value.is_empty()),
        title: string_attribute(attributes, "title", part)?.filter(|value| !value.is_empty()),
        description: string_attribute(attributes, "descr", part)?.filter(|value| !value.is_empty()),
        hidden: boolean_attribute(attributes, "hidden", part)?.unwrap_or(false),
        click_action: None,
        hover_action: None,
    })
}

fn slide_jump_target(action: &str) -> Option<String> {
    let jump = action
        .strip_prefix("ppaction://hlinkshowjump?jump=")?
        .to_ascii_lowercase();
    Some(
        match jump.as_str() {
            "nextslide" => "next",
            "previousslide" => "previous",
            "firstslide" => "first",
            "lastslide" => "last",
            "endshow" => "end",
            _ => jump.as_str(),
        }
        .to_owned(),
    )
}

fn drawingml_action(
    attributes: &[XmlAttribute<'_>],
    relationships: &HashMap<&str, &Relationship>,
    part: &str,
) -> Result<PptxAction, Diagnostic> {
    let action = string_attribute(attributes, "action", part)?.filter(|value| !value.is_empty());
    let tooltip = string_attribute(attributes, "tooltip", part)?.filter(|value| !value.is_empty());
    let relationship = relationship_id(attributes, part)?
        .as_deref()
        .and_then(|id| relationships.get(id))
        .copied();
    let relationship_target = relationship.map(|relationship| relationship.target.clone());
    let jump_target = action.as_deref().and_then(slide_jump_target);
    let kind = if jump_target.is_some()
        || relationship.is_some_and(|relationship| relationship.type_uri.ends_with("/slide"))
    {
        "slide"
    } else if relationship.is_some_and(|relationship| relationship.type_uri.ends_with("/hyperlink"))
    {
        "hyperlink"
    } else if action.is_some() {
        "command"
    } else {
        "unknown"
    };
    Ok(PptxAction {
        kind: kind.to_owned(),
        action,
        target: jump_target.or(relationship_target),
        tooltip,
    })
}

#[derive(Debug)]
struct ShapeState {
    depth: usize,
    is_connector: bool,
    shape_id: Option<u32>,
    metadata: PptxObjectMetadata,
    non_visual_depth: Option<usize>,
    x: Option<i64>,
    y: Option<i64>,
    width: Option<u64>,
    height: Option<u64>,
    preset: Option<String>,
    preset_adjustments: HashMap<String, f32>,
    geometry: Geometry,
    fill: Paint,
    style_fill: Option<Paint>,
    explicit_fill: bool,
    stroke: Paint,
    style_stroke: Option<Paint>,
    explicit_stroke: bool,
    stroke_width: f32,
    style_stroke_width: Option<f32>,
    explicit_stroke_width: bool,
    shape_properties_depth: Option<usize>,
    line_depth: Option<usize>,
    head_arrow: Option<DrawingMlLineEnd>,
    tail_arrow: Option<DrawingMlLineEnd>,
    dash_pattern: DrawingMlDashPattern,
    line_cap: LineCap,
    line_join: LineJoin,
    line_compound: LineCompound,
    line_alignment: LineAlignment,
    miter_limit: f32,
    paint_capture: Option<(PaintTarget, PaintCapture)>,
    image_fill: Option<ShapeImageFillState>,
    use_group_fill: bool,
    transform_depth: Option<usize>,
    text_transform_depth: Option<usize>,
    text_x: Option<i64>,
    text_y: Option<i64>,
    text_width: Option<u64>,
    text_height: Option<u64>,
    rotation_degrees: f32,
    flip_horizontal: bool,
    flip_vertical: bool,
    text: String,
    collecting_text: bool,
    font_family: String,
    font_east_asian: Option<String>,
    font_complex_script: Option<String>,
    font_size: f32,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    highlight: u32,
    baseline_shift: f32,
    font_color: u32,
    font_paint: Option<Box<Paint>>,
    explicit_text_color: bool,
    diagram_text_color: Option<u32>,
    letter_spacing: f32,
    east_asian_line_breaks: bool,
    capitalization: TextCapitalization,
    paragraph_count: u32,
    paragraph_styles: Vec<ParagraphStyle>,
    paragraph_style_capture: Option<ParagraphStyleCapture>,
    current_paragraph_style: ParagraphStyle,
    current_paragraph_level: usize,
    paragraph_pending: bool,
    pending_bullet_style: Option<usize>,
    numbering_counters: [u32; TEXT_LEVEL_COUNT],
    paragraph_layout_style: Option<ParagraphStyle>,
    paragraph_layouts: Vec<TextParagraphLayout>,
    paragraph_layout_styles: Vec<ParagraphStyle>,
    paragraph_run_starts: Vec<usize>,
    run: Option<TextRunState>,
    runs: Vec<TextRun>,
    run_effects: Vec<TextEffect>,
    inherited_run_effect: TextEffect,
    run_properties_depth: Option<usize>,
    run_properties_target: Option<RunPropertiesTarget>,
    text_body_properties_depth: Option<usize>,
    fill_reference_depth: Option<usize>,
    line_reference_depth: Option<usize>,
    font_reference_depth: Option<usize>,
    align: TextAlign,
    vertical_align: Option<TextVerticalAlign>,
    text_direction: TextDirection,
    text_orientation: TextOrientation,
    text_auto_fit: TextAutoFit,
    text_min_scale: f32,
    text_font_scale: f32,
    text_line_spacing_reduction: f32,
    text_column_count: u32,
    text_column_spacing: f32,
    text_rotation_degrees: f32,
    text_horizontal_overflow: TextHorizontalOverflow,
    text_vertical_overflow: TextVerticalOverflow,
    text_wrap: bool,
    text_warp: Option<String>,
    text_inset_left: f32,
    text_inset_right: f32,
    text_inset_top: f32,
    text_inset_bottom: f32,
    text_space_first_last_paragraph: bool,
    shadow: Option<Shadow>,
    outer_shadow: Option<OuterShadow>,
    inner_shadow: Option<Shadow>,
    glow: Option<Glow>,
    reflection: Option<Reflection>,
    soft_edge: Option<f32>,
    three_d: Option<ThreeDStyle>,
    scene_3d_depth: Option<usize>,
    backdrop_3d_depth: Option<usize>,
    camera_3d_depth: Option<usize>,
    light_3d_depth: Option<usize>,
    shape_3d_depth: Option<usize>,
    shadow_kind: PptxShadowKind,
    shadow_depth: Option<usize>,
    shadow_color: Option<u32>,
    shadow_blur: f32,
    shadow_distance: f32,
    shadow_direction_degrees: f32,
    custom_geometry: Option<CustomGeometryState>,
    math_capture: Option<OmmlCapture>,
    math_root: Option<(OmmlNode, String)>,
    math_present: bool,
    omit_empty_math: bool,
    placeholder_type: Option<String>,
    placeholder_index: Option<u32>,
    is_custom_prompt: bool,
}

impl ShapeState {
    fn new(depth: usize, theme: &PptxTheme, is_connector: bool) -> Self {
        Self {
            depth,
            is_connector,
            shape_id: None,
            metadata: PptxObjectMetadata::default(),
            non_visual_depth: None,
            x: None,
            y: None,
            width: None,
            height: None,
            preset: None,
            preset_adjustments: HashMap::new(),
            geometry: Geometry::Rectangle,
            fill: Paint::None,
            style_fill: None,
            explicit_fill: false,
            stroke: Paint::None,
            style_stroke: None,
            explicit_stroke: false,
            stroke_width: 0.0,
            style_stroke_width: None,
            explicit_stroke_width: false,
            shape_properties_depth: None,
            line_depth: None,
            head_arrow: None,
            tail_arrow: None,
            dash_pattern: DrawingMlDashPattern::Solid,
            line_cap: LineCap::Flat,
            line_join: LineJoin::Miter,
            line_compound: LineCompound::Single,
            line_alignment: LineAlignment::Center,
            miter_limit: 10.0,
            paint_capture: None,
            image_fill: None,
            use_group_fill: false,
            transform_depth: None,
            text_transform_depth: None,
            text_x: None,
            text_y: None,
            text_width: None,
            text_height: None,
            rotation_degrees: 0.0,
            flip_horizontal: false,
            flip_vertical: false,
            text: String::new(),
            collecting_text: false,
            font_family: theme.resolve_typeface("+mn-lt", "latin"),
            font_east_asian: None,
            font_complex_script: None,
            font_size: 18.0 * POINTS_TO_CSS_PIXELS,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            highlight: 0,
            baseline_shift: 0.0,
            font_paint: None,
            font_color: 0x0000_00ff,
            explicit_text_color: false,
            diagram_text_color: None,
            letter_spacing: 0.0,
            east_asian_line_breaks: true,
            capitalization: TextCapitalization::None,
            paragraph_count: 0,
            paragraph_styles: vec![ParagraphStyle::default(); TEXT_LEVEL_COUNT],
            paragraph_style_capture: None,
            current_paragraph_style: ParagraphStyle::default(),
            current_paragraph_level: 0,
            paragraph_pending: false,
            pending_bullet_style: None,
            numbering_counters: [0; TEXT_LEVEL_COUNT],
            paragraph_layout_style: None,
            paragraph_layouts: Vec::new(),
            paragraph_layout_styles: Vec::new(),
            paragraph_run_starts: Vec::new(),
            run: None,
            runs: Vec::new(),
            run_effects: Vec::new(),
            inherited_run_effect: TextEffect::default(),
            run_properties_depth: None,
            run_properties_target: None,
            text_body_properties_depth: None,
            fill_reference_depth: None,
            line_reference_depth: None,
            font_reference_depth: None,
            align: TextAlign::Start,
            vertical_align: None,
            text_direction: TextDirection::Auto,
            text_orientation: TextOrientation::Horizontal,
            text_auto_fit: TextAutoFit::None,
            text_min_scale: 0.1,
            text_font_scale: 1.0,
            text_line_spacing_reduction: 0.0,
            text_column_count: 1,
            text_column_spacing: 0.0,
            text_rotation_degrees: 0.0,
            text_horizontal_overflow: TextHorizontalOverflow::Overflow,
            text_vertical_overflow: TextVerticalOverflow::Overflow,
            text_wrap: true,
            text_warp: None,
            text_inset_left: 9.6,
            text_inset_right: 9.6,
            text_inset_top: 4.8,
            text_inset_bottom: 4.8,
            text_space_first_last_paragraph: false,
            shadow: None,
            outer_shadow: None,
            inner_shadow: None,
            glow: None,
            reflection: None,
            soft_edge: None,
            three_d: None,
            scene_3d_depth: None,
            backdrop_3d_depth: None,
            camera_3d_depth: None,
            light_3d_depth: None,
            shape_3d_depth: None,
            shadow_kind: PptxShadowKind::Outer,
            shadow_depth: None,
            shadow_color: None,
            shadow_blur: 0.0,
            shadow_distance: 0.0,
            shadow_direction_degrees: 0.0,
            custom_geometry: None,
            math_capture: None,
            math_root: None,
            math_present: false,
            omit_empty_math: false,
            placeholder_type: None,
            placeholder_index: None,
            is_custom_prompt: false,
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum PptxShadowKind {
    Outer,
    Inner,
    Glow,
}

#[derive(Clone, Debug)]
struct ShapeImageFillState {
    target: PaintTarget,
    mapping: crate::model::ImageFillMapping,
    depth: usize,
    embedded_relationship_id: Option<String>,
    preferred_svg_relationship_id: Option<String>,
    linked_relationship_id: Option<String>,
    crop: ImageCrop,
    tile: bool,
}

impl ShapeImageFillState {
    fn new(depth: usize) -> Self {
        Self {
            target: PaintTarget::Fill,
            mapping: crate::model::ImageFillMapping::default(),
            depth,
            embedded_relationship_id: None,
            preferred_svg_relationship_id: None,
            linked_relationship_id: None,
            crop: ImageCrop::default(),
            tile: false,
        }
    }
    fn start(
        &mut self,
        local: &str,
        attributes: &[XmlAttribute<'_>],
        part: &str,
    ) -> Result<(), Diagnostic> {
        super::drawingml::drawingml_image_fill_mapping(&mut self.mapping, local, attributes, part)?;
        match local {
            "blip" => {
                let embedded = string_attribute(attributes, "embed", part)?;
                let linked = string_attribute(attributes, "link", part)?;
                if embedded.is_some() && linked.is_some() {
                    return Err(format_error(
                        part,
                        "shape image fill cannot be both embedded and linked",
                    ));
                }
                if let Some(embedded) = embedded {
                    self.embedded_relationship_id = Some(embedded);
                }
                if let Some(linked) = linked {
                    self.linked_relationship_id = Some(linked);
                }
            }
            "svgBlip" => {
                self.preferred_svg_relationship_id = string_attribute(attributes, "embed", part)?;
            }
            "srcRect" => self.crop = super::drawingml::drawingml_picture_crop(attributes, part)?,
            "tile" => self.tile = true,
            "fillRect" => {}
            _ => {}
        }
        Ok(())
    }
}

#[derive(Clone, Copy, Debug)]
enum CustomPathCommandKind {
    Move,
    Line,
    Quadratic,
    Bezier,
}

#[derive(Debug)]
struct PendingCustomPathCommand {
    depth: usize,
    kind: CustomPathCommandKind,
    points: Vec<(f32, f32)>,
}

#[derive(Debug)]
struct CustomPathLayer {
    start: usize,
    end: usize,
    fill: PathFillMode,
    stroke: bool,
}

#[derive(Debug)]
pub(super) struct CustomGeometryState {
    depth: usize,
    default_width: f32,
    default_height: f32,
    width: f32,
    height: f32,
    formulas: HashMap<String, String>,
    guides: Vec<(String, String)>,
    formula_cache: RefCell<HashMap<String, f32>>,
    preserve_formula_cache_across_paths: bool,
    text_rectangle: Option<Rect>,
    commands: Vec<PathCommand>,
    path_layers: Vec<CustomPathLayer>,
    path_start: usize,
    path_fill: PathFillMode,
    path_stroke: bool,
    pending: Option<PendingCustomPathCommand>,
    current_point: Option<(f32, f32)>,
    subpath_start: Option<(f32, f32)>,
    invalid: bool,
}

impl CustomGeometryState {
    pub(super) fn new(depth: usize, width: f32, height: f32) -> Self {
        Self {
            depth,
            default_width: width,
            default_height: height,
            width,
            height,
            formulas: HashMap::new(),
            guides: Vec::new(),
            formula_cache: RefCell::new(HashMap::new()),
            preserve_formula_cache_across_paths: false,
            text_rectangle: None,
            commands: Vec::new(),
            path_layers: Vec::new(),
            path_start: 0,
            path_fill: PathFillMode::Normal,
            path_stroke: true,
            pending: None,
            current_point: None,
            subpath_start: None,
            invalid: false,
        }
    }

    fn builtin(&self, name: &str) -> Option<f32> {
        let width = self.width;
        let height = self.height;
        Some(match name {
            "l" | "t" => 0.0,
            "w" | "r" => width,
            "h" | "b" => height,
            "hc" => width / 2.0,
            "vc" => height / 2.0,
            "ss" => width.min(height),
            "ls" => width.max(height),
            "wd2" => width / 2.0,
            "wd3" => width / 3.0,
            "wd4" => width / 4.0,
            "wd5" => width / 5.0,
            "wd6" => width / 6.0,
            "wd8" => width / 8.0,
            "wd10" => width / 10.0,
            "wd12" => width / 12.0,
            "wd32" => width / 32.0,
            "hd2" => height / 2.0,
            "hd3" => height / 3.0,
            "hd4" => height / 4.0,
            "hd5" => height / 5.0,
            "hd6" => height / 6.0,
            "hd8" => height / 8.0,
            "hd10" => height / 10.0,
            "ssd2" => width.min(height) / 2.0,
            "ssd6" => width.min(height) / 6.0,
            "ssd8" => width.min(height) / 8.0,
            "ssd16" => width.min(height) / 16.0,
            "ssd32" => width.min(height) / 32.0,
            "cd2" => 10_800_000.0,
            "cd4" => 5_400_000.0,
            "cd8" => 2_700_000.0,
            "3cd4" => 16_200_000.0,
            "3cd8" => 8_100_000.0,
            "5cd8" => 13_500_000.0,
            "7cd8" => 18_900_000.0,
            _ => return None,
        })
    }

    fn set_guide(&mut self, name: String, formula: String) {
        self.formula_cache.borrow_mut().remove(&name);
        self.formulas.insert(name.clone(), formula);
        let _ = self.value(&name);
    }

    fn value(&self, token: &str) -> Option<f32> {
        self.value_inner(token, &mut HashSet::new())
    }

    fn value_inner(&self, token: &str, visiting: &mut HashSet<String>) -> Option<f32> {
        if let Ok(value) = token.parse::<f32>() {
            return value.is_finite().then_some(value);
        }
        if let Some(value) = self.builtin(token) {
            return Some(value);
        }
        if let Some(value) = self.formula_cache.borrow().get(token).copied() {
            return Some(value);
        }
        if !visiting.insert(token.to_owned()) {
            return None;
        }
        let formula = self.formulas.get(token)?;
        let mut fields = formula.split_ascii_whitespace();
        let operator = fields.next()?;
        let arguments = fields
            .map(|field| self.value_inner(field, visiting))
            .collect::<Option<Vec<_>>>()?;
        visiting.remove(token);
        let argument = |index: usize| arguments.get(index).copied();
        let value = match operator {
            "val" => argument(0)?,
            "+-" => argument(0)? + argument(1)? - argument(2)?,
            "*/" => argument(0)? * argument(1)? / argument(2)?,
            "+/" => (argument(0)? + argument(1)?) / argument(2)?,
            "?:" => {
                if argument(0)? > 0.0 {
                    argument(1)?
                } else {
                    argument(2)?
                }
            }
            "abs" => argument(0)?.abs(),
            // DrawingML defines `at2 x y` as the angle whose tangent is y / x.
            // Rust's atan2 receiver is the y coordinate, so the operands are
            // intentionally reversed here.
            "at2" => argument(1)?.atan2(argument(0)?) * 180.0 / std::f32::consts::PI * 60_000.0,
            "cat2" => {
                let angle = argument(2)?.atan2(argument(1)?);
                argument(0)? * angle.cos()
            }
            "cos" => argument(0)? * (argument(1)? / 60_000.0 * std::f32::consts::PI / 180.0).cos(),
            "max" => argument(0)?.max(argument(1)?),
            "min" => argument(0)?.min(argument(1)?),
            "mod" => (argument(0)?.powi(2) + argument(1)?.powi(2) + argument(2)?.powi(2)).sqrt(),
            "pin" => argument(1)?.clamp(argument(0)?, argument(2)?),
            "sat2" => {
                let angle = argument(2)?.atan2(argument(1)?);
                argument(0)? * angle.sin()
            }
            "sin" => argument(0)? * (argument(1)? / 60_000.0 * std::f32::consts::PI / 180.0).sin(),
            "sqrt" => argument(0)?.max(0.0).sqrt(),
            "tan" => argument(0)? * (argument(1)? / 60_000.0 * std::f32::consts::PI / 180.0).tan(),
            _ => return None,
        };
        let value = value.is_finite().then_some(value)?;
        self.formula_cache
            .borrow_mut()
            .insert(token.to_owned(), value);
        Some(value)
    }

    fn append_arc(
        &mut self,
        attributes: &[XmlAttribute<'_>],
        part: &str,
    ) -> Result<(), Diagnostic> {
        let read = |name| -> Result<Option<f32>, Diagnostic> {
            Ok(string_attribute(attributes, name, part)?
                .as_deref()
                .and_then(|value| self.value(value)))
        };
        let (
            Some((start_x, start_y)),
            Some(radius_x),
            Some(radius_y),
            Some(start_angle),
            Some(sweep),
        ) = (
            self.current_point,
            read("wR")?,
            read("hR")?,
            read("stAng")?,
            read("swAng")?,
        )
        else {
            self.invalid = true;
            return Ok(());
        };
        if radius_x < 0.0 || radius_y < 0.0 {
            self.invalid = true;
            return Ok(());
        }
        let radial_start = (start_angle / 60_000.0).to_radians();
        let radial_sweep = (sweep / 60_000.0).to_radians();
        let parameter_angle = |angle: f32| (radius_x * angle.sin()).atan2(radius_y * angle.cos());
        // DrawingML angles are radial, unlike the parameter angle in x=rx*cos(t).
        // Keep the degenerate-radius behavior below independent of that conversion.
        let (start, total) = if radius_x == radius_y || radius_x == 0.0 || radius_y == 0.0 {
            (radial_start, radial_sweep)
        } else {
            let start = parameter_angle(radial_start);
            let turns = (radial_sweep / std::f32::consts::TAU).trunc();
            let remainder = radial_sweep % std::f32::consts::TAU;
            let difference = parameter_angle(radial_start + remainder) - start;
            let partial = if remainder > 0.0 {
                difference.rem_euclid(std::f32::consts::TAU)
            } else if remainder < 0.0 {
                -(-difference).rem_euclid(std::f32::consts::TAU)
            } else {
                0.0
            };
            (start, turns * std::f32::consts::TAU + partial)
        };
        let center_x = start_x - radius_x * start.cos();
        let center_y = start_y - radius_y * start.sin();
        if sweep == 0.0 {
            return Ok(());
        }
        if radius_x == 0.0 || radius_y == 0.0 {
            let end_angle = start + total;
            let end_x = center_x + radius_x * end_angle.cos();
            let end_y = center_y + radius_y * end_angle.sin();
            if (end_x - start_x).abs() > f32::EPSILON || (end_y - start_y).abs() > f32::EPSILON {
                self.commands
                    .push(PathCommand::LineTo { x: end_x, y: end_y });
            }
            self.current_point = Some((end_x, end_y));
            return Ok(());
        }
        // Float conversion may put an exact quarter turn just above its boundary.
        let segment_count = (total.abs() / std::f32::consts::FRAC_PI_2 - 1e-6)
            .ceil()
            .max(1.0) as usize;
        let delta = total / segment_count as f32;
        let mut angle = start;
        for _ in 0..segment_count {
            let next = angle + delta;
            let [control_1, control_2, (end_x, end_y)] =
                ellipse_arc_bezier_points((center_x, center_y), (radius_x, radius_y), angle, next);
            self.commands.push(PathCommand::BezierCurveTo {
                cp1x: control_1.0,
                cp1y: control_1.1,
                cp2x: control_2.0,
                cp2y: control_2.1,
                x: end_x,
                y: end_y,
            });
            self.current_point = Some((end_x, end_y));
            angle = next;
        }
        Ok(())
    }

    pub(super) fn start(
        &mut self,
        local: &str,
        attributes: &[XmlAttribute<'_>],
        part: &str,
        depth: usize,
        empty: bool,
    ) -> Result<(), Diagnostic> {
        match local {
            "path" => {
                self.finish_current_path();
                self.width = string_attribute(attributes, "w", part)?
                    .and_then(|value| value.parse::<f32>().ok())
                    .unwrap_or(self.default_width);
                self.height = string_attribute(attributes, "h", part)?
                    .and_then(|value| value.parse::<f32>().ok())
                    .unwrap_or(self.default_height);
                self.path_start = self.commands.len();
                self.path_fill = match string_attribute(attributes, "fill", part)?.as_deref() {
                    Some("none") => PathFillMode::None,
                    Some("darken") => PathFillMode::Darken,
                    Some("darkenLess") => PathFillMode::DarkenLess,
                    Some("lighten") => PathFillMode::Lighten,
                    Some("lightenLess") => PathFillMode::LightenLess,
                    _ => PathFillMode::Normal,
                };
                self.path_stroke = !matches!(
                    string_attribute(attributes, "stroke", part)?.as_deref(),
                    Some("false" | "0" | "off")
                );
                self.current_point = None;
                self.subpath_start = None;
                if !self.preserve_formula_cache_across_paths {
                    // Custom paths may specify their own coordinate space. Replay guides
                    // in declaration order so duplicate names cannot rewrite earlier results.
                    self.formula_cache.borrow_mut().clear();
                    self.formulas.clear();
                    for (name, formula) in self.guides.clone() {
                        self.set_guide(name, formula);
                    }
                }
                if self.width <= 0.0 || self.height <= 0.0 {
                    self.invalid = true;
                }
            }
            "gd" => {
                if let (Some(name), Some(formula)) = (
                    string_attribute(attributes, "name", part)?,
                    string_attribute(attributes, "fmla", part)?,
                ) {
                    self.guides.push((name.clone(), formula.clone()));
                    self.set_guide(name, formula);
                }
            }
            "rect" => {
                let read = |name| -> Result<Option<f32>, Diagnostic> {
                    Ok(string_attribute(attributes, name, part)?
                        .as_deref()
                        .and_then(|value| self.value(value)))
                };
                if let (Some(left), Some(top), Some(right), Some(bottom)) =
                    (read("l")?, read("t")?, read("r")?, read("b")?)
                    && self.width > 0.0
                    && self.height > 0.0
                    && right >= left
                    && bottom >= top
                {
                    let scale_x = 21_600.0 / self.width;
                    let scale_y = 21_600.0 / self.height;
                    self.text_rectangle = Some(Rect {
                        x: left * scale_x,
                        y: top * scale_y,
                        width: (right - left) * scale_x,
                        height: (bottom - top) * scale_y,
                    });
                }
            }
            "moveTo" | "lnTo" | "quadBezTo" | "cubicBezTo" => {
                if self.pending.is_some() || empty {
                    self.invalid = true;
                } else {
                    let kind = match local {
                        "moveTo" => CustomPathCommandKind::Move,
                        "lnTo" => CustomPathCommandKind::Line,
                        "quadBezTo" => CustomPathCommandKind::Quadratic,
                        _ => CustomPathCommandKind::Bezier,
                    };
                    self.pending = Some(PendingCustomPathCommand {
                        depth,
                        kind,
                        points: Vec::new(),
                    });
                }
            }
            "pt" => {
                let x = string_attribute(attributes, "x", part)?
                    .as_deref()
                    .and_then(|value| self.value(value));
                let y = string_attribute(attributes, "y", part)?
                    .as_deref()
                    .and_then(|value| self.value(value));
                let Some(pending) = self.pending.as_mut() else {
                    return Ok(());
                };
                if let (Some(x), Some(y)) = (x, y)
                    && x.is_finite()
                    && y.is_finite()
                {
                    pending.points.push((x, y));
                } else {
                    self.invalid = true;
                }
            }
            "close" => {
                self.commands.push(PathCommand::ClosePath);
                self.current_point = self.subpath_start;
            }
            "arcTo" => self.append_arc(attributes, part)?,
            _ => {}
        }
        Ok(())
    }

    pub(super) fn end(&mut self, local: &str, depth: usize) {
        if !matches!(local, "moveTo" | "lnTo" | "quadBezTo" | "cubicBezTo")
            || self
                .pending
                .as_ref()
                .is_none_or(|pending| pending.depth != depth)
        {
            return;
        }
        let Some(pending) = self.pending.take() else {
            self.invalid = true;
            return;
        };
        let command = match (pending.kind, pending.points.as_slice()) {
            (CustomPathCommandKind::Move, &[(x, y)]) => Some(PathCommand::MoveTo { x, y }),
            (CustomPathCommandKind::Line, &[(x, y)]) => Some(PathCommand::LineTo { x, y }),
            (CustomPathCommandKind::Quadratic, &[(cpx, cpy), (x, y)]) => {
                Some(PathCommand::QuadraticCurveTo { cpx, cpy, x, y })
            }
            (CustomPathCommandKind::Bezier, &[(cp1x, cp1y), (cp2x, cp2y), (x, y)]) => {
                Some(PathCommand::BezierCurveTo {
                    cp1x,
                    cp1y,
                    cp2x,
                    cp2y,
                    x,
                    y,
                })
            }
            _ => None,
        };
        if let Some(command) = command {
            match &command {
                PathCommand::MoveTo { x, y } => {
                    self.current_point = Some((*x, *y));
                    self.subpath_start = Some((*x, *y));
                }
                PathCommand::LineTo { x, y }
                | PathCommand::QuadraticCurveTo { x, y, .. }
                | PathCommand::BezierCurveTo { x, y, .. } => {
                    self.current_point = Some((*x, *y));
                }
                PathCommand::ClosePath => {}
            }
            self.commands.push(command);
        } else {
            self.invalid = true;
        }
    }

    fn normalize_current_path(&mut self) {
        if self.path_start >= self.commands.len() || self.width <= 0.0 || self.height <= 0.0 {
            return;
        }
        let scale_x = 21_600.0 / self.width;
        let scale_y = 21_600.0 / self.height;
        for command in &mut self.commands[self.path_start..] {
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
    }

    fn finish_current_path(&mut self) {
        self.normalize_current_path();
        if self.path_start < self.commands.len() {
            self.path_layers.push(CustomPathLayer {
                start: self.path_start,
                end: self.commands.len(),
                fill: self.path_fill,
                stroke: self.path_stroke,
            });
            self.path_start = self.commands.len();
        }
    }

    fn text_rectangle(&self, bounds: Rect) -> Option<Rect> {
        self.text_rectangle.map(|rectangle| Rect {
            x: rectangle.x * bounds.width / 21_600.0,
            y: rectangle.y * bounds.height / 21_600.0,
            width: rectangle.width * bounds.width / 21_600.0,
            height: rectangle.height * bounds.height / 21_600.0,
        })
    }

    pub(super) fn into_geometry(mut self, bounds: Rect) -> Option<Geometry> {
        self.finish_current_path();
        if self.invalid
            || self.pending.is_some()
            || self.width <= 0.0
            || self.height <= 0.0
            || self.commands.is_empty()
        {
            return None;
        }
        // Every completed path is normalized independently. A following path
        // without explicit `w`/`h` resets to the shape coordinate space, so
        // neither its authored dimensions nor cached guide values may be
        // replaced with the canonical output dimensions between paths.
        let scale_x = bounds.width / 21_600.0;
        let scale_y = bounds.height / 21_600.0;
        let commands: Vec<PathCommand> = self
            .commands
            .into_iter()
            .map(|command| match command {
                PathCommand::MoveTo { x, y } => PathCommand::MoveTo {
                    x: x * scale_x,
                    y: y * scale_y,
                },
                PathCommand::LineTo { x, y } => PathCommand::LineTo {
                    x: x * scale_x,
                    y: y * scale_y,
                },
                PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => PathCommand::QuadraticCurveTo {
                    cpx: cpx * scale_x,
                    cpy: cpy * scale_y,
                    x: x * scale_x,
                    y: y * scale_y,
                },
                PathCommand::BezierCurveTo {
                    cp1x,
                    cp1y,
                    cp2x,
                    cp2y,
                    x,
                    y,
                } => PathCommand::BezierCurveTo {
                    cp1x: cp1x * scale_x,
                    cp1y: cp1y * scale_y,
                    cp2x: cp2x * scale_x,
                    cp2y: cp2y * scale_y,
                    x: x * scale_x,
                    y: y * scale_y,
                },
                PathCommand::ClosePath => PathCommand::ClosePath,
            })
            .collect();
        if self.path_layers.len() == 1
            && self.path_layers[0].fill == PathFillMode::Normal
            && self.path_layers[0].stroke
        {
            return Some(Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands,
            });
        }
        let layers = self
            .path_layers
            .into_iter()
            .map(|layer| PathLayer {
                fill_rule: FillRule::NonZero,
                fill: layer.fill,
                stroke: layer.stroke,
                commands: commands[layer.start..layer.end].to_vec(),
            })
            .collect();
        Some(Geometry::LayeredPath { layers })
    }
}

pub(super) fn drawingml_preset_geometry(
    preset: &str,
    bounds: Rect,
    adjustments: &HashMap<String, f32>,
) -> Option<(Geometry, Option<Rect>)> {
    if matches!(
        preset,
        "actionButtonForwardNext" | "actionButtonBackPrevious"
    ) {
        return Some((
            super::horizontal_action_button_geometry(bounds, preset == "actionButtonForwardNext"),
            Some(Rect {
                x: 0.0,
                y: 0.0,
                width: bounds.width,
                height: bounds.height,
            }),
        ));
    }
    if preset == "lightningBolt" {
        return Some((
            super::lightning_bolt_geometry(bounds),
            Some(Rect {
                x: bounds.width * 8_757.0 / 21_600.0,
                y: bounds.height * 7_437.0 / 21_600.0,
                width: bounds.width * (13_917.0 - 8_757.0) / 21_600.0,
                height: bounds.height * (14_277.0 - 7_437.0) / 21_600.0,
            }),
        ));
    }
    let index = PRESET_SHAPE_DEFINITIONS
        .binary_search_by_key(&preset, |(name, _)| *name)
        .ok()?;
    let definition = PRESET_SHAPE_DEFINITIONS[index].1;
    let mut geometry = CustomGeometryState::new(0, bounds.width, bounds.height);
    geometry.preserve_formula_cache_across_paths = true;
    let mut depth = 0_usize;
    parse_xml(
        definition.as_bytes(),
        crate::limits::Limits::default(),
        |event| {
            match event {
                XmlEvent::StartElement {
                    name,
                    attributes,
                    empty,
                } => {
                    let local = local_name(name);
                    geometry.start(
                        local,
                        &attributes,
                        "presetShapeDefinitions.xml",
                        depth,
                        empty,
                    )?;
                    if local == "gd"
                        && let Some(name) =
                            string_attribute(&attributes, "name", "presetShapeDefinitions.xml")?
                    {
                        if let Some(value) = adjustments.get(&name) {
                            geometry
                                .formulas
                                .insert(name.clone(), format!("val {value}"));
                            geometry.formula_cache.borrow_mut().insert(name, *value);
                        }
                    }
                    if !empty {
                        depth = depth.saturating_add(1);
                    }
                }
                XmlEvent::EndElement { name } => {
                    depth = depth.saturating_sub(1);
                    geometry.end(local_name(name), depth);
                }
                XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
            }
            Ok(())
        },
    )
    .ok()?;
    let text_rectangle = geometry.text_rectangle(bounds);
    geometry
        .into_geometry(bounds)
        .map(|geometry| (geometry, text_rectangle))
}

#[derive(Clone, Copy, Debug)]
enum RunPropertiesTarget {
    Default,
    Current,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
enum TextCapitalization {
    #[default]
    None,
    All,
    Small,
}

fn drawingml_underline(value: &str) -> (bool, bool, bool, bool, bool, bool) {
    let underline = !matches!(value, "none" | "0" | "false");
    (
        underline,
        underline && value.starts_with("wavy"),
        underline && value.starts_with("dotted"),
        underline && value.ends_with("Heavy"),
        underline && matches!(value, "dbl" | "wavyDbl"),
        underline && value == "dotDash",
    )
}

#[derive(Debug)]
struct TextRunState {
    stroke: Option<Box<Paint>>,
    stroke_width: f32,
    glow: Option<Glow>,
    paint: Option<Box<Paint>>,
    depth: usize,
    text: String,
    replacement_text: Option<String>,
    font_family: String,
    font_east_asian: Option<String>,
    font_complex_script: Option<String>,
    font_symbol: Option<String>,
    font_size: f32,
    color: u32,
    bold: bool,
    italic: bool,
    underline: bool,
    wavy_underline: bool,
    dotted_underline: bool,
    heavy_underline: bool,
    double_underline: bool,
    dot_dash_underline: bool,
    double_strikethrough: bool,
    strikethrough: bool,
    highlight: u32,
    baseline_shift: f32,
    letter_spacing: f32,
    east_asian_line_breaks: bool,
    capitalization: TextCapitalization,
    shadow: Option<Shadow>,
    inner_shadow: Option<Shadow>,
    reflection: Option<Reflection>,
    shadow_scale_x: f32,
    shadow_scale_y: f32,
    shadow_skew_x: f32,
    shadow_skew_y: f32,
    shadow_alignment: u8,
}

impl TextRunState {
    fn effect(&self) -> TextEffect {
        TextEffect {
            stroke: self.stroke.clone(),
            stroke_width: self.stroke_width,
            glow: self.glow,
            fill_to_text: self.paint.is_some(),
            shadow: self.shadow,
            inner_shadow: self.inner_shadow,
            reflection: self.reflection,
            wavy_underline: self.wavy_underline,
            dotted_underline: self.dotted_underline,
            heavy_underline: self.heavy_underline,
            double_underline: self.double_underline,
            dot_dash_underline: self.dot_dash_underline,
            double_strikethrough: self.double_strikethrough,
            shadow_scale_x: self.shadow_scale_x,
            shadow_scale_y: self.shadow_scale_y,
            shadow_skew_x: self.shadow_skew_x,
            shadow_skew_y: self.shadow_skew_y,
            shadow_alignment: self.shadow_alignment,
        }
    }

    fn from_shape(shape: &ShapeState, depth: usize) -> Self {
        Self {
            stroke: shape.inherited_run_effect.stroke.clone(),
            stroke_width: shape.inherited_run_effect.stroke_width,
            glow: shape.inherited_run_effect.glow,
            paint: shape.font_paint.clone(),
            depth,
            text: String::new(),
            replacement_text: None,
            font_family: shape.font_family.clone(),
            font_east_asian: shape.font_east_asian.clone(),
            font_complex_script: shape.font_complex_script.clone(),
            font_symbol: None,
            font_size: shape.font_size,
            color: shape.font_color,
            bold: shape.bold,
            italic: shape.italic,
            underline: shape.underline,
            wavy_underline: shape.inherited_run_effect.wavy_underline,
            dotted_underline: false,
            heavy_underline: false,
            double_underline: false,
            dot_dash_underline: false,
            double_strikethrough: shape.inherited_run_effect.double_strikethrough,
            strikethrough: shape.strikethrough,
            highlight: shape.highlight,
            baseline_shift: shape.baseline_shift,
            letter_spacing: shape.letter_spacing,
            east_asian_line_breaks: shape.east_asian_line_breaks,
            capitalization: shape.capitalization,
            shadow: shape.inherited_run_effect.shadow,
            inner_shadow: shape.inherited_run_effect.inner_shadow,
            reflection: shape.inherited_run_effect.reflection,
            shadow_scale_x: shape.inherited_run_effect.shadow_scale_x,
            shadow_scale_y: shape.inherited_run_effect.shadow_scale_y,
            shadow_skew_x: shape.inherited_run_effect.shadow_skew_x,
            shadow_skew_y: shape.inherited_run_effect.shadow_skew_y,
            shadow_alignment: shape.inherited_run_effect.shadow_alignment,
        }
    }

    fn from_table_cell(cell: &TableCellState, depth: usize) -> Self {
        Self {
            stroke: None,
            stroke_width: 0.0,
            glow: None,
            paint: None,
            depth,
            text: String::new(),
            replacement_text: None,
            font_family: cell.font_family.clone(),
            font_east_asian: cell.font_east_asian.clone(),
            font_complex_script: cell.font_complex_script.clone(),
            font_symbol: None,
            font_size: cell.font_size,
            color: cell.font_color,
            bold: cell.bold,
            italic: cell.italic,
            underline: cell.underline,
            wavy_underline: false,
            dotted_underline: false,
            heavy_underline: false,
            double_underline: false,
            dot_dash_underline: false,
            double_strikethrough: false,
            strikethrough: cell.strikethrough,
            highlight: 0,
            baseline_shift: cell.baseline_shift,
            letter_spacing: cell.letter_spacing,
            east_asian_line_breaks: cell.east_asian_line_breaks,
            capitalization: cell.capitalization,
            shadow: None,
            inner_shadow: None,
            reflection: None,
            shadow_scale_x: 1.0,
            shadow_scale_y: 1.0,
            shadow_skew_x: 0.0,
            shadow_skew_y: 0.0,
            shadow_alignment: 7,
        }
    }

    fn set_font(&mut self, script: &str, family: String) {
        match script {
            "ea" => self.font_east_asian = Some(family),
            "cs" => self.font_complex_script = Some(family),
            "sym" => self.font_symbol = Some(family),
            _ => self.font_family = family,
        }
    }

    fn family_for(&self, character: char) -> &str {
        if is_private_use_character(character)
            && let Some(symbol) = self.font_symbol.as_deref()
        {
            return symbol;
        }
        match drawingml_font_class(character) {
            DrawingMlFontClass::Latin => &self.font_family,
            DrawingMlFontClass::EastAsian => {
                self.font_east_asian.as_deref().unwrap_or(&self.font_family)
            }
            DrawingMlFontClass::ComplexScript => self
                .font_complex_script
                .as_deref()
                .unwrap_or(&self.font_family),
        }
    }

    fn text_run_with_text(&self, text: String, font_family: String) -> TextRun {
        TextRun {
            paint: self.paint.clone(),
            east_asian_line_breaks: self.east_asian_line_breaks,
            text,
            font_family,
            font_size: self.font_size,
            color: self.color,
            bold: self.bold,
            italic: self.italic,
            underline: self.underline,
            strikethrough: self.strikethrough,
            highlight: self.highlight,
            baseline_shift: self.baseline_shift,
            letter_spacing: self.letter_spacing,
            horizontal_scale: 1.0,
        }
    }

    fn text_run(&self, text: String, font_family: String) -> TextRun {
        self.text_run_with_text(text, font_family)
    }

    fn finish(&self) -> Vec<TextRun> {
        let mut runs = Vec::new();
        let mut segment_start = 0;
        let mut segment_family: Option<&str> = None;
        for (byte_index, character) in self.text.char_indices() {
            let family = self.family_for(character);
            if segment_family.is_some_and(|current| current != family) {
                runs.push(self.text_run(
                    self.text[segment_start..byte_index].to_owned(),
                    segment_family.unwrap_or(&self.font_family).to_owned(),
                ));
                segment_start = byte_index;
            }
            segment_family = Some(family);
        }
        if segment_start < self.text.len() {
            runs.push(self.text_run(
                self.text[segment_start..].to_owned(),
                segment_family.unwrap_or(&self.font_family).to_owned(),
            ));
        }
        if runs.is_empty() {
            runs.push(self.text_run(String::new(), self.font_family.clone()));
        }
        runs
    }

    fn finish_single(&self) -> TextRun {
        self.finish()
            .into_iter()
            .next()
            .unwrap_or_else(|| self.text_run(String::new(), self.font_family.clone()))
    }

    /// Keeps the source character for legacy symbol-font glyph selection.
    /// The containing object's text separately retains the normalized Unicode
    /// equivalent used by search and accessibility.
    fn finish_single_verbatim(&self) -> TextRun {
        self.text_run_with_text(self.text.clone(), self.font_family.clone())
    }
}

fn is_private_use_character(character: char) -> bool {
    matches!(character as u32, 0xe000..=0xf8ff | 0xf0000..=0xffffd | 0x100000..=0x10fffd)
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum DrawingMlFontClass {
    Latin,
    EastAsian,
    ComplexScript,
}

fn drawingml_font_class(character: char) -> DrawingMlFontClass {
    let code = character as u32;
    if matches!(
        code,
        0x0590..=0x109f
            | 0x1780..=0x17ff
            | 0x1a20..=0x1cff
            | 0xa800..=0xa8ff
            | 0xa980..=0xaaff
            | 0xabc0..=0xabff
            | 0x10e60..=0x10e7f
            | 0x1e800..=0x1efff
    ) {
        DrawingMlFontClass::ComplexScript
    } else if matches!(
        code,
        0x1100..=0x11ff
            // DrawingML routes the Geometric Shapes, Miscellaneous Symbols,
            // and Dingbats blocks through an authored East Asian run font.
            // `family_for` still falls back to Latin when no `a:ea` font is
            // explicit or inherited.
            | 0x25a0..=0x27bf
            | 0x2e80..=0x4dff
            | 0x4e00..=0x9fff
            | 0xa960..=0xa97f
            | 0xac00..=0xd7ff
            | 0xe000..=0xfaff
            | 0xfe10..=0xfe1f
            | 0xfe30..=0xfe6f
            | 0xff00..=0xffef
            | 0x1aff0..=0x1b2ff
            | 0x1f200..=0x1f2ff
            | 0x20000..=0x323af
    ) {
        DrawingMlFontClass::EastAsian
    } else {
        DrawingMlFontClass::Latin
    }
}

#[derive(Clone, Copy, Debug)]
enum PaintTarget {
    Fill,
    Stroke,
    Text,
    Highlight,
    Extrusion,
    Contour,
}

#[derive(Debug)]
struct PictureState {
    image_mapping: crate::model::ImageFillMapping,
    image_tile: bool,
    depth: usize,
    parent_numeric_id: Option<u32>,
    shape_properties_depth: Option<usize>,
    transform_depth: Option<usize>,
    has_transform: bool,
    explicit_bounds: Option<Rect>,
    is_background: bool,
    crop: ImageCrop,
    shape_id: Option<u32>,
    metadata: PptxObjectMetadata,
    non_visual_depth: Option<usize>,
    placeholder_type: Option<String>,
    placeholder_index: Option<u32>,
    x: Option<i64>,
    y: Option<i64>,
    width: Option<u64>,
    height: Option<u64>,
    rotation_degrees: f32,
    flip_horizontal: bool,
    flip_vertical: bool,
    preferred_svg_relationship_id: Option<String>,
    embedded_relationship_id: Option<String>,
    linked_relationship_id: Option<String>,
    media_relationship_id: Option<String>,
    media_kind_hint: Option<MediaKind>,
    blip_depth: Option<usize>,
    opacity: f32,
    color_change_depth: Option<usize>,
    color_change_from_depth: Option<usize>,
    color_change_to_depth: Option<usize>,
    color_change_from: Option<u32>,
    color_change_to: Option<u32>,
    color_change_use_alpha: bool,
    geometry: super::drawingml::DrawingMlPictureGeometry,
    stroke: Paint,
    stroke_width: f32,
    line_depth: Option<usize>,
    dash_pattern: DrawingMlDashPattern,
    line_cap: LineCap,
    line_join: LineJoin,
    line_compound: LineCompound,
    line_alignment: LineAlignment,
    miter_limit: f32,
    fill: Paint,
    image_adjustment: ImageAdjustment,
    background_image_fill_index: Option<usize>,
    background_fill_reference: Option<u64>,
    background_fill_defined: bool,
    fill_capture: Option<PaintCapture>,
    fill_reference_depth: Option<usize>,
    line_reference_depth: Option<usize>,
    style_shadow: bool,
    effects: DrawingMlPictureEffectsCapture,
}

impl PictureState {
    fn new(depth: usize, parent_numeric_id: Option<u32>) -> Self {
        Self {
            image_mapping: crate::model::ImageFillMapping::default(),
            image_tile: false,
            depth,
            parent_numeric_id,
            shape_properties_depth: None,
            transform_depth: None,
            has_transform: false,
            explicit_bounds: None,
            is_background: false,
            crop: ImageCrop::default(),
            shape_id: None,
            metadata: PptxObjectMetadata::default(),
            non_visual_depth: None,
            placeholder_type: None,
            placeholder_index: None,
            x: None,
            y: None,
            width: None,
            height: None,
            rotation_degrees: 0.0,
            flip_horizontal: false,
            flip_vertical: false,
            preferred_svg_relationship_id: None,
            embedded_relationship_id: None,
            linked_relationship_id: None,
            media_relationship_id: None,
            media_kind_hint: None,
            blip_depth: None,
            opacity: 1.0,
            color_change_depth: None,
            color_change_from_depth: None,
            color_change_to_depth: None,
            color_change_from: None,
            color_change_to: None,
            color_change_use_alpha: true,
            geometry: super::drawingml::DrawingMlPictureGeometry::default(),
            stroke: Paint::None,
            stroke_width: 0.0,
            line_depth: None,
            dash_pattern: DrawingMlDashPattern::Solid,
            line_cap: LineCap::Flat,
            line_join: LineJoin::Miter,
            line_compound: LineCompound::Single,
            line_alignment: LineAlignment::Center,
            miter_limit: 10.0,
            fill: Paint::None,
            image_adjustment: ImageAdjustment::default(),
            background_image_fill_index: None,
            background_fill_reference: None,
            background_fill_defined: false,
            fill_capture: None,
            fill_reference_depth: None,
            line_reference_depth: None,
            style_shadow: false,
            effects: DrawingMlPictureEffectsCapture::default(),
        }
    }

    fn background(depth: usize, bounds: Rect) -> Self {
        Self {
            image_mapping: crate::model::ImageFillMapping::default(),
            image_tile: false,
            depth,
            parent_numeric_id: None,
            shape_properties_depth: None,
            transform_depth: None,
            has_transform: false,
            explicit_bounds: Some(bounds),
            is_background: true,
            crop: ImageCrop::default(),
            shape_id: None,
            metadata: PptxObjectMetadata::default(),
            non_visual_depth: None,
            placeholder_type: None,
            placeholder_index: None,
            x: None,
            y: None,
            width: None,
            height: None,
            rotation_degrees: 0.0,
            flip_horizontal: false,
            flip_vertical: false,
            preferred_svg_relationship_id: None,
            embedded_relationship_id: None,
            linked_relationship_id: None,
            media_relationship_id: None,
            media_kind_hint: None,
            blip_depth: None,
            opacity: 1.0,
            color_change_depth: None,
            color_change_from_depth: None,
            color_change_to_depth: None,
            color_change_from: None,
            color_change_to: None,
            color_change_use_alpha: true,
            geometry: super::drawingml::DrawingMlPictureGeometry::default(),
            stroke: Paint::None,
            stroke_width: 0.0,
            line_depth: None,
            dash_pattern: DrawingMlDashPattern::Solid,
            line_cap: LineCap::Flat,
            line_join: LineJoin::Miter,
            line_compound: LineCompound::Single,
            line_alignment: LineAlignment::Center,
            miter_limit: 10.0,
            fill: Paint::None,
            image_adjustment: ImageAdjustment::default(),
            background_image_fill_index: None,
            background_fill_reference: None,
            background_fill_defined: false,
            fill_capture: None,
            fill_reference_depth: None,
            line_reference_depth: None,
            style_shadow: false,
            effects: DrawingMlPictureEffectsCapture::default(),
        }
    }
}

#[derive(Debug)]
struct GroupState {
    depth: usize,
    numeric_id: u32,
    shape_id: Option<u32>,
    metadata: PptxObjectMetadata,
    non_visual_depth: Option<usize>,
    shape_properties_depth: Option<usize>,
    fill: Paint,
    explicit_fill: bool,
    paint_capture: Option<PaintCapture>,
    transform_depth: Option<usize>,
    x: Option<i64>,
    y: Option<i64>,
    width: Option<u64>,
    height: Option<u64>,
    child_x: Option<i64>,
    child_y: Option<i64>,
    child_width: Option<u64>,
    child_height: Option<u64>,
    rotation_degrees: f32,
    flip_horizontal: bool,
    flip_vertical: bool,
    three_d: Option<ThreeDStyle>,
    scene_3d_depth: Option<usize>,
    backdrop_3d_depth: Option<usize>,
    camera_3d_depth: Option<usize>,
    light_3d_depth: Option<usize>,
}

#[derive(Debug)]
struct GraphicFrameState {
    depth: usize,
    parent_numeric_id: Option<u32>,
    shape_id: Option<u32>,
    metadata: PptxObjectMetadata,
    non_visual_depth: Option<usize>,
    transform_depth: Option<usize>,
    x: Option<i64>,
    y: Option<i64>,
    width: Option<u64>,
    height: Option<u64>,
    table_depth: Option<usize>,
    table_style_id: String,
    table_fill: Option<super::drawingml::ChartFill>,
    table_effects: Option<super::drawingml::DrawingMlPictureEffects>,
    collecting_table_style_id: bool,
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
    band_rows: bool,
    band_columns: bool,
    columns: Vec<u64>,
    rows: Vec<TableRowState>,
    row: Option<TableRowState>,
    cell: Option<TableCellState>,
    chart: Option<BasicChart>,
    diagram: Option<PptxDiagram>,
    embedded_raster: Option<(String, Vec<u8>, String)>,
    embedded_text: Option<(Document, f32, String)>,
    embedded_plain_text: Option<(String, f32, String)>,
    embedded_document: Option<(Document, String)>,
    placeholder: Option<&'static str>,
}

#[derive(Debug)]
struct TableRowState {
    depth: usize,
    height: u64,
    cells: Vec<TableCellState>,
}

#[derive(Clone, Debug)]
struct TableCellRunDefaults {
    font_family: String,
    font_east_asian: Option<String>,
    font_complex_script: Option<String>,
    font_size: f32,
    font_color: u32,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    baseline_shift: f32,
    letter_spacing: f32,
    east_asian_line_breaks: bool,
    capitalization: TextCapitalization,
}

impl TableCellRunDefaults {
    fn from_presentation_defaults(theme: &PptxTheme, style: Option<&ParagraphStyle>) -> Self {
        let style = style.cloned().unwrap_or_default();
        Self {
            font_family: style
                .font_family
                .unwrap_or_else(|| theme.resolve_typeface("+mn-lt", "latin")),
            font_east_asian: style.font_east_asian,
            font_complex_script: style.font_complex_script,
            font_size: style.font_size.unwrap_or(18.0 * POINTS_TO_CSS_PIXELS),
            font_color: style.font_color.unwrap_or(0x0000_00ff),
            bold: style.bold.unwrap_or(false),
            italic: style.italic.unwrap_or(false),
            underline: style.underline.unwrap_or(false),
            strikethrough: style.strikethrough.unwrap_or(false),
            baseline_shift: style.baseline_ratio.unwrap_or(0.0)
                * style.font_size.unwrap_or(18.0 * POINTS_TO_CSS_PIXELS),
            letter_spacing: style.letter_spacing.unwrap_or(0.0),
            east_asian_line_breaks: style.east_asian_line_breaks.unwrap_or(true),
            capitalization: style.capitalization.unwrap_or_default(),
        }
    }
}

#[derive(Debug)]
struct TableCellState {
    depth: usize,
    text: String,
    collecting_text: bool,
    fill: Paint,
    explicit_fill: bool,
    style_fill: Option<super::drawingml::ChartFill>,
    effects: Option<super::drawingml::DrawingMlPictureEffects>,
    borders: [TableCellBorder; 4],
    grid_span: usize,
    row_span: usize,
    horizontal_merge: bool,
    vertical_merge: bool,
    properties_depth: Option<usize>,
    border_depth: Option<(usize, TableBorderSide)>,
    ignored_border_depth: Option<usize>,
    paint_capture: Option<(CellPaintTarget, PaintCapture)>,
    paragraph_count: u32,
    paragraph_styles: Vec<ParagraphStyle>,
    local_paragraph_styles: Vec<ParagraphStyle>,
    run_style: ParagraphStyle,
    run_styles: Vec<ParagraphStyle>,
    paragraph_style_capture: Option<ParagraphStyleCapture>,
    paragraph_default_run_depth: Option<usize>,
    current_paragraph_style: ParagraphStyle,
    current_paragraph_level: usize,
    paragraph_pending: bool,
    paragraph_layouts: Vec<TextParagraphLayout>,
    paragraph_layout_styles: Vec<ParagraphStyle>,
    paragraph_run_starts: Vec<usize>,
    numbering_counters: [u32; TEXT_LEVEL_COUNT],
    align: TextAlign,
    text_inset_left: f32,
    text_inset_right: f32,
    text_inset_top: f32,
    text_inset_bottom: f32,
    vertical_align: TextVerticalAlign,
    text_orientation: TextOrientation,
    base_run: TableCellRunDefaults,
    font_family: String,
    font_east_asian: Option<String>,
    font_complex_script: Option<String>,
    font_size: f32,
    font_color: u32,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    baseline_shift: f32,
    letter_spacing: f32,
    east_asian_line_breaks: bool,
    capitalization: TextCapitalization,
    text_direction: TextDirection,
    run: Option<TextRunState>,
    runs: Vec<TextRun>,
    run_effects: Vec<TextEffect>,
    run_properties_depth: Option<usize>,
    explicit_text_color: bool,
}

#[derive(Clone, Copy, Debug)]
enum CellPaintTarget {
    Fill,
    Stroke(TableBorderSide),
    Text,
    ParagraphText,
    Highlight,
}

#[derive(Clone, Copy, Debug, Eq, Ord, PartialEq, PartialOrd)]
enum TableBorderSide {
    Left,
    Right,
    Top,
    Bottom,
}

impl TableBorderSide {
    const ALL: [Self; 4] = [Self::Left, Self::Right, Self::Top, Self::Bottom];

    fn from_element(local: &str) -> Option<Self> {
        match local {
            "lnL" => Some(Self::Left),
            "lnR" => Some(Self::Right),
            "lnT" => Some(Self::Top),
            "lnB" => Some(Self::Bottom),
            _ => None,
        }
    }

    const fn index(self) -> usize {
        match self {
            Self::Left => 0,
            Self::Right => 1,
            Self::Top => 2,
            Self::Bottom => 3,
        }
    }
}

#[derive(Clone, Debug)]
struct TableCellBorder {
    paint: Paint,
    width: f32,
    dash_pattern: DrawingMlDashPattern,
    stroke_style: StrokeStyle,
    explicit: bool,
}

impl Default for TableCellBorder {
    fn default() -> Self {
        Self {
            paint: Paint::Solid(0x8080_80ff),
            width: 1.0,
            dash_pattern: DrawingMlDashPattern::Solid,
            stroke_style: StrokeStyle::default(),
            explicit: false,
        }
    }
}

impl TableCellState {
    fn border(&self, side: TableBorderSide) -> &TableCellBorder {
        &self.borders[side.index()]
    }

    fn border_mut(&mut self, side: TableBorderSide) -> &mut TableCellBorder {
        &mut self.borders[side.index()]
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum TableStyleBorderSide {
    Left,
    Right,
    Top,
    Bottom,
    InsideHorizontal,
    InsideVertical,
}

impl TableStyleBorderSide {
    fn from_element(local: &str) -> Option<Self> {
        match local {
            "left" => Some(Self::Left),
            "right" => Some(Self::Right),
            "top" => Some(Self::Top),
            "bottom" => Some(Self::Bottom),
            "insideH" => Some(Self::InsideHorizontal),
            "insideV" => Some(Self::InsideVertical),
            _ => None,
        }
    }

    const fn index(self) -> usize {
        match self {
            Self::Left => 0,
            Self::Right => 1,
            Self::Top => 2,
            Self::Bottom => 3,
            Self::InsideHorizontal => 4,
            Self::InsideVertical => 5,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
enum TableStyleRegion {
    WholeTable,
    Band1Horizontal,
    Band2Horizontal,
    Band1Vertical,
    Band2Vertical,
    FirstRow,
    LastRow,
    FirstColumn,
    LastColumn,
    NorthWest,
    NorthEast,
    SouthWest,
    SouthEast,
}

impl TableStyleRegion {
    fn from_element(local: &str) -> Option<Self> {
        Some(match local {
            "wholeTbl" => Self::WholeTable,
            "band1H" => Self::Band1Horizontal,
            "band2H" => Self::Band2Horizontal,
            "band1V" => Self::Band1Vertical,
            "band2V" => Self::Band2Vertical,
            "firstRow" => Self::FirstRow,
            "lastRow" => Self::LastRow,
            "firstCol" => Self::FirstColumn,
            "lastCol" => Self::LastColumn,
            "nwCell" => Self::NorthWest,
            "neCell" => Self::NorthEast,
            "swCell" => Self::SouthWest,
            "seCell" => Self::SouthEast,
            _ => return None,
        })
    }
}

#[derive(Clone, Debug, Default)]
struct TablePartStyle {
    borders: [Option<TableCellBorder>; 6],
    fill: Option<super::drawingml::ChartFill>,
    text: ParagraphStyle,
    effects: Option<super::drawingml::DrawingMlPictureEffects>,
}

#[derive(Debug)]
struct PptxTableStyle {
    parts: HashMap<TableStyleRegion, TablePartStyle>,
    background_fill: Option<super::drawingml::ChartFill>,
    background_effects: Option<super::drawingml::DrawingMlPictureEffects>,
    image_fills: Vec<(Option<TableStyleRegion>, String, ShapeImageFillState)>,
    diagnostics: Vec<Diagnostic>,
    empty_border: TableCellBorder,
}

#[derive(Clone, Copy, Debug)]
struct TableCellStyleContext {
    row: usize,
    column: usize,
    row_span: usize,
    column_span: usize,
    row_count: usize,
    column_count: usize,
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
    band_rows: bool,
    band_columns: bool,
}

impl TableCellStyleContext {
    fn is_first_row(self) -> bool {
        self.first_row && self.row == 0
    }

    fn is_last_row(self) -> bool {
        self.last_row && self.row.saturating_add(self.row_span) >= self.row_count
    }

    fn is_first_column(self) -> bool {
        self.first_column && self.column == 0
    }

    fn is_last_column(self) -> bool {
        self.last_column && self.column.saturating_add(self.column_span) >= self.column_count
    }

    fn row_band(self) -> Option<usize> {
        self.band_rows
            .then_some(())
            .filter(|_| !self.is_first_row() && !self.is_last_row())
            .map(|_| self.row.saturating_sub(usize::from(self.first_row)) % 2)
    }

    fn column_band(self) -> Option<usize> {
        self.band_columns
            .then_some(())
            .filter(|_| !self.is_first_column() && !self.is_last_column())
            .map(|_| self.column.saturating_sub(usize::from(self.first_column)) % 2)
    }

    // Highest priority first; omitted properties continue through the cascade.
    fn regions(self) -> impl Iterator<Item = TableStyleRegion> {
        use TableStyleRegion::*;
        [
            (self.is_first_row() && self.is_first_column()).then_some(NorthWest),
            (self.is_first_row() && self.is_last_column()).then_some(NorthEast),
            (self.is_last_row() && self.is_first_column()).then_some(SouthWest),
            (self.is_last_row() && self.is_last_column()).then_some(SouthEast),
            self.is_first_row().then_some(FirstRow),
            self.is_last_row().then_some(LastRow),
            self.is_first_column().then_some(FirstColumn),
            self.is_last_column().then_some(LastColumn),
            self.row_band().map(|band| {
                if band == 0 {
                    Band1Horizontal
                } else {
                    Band2Horizontal
                }
            }),
            self.column_band().map(|band| {
                if band == 0 {
                    Band1Vertical
                } else {
                    Band2Vertical
                }
            }),
            Some(WholeTable),
        ]
        .into_iter()
        .flatten()
    }
}

impl PptxTableStyle {
    fn cell_border(
        &self,
        side: TableBorderSide,
        context: TableCellStyleContext,
    ) -> &TableCellBorder {
        use TableStyleRegion::*;
        for region in context.regions() {
            let Some(part) = self.parts.get(&region) else {
                continue;
            };
            let full_rows = matches!(
                region,
                WholeTable | Band1Vertical | Band2Vertical | FirstColumn | LastColumn
            );
            let full_columns = matches!(
                region,
                WholeTable | Band1Horizontal | Band2Horizontal | FirstRow | LastRow
            );
            let side = match side {
                TableBorderSide::Left if !full_columns || context.column == 0 => {
                    TableStyleBorderSide::Left
                }
                TableBorderSide::Left => TableStyleBorderSide::InsideVertical,
                TableBorderSide::Right
                    if !full_columns
                        || context.column + context.column_span >= context.column_count =>
                {
                    TableStyleBorderSide::Right
                }
                TableBorderSide::Right => TableStyleBorderSide::InsideVertical,
                TableBorderSide::Top if !full_rows || context.row == 0 => TableStyleBorderSide::Top,
                TableBorderSide::Top => TableStyleBorderSide::InsideHorizontal,
                TableBorderSide::Bottom
                    if !full_rows || context.row + context.row_span >= context.row_count =>
                {
                    TableStyleBorderSide::Bottom
                }
                TableBorderSide::Bottom => TableStyleBorderSide::InsideHorizontal,
            };
            if let Some(border) = part.borders[side.index()].as_ref() {
                return border;
            }
        }
        &self.empty_border
    }

    fn cell_fill(&self, context: TableCellStyleContext) -> Option<&super::drawingml::ChartFill> {
        context
            .regions()
            .find_map(|region| self.parts.get(&region)?.fill.as_ref())
    }

    fn cell_text_style(&self, context: TableCellStyleContext) -> ParagraphStyle {
        let mut result = ParagraphStyle::default();
        let regions = context.regions().collect::<Vec<_>>();
        for region in regions.into_iter().rev() {
            if let Some(part) = self.parts.get(&region) {
                result.apply(&part.text);
            }
        }
        result
    }

    fn cell_effects(
        &self,
        context: TableCellStyleContext,
    ) -> Option<&super::drawingml::DrawingMlPictureEffects> {
        context
            .regions()
            .find_map(|region| self.parts.get(&region)?.effects.as_ref())
    }
}
#[derive(Clone, Debug)]
struct PptxTheme {
    part: Option<String>,
    colors: HashMap<String, u32>,
    fonts: HashMap<String, String>,
    color_map: HashMap<String, String>,
    line_styles: DrawingMlThemeLineStyles,
    background_image_fills: Vec<PptxThemeBackgroundImageFill>,
}

#[derive(Clone, Debug, Default)]
struct PptxThemeBackgroundImageFill {
    mapping: crate::model::ImageFillMapping,
    relationship: Option<Relationship>,
    tile: bool,
    duotone: Vec<PptxThemeColor>,
}

#[derive(Clone, Debug)]
struct PptxThemeColor {
    kind: String,
    value: String,
    transforms: Vec<(String, f32)>,
}

impl PptxThemeBackgroundImageFill {
    fn adjustment(&self, placeholder: u32, theme: &PptxTheme) -> ImageAdjustment {
        let colors = self
            .duotone
            .iter()
            .filter_map(|color| {
                let mut value = match color.kind.as_str() {
                    "srgbClr" => parse_rgb_color(&color.value),
                    "schemeClr" if color.value == "phClr" => Some(placeholder),
                    "schemeClr" => theme.color(&color.value),
                    "prstClr" => drawingml_preset_color(&color.value),
                    _ => None,
                }?;
                for (name, ratio) in &color.transforms {
                    apply_drawingml_color_transform(&mut value, name, *ratio);
                }
                Some(value)
            })
            .collect::<Vec<_>>();
        ImageAdjustment {
            duotone: (colors.len() == 2).then(|| [colors[0], colors[1]]),
            ..ImageAdjustment::default()
        }
    }
}

impl Default for PptxTheme {
    fn default() -> Self {
        let colors = [
            "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
            "accent6", "hlink", "folHlink",
        ]
        .into_iter()
        .filter_map(|name| default_scheme_color(name).map(|color| (name.to_owned(), color)))
        .collect();
        // Empty DrawingML east-Asian/complex theme faces inherit the Latin
        // face. Seeding explicit Arial entries for those slots prevents that
        // fallback and silently replaces authored major/minor fonts.
        let fonts = ["+mj-lt", "+mn-lt"]
            .into_iter()
            .map(|name| (name.to_owned(), "Arial".to_owned()))
            .collect();
        Self {
            part: None,
            colors,
            fonts,
            color_map: default_color_map(),
            line_styles: DrawingMlThemeLineStyles::default(),
            background_image_fills: Vec::new(),
        }
    }
}

impl PptxTheme {
    fn color(&self, value: &str) -> Option<u32> {
        if value == "phClr"
            && let Some(color) = self.colors.get(value)
        {
            return Some(*color);
        }
        let value = if value == "phClr" { "accent1" } else { value };
        let slot = self.color_map.get(value).map_or(value, String::as_str);
        self.colors.get(slot).copied()
    }

    fn with_color_map(mut self, color_map: HashMap<String, String>) -> Self {
        self.color_map = color_map;
        self
    }

    fn line_width(&self, index: usize) -> Option<f32> {
        self.line_styles.width(index as u64)
    }

    fn resolve_typeface(&self, typeface: &str, script: &str) -> String {
        if !typeface.starts_with("+mj") && !typeface.starts_with("+mn") {
            return typeface.to_owned();
        }
        let family = if typeface.starts_with("+mj") {
            "mj"
        } else {
            "mn"
        };
        let script = match script {
            "ea" => "ea",
            "cs" => "cs",
            _ => "lt",
        };
        let exact = if typeface.len() > 3 { typeface } else { "" };
        self.fonts
            .get(exact)
            .or_else(|| self.fonts.get(&format!("+{family}-{script}")))
            .or_else(|| self.fonts.get(&format!("+{family}-lt")))
            .cloned()
            .unwrap_or_else(|| "Arial".to_owned())
    }
}

fn default_color_map() -> HashMap<String, String> {
    [
        ("bg1", "lt1"),
        ("tx1", "dk1"),
        ("bg2", "lt2"),
        ("tx2", "dk2"),
        ("accent1", "accent1"),
        ("accent2", "accent2"),
        ("accent3", "accent3"),
        ("accent4", "accent4"),
        ("accent5", "accent5"),
        ("accent6", "accent6"),
        ("hlink", "hlink"),
        ("folHlink", "folHlink"),
    ]
    .into_iter()
    .map(|(key, value)| (key.to_owned(), value.to_owned()))
    .collect()
}

fn parse_pptx_theme(package: &Package<'_>, part: &str) -> Result<PptxTheme, Diagnostic> {
    let bytes = package.required_part(part)?;
    let relationships = package.relationships(Some(part))?;
    let mut theme = PptxTheme::default();
    theme.part = Some(part.to_owned());
    theme
        .colors
        .extend(super::drawingml::drawingml_theme_colors(
            &bytes,
            package.limits(),
            part,
        )?);
    let mut depth = 0_usize;
    let mut font_family: Option<(usize, &'static str)> = None;
    let mut background_fill_list_depth = None;
    let mut background_fill: Option<(usize, usize)> = None;
    let mut background_duotone_depth = None;
    let mut background_duotone_color: Option<(usize, usize)> = None;
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "bgFillStyleLst" {
                    background_fill_list_depth = (!empty).then_some(depth);
                } else if background_fill_list_depth
                    .is_some_and(|start| depth == start.saturating_add(1))
                {
                    let index = theme.background_image_fills.len();
                    theme
                        .background_image_fills
                        .push(PptxThemeBackgroundImageFill::default());
                    super::drawingml::drawingml_image_fill_mapping(
                        &mut theme.background_image_fills[index].mapping,
                        local,
                        &attributes,
                        part,
                    )?;
                    background_fill = (!empty).then_some((depth, index));
                } else if local == "blip"
                    && let Some((_, index)) = background_fill
                    && let Some(relationship_id) = string_attribute(&attributes, "embed", part)?
                    && let Some(relationship) = relationships
                        .iter()
                        .find(|relationship| relationship.id == relationship_id)
                {
                    theme.background_image_fills[index].relationship = Some(relationship.clone());
                } else if matches!(local, "tile" | "fillRect")
                    && let Some((_, index)) = background_fill
                {
                    theme.background_image_fills[index].tile = local == "tile";
                    super::drawingml::drawingml_image_fill_mapping(
                        &mut theme.background_image_fills[index].mapping,
                        local,
                        &attributes,
                        part,
                    )?;
                } else if local == "duotone" && background_fill.is_some() {
                    background_duotone_depth = (!empty).then_some(depth);
                } else if matches!(local, "srgbClr" | "schemeClr" | "prstClr")
                    && background_duotone_depth
                        .is_some_and(|start| depth == start.saturating_add(1))
                    && let Some((_, index)) = background_fill
                    && let Some(value) = string_attribute(&attributes, "val", part)?
                {
                    let color_index = theme.background_image_fills[index].duotone.len();
                    theme.background_image_fills[index]
                        .duotone
                        .push(PptxThemeColor {
                            kind: local.to_owned(),
                            value,
                            transforms: Vec::new(),
                        });
                    background_duotone_color = (!empty).then_some((depth, color_index));
                } else if matches!(
                    local,
                    "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff"
                ) && let (Some((_, fill_index)), Some((_, color_index))) =
                    (background_fill, background_duotone_color)
                    && let Some(value) = numeric_attribute(&attributes, "val", part)?
                {
                    theme.background_image_fills[fill_index].duotone[color_index]
                        .transforms
                        .push((local.to_owned(), value as f32 / 100_000.0));
                }
                if local == "majorFont" {
                    font_family = (!empty).then_some((depth, "mj"));
                } else if local == "minorFont" {
                    font_family = (!empty).then_some((depth, "mn"));
                } else if matches!(local, "latin" | "ea" | "cs")
                    && let Some((_, family)) = font_family
                    && let Some(typeface) = string_attribute(&attributes, "typeface", part)?
                        .filter(|typeface| !typeface.is_empty())
                {
                    let script = if local == "latin" { "lt" } else { local };
                    theme.fonts.insert(format!("+{family}-{script}"), typeface);
                }
                theme
                    .line_styles
                    .start(local, &attributes, depth, empty, part)?;
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if background_fill.is_some_and(|(start, _)| start == depth) {
                    background_fill = None;
                }
                if background_duotone_color.is_some_and(|(start, _)| start == depth) {
                    background_duotone_color = None;
                }
                if local == "duotone" && background_duotone_depth == Some(depth) {
                    background_duotone_depth = None;
                }
                if local == "bgFillStyleLst" && background_fill_list_depth == Some(depth) {
                    background_fill_list_depth = None;
                }
                if font_family
                    .as_ref()
                    .is_some_and(|(start, _)| *start == depth)
                {
                    font_family = None;
                }
                theme.line_styles.end(local, depth);
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(theme)
}

#[derive(Debug)]
enum PaintCaptureKind {
    Solid,
    Gradient,
    Pattern { preset: String },
}

#[derive(Clone, Copy, Debug)]
enum PatternColorTarget {
    Foreground,
    Background,
}

#[derive(Debug)]
struct PaintCapture {
    gradient_mapping: super::drawingml::DrawingMlGradientMapping,
    depth: usize,
    kind: PaintCaptureKind,
    color: Option<u32>,
    background_color: Option<u32>,
    pattern_color_target: PatternColorTarget,
    stops: Vec<GradientStop>,
    current_stop: Option<(f32, Option<u32>)>,
    angle_degrees: f32,
    angle_scaled: bool,
    shape_path: bool,
    circular_path: bool,
    rectangular_path: bool,
    fill_to_rectangle: Option<DrawingMlRelativeRectangle>,
}

fn apply_paint_color_transform(paint: &mut Paint, local: &str, ratio: f32) {
    if let Paint::Solid(color) = paint {
        apply_drawingml_color_transform(color, local, ratio);
    }
}

fn drawingml_color_value(
    local: &str,
    attributes: &[XmlAttribute<'_>],
    part: &str,
    theme: &PptxTheme,
) -> Result<Option<u32>, Diagnostic> {
    match local {
        "srgbClr" => string_attribute(attributes, "val", part)?
            .map(|value| {
                parse_rgb_color(&value).ok_or_else(|| {
                    format_error(
                        part,
                        "DrawingML srgbClr must contain six hexadecimal digits",
                    )
                })
            })
            .transpose(),
        "scrgbClr" => {
            let channel = |name| -> Result<u32, Diagnostic> {
                let linear = signed_numeric_attribute(attributes, name, part)?.unwrap_or(0) as f32
                    / 100_000.0;
                Ok((linear_to_srgb(linear.clamp(0.0, 1.0)) * 255.0).round() as u32)
            };
            Ok(Some(
                (channel("r")? << 24) | (channel("g")? << 16) | (channel("b")? << 8) | 0xff,
            ))
        }
        "schemeClr" => Ok(string_attribute(attributes, "val", part)?
            .as_deref()
            .and_then(|value| theme.color(value))),
        "sysClr" => Ok(string_attribute(attributes, "lastClr", part)?
            .as_deref()
            .and_then(parse_rgb_color)),
        "prstClr" => Ok(string_attribute(attributes, "val", part)?
            .as_deref()
            .and_then(drawingml_preset_color)),
        _ => Ok(None),
    }
}

fn picture_effect_color(
    kind: &str,
    value: &str,
    theme: &PptxTheme,
    part: &str,
) -> Result<u32, Diagnostic> {
    match kind {
        "srgbClr" | "sysClr" => parse_rgb_color(value).ok_or_else(|| {
            format_error(
                part,
                "PPTX picture effect color must contain six hexadecimal digits",
            )
        }),
        "schemeClr" => theme
            .color(value)
            .ok_or_else(|| format_error(part, format!("unknown PPTX theme color `{value}`"))),
        "prstClr" => Ok(drawingml_preset_color(value).unwrap_or(0x0000_00ff)),
        _ => Err(format_error(part, "unsupported PPTX picture effect color")),
    }
}

impl PaintCapture {
    fn new(
        local: &str,
        attributes: &[XmlAttribute<'_>],
        part: &str,
        depth: usize,
    ) -> Result<Option<Self>, Diagnostic> {
        let kind = match local {
            "solidFill" => PaintCaptureKind::Solid,
            "gradFill" => PaintCaptureKind::Gradient,
            "pattFill" => PaintCaptureKind::Pattern {
                preset: string_attribute(attributes, "prst", part)?
                    .unwrap_or_else(|| "pct5".to_owned()),
            },
            _ => return Ok(None),
        };
        let mut gradient_mapping = super::drawingml::DrawingMlGradientMapping::default();
        gradient_mapping.start(local, attributes, part)?;
        Ok(Some(Self {
            gradient_mapping,
            depth,
            kind,
            color: None,
            background_color: None,
            pattern_color_target: PatternColorTarget::Foreground,
            stops: Vec::new(),
            current_stop: None,
            angle_degrees: 0.0,
            angle_scaled: false,
            shape_path: false,
            circular_path: false,
            rectangular_path: false,
            fill_to_rectangle: None,
        }))
    }

    fn set_color(&mut self, color: u32) {
        if let Some((_, stop_color)) = self.current_stop.as_mut() {
            *stop_color = Some(color);
        } else if matches!(self.kind, PaintCaptureKind::Pattern { .. })
            && matches!(self.pattern_color_target, PatternColorTarget::Background)
        {
            self.background_color = Some(color);
        } else {
            self.color = Some(color);
        }
    }

    fn start(
        &mut self,
        local: &str,
        attributes: &[XmlAttribute<'_>],
        part: &str,
        theme: &PptxTheme,
    ) -> Result<(), Diagnostic> {
        self.gradient_mapping.start(local, attributes, part)?;
        match local {
            "fgClr" if matches!(self.kind, PaintCaptureKind::Pattern { .. }) => {
                self.pattern_color_target = PatternColorTarget::Foreground;
            }
            "bgClr" if matches!(self.kind, PaintCaptureKind::Pattern { .. }) => {
                self.pattern_color_target = PatternColorTarget::Background;
            }
            "gs" => {
                let offset =
                    numeric_attribute(attributes, "pos", part)?.unwrap_or(0) as f32 / 100_000.0;
                self.current_stop = Some((offset.clamp(0.0, 1.0), None));
            }
            "srgbClr" | "scrgbClr" | "schemeClr" | "sysClr" | "prstClr" => {
                if let Some(color) = drawingml_color_value(local, attributes, part, theme)? {
                    self.set_color(color);
                }
            }
            "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff" => {
                let ratio = color_transform_percentage(attributes, part)?;
                if let Some((_, Some(color))) = self.current_stop.as_mut() {
                    apply_drawingml_color_transform(color, local, ratio);
                } else if matches!(self.kind, PaintCaptureKind::Pattern { .. })
                    && matches!(self.pattern_color_target, PatternColorTarget::Background)
                    && let Some(color) = self.background_color.as_mut()
                {
                    apply_drawingml_color_transform(color, local, ratio);
                } else if let Some(color) = self.color.as_mut() {
                    apply_drawingml_color_transform(color, local, ratio);
                }
            }
            "lin" => {
                self.angle_degrees = signed_numeric_attribute(attributes, "ang", part)?.unwrap_or(0)
                    as f32
                    / 60_000.0;
                self.angle_scaled = boolean_attribute(attributes, "scaled", part)?.unwrap_or(false);
            }
            "path" => {
                let path = string_attribute(attributes, "path", part)?;
                self.shape_path = path.as_deref() == Some("shape");
                self.circular_path = path.as_deref() == Some("circle");
                self.rectangular_path = path.as_deref() == Some("rect");
            }
            "fillToRect" if self.circular_path || self.rectangular_path || self.shape_path => {
                self.fill_to_rectangle = Some(drawingml_relative_rectangle(attributes, part)?);
            }
            _ => {}
        }
        Ok(())
    }

    fn end(&mut self, local: &str) {
        if local == "gs"
            && let Some((offset, Some(color))) = self.current_stop.take()
        {
            self.stops.push(GradientStop { offset, color });
        }
        if matches!(local, "fgClr" | "bgClr") {
            self.pattern_color_target = PatternColorTarget::Foreground;
        }
    }

    fn into_paint(self, bounds: Rect) -> Paint {
        self.into_fill().paint(bounds)
    }

    fn into_fill(mut self) -> super::drawingml::ChartFill {
        use super::drawingml::ChartFill;
        match self.kind {
            PaintCaptureKind::Solid => self
                .color
                .map_or(ChartFill::Image(Paint::None), ChartFill::Solid),
            PaintCaptureKind::Pattern { preset } => ChartFill::Pattern {
                preset,
                foreground: self.color.unwrap_or(0x0000_00ff),
                background: self.background_color.unwrap_or(0xffff_ffff),
            },
            PaintCaptureKind::Gradient => {
                self.stops.sort_by(|a, b| a.offset.total_cmp(&b.offset));
                if self.stops.is_empty() {
                    return ChartFill::Image(Paint::None);
                }
                let fill = if self.shape_path {
                    ChartFill::ShapeGradient {
                        focus: self.fill_to_rectangle,
                        stops: self.stops,
                    }
                } else if self.circular_path || self.rectangular_path {
                    ChartFill::PathGradient {
                        circular: self.circular_path,
                        fill_to_rectangle: self.fill_to_rectangle,
                        stops: self.stops,
                    }
                } else {
                    ChartFill::LinearGradient {
                        angle_degrees: self.angle_degrees,
                        angle_scaled: self.angle_scaled,
                        stops: self.stops,
                    }
                };
                ChartFill::MappedGradient {
                    fill: Box::new(fill),
                    mapping: self.gradient_mapping,
                }
            }
        }
    }
}

fn visit_pptx_theme_style(
    package: &Package<'_>,
    theme: &PptxTheme,
    list: &str,
    index: u64,
    visit: impl FnMut(XmlEvent<'_>, usize) -> Result<(), Diagnostic>,
) -> Result<(), Diagnostic> {
    let Some(part) = theme.part.as_deref() else {
        return Ok(());
    };
    super::drawingml::visit_drawingml_theme_style(package, part, list, index, visit)
}

fn pptx_theme_fill_slot(index: u64) -> Option<(&'static str, u64)> {
    match index {
        1..=999 => Some(("fillStyleLst", index - 1)),
        1001.. => Some(("bgFillStyleLst", index - 1001)),
        _ => None,
    }
}

fn resolve_pptx_theme_fill(
    package: &Package<'_>,
    theme: &PptxTheme,
    index: u64,
    placeholder: u32,
) -> Result<Option<super::drawingml::ChartFill>, Diagnostic> {
    let Some((list, index)) = pptx_theme_fill_slot(index) else {
        return Ok(None);
    };
    let part = theme.part.as_deref().unwrap_or(TABLE_STYLES_PART);
    let mut fill_theme = theme.clone();
    fill_theme.colors.insert("phClr".to_owned(), placeholder);
    let mut paint = None;
    let mut capture: Option<PaintCapture> = None;
    visit_pptx_theme_style(package, theme, list, index, |event, depth| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if let Some(capture) = capture.as_mut() {
                    capture.start(local, &attributes, part, &fill_theme)?;
                } else if local == "noFill" {
                    paint = Some(super::drawingml::ChartFill::Image(Paint::None));
                } else if let Some(current) = PaintCapture::new(local, &attributes, part, depth)? {
                    if empty {
                        paint = Some(current.into_fill());
                    } else {
                        capture = Some(current);
                    }
                }
            }
            XmlEvent::EndElement { name } => {
                if let Some(capture) = capture.as_mut() {
                    capture.end(local_name(name));
                }
                if capture.as_ref().is_some_and(|c| c.depth == depth) {
                    paint = capture.take().map(|c| c.into_fill());
                }
            }
            _ => {}
        }
        Ok(())
    })?;
    Ok(paint)
}

fn resolve_pptx_theme_image_fill(
    package: &Package<'_>,
    theme: &PptxTheme,
    index: u64,
) -> Result<Option<ShapeImageFillState>, Diagnostic> {
    let Some((list, index)) = pptx_theme_fill_slot(index) else {
        return Ok(None);
    };
    let mut image: Option<ShapeImageFillState> = None;
    visit_pptx_theme_style(package, theme, list, index, |event, depth| {
        if let XmlEvent::StartElement {
            name, attributes, ..
        } = event
        {
            if let Some(image) = image.as_mut() {
                image.start(
                    local_name(name),
                    &attributes,
                    theme.part.as_deref().unwrap_or(TABLE_STYLES_PART),
                )?;
            } else if local_name(name) == "blipFill" {
                image = Some(ShapeImageFillState::new(depth));
            }
        }
        Ok(())
    })?;
    Ok(image)
}

fn resolve_pptx_theme_effects(
    package: &Package<'_>,
    theme: &PptxTheme,
    index: u64,
    placeholder: u32,
) -> Result<Option<super::drawingml::DrawingMlPictureEffects>, Diagnostic> {
    let Some(index) = index.checked_sub(1) else {
        return Ok(None);
    };
    let Some(part) = theme.part.as_deref() else {
        return Ok(None);
    };
    let mut theme_colors = theme.clone();
    theme_colors.colors.insert("phClr".to_owned(), placeholder);
    super::drawingml::drawingml_theme_effects(package, part, index, |kind, value| {
        picture_effect_color(kind, value, &theme_colors, part)
    })
}

struct PptxParseState {
    objects: Vec<Object>,
    next_z: i32,
    diagnostics: Vec<Diagnostic>,
    image_cache: HashMap<String, ImageCacheEntry>,
    materialized_image_bytes: usize,
    reported_unsupported_backgrounds: HashSet<String>,
    placeholders: HashMap<(u32, String), u32>,
    placeholder_text_styles: HashMap<u32, Vec<ParagraphStyle>>,
    placeholder_presets: HashMap<u32, (String, HashMap<String, f32>)>,
    pending_placeholder_objects: HashSet<u32>,
    default_theme: PptxTheme,
    default_text_styles: Vec<ParagraphStyle>,
    theme_cache: HashMap<String, PptxTheme>,
}

impl PptxParseState {
    fn take_z(&mut self) -> i32 {
        let z = self.next_z;
        self.next_z = self.next_z.saturating_add(1);
        z
    }
}
#[derive(Clone, Debug)]
struct PresentationPartProperties {
    has_background: bool,
    show_master_shapes: bool,
    shown: bool,
    name: Option<String>,
    color_map: Option<HashMap<String, String>>,
    has_transition: bool,
    has_timing: bool,
    unsupported_features: HashSet<String>,
}

#[derive(Clone, Copy)]
struct PartContent {
    backgrounds: bool,
    pictures: bool,
    shapes: bool,
    inherited: bool,
}

impl PartContent {
    const BACKGROUND: Self = Self {
        backgrounds: true,
        pictures: false,
        shapes: false,
        inherited: false,
    };
    const INHERITED: Self = Self {
        backgrounds: false,
        pictures: true,
        shapes: true,
        inherited: true,
    };
    const SLIDE: Self = Self {
        backgrounds: false,
        pictures: true,
        shapes: true,
        inherited: false,
    };
}

struct SlideParseContext<'a> {
    part: &'a str,
    unit_index: u32,
    properties: PresentationPartProperties,
    bounds: Rect,
}

fn parse_slide(
    package: &Package<'_>,
    context: SlideParseContext<'_>,
    state: &mut PptxParseState,
    content_types: &ContentTypes,
    font_metrics: &FontMetricTable,
) -> Result<(), Diagnostic> {
    let SlideParseContext {
        part,
        unit_index,
        properties: slide_properties,
        bounds: slide_bounds,
    } = context;
    let relationships = package.relationships(Some(part))?;
    report_unsupported_slide_features(
        part,
        &relationships,
        &slide_properties.unsupported_features,
        &mut state.diagnostics,
    );
    let layout_part = related_presentation_part(
        &relationships,
        "/slideLayout",
        "slide layout",
        part,
        &mut state.diagnostics,
    )?;
    let layout_relationships = if let Some(layout_part) = layout_part.as_deref() {
        package.relationships(Some(layout_part))?
    } else {
        Vec::new()
    };
    let master_part = if let Some(layout_part) = layout_part.as_deref() {
        related_presentation_part(
            &layout_relationships,
            "/slideMaster",
            "slide master",
            layout_part,
            &mut state.diagnostics,
        )?
    } else {
        None
    };
    let master_relationships = if let Some(master_part) = master_part.as_deref() {
        package.relationships(Some(master_part))?
    } else {
        Vec::new()
    };
    let theme_part = if let Some(master_part) = master_part.as_deref() {
        related_presentation_part(
            &master_relationships,
            "/theme",
            "slide theme",
            master_part,
            &mut state.diagnostics,
        )?
    } else {
        None
    }
    .or(if let Some(layout_part) = layout_part.as_deref() {
        related_presentation_part(
            &layout_relationships,
            "/theme",
            "slide theme",
            layout_part,
            &mut state.diagnostics,
        )?
    } else {
        None
    })
    .or(related_presentation_part(
        &relationships,
        "/theme",
        "slide theme",
        part,
        &mut state.diagnostics,
    )?);
    let theme = if let Some(theme_part) = theme_part {
        if let Some(theme) = state.theme_cache.get(&theme_part) {
            theme.clone()
        } else {
            let theme = parse_pptx_theme(package, &theme_part)?;
            state.theme_cache.insert(theme_part, theme.clone());
            theme
        }
    } else {
        state.default_theme.clone()
    };
    let layout_properties = layout_part
        .as_deref()
        .map(|layout_part| presentation_part_properties(package, layout_part))
        .transpose()?;
    let master_properties = master_part
        .as_deref()
        .map(|master_part| presentation_part_properties(package, master_part))
        .transpose()?;
    let mut color_map = default_color_map();
    if let Some(mapping) = master_properties
        .as_ref()
        .and_then(|properties| properties.color_map.as_ref())
    {
        color_map.clone_from(mapping);
    }
    if let Some(mapping) = layout_properties
        .as_ref()
        .and_then(|properties| properties.color_map.as_ref())
    {
        color_map.clone_from(mapping);
    }
    if let Some(mapping) = slide_properties.color_map.as_ref() {
        color_map.clone_from(mapping);
    }
    let theme = theme.with_color_map(color_map);
    let master_text_styles = master_part
        .as_deref()
        .map(|master_part| parse_master_text_styles(package, master_part, &theme))
        .transpose()?;
    if slide_properties.has_transition {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Unsupported,
                "PPTX slide transition is preserved as metadata but is not played",
            )
            .in_part(part)
            .with_detail("feature", "slide-transition"),
        );
    }
    if slide_properties.has_timing {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Unsupported,
                "PPTX animation timing is preserved structurally but is not played",
            )
            .in_part(part)
            .with_detail("feature", "animation-timing"),
        );
    }
    let background_part = if slide_properties.has_background {
        Some(part)
    } else if layout_properties
        .as_ref()
        .is_some_and(|properties| properties.has_background)
    {
        layout_part.as_deref()
    } else if master_properties
        .as_ref()
        .is_some_and(|properties| properties.has_background)
    {
        master_part.as_deref()
    } else {
        None
    };
    if let Some(background_part) = background_part {
        parse_presentation_part(
            package,
            background_part,
            unit_index,
            slide_bounds,
            PartContent::BACKGROUND,
            state,
            content_types,
            font_metrics,
            &theme,
            master_text_styles.as_ref(),
            None,
        )?;
    }
    let show_master_shapes = slide_properties.show_master_shapes
        && layout_properties
            .as_ref()
            .is_none_or(|properties| properties.show_master_shapes);
    if show_master_shapes && let Some(master_part) = master_part.as_deref() {
        parse_presentation_part(
            package,
            master_part,
            unit_index,
            slide_bounds,
            PartContent::INHERITED,
            state,
            content_types,
            font_metrics,
            &theme,
            master_text_styles.as_ref(),
            None,
        )?;
    }
    if let Some(layout_part) = layout_part.as_deref() {
        parse_presentation_part(
            package,
            layout_part,
            unit_index,
            slide_bounds,
            PartContent::INHERITED,
            state,
            content_types,
            font_metrics,
            &theme,
            master_text_styles.as_ref(),
            None,
        )?;
    }
    parse_presentation_part(
        package,
        part,
        unit_index,
        slide_bounds,
        PartContent::SLIDE,
        state,
        content_types,
        font_metrics,
        &theme,
        master_text_styles.as_ref(),
        None,
    )?;
    parse_slide_ink(package, &relationships, unit_index, state)?;
    report_blocked_pptx_actions(unit_index, state);
    hide_uninstantiated_prompt_objects(unit_index, state);
    Ok(())
}

type SpeakerNotes = (Option<String>, Option<String>, Vec<SpeakerNoteParagraph>);

fn parse_speaker_notes(
    package: &Package<'_>,
    slide_part: &str,
    theme: &PptxTheme,
) -> Result<SpeakerNotes, Diagnostic> {
    let relationships = package.relationships(Some(slide_part))?;
    let Some(relationship) = relationships
        .iter()
        .find(|relationship| relationship.type_uri.ends_with("/notesSlide"))
    else {
        return Ok((None, None, Vec::new()));
    };
    if relationship.external {
        return Ok((None, None, Vec::new()));
    }
    let part = relationship.target.as_str();
    let note_relationships = package.relationships(Some(part))?;
    let notes_master_part = note_relationships
        .iter()
        .find(|relationship| {
            !relationship.external && relationship.type_uri.ends_with("/notesMaster")
        })
        .map(|relationship| relationship.target.as_str());
    let mut resolved_theme = theme.clone();
    if let Some(notes_master_part) = notes_master_part {
        let notes_master_relationships = package.relationships(Some(notes_master_part))?;
        if let Some(theme_part) = notes_master_relationships
            .iter()
            .find(|relationship| {
                !relationship.external && relationship.type_uri.ends_with("/theme")
            })
            .map(|relationship| relationship.target.as_str())
        {
            resolved_theme = parse_pptx_theme(package, theme_part)?;
        }
        if let Some(color_map) = presentation_part_properties(package, notes_master_part)?.color_map
        {
            resolved_theme = resolved_theme.with_color_map(color_map);
        }
    }
    if let Some(color_map) = presentation_part_properties(package, part)?.color_map {
        resolved_theme = resolved_theme.with_color_map(color_map);
    }
    let note_styles = notes_master_part
        .map(|notes_master_part| {
            parse_notes_master_styles(package, notes_master_part, &resolved_theme)
        })
        .transpose()?
        .unwrap_or_else(|| vec![ParagraphStyle::default(); TEXT_LEVEL_COUNT]);
    let theme = &resolved_theme;
    let bytes = package.required_part(part)?;
    let mut depth = 0_usize;
    let mut shape: Option<ShapeState> = None;
    let mut ignored_shape = false;
    let mut highlight_depth: Option<usize> = None;
    let mut paragraphs = Vec::new();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "sp" && shape.is_none() {
                    let mut current = ShapeState::new(depth, theme, false);
                    current.paragraph_styles.clone_from(&note_styles);
                    shape = Some(current);
                    ignored_shape = false;
                }
                if let Some(current) = shape.as_mut() {
                    match local {
                        "ph" => {
                            let placeholder_type = string_attribute(&attributes, "type", part)?;
                            ignored_shape = matches!(
                                placeholder_type.as_deref(),
                                Some("sldImg" | "sldNum" | "hdr" | "ftr" | "dt")
                            );
                            current.placeholder_type = placeholder_type;
                        }
                        "p" => {
                            current.paragraph_count = current.paragraph_count.saturating_add(1);
                            current.pending_bullet_style = None;
                            current.current_paragraph_style = ParagraphStyle::default();
                            current.current_paragraph_level = 0;
                            current.paragraph_pending = true;
                        }
                        local if paragraph_level(local).is_some() => {
                            current.paragraph_style_capture = Some(ParagraphStyleCapture {
                                depth,
                                target: ParagraphStyleTarget::Level(
                                    paragraph_level(local).unwrap_or(0),
                                ),
                                style: paragraph_style_from_attributes(&attributes, part)?,
                                spacing: None,
                                bullet_color_depth: None,
                                text_effect_depth: None,
                            });
                        }
                        "defPPr" => {
                            current.paragraph_style_capture = Some(ParagraphStyleCapture {
                                depth,
                                target: ParagraphStyleTarget::Default,
                                style: paragraph_style_from_attributes(&attributes, part)?,
                                spacing: None,
                                bullet_color_depth: None,
                                text_effect_depth: None,
                            });
                        }
                        "pPr" => {
                            current.current_paragraph_level =
                                numeric_attribute(&attributes, "lvl", part)?
                                    .and_then(|value| usize::try_from(value).ok())
                                    .unwrap_or(0)
                                    .min(TEXT_LEVEL_COUNT.saturating_sub(1));
                            current.paragraph_style_capture = Some(ParagraphStyleCapture {
                                depth,
                                target: ParagraphStyleTarget::Current,
                                style: paragraph_style_from_attributes(&attributes, part)?,
                                spacing: None,
                                bullet_color_depth: None,
                                text_effect_depth: None,
                            });
                        }
                        "lnSpc" | "spcBef" | "spcAft" | "spcPct" | "spcPts" | "buNone"
                        | "buChar" | "buAutoNum" | "buFont" | "buSzPct" | "buSzPts" | "buClr"
                            if current.paragraph_style_capture.is_some() =>
                        {
                            update_paragraph_style_capture(
                                current.paragraph_style_capture.as_mut().ok_or_else(|| {
                                    format_error(part, "speaker-note paragraph style is missing")
                                })?,
                                local,
                                &attributes,
                                depth,
                                part,
                                theme,
                            )?;
                        }
                        "srgbClr" | "schemeClr"
                            if current.run_properties_depth.is_some()
                                || highlight_depth.is_some() =>
                        {
                            let color = if local == "srgbClr" {
                                string_attribute(&attributes, "val", part)?
                                    .and_then(|value| parse_rgb_color(&value))
                            } else {
                                string_attribute(&attributes, "val", part)?
                                    .and_then(|value| theme.color(&value))
                            };
                            if let Some(color) = color {
                                if highlight_depth.is_some() {
                                    if let Some(run) = current.run.as_mut() {
                                        run.highlight = color;
                                    } else {
                                        current.highlight = color;
                                    }
                                } else if let Some(run) = current.run.as_mut() {
                                    run.color = color;
                                } else {
                                    current.font_color = color;
                                }
                            }
                        }
                        "r" | "fld" => {
                            begin_shape_paragraph(current, true);
                            current.run = Some(TextRunState::from_shape(current, depth));
                        }
                        "br" => {
                            begin_shape_paragraph(current, true);
                            append_shape_line_break(current, depth, '\u{2028}');
                        }
                        "rPr" | "defRPr" => {
                            begin_shape_run_properties(current, &attributes, depth, empty, part)?;
                        }
                        "latin" | "ea" | "cs" | "sym" => {
                            apply_shape_run_typeface(current, local, &attributes, part, theme)?;
                        }
                        "highlight" => highlight_depth = (!empty).then_some(depth),
                        "t" => current.collecting_text = !empty,
                        _ => {}
                    }
                    if empty
                        && current
                            .paragraph_style_capture
                            .as_ref()
                            .is_some_and(|capture| capture.depth == depth)
                    {
                        commit_shape_paragraph_style(current);
                    }
                    if empty && local == "p" && current.paragraph_pending {
                        begin_shape_paragraph(current, false);
                    }
                    if empty && matches!(local, "r" | "fld") {
                        finish_speaker_note_run(current, part)?;
                    }
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if let Some(current) = shape.as_mut() {
                    if local == "t" {
                        current.collecting_text = false;
                    }
                    if let Some(capture) = current.paragraph_style_capture.as_mut()
                        && capture
                            .spacing
                            .is_some_and(|(spacing_depth, _)| spacing_depth == depth)
                    {
                        capture.spacing = None;
                    }
                    if let Some(capture) = current.paragraph_style_capture.as_mut()
                        && capture.bullet_color_depth == Some(depth)
                    {
                        capture.bullet_color_depth = None;
                    }
                    if current
                        .paragraph_style_capture
                        .as_ref()
                        .is_some_and(|capture| capture.depth == depth)
                    {
                        commit_shape_paragraph_style(current);
                    }
                    if local == "p" && current.paragraph_pending {
                        begin_shape_paragraph(current, false);
                    }
                    if matches!(local, "rPr" | "defRPr")
                        && current.run_properties_depth == Some(depth)
                    {
                        current.run_properties_depth = None;
                        current.run_properties_target = None;
                    }
                    if matches!(local, "r" | "fld")
                        && current.run.as_ref().is_some_and(|run| run.depth == depth)
                    {
                        finish_speaker_note_run(current, part)?;
                    }
                }
                if local == "highlight" && highlight_depth == Some(depth) {
                    highlight_depth = None;
                }
                if local == "sp" && shape.as_ref().is_some_and(|current| current.depth == depth) {
                    let current = shape.take().ok_or_else(|| {
                        format_error(part, "speaker-note shape parser state is missing")
                    })?;
                    if !ignored_shape {
                        paragraphs.extend(finish_speaker_note_shape(current));
                    }
                    ignored_shape = false;
                }
            }
            XmlEvent::Text(value) => {
                if let Some(current) = shape.as_mut().filter(|current| current.collecting_text) {
                    let value = decode_xml_text(value).map_err(|error| with_part(error, part))?;
                    append_shape_text(current, &value);
                }
            }
            XmlEvent::Cdata(value) => {
                if let Some(current) = shape.as_mut().filter(|current| current.collecting_text) {
                    append_shape_text(current, value);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    let text = paragraphs
        .iter()
        .map(|paragraph| {
            paragraph
                .runs
                .iter()
                .filter(|run| !run.text.ends_with('\t'))
                .map(|run| run.text.as_str())
                .collect::<String>()
        })
        .collect::<Vec<_>>()
        .join("\n")
        .trim()
        .to_owned();
    Ok((
        (!text.is_empty()).then_some(text),
        Some(part.to_owned()),
        paragraphs,
    ))
}

fn begin_shape_run_properties(
    shape: &mut ShapeState,
    attributes: &[XmlAttribute<'_>],
    depth: usize,
    empty: bool,
    part: &str,
) -> Result<(), Diagnostic> {
    let target = if shape.run.is_some() {
        RunPropertiesTarget::Current
    } else {
        RunPropertiesTarget::Default
    };
    shape.run_properties_depth = (!empty).then_some(depth);
    shape.run_properties_target = (!empty).then_some(target);
    if let Some(size) = numeric_attribute(attributes, "sz", part)? {
        let size = size as f32 / 100.0 * POINTS_TO_CSS_PIXELS;
        if let Some(run) = shape.run.as_mut() {
            run.font_size = size;
        } else {
            shape.font_size = size;
        }
    }
    for (attribute, target) in [("b", 0_u8), ("i", 1_u8)] {
        if let Some(value) = boolean_attribute(attributes, attribute, part)? {
            if let Some(run) = shape.run.as_mut() {
                if target == 0 {
                    run.bold = value
                } else {
                    run.italic = value
                }
            } else if target == 0 {
                shape.bold = value;
            } else {
                shape.italic = value;
            }
        }
    }
    if let Some(value) = string_attribute(attributes, "u", part)? {
        let (underline, wavy, dotted, heavy, double, dot_dash) = drawingml_underline(&value);
        if let Some(run) = shape.run.as_mut() {
            run.underline = underline;
            run.wavy_underline = wavy;
            run.dotted_underline = dotted;
            run.heavy_underline = heavy;
            run.double_underline = double;
            run.dot_dash_underline = dot_dash;
        } else {
            shape.underline = underline
        }
    }
    if let Some(value) = string_attribute(attributes, "strike", part)? {
        let double = value == "dblStrike";
        let value = !matches!(value.as_str(), "noStrike" | "none" | "0" | "false");
        if let Some(run) = shape.run.as_mut() {
            run.strikethrough = value;
            run.double_strikethrough = double;
        } else {
            shape.strikethrough = value;
            shape.inherited_run_effect.double_strikethrough = double;
        }
    }
    if let Some(value) = signed_numeric_attribute(attributes, "baseline", part)? {
        let ratio = value as f32 / 100_000.0;
        if let Some(run) = shape.run.as_mut() {
            run.baseline_shift = run.font_size * ratio;
            if value != 0 {
                run.font_size *= 0.65;
            }
        } else {
            shape.baseline_shift = shape.font_size * ratio;
            if value != 0 {
                shape.font_size *= 0.65;
            }
        }
    }
    if let Some(value) = signed_numeric_attribute(attributes, "spc", part)? {
        let value = value as f32 / 100.0 * POINTS_TO_CSS_PIXELS;
        if let Some(run) = shape.run.as_mut() {
            run.letter_spacing = value
        } else {
            shape.letter_spacing = value
        }
    }
    if let Some(value) = string_attribute(attributes, "cap", part)? {
        let capitalization = match value.as_str() {
            "all" => TextCapitalization::All,
            "small" => TextCapitalization::Small,
            _ => TextCapitalization::None,
        };
        if let Some(run) = shape.run.as_mut() {
            run.capitalization = capitalization;
        } else {
            shape.capitalization = capitalization;
        }
    }
    Ok(())
}

fn apply_shape_run_typeface(
    shape: &mut ShapeState,
    script: &str,
    attributes: &[XmlAttribute<'_>],
    part: &str,
    theme: &PptxTheme,
) -> Result<(), Diagnostic> {
    let Some(typeface) = string_attribute(attributes, "typeface", part)? else {
        return Ok(());
    };
    if typeface.is_empty() {
        return Ok(());
    }
    let typeface = theme.resolve_typeface(&typeface, script);
    if matches!(
        shape.run_properties_target,
        Some(RunPropertiesTarget::Current)
    ) {
        if let Some(run) = shape.run.as_mut() {
            run.set_font(script, typeface);
        }
    } else {
        match script {
            "ea" => shape.font_east_asian = Some(typeface),
            "cs" => shape.font_complex_script = Some(typeface),
            "sym" => {}
            _ => shape.font_family = typeface,
        }
    }
    Ok(())
}

fn finish_speaker_note_run(shape: &mut ShapeState, part: &str) -> Result<(), Diagnostic> {
    let runs = shape
        .run
        .take()
        .ok_or_else(|| format_error(part, "speaker-note text run parser state is missing"))?
        .finish();
    if let Some(bullet) = shape
        .runs
        .last_mut()
        .filter(|candidate| candidate.text.ends_with('\t'))
        && let Some(run) = runs.first()
    {
        bullet.font_size = shape
            .paragraph_layout_styles
            .last()
            .and_then(|style| style.bullet_size.as_ref())
            .map_or(run.font_size, |size| size.resolve(run.font_size));
    }
    shape.run_effects.resize(
        shape.run_effects.len().saturating_add(runs.len()),
        TextEffect::default(),
    );
    shape.runs.extend(runs);
    Ok(())
}

fn finish_speaker_note_shape(shape: ShapeState) -> Vec<SpeakerNoteParagraph> {
    let mut output = Vec::new();
    for (index, layout) in shape.paragraph_layouts.iter().copied().enumerate() {
        let start = shape
            .paragraph_run_starts
            .get(index)
            .copied()
            .unwrap_or(0)
            .min(shape.runs.len());
        let end = shape
            .paragraph_run_starts
            .get(index + 1)
            .copied()
            .unwrap_or(shape.runs.len())
            .min(shape.runs.len());
        let runs = shape.runs[start..end].to_vec();
        if runs.iter().all(|run| run.text.is_empty()) {
            continue;
        }
        let font_size = runs
            .iter()
            .filter(|run| !run.text.trim().is_empty())
            .map(|run| run.font_size)
            .reduce(f32::max)
            .unwrap_or(shape.font_size)
            .max(1.0);
        let style = shape.paragraph_layout_styles.get(index);
        let mut layout = layout;
        layout.line_height = style
            .and_then(|style| style.line_spacing.as_ref())
            .map_or(font_size * 1.2, |spacing| {
                spacing.resolve_for_layout(Some(font_size))
            });
        layout.space_before = style
            .and_then(|style| style.space_before.as_ref())
            .map_or(0.0, |spacing| spacing.resolve(font_size));
        layout.space_after = style
            .and_then(|style| style.space_after.as_ref())
            .map_or(0.0, |spacing| spacing.resolve(font_size));
        output.push(SpeakerNoteParagraph { layout, runs });
    }
    output
}

// VML ink has authored vector paths; the ISF payload is unnecessary for static viewing.
// Other VML objects and grouped coordinates remain outside this adapter's boundary.
fn parse_slide_ink(
    package: &Package<'_>,
    relationships: &[Relationship],
    unit_index: u32,
    state: &mut PptxParseState,
) -> Result<(), Diagnostic> {
    for relationship in relationships
        .iter()
        .filter(|r| r.type_uri.ends_with("/vmlDrawing"))
    {
        let part = relationship.target.as_str();
        if relationship.external {
            state.diagnostics.push(unsupported_slide_feature_diagnostic(
                part,
                "legacy-vml-drawing",
            ));
            continue;
        }
        let result = (|| {
            let bytes = package.required_part(part)?;
            let mut depth = 0;
            let mut shape = None;
            let mut ink = false;
            let mut cap = LineCap::Flat;
            let mut opacity = Ok(1.0);
            parse_xml(&bytes, package.limits(), |event| {
                match event {
                    XmlEvent::StartElement {
                        name,
                        attributes,
                        empty,
                    } => {
                        depth += 1;
                        match local_name(name) {
                            "shape" if depth == 2 && !empty => {
                                shape = Some(parse_vml_ink_shape(&attributes, part));
                                ink = false;
                                cap = LineCap::Flat;
                                opacity = Ok(1.0);
                            }
                            "ink" if shape.is_some() => ink = true,
                            "stroke" if shape.is_some() => {
                                opacity = super::vml::parse_vml_opacity(&attributes, part);
                                cap = match string_attribute(&attributes, "endcap", part)?
                                    .as_deref()
                                {
                                    Some("round") => LineCap::Round,
                                    Some("square") => LineCap::Square,
                                    _ => LineCap::Flat,
                                };
                            }
                            "group" if depth == 2 => state.diagnostics.push(
                                unsupported_slide_feature_diagnostic(part, "legacy-vml-drawing"),
                            ),
                            _ => {}
                        }
                        if empty {
                            depth -= 1;
                        }
                    }
                    XmlEvent::EndElement { name } => {
                        if local_name(name) == "shape" && depth == 2 {
                            if let Some(shape) = shape.take() {
                                let parsed = shape
                                    .and_then(|value| opacity.clone().map(|alpha| (value, alpha)));
                                match parsed {
                                    Ok(((bounds, geometry, color, width, shape_id), alpha))
                                        if ink =>
                                    {
                                        if state.objects.len()
                                            >= package.limits().max_document_objects
                                        {
                                            return Err(Diagnostic::fatal(
                                                DiagnosticCode::ObjectLimit,
                                                Phase::Parse,
                                                None,
                                                "document exceeds the configured object limit",
                                            )
                                            .in_part(part));
                                        }
                                        let numeric_id = u32::try_from(state.objects.len())
                                            .map_err(|_| {
                                                format_error(
                                                    part,
                                                    "object count exceeds supported range",
                                                )
                                            })?;
                                        let z = state.take_z();
                                        state.objects.push(Object {
                                            numeric_id,
                                            parent_numeric_id: None,
                                            stable_id: format!("object:{numeric_id}"),
                                            parent_stable_id: None,
                                            kind: ObjectKind::Shape,
                                            unit_index,
                                            bounds,
                                            z,
                                            text: None,
                                            source: SourceRef {
                                                part: part.to_owned(),
                                                mapping: MappingQuality::Exact,
                                                locator: SourceLocator::PptxShape {
                                                    shape_id,
                                                    row: None,
                                                    column: None,
                                                    text_range: None,
                                                    metadata: PptxObjectMetadata::default(),
                                                },
                                            },
                                            visual: Visual::StrokeStyle {
                                                style: StrokeStyle {
                                                    cap,
                                                    join: LineJoin::Round,
                                                    ..StrokeStyle::default()
                                                },
                                                visual: Box::new(Visual::Shape {
                                                    geometry,
                                                    fill: 0,
                                                    stroke: (color & 0xffffff00)
                                                        | ((color & 255) as f32 * alpha).round()
                                                            as u32,
                                                    stroke_width: width,
                                                }),
                                            },
                                        });
                                    }
                                    result => {
                                        let mut diagnostic = unsupported_slide_feature_diagnostic(
                                            part,
                                            "legacy-vml-drawing",
                                        );
                                        if let Err(error) = result {
                                            diagnostic =
                                                diagnostic.with_detail("reason", error.message);
                                        }
                                        state.diagnostics.push(diagnostic);
                                    }
                                }
                            }
                        }
                        depth -= 1;
                    }
                    _ => {}
                }
                Ok(())
            })?;
            Ok(())
        })();
        if let Err(error) = result {
            let error: Diagnostic = error;
            if !matches!(
                error.code,
                DiagnosticCode::FormatInvalid
                    | DiagnosticCode::XmlInvalid
                    | DiagnosticCode::XmlEncodingUnsupported
            ) {
                return Err(error);
            }
            state.diagnostics.push(
                unsupported_slide_feature_diagnostic(part, "legacy-vml-drawing")
                    .with_detail("reason", error.message),
            );
        }
    }
    Ok(())
}

fn parse_vml_ink_shape(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<(Rect, Geometry, u32, f32, u32), Diagnostic> {
    use super::vml::{parse_vml_color, parse_vml_length, parse_vml_path_geometry};
    let attr = |name| string_attribute(attributes, name, part);
    let mut bounds = Rect::default();
    for declaration in attr("style")?.unwrap_or_default().split(';') {
        let Some((key, value)) = declaration.split_once(':') else {
            continue;
        };
        let target = match key.trim() {
            "left" => &mut bounds.x,
            "top" => &mut bounds.y,
            "width" => &mut bounds.width,
            "height" => &mut bounds.height,
            "rotation" | "flip" => {
                return Err(format_error(part, "transformed VML ink is unsupported"));
            }
            _ => continue,
        };
        *target = parse_vml_length(value, part)?;
    }
    if bounds.width <= 0.0 || bounds.height <= 0.0 {
        return Err(format_error(part, "VML ink has invalid bounds"));
    }
    let geometry = parse_vml_path_geometry(
        &attr("path")?.unwrap_or_default(),
        attr("coordsize")?.as_deref(),
        attr("coordorigin")?.as_deref(),
        bounds.width,
        bounds.height,
    )
    .ok_or_else(|| format_error(part, "unsupported VML ink path"))?;
    let color = parse_vml_color(attr("strokecolor")?.as_deref().unwrap_or("black"), part)?;
    let width = parse_vml_length(attr("strokeweight")?.as_deref().unwrap_or("1pt"), part)?;
    let shape_id = attr("id")?
        .and_then(|id| id.rsplit('_').next()?.strip_prefix('s')?.parse().ok())
        .ok_or_else(|| format_error(part, "VML ink shape ID is invalid"))?;
    Ok((bounds, geometry, color, width.max(0.0), shape_id))
}
fn report_unsupported_slide_features(
    part: &str,
    relationships: &[Relationship],
    inline_features: &HashSet<String>,
    diagnostics: &mut Vec<Diagnostic>,
) {
    let mut reported = HashSet::new();
    for feature in inline_features {
        if reported.insert(feature.clone()) {
            diagnostics.push(unsupported_slide_feature_diagnostic(part, feature));
        }
    }
    for relationship in relationships {
        let feature = if relationship.type_uri.ends_with("/notesSlide") {
            None
        } else if relationship.type_uri.ends_with("/comments") {
            Some("comments")
        } else if relationship.type_uri.ends_with("/tags") {
            Some("slide-tags")
        } else if relationship.type_uri.ends_with("/audio")
            || relationship.type_uri.ends_with("/video")
            || relationship.type_uri.ends_with("/media")
        {
            Some("media-playback")
        } else if relationship.type_uri.ends_with("/package") {
            Some("embedded-object")
        } else if relationship.type_uri.ends_with("/control")
            || relationship.type_uri.ends_with("/controlProperties")
            || relationship.type_uri.ends_with("/activeX")
        {
            Some("activex-control")
        } else if relationship.type_uri.ends_with("/model3d") {
            Some("three-dimensional-model")
        } else {
            None
        };
        let Some(feature) = feature else {
            continue;
        };
        let report_key = if feature == "media-playback" {
            format!("{feature}:{}", relationship.target)
        } else {
            feature.to_owned()
        };
        if reported.insert(report_key) {
            diagnostics.push(
                unsupported_slide_feature_diagnostic(part, feature)
                    .with_detail("relationshipType", relationship.type_uri.clone())
                    .with_detail("target", relationship.target.clone()),
            );
        }
    }
}

fn unsupported_slide_feature_diagnostic(part: &str, feature: &str) -> Diagnostic {
    let message = match feature {
        "speaker-notes" => "PPTX speaker notes are retained in the package but are not rendered",
        "comments" => "PPTX comments are retained in the package but are not rendered",
        "slide-tags" => "PPTX slide tags are retained in the package but are not exposed",
        "media-playback" => {
            "PPTX audio/video relationship is not attached to a supported media object"
        }
        "timed-media" => "PPTX timed audio/video sequencing is not supported",
        "embedded-object" => "PPTX embedded OLE or package objects are blocked and not rendered",
        "activex-control" => "PPTX ActiveX and form controls are blocked and not rendered",
        "three-dimensional-model" => "PPTX embedded 3D models are not rendered",
        "legacy-vml-drawing" => "PPTX legacy VML drawing content is not rendered",
        _ => "PPTX content uses an unsupported slide feature",
    };
    let (code, phase, fidelity) = match feature {
        "embedded-object" | "activex-control" => (
            DiagnosticCode::ActiveContentBlocked,
            Phase::Security,
            Fidelity::Blocked,
        ),
        _ => (
            DiagnosticCode::UnsupportedFeature,
            Phase::Render,
            Fidelity::Omitted,
        ),
    };
    Diagnostic::warning(code, phase, fidelity, message)
        .in_part(part)
        .with_detail("feature", feature)
}

fn report_blocked_pptx_actions(unit_index: u32, state: &mut PptxParseState) {
    let mut diagnostics = Vec::new();
    for object in state
        .objects
        .iter()
        .filter(|object| object.unit_index == unit_index)
    {
        let SourceLocator::PptxShape { metadata, .. } = &object.source.locator else {
            continue;
        };
        for (trigger, action) in [
            ("click", metadata.click_action.as_ref()),
            ("hover", metadata.hover_action.as_ref()),
        ] {
            let Some(action) = action else {
                continue;
            };
            if action.kind == "command" {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::ActiveContentBlocked,
                        Phase::Security,
                        Fidelity::Blocked,
                        "PPTX command or macro action was exposed as metadata but execution is blocked",
                    )
                    .in_part(&object.source.part)
                    .with_detail("feature", "native-action-command")
                    .with_detail("trigger", trigger)
                    .with_detail("objectId", object.stable_id.clone()),
                );
            } else if action.kind == "unknown" {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Parse,
                        Fidelity::Omitted,
                        "PPTX action could not be classified and is exposed as read-only metadata",
                    )
                    .in_part(&object.source.part)
                    .with_detail("feature", "native-action-unknown")
                    .with_detail("trigger", trigger)
                    .with_detail("objectId", object.stable_id.clone()),
                );
            }
        }
    }
    state.diagnostics.extend(diagnostics);
}

#[derive(Clone, Copy)]
enum MasterTextStyleSection {
    Title,
    Body,
    Other,
}

fn parse_master_text_styles(
    package: &Package<'_>,
    part: &str,
    theme: &PptxTheme,
) -> Result<MasterTextStyles, Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut styles = MasterTextStyles::default();
    let mut section: Option<(usize, MasterTextStyleSection)> = None;
    let mut capture: Option<ParagraphStyleCapture> = None;
    let mut depth = 0_usize;
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                let section_kind = match local {
                    "titleStyle" => Some(MasterTextStyleSection::Title),
                    "bodyStyle" => Some(MasterTextStyleSection::Body),
                    "otherStyle" => Some(MasterTextStyleSection::Other),
                    _ => None,
                };
                if let Some(kind) = section_kind {
                    section = Some((depth, kind));
                } else if section.is_some()
                    && (local == "defPPr" || paragraph_level(local).is_some())
                {
                    capture = Some(ParagraphStyleCapture {
                        depth,
                        target: paragraph_level(local)
                            .map_or(ParagraphStyleTarget::Default, ParagraphStyleTarget::Level),
                        style: paragraph_style_from_attributes(&attributes, part)?,
                        spacing: None,
                        bullet_color_depth: None,
                        text_effect_depth: None,
                    });
                    if empty {
                        commit_master_paragraph_style(&mut styles, section, capture.take());
                    }
                } else if let Some(current) = capture.as_mut() {
                    update_paragraph_style_capture(
                        current,
                        local,
                        &attributes,
                        depth,
                        part,
                        theme,
                    )?;
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if let Some(current) = capture.as_mut()
                    && current
                        .spacing
                        .is_some_and(|(spacing_depth, _)| spacing_depth == depth)
                {
                    current.spacing = None;
                }
                if let Some(current) = capture.as_mut()
                    && current.bullet_color_depth == Some(depth)
                {
                    current.bullet_color_depth = None;
                }
                if capture
                    .as_ref()
                    .is_some_and(|current| current.depth == depth)
                {
                    commit_master_paragraph_style(&mut styles, section, capture.take());
                }
                if section.is_some_and(|(section_depth, _)| section_depth == depth)
                    && matches!(local, "titleStyle" | "bodyStyle" | "otherStyle")
                {
                    section = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(styles)
}

fn parse_default_text_styles(
    bytes: &[u8],
    limits: crate::limits::Limits,
    part: &str,
    theme: &PptxTheme,
) -> Result<Vec<ParagraphStyle>, Diagnostic> {
    let mut styles = vec![ParagraphStyle::default(); TEXT_LEVEL_COUNT];
    let mut section_depth = None;
    let mut capture: Option<ParagraphStyleCapture> = None;
    let mut depth = 0_usize;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "defaultTextStyle" {
                    section_depth = Some(depth);
                } else if section_depth.is_some()
                    && (local == "defPPr" || paragraph_level(local).is_some())
                {
                    capture = Some(ParagraphStyleCapture {
                        depth,
                        target: paragraph_level(local)
                            .map_or(ParagraphStyleTarget::Default, ParagraphStyleTarget::Level),
                        style: paragraph_style_from_attributes(&attributes, part)?,
                        spacing: None,
                        bullet_color_depth: None,
                        text_effect_depth: None,
                    });
                    if empty {
                        commit_notes_master_style(&mut styles, capture.take());
                    }
                } else if let Some(current) = capture.as_mut() {
                    update_paragraph_style_capture(
                        current,
                        local,
                        &attributes,
                        depth,
                        part,
                        theme,
                    )?;
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if let Some(current) = capture.as_mut()
                    && current
                        .spacing
                        .is_some_and(|(spacing_depth, _)| spacing_depth == depth)
                {
                    current.spacing = None;
                }
                if let Some(current) = capture.as_mut()
                    && current.bullet_color_depth == Some(depth)
                {
                    current.bullet_color_depth = None;
                }
                if capture
                    .as_ref()
                    .is_some_and(|current| current.depth == depth)
                {
                    commit_notes_master_style(&mut styles, capture.take());
                }
                if local == "defaultTextStyle" && section_depth == Some(depth) {
                    section_depth = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(styles)
}

fn parse_notes_master_styles(
    package: &Package<'_>,
    part: &str,
    theme: &PptxTheme,
) -> Result<Vec<ParagraphStyle>, Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut styles = vec![ParagraphStyle::default(); TEXT_LEVEL_COUNT];
    let mut notes_style_depth: Option<usize> = None;
    let mut capture: Option<ParagraphStyleCapture> = None;
    let mut depth = 0_usize;
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "notesStyle" {
                    notes_style_depth = Some(depth);
                } else if notes_style_depth.is_some()
                    && (local == "defPPr" || paragraph_level(local).is_some())
                {
                    capture = Some(ParagraphStyleCapture {
                        depth,
                        target: paragraph_level(local)
                            .map_or(ParagraphStyleTarget::Default, ParagraphStyleTarget::Level),
                        style: paragraph_style_from_attributes(&attributes, part)?,
                        spacing: None,
                        bullet_color_depth: None,
                        text_effect_depth: None,
                    });
                    if empty {
                        commit_notes_master_style(&mut styles, capture.take());
                    }
                } else if let Some(current) = capture.as_mut() {
                    update_paragraph_style_capture(
                        current,
                        local,
                        &attributes,
                        depth,
                        part,
                        theme,
                    )?;
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if let Some(current) = capture.as_mut()
                    && current
                        .spacing
                        .is_some_and(|(spacing_depth, _)| spacing_depth == depth)
                {
                    current.spacing = None;
                }
                if let Some(current) = capture.as_mut()
                    && current.bullet_color_depth == Some(depth)
                {
                    current.bullet_color_depth = None;
                }
                if capture
                    .as_ref()
                    .is_some_and(|current| current.depth == depth)
                {
                    commit_notes_master_style(&mut styles, capture.take());
                }
                if local == "notesStyle" && notes_style_depth == Some(depth) {
                    notes_style_depth = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(styles)
}

fn commit_notes_master_style(
    styles: &mut [ParagraphStyle],
    capture: Option<ParagraphStyleCapture>,
) {
    let Some(capture) = capture else {
        return;
    };
    match capture.target {
        ParagraphStyleTarget::Default => {
            for style in styles {
                style.apply(&capture.style);
            }
        }
        ParagraphStyleTarget::Level(level) => {
            if let Some(style) = styles.get_mut(level) {
                style.apply(&capture.style);
            }
        }
        ParagraphStyleTarget::Current => {}
    }
}

fn commit_master_paragraph_style(
    styles: &mut MasterTextStyles,
    section: Option<(usize, MasterTextStyleSection)>,
    capture: Option<ParagraphStyleCapture>,
) {
    let (Some((_, section)), Some(capture)) = (section, capture) else {
        return;
    };
    let levels = match section {
        MasterTextStyleSection::Title => &mut styles.title,
        MasterTextStyleSection::Body => &mut styles.body,
        MasterTextStyleSection::Other => &mut styles.other,
    };
    match capture.target {
        ParagraphStyleTarget::Default => {
            for style in levels {
                style.apply(&capture.style);
            }
        }
        ParagraphStyleTarget::Level(level) => {
            if let Some(style) = levels.get_mut(level) {
                style.apply(&capture.style);
            }
        }
        ParagraphStyleTarget::Current => {}
    }
}

fn paragraph_level(local: &str) -> Option<usize> {
    let value = local.strip_prefix("lvl")?.strip_suffix("pPr")?;
    let level = value.parse::<usize>().ok()?;
    (1..=TEXT_LEVEL_COUNT).contains(&level).then_some(level - 1)
}

fn paragraph_style_from_attributes(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<ParagraphStyle, Diagnostic> {
    Ok(ParagraphStyle {
        align: string_attribute(attributes, "algn", part)?.map(|value| match value.as_str() {
            "ctr" => TextAlign::Center,
            "r" => TextAlign::End,
            "just" => TextAlign::Justify,
            "justLow" => TextAlign::LowKashida,
            "dist" => TextAlign::Distribute,
            "thaiDist" => TextAlign::ThaiDistribute,
            _ => TextAlign::Start,
        }),
        margin_left: numeric_attribute(attributes, "marL", part)?
            .map(|value| value as f32 / EMU_PER_CSS_PIXEL),
        margin_right: numeric_attribute(attributes, "marR", part)?
            .map(|value| value as f32 / EMU_PER_CSS_PIXEL),
        indent: signed_numeric_attribute(attributes, "indent", part)?
            .map(|value| value as f32 / EMU_PER_CSS_PIXEL),
        default_tab_stop: numeric_attribute(attributes, "defTabSz", part)?
            .map(|value| value as f32 / EMU_PER_CSS_PIXEL),
        latin_line_break: boolean_attribute(attributes, "latinLnBrk", part)?,
        hanging_punctuation: boolean_attribute(attributes, "hangingPunct", part)?,
        east_asian_line_breaks: boolean_attribute(attributes, "eaLnBrk", part)?,
        ..ParagraphStyle::default()
    })
}

fn update_paragraph_style_capture(
    capture: &mut ParagraphStyleCapture,
    local: &str,
    attributes: &[XmlAttribute<'_>],
    depth: usize,
    part: &str,
    theme: &PptxTheme,
) -> Result<(), Diagnostic> {
    match local {
        "defRPr" => {
            capture.style.font_size = numeric_attribute(attributes, "sz", part)?
                .map(|value| value as f32 / 100.0 * POINTS_TO_CSS_PIXELS);
            capture.style.bold = boolean_attribute(attributes, "b", part)?;
            capture.style.italic = boolean_attribute(attributes, "i", part)?;
            capture.style.underline = string_attribute(attributes, "u", part)?
                .map(|value| !matches!(value.as_str(), "none" | "0" | "false"));
            capture.style.strikethrough = string_attribute(attributes, "strike", part)?
                .map(|value| !matches!(value.as_str(), "noStrike" | "none" | "0" | "false"));
            capture.style.baseline_ratio = signed_numeric_attribute(attributes, "baseline", part)?
                .map(|value| value as f32 / 100_000.0);
            capture.style.letter_spacing = signed_numeric_attribute(attributes, "spc", part)?
                .map(|value| value as f32 / 100.0 * POINTS_TO_CSS_PIXELS);
            capture.style.capitalization =
                string_attribute(attributes, "cap", part)?.map(|value| match value.as_str() {
                    "all" => TextCapitalization::All,
                    "small" => TextCapitalization::Small,
                    _ => TextCapitalization::None,
                });
        }
        "outerShdw" => {
            let distance = numeric_attribute(attributes, "dist", part)?.unwrap_or(0) as f32
                / EMU_PER_CSS_PIXEL;
            let direction =
                signed_numeric_attribute(attributes, "dir", part)?.unwrap_or(0) as f32 / 60_000.0;
            let radians = direction.to_radians();
            capture.text_effect_depth = Some(depth);
            capture.style.text_effect = Some(TextEffect {
                shadow: Some(Shadow {
                    color: 0x0000_0080,
                    blur: numeric_attribute(attributes, "blurRad", part)?.unwrap_or(0) as f32
                        / EMU_PER_CSS_PIXEL,
                    offset_x: distance * radians.cos(),
                    offset_y: distance * radians.sin(),
                }),
                shadow_scale_x: signed_numeric_attribute(attributes, "sx", part)?
                    .map_or(1.0, |value| value as f32 / 100_000.0),
                shadow_scale_y: signed_numeric_attribute(attributes, "sy", part)?
                    .map_or(1.0, |value| value as f32 / 100_000.0),
                shadow_skew_x: signed_numeric_attribute(attributes, "kx", part)?
                    .map_or(0.0, |value| value as f32 / 60_000.0),
                shadow_skew_y: signed_numeric_attribute(attributes, "ky", part)?
                    .map_or(0.0, |value| value as f32 / 60_000.0),
                shadow_alignment: drawingml_shadow_alignment(
                    string_attribute(attributes, "algn", part)?.as_deref(),
                ),
                ..TextEffect::default()
            });
        }
        "latin" | "ea" | "cs" => {
            if let Some(typeface) = string_attribute(attributes, "typeface", part)?
                .filter(|typeface| !typeface.is_empty())
            {
                let typeface = theme.resolve_typeface(&typeface, local);
                match local {
                    "ea" => capture.style.font_east_asian = Some(typeface),
                    "cs" => capture.style.font_complex_script = Some(typeface),
                    _ => capture.style.font_family = Some(typeface),
                }
            }
        }
        "lnSpc" => capture.spacing = Some((depth, ParagraphSpacingTarget::Line)),
        "spcBef" => capture.spacing = Some((depth, ParagraphSpacingTarget::Before)),
        "spcAft" => capture.spacing = Some((depth, ParagraphSpacingTarget::After)),
        "spcPct" | "spcPts" => {
            let Some((_, target)) = capture.spacing else {
                return Ok(());
            };
            let spacing = if local == "spcPct" {
                let Some(value) = optional_percentage_attribute(attributes, "val", part)? else {
                    return Ok(());
                };
                ParagraphSpacing::Percent(value)
            } else {
                let Some(value) = numeric_attribute(attributes, "val", part)? else {
                    return Ok(());
                };
                ParagraphSpacing::Points(value as f32 / 100.0 * POINTS_TO_CSS_PIXELS)
            };
            match target {
                ParagraphSpacingTarget::Line => capture.style.line_spacing = Some(spacing),
                ParagraphSpacingTarget::Before => capture.style.space_before = Some(spacing),
                ParagraphSpacingTarget::After => capture.style.space_after = Some(spacing),
            }
        }
        "buFont" => {
            capture.style.bullet_font_family = string_attribute(attributes, "typeface", part)?;
        }
        "buSzPct" | "buSzPts" => {
            if local == "buSzPct" {
                if let Some(value) = optional_percentage_attribute(attributes, "val", part)? {
                    capture.style.bullet_size = Some(ParagraphBulletSize::Percent(value));
                }
            } else if let Some(value) = numeric_attribute(attributes, "val", part)? {
                capture.style.bullet_size = Some(ParagraphBulletSize::Points(
                    value as f32 / 100.0 * POINTS_TO_CSS_PIXELS,
                ));
            }
        }
        "buClr" => capture.bullet_color_depth = Some(depth),
        "srgbClr" => {
            let color = string_attribute(attributes, "val", part)?
                .and_then(|value| parse_rgb_color(&value));
            if !update_paragraph_text_shadow_color(capture, depth, color) {
                if capture.bullet_color_depth.is_some() {
                    capture.style.bullet_color = color;
                } else {
                    capture.style.font_color = color;
                }
            }
        }
        "sysClr" => {
            let color = string_attribute(attributes, "lastClr", part)?
                .and_then(|value| parse_rgb_color(&value));
            if !update_paragraph_text_shadow_color(capture, depth, color) {
                if capture.bullet_color_depth.is_some() {
                    capture.style.bullet_color = color;
                } else {
                    capture.style.font_color = color;
                }
            }
        }
        "schemeClr" => {
            let color =
                string_attribute(attributes, "val", part)?.and_then(|value| theme.color(&value));
            if !update_paragraph_text_shadow_color(capture, depth, color) {
                if capture.bullet_color_depth.is_some() {
                    capture.style.bullet_color = color;
                } else {
                    capture.style.font_color = color;
                }
            }
        }
        "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff"
            if capture.text_effect_depth.is_some_and(|start| depth > start) =>
        {
            let ratio = color_transform_percentage(attributes, part)?;
            if let Some(color) = capture
                .style
                .text_effect
                .as_mut()
                .and_then(|effect| effect.shadow.as_mut())
                .map(|shadow| &mut shadow.color)
            {
                apply_drawingml_color_transform(color, local, ratio);
            }
        }
        "buNone" => capture.style.bullet = Some(ParagraphBullet::None),
        "buChar" => {
            if let Some(value) = string_attribute(attributes, "char", part)? {
                capture.style.bullet = value
                    .chars()
                    .next()
                    .map(|character| ParagraphBullet::Character(character.to_string()));
            }
        }
        "buAutoNum" => {
            let kind = string_attribute(attributes, "type", part)?
                .unwrap_or_else(|| "arabicPeriod".to_owned());
            let start_at = numeric_attribute(attributes, "startAt", part)?
                .map(u32::try_from)
                .transpose()
                .map_err(|_| format_error(part, "automatic-numbering start exceeds u32 range"))?
                .unwrap_or(1);
            capture.style.bullet = Some(ParagraphBullet::AutoNumber { kind, start_at });
        }
        "blip" => {
            if let Some(relationship_id) = string_attribute(attributes, "embed", part)? {
                capture.style.bullet = Some(ParagraphBullet::Image(relationship_id));
            }
        }
        _ => {}
    }
    Ok(())
}

fn update_paragraph_text_shadow_color(
    capture: &mut ParagraphStyleCapture,
    depth: usize,
    color: Option<u32>,
) -> bool {
    if !capture.text_effect_depth.is_some_and(|start| depth > start) {
        return false;
    }
    if let (Some(color), Some(shadow)) = (
        color,
        capture
            .style
            .text_effect
            .as_mut()
            .and_then(|effect| effect.shadow.as_mut()),
    ) {
        shadow.color = color;
    }
    true
}

fn commit_shape_paragraph_style(shape: &mut ShapeState) {
    let Some(capture) = shape.paragraph_style_capture.take() else {
        return;
    };
    match capture.target {
        ParagraphStyleTarget::Default => {
            for style in &mut shape.paragraph_styles {
                style.apply(&capture.style);
            }
        }
        ParagraphStyleTarget::Level(level) => {
            if let Some(style) = shape.paragraph_styles.get_mut(level) {
                style.apply(&capture.style);
            }
        }
        ParagraphStyleTarget::Current => shape.current_paragraph_style.apply(&capture.style),
    }
}

fn apply_end_paragraph_font_size(
    paragraph_pending: bool,
    current_style: &mut ParagraphStyle,
    paragraph_styles: &mut [ParagraphStyle],
    font_size: f32,
) {
    if paragraph_pending {
        current_style.font_size = Some(font_size);
    } else if let Some(style) = paragraph_styles.last_mut() {
        style.font_size = Some(font_size);
    }
}

fn begin_shape_paragraph(shape: &mut ShapeState, emit_prefix: bool) {
    if !shape.paragraph_pending {
        return;
    }
    let level = shape
        .current_paragraph_level
        .min(TEXT_LEVEL_COUNT.saturating_sub(1));
    let mut style = shape
        .paragraph_styles
        .get(level)
        .cloned()
        .unwrap_or_default();
    style.apply(&shape.current_paragraph_style);
    style.resolve_bullet_indentation(&shape.current_paragraph_style);
    shape.inherited_run_effect = style.text_effect.clone().unwrap_or_default();
    shape.font_paint.clone_from(&style.font_paint);
    if let Some(font_family) = &style.font_family {
        shape.font_family.clone_from(font_family);
    }
    if style.font_east_asian.is_some() {
        shape.font_east_asian.clone_from(&style.font_east_asian);
    }
    if style.font_complex_script.is_some() {
        shape
            .font_complex_script
            .clone_from(&style.font_complex_script);
    }
    if let Some(font_size) = style.font_size {
        shape.font_size = font_size;
    }
    if !shape.explicit_text_color
        && let Some(font_color) = style.font_color
    {
        shape.font_color = font_color;
    }
    if let Some(bold) = style.bold {
        shape.bold = bold;
    }
    if let Some(italic) = style.italic {
        shape.italic = italic;
    }
    if let Some(underline) = style.underline {
        shape.underline = underline;
    }
    if let Some(strikethrough) = style.strikethrough {
        shape.strikethrough = strikethrough;
    }
    if let Some(baseline_ratio) = style.baseline_ratio {
        shape.baseline_shift = shape.font_size * baseline_ratio;
    }
    if let Some(letter_spacing) = style.letter_spacing {
        shape.letter_spacing = letter_spacing;
    }
    if let Some(east_asian_line_breaks) = style.east_asian_line_breaks {
        shape.east_asian_line_breaks = east_asian_line_breaks;
    }
    if let Some(capitalization) = style.capitalization {
        shape.capitalization = capitalization;
    }
    if shape.paragraph_layout_style.is_none() {
        shape.paragraph_layout_style = Some(style.clone());
        shape.align = style.align.unwrap_or(shape.align);
    }
    shape.paragraph_layout_styles.push(style.clone());
    shape.paragraph_run_starts.push(shape.runs.len());
    shape.paragraph_layouts.push(TextParagraphLayout {
        align: style.align.unwrap_or(shape.align),
        margin_left: style.margin_left.unwrap_or(0.0).max(0.0),
        margin_right: style.margin_right.unwrap_or(0.0).max(0.0),
        first_line_indent: style.indent.unwrap_or(0.0),
        default_tab_stop: style.default_tab_stop.unwrap_or(36.0).max(1.0),
        line_height: style
            .line_spacing
            .as_ref()
            .map_or(0.0, |spacing| spacing.resolve_for_layout(style.font_size)),
        space_before: style
            .space_before
            .as_ref()
            .map_or(0.0, |spacing| spacing.resolve(shape.font_size)),
        space_after: style
            .space_after
            .as_ref()
            .map_or(0.0, |spacing| spacing.resolve(shape.font_size)),
        // Office defaults from [MS-OE376] §5.1.5.4.13.
        latin_line_break: style.latin_line_break.unwrap_or(false),
        hanging_punctuation: style.hanging_punctuation.unwrap_or(true),
        rule_above: None,
        rule_below: None,
        drop_cap: None,
    });
    let prefix = emit_prefix
        .then(|| match style.bullet.as_ref() {
            Some(ParagraphBullet::Character(value)) => Some(character_bullet_prefix(
                value,
                style.bullet_font_family.as_deref(),
            )),
            Some(ParagraphBullet::AutoNumber { kind, start_at }) => {
                for counter in shape.numbering_counters.iter_mut().skip(level + 1) {
                    *counter = 0;
                }
                let counter = &mut shape.numbering_counters[level];
                if *counter == 0 {
                    *counter = *start_at;
                } else {
                    *counter = counter.saturating_add(1);
                }
                let prefix = format!("{}\t", format_auto_number(kind, *counter));
                Some((prefix.clone(), prefix, None))
            }
            Some(ParagraphBullet::None | ParagraphBullet::Image(_)) | None => None,
        })
        .flatten();
    if let Some((semantic_prefix, glyph_prefix, bullet_font_family)) = prefix {
        shape.text.push_str(&semantic_prefix);
        let mut run = TextRunState::from_shape(shape, shape.depth);
        run.text = glyph_prefix;
        if let Some(font_family) = bullet_font_family {
            run.font_family = font_family;
        }
        if let Some(color) = style.bullet_color {
            run.color = color;
        }
        shape.pending_bullet_style = style.bullet_color.is_none().then_some(shape.runs.len());
        shape.runs.push(run.finish_single_verbatim());
        shape.run_effects.push(shape.inherited_run_effect.clone());
    }
    shape.paragraph_pending = false;
}

fn begin_table_cell_paragraph(cell: &mut TableCellState) {
    if !cell.paragraph_pending {
        return;
    }
    let level = cell
        .current_paragraph_level
        .min(TEXT_LEVEL_COUNT.saturating_sub(1));
    let mut style = cell
        .paragraph_styles
        .get(level)
        .cloned()
        .unwrap_or_default();
    style.apply(&cell.current_paragraph_style);
    style.resolve_bullet_indentation(&cell.current_paragraph_style);

    cell.font_family.clone_from(&cell.base_run.font_family);
    cell.font_east_asian
        .clone_from(&cell.base_run.font_east_asian);
    cell.font_complex_script
        .clone_from(&cell.base_run.font_complex_script);
    cell.font_size = cell.base_run.font_size;
    cell.font_color = cell.base_run.font_color;
    cell.bold = cell.base_run.bold;
    cell.italic = cell.base_run.italic;
    cell.underline = cell.base_run.underline;
    cell.strikethrough = cell.base_run.strikethrough;
    cell.baseline_shift = cell.base_run.baseline_shift;
    cell.letter_spacing = cell.base_run.letter_spacing;
    cell.east_asian_line_breaks = cell.base_run.east_asian_line_breaks;
    cell.capitalization = cell.base_run.capitalization;
    if let Some(font_family) = &style.font_family {
        cell.font_family.clone_from(font_family);
    }
    if style.font_east_asian.is_some() {
        cell.font_east_asian.clone_from(&style.font_east_asian);
    }
    if style.font_complex_script.is_some() {
        cell.font_complex_script
            .clone_from(&style.font_complex_script);
    }
    if let Some(font_size) = style.font_size {
        cell.font_size = font_size;
    }
    if let Some(font_color) = style.font_color {
        cell.font_color = font_color;
    }
    if let Some(bold) = style.bold {
        cell.bold = bold;
    }
    if let Some(italic) = style.italic {
        cell.italic = italic;
    }
    if let Some(underline) = style.underline {
        cell.underline = underline;
    }
    if let Some(strikethrough) = style.strikethrough {
        cell.strikethrough = strikethrough;
    }
    if let Some(baseline_ratio) = style.baseline_ratio {
        cell.baseline_shift = cell.font_size * baseline_ratio;
    }
    if let Some(letter_spacing) = style.letter_spacing {
        cell.letter_spacing = letter_spacing;
    }
    if let Some(east_asian_line_breaks) = style.east_asian_line_breaks {
        cell.east_asian_line_breaks = east_asian_line_breaks;
    }
    if let Some(capitalization) = style.capitalization {
        cell.capitalization = capitalization;
    }
    let align = style.align.unwrap_or(cell.align);
    if cell.paragraph_layouts.is_empty() {
        cell.align = align;
    }
    cell.paragraph_layout_styles.push(style.clone());
    cell.paragraph_run_starts.push(cell.runs.len());
    cell.paragraph_layouts.push(TextParagraphLayout {
        align,
        margin_left: style.margin_left.unwrap_or(0.0).max(0.0),
        margin_right: style.margin_right.unwrap_or(0.0).max(0.0),
        first_line_indent: style.indent.unwrap_or(0.0),
        default_tab_stop: style.default_tab_stop.unwrap_or(36.0).max(1.0),
        line_height: style
            .line_spacing
            .as_ref()
            .map_or(0.0, |spacing| spacing.resolve_for_layout(style.font_size)),
        space_before: style
            .space_before
            .as_ref()
            .map_or(0.0, |spacing| spacing.resolve(cell.font_size)),
        space_after: style
            .space_after
            .as_ref()
            .map_or(0.0, |spacing| spacing.resolve(cell.font_size)),
        latin_line_break: style.latin_line_break.unwrap_or(false),
        hanging_punctuation: style.hanging_punctuation.unwrap_or(true),
        rule_above: None,
        rule_below: None,
        drop_cap: None,
    });
    let prefix = match style.bullet.as_ref() {
        Some(ParagraphBullet::Character(value)) => Some(character_bullet_prefix(
            value,
            style.bullet_font_family.as_deref(),
        )),
        Some(ParagraphBullet::AutoNumber { kind, start_at }) => {
            for counter in cell.numbering_counters.iter_mut().skip(level + 1) {
                *counter = 0;
            }
            let counter = &mut cell.numbering_counters[level];
            if *counter == 0 {
                *counter = *start_at;
            } else {
                *counter = counter.saturating_add(1);
            }
            let prefix = format!("{}\t", format_auto_number(kind, *counter));
            Some((prefix.clone(), prefix, None))
        }
        Some(ParagraphBullet::None | ParagraphBullet::Image(_)) | None => None,
    };
    if let Some((semantic_prefix, glyph_prefix, bullet_font_family)) = prefix {
        cell.text.push_str(&semantic_prefix);
        let mut run = TextRunState::from_table_cell(cell, cell.depth);
        run.text = glyph_prefix;
        if let Some(font_family) = bullet_font_family {
            run.font_family = font_family;
        }
        if let Some(color) = style.bullet_color {
            run.color = color;
        }
        cell.run_styles.push(style.clone());
        cell.runs.push(run.finish_single_verbatim());
        cell.run_effects.push(TextEffect::default());
    }
    cell.paragraph_pending = false;
}

fn character_bullet_prefix(
    value: &str,
    font_family: Option<&str>,
) -> (String, String, Option<String>) {
    let normalized = normalize_symbol_font_character(value, font_family);
    let preserved_font = if normalized == value {
        font_family.map(str::to_owned)
    } else {
        None
    };
    let prefix = format!("{normalized}\t");
    (prefix.clone(), prefix, preserved_font)
}

fn commit_table_cell_paragraph_style(cell: &mut TableCellState) {
    let Some(capture) = cell.paragraph_style_capture.take() else {
        return;
    };
    match capture.target {
        ParagraphStyleTarget::Default => {
            for style in cell
                .paragraph_styles
                .iter_mut()
                .chain(cell.local_paragraph_styles.iter_mut())
            {
                style.apply(&capture.style);
            }
        }
        ParagraphStyleTarget::Level(level) => {
            if let Some(style) = cell.local_paragraph_styles.get_mut(level) {
                style.apply(&capture.style);
            }
            if let Some(style) = cell.paragraph_styles.get_mut(level) {
                style.apply(&capture.style);
            }
        }
        ParagraphStyleTarget::Current => cell.current_paragraph_style.apply(&capture.style),
    }
}

fn resolve_table_cell_paragraph_layouts(cell: &mut TableCellState) {
    let mut layouts = std::mem::take(&mut cell.paragraph_layouts);
    let styles = std::mem::take(&mut cell.paragraph_layout_styles);
    let run_starts = std::mem::take(&mut cell.paragraph_run_starts);
    let mut empty_prefixes = Vec::new();

    for (index, layout) in layouts.iter_mut().enumerate() {
        let style = styles.get(index);
        let start = run_starts
            .get(index)
            .copied()
            .unwrap_or(0)
            .min(cell.runs.len());
        let end = run_starts
            .get(index + 1)
            .copied()
            .unwrap_or(cell.runs.len())
            .min(cell.runs.len());
        let has_generated_prefix = style.and_then(|style| style.bullet.as_ref()).is_some()
            && cell
                .runs
                .get(start)
                .is_some_and(|run| run.text.ends_with('\t'));
        let content_start = start
            .saturating_add(usize::from(has_generated_prefix))
            .min(end);
        let has_content = cell.runs[content_start..end].iter().any(|run| {
            run.text
                .chars()
                .any(|character| !matches!(character, '\r' | '\n' | '\u{2028}'))
        });
        if has_generated_prefix && !has_content {
            empty_prefixes.push(start);
        }
        let font_size = cell.runs[content_start..end]
            .iter()
            .filter(|run| {
                run.text
                    .chars()
                    .any(|character| !matches!(character, '\r' | '\n' | '\u{2028}'))
            })
            .map(|run| run.font_size)
            .reduce(f32::max)
            .or_else(|| style.and_then(|style| style.font_size))
            .unwrap_or(cell.font_size)
            .max(1.0);

        if has_generated_prefix
            && has_content
            && let Some(prefix) = cell.runs.get_mut(start)
        {
            prefix.font_size = style
                .and_then(|style| style.bullet_size.as_ref())
                .map_or(font_size, |size| size.resolve(font_size));
        }
        layout.line_height = style
            .and_then(|style| style.line_spacing.as_ref())
            .map_or(font_size * 1.2, |spacing| {
                spacing.resolve_for_layout(Some(font_size))
            });
        layout.space_before = style
            .and_then(|style| style.space_before.as_ref())
            .map_or(0.0, |spacing| spacing.resolve(font_size));
        layout.space_after = style
            .and_then(|style| style.space_after.as_ref())
            .map_or(0.0, |spacing| spacing.resolve(font_size));
    }
    if !empty_prefixes.is_empty() {
        for index in empty_prefixes.into_iter().rev() {
            cell.run_styles.remove(index);
            cell.runs.remove(index);
            cell.run_effects.remove(index);
        }
        cell.text = cell
            .runs
            .iter()
            .map(|run| run.text.as_str())
            .collect::<String>()
            .replace('\u{2028}', "\n");
    }
    if let Some(first) = layouts.first_mut() {
        first.space_before = 0.0;
    }
    if let Some(last) = layouts.last_mut() {
        last.space_after = 0.0;
    }
    cell.paragraph_layouts = layouts;
}

fn begin_table_cell_run_properties(
    cell: &mut TableCellState,
    attributes: &[XmlAttribute<'_>],
    depth: usize,
    empty: bool,
    part: &str,
) -> Result<(), Diagnostic> {
    cell.run_properties_depth = (!empty).then_some(depth);
    if let Some(value) = boolean_attribute(attributes, "b", part)? {
        cell.run_style.bold = Some(value);
    }
    if let Some(value) = boolean_attribute(attributes, "i", part)? {
        cell.run_style.italic = Some(value);
    }
    if let Some(size) = numeric_attribute(attributes, "sz", part)? {
        let size = size as f32 / 100.0 * POINTS_TO_CSS_PIXELS;
        if let Some(run) = cell.run.as_mut() {
            run.font_size = size;
        } else {
            cell.font_size = size;
        }
    }
    if let Some(value) = boolean_attribute(attributes, "b", part)? {
        if let Some(run) = cell.run.as_mut() {
            run.bold = value;
        } else {
            cell.bold = value;
        }
    }
    if let Some(value) = boolean_attribute(attributes, "i", part)? {
        if let Some(run) = cell.run.as_mut() {
            run.italic = value;
        } else {
            cell.italic = value;
        }
    }
    if let Some(baseline) = signed_numeric_attribute(attributes, "baseline", part)? {
        let baseline_ratio = baseline as f32 / 100_000.0;
        if let Some(run) = cell.run.as_mut() {
            run.baseline_shift = run.font_size * baseline_ratio;
            if baseline != 0 {
                run.font_size *= 0.65;
            }
        } else {
            cell.baseline_shift = cell.font_size * baseline_ratio;
            if baseline != 0 {
                cell.font_size *= 0.65;
            }
        }
    }
    if let Some(spacing) = signed_numeric_attribute(attributes, "spc", part)? {
        let spacing = spacing as f32 / 100.0 * POINTS_TO_CSS_PIXELS;
        if let Some(run) = cell.run.as_mut() {
            run.letter_spacing = spacing;
        } else {
            cell.letter_spacing = spacing;
        }
    }
    if let Some(value) = string_attribute(attributes, "u", part)? {
        let (underline, wavy, dotted, heavy, double, dot_dash) = drawingml_underline(&value);
        if let Some(run) = cell.run.as_mut() {
            run.underline = underline;
            run.wavy_underline = wavy;
            run.dotted_underline = dotted;
            run.heavy_underline = heavy;
            run.double_underline = double;
            run.dot_dash_underline = dot_dash;
        } else {
            cell.underline = underline;
        }
    }
    if let Some(value) = string_attribute(attributes, "strike", part)? {
        let double = value == "dblStrike";
        let value = !matches!(value.as_str(), "noStrike" | "none" | "0" | "false");
        if let Some(run) = cell.run.as_mut() {
            run.strikethrough = value;
            run.double_strikethrough = double;
        } else {
            cell.strikethrough = value;
        }
    }
    if let Some(value) = string_attribute(attributes, "cap", part)? {
        let capitalization = match value.as_str() {
            "all" => TextCapitalization::All,
            "small" => TextCapitalization::Small,
            _ => TextCapitalization::None,
        };
        if let Some(run) = cell.run.as_mut() {
            run.capitalization = capitalization;
        } else {
            cell.capitalization = capitalization;
        }
    }
    Ok(())
}

// Rich runs distinguish manual line separators (U+2028) from paragraph breaks (LF).
// Plain object text keeps LF for search and copy; both occupy one UTF-16 code unit.
fn append_table_cell_line_break(cell: &mut TableCellState, depth: usize, separator: char) {
    cell.text.push('\n');
    let line_break = cell.runs.last().cloned().map_or_else(
        || {
            let mut run = TextRunState::from_table_cell(cell, depth);
            run.text.push(separator);
            run.finish_single()
        },
        |mut run| {
            run.text.clear();
            run.text.push(separator);
            run
        },
    );
    cell.run_styles
        .push(cell.run_styles.last().cloned().unwrap_or_default());
    cell.runs.push(line_break);
    cell.run_effects
        .push(cell.run_effects.last().cloned().unwrap_or_default());
}

fn append_shape_line_break(shape: &mut ShapeState, depth: usize, separator: char) {
    shape.text.push('\n');
    let line_break = shape.runs.last().cloned().map_or_else(
        || {
            let mut run = TextRunState::from_shape(shape, depth);
            run.text.push(separator);
            run.finish_single()
        },
        |mut run| {
            run.text.clear();
            run.text.push(separator);
            run
        },
    );
    shape.runs.push(line_break);
    shape.run_effects.push(shape.inherited_run_effect.clone());
}

fn format_auto_number(kind: &str, value: u32) -> String {
    match kind {
        "arabicParenR" => format!("{value})"),
        "arabicParenBoth" => format!("({value})"),
        "alphaLcParenR" => format!("{})", super::alphabetic_number(value, false)),
        "alphaLcPeriod" => format!("{}.", super::alphabetic_number(value, false)),
        "alphaUcParenR" => format!("{})", super::alphabetic_number(value, true)),
        "alphaUcPeriod" => format!("{}.", super::alphabetic_number(value, true)),
        "romanLcParenR" => format!("{})", super::roman_number(value, false)),
        "romanLcParenBoth" => format!("({})", super::roman_number(value, false)),
        "romanLcPeriod" => format!("{}.", super::roman_number(value, false)),
        "romanUcParenR" => format!("{})", super::roman_number(value, true)),
        "romanUcParenBoth" => format!("({})", super::roman_number(value, true)),
        "romanUcPeriod" => format!("{}.", super::roman_number(value, true)),
        "circleNumDbPlain" => super::circled_number(value).unwrap_or_else(|| format!("{value}.")),
        "ea1JpnKorPeriod" => format!("{}.", super::chinese_number(value)),
        "ea1JpnChsDbPeriod" => format!("{}、", super::chinese_number(value)),
        _ => format!("{value}."),
    }
}

fn append_shape_text(shape: &mut ShapeState, text: &str) {
    if let Some(run) = shape.run.as_mut()
        && let Some(replacement) = run.replacement_text.as_deref()
    {
        if run.text.is_empty() {
            shape.text.push_str(replacement);
            run.text.push_str(replacement);
        }
        return;
    }
    if let Some(run) = shape.run.as_mut() {
        let text = drawingml_capitalized_text(text, run.capitalization);
        shape.text.push_str(&normalized_symbol_run_text(&text, run));
        run.text.push_str(&text);
    } else {
        shape
            .text
            .push_str(&drawingml_capitalized_text(text, shape.capitalization));
    }
}

fn append_table_cell_text(cell: &mut TableCellState, text: &str) {
    if let Some(run) = cell.run.as_mut() {
        let text = drawingml_capitalized_text(text, run.capitalization);
        cell.text.push_str(&normalized_symbol_run_text(&text, run));
        run.text.push_str(&text);
    } else {
        cell.text.push_str(text);
    }
}

fn drawingml_capitalized_text(text: &str, capitalization: TextCapitalization) -> String {
    match capitalization {
        TextCapitalization::None => text.to_owned(),
        TextCapitalization::All | TextCapitalization::Small => text.to_uppercase(),
    }
}

fn normalized_symbol_run_text(text: &str, run: &TextRunState) -> String {
    let mut normalized = String::with_capacity(text.len());
    for character in text.chars() {
        normalized.push_str(&normalize_symbol_font_character(
            &character.to_string(),
            Some(run.family_for(character)),
        ));
    }
    normalized
}

fn apply_master_text_styles(shape: &mut ShapeState, master_text_styles: Option<&MasterTextStyles>) {
    let Some(styles) = master_text_styles else {
        return;
    };
    shape.paragraph_styles = styles
        .for_placeholder(shape.placeholder_type.as_deref())
        .to_vec();
}

fn presentation_part_properties(
    package: &Package<'_>,
    part: &str,
) -> Result<PresentationPartProperties, Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut properties = PresentationPartProperties {
        has_background: false,
        show_master_shapes: true,
        shown: true,
        name: None,
        color_map: None,
        has_transition: false,
        has_timing: false,
        unsupported_features: HashSet::new(),
    };
    let mut found_root = false;
    parse_xml(&bytes, package.limits(), |event| {
        let XmlEvent::StartElement {
            name, attributes, ..
        } = event
        else {
            return Ok(());
        };
        let local = local_name(name);
        if !found_root && matches!(local, "sld" | "sldLayout" | "sldMaster") {
            properties.show_master_shapes =
                boolean_attribute(&attributes, "showMasterSp", part)?.unwrap_or(true);
            properties.shown = boolean_attribute(&attributes, "show", part)?.unwrap_or(true);
            found_root = true;
        } else if local == "cSld" && properties.name.is_none() {
            properties.name =
                string_attribute(&attributes, "name", part)?.filter(|value| !value.is_empty());
        } else if local == "bg" {
            properties.has_background = true;
        } else if matches!(local, "clrMap" | "overrideClrMapping") {
            let mut mapping = default_color_map();
            for key in [
                "bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5",
                "accent6", "hlink", "folHlink",
            ] {
                if let Some(value) = string_attribute(&attributes, key, part)? {
                    mapping.insert(key.to_owned(), value);
                }
            }
            properties.color_map = Some(mapping);
        } else if local == "transition" {
            properties.has_transition = true;
        } else if local == "timing" {
            properties.has_timing = true;
        } else if matches!(local, "audio" | "video") {
            properties
                .unsupported_features
                .insert("timed-media".to_owned());
        } else if local == "control" {
            properties
                .unsupported_features
                .insert("activex-control".to_owned());
        } else if local == "model3d" {
            properties
                .unsupported_features
                .insert("three-dimensional-model".to_owned());
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(properties)
}

#[allow(clippy::too_many_arguments)]
fn parse_presentation_part(
    package: &Package<'_>,
    part: &str,
    unit_index: u32,
    slide_bounds: Rect,
    content: PartContent,
    state: &mut PptxParseState,
    content_types: &ContentTypes,
    font_metrics: &FontMetricTable,
    theme: &PptxTheme,
    master_text_styles: Option<&MasterTextStyles>,
    diagram_text_colors: Option<&HashMap<String, u32>>,
) -> Result<(), Diagnostic> {
    let bytes = package.required_part(part)?;
    let relationships = package.relationships(Some(part))?;
    let relationship_map: HashMap<&str, &Relationship> = relationships
        .iter()
        .map(|relationship| (relationship.id.as_str(), relationship))
        .collect();
    let mut depth = 0_usize;
    let mut word_text_depth = None;
    let mut shape: Option<ShapeState> = None;
    let mut picture: Option<PictureState> = None;
    let mut graphic_frame: Option<GraphicFrameState> = None;
    let mut groups: Vec<GroupState> = Vec::new();
    let mut unsupported_formula_reported = false;
    let mut unsupported_geometry_reported = false;
    let mut wordart_warp_reported = false;
    let mut chart_choice = super::drawingml::DrawingMlChartChoice::default();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if chart_choice.start(local, depth, empty) {
                    depth += usize::from(!empty); return Ok(());
                }
                if local == "txbxContent" || word_text_depth.is_some() {
                    if !empty {
                        if word_text_depth.is_none() {
                            word_text_depth = Some(depth);
                        }
                        depth = depth.saturating_add(1);
                    }
                    return Ok(());
                }
                if (content.pictures || content.shapes)
                    && matches!(local, "grpSp" | "lockedCanvas" | "wgp")
                {
                    if state.objects.len() >= package.limits().max_document_objects {
                        return Err(Diagnostic::fatal(
                            DiagnosticCode::ObjectLimit,
                            Phase::Parse,
                            None,
                            "document exceeds the configured object limit",
                        )
                        .in_part(part));
                    }
                    let numeric_id = u32::try_from(state.objects.len())
                        .map_err(|_| format_error(part, "object count exceeds supported range"))?;
                    let parent_numeric_id = groups.last().map(|group| group.numeric_id);
                    let z = state.take_z();
                    state.objects.push(Object {
                        numeric_id,
                        parent_numeric_id,
                        stable_id: format!("object:{numeric_id}"),
                        parent_stable_id: parent_numeric_id
                            .map(|parent| format!("object:{parent}")),
                        kind: ObjectKind::Group,
                        unit_index,
                        bounds: Rect::default(),
                        z,
                        text: None,
                        source: SourceRef {
                            part: part.to_owned(),
                            mapping: MappingQuality::Derived,
                            locator: SourceLocator::PptxShape {
                                shape_id: numeric_id,
                                row: None,
                                column: None,
                                text_range: None,
                                metadata: PptxObjectMetadata::default(),
                            },
                        },
                        visual: Visual::None,
                    });
                    groups.push(GroupState {
                        depth,
                        numeric_id,
                        shape_id: None,
                        metadata: PptxObjectMetadata::default(),
                        non_visual_depth: None,
                        shape_properties_depth: None,
                        fill: Paint::None,
                        explicit_fill: false,
                        paint_capture: None,
                        transform_depth: None,
                        x: None,
                        y: None,
                        width: None,
                        height: None,
                        child_x: None,
                        child_y: None,
                        child_width: None,
                        child_height: None,
                        rotation_degrees: 0.0,
                        flip_horizontal: false,
                        flip_vertical: false,
                        three_d: None,
                        scene_3d_depth: None,
                        backdrop_3d_depth: None,
                        camera_3d_depth: None,
                        light_3d_depth: None,
                    });
                }
                if content.shapes && local == "graphicFrame" {
                    if graphic_frame.is_some() {
                        return Err(format_error(part, "nested graphic frames are invalid"));
                    }
                    if empty {
                        return Err(format_error(part, "graphic frame element has no content"));
                    }
                    graphic_frame = Some(GraphicFrameState {
                        depth,
                        parent_numeric_id: groups.last().map(|group| group.numeric_id),
                        shape_id: None,
                        metadata: PptxObjectMetadata::default(),
                        non_visual_depth: None,
                        transform_depth: None,
                        x: None,
                        y: None,
                        width: None,
                        height: None,
                        table_depth: None,
                        table_style_id: String::new(),
                        table_fill: None,
                        table_effects: None,
                        collecting_table_style_id: false,
                        first_row: false,
                        last_row: false,
                        first_column: false,
                        last_column: false,
                        band_rows: false,
                        band_columns: false,
                        columns: Vec::new(),
                        rows: Vec::new(),
                        row: None,
                        cell: None,
                        chart: None,
                        diagram: None,
                        embedded_raster: None,
                        embedded_text: None,
                        embedded_plain_text: None,
                        embedded_document: None,
                        placeholder: None,
                    });
                }
                if content.shapes && matches!(local, "sp" | "cxnSp" | "wsp") && shape.is_none() {
                    let mut current = ShapeState::new(depth, theme, local == "cxnSp");
                    current.diagram_text_color = string_attribute(&attributes, "modelId", part)?
                        .and_then(|model_id| {
                            diagram_text_colors.and_then(|colors| colors.get(&model_id).copied())
                        });
                    current
                        .paragraph_styles
                        .clone_from(&state.default_text_styles);
                    shape = Some(current);
                }
                if content.pictures && local == "pic" {
                    if picture.is_some() {
                        return Err(format_error(part, "nested pictures are invalid"));
                    }
                    if empty {
                        return Err(format_error(part, "picture element has no content"));
                    }
                    picture = Some(PictureState::new(
                        depth,
                        groups.last().map(|group| group.numeric_id),
                    ));
                } else if content.backgrounds && local == "bg" && picture.is_none() && !empty {
                    picture = Some(PictureState::background(depth, slide_bounds));
                }
                if let Some(current) = shape.as_mut() {
                    if local == "cNvPr" && current.non_visual_depth == Some(depth) {
                        current.non_visual_depth = None;
                    }
                    if matches!(local, "oMath" | "oMathPara")
                        && current.math_capture.is_none()
                    {
                        current.math_present = true;
                        if empty {
                            current.omit_empty_math = true;
                        } else {
                            current.math_capture = Some(OmmlCapture::new(depth, local));
                        }
                    } else if let Some(math) = current.math_capture.as_mut() {
                        math.start(local, &attributes, depth, empty, part)?;
                    }
                    if local == "custGeom" {
                        current.preset = None;
                        current.preset_adjustments.clear();
                        // Without path w/h, DrawingML coordinates use the shape's EMU extents.
                        current.custom_geometry = (!empty).then(|| {
                            CustomGeometryState::new(
                                depth,
                                current.width.unwrap_or(0) as f32,
                                current.height.unwrap_or(0) as f32,
                            )
                        });
                    } else if let Some(custom_geometry) = current.custom_geometry.as_mut() {
                        custom_geometry.start(local, &attributes, part, depth, empty)?;
                    }
                    if let Some(image_fill) = current.image_fill.as_mut() {
                        image_fill.start(local, &attributes, part)?;
                    }
                    if current.text_body_properties_depth.is_some()
                        && let Some((_, capture)) = current.paint_capture.as_mut()
                    {
                        capture.start(local, &attributes, part, theme)?;
                    }
                    if current.shape_properties_depth.is_some() {
                        if let Some((_, capture)) = current.paint_capture.as_mut() {
                            capture.start(local, &attributes, part, theme)?;
                        } else {
                            let direct_target = if current
                                .line_depth
                                .is_some_and(|start| depth == start.saturating_add(1))
                            {
                                Some(PaintTarget::Stroke)
                            } else if current
                                .shape_properties_depth
                                .is_some_and(|start| depth == start.saturating_add(1))
                            {
                                Some(PaintTarget::Fill)
                            } else {
                                None
                            };
                            if matches!(local, "solidFill" | "gradFill" | "pattFill") {
                                if let Some(target) = direct_target {
                                    match target {
                                        PaintTarget::Fill => current.explicit_fill = true,
                                        PaintTarget::Stroke => current.explicit_stroke = true,
                                        PaintTarget::Text
                                        | PaintTarget::Highlight
                                        | PaintTarget::Extrusion
                                        | PaintTarget::Contour => {}
                                    }
                                    current.paint_capture =
                                        PaintCapture::new(local, &attributes, part, depth)?
                                        .map(|capture| (target, capture));
                                }
                            } else if local == "blipFill"
                                && matches!(direct_target, Some(PaintTarget::Fill))
                            {
                                current.explicit_fill = true;
                                current.fill = Paint::None;
                                current.image_fill = (!empty).then(|| ShapeImageFillState::new(depth));
                                if let Some(image_fill) = current.image_fill.as_mut() {
                                    super::drawingml::drawingml_image_fill_mapping(&mut image_fill.mapping, local, &attributes, part)?;
                                }
                            } else if matches!(local, "noFill" | "grpFill")
                                && let Some(target) = direct_target
                            {
                                match target {
                                    PaintTarget::Fill => {
                                        current.fill = Paint::None;
                                        current.explicit_fill = true;
                                        current.use_group_fill = local == "grpFill";
                                    }
                                    PaintTarget::Stroke => {
                                        current.stroke = Paint::None;
                                        current.explicit_stroke = true;
                                    }
                                    PaintTarget::Text
                                    | PaintTarget::Highlight
                                    | PaintTarget::Extrusion
                                    | PaintTarget::Contour => {}
                                }
                            }
                        }
                    }
                    if current.run_properties_depth.is_some() {
                        if local == "noFill" && current.line_depth.is_some() {
                            if let Some(run) = current.run.as_mut() { run.stroke = Some(Box::new(Paint::None)); }
                        } else if local == "noFill" {
                            if let Some(run) = current.run.as_mut() { run.paint = Some(Box::new(Paint::None)); }
                            else { current.font_paint = Some(Box::new(Paint::None)); }
                        }
                        if local == "blipFill" {
                            let mut image = ShapeImageFillState::new(depth);
                            image.target = PaintTarget::Text;
                            super::drawingml::drawingml_image_fill_mapping(&mut image.mapping, local, &attributes, part)?;
                            current.image_fill = Some(image);
                        }
                        if let Some((_, capture)) = current.paint_capture.as_mut() {
                            capture.start(local, &attributes, part, theme)?;
                        } else if local == "highlight" {
                            current.paint_capture =
                                PaintCapture::new("solidFill", &attributes, part, depth)?
                                .map(|capture| (PaintTarget::Highlight, capture));
                        } else if matches!(local, "solidFill" | "gradFill" | "pattFill") {
                            current.paint_capture =
                                PaintCapture::new(local, &attributes, part, depth)?
                                .map(|capture| (if current.line_depth.is_some() { PaintTarget::Stroke } else { PaintTarget::Text }, capture));
                        }
                    }
                    match local {
                        "cNvPr" if current.shape_id.is_none() => {
                            current.shape_id = numeric_attribute(&attributes, "id", part)?
                                .and_then(|value| u32::try_from(value).ok());
                            current.metadata = non_visual_metadata(&attributes, part)?;
                            current.non_visual_depth = (!empty).then_some(depth);
                        }
                        "hlinkClick" if current.non_visual_depth.is_some() => {
                            current.metadata.click_action =
                                Some(drawingml_action(&attributes, &relationship_map, part)?);
                        }
                        "hlinkHover" if current.non_visual_depth.is_some() => {
                            current.metadata.hover_action =
                                Some(drawingml_action(&attributes, &relationship_map, part)?);
                        }
                        "ph" => {
                            current.placeholder_type = Some(
                                string_attribute(&attributes, "type", part)?
                                    .unwrap_or_else(|| "obj".to_owned()),
                            );
                            current.placeholder_index = numeric_attribute(&attributes, "idx", part)?
                                .map(u32::try_from)
                                .transpose()
                                .map_err(|_| {
                                    format_error(part, "placeholder index exceeds u32 range")
                                })?;
                            current.is_custom_prompt = boolean_attribute(
                                &attributes,
                                "hasCustomPrompt",
                                part,
                            )?
                            .unwrap_or(false);
                            apply_master_text_styles(current, master_text_styles);
                            inherit_placeholder_style(current, unit_index, state);
                        }
                        "spPr" if depth == current.depth.saturating_add(1) => {
                            current.shape_properties_depth = (!empty).then_some(depth);
                        }
                        "scene3d"
                            if current.shape_properties_depth.is_some()
                                || current.text_body_properties_depth.is_some() =>
                        {
                            current
                                .three_d
                                .get_or_insert_with(ThreeDStyle::default)
                                .applies_to_text = current.text_body_properties_depth.is_some();
                            current.scene_3d_depth = (!empty).then_some(depth);
                        }
                        "camera" if current.scene_3d_depth.is_some() => {
                            parse_three_d_camera(
                                current.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                            )?;
                            current.camera_3d_depth = (!empty).then_some(depth);
                        }
                        "lightRig" if current.scene_3d_depth.is_some() => {
                            parse_three_d_light(
                                current.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                            )?;
                            current.light_3d_depth = (!empty).then_some(depth);
                        }
                        "backdrop" if current.scene_3d_depth.is_some() => {
                            current
                                .three_d
                                .get_or_insert_with(ThreeDStyle::default)
                                .backdrop = Some(Backdrop3D::default());
                            current.backdrop_3d_depth = (!empty).then_some(depth);
                        }
                        "anchor" | "norm" | "up" if current.backdrop_3d_depth.is_some() => {
                            let backdrop = current
                                .three_d
                                .get_or_insert_with(ThreeDStyle::default)
                                .backdrop
                                .get_or_insert_with(Backdrop3D::default);
                            parse_three_d_backdrop_point(backdrop, local, &attributes, part)?;
                        }
                        "rot" if current.camera_3d_depth.is_some() => {
                            parse_three_d_rotation(
                                current.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                                true,
                            )?;
                        }
                        "rot" if current.light_3d_depth.is_some() => {
                            parse_three_d_rotation(
                                current.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                                false,
                            )?;
                        }
                        "sp3d"
                            if current.shape_properties_depth.is_some()
                                || current.text_body_properties_depth.is_some() =>
                        {
                            parse_three_d_shape(
                                current.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                                current.text_body_properties_depth.is_some(),
                            )?;
                            current.shape_3d_depth = (!empty).then_some(depth);
                        }
                        "bevelT" if current.shape_3d_depth.is_some() => {
                            current
                                .three_d
                                .get_or_insert_with(ThreeDStyle::default)
                                .bevel_top = Some(parse_three_d_bevel(&attributes, part)?);
                        }
                        "bevelB" if current.shape_3d_depth.is_some() => {
                            current
                                .three_d
                                .get_or_insert_with(ThreeDStyle::default)
                                .bevel_bottom = Some(parse_three_d_bevel(&attributes, part)?);
                        }
                        "extrusionClr" if current.shape_3d_depth.is_some() && !empty => {
                            current.paint_capture = PaintCapture::new(
                                "solidFill",
                                &[],
                                part,
                                depth,
                            )?
                            .map(|capture| (PaintTarget::Extrusion, capture));
                        }
                        "contourClr" if current.shape_3d_depth.is_some() && !empty => {
                            current.paint_capture = PaintCapture::new(
                                "solidFill",
                                &[],
                                part,
                                depth,
                            )?
                            .map(|capture| (PaintTarget::Contour, capture));
                        }
                        "txXfrm" => {
                            current.text_transform_depth = (!empty).then_some(depth);
                        }
                        "off"
                            if current
                                .text_transform_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            current.text_x = signed_numeric_attribute(&attributes, "x", part)?;
                            current.text_y = signed_numeric_attribute(&attributes, "y", part)?;
                        }
                        "ext"
                            if current
                                .text_transform_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            current.text_width = numeric_attribute(&attributes, "cx", part)?;
                            current.text_height = numeric_attribute(&attributes, "cy", part)?;
                        }
                        "xfrm"
                            if current
                                .shape_properties_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            current.transform_depth = (!empty).then_some(depth);
                            current.rotation_degrees = signed_numeric_attribute(
                                &attributes,
                                "rot",
                                part,
                            )?
                            .unwrap_or(0) as f32
                                / 60_000.0;
                            current.flip_horizontal = boolean_attribute(
                                &attributes,
                                "flipH",
                                part,
                            )?
                            .unwrap_or(false);
                            current.flip_vertical = boolean_attribute(
                                &attributes,
                                "flipV",
                                part,
                            )?
                            .unwrap_or(false);
                        }
                        "off"
                            if current
                                .transform_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            current.x = signed_numeric_attribute(&attributes, "x", part)?;
                            current.y = signed_numeric_attribute(&attributes, "y", part)?;
                        }
                        "ext"
                            if current
                                .transform_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            current.width = numeric_attribute(&attributes, "cx", part)?;
                            current.height = numeric_attribute(&attributes, "cy", part)?;
                        }
                        "ln" if current.run_properties_depth.is_some() => {
                            current.line_depth = (!empty).then_some(depth);
                            let (width, _) = drawingml_stroke_style(&attributes, part)?;
                            if let Some(run) = current.run.as_mut() { run.stroke_width = width; }
                        }
                        "ln" if current.shape_properties_depth.is_some() => {
                            current.line_depth = (!empty).then_some(depth);
                            let width = numeric_attribute(&attributes, "w", part)?;
                            current.explicit_stroke_width = width.is_some();
                            let (stroke_width, style) =
                                drawingml_stroke_style(&attributes, part)?;
                            current.stroke_width = stroke_width;
                            current.line_cap = style.cap;
                            current.line_compound = style.compound;
                            current.line_alignment = style.alignment;
                        }
                        "round" if current.line_depth.is_some() => {
                            current.line_join = LineJoin::Round;
                        }
                        "bevel" if current.line_depth.is_some() => {
                            current.line_join = LineJoin::Bevel;
                        }
                        "miter" if current.line_depth.is_some() => {
                            current.line_join = LineJoin::Miter;
                            current.miter_limit = optional_percentage_attribute(
                                &attributes,
                                "lim",
                                part,
                            )?
                            .unwrap_or(8.0);
                        }
                        "prstGeom" => {
                            let preset = string_attribute(&attributes, "prst", part)?;
                            current.preset = preset.clone();
                            current.preset_adjustments.clear();
                            current.geometry = if let Some(geometry) = preset
                                .as_deref()
                                .and_then(|preset| {
                                    preset_geometry(preset, shape_bounds(current))
                                        .map(|(geometry, _)| geometry)
                                })
                            {
                                geometry
                            } else {
                                match preset.as_deref() {
                                    Some("rect") | None => Geometry::Rectangle,
                                    Some(_) => Geometry::Rectangle,
                                }
                            };
                        }
                        "gd" if current.preset.is_some() => {
                            if let (Some(name), Some(formula)) = (
                                string_attribute(&attributes, "name", part)?,
                                string_attribute(&attributes, "fmla", part)?,
                            ) && let Some(value) = formula
                                .strip_prefix("val ")
                                .and_then(|value| value.parse::<f32>().ok())
                                .filter(|value| value.is_finite())
                            {
                                current.preset_adjustments.insert(name, value);
                            }
                        }
                        "headEnd" if current.line_depth.is_some() => {
                            current.head_arrow = drawingml_line_end(&attributes, part)?;
                        }
                        "tailEnd" if current.line_depth.is_some() => {
                            current.tail_arrow = drawingml_line_end(&attributes, part)?;
                        }
                        "prstDash" if current.line_depth.is_some() => {
                            let value = string_attribute(&attributes, "val", part)?;
                            current.dash_pattern =
                                DrawingMlDashPattern::from_attribute(value.as_deref());
                        }
                        "bodyPr" => {
                            current.text_body_properties_depth = (!empty).then_some(depth);
                            current.text_space_first_last_paragraph = boolean_attribute(
                                &attributes,
                                "spcFirstLastPara",
                                part,
                            )?
                            .unwrap_or(false);
                            if boolean_attribute(&attributes, "rtlCol", part)?.unwrap_or(false) {
                                current.text_direction = TextDirection::Rtl;
                            }
                            if boolean_attribute(&attributes, "anchorCtr", part)?.unwrap_or(false) {
                                current.align = TextAlign::Center;
                            }
                            current.text_column_count =
                                numeric_attribute(&attributes, "numCol", part)?
                                    .and_then(|value| u32::try_from(value).ok())
                                    .unwrap_or(1)
                                    .clamp(1, 64);
                            current.text_column_spacing =
                                numeric_attribute(&attributes, "spcCol", part)?
                                    .map_or(0.0, |value| value as f32 / EMU_PER_CSS_PIXEL);
                            current.text_rotation_degrees =
                                signed_numeric_attribute(&attributes, "rot", part)?
                                    .unwrap_or(0) as f32
                                    / 60_000.0;
                            if boolean_attribute(&attributes, "upright", part)?.unwrap_or(false) {
                                current.text_rotation_degrees = 0.0;
                            }
                            current.text_horizontal_overflow =
                                match string_attribute(&attributes, "horzOverflow", part)?
                                    .as_deref()
                                {
                                    Some("clip") => TextHorizontalOverflow::Clip,
                                    _ => TextHorizontalOverflow::Overflow,
                                };
                            current.text_vertical_overflow =
                                match string_attribute(&attributes, "vertOverflow", part)?
                                    .as_deref()
                                {
                                    Some("clip") => TextVerticalOverflow::Clip,
                                    Some("ellipsis") => TextVerticalOverflow::Ellipsis,
                                    _ => TextVerticalOverflow::Overflow,
                                };
                            current.text_wrap = string_attribute(&attributes, "wrap", part)?
                                .as_deref()
                                != Some("none");
                            current.text_inset_left = signed_numeric_attribute(&attributes, "lIns", part)?
                                .map_or(9.6, |value| value as f32 / EMU_PER_CSS_PIXEL);
                            current.text_inset_right = signed_numeric_attribute(&attributes, "rIns", part)?
                                .map_or(9.6, |value| value as f32 / EMU_PER_CSS_PIXEL);
                            current.text_inset_top = signed_numeric_attribute(&attributes, "tIns", part)?
                                .map_or(4.8, |value| value as f32 / EMU_PER_CSS_PIXEL);
                            current.text_inset_bottom = signed_numeric_attribute(&attributes, "bIns", part)?
                                .map_or(4.8, |value| value as f32 / EMU_PER_CSS_PIXEL);
                            if let Some(anchor) = string_attribute(&attributes, "anchor", part)? {
                                current.vertical_align = Some(match anchor.as_str() {
                                    "ctr" => TextVerticalAlign::Center,
                                    "b" => TextVerticalAlign::Bottom,
                                    _ => TextVerticalAlign::Top,
                                });
                            }
                            if let Some(vertical) = string_attribute(&attributes, "vert", part)? {
                                current.text_orientation =
                                    drawingml_text_orientation(vertical.as_str());
                            }
                            let rotation = current.text_rotation_degrees.rem_euclid(360.0);
                            if current.text_orientation == TextOrientation::Horizontal
                                && (rotation - 90.0).abs() < 0.01
                            {
                                current.text_orientation = TextOrientation::Rotated90;
                                current.text_rotation_degrees = 0.0;
                            } else if current.text_orientation == TextOrientation::Horizontal
                                && (rotation - 270.0).abs() < 0.01
                            {
                                current.text_orientation = TextOrientation::Rotated270;
                                current.text_rotation_degrees = 0.0;
                            }
                        }
                        "spAutoFit" => {
                            // Static rendering keeps the persisted a:xfrm extent. A shape resize
                            // requires final font metrics and connector relayout; estimating it
                            // while parsing distorts authored geometry. bodyPr overflow still
                            // controls how text outside that extent is painted.
                        }
                        "flatTx" => {
                            current
                                .three_d
                                .get_or_insert_with(ThreeDStyle::default)
                                .flat_text_z = Some(
                                signed_numeric_attribute(&attributes, "z", part)?.unwrap_or(0)
                                    as f32
                                    / EMU_PER_CSS_PIXEL,
                            );
                        }
                        "normAutofit" | "normAutoFit" => {
                            // The file stores the final fontScale (default 100%). Re-fitting
                            // with browser metrics would apply an additional, unrequested shrink.
                            current.text_auto_fit = TextAutoFit::None;
                            current.text_font_scale =
                                numeric_attribute(&attributes, "fontScale", part)?
                                    .map_or(1.0, |value| value as f32 / 100_000.0)
                                    .clamp(0.01, 1.0);
                            current.text_line_spacing_reduction =
                                numeric_attribute(&attributes, "lnSpcReduction", part)?
                                    .map_or(0.0, |value| value as f32 / 100_000.0)
                                    .clamp(0.0, 0.99);
                        }
                        "noAutofit" => {
                            current.text_auto_fit = TextAutoFit::None;
                            current.text_font_scale = 1.0;
                            current.text_line_spacing_reduction = 0.0;
                        }
                        "prstTxWarp" => {
                            current.text_warp = string_attribute(&attributes, "prst", part)?
                                .filter(|value| !value.is_empty() && value != "textNoShape");
                            if current.text_warp.is_some() && !wordart_warp_reported {
                                state.diagnostics.push(
                                    Diagnostic::warning(
                                        DiagnosticCode::ApproximateLayout,
                                        Phase::Render,
                                        Fidelity::Approximate,
                                        "PPTX WordArt text warp is rendered with a deterministic Canvas approximation",
                                    )
                                    .in_part(part)
                                    .with_detail("feature", "wordart-warp"),
                                );
                                wordart_warp_reported = true;
                            }
                        }
                        "outerShdw" | "prstShdw" | "innerShdw"
                            if current.shape_properties_depth.is_some()
                                || current.run_properties_depth.is_some()
                                    && current.run.is_some() =>
                        {
                            current.shadow_kind = match local {
                                "innerShdw" => PptxShadowKind::Inner,
                                "glow" => PptxShadowKind::Glow,
                                _ => PptxShadowKind::Outer,
                            };
                            current.shadow_depth = (!empty).then_some(depth);
                            current.shadow_color = None;
                            current.shadow_blur = numeric_attribute(
                                &attributes,
                                if local == "glow" { "rad" } else { "blurRad" },
                                part,
                            )?
                            .unwrap_or(0) as f32
                                / EMU_PER_CSS_PIXEL;
                            current.shadow_distance = numeric_attribute(
                                &attributes,
                                "dist",
                                part,
                            )?
                            .unwrap_or(0) as f32
                                / EMU_PER_CSS_PIXEL;
                            current.shadow_direction_degrees = signed_numeric_attribute(
                                &attributes,
                                "dir",
                                part,
                            )?
                            .unwrap_or(0) as f32
                                / 60_000.0;
                            if matches!(current.shadow_kind, PptxShadowKind::Outer)
                                && current.shape_properties_depth.is_some()
                            {
                                let radians = current.shadow_direction_degrees.to_radians();
                                current.outer_shadow = Some(drawingml_outer_shadow(
                                    Shadow {
                                        color: 0x0000_0080,
                                        blur: current.shadow_blur,
                                        offset_x: current.shadow_distance * radians.cos(),
                                        offset_y: current.shadow_distance * radians.sin(),
                                    },
                                    &attributes,
                                    part,
                                )?);
                            }
                            if matches!(current.shadow_kind, PptxShadowKind::Outer)
                                && let Some(run) = current.run.as_mut()
                                && current.run_properties_depth.is_some()
                            {
                                run.shadow_scale_x = signed_numeric_attribute(
                                    &attributes, "sx", part,
                                )?.map_or(1.0, |value| value as f32 / 100_000.0);
                                run.shadow_scale_y = signed_numeric_attribute(
                                    &attributes, "sy", part,
                                )?.map_or(1.0, |value| value as f32 / 100_000.0);
                                run.shadow_skew_x = signed_numeric_attribute(
                                    &attributes, "kx", part,
                                )?.map_or(0.0, |value| value as f32 / 60_000.0);
                                run.shadow_skew_y = signed_numeric_attribute(
                                    &attributes, "ky", part,
                                )?.map_or(0.0, |value| value as f32 / 60_000.0);
                                run.shadow_alignment = drawingml_shadow_alignment(
                                    string_attribute(&attributes, "algn", part)?.as_deref(),
                                );
                            }
                            if empty {
                                let radians = current.shadow_direction_degrees.to_radians();
                                let shadow = Shadow {
                                    color: 0x0000_0080,
                                    blur: current.shadow_blur,
                                    offset_x: current.shadow_distance * radians.cos(),
                                    offset_y: current.shadow_distance * radians.sin(),
                                };
                                match current.shadow_kind {
                                    PptxShadowKind::Outer => {
                                        if let Some(run) = current.run.as_mut()
                                            && current.run_properties_depth.is_some()
                                        {
                                            run.shadow = Some(shadow);
                                        } else if let Some(effect) = current.outer_shadow.as_mut() {
                                            effect.shadow = shadow;
                                        } else {
                                            current.shadow = Some(shadow);
                                        }
                                    }
                                    PptxShadowKind::Inner => {
                                        if let Some(run) = current.run.as_mut()
                                            && current.run_properties_depth.is_some()
                                        {
                                            run.inner_shadow = Some(shadow);
                                        } else {
                                            current.inner_shadow = Some(shadow);
                                        }
                                    }
                                    PptxShadowKind::Glow => {
                                        current.glow = Some(Glow {
                                            color: shadow.color,
                                            radius: shadow.blur,
                                        });
                                    }
                                }
                            }
                        }
                        "glow" if current.shape_properties_depth.is_some() || current.run_properties_depth.is_some() => {
                            current.shadow_kind = PptxShadowKind::Glow;
                            current.shadow_depth = (!empty).then_some(depth);
                            current.shadow_color = None;
                            current.shadow_blur = numeric_attribute(&attributes, "rad", part)?
                                .unwrap_or(0) as f32
                                / EMU_PER_CSS_PIXEL;
                            current.shadow_distance = 0.0;
                            current.shadow_direction_degrees = 0.0;
                            if empty {
                                current.glow = Some(Glow {
                                    color: 0x0000_0080,
                                    radius: current.shadow_blur,
                                });
                            }
                        }
                        "reflection" if current.run_properties_depth.is_some()
                            || current.shape_properties_depth.is_some() =>
                        {
                            let reflection = parse_drawingml_reflection(&attributes, part)?;
                            if let Some(run) = current.run.as_mut()
                                && current.run_properties_depth.is_some()
                            {
                                run.reflection = Some(reflection);
                            } else {
                                current.reflection = Some(reflection);
                            }
                        }
                        "softEdge" if current.shape_properties_depth.is_some() => {
                            current.soft_edge = Some(
                                numeric_attribute(&attributes, "rad", part)?.unwrap_or(0) as f32
                                    / EMU_PER_CSS_PIXEL,
                            );
                        }
                        "srgbClr" if current.shadow_depth.is_some() => {
                            current.shadow_color = string_attribute(&attributes, "val", part)?
                                .and_then(|value| parse_rgb_color(&value));
                        }
                        "sysClr" if current.shadow_depth.is_some() => {
                            current.shadow_color = string_attribute(
                                &attributes,
                                "lastClr",
                                part,
                            )?
                            .and_then(|value| parse_rgb_color(&value));
                        }
                        "prstClr" if current.shadow_depth.is_some() => {
                            current.shadow_color = string_attribute(&attributes, "val", part)?
                                .as_deref()
                                .and_then(drawingml_preset_color);
                        }
                        "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff"
                            if current.shadow_depth.is_some() =>
                        {
                            let ratio = color_transform_percentage(&attributes, part)?;
                            if current.shadow_color.is_none() {
                                current.shadow_color = Some(0x0000_00ff);
                            }
                            if let Some(color) = current.shadow_color.as_mut() {
                                apply_drawingml_color_transform(color, local, ratio);
                            }
                        }
                        "p" => {
                            if current.paragraph_count != 0 {
                                append_shape_line_break(current, depth, '\n');
                            }
                            current.paragraph_count = current.paragraph_count.saturating_add(1);
                            current.current_paragraph_style = ParagraphStyle::default();
                            current.current_paragraph_level = 0;
                            current.paragraph_pending = true;
                        }
                        local if paragraph_level(local).is_some() => {
                            let level = paragraph_level(local).unwrap_or(0);
                            current.paragraph_style_capture = Some(ParagraphStyleCapture {
                                depth,
                                target: ParagraphStyleTarget::Level(level),
                                style: paragraph_style_from_attributes(&attributes, part)?,
                                spacing: None,
                                bullet_color_depth: None,
                                text_effect_depth: None,
                            });
                        }
                        "defPPr" => {
                            current.paragraph_style_capture = Some(ParagraphStyleCapture {
                                depth,
                                target: ParagraphStyleTarget::Default,
                                style: paragraph_style_from_attributes(&attributes, part)?,
                                spacing: None,
                                bullet_color_depth: None,
                                text_effect_depth: None,
                            });
                        }
                        "pPr" => {
                            if boolean_attribute(&attributes, "rtl", part)?.unwrap_or(false) {
                                current.text_direction = TextDirection::Rtl;
                            }
                            current.current_paragraph_level = numeric_attribute(
                                &attributes,
                                "lvl",
                                part,
                            )?
                            .and_then(|value| usize::try_from(value).ok())
                            .unwrap_or(0)
                            .min(TEXT_LEVEL_COUNT.saturating_sub(1));
                            current.paragraph_style_capture = Some(ParagraphStyleCapture {
                                depth,
                                target: ParagraphStyleTarget::Current,
                                style: paragraph_style_from_attributes(&attributes, part)?,
                                spacing: None,
                                bullet_color_depth: None,
                                text_effect_depth: None,
                            });
                        }
                        "lnSpc" | "spcBef" | "spcAft" | "spcPct" | "spcPts"
                        | "buNone" | "buChar" | "buAutoNum" | "buFont" | "buSzPct"
                        | "buSzPts" | "buClr" | "blip"
                            if current.paragraph_style_capture.is_some() =>
                        {
                            update_paragraph_style_capture(
                                current.paragraph_style_capture.as_mut().ok_or_else(|| {
                                    format_error(part, "paragraph style capture is missing")
                                })?,
                                local,
                                &attributes,
                                depth,
                                part,
                                theme,
                            )?;
                        }
                        "srgbClr" | "sysClr"
                            if current
                                .paragraph_style_capture
                                .as_ref()
                                .is_some_and(|capture| capture.bullet_color_depth.is_some()) =>
                        {
                            update_paragraph_style_capture(
                                current.paragraph_style_capture.as_mut().ok_or_else(|| {
                                    format_error(part, "paragraph style capture is missing")
                                })?,
                                local,
                                &attributes,
                                depth,
                                part,
                                theme,
                            )?;
                        }
                        "r" | "fld" => {
                            if current.run.is_some() {
                                return Err(format_error(part, "nested DrawingML text runs are invalid"));
                            }
                            begin_shape_paragraph(current, true);
                            let mut run = TextRunState::from_shape(current, depth);
                            if local == "fld"
                                && string_attribute(&attributes, "type", part)?.as_deref()
                                    == Some("slidenum")
                            {
                                run.replacement_text = Some(unit_index.saturating_add(1).to_string());
                            }
                            current.run = Some(run);
                        }
                        "br" => {
                            begin_shape_paragraph(current, true);
                            append_shape_line_break(current, depth, '\u{2028}');
                        }
                        "endParaRPr" => {
                            if let Some(size) = numeric_attribute(&attributes, "sz", part)? {
                                apply_end_paragraph_font_size(
                                    current.paragraph_pending,
                                    &mut current.current_paragraph_style,
                                    &mut current.paragraph_layout_styles,
                                    size as f32 / 100.0 * POINTS_TO_CSS_PIXELS,
                                );
                            }
                        }
                        "rPr" | "defRPr" => {
                            if local == "defRPr"
                                && let Some(capture) = current.paragraph_style_capture.as_mut()
                            {
                                update_paragraph_style_capture(
                                    capture,
                                    local,
                                    &attributes,
                                    depth,
                                    part,
                                    theme,
                                )?;
                            }
                            begin_shape_run_properties(
                                current,
                                &attributes,
                                depth,
                                empty,
                                part,
                            )?;
                        }
                        "hlinkClick" if current.run_properties_depth.is_some() => {
                            if let Some(run) = current.run.as_mut() {
                                run.color = theme.color("hlink").unwrap_or(0x0563_c1ff);
                                run.underline = true;
                            }
                        }
                        "latin" | "ea" | "cs" | "sym" => {
                            apply_shape_run_typeface(current, local, &attributes, part, theme)?;
                        }
                        "fillRef" => {
                            current.fill_reference_depth = (!empty
                                && drawingml_fill_reference_has_paint(&attributes, part)?)
                                .then_some(depth);
                        }
                        "lnRef" => {
                            current.line_reference_depth = (!empty).then_some(depth);
                            current.style_stroke_width = numeric_attribute(&attributes, "idx", part)?
                                .and_then(|index| usize::try_from(index).ok())
                                .and_then(|index| theme.line_width(index));
                        }
                        "fontRef" => current.font_reference_depth = (!empty).then_some(depth),
                        "schemeClr" => {
                            if let Some(value) = string_attribute(&attributes, "val", part)?
                                && let Some(color) = theme.color(&value)
                            {
                                if current.shadow_depth.is_some() {
                                    current.shadow_color = Some(color);
                                } else if current.fill_reference_depth.is_some() {
                                    current.style_fill = Some(Paint::Solid(color));
                                } else if current.line_reference_depth.is_some() {
                                    current.style_stroke = Some(Paint::Solid(color));
                                } else if current.font_reference_depth.is_some() {
                                    current.font_color =
                                        current.diagram_text_color.unwrap_or(color);
                                    current.explicit_text_color = true;
                                }
                            }
                        }
                        "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff"
                            if current.fill_reference_depth.is_some()
                                || current.line_reference_depth.is_some()
                                || current.font_reference_depth.is_some() =>
                        {
                            let ratio = color_transform_percentage(&attributes, part)?;
                            if current.fill_reference_depth.is_some() {
                                if let Some(fill) = current.style_fill.as_mut() {
                                    apply_paint_color_transform(fill, local, ratio);
                                }
                            } else if current.line_reference_depth.is_some() {
                                if let Some(stroke) = current.style_stroke.as_mut() {
                                    apply_paint_color_transform(stroke, local, ratio);
                                }
                            } else if current.font_reference_depth.is_some() {
                                apply_drawingml_color_transform(
                                    &mut current.font_color,
                                    local,
                                    ratio,
                                );
                            }
                        }
                        "t" if current.math_capture.is_none() => {
                            current.collecting_text = !empty;
                        }
                        _ => {}
                    }
                    if empty
                        && current
                            .paragraph_style_capture
                            .as_ref()
                            .is_some_and(|capture| capture.depth == depth)
                    {
                        commit_shape_paragraph_style(current);
                    }
                }
                if let Some(current) = picture.as_mut() {
                    super::drawingml::drawingml_image_fill_mapping(&mut current.image_mapping, local, &attributes, part)?;
                    if local == "tile" { current.image_tile = true; }
                    if !current.is_background && current.shape_properties_depth.is_none() {
                        if local == "lnRef" {
                            current.line_reference_depth = (!empty).then_some(depth);
                            current.stroke_width = numeric_attribute(&attributes, "idx", part)?
                                .and_then(|index| usize::try_from(index).ok())
                                .and_then(|index| theme.line_width(index))
                                .unwrap_or(0.0);
                        } else if local == "effectRef" {
                            current.style_shadow = numeric_attribute(&attributes, "idx", part)?
                                .is_some_and(|index| index > 0);
                        } else if local == "schemeClr"
                            && current.line_reference_depth.is_some()
                            && let Some(value) = string_attribute(&attributes, "val", part)?
                            && let Some(color) = theme.color(&value)
                        {
                            current.stroke = Paint::Solid(color);
                        } else if matches!(
                            local,
                            "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff"
                        ) && current.line_reference_depth.is_some()
                        {
                            apply_paint_color_transform(
                                &mut current.stroke,
                                local,
                                color_transform_percentage(&attributes, part)?,
                            );
                        }
                    }
                    if current.is_background {
                        if let Some(capture) = current.fill_capture.as_mut() {
                            capture.start(local, &attributes, part, theme)?;
                        } else if local == "noFill" {
                            current.fill = Paint::None;
                            current.background_fill_defined = true;
                        } else if local == "bgRef" {
                            current.background_fill_reference = numeric_attribute(&attributes, "idx", part)?;
                            current.fill_reference_depth = (!empty).then_some(depth);
                            if let Some(index) = numeric_attribute(&attributes, "idx", part)?
                                .and_then(|index| index.checked_sub(1001))
                                .and_then(|index| usize::try_from(index).ok())
                                && let Some(image_fill) = theme.background_image_fills.get(index)
                                && let Some(relationship) = image_fill.relationship.as_ref()
                            {
                                current.fill = resolve_image_fill_relationship(
                                    relationship,
                                    part,
                                    ImageCrop::default(),
                                    image_fill.tile,
                                    package,
                                    state,
                                    content_types,
                                )?;
                                if let Paint::Image { mapping, .. } = &mut current.fill {
                                    *mapping = (image_fill.tile || image_fill.mapping != crate::model::ImageFillMapping::default()).then(|| Box::new(image_fill.mapping.clone()));
                                }
                                current.image_adjustment = image_fill.adjustment(
                                    theme.color("accent1").unwrap_or(0x0000_00ff),
                                    theme,
                                );
                                current.background_image_fill_index = Some(index);
                            }
                        } else if current.fill_reference_depth.is_some()
                            && let Some(color) = drawingml_color_value(local, &attributes, part, theme)?
                        {
                            if matches!(current.fill, Paint::Image { .. }) {
                                if let Some(image_fill) = current.background_image_fill_index
                                    .and_then(|index| theme.background_image_fills.get(index))
                                {
                                    current.image_adjustment = image_fill.adjustment(color, theme);
                                }
                            } else {
                                current.fill = Paint::Solid(color);
                            }
                        } else if matches!(
                            local,
                            "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff"
                        )
                            && current.fill_reference_depth.is_some()
                        {
                            let ratio = color_transform_percentage(&attributes, part)?;
                            apply_paint_color_transform(&mut current.fill, local, ratio);
                        } else if let Some(capture) =
                            PaintCapture::new(local, &attributes, part, depth)?
                        {
                            current.background_fill_defined = true;
                            current.fill_capture = Some(capture);
                        }
                    } else if current.shape_properties_depth.is_some() {
                        if let Some(capture) = current.fill_capture.as_mut() {
                            capture.start(local, &attributes, part, theme)?;
                        } else if current.line_depth.is_some() {
                            if local == "noFill" {
                                current.stroke = Paint::None;
                            } else if let Some(capture) =
                                PaintCapture::new(local, &attributes, part, depth)?
                            {
                                current.fill_capture = Some(capture);
                            }
                        }
                    }
                    if current.shape_properties_depth.is_some() {
                        current.geometry.start(local, &attributes, empty, depth, part,
                            current.width.unwrap_or(0) as f32, current.height.unwrap_or(0) as f32)?;
                    }
                    match local {
                        "cNvPr" if current.shape_id.is_none() => {
                            current.shape_id = numeric_attribute(&attributes, "id", part)?
                                .map(u32::try_from)
                                .transpose()
                                .map_err(|_| {
                                    format_error(part, "picture source ID exceeds u32 range")
                                })?;
                            current.metadata = non_visual_metadata(&attributes, part)?;
                            current.non_visual_depth = (!empty).then_some(depth);
                        }
                        "hlinkClick" if current.non_visual_depth.is_some() => {
                            current.metadata.click_action =
                                Some(drawingml_action(&attributes, &relationship_map, part)?);
                        }
                        "hlinkHover" if current.non_visual_depth.is_some() => {
                            current.metadata.hover_action =
                                Some(drawingml_action(&attributes, &relationship_map, part)?);
                        }
                        "ph" => {
                            current.placeholder_type = string_attribute(&attributes, "type", part)?;
                            current.placeholder_index = numeric_attribute(&attributes, "idx", part)?
                                .map(u32::try_from)
                                .transpose()
                                .map_err(|_| {
                                    format_error(part, "placeholder index exceeds u32 range")
                                })?;
                        }
                        "spPr" if depth == current.depth.saturating_add(1) => {
                            current.shape_properties_depth = (!empty).then_some(depth);
                        }
                        "xfrm"
                            if current
                                .shape_properties_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            if current.transform_depth.is_some() || current.has_transform {
                                return Err(format_error(
                                    part,
                                    "picture must contain exactly one non-nested transform",
                                ));
                            }
                            current.has_transform = true;
                            current.transform_depth = (!empty).then_some(depth);
                            current.rotation_degrees = signed_numeric_attribute(
                                &attributes,
                                "rot",
                                part,
                            )?
                            .unwrap_or(0) as f32
                                / 60_000.0;
                            current.flip_horizontal = boolean_attribute(
                                &attributes,
                                "flipH",
                                part,
                            )?
                            .unwrap_or(false);
                            current.flip_vertical = boolean_attribute(
                                &attributes,
                                "flipV",
                                part,
                            )?
                            .unwrap_or(false);
                        }
                        "off"
                            if current
                                .transform_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            current.x = Some(
                                signed_numeric_attribute(&attributes, "x", part)?.ok_or_else(
                                    || format_error(part, "picture transform is missing x"),
                                )?,
                            );
                            current.y = Some(
                                signed_numeric_attribute(&attributes, "y", part)?.ok_or_else(
                                    || format_error(part, "picture transform is missing y"),
                                )?,
                            );
                        }
                        "ext"
                            if current
                                .transform_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            current.width = Some(
                                numeric_attribute(&attributes, "cx", part)?.ok_or_else(|| {
                                    format_error(part, "picture transform is missing cx")
                                })?,
                            );
                            current.height = Some(
                                numeric_attribute(&attributes, "cy", part)?.ok_or_else(|| {
                                    format_error(part, "picture transform is missing cy")
                                })?,
                            );
                        }
                        _ if (current.shape_properties_depth.is_some() || current.is_background)
                            && current.effects.start(
                                local,
                                &attributes,
                                empty,
                                depth,
                                part,
                                |kind, value| picture_effect_color(kind, value, theme, part),
                            )? => {}
                        "ln" if current.shape_properties_depth.is_some() => {
                            current.line_depth = (!empty).then_some(depth);
                            let (stroke_width, style) =
                                drawingml_stroke_style(&attributes, part)?;
                            current.stroke_width = stroke_width;
                            current.line_cap = style.cap;
                            current.line_compound = style.compound;
                            current.line_alignment = style.alignment;
                        }
                        "round" if current.line_depth.is_some() => {
                            current.line_join = LineJoin::Round;
                        }
                        "bevel" if current.line_depth.is_some() => {
                            current.line_join = LineJoin::Bevel;
                        }
                        "miter" if current.line_depth.is_some() => {
                            current.line_join = LineJoin::Miter;
                            current.miter_limit = optional_percentage_attribute(
                                &attributes,
                                "lim",
                                part,
                            )?
                            .unwrap_or(8.0);
                        }
                        "prstDash" if current.line_depth.is_some() => {
                            current.dash_pattern = DrawingMlDashPattern::from_attribute(
                                string_attribute(&attributes, "val", part)?.as_deref(),
                            );
                        }
                        "audioFile" | "videoFile" => {
                            let relationship_id = string_attribute(&attributes, "embed", part)?
                                .or(string_attribute(&attributes, "link", part)?);
                            if let Some(relationship_id) = relationship_id
                                && current.media_relationship_id.is_none()
                            {
                                current.media_relationship_id = Some(relationship_id);
                                current.media_kind_hint = Some(if local == "audioFile" {
                                    MediaKind::Audio
                                } else {
                                    MediaKind::Video
                                });
                            }
                        }
                        "media" => {
                            let relationship_id = string_attribute(&attributes, "embed", part)?
                                .or(string_attribute(&attributes, "link", part)?);
                            if let Some(relationship_id) = relationship_id {
                                current.media_relationship_id = Some(relationship_id);
                            }
                        }
                        "blip" => {
                            current.blip_depth = (!empty).then_some(depth);
                            let embedded = string_attribute(&attributes, "embed", part)?;
                            let linked = string_attribute(&attributes, "link", part)?;
                            if embedded.is_some() && linked.is_some() {
                                return Err(format_error(
                                    part,
                                    "picture blip cannot be both embedded and linked",
                                ));
                            }
                            if let Some(embedded) = embedded
                                && current.embedded_relationship_id.replace(embedded).is_some()
                            {
                                return Err(format_error(
                                    part,
                                    "picture contains multiple embedded image references",
                                ));
                            }
                            if let Some(linked) = linked
                                && current.linked_relationship_id.replace(linked).is_some()
                            {
                                return Err(format_error(
                                    part,
                                    "picture contains multiple linked image references",
                                ));
                            }
                        }
                        "clrChange"
                            if current
                                .blip_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            if current.color_change_depth.is_some() {
                                return Err(format_error(
                                    part,
                                    "picture contains nested color-change effects",
                                ));
                            }
                            current.color_change_use_alpha =
                                boolean_attribute(&attributes, "useA", part)?.unwrap_or(true);
                            current.color_change_depth = (!empty).then_some(depth);
                        }
                        "alphaModFix"
                            if current
                                .blip_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            current.opacity *= numeric_attribute(&attributes, "amt", part)?
                                .unwrap_or(100_000) as f32
                                / 100_000.0;
                            current.opacity = current.opacity.clamp(0.0, 1.0);
                        }
                        "clrFrom"
                            if current
                                .color_change_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            current.color_change_from_depth = (!empty).then_some(depth);
                        }
                        "clrTo"
                            if current
                                .color_change_depth
                                .is_some_and(|start| depth == start.saturating_add(1)) =>
                        {
                            current.color_change_to_depth = (!empty).then_some(depth);
                        }
                        "srgbClr" | "scrgbClr" | "schemeClr" | "sysClr" | "prstClr"
                            if current.color_change_from_depth.is_some()
                                || current.color_change_to_depth.is_some() =>
                        {
                            if let Some(color) =
                                drawingml_color_value(local, &attributes, part, theme)?
                            {
                                if current.color_change_to_depth.is_some() {
                                    current.color_change_to = Some(color);
                                } else {
                                    current.color_change_from = Some(color);
                                }
                            }
                        }
                        "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff"
                            if current.color_change_from_depth.is_some()
                                || current.color_change_to_depth.is_some() =>
                        {
                            let ratio = color_transform_percentage(&attributes, part)?;
                            let color = if current.color_change_to_depth.is_some() {
                                current.color_change_to.as_mut()
                            } else {
                                current.color_change_from.as_mut()
                            };
                            if let Some(color) = color {
                                apply_drawingml_color_transform(color, local, ratio);
                            }
                        }
                        "svgBlip" => {
                            let embedded = string_attribute(&attributes, "embed", part)?
                                .ok_or_else(|| {
                                    format_error(
                                        part,
                                        "preferred SVG reference is missing r:embed",
                                    )
                                })?;
                            if current
                                .preferred_svg_relationship_id
                                .replace(embedded)
                                .is_some()
                            {
                                return Err(format_error(
                                    part,
                                    "picture contains multiple preferred SVG references",
                                ));
                            }
                        }
                        "srcRect" => {
                            current.crop = super::drawingml::drawingml_picture_crop(&attributes, part)?;
                        }
                        _ => {}
                    }
                }
                if let Some(current) = graphic_frame.as_mut() {
                    if let Some(cell) = current.cell.as_mut() {
                        if let Some((_, capture)) = cell.paint_capture.as_mut() {
                            capture.start(local, &attributes, part, theme)?;
                        } else if cell.ignored_border_depth.is_some() {
                            // Diagonal table borders are not represented yet. Their nested
                            // paint must not be mistaken for the cell background fill.
                        } else if local == "prstDash"
                            && let Some((_, side)) = cell.border_depth
                        {
                            cell.border_mut(side).dash_pattern =
                                DrawingMlDashPattern::from_attribute(
                                    string_attribute(&attributes, "val", part)?.as_deref(),
                                );
                        } else if local == "highlight" && cell.run_properties_depth.is_some() {
                            cell.paint_capture =
                                PaintCapture::new("solidFill", &attributes, part, depth)?
                                .map(|capture| (CellPaintTarget::Highlight, capture));
                        } else if matches!(local, "solidFill" | "gradFill" | "pattFill")
                            && (cell.properties_depth.is_some()
                                || cell.run_properties_depth.is_some()
                                || cell.paragraph_default_run_depth.is_some())
                        {
                            let target = if cell.paragraph_default_run_depth.is_some() {
                                CellPaintTarget::ParagraphText
                            } else if cell.run_properties_depth.is_some() {
                                CellPaintTarget::Text
                            } else if let Some((_, side)) = cell.border_depth {
                                CellPaintTarget::Stroke(side)
                            } else {
                                CellPaintTarget::Fill
                            };
                            if matches!(target, CellPaintTarget::Fill) {
                                cell.explicit_fill = true;
                            }
                            cell.paint_capture =
                                PaintCapture::new(local, &attributes, part, depth)?
                                    .map(|capture| (target, capture));
                        } else if local == "noFill" && cell.properties_depth.is_some() {
                            if let Some((_, side)) = cell.border_depth {
                                let border = cell.border_mut(side);
                                border.paint = Paint::None;
                                border.explicit = true;
                            } else {
                                cell.fill = Paint::None;
                                cell.explicit_fill = true;
                            }
                        }
                    }
                    match local {
                        "cNvPr" if current.shape_id.is_none() => {
                            current.shape_id = numeric_attribute(&attributes, "id", part)?
                                .map(u32::try_from)
                                .transpose()
                                .map_err(|_| {
                                    format_error(part, "graphic-frame source ID exceeds u32 range")
                                })?;
                            current.metadata = non_visual_metadata(&attributes, part)?;
                            current.non_visual_depth = (!empty).then_some(depth);
                        }
                        "hlinkClick" if current.non_visual_depth.is_some() => {
                            current.metadata.click_action =
                                Some(drawingml_action(&attributes, &relationship_map, part)?);
                        }
                        "hlinkHover" if current.non_visual_depth.is_some() => {
                            current.metadata.hover_action =
                                Some(drawingml_action(&attributes, &relationship_map, part)?);
                        }
                        "xfrm" if current.transform_depth.is_none() => {
                            current.transform_depth = (!empty).then_some(depth);
                        }
                        "off" if current.transform_depth.is_some() => {
                            current.x = signed_numeric_attribute(&attributes, "x", part)?;
                            current.y = signed_numeric_attribute(&attributes, "y", part)?;
                        }
                        "ext" if current.transform_depth.is_some() => {
                            current.width = numeric_attribute(&attributes, "cx", part)?;
                            current.height = numeric_attribute(&attributes, "cy", part)?;
                        }
                        "tbl" => current.table_depth = (!empty).then_some(depth),
                        "tblPr" if current.table_depth.is_some() => {
                            current.first_row =
                                boolean_attribute(&attributes, "firstRow", part)?.unwrap_or(false);
                            current.last_row =
                                boolean_attribute(&attributes, "lastRow", part)?.unwrap_or(false);
                            current.first_column =
                                boolean_attribute(&attributes, "firstCol", part)?.unwrap_or(false);
                            current.last_column =
                                boolean_attribute(&attributes, "lastCol", part)?.unwrap_or(false);
                            current.band_rows =
                                boolean_attribute(&attributes, "bandRow", part)?.unwrap_or(false);
                            current.band_columns =
                                boolean_attribute(&attributes, "bandCol", part)?.unwrap_or(false);
                        }
                        "tableStyleId" if current.table_depth.is_some() => {
                            current.table_style_id.clear();
                            current.collecting_table_style_id = !empty;
                        }
                        "gridCol" if current.table_depth.is_some() => {
                            current
                                .columns
                                .push(numeric_attribute(&attributes, "w", part)?.unwrap_or(0));
                        }
                        "tr" if current.table_depth.is_some() => {
                            if current.row.is_some() {
                                return Err(format_error(part, "nested table rows are invalid"));
                            }
                            current.row = Some(TableRowState {
                                depth,
                                height: numeric_attribute(&attributes, "h", part)?.unwrap_or(0),
                                cells: Vec::new(),
                            });
                        }
                        "tc" if current.row.is_some() => {
                            if current.cell.is_some() {
                                return Err(format_error(part, "nested table cells are invalid"));
                            }
                            let base_run = TableCellRunDefaults::from_presentation_defaults(
                                theme,
                                state.default_text_styles.first(),
                            );
                            let font_family = base_run.font_family.clone();
                            let font_east_asian = base_run.font_east_asian.clone();
                            let font_complex_script = base_run.font_complex_script.clone();
                            let font_size = base_run.font_size;
                            let font_color = base_run.font_color;
                            let bold = base_run.bold;
                            let italic = base_run.italic;
                            let underline = base_run.underline;
                            let strikethrough = base_run.strikethrough;
                            let baseline_shift = base_run.baseline_shift;
                            let letter_spacing = base_run.letter_spacing;
                            let east_asian_line_breaks = base_run.east_asian_line_breaks;
                            let capitalization = base_run.capitalization;
                            current.cell = Some(TableCellState {
                                depth,
                                text: String::new(),
                                collecting_text: false,
                                fill: Paint::None,
                                explicit_fill: false,
                                style_fill: None,
                                effects: None,
                                borders: std::array::from_fn(|_| TableCellBorder::default()),
                                grid_span: numeric_attribute(&attributes, "gridSpan", part)?
                                    .and_then(|span| usize::try_from(span).ok())
                                    .unwrap_or(1)
                                    .max(1),
                                row_span: numeric_attribute(&attributes, "rowSpan", part)?
                                    .and_then(|span| usize::try_from(span).ok())
                                    .unwrap_or(1)
                                    .max(1),
                                horizontal_merge: boolean_attribute(
                                    &attributes,
                                    "hMerge",
                                    part,
                                )?
                                .unwrap_or(false),
                                vertical_merge: boolean_attribute(
                                    &attributes,
                                    "vMerge",
                                    part,
                                )?
                                .unwrap_or(false),
                                properties_depth: None,
                                border_depth: None,
                                ignored_border_depth: None,
                                paint_capture: None,
                                paragraph_count: 0,
                                paragraph_styles: state.default_text_styles.clone(),
                                local_paragraph_styles: vec![ParagraphStyle::default(); TEXT_LEVEL_COUNT],
                                run_style: ParagraphStyle::default(),
                                run_styles: Vec::new(),
                                paragraph_style_capture: None,
                                paragraph_default_run_depth: None,
                                current_paragraph_style: ParagraphStyle::default(),
                                current_paragraph_level: 0,
                                paragraph_pending: false,
                                paragraph_layouts: Vec::new(),
                                paragraph_layout_styles: Vec::new(),
                                paragraph_run_starts: Vec::new(),
                                numbering_counters: [0; TEXT_LEVEL_COUNT],
                                align: TextAlign::Start,
                                text_inset_left: 9.6,
                                text_inset_right: 9.6,
                                text_inset_top: 4.8,
                                text_inset_bottom: 4.8,
                                vertical_align: TextVerticalAlign::Top,
                                text_orientation: TextOrientation::Horizontal,
                                base_run,
                                font_family,
                                font_east_asian,
                                font_complex_script,
                                font_size,
                                font_color,
                                bold,
                                italic,
                                underline,
                                strikethrough,
                                baseline_shift,
                                letter_spacing,
                                east_asian_line_breaks,
                                capitalization,
                                text_direction: TextDirection::Auto,
                                run: None,
                                runs: Vec::new(),
                                run_effects: Vec::new(),
                                run_properties_depth: None,
                                explicit_text_color: false,
                            });
                        }
                        "tcPr" if current.cell.is_some() => {
                            if let Some(cell) = current.cell.as_mut() {
                                cell.properties_depth = (!empty).then_some(depth);
                                cell.text_inset_left = signed_numeric_attribute(&attributes, "marL", part)?
                                    .or(signed_numeric_attribute(&attributes, "lIns", part)?)
                                    .map_or(9.6, |value| value as f32 / EMU_PER_CSS_PIXEL);
                                cell.text_inset_right = signed_numeric_attribute(&attributes, "marR", part)?
                                    .or(signed_numeric_attribute(&attributes, "rIns", part)?)
                                    .map_or(9.6, |value| value as f32 / EMU_PER_CSS_PIXEL);
                                cell.text_inset_top = signed_numeric_attribute(&attributes, "marT", part)?
                                    .or(signed_numeric_attribute(&attributes, "tIns", part)?)
                                    .map_or(4.8, |value| value as f32 / EMU_PER_CSS_PIXEL);
                                cell.text_inset_bottom = signed_numeric_attribute(&attributes, "marB", part)?
                                    .or(signed_numeric_attribute(&attributes, "bIns", part)?)
                                    .map_or(4.8, |value| value as f32 / EMU_PER_CSS_PIXEL);
                                if let Some(anchor) = string_attribute(&attributes, "anchor", part)? {
                                    cell.vertical_align = match anchor.as_str() {
                                        "ctr" => TextVerticalAlign::Center,
                                        "b" => TextVerticalAlign::Bottom,
                                        _ => TextVerticalAlign::Top,
                                    };
                                }
                                if let Some(vertical) = string_attribute(&attributes, "vert", part)? {
                                    cell.text_orientation =
                                        drawingml_text_orientation(vertical.as_str());
                                }
                            }
                        }
                        "lnL" | "lnR" | "lnT" | "lnB" if current.cell.is_some() => {
                            if let Some(cell) = current.cell.as_mut() {
                                let side = TableBorderSide::from_element(local).ok_or_else(|| {
                                    format_error(part, "table border side is invalid")
                                })?;
                                let border = cell.border_mut(side);
                                border.explicit = true;
                                border.stroke_style = drawingml_stroke_style(&attributes, part)?.1;
                                border.width = numeric_attribute(&attributes, "w", part)?
                                    .unwrap_or(12_700) as f32
                                    / EMU_PER_CSS_PIXEL;
                                cell.border_depth = (!empty).then_some((depth, side));
                            }
                        }
                        "lnTlToBr" | "lnBlToTr" if current.cell.is_some() => {
                            if let Some(cell) = current.cell.as_mut() {
                                cell.ignored_border_depth = (!empty).then_some(depth);
                            }
                        }
                        "p" if current.cell.is_some() => {
                            if let Some(cell) = current.cell.as_mut() {
                                if cell.paragraph_count != 0 {
                                    append_table_cell_line_break(cell, depth, '\n');
                                }
                                cell.paragraph_count = cell.paragraph_count.saturating_add(1);
                                cell.current_paragraph_style = ParagraphStyle::default();
                                cell.current_paragraph_level = 0;
                                cell.paragraph_pending = true;
                            }
                        }
                        local if paragraph_level(local).is_some() && current.cell.is_some() => {
                            let level = paragraph_level(local).unwrap_or(0);
                            if let Some(cell) = current.cell.as_mut() {
                                cell.paragraph_style_capture = Some(ParagraphStyleCapture {
                                    depth,
                                    target: ParagraphStyleTarget::Level(level),
                                    style: paragraph_style_from_attributes(&attributes, part)?,
                                    spacing: None,
                                    bullet_color_depth: None,
                                    text_effect_depth: None,
                                });
                                if empty {
                                    commit_table_cell_paragraph_style(cell);
                                }
                            }
                        }
                        "defPPr" if current.cell.is_some() => {
                            if let Some(cell) = current.cell.as_mut() {
                                cell.paragraph_style_capture = Some(ParagraphStyleCapture {
                                    depth,
                                    target: ParagraphStyleTarget::Default,
                                    style: paragraph_style_from_attributes(&attributes, part)?,
                                    spacing: None,
                                    bullet_color_depth: None,
                                    text_effect_depth: None,
                                });
                                if empty {
                                    commit_table_cell_paragraph_style(cell);
                                }
                            }
                        }
                        "pPr" if current.cell.is_some() => {
                            if let Some(cell) = current.cell.as_mut() {
                                if boolean_attribute(&attributes, "rtl", part)?.unwrap_or(false) {
                                    cell.text_direction = TextDirection::Rtl;
                                }
                                cell.current_paragraph_level = numeric_attribute(
                                    &attributes,
                                    "lvl",
                                    part,
                                )?
                                .and_then(|value| usize::try_from(value).ok())
                                .unwrap_or(0)
                                .min(TEXT_LEVEL_COUNT.saturating_sub(1));
                                cell.paragraph_style_capture = Some(ParagraphStyleCapture {
                                    depth,
                                    target: ParagraphStyleTarget::Current,
                                    style: paragraph_style_from_attributes(&attributes, part)?,
                                    spacing: None,
                                    bullet_color_depth: None,
                                    text_effect_depth: None,
                                });
                                if empty {
                                    commit_table_cell_paragraph_style(cell);
                                }
                            }
                        }
                        "defRPr"
                            if current
                                .cell
                                .as_ref()
                                .is_some_and(|cell| cell.paragraph_style_capture.is_some()) =>
                        {
                            if let Some(cell) = current.cell.as_mut() {
                                update_paragraph_style_capture(
                                    cell.paragraph_style_capture.as_mut().ok_or_else(|| {
                                        format_error(
                                            part,
                                            "table paragraph style capture is missing",
                                        )
                                    })?,
                                    local,
                                    &attributes,
                                    depth,
                                    part,
                                    theme,
                                )?;
                                cell.paragraph_default_run_depth = (!empty).then_some(depth);
                            }
                        }
                        "lnSpc" | "spcBef" | "spcAft" | "spcPct" | "spcPts"
                        | "buNone" | "buChar" | "buAutoNum" | "buFont" | "buSzPct"
                        | "buSzPts" | "buClr" | "srgbClr" | "sysClr"
                            if current
                                .cell
                                .as_ref()
                                .is_some_and(|cell| cell.paragraph_style_capture.is_some()) =>
                        {
                            if let Some(cell) = current.cell.as_mut() {
                                update_paragraph_style_capture(
                                    cell.paragraph_style_capture.as_mut().ok_or_else(|| {
                                        format_error(
                                            part,
                                            "table paragraph style capture is missing",
                                        )
                                    })?,
                                    local,
                                    &attributes,
                                    depth,
                                    part,
                                    theme,
                                )?;
                            }
                        }
                        "r" | "fld" if current.cell.is_some() => {
                            if let Some(cell) = current.cell.as_mut() {
                                if cell.run.is_some() {
                                    return Err(format_error(
                                        part,
                                        "nested table-cell text runs are invalid",
                                    ));
                                }
                                begin_table_cell_paragraph(cell);
                                cell.run_style = cell.local_paragraph_styles.get(cell.current_paragraph_level).cloned().unwrap_or_default();
                                cell.run_style.apply(&cell.current_paragraph_style);
                                cell.run = Some(TextRunState::from_table_cell(cell, depth));
                            }
                        }
                        "br" if current.cell.is_some() => {
                            if let Some(cell) = current.cell.as_mut() {
                                begin_table_cell_paragraph(cell);
                                append_table_cell_line_break(cell, depth, '\u{2028}');
                            }
                        }
                        "endParaRPr" if current.cell.is_some() => {
                            if let Some(size) = numeric_attribute(&attributes, "sz", part)?
                                && let Some(cell) = current.cell.as_mut()
                            {
                                apply_end_paragraph_font_size(
                                    cell.paragraph_pending,
                                    &mut cell.current_paragraph_style,
                                    &mut cell.paragraph_layout_styles,
                                    size as f32 / 100.0 * POINTS_TO_CSS_PIXELS,
                                );
                            }
                        }
                        "rPr" | "defRPr" if current.cell.is_some() => {
                            if let Some(cell) = current.cell.as_mut() {
                                begin_table_cell_run_properties(
                                    cell,
                                    &attributes,
                                    depth,
                                    empty,
                                    part,
                                )?;
                            }
                        }
                        "hlinkClick"
                            if current
                                .cell
                                .as_ref()
                                .is_some_and(|cell| cell.run_properties_depth.is_some()) =>
                        {
                            if let Some(cell) = current.cell.as_mut() && let Some(run) = cell.run.as_mut() {
                                run.color = theme.color("hlink").unwrap_or(0x0563_c1ff);
                                cell.run_style.font_color = Some(run.color);
                                run.underline = true;
                            }
                        }
                        "latin" | "ea" | "cs" | "sym" if current.cell.is_some() => {
                            if let Some(typeface) = string_attribute(&attributes, "typeface", part)?
                                && !typeface.is_empty()
                                && let Some(cell) = current.cell.as_mut()
                            {
                                let typeface = theme.resolve_typeface(&typeface, local);
                                if cell.paragraph_default_run_depth.is_some() {
                                    let style = &mut cell
                                        .paragraph_style_capture
                                        .as_mut()
                                        .ok_or_else(|| {
                                            format_error(
                                                part,
                                                "table paragraph style capture is missing",
                                            )
                                        })?
                                        .style;
                                    match local {
                                        "ea" => style.font_east_asian = Some(typeface),
                                        "cs" => style.font_complex_script = Some(typeface),
                                        "sym" => {}
                                        _ => style.font_family = Some(typeface),
                                    }
                                } else if let Some(run) = cell.run.as_mut() {
                                    match local {
                                        "ea" => cell.run_style.font_east_asian = Some(typeface.clone()),
                                        "cs" => cell.run_style.font_complex_script = Some(typeface.clone()),
                                        _ => cell.run_style.font_family = Some(typeface.clone()),
                                    }
                                    run.set_font(local, typeface);
                                } else {
                                    match local {
                                        "ea" => cell.font_east_asian = Some(typeface),
                                        "cs" => cell.font_complex_script = Some(typeface),
                                        "sym" => {}
                                        _ => cell.font_family = typeface,
                                    }
                                }
                            }
                        }
                        "t" if current.cell.is_some() => {
                            if let Some(cell) = current.cell.as_mut() {
                                cell.collecting_text = !empty;
                            }
                        }
                        "chart" => {
                            chart_choice.chart();
                            current.placeholder = Some("Chart");
                            if let Some(relationship_id) = relationship_id(&attributes, part)?
                                && let Some(relationship) =
                                    relationship_map.get(relationship_id.as_str())
                                && !relationship.external
                                && (relationship.type_uri.ends_with("/chart") || relationship.type_uri.ends_with("/chartEx"))
                            {
                                current.chart = if relationship.type_uri.ends_with("/chartEx") {
                                    super::drawingml::parse_chart_ex(package, &relationship.target, |value| theme.color(value))?
                                } else { parse_basic_pptx_chart(package, &relationship.target, theme)? };
                            }
                        }
                        "oleObj" => {
                            current.placeholder = Some("Embedded object");
                            let relationship_id = relationship_id(&attributes, part)?;
                            let program_id = string_attribute(&attributes, "progId", part)?;
                            let native_width = numeric_attribute(&attributes, "imgW", part)?
                                .map(|width| width as f32 / EMU_PER_CSS_PIXEL);
                            let native_height = numeric_attribute(&attributes, "imgH", part)?
                                .map(|height| height as f32 / EMU_PER_CSS_PIXEL);
                            let relationship = relationship_id
                                .as_deref()
                                .and_then(|id| relationship_map.get(id));
                            if program_id.as_deref() == Some("MSGraph.Chart.8")
                                && let Some(relationship) = relationship
                                && !relationship.external
                                && relationship.type_uri.ends_with("/oleObject")
                            {
                                let bytes = package.required_part(&relationship.target)?;
                                let colors = [
                                    "accent1", "accent2", "hlink", "folHlink", "lt2", "accent6",
                                ].map(|slot| theme.color(slot).unwrap_or(0x4472_c4ff));
                                match super::legacy::parse_msgraph_chart(
                                    &bytes,
                                    package.limits(),
                                    colors,
                                ) {
                                    Ok(chart) => current.chart = chart,
                                    Err(error) => state.diagnostics.push(
                                        Diagnostic::warning(
                                            DiagnosticCode::FormatInvalid,
                                            Phase::Parse,
                                            Fidelity::Omitted,
                                            "embedded Microsoft Graph chart could not be parsed",
                                        )
                                        .in_part(&relationship.target)
                                        .with_detail("cause", error.message),
                                    ),
                                }
                                if let Some(chart) = current.chart.as_mut() {
                                    chart.source_part.clone_from(&relationship.target);
                                    chart.native_size = native_width.zip(native_height);
                                    current.placeholder = Some("Chart");
                                }
                            }
                            if current.chart.is_none() {
                                if let Some(relationship) = relationship
                                    && !relationship.external
                                    && relationship.type_uri.ends_with("/oleObject")
                                {
                                    let bytes = package.required_part(&relationship.target)?;
                                    match super::legacy::parse_embedded_raster(
                                        &bytes,
                                        package.limits(),
                                    ) {
                                        Ok(Some((media_type, bytes))) => {
                                            current.embedded_raster = Some((
                                                media_type.to_owned(),
                                                bytes,
                                                relationship.target.clone(),
                                            ));
                                        }
                                        Ok(None) => {}
                                        Err(error) => state.diagnostics.push(
                                            Diagnostic::warning(
                                                DiagnosticCode::FormatInvalid,
                                                Phase::Parse,
                                                Fidelity::Omitted,
                                                "embedded OLE raster could not be parsed",
                                            )
                                            .in_part(&relationship.target)
                                            .with_detail("cause", error.message),
                                        ),
                                    }
                                    if current.embedded_raster.is_none()
                                        && let Some(rtf) = super::legacy::parse_embedded_rtf(
                                            &bytes,
                                            package.limits(),
                                        )?
                                        && let Some(native_width) = native_width
                                    {
                                        current.embedded_text = Some((
                                            super::flat::parse_rtf(
                                                &rtf,
                                                package.limits(),
                                                font_metrics,
                                            )?,
                                            native_width,
                                            relationship.target.clone(),
                                        ));
                                    }
                                    if current.embedded_raster.is_none()
                                        && current.embedded_text.is_none()
                                        && let Some((text, scale)) = match program_id.as_deref() {
                                            Some("Equation.3") => super::legacy::parse_embedded_equation(
                                                &bytes,
                                                package.limits(),
                                            )?
                                            .map(|text| (text, 0.55)),
                                            Some("MSWordArt.2") => super::legacy::parse_embedded_wordart(
                                                &bytes,
                                                package.limits(),
                                            )?
                                            .map(|text| (text, 0.32)),
                                            _ => None,
                                        }
                                    {
                                        current.embedded_plain_text = Some((
                                            text,
                                            scale,
                                            relationship.target.clone(),
                                        ));
                                    }
                                    if current.embedded_raster.is_none()
                                        && current.embedded_text.is_none()
                                        && let Some(document) =
                                            super::legacy::detect_and_parse(
                                                &bytes,
                                                package.limits(),
                                            )?
                                    {
                                        current.embedded_document = Some((
                                            document,
                                            relationship.target.clone(),
                                        ));
                                    }
                                }
                                if current.embedded_raster.is_none()
                                    && current.embedded_text.is_none()
                                    && current.embedded_plain_text.is_none()
                                    && current.embedded_document.is_none()
                                {
                                    let mut diagnostic = unsupported_slide_feature_diagnostic(
                                        part,
                                        "embedded-object",
                                    );
                                    if let Some(relationship) = relationship {
                                        diagnostic = diagnostic
                                            .with_detail(
                                                "relationshipType",
                                                relationship.type_uri.clone(),
                                            )
                                            .with_detail("target", relationship.target.clone());
                                    }
                                    state.diagnostics.push(diagnostic);
                                }
                            }
                        }
                        "relIds" => {
                            let colors_part = string_attribute(&attributes, "cs", part)?
                                .and_then(|relationship_id| {
                                    relationship_map.get(relationship_id.as_str())
                                })
                                .filter(|relationship| {
                                    !relationship.external
                                        && relationship.type_uri.ends_with("/diagramColors")
                                })
                                .map(|relationship| relationship.target.as_str());
                            if let Some(relationship_id) =
                                string_attribute(&attributes, "dm", part)?
                                && let Some(relationship) =
                                    relationship_map.get(relationship_id.as_str())
                            {
                                if relationship.external {
                                    state.diagnostics.push(
                                        Diagnostic::warning(
                                            DiagnosticCode::ExternalResourceBlocked,
                                            Phase::Security,
                                            Fidelity::Blocked,
                                            "external PPTX diagram data relationship was blocked",
                                        )
                                        .in_part(part),
                                    );
                                } else if relationship.type_uri.ends_with("/diagramData") {
                                    current.diagram = Some(parse_pptx_diagram(
                                        package,
                                        relationship.target.as_str(),
                                        &relationship_map,
                                        colors_part,
                                        theme.part.as_deref(),
                                        |value| theme.color(value),
                                    )?);
                                    current.placeholder = None;
                                }
                            }
                        }
                        "graphicData" => {
                            if string_attribute(&attributes, "uri", part)?
                                .is_some_and(|uri| uri.to_ascii_lowercase().contains("diagram"))
                            {
                                current.placeholder = Some("Diagram");
                            }
                        }
                        "oMath" | "oMathPara" => current.placeholder = Some("Formula"),
                        _ => {}
                    }
                }
                if shape.is_none()
                    && picture.is_none()
                    && graphic_frame.is_none()
                    && let Some(current) = groups.last_mut()
                {
                    if let Some(capture) = current.paint_capture.as_mut() {
                        capture.start(local, &attributes, part, theme)?;
                    } else if current
                        .shape_properties_depth
                        .is_some_and(|start| depth == start.saturating_add(1))
                        && matches!(local, "solidFill" | "gradFill" | "pattFill")
                    {
                        current.explicit_fill = true;
                        current.paint_capture =
                            PaintCapture::new(local, &attributes, part, depth)?;
                    } else if current
                        .shape_properties_depth
                        .is_some_and(|start| depth == start.saturating_add(1))
                        && local == "noFill"
                    {
                        current.explicit_fill = true;
                        current.fill = Paint::None;
                    }
                    match local {
                        "cNvPr" if current.shape_id.is_none() => {
                            current.shape_id = numeric_attribute(&attributes, "id", part)?
                                .map(u32::try_from)
                                .transpose()
                                .map_err(|_| format_error(part, "group source ID exceeds u32 range"))?;
                            current.metadata = non_visual_metadata(&attributes, part)?;
                            current.non_visual_depth = (!empty).then_some(depth);
                        }
                        "hlinkClick" if current.non_visual_depth.is_some() => {
                            current.metadata.click_action =
                                Some(drawingml_action(&attributes, &relationship_map, part)?);
                        }
                        "hlinkHover" if current.non_visual_depth.is_some() => {
                            current.metadata.hover_action =
                                Some(drawingml_action(&attributes, &relationship_map, part)?);
                        }
                        "grpSpPr" if depth == current.depth.saturating_add(1) => {
                            current.shape_properties_depth = (!empty).then_some(depth);
                        }
                        "scene3d" if current.shape_properties_depth.is_some() => {
                            current.three_d.get_or_insert_with(ThreeDStyle::default);
                            current.scene_3d_depth = (!empty).then_some(depth);
                        }
                        "camera" if current.scene_3d_depth.is_some() => {
                            parse_three_d_camera(
                                current.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                            )?;
                            current.camera_3d_depth = (!empty).then_some(depth);
                        }
                        "lightRig" if current.scene_3d_depth.is_some() => {
                            parse_three_d_light(
                                current.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                            )?;
                            current.light_3d_depth = (!empty).then_some(depth);
                        }
                        "backdrop" if current.scene_3d_depth.is_some() => {
                            current
                                .three_d
                                .get_or_insert_with(ThreeDStyle::default)
                                .backdrop = Some(Backdrop3D::default());
                            current.backdrop_3d_depth = (!empty).then_some(depth);
                        }
                        "anchor" | "norm" | "up" if current.backdrop_3d_depth.is_some() => {
                            let backdrop = current
                                .three_d
                                .get_or_insert_with(ThreeDStyle::default)
                                .backdrop
                                .get_or_insert_with(Backdrop3D::default);
                            parse_three_d_backdrop_point(backdrop, local, &attributes, part)?;
                        }
                        "rot" if current.camera_3d_depth.is_some() => {
                            parse_three_d_rotation(
                                current.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                                true,
                            )?;
                        }
                        "rot" if current.light_3d_depth.is_some() => {
                            parse_three_d_rotation(
                                current.three_d.get_or_insert_with(ThreeDStyle::default),
                                &attributes,
                                part,
                                false,
                            )?;
                        }
                        "xfrm" if current.transform_depth.is_none() => {
                            current.transform_depth = (!empty).then_some(depth);
                            current.rotation_degrees = signed_numeric_attribute(
                                &attributes,
                                "rot",
                                part,
                            )?
                            .unwrap_or(0) as f32
                                / 60_000.0;
                            current.flip_horizontal = boolean_attribute(
                                &attributes,
                                "flipH",
                                part,
                            )?
                            .unwrap_or(false);
                            current.flip_vertical = boolean_attribute(
                                &attributes,
                                "flipV",
                                part,
                            )?
                            .unwrap_or(false);
                        }
                        "off" if current.transform_depth.is_some() => {
                            current.x = signed_numeric_attribute(&attributes, "x", part)?;
                            current.y = signed_numeric_attribute(&attributes, "y", part)?;
                        }
                        "ext" if current.transform_depth.is_some() => {
                            current.width = numeric_attribute(&attributes, "cx", part)?;
                            current.height = numeric_attribute(&attributes, "cy", part)?;
                        }
                        "chOff" if current.transform_depth.is_some() => {
                            current.child_x = signed_numeric_attribute(&attributes, "x", part)?;
                            current.child_y = signed_numeric_attribute(&attributes, "y", part)?;
                        }
                        "chExt" if current.transform_depth.is_some() => {
                            current.child_width = numeric_attribute(&attributes, "cx", part)?;
                            current.child_height = numeric_attribute(&attributes, "cy", part)?;
                        }
                        _ => {}
                    }
                }
                if matches!(local, "grpSp" | "lockedCanvas" | "wgp") && empty {
                    let group = groups.pop().ok_or_else(|| {
                        format_error(part, "group parser state ended unexpectedly")
                    })?;
                    finish_group(group, part, state)?;
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if chart_choice.end(local_name(name), depth) { return Ok(()); }
                if word_text_depth.is_some() {
                    if word_text_depth == Some(depth) {
                        word_text_depth = None;
                    }
                    return Ok(());
                }
                let local = local_name(name);
                if local == "t"
                    && let Some(current) = shape.as_mut()
                    && current.math_capture.is_none()
                {
                    current.collecting_text = false;
                }
                if let Some(current) = shape.as_mut() {
                    if matches!(local, "oMath" | "oMathPara")
                        && current
                            .math_capture
                            .as_ref()
                            .is_some_and(|math| math.depth == depth)
                    {
                        let result = current
                            .math_capture
                            .take()
                            .ok_or_else(|| format_error(part, "OMML parser state is missing"))?
                            .finish(part)?;
                        if result.text.trim().is_empty() {
                            current.omit_empty_math = true;
                            if !unsupported_formula_reported {
                                state.diagnostics.push(
                                    Diagnostic::warning(
                                        DiagnosticCode::UnsupportedFeature,
                                        Phase::Render,
                                        Fidelity::Omitted,
                                        "PPTX formula contains no readable semantic content and was omitted",
                                    )
                                    .in_part(part),
                                );
                                unsupported_formula_reported = true;
                            }
                        } else {
                            if current.text.trim().is_empty() && current.math_root.is_none() {
                                current.math_root = Some((result.root, result.text.clone()));
                            } else {
                                current.math_root = None;
                            }
                            current.text.push_str(&result.text);
                            current.runs.push(TextRun {
                                paint: None,
                                east_asian_line_breaks: true,
                                text: result.text,
                                font_family: "Cambria Math".to_owned(),
                                font_size: current.font_size,
                                color: current.font_color,
                                bold: current.bold,
                                italic: current.italic,
                                underline: current.underline,
                                strikethrough: current.strikethrough,
                                highlight: current.highlight,
                                baseline_shift: current.baseline_shift,
                                letter_spacing: current.letter_spacing,
                                horizontal_scale: 1.0,
                            });
                            if result.approximate && !unsupported_formula_reported {
                                let message = if current
                                    .math_root
                                    .as_ref()
                                    .is_some_and(|(root, _)| root.needs_box_layout())
                                {
                                    "PPTX formula has unsupported OMML properties and uses approximate layout"
                                } else {
                                    "complex PPTX formula structure was flattened to readable text"
                                };
                                state.diagnostics.push(
                                    Diagnostic::warning(
                                        DiagnosticCode::UnsupportedFeature,
                                        Phase::Render,
                                        Fidelity::Approximate,
                                        message,
                                    )
                                    .in_part(part),
                                );
                                unsupported_formula_reported = true;
                            }
                        }
                    } else if let Some(math) = current.math_capture.as_mut() {
                        math.end(depth, part)?;
                    }
                    if let Some(capture) = current.paragraph_style_capture.as_mut()
                        && capture
                            .spacing
                            .is_some_and(|(spacing_depth, _)| spacing_depth == depth)
                    {
                        capture.spacing = None;
                    }
                    if let Some(capture) = current.paragraph_style_capture.as_mut()
                        && capture.bullet_color_depth == Some(depth)
                    {
                        capture.bullet_color_depth = None;
                    }
                    if current
                        .paragraph_style_capture
                        .as_ref()
                        .is_some_and(|capture| capture.depth == depth)
                    {
                        commit_shape_paragraph_style(current);
                    }
                    if local == "p" && current.paragraph_pending {
                        begin_shape_paragraph(current, false);
                    }
                    if let Some(custom_geometry) = current.custom_geometry.as_mut() {
                        custom_geometry.end(local, depth);
                    }
                    if local == "custGeom"
                        && current
                            .custom_geometry
                            .as_ref()
                            .is_some_and(|geometry| geometry.depth == depth)
                    {
                        let custom_geometry = current.custom_geometry.take().ok_or_else(|| {
                            format_error(part, "custom geometry parser state ended unexpectedly")
                        })?;
                        if let Some(geometry) = custom_geometry.into_geometry(shape_bounds(current)) {
                            current.geometry = geometry;
                            current.preset = None;
                        } else if !unsupported_geometry_reported {
                            state.diagnostics.push(
                                Diagnostic::warning(
                                    DiagnosticCode::UnsupportedFeature,
                                    Phase::Render,
                                    Fidelity::Approximate,
                                    "PPTX custom or unsupported shape geometry uses a rectangle fallback",
                                )
                                .in_part(part),
                            );
                            unsupported_geometry_reported = true;
                        }
                    }
                    if local == "blipFill"
                        && current
                            .image_fill
                            .as_ref()
                            .is_some_and(|fill| fill.depth == depth)
                    {
                        let fill = current.image_fill.take().ok_or_else(|| {
                            format_error(part, "shape image-fill parser state ended unexpectedly")
                        })?;
                        let target = fill.target;
                        let paint = resolve_shape_image_fill(
                            fill,
                            part,
                            &relationship_map,
                            package,
                            state,
                            content_types,
                        )?;
                        if matches!(target, PaintTarget::Text) {
                            if let Some(run) = current.run.as_mut() { run.paint = Some(Box::new(paint)); }
                            else {
                                if let Some(capture) = current.paragraph_style_capture.as_mut() { capture.style.font_paint = Some(Box::new(paint.clone())); }
                                current.font_paint = Some(Box::new(paint));
                            }
                        } else { current.fill = paint; }

                    }
                    if let Some((_, capture)) = current.paint_capture.as_mut() {
                        capture.end(local);
                    }
                    if current
                        .paint_capture
                        .as_ref()
                        .is_some_and(|(_, capture)| capture.depth == depth)
                    {
                        let (target, capture) = current.paint_capture.take().ok_or_else(|| {
                            format_error(part, "shape paint parser state ended unexpectedly")
                        })?;
                        let bounds = resolved_shape_bounds(current, unit_index, state);
                        let paint = capture.into_paint(bounds);
                        match target {
                            PaintTarget::Fill => current.fill = paint,
                            PaintTarget::Stroke if current.run_properties_depth.is_some() => {
                                if let Some(run) = current.run.as_mut() { run.stroke = Some(Box::new(paint)); }
                            }
                            PaintTarget::Stroke => current.stroke = paint,
                            PaintTarget::Text => {
                                let glyph_paint = (!matches!(paint, Paint::Solid(_))).then(|| Box::new(paint.clone()));
                                if let Some(run) = current.run.as_mut() {
                                    run.paint = glyph_paint;
                                } else {
                                    if let Some(capture) = current.paragraph_style_capture.as_mut() { capture.style.font_paint = glyph_paint.clone(); }
                                    current.font_paint = glyph_paint;
                                }
                                if let Paint::Solid(color) = paint {
                                    if let Some(run) = current.run.as_mut() {
                                        run.color = color;
                                    } else {
                                        current.font_color = color;
                                        if let Some(capture) =
                                            current.paragraph_style_capture.as_mut()
                                        {
                                            capture.style.font_color = Some(color);
                                        }
                                    }
                                }
                            }
                            PaintTarget::Highlight => {
                                if let Paint::Solid(color) = paint {
                                    if let Some(run) = current.run.as_mut() {
                                        run.highlight = color;
                                    } else {
                                        current.highlight = color;
                                    }
                                }
                            }
                            PaintTarget::Extrusion => {
                                if let Paint::Solid(color) = paint {
                                    current
                                        .three_d
                                        .get_or_insert_with(ThreeDStyle::default)
                                        .extrusion_color = Some(color);
                                }
                            }
                            PaintTarget::Contour => {
                                if let Paint::Solid(color) = paint {
                                    current
                                        .three_d
                                        .get_or_insert_with(ThreeDStyle::default)
                                        .contour_color = Some(color);
                                }
                            }
                        }
                    }
                    if local == "ln" && current.line_depth == Some(depth) {
                        current.line_depth = None;
                    } else if local == "camera" && current.camera_3d_depth == Some(depth) {
                        current.camera_3d_depth = None;
                    } else if local == "lightRig" && current.light_3d_depth == Some(depth) {
                        current.light_3d_depth = None;
                    } else if local == "backdrop" && current.backdrop_3d_depth == Some(depth) {
                        current.backdrop_3d_depth = None;
                    } else if local == "scene3d" && current.scene_3d_depth == Some(depth) {
                        current.scene_3d_depth = None;
                    } else if local == "sp3d" && current.shape_3d_depth == Some(depth) {
                        current.shape_3d_depth = None;
                    } else if matches!(local, "outerShdw" | "prstShdw" | "innerShdw" | "glow")
                        && current.shadow_depth == Some(depth)
                    {
                        let radians = current.shadow_direction_degrees.to_radians();
                        let shadow = Shadow {
                            color: current.shadow_color.unwrap_or(0x0000_0080),
                            blur: current.shadow_blur,
                            offset_x: current.shadow_distance * radians.cos(),
                            offset_y: current.shadow_distance * radians.sin(),
                        };
                        match current.shadow_kind {
                            PptxShadowKind::Outer => {
                                if let Some(run) = current.run.as_mut()
                                    && current.run_properties_depth.is_some()
                                {
                                    run.shadow = Some(shadow);
                                } else if let Some(effect) = current.outer_shadow.as_mut() {
                                    effect.shadow = shadow;
                                } else {
                                    current.shadow = Some(shadow);
                                }
                            }
                            PptxShadowKind::Inner => {
                                if let Some(run) = current.run.as_mut()
                                    && current.run_properties_depth.is_some()
                                {
                                    run.inner_shadow = Some(shadow);
                                } else {
                                    current.inner_shadow = Some(shadow);
                                }
                            }
                            PptxShadowKind::Glow => {
                                let glow = Glow { color: shadow.color, radius: shadow.blur };
                                if let Some(run) = current.run.as_mut()
                                    && current.run_properties_depth.is_some()
                                { run.glow = Some(glow); }
                                else { current.glow = Some(glow); }
                            }
                        }
                        current.shadow_depth = None;
                    } else if local == "txXfrm" && current.text_transform_depth == Some(depth) {
                        current.text_transform_depth = None;
                    } else if local == "xfrm" && current.transform_depth == Some(depth) {
                        current.transform_depth = None;
                    } else if local == "spPr"
                        && current.shape_properties_depth == Some(depth)
                    {
                        current.shape_properties_depth = None;
                    } else if local == "bodyPr"
                        && current.text_body_properties_depth == Some(depth)
                    {
                        current.text_body_properties_depth = None;
                    }
                    if matches!(local, "rPr" | "defRPr")
                        && current.run_properties_depth == Some(depth)
                    {
                        current.run_properties_depth = None;
                        current.run_properties_target = None;
                    }
                    if matches!(local, "r" | "fld")
                        && current.run.as_ref().is_some_and(|run| run.depth == depth)
                    {
                        let run = current.run.take().ok_or_else(|| {
                            format_error(part, "text run parser state ended unexpectedly")
                        })?;
                        let effect = run.effect();
                        if let Some(index) = current.pending_bullet_style.take() {
                            current.runs[index].color = run.color;
                            current.runs[index].paint = run.paint.clone();
                            current.run_effects[index] = effect.clone();
                        }
                        let runs = run.finish();
                        current.run_effects.extend(std::iter::repeat_n(effect, runs.len()));
                        current.runs.extend(runs);
                    }
                    if local == "fillRef" && current.fill_reference_depth == Some(depth) {
                        current.fill_reference_depth = None;
                    } else if local == "lnRef" && current.line_reference_depth == Some(depth) {
                        current.line_reference_depth = None;
                    } else if local == "fontRef"
                        && current.font_reference_depth == Some(depth)
                    {
                        current.font_reference_depth = None;
                    }
                }
                if local == "cNvPr"
                    && let Some(current) = shape.as_mut()
                    && current.non_visual_depth == Some(depth)
                {
                    current.non_visual_depth = None;
                }
                if local == "xfrm"
                    && let Some(current) = picture.as_mut()
                    && current.transform_depth == Some(depth)
                {
                    current.transform_depth = None;
                }
                if local == "spPr"
                    && let Some(current) = picture.as_mut()
                    && current.shape_properties_depth == Some(depth)
                {
                    current.shape_properties_depth = None;
                }
                if local == "cNvPr"
                    && let Some(current) = picture.as_mut()
                    && current.non_visual_depth == Some(depth)
                {
                    current.non_visual_depth = None;
                }
                if let Some(current) = picture.as_mut() {
                    current.effects.end(depth);
                    current.geometry.end(local, depth);
                    if local == "clrFrom" && current.color_change_from_depth == Some(depth) {
                        current.color_change_from_depth = None;
                    } else if local == "clrTo" && current.color_change_to_depth == Some(depth) {
                        current.color_change_to_depth = None;
                    } else if local == "clrChange" && current.color_change_depth == Some(depth) {
                        current.color_change_depth = None;
                    } else if local == "blip" && current.blip_depth == Some(depth) {
                        current.blip_depth = None;
                    }
                }
                if let Some(current) = picture.as_mut() {
                    if local == "lnRef" && current.line_reference_depth == Some(depth) {
                        current.line_reference_depth = None;
                    }
                    if current.is_background {
                        if let Some(capture) = current.fill_capture.as_mut() {
                            capture.end(local);
                        }
                        if current
                            .fill_capture
                            .as_ref()
                            .is_some_and(|capture| capture.depth == depth)
                        {
                            let capture = current.fill_capture.take().ok_or_else(|| {
                                format_error(
                                    part,
                                    "background fill parser state ended unexpectedly",
                                )
                            })?;
                            current.fill = capture.into_paint(slide_bounds);
                        }
                        if local == "bgRef" && current.fill_reference_depth == Some(depth) {
                            if matches!(current.fill, Paint::Image { .. }) {
                                current.background_fill_defined = true;
                            } else if let Some(index) = current.background_fill_reference {
                                if matches!(index, 0 | 1000) {
                                    current.fill = Paint::None;
                                    current.background_fill_defined = true;
                                }
                                let placeholder = match current.fill { Paint::Solid(color) => color, _ => 0x0000_00ff };
                                if let Some(fill) = resolve_pptx_theme_fill(package, theme, index, placeholder)? {
                                    current.fill = fill.paint(slide_bounds);
                                    current.background_fill_defined = true;
                                }
                            }
                            current.fill_reference_depth = None;
                        }
                    } else {
                        if let Some(capture) = current.fill_capture.as_mut() {
                            capture.end(local);
                        }
                        if current
                            .fill_capture
                            .as_ref()
                            .is_some_and(|capture| capture.depth == depth)
                        {
                            let capture = current.fill_capture.take().ok_or_else(|| {
                                format_error(part, "picture line parser state ended unexpectedly")
                            })?;
                            current.stroke = capture.into_paint(picture_shape_bounds(current));
                        }
                        if local == "ln" && current.line_depth == Some(depth) {
                            current.line_depth = None;
                        }
                    }
                }
                let finished = matches!(local, "sp" | "cxnSp" | "wsp")
                    && shape.as_ref().is_some_and(|current| current.depth == depth);
                if finished {
                    let current = shape.take().ok_or_else(|| {
                        format_error(part, "shape parser state ended unexpectedly")
                    })?;
                    let group_fill = groups
                        .iter()
                        .rev()
                        .find(|group| group.explicit_fill)
                        .map(|group| group.fill.clone());
                    push_shape(
                        current,
                        ShapePlacement {
                            part,
                            unit_index,
                            parent_numeric_id: groups.last().map(|group| group.numeric_id),
                            group_fill,
                            inherited: content.inherited,
                            object_limit: package.limits().max_document_objects,
                        },
                        state,
                        font_metrics,
                        &relationship_map,
                        package,
                        content_types,
                    )?;
                }
                let picture_finished = (local == "pic"
                    || (local == "bg"
                        && picture
                            .as_ref()
                            .is_some_and(|current| current.is_background)))
                    && picture
                        .as_ref()
                        .is_some_and(|current| current.depth == depth);
                if picture_finished {
                    let current = picture.take().ok_or_else(|| {
                        format_error(part, "picture parser state ended unexpectedly")
                    })?;
                    if current.is_background
                        && current.preferred_svg_relationship_id.is_none()
                        && current.embedded_relationship_id.is_none()
                        && current.linked_relationship_id.is_none()
                    {
                        if !matches!(current.fill, Paint::None) {
                            push_painted_background(
                                current,
                                part,
                                unit_index,
                                state,
                                package.limits().max_document_objects,
                            )?;
                        } else if !current.background_fill_defined && state
                            .reported_unsupported_backgrounds
                            .insert(part.to_owned())
                        {
                            state.diagnostics.push(
                                Diagnostic::warning(
                                    DiagnosticCode::UnsupportedFeature,
                                    Phase::Render,
                                    Fidelity::Omitted,
                                    "PPTX background fill could not be resolved",
                                )
                                .in_part(part),
                            );
                        }
                    } else {
                        push_picture(
                            current,
                            part,
                            unit_index,
                            &relationship_map,
                            package,
                            state,
                            content_types,
                        )?;
                    }
                }
                if let Some(current) = graphic_frame.as_mut() {
                    if local == "cNvPr" && current.non_visual_depth == Some(depth) {
                        current.non_visual_depth = None;
                    }
                    if let Some(cell) = current.cell.as_mut() {
                        if let Some((_, capture)) = cell.paint_capture.as_mut() {
                            capture.end(local);
                        }
                        if cell
                            .paint_capture
                            .as_ref()
                            .is_some_and(|(_, capture)| capture.depth == depth)
                        {
                            let (target, capture) = cell.paint_capture.take().ok_or_else(|| {
                                format_error(part, "table-cell paint parser state ended unexpectedly")
                            })?;
                            let paint = capture.into_paint(Rect::default());
                            match target {
                                CellPaintTarget::Fill => cell.fill = paint,
                                CellPaintTarget::Stroke(side) => {
                                    cell.border_mut(side).paint = paint;
                                }
                                CellPaintTarget::Text => {
                                    if let Paint::Solid(color) = paint {
                                        cell.explicit_text_color = true;
                                        if let Some(run) = cell.run.as_mut() {
                                            run.color = color;
                                            cell.run_style.font_color = Some(color);
                                        } else {
                                            cell.font_color = color;
                                        }
                                    }
                                }
                                CellPaintTarget::ParagraphText => {
                                    if let Paint::Solid(color) = paint {
                                        cell.explicit_text_color = true;
                                        cell.paragraph_style_capture
                                            .as_mut()
                                            .ok_or_else(|| {
                                                format_error(
                                                    part,
                                                    "table paragraph style capture is missing",
                                                )
                                            })?
                                            .style
                                            .font_color = Some(color);
                                    }
                                }
                                CellPaintTarget::Highlight => {
                                    if let Paint::Solid(color) = paint
                                        && let Some(run) = cell.run.as_mut()
                                    {
                                        run.highlight = color;
                                    }
                                }
                            }
                        }
                        if let Some(capture) = cell.paragraph_style_capture.as_mut()
                            && capture
                                .spacing
                                .is_some_and(|(spacing_depth, _)| spacing_depth == depth)
                        {
                            capture.spacing = None;
                        }
                        if let Some(capture) = cell.paragraph_style_capture.as_mut()
                            && capture.bullet_color_depth == Some(depth)
                        {
                            capture.bullet_color_depth = None;
                        }
                        if local == "defRPr"
                            && cell.paragraph_default_run_depth == Some(depth)
                        {
                            cell.paragraph_default_run_depth = None;
                        }
                        if cell
                            .paragraph_style_capture
                            .as_ref()
                            .is_some_and(|capture| capture.depth == depth)
                        {
                            commit_table_cell_paragraph_style(cell);
                        }
                        if local == "t" {
                            cell.collecting_text = false;
                        } else if matches!(local, "r" | "fld")
                            && cell.run.as_ref().is_some_and(|run| run.depth == depth)
                        {
                            let run = cell.run.take().ok_or_else(|| {
                                format_error(part, "table-cell text run parser state ended unexpectedly")
                            })?;
                            let effect = run.effect();
                            let runs = run.finish();
                            cell.run_effects
                                .extend(std::iter::repeat_n(effect, runs.len()));
                            cell.run_styles.extend(std::iter::repeat_n(cell.run_style.clone(), runs.len()));
                            cell.runs.extend(runs);
                        } else if matches!(local, "rPr" | "defRPr")
                            && cell.run_properties_depth == Some(depth)
                        {
                            cell.run_properties_depth = None;
                        } else if matches!(local, "lnL" | "lnR" | "lnT" | "lnB")
                            && cell
                                .border_depth
                                .is_some_and(|(border_depth, _)| border_depth == depth)
                        {
                            cell.border_depth = None;
                        } else if matches!(local, "lnTlToBr" | "lnBlToTr")
                            && cell.ignored_border_depth == Some(depth)
                        {
                            cell.ignored_border_depth = None;
                        } else if local == "tcPr" && cell.properties_depth == Some(depth) {
                            cell.properties_depth = None;
                        }
                        if local == "p" && cell.paragraph_pending {
                            begin_table_cell_paragraph(cell);
                        }
                    }
                    if local == "tc"
                        && current.cell.as_ref().is_some_and(|cell| cell.depth == depth)
                    {
                        let cell = current.cell.take().ok_or_else(|| {
                            format_error(part, "table-cell parser state ended unexpectedly")
                        })?;
                        current
                            .row
                            .as_mut()
                            .ok_or_else(|| format_error(part, "table cell has no row"))?
                            .cells
                            .push(cell);
                    } else if local == "tr"
                        && current.row.as_ref().is_some_and(|row| row.depth == depth)
                    {
                        let row = current.row.take().ok_or_else(|| {
                            format_error(part, "table-row parser state ended unexpectedly")
                        })?;
                        current.rows.push(row);
                    } else if local == "tbl" && current.table_depth == Some(depth) {
                        current.table_depth = None;
                    } else if local == "tableStyleId" {
                        current.collecting_table_style_id = false;
                    } else if local == "xfrm" && current.transform_depth == Some(depth) {
                        current.transform_depth = None;
                    }
                }
                let graphic_finished = local == "graphicFrame"
                    && graphic_frame
                        .as_ref()
                        .is_some_and(|current| current.depth == depth);
                if graphic_finished {
                    let current = graphic_frame.take().ok_or_else(|| {
                        format_error(part, "graphic-frame parser state ended unexpectedly")
                    })?;
                    finish_graphic_frame(
                        current,
                        part,
                        unit_index,
                        state,
                        package.limits().max_document_objects,
                        package,
                        content_types,
                        font_metrics,
                        theme,
                    )?;
                }
                if local == "xfrm"
                    && shape.is_none()
                    && picture.is_none()
                    && graphic_frame.is_none()
                    && let Some(current) = groups.last_mut()
                    && current.transform_depth == Some(depth)
                {
                    current.transform_depth = None;
                }
                if local == "cNvPr"
                    && shape.is_none()
                    && picture.is_none()
                    && graphic_frame.is_none()
                    && let Some(current) = groups.last_mut()
                    && current.non_visual_depth == Some(depth)
                {
                    current.non_visual_depth = None;
                }
                if shape.is_none()
                    && picture.is_none()
                    && graphic_frame.is_none()
                    && let Some(current) = groups.last_mut()
                {
                    if let Some(capture) = current.paint_capture.as_mut() {
                        capture.end(local);
                    }
                    if current
                        .paint_capture
                        .as_ref()
                        .is_some_and(|capture| capture.depth == depth)
                    {
                        let capture = current.paint_capture.take().ok_or_else(|| {
                            format_error(part, "group fill parser state ended unexpectedly")
                        })?;
                        current.fill = capture.into_paint(Rect {
                            x: 0.0,
                            y: 0.0,
                            width: current.width.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
                            height: current.height.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
                        });
                    }
                    if local == "grpSpPr" && current.shape_properties_depth == Some(depth) {
                        current.shape_properties_depth = None;
                    } else if local == "camera" && current.camera_3d_depth == Some(depth) {
                        current.camera_3d_depth = None;
                    } else if local == "lightRig" && current.light_3d_depth == Some(depth) {
                        current.light_3d_depth = None;
                    } else if local == "backdrop" && current.backdrop_3d_depth == Some(depth) {
                        current.backdrop_3d_depth = None;
                    } else if local == "scene3d" && current.scene_3d_depth == Some(depth) {
                        current.scene_3d_depth = None;
                    }
                }
                if matches!(local, "grpSp" | "lockedCanvas" | "wgp")
                    && groups.last().is_some_and(|current| current.depth == depth)
                {
                    let group = groups.pop().ok_or_else(|| {
                        format_error(part, "group parser state ended unexpectedly")
                    })?;
                    finish_group(group, part, state)?;
                }
            }
            XmlEvent::Text(text) => {
                if chart_choice.skipping() { return Ok(()); }
                if let Some(math) = shape
                    .as_mut()
                    .and_then(|current| current.math_capture.as_mut())
                {
                    math.text(
                        &decode_xml_text(text).map_err(|error| with_part(error, part))?,
                    );
                } else if let Some(current) =
                    shape.as_mut().filter(|current| current.collecting_text)
                {
                    let text = decode_xml_text(text).map_err(|error| with_part(error, part))?;
                    append_shape_text(current, &text);
                } else if let Some(cell) = graphic_frame
                    .as_mut()
                    .and_then(|current| current.cell.as_mut())
                    .filter(|cell| cell.collecting_text)
                {
                    let text = decode_xml_text(text).map_err(|error| with_part(error, part))?;
                    append_table_cell_text(cell, &text);
                } else if let Some(current) = graphic_frame
                    .as_mut()
                    .filter(|current| current.collecting_table_style_id)
                {
                    current.table_style_id.push_str(
                        &decode_xml_text(text).map_err(|error| with_part(error, part))?,
                    );
                }
            }
            XmlEvent::Cdata(text) => {
                if chart_choice.skipping() { return Ok(()); }
                if let Some(math) = shape
                    .as_mut()
                    .and_then(|current| current.math_capture.as_mut())
                {
                    math.text(text);
                } else if let Some(current) =
                    shape.as_mut().filter(|current| current.collecting_text)
                {
                    append_shape_text(current, text);
                } else if let Some(cell) = graphic_frame
                    .as_mut()
                    .and_then(|current| current.cell.as_mut())
                    .filter(|cell| cell.collecting_text)
                {
                    append_table_cell_text(cell, text);
                } else if let Some(current) = graphic_frame
                    .as_mut()
                    .filter(|current| current.collecting_table_style_id)
                {
                    current.table_style_id.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(())
}

fn push_painted_background(
    background: PictureState,
    part: &str,
    unit_index: u32,
    state: &mut PptxParseState,
    object_limit: usize,
) -> Result<(), Diagnostic> {
    if matches!(background.fill, Paint::None) {
        return Ok(());
    }
    if state.objects.len() >= object_limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "document exceeds the configured object limit",
        )
        .in_part(part));
    }
    let bounds = background
        .explicit_bounds
        .ok_or_else(|| format_error(part, "background is missing slide bounds"))?;
    let numeric_id = u32::try_from(state.objects.len())
        .map_err(|_| format_error(part, "object count exceeds supported range"))?;
    let z = state.take_z();
    let visual = Visual::PaintedShape {
        geometry: Geometry::Rectangle,
        fill: background.fill,
        stroke: Paint::None,
        stroke_width: 0.0,
    };
    let visual = if background.image_adjustment == ImageAdjustment::default() {
        visual
    } else {
        Visual::ImageAdjustment {
            adjustment: background.image_adjustment,
            visual: Box::new(visual),
        }
    };
    let visual = background.effects.finish().wrap(visual, None);
    state.objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: None,
        kind: ObjectKind::Shape,
        unit_index,
        bounds,
        z,
        text: None,
        source: SourceRef {
            part: part.to_owned(),
            mapping: MappingQuality::Derived,
            locator: SourceLocator::PptxShape {
                shape_id: numeric_id,
                row: None,
                column: None,
                text_range: None,
                metadata: PptxObjectMetadata::default(),
            },
        },
        visual,
    });
    Ok(())
}

fn finish_group(
    group: GroupState,
    part: &str,
    state: &mut PptxParseState,
) -> Result<(), Diagnostic> {
    let three_d = group.three_d;
    let target = Rect {
        x: group.x.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        y: group.y.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        width: group.width.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        height: group.height.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
    };
    let child_bounds = Rect {
        x: group.child_x.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        y: group.child_y.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        width: group.child_width.unwrap_or(group.width.unwrap_or(0)) as f32 / EMU_PER_CSS_PIXEL,
        height: group.child_height.unwrap_or(group.height.unwrap_or(0)) as f32 / EMU_PER_CSS_PIXEL,
    };
    if !target.is_valid() || !child_bounds.is_valid() {
        return Err(format_error(part, "group transform has invalid bounds"));
    }
    let scale_x = if child_bounds.width > 0.0 {
        target.width / child_bounds.width
    } else {
        1.0
    };
    let scale_y = if child_bounds.height > 0.0 {
        target.height / child_bounds.height
    } else {
        1.0
    };
    let child_to_target = AffineTransform {
        a: scale_x,
        b: 0.0,
        c: 0.0,
        d: scale_y,
        e: target.x - child_bounds.x * scale_x,
        f: target.y - child_bounds.y * scale_y,
    };
    let transform = shape_transform(
        target,
        group.rotation_degrees,
        group.flip_horizontal,
        group.flip_vertical,
    )
    .concat(child_to_target);
    let mapping = if group.shape_id.is_some() {
        MappingQuality::Exact
    } else {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Approximate,
                "group has no source ID; a session-local mapping was derived",
            )
            .in_part(part),
        );
        MappingQuality::Derived
    };
    let index = usize::try_from(group.numeric_id)
        .map_err(|_| format_error(part, "group object ID exceeds addressable range"))?;
    let object = state
        .objects
        .get_mut(index)
        .ok_or_else(|| format_error(part, "group object mapping is missing"))?;
    object.bounds = child_bounds;
    object.source.mapping = mapping;
    object.source.locator = SourceLocator::PptxShape {
        shape_id: group.shape_id.unwrap_or(group.numeric_id),
        row: None,
        column: None,
        text_range: None,
        metadata: group.metadata,
    };
    let visual = Visual::Layer {
        transform,
        opacity: 1.0,
        blend_mode: crate::model::BlendMode::Normal,
        visual: Box::new(Visual::None),
    };
    object.visual = if let Some(three_d) = three_d {
        Visual::AdvancedEffect {
            outer_shadow: None,
            inner_shadow: None,
            glow: None,
            reflection: None,
            soft_edge: None,
            three_d: Some(three_d),
            visual: Box::new(visual),
        }
    } else {
        visual
    };
    Ok(())
}

fn parse_basic_pptx_chart(
    package: &Package<'_>,
    part: &str,
    theme: &PptxTheme,
) -> Result<Option<BasicChart>, Diagnostic> {
    parse_drawingml_chart(package, part, theme.part.as_deref(), |value| {
        theme.color(value)
    })
}

pub(super) fn parse_ooxml_diagram_data(
    package: &Package<'_>,
    owner_part: &str,
    part: &str,
    colors_relationship_id: Option<&str>,
    bounds: Rect,
    scheme_color: impl Fn(&str) -> Option<u32>,
) -> Result<(Vec<Object>, Vec<Diagnostic>), Diagnostic> {
    let mut theme = PptxTheme::default();
    for (name, color) in &mut theme.colors {
        *color = scheme_color(name).unwrap_or(*color);
    }
    let owner_relationships = package.relationships(Some(owner_part))?;
    let owner_relationship_map = owner_relationships
        .iter()
        .map(|relationship| (relationship.id.as_str(), relationship))
        .collect::<HashMap<_, _>>();
    let colors_part = colors_relationship_id
        .and_then(|relationship_id| owner_relationship_map.get(relationship_id))
        .filter(|relationship| {
            !relationship.external && relationship.type_uri.ends_with("/diagramColors")
        })
        .map(|relationship| relationship.target.as_str());
    let diagram = parse_pptx_diagram(
        package,
        part,
        &owner_relationship_map,
        colors_part,
        theme.part.as_deref(),
        |value| theme.color(value),
    )?;
    let mut state = PptxParseState {
        objects: Vec::new(),
        next_z: 0,
        diagnostics: Vec::new(),
        image_cache: HashMap::new(),
        materialized_image_bytes: 0,
        reported_unsupported_backgrounds: HashSet::new(),
        placeholders: HashMap::new(),
        placeholder_text_styles: HashMap::new(),
        placeholder_presets: HashMap::new(),
        pending_placeholder_objects: HashSet::new(),
        default_theme: PptxTheme::default(),
        default_text_styles: vec![ParagraphStyle::default(); TEXT_LEVEL_COUNT],
        theme_cache: HashMap::new(),
    };
    if try_push_pptx_diagram_drawing(
        package,
        &diagram,
        bounds,
        0,
        &mut state,
        &ContentTypes::default(),
        &FontMetricTable::default(),
        &theme,
        &PptxObjectMetadata::default(),
    )? {
        return Ok((state.objects, state.diagnostics));
    }
    push_pptx_diagram_data(
        diagram,
        bounds,
        0,
        None,
        None,
        PptxObjectMetadata::default(),
        &mut state,
        package.limits().max_document_objects,
        &theme,
    )?;
    Ok((state.objects, state.diagnostics))
}

pub(super) fn parse_ooxml_locked_canvases(
    package: &Package<'_>,
    part: &str,
    font_metrics: &FontMetricTable,
) -> Result<(Vec<Vec<Object>>, Vec<Diagnostic>), Diagnostic> {
    let theme = package
        .relationships(Some(part))?
        .into_iter()
        .find(|relationship| !relationship.external && relationship.type_uri.ends_with("/theme"))
        .map(|relationship| parse_pptx_theme(package, &relationship.target))
        .transpose()?
        .unwrap_or_default();
    let mut state = PptxParseState {
        objects: Vec::new(),
        next_z: 0,
        diagnostics: Vec::new(),
        image_cache: HashMap::new(),
        materialized_image_bytes: 0,
        reported_unsupported_backgrounds: HashSet::new(),
        placeholders: HashMap::new(),
        placeholder_text_styles: HashMap::new(),
        placeholder_presets: HashMap::new(),
        pending_placeholder_objects: HashSet::new(),
        default_theme: theme.clone(),
        default_text_styles: vec![ParagraphStyle::default(); TEXT_LEVEL_COUNT],
        theme_cache: HashMap::new(),
    };
    parse_presentation_part(
        package,
        part,
        0,
        Rect::default(),
        PartContent {
            backgrounds: false,
            pictures: true,
            shapes: true,
            inherited: false,
        },
        &mut state,
        &ContentTypes::default(),
        font_metrics,
        &theme,
        None,
        None,
    )?;
    let mut roots = HashMap::new();
    for object in &state.objects {
        let root = object
            .parent_numeric_id
            .and_then(|parent| roots.get(&parent).copied())
            .unwrap_or(object.numeric_id);
        roots.insert(object.numeric_id, root);
    }
    let root_ids = state
        .objects
        .iter()
        .filter(|object| object.parent_numeric_id.is_none() && object.kind == ObjectKind::Group)
        .map(|object| object.numeric_id)
        .collect::<Vec<_>>();
    let canvases = root_ids
        .into_iter()
        .map(|root| {
            state
                .objects
                .iter()
                .filter(|object| roots.get(&object.numeric_id) == Some(&root))
                .cloned()
                .collect()
        })
        .collect();
    Ok((canvases, state.diagnostics))
}

#[allow(clippy::too_many_arguments)]
fn finish_graphic_frame(
    mut frame: GraphicFrameState,
    part: &str,
    unit_index: u32,
    state: &mut PptxParseState,
    object_limit: usize,
    package: &Package<'_>,
    content_types: &ContentTypes,
    font_metrics: &FontMetricTable,
    theme: &PptxTheme,
) -> Result<(), Diagnostic> {
    if let Some(diagram) = frame.diagram.take() {
        let bounds = Rect {
            x: frame.x.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
            y: frame.y.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
            width: frame.width.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
            height: frame.height.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        };
        if !bounds.is_valid() {
            return Err(format_error(part, "diagram frame has invalid bounds"));
        }
        if try_push_pptx_diagram_drawing(
            package,
            &diagram,
            bounds,
            unit_index,
            state,
            content_types,
            font_metrics,
            theme,
            &frame.metadata,
        )? {
            return Ok(());
        }
        return push_pptx_diagram_data(
            diagram,
            bounds,
            unit_index,
            frame.parent_numeric_id,
            frame.shape_id,
            frame.metadata,
            state,
            object_limit,
            theme,
        );
    }
    if !frame.table_style_id.trim().is_empty() {
        let style_id = frame.table_style_id.trim();
        let bounds = Rect {
            x: frame.x.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
            y: frame.y.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
            width: frame.width.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
            height: frame.height.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        };
        let style = parse_pptx_table_style(package, style_id, theme, bounds)?;
        if let Some(mut style) = style {
            state.diagnostics.append(&mut style.diagnostics);
            for (region, source_part, image) in std::mem::take(&mut style.image_fills) {
                let resolved =
                    package
                        .relationships(Some(&source_part))
                        .and_then(|relationships| {
                            let map = relationships.iter().map(|r| (r.id.as_str(), r)).collect();
                            resolve_shape_image_fill(
                                image,
                                &source_part,
                                &map,
                                package,
                                state,
                                content_types,
                            )
                        });
                match resolved {
                    Ok(paint) => {
                        if let Some(region) = region {
                            style.parts.entry(region).or_default().fill =
                                Some(super::drawingml::ChartFill::Image(paint));
                        } else {
                            style.background_fill = Some(super::drawingml::ChartFill::Image(paint));
                        }
                    }
                    Err(error)
                        if matches!(
                            error.code,
                            DiagnosticCode::ObjectLimit
                                | DiagnosticCode::ZipTotalSizeLimit
                                | DiagnosticCode::ZipEntryLimit
                                | DiagnosticCode::ImageDimensionLimit
                        ) =>
                    {
                        return Err(error);
                    }
                    Err(error) => state.diagnostics.push(
                        Diagnostic::warning(
                            DiagnosticCode::FormatInvalid,
                            Phase::Parse,
                            Fidelity::Omitted,
                            "table image fill could not be decoded",
                        )
                        .in_part(&source_part)
                        .with_detail("cause", error.message),
                    ),
                }
            }
            frame.table_effects = style.background_effects.clone();
            frame.table_fill = style.background_fill.clone();
            let row_count = frame.rows.len();
            let column_count = frame.columns.len().max(
                frame
                    .rows
                    .iter()
                    .map(|row| {
                        row.cells
                            .iter()
                            .filter(|cell| !cell.horizontal_merge)
                            .map(|cell| cell.grid_span.max(1))
                            .sum()
                    })
                    .max()
                    .unwrap_or(0),
            );
            for (row_index, row) in frame.rows.iter_mut().enumerate() {
                let mut logical_column = 0_usize;
                for cell in &mut row.cells {
                    if cell.horizontal_merge {
                        continue;
                    }
                    let column_index = logical_column;
                    logical_column = logical_column.saturating_add(cell.grid_span.max(1));
                    let style_context = TableCellStyleContext {
                        row: row_index,
                        column: column_index,
                        row_span: cell.row_span,
                        column_span: cell.grid_span,
                        row_count,
                        column_count,
                        first_row: frame.first_row,
                        last_row: frame.last_row,
                        first_column: frame.first_column,
                        last_column: frame.last_column,
                        band_rows: frame.band_rows,
                        band_columns: frame.band_columns,
                    };
                    for side in TableBorderSide::ALL {
                        if !cell.border(side).explicit {
                            let styled = style.cell_border(side, style_context);
                            let border = cell.border_mut(side);
                            border.paint = styled.paint.clone();
                            border.width = styled.width;
                            border.dash_pattern = styled.dash_pattern;
                            border.stroke_style = styled.stroke_style.clone();
                        }
                    }
                    if !cell.explicit_fill
                        && let Some(fill) = style.cell_fill(style_context)
                    {
                        if let super::drawingml::ChartFill::Image(Paint::Image { bytes, .. }) = fill
                        {
                            reserve_materialized_image_bytes(
                                &mut state.materialized_image_bytes,
                                bytes.len(),
                                package.limits().max_total_uncompressed_bytes,
                                TABLE_STYLES_PART,
                            )?;
                        }
                        cell.style_fill = Some(fill.clone());
                    }
                    cell.effects = style.cell_effects(style_context).cloned();
                    let text = style.cell_text_style(style_context);
                    if let Some(color) = text.font_color {
                        cell.font_color = color;
                    }
                    if let Some(bold) = text.bold {
                        cell.bold = bold;
                    }
                    if let Some(italic) = text.italic {
                        cell.italic = italic;
                    }
                    if let Some(font) = &text.font_family {
                        cell.font_family.clone_from(font);
                    }
                    for (run, local) in cell.runs.iter_mut().zip(&cell.run_styles) {
                        if local.font_color.is_none()
                            && let Some(color) = text.font_color
                        {
                            run.color = color;
                        }
                        if local.bold.is_none()
                            && let Some(bold) = text.bold
                        {
                            run.bold = bold;
                        }
                        if local.italic.is_none()
                            && let Some(italic) = text.italic
                        {
                            run.italic = italic;
                        }
                        let class = run
                            .text
                            .chars()
                            .find(|c| !c.is_whitespace())
                            .map(drawingml_font_class);
                        let (declared, inherited) = match class {
                            Some(DrawingMlFontClass::EastAsian) => {
                                (&local.font_east_asian, &text.font_east_asian)
                            }
                            Some(DrawingMlFontClass::ComplexScript) => {
                                (&local.font_complex_script, &text.font_complex_script)
                            }
                            _ => (&local.font_family, &text.font_family),
                        };
                        if declared.is_none()
                            && let Some(font) = inherited.as_ref().or(text.font_family.as_ref())
                        {
                            run.font_family.clone_from(font);
                        }
                    }
                }
            }
        }
    }
    push_graphic_frame(frame, part, unit_index, state, object_limit, font_metrics)
}

fn parse_pptx_table_style(
    package: &Package<'_>,
    style_id: &str,
    theme: &PptxTheme,
    bounds: Rect,
) -> Result<Option<PptxTableStyle>, Diagnostic> {
    if let Some(bytes) = package.part(TABLE_STYLES_PART)?
        && let Some(style) = parse_pptx_table_style_xml(&bytes, package, style_id, theme, bounds)?
    {
        return Ok(Some(style));
    }
    let builtins = include_str!("pptx_table_styles.xml");
    let key = format!("<a:tblStyle styleId=\"{style_id}\"");
    let Some(start) = builtins.find(&key) else {
        return Ok(None);
    };
    let definition = &builtins[start..];
    let end = definition
        .find("</a:tblStyle>")
        .expect("bundled table styles are complete")
        + "</a:tblStyle>".len();
    parse_pptx_table_style_xml(
        definition[..end].as_bytes(),
        package,
        style_id,
        theme,
        bounds,
    )
}

fn parse_pptx_table_style_xml(
    bytes: &[u8],
    package: &Package<'_>,
    style_id: &str,
    theme: &PptxTheme,
    bounds: Rect,
) -> Result<Option<PptxTableStyle>, Diagnostic> {
    use TableStyleRegion as Region;
    #[derive(Debug)]
    enum PaintTarget {
        Stroke(TableStyleBorderSide),
        Fill(Region),
        Background,
        Text(Region),
        FillReference(u64, Option<Region>),
        EffectReference(u64, Option<Region>),
    }

    let mut depth = 0_usize;
    let mut matched_style = false;
    let mut style_depth = None;
    let mut region = None;
    let mut cell_style_depth = None;
    let mut table_background_depth = None;
    let mut cell_fill_depth = None;
    let mut text_style_depth = None;
    let mut borders_depth = None;
    let mut side_depth: Option<(usize, TableStyleBorderSide)> = None;
    let mut line_depth: Option<(usize, TableStyleBorderSide)> = None;
    let empty_border = TableCellBorder {
        paint: Paint::None,
        ..TableCellBorder::default()
    };
    let mut parts = HashMap::<Region, TablePartStyle>::new();
    let mut borders = std::array::from_fn::<_, 6, _>(|_| empty_border.clone());
    let mut background_fill = None;
    let mut background_effects = None;
    let mut image_fills = Vec::new();
    let mut image_capture: Option<(Option<Region>, ShapeImageFillState)> = None;
    let mut effects_capture: Option<(usize, Option<Region>, DrawingMlPictureEffectsCapture)> = None;
    let mut diagnostics = Vec::new();
    let mut failed_image_depth = None;
    let mut paint_capture: Option<(PaintTarget, PaintCapture)> = None;
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if failed_image_depth.is_some() {
                    // Ignore only the malformed image subtree; retain the rest of the table.
                } else if let Some((_, image)) = image_capture.as_mut() {
                    if let Err(error) = image.start(local, &attributes, TABLE_STYLES_PART) {
                        failed_image_depth = Some(image.depth);
                        image_capture = None;
                        diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::FormatInvalid,
                                Phase::Parse,
                                Fidelity::Omitted,
                                "invalid table image fill",
                            )
                            .in_part(TABLE_STYLES_PART)
                            .with_detail("cause", error.message),
                        );
                    }
                } else if let Some((_, _, effects)) = effects_capture.as_mut() {
                    effects.start(
                        if local == "bevel" { "bevelT" } else { local },
                        &attributes,
                        empty,
                        depth,
                        TABLE_STYLES_PART,
                        |kind, value| picture_effect_color(kind, value, theme, TABLE_STYLES_PART),
                    )?;
                } else if (cell_style_depth.is_some() || table_background_depth.is_some())
                    && local == "blipFill"
                {
                    if !empty {
                        image_capture =
                            Some((region.map(|(_, r)| r), ShapeImageFillState::new(depth)));
                    }
                } else if (cell_style_depth.is_some() || table_background_depth.is_some())
                    && matches!(local, "effectLst" | "effectDag" | "cell3D")
                {
                    if !empty {
                        let mut capture = DrawingMlPictureEffectsCapture::default();
                        // CT_Cell3D combines the shared shape material/bevel and scene lighting.
                        if local == "cell3D" {
                            capture.start(
                                "scene3d",
                                &[],
                                false,
                                depth,
                                TABLE_STYLES_PART,
                                |kind, value| {
                                    picture_effect_color(kind, value, theme, TABLE_STYLES_PART)
                                },
                            )?;
                        }
                        capture.start(
                            if local == "cell3D" { "sp3d" } else { local },
                            &attributes,
                            empty,
                            depth,
                            TABLE_STYLES_PART,
                            |kind, value| {
                                picture_effect_color(kind, value, theme, TABLE_STYLES_PART)
                            },
                        )?;
                        effects_capture = Some((depth, region.map(|(_, r)| r), capture));
                    }
                } else if let Some((_, capture)) = paint_capture.as_mut() {
                    capture.start(local, &attributes, TABLE_STYLES_PART, theme)?;
                } else if let Some((_, side)) = line_depth
                    && matches!(local, "solidFill" | "gradFill" | "pattFill")
                {
                    if let Some(capture) =
                        PaintCapture::new(local, &attributes, TABLE_STYLES_PART, depth)?
                    {
                        if empty {
                            borders[side.index()].paint = capture.into_paint(Rect::default());
                        } else {
                            paint_capture = Some((PaintTarget::Stroke(side), capture));
                        }
                    }
                } else if let Some((_, side)) = line_depth
                    && local == "noFill"
                {
                    borders[side.index()].paint = Paint::None;
                } else if let Some((_, side)) = line_depth
                    && matches!(
                        local,
                        "srgbClr" | "scrgbClr" | "schemeClr" | "sysClr" | "prstClr"
                    )
                    && let Some(mut capture) =
                        PaintCapture::new("solidFill", &[], TABLE_STYLES_PART, depth)?
                {
                    capture.start(local, &attributes, TABLE_STYLES_PART, theme)?;
                    if empty {
                        borders[side.index()].paint = capture.into_paint(Rect::default());
                    } else {
                        paint_capture = Some((PaintTarget::Stroke(side), capture));
                    }
                } else if table_background_depth.is_some()
                    && matches!(local, "solidFill" | "gradFill" | "pattFill" | "noFill")
                {
                    if local == "noFill" {
                        background_fill = Some(super::drawingml::ChartFill::Image(Paint::None));
                    } else if let Some(capture) =
                        PaintCapture::new(local, &attributes, TABLE_STYLES_PART, depth)?
                    {
                        if empty {
                            background_fill = Some(capture.into_fill());
                        } else {
                            paint_capture = Some((PaintTarget::Background, capture));
                        }
                    }
                } else if cell_fill_depth.is_some()
                    && matches!(local, "solidFill" | "gradFill" | "pattFill")
                {
                    if let (Some((_, current_region)), Some(capture)) = (
                        region,
                        PaintCapture::new(local, &attributes, TABLE_STYLES_PART, depth)?,
                    ) {
                        if empty {
                            parts.entry(current_region).or_default().fill =
                                Some(capture.into_fill());
                        } else {
                            paint_capture = Some((PaintTarget::Fill(current_region), capture));
                        }
                    }
                } else if cell_fill_depth.is_some() && local == "noFill" {
                    if let Some((_, current_region)) = region {
                        parts.entry(current_region).or_default().fill =
                            Some(super::drawingml::ChartFill::Image(Paint::None));
                    }
                } else if text_style_depth.is_some_and(|start| depth == start + 1)
                    && matches!(
                        local,
                        "srgbClr" | "scrgbClr" | "schemeClr" | "sysClr" | "prstClr"
                    )
                    && let Some((_, current_region)) = region
                    && let Some(mut capture) =
                        PaintCapture::new("solidFill", &[], TABLE_STYLES_PART, depth)?
                {
                    capture.start(local, &attributes, TABLE_STYLES_PART, theme)?;
                    if empty {
                        if let Paint::Solid(color) = capture.into_paint(Rect::default()) {
                            parts.entry(current_region).or_default().text.font_color = Some(color);
                        }
                    } else {
                        paint_capture = Some((PaintTarget::Text(current_region), capture));
                    }
                } else if text_style_depth.is_some() && local == "fontRef" {
                    let font = string_attribute(&attributes, "idx", TABLE_STYLES_PART)?;
                    if matches!(font.as_deref(), Some("major" | "minor")) {
                        let prefix = if font.as_deref() == Some("major") {
                            "mj"
                        } else {
                            "mn"
                        };
                        let text = &mut parts.entry(region.unwrap().1).or_default().text;
                        text.font_family =
                            Some(theme.resolve_typeface(&format!("+{prefix}-lt"), "latin"));
                        text.font_east_asian =
                            Some(theme.resolve_typeface(&format!("+{prefix}-ea"), "ea"));
                        text.font_complex_script =
                            Some(theme.resolve_typeface(&format!("+{prefix}-cs"), "cs"));
                    }
                } else if text_style_depth.is_some() && matches!(local, "latin" | "ea" | "cs") {
                    if let Some(font) =
                        string_attribute(&attributes, "typeface", TABLE_STYLES_PART)?
                            .filter(|font| !font.is_empty())
                    {
                        let font = Some(theme.resolve_typeface(&font, local));
                        let text = &mut parts.entry(region.unwrap().1).or_default().text;
                        match local {
                            "ea" => text.font_east_asian = font,
                            "cs" => text.font_complex_script = font,
                            _ => text.font_family = font,
                        }
                    }
                } else if local == "tblStyle"
                    && string_attribute(&attributes, "styleId", TABLE_STYLES_PART)?.as_deref()
                        == Some(style_id)
                {
                    matched_style = true;
                    style_depth = (!empty).then_some(depth);
                } else if style_depth.is_some() && local == "tblBg" {
                    table_background_depth = (!empty).then_some(depth);
                } else if (table_background_depth.is_some() || cell_style_depth.is_some())
                    && matches!(local, "fillRef" | "effectRef")
                    && let Some(index) = numeric_attribute(&attributes, "idx", TABLE_STYLES_PART)?
                {
                    if local == "fillRef" && matches!(index, 0 | 1000) {
                        let fill = Some(super::drawingml::ChartFill::Image(Paint::None));
                        if let Some((_, region)) = region {
                            parts.entry(region).or_default().fill = fill;
                        } else {
                            background_fill = fill;
                        }
                    } else if local == "effectRef" && index == 0 {
                        let effects = Some(super::drawingml::DrawingMlPictureEffects::default());
                        if let Some((_, region)) = region {
                            parts.entry(region).or_default().effects = effects;
                        } else {
                            background_effects = effects;
                        }
                    } else if !empty
                        && let Some(capture) =
                            PaintCapture::new("solidFill", &[], TABLE_STYLES_PART, depth)?
                    {
                        paint_capture = Some((
                            if local == "fillRef" {
                                PaintTarget::FillReference(
                                    index,
                                    region.map(|(_, current)| current),
                                )
                            } else {
                                PaintTarget::EffectReference(
                                    index,
                                    region.map(|(_, current)| current),
                                )
                            },
                            capture,
                        ));
                    }
                } else if style_depth.is_some()
                    && let Some(current_region) = Region::from_element(local)
                {
                    region = (!empty).then_some((depth, current_region));
                } else if region.is_some() && local == "tcStyle" {
                    cell_style_depth = (!empty).then_some(depth);
                } else if region.is_some() && local == "tcTxStyle" {
                    text_style_depth = (!empty).then_some(depth);
                    let text = &mut parts.entry(region.unwrap().1).or_default().text;
                    text.bold = boolean_attribute(&attributes, "b", TABLE_STYLES_PART)?;
                    text.italic = boolean_attribute(&attributes, "i", TABLE_STYLES_PART)?;
                } else if cell_style_depth.is_some() && local == "fill" {
                    cell_fill_depth = (!empty).then_some(depth);
                } else if cell_style_depth.is_some() && local == "tcBdr" {
                    borders_depth = (!empty).then_some(depth);
                } else if borders_depth.is_some()
                    && let Some(side) = TableStyleBorderSide::from_element(local)
                {
                    side_depth = (!empty).then_some((depth, side));
                } else if let Some((_, side)) = side_depth
                    && matches!(local, "ln" | "lnRef")
                {
                    let border = &mut borders[side.index()];
                    *border = empty_border.clone();
                    border.width = if local == "lnRef" {
                        numeric_attribute(&attributes, "idx", TABLE_STYLES_PART)?
                            .and_then(|index| usize::try_from(index).ok())
                            .and_then(|index| theme.line_width(index))
                            .unwrap_or(1.0)
                    } else {
                        numeric_attribute(&attributes, "w", TABLE_STYLES_PART)?.unwrap_or(12_700)
                            as f32
                            / EMU_PER_CSS_PIXEL
                    };
                    border.dash_pattern = DrawingMlDashPattern::Solid;
                    border.stroke_style = drawingml_stroke_style(&attributes, TABLE_STYLES_PART)?.1;
                    line_depth = (!empty).then_some((depth, side));
                    if empty {
                        parts.entry(region.unwrap().1).or_default().borders[side.index()] =
                            Some(border.clone());
                    }
                } else if let Some((_, side)) = line_depth
                    && local == "prstDash"
                {
                    borders[side.index()].dash_pattern = DrawingMlDashPattern::from_attribute(
                        string_attribute(&attributes, "val", TABLE_STYLES_PART)?.as_deref(),
                    );
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if let Some((_, _, effects)) = effects_capture.as_mut() {
                    effects.end(depth);
                }
                if effects_capture
                    .as_ref()
                    .is_some_and(|(start, _, _)| *start == depth)
                {
                    let (_, target, effects) = effects_capture.take().unwrap();
                    if let Some(target) = target {
                        parts.entry(target).or_default().effects = Some(effects.finish());
                    } else {
                        background_effects = Some(effects.finish());
                    }
                }
                if failed_image_depth == Some(depth) {
                    failed_image_depth = None;
                }
                if image_capture
                    .as_ref()
                    .is_some_and(|(_, image)| image.depth == depth)
                {
                    let (target, image) = image_capture.take().unwrap();
                    image_fills.push((target, TABLE_STYLES_PART.to_owned(), image));
                }
                if let Some((_, capture)) = paint_capture.as_mut() {
                    capture.end(local);
                }
                if paint_capture
                    .as_ref()
                    .is_some_and(|(_, capture)| capture.depth == depth)
                {
                    let (target, capture) = paint_capture.take().ok_or_else(|| {
                        format_error(
                            TABLE_STYLES_PART,
                            "table-style paint parser state ended unexpectedly",
                        )
                    })?;
                    let fill = capture.into_fill();
                    let paint = fill.paint(bounds);
                    match target {
                        PaintTarget::Stroke(side) => borders[side.index()].paint = paint,
                        PaintTarget::Fill(current_region) => {
                            parts.entry(current_region).or_default().fill = Some(fill);
                        }
                        PaintTarget::Background => background_fill = Some(fill),
                        PaintTarget::Text(current_region) => {
                            if let Paint::Solid(color) = paint {
                                parts.entry(current_region).or_default().text.font_color =
                                    Some(color);
                            }
                        }
                        PaintTarget::EffectReference(index, target) => {
                            if let Paint::Solid(color) = paint {
                                let effects =
                                    resolve_pptx_theme_effects(package, theme, index, color)?;
                                if let Some(target) = target {
                                    parts.entry(target).or_default().effects = effects;
                                } else {
                                    background_effects = effects;
                                }
                            }
                        }
                        PaintTarget::FillReference(index, current_region) => {
                            if let Paint::Solid(placeholder) = paint
                                && let Some(fill) =
                                    resolve_pptx_theme_fill(package, theme, index, placeholder)?
                            {
                                if let Some(current_region) = current_region {
                                    parts.entry(current_region).or_default().fill = Some(fill);
                                } else {
                                    background_fill = Some(fill);
                                }
                            }
                            match resolve_pptx_theme_image_fill(package, theme, index) {
                                Ok(Some(image)) => image_fills.push((
                                    current_region,
                                    theme.part.clone().unwrap_or_default(),
                                    image,
                                )),
                                Ok(None) => {}
                                Err(error) if error.code == DiagnosticCode::FormatInvalid => {
                                    diagnostics.push(
                                        Diagnostic::warning(
                                            error.code,
                                            Phase::Parse,
                                            Fidelity::Omitted,
                                            "invalid theme table image fill",
                                        )
                                        .in_part(TABLE_STYLES_PART)
                                        .with_detail("cause", error.message),
                                    )
                                }
                                Err(error) => return Err(error),
                            }
                        }
                    }
                }
                if matches!(local, "ln" | "lnRef")
                    && line_depth.is_some_and(|(line_start, _)| line_start == depth)
                {
                    if let Some((_, current_region)) = region {
                        parts.entry(current_region).or_default().borders
                            [line_depth.unwrap().1.index()] =
                            Some(borders[line_depth.unwrap().1.index()].clone());
                    }
                    line_depth = None;
                } else if TableStyleBorderSide::from_element(local).is_some()
                    && side_depth.is_some_and(|(side_start, _)| side_start == depth)
                {
                    side_depth = None;
                } else if local == "tcBdr" && borders_depth == Some(depth) {
                    borders_depth = None;
                } else if local == "fill" && cell_fill_depth == Some(depth) {
                    cell_fill_depth = None;
                } else if local == "tcStyle" && cell_style_depth == Some(depth) {
                    cell_style_depth = None;
                } else if local == "tcTxStyle" && text_style_depth == Some(depth) {
                    text_style_depth = None;
                } else if local == "tblBg" && table_background_depth == Some(depth) {
                    table_background_depth = None;
                } else if region.is_some_and(|(start, _)| start == depth)
                    && Region::from_element(local).is_some()
                {
                    region = None;
                } else if local == "tblStyle" && style_depth == Some(depth) {
                    style_depth = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, TABLE_STYLES_PART))?;
    if !matched_style {
        return Ok(None);
    }
    Ok(Some(PptxTableStyle {
        parts,
        background_fill,
        background_effects,
        image_fills,
        diagnostics,
        empty_border,
    }))
}
#[cfg(test)]
fn built_in_pptx_table_style(style_id: &str, theme: &PptxTheme) -> Option<PptxTableStyle> {
    let bytes = super::presentation_image::stored_zip(&[]);
    let package = Package::open(&bytes, crate::limits::Limits::default()).ok()?;
    parse_pptx_table_style_xml(
        include_bytes!("pptx_table_styles.xml"),
        &package,
        style_id,
        theme,
        Rect::default(),
    )
    .ok()
    .flatten()
}

#[allow(clippy::too_many_arguments)]
fn cycle_arrow_commands(index: usize, width: f32, height: f32) -> Vec<PathCommand> {
    let band = width.min(height) * 0.12;
    let head = band * 1.25;
    let k = 0.552_284_8;
    let inner_width = (width - band).max(0.0);
    let inner_height = (height - band).max(0.0);
    let transform = |x: f32, y: f32| match index % 4 {
        0 => (x, y),
        1 => (width - y, x),
        2 => (width - x, height - y),
        _ => (y, height - x),
    };
    let point = |x, y| {
        let (x, y) = transform(x, y);
        PathCommand::LineTo { x, y }
    };
    let (start_x, start_y) = transform(0.0, 0.0);
    let (cp1x, cp1y) = transform(width * k, 0.0);
    let (cp2x, cp2y) = transform(width, height * (1.0 - k));
    let (end_x, end_y) = transform(width, height);
    let (inner_cp1x, inner_cp1y) = transform(inner_width, band + inner_height * (1.0 - k));
    let (inner_cp2x, inner_cp2y) = transform(band + inner_width * k, band);
    let (inner_end_x, inner_end_y) = transform(band, band);
    vec![
        PathCommand::MoveTo {
            x: start_x,
            y: start_y,
        },
        PathCommand::BezierCurveTo {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x: end_x,
            y: end_y,
        },
        point(width + head, height),
        point(width - band, height + head),
        point(inner_width, inner_height),
        PathCommand::BezierCurveTo {
            cp1x: inner_cp1x,
            cp1y: inner_cp1y,
            cp2x: inner_cp2x,
            cp2y: inner_cp2y,
            x: inner_end_x,
            y: inner_end_y,
        },
        PathCommand::ClosePath,
    ]
}

fn push_pptx_diagram_data(
    diagram: PptxDiagram,
    bounds: Rect,
    unit_index: u32,
    parent_numeric_id: Option<u32>,
    shape_id: Option<u32>,
    metadata: PptxObjectMetadata,
    state: &mut PptxParseState,
    object_limit: usize,
    theme: &PptxTheme,
) -> Result<(), Diagnostic> {
    state
        .diagnostics
        .extend(diagram.diagnostics.iter().cloned());
    if diagram.nodes.is_empty() {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Omitted,
                "PPTX diagram has neither a usable drawing fallback nor visible data nodes",
            )
            .in_part(&diagram.data_part),
        );
        return Ok(());
    }
    let edge_count = diagram
        .nodes
        .iter()
        .filter(|node| node.parent_id.is_some())
        .count();
    let accent = theme.color("accent1").unwrap_or(0x4f81_bdff);
    let fallback_elements = super::drawingml::diagram_fallback_elements(&diagram, bounds, accent);
    let required = state
        .objects
        .len()
        .checked_add(
            fallback_elements
                .as_ref()
                .map_or_else(|| diagram.nodes.len().saturating_mul(2), Vec::len),
        )
        .and_then(|count| count.checked_add(edge_count))
        .ok_or_else(|| format_error(&diagram.data_part, "diagram object count overflow"))?;
    if required > object_limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "diagram nodes exceed the configured object limit",
        )
        .in_part(&diagram.data_part));
    }
    state.diagnostics.push(
        Diagnostic::warning(
            DiagnosticCode::ApproximateLayout,
            Phase::Layout,
            Fidelity::Approximate,
            "PPTX diagram data uses an approximate family layout because no generated DrawingML geometry exists",
        )
        .in_part(&diagram.data_part),
    );
    if let Some(fill) = diagram.background_fill.as_ref() {
        let numeric_id = u32::try_from(state.objects.len())
            .map_err(|_| format_error(&diagram.data_part, "diagram object count exceeds range"))?;
        let z = state.take_z();
        let visual = Visual::PaintedShape {
            geometry: Geometry::RoundedRectangle {
                radius_x: 6.0,
                radius_y: 6.0,
            },
            fill: if diagram.scene_three_d {
                match fill {
                    super::drawingml::ChartFill::Solid(color) => {
                        let color = *color;
                        let alpha = (color & 0xff) as f32 / 255.0;
                        let channel = |shift: u32, light: f32| {
                            (((((color >> shift) & 0xff_u32) as f32 * alpha
                                + 255.0 * (1.0 - alpha))
                                * light)
                                .round() as u32)
                                .min(255)
                        };
                        Paint::Solid(
                            (channel(24, 0.909_09) << 24)
                                | (channel(16, 0.9) << 16)
                                | (channel(8, 0.895_45) << 8)
                                | 0xff,
                        )
                    }
                    _ => fill.paint(bounds),
                }
            } else {
                fill.paint(bounds)
            },
            stroke: Paint::None,
            stroke_width: 0.0,
        };
        let background_shadow = diagram.background_shadow.map(|mut shadow| {
            if diagram.scene_three_d
                && let super::drawingml::ChartFill::Solid(color) = fill
            {
                let fill_alpha = (color & 0xff) as f32 / 255.0;
                let shadow_alpha = (shadow.color & 0xff) as f32 / 255.0;
                shadow.color = (shadow.color & 0xffff_ff00)
                    | (fill_alpha * shadow_alpha * 255.0).round() as u32;
            }
            shadow
        });
        let visual = if let Some(shadow) = background_shadow {
            Visual::Effect {
                shadow: Some(shadow),
                clip: None,
                visual: Box::new(visual),
            }
        } else {
            visual
        };
        state.objects.push(Object {
            numeric_id,
            parent_numeric_id,
            stable_id: format!("object:{numeric_id}"),
            parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
            kind: ObjectKind::Shape,
            unit_index,
            bounds,
            z,
            text: None,
            source: SourceRef {
                part: diagram.data_part.clone(),
                mapping: MappingQuality::Exact,
                locator: SourceLocator::PptxShape {
                    shape_id: shape_id.unwrap_or(numeric_id),
                    row: None,
                    column: None,
                    text_range: None,
                    metadata: metadata.clone(),
                },
            },
            visual,
        });
    }
    if let Some(elements) = fallback_elements {
        for (index, element) in elements.into_iter().enumerate() {
            let numeric_id = u32::try_from(state.objects.len()).map_err(|_| {
                format_error(&diagram.data_part, "diagram object count exceeds range")
            })?;
            let text_length = element.text.as_deref().map_or(0, |text| {
                u32::try_from(text.chars().count()).unwrap_or(u32::MAX)
            });
            let z = state.take_z();
            state.objects.push(Object {
                numeric_id,
                parent_numeric_id,
                stable_id: format!("object:{numeric_id}"),
                parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
                kind: if element.text.is_some() {
                    ObjectKind::TextBox
                } else {
                    ObjectKind::Shape
                },
                unit_index,
                bounds: element.bounds,
                z,
                text: element.text,
                source: SourceRef {
                    part: diagram.data_part.clone(),
                    mapping: MappingQuality::Approximate,
                    locator: SourceLocator::PptxShape {
                        shape_id: shape_id
                            .unwrap_or(numeric_id)
                            .saturating_add(u32::try_from(index).unwrap_or(u32::MAX)),
                        row: Some(u32::try_from(index).unwrap_or(u32::MAX)),
                        column: None,
                        text_range: Some((0, text_length)),
                        metadata: metadata.clone(),
                    },
                },
                visual: element.visual,
            });
        }
        return Ok(());
    }
    let gap = 12.0_f32.min(bounds.height / 8.0);
    let inset = 16.0_f32.min(bounds.width / 8.0);
    let horizontal_process = diagram
        .layout_type
        .as_deref()
        .is_some_and(|value| value.contains("/hProcess7"));
    let picture_list = diagram
        .layout_type
        .as_deref()
        .is_some_and(|value| value.contains("/pList1"));
    let cycle = diagram
        .layout_type
        .as_deref()
        .is_some_and(|value| value.contains("/cycle8"));
    let process_visual = |visual| {
        if !horizontal_process {
            return visual;
        }
        let center_x = bounds.x + bounds.width / 2.0;
        let center_y = bounds.y + bounds.height / 2.0;
        let latitude = 26.0_f32.to_radians();
        let longitude = 18.0_f32.to_radians();
        let a = longitude.cos();
        let b = -latitude.sin() * longitude.sin();
        let d = latitude.cos();
        Visual::Layer {
            transform: AffineTransform {
                a,
                b,
                c: 0.0,
                d,
                e: center_x - a * center_x,
                f: center_y - b * center_x - d * center_y,
            },
            opacity: 1.0,
            blend_mode: crate::model::BlendMode::Normal,
            visual: Box::new(visual),
        }
    };
    let visible_ids = diagram
        .nodes
        .iter()
        .map(|node| node.model_id.clone())
        .collect::<HashSet<_>>();
    let root_ids = diagram
        .nodes
        .iter()
        .filter(|node| {
            node.parent_id
                .as_deref()
                .is_none_or(|parent_id| !visible_ids.contains(parent_id))
        })
        .map(|node| node.model_id.clone())
        .collect::<Vec<_>>();
    let available_height = (bounds.height - gap * (diagram.nodes.len() + 1) as f32).max(1.0);
    let node_height = (available_height / diagram.nodes.len() as f32).max(1.0);
    let root_width = (bounds.width - inset * 2.0).max(1.0);
    let explicit_fills = diagram
        .nodes
        .iter()
        .filter_map(|node| {
            node.fill
                .as_ref()
                .map(|fill| (node.model_id.clone(), fill.clone()))
        })
        .collect::<HashMap<_, _>>();
    let no_fill_ids = diagram
        .nodes
        .iter()
        .filter(|node| node.no_fill)
        .map(|node| node.model_id.clone())
        .collect::<HashSet<_>>();
    let fallback_shape_id =
        shape_id.unwrap_or_else(|| u32::try_from(state.objects.len()).unwrap_or(u32::MAX));
    let diagram_node_count = diagram.nodes.len();
    let layout = diagram
        .nodes
        .into_iter()
        .enumerate()
        .map(|(index, node)| {
            let visible_parent = node
                .parent_id
                .as_ref()
                .filter(|parent_id| visible_ids.contains(parent_id.as_str()));
            let nested = visible_parent.is_some();
            let node_bounds = if cycle {
                let size = bounds.width.min(bounds.height) * 0.76;
                let half = size / 2.0;
                let center_x = bounds.x + bounds.width / 2.0;
                let center_y = bounds.y + bounds.height / 2.0;
                Rect {
                    x: center_x - if index >= 2 { half } else { 0.0 },
                    y: center_y - if index == 0 || index == 3 { half } else { 0.0 },
                    width: half,
                    height: half,
                }
            } else if picture_list {
                let columns = 2_usize.min(diagram_node_count).max(1);
                let rows = diagram_node_count.div_ceil(columns).max(1);
                let cell_width = (bounds.width * 0.313_043_48).max(1.0);
                let cell_height = cell_width * 1.06;
                let column_gap = cell_width / 9.0;
                let row_gap = cell_width * 0.115;
                let grid_width =
                    cell_width * columns as f32 + column_gap * columns.saturating_sub(1) as f32;
                let grid_height =
                    cell_height * rows as f32 + row_gap * rows.saturating_sub(1) as f32;
                let column = index % columns;
                let row = index / columns;
                let cell_x = bounds.x
                    + (bounds.width - grid_width).max(0.0) / 2.0
                    + column as f32 * (cell_width + column_gap);
                let cell_y = bounds.y
                    + (bounds.height - grid_height).max(0.0) / 2.0
                    + row as f32 * (cell_height + row_gap);
                let picture_height = cell_height * 0.65;
                Rect {
                    x: cell_x,
                    y: cell_y
                        + if node.custom_geometry {
                            picture_height
                        } else {
                            0.0
                        },
                    width: cell_width,
                    height: cell_height * if node.custom_geometry { 0.35 } else { 0.65 },
                }
            } else if horizontal_process {
                let root_id = visible_parent.unwrap_or(&node.model_id);
                let column = root_ids
                    .iter()
                    .position(|candidate| candidate == root_id)
                    .unwrap_or(index.min(root_ids.len().saturating_sub(1)));
                let column_count = root_ids.len().max(1);
                let segment_width = bounds.width / column_count as f32;
                let column_width = segment_width * if column == 0 { 0.8 } else { 0.95 };
                let column_x = if column == 0 {
                    bounds.x + inset
                } else {
                    bounds.x + column as f32 * segment_width - inset * 0.625
                };
                Rect {
                    x: column_x + if nested { column_width * 0.20 } else { 0.0 },
                    y: bounds.y + bounds.height * if column == 0 { 0.28 } else { 0.30 },
                    width: column_width * if nested { 0.745 } else { 1.0 },
                    height: bounds.height * 0.37,
                }
            } else {
                let left_inset = if nested { inset * 2.0 } else { inset };
                Rect {
                    x: bounds.x + left_inset,
                    y: bounds.y + gap + index as f32 * (node_height + gap),
                    width: (root_width - if nested { inset } else { 0.0 }).max(1.0),
                    height: node_height,
                }
            };
            (index, node, node_bounds)
        })
        .collect::<Vec<_>>();
    let bounds_by_id = layout
        .iter()
        .map(|(_, node, node_bounds)| (node.model_id.clone(), *node_bounds))
        .collect::<HashMap<_, _>>();
    for (index, node, node_bounds) in &layout {
        let Some(parent_bounds) = node
            .parent_id
            .as_ref()
            .and_then(|parent_id| bounds_by_id.get(parent_id))
        else {
            continue;
        };
        if cycle || horizontal_process {
            continue;
        }
        let (start_x, start_y, end_x, end_y) = if horizontal_process {
            (
                parent_bounds.x + parent_bounds.width,
                parent_bounds.y + parent_bounds.height / 2.0,
                node_bounds.x,
                node_bounds.y + node_bounds.height / 2.0,
            )
        } else {
            (
                parent_bounds.x + parent_bounds.width / 2.0,
                parent_bounds.y + parent_bounds.height,
                node_bounds.x + node_bounds.width / 2.0,
                node_bounds.y,
            )
        };
        let connector_bounds = Rect {
            x: start_x.min(end_x),
            y: start_y.min(end_y),
            width: (end_x - start_x).abs().max(0.01),
            height: (end_y - start_y).abs().max(0.01),
        };
        let numeric_id = u32::try_from(state.objects.len())
            .map_err(|_| format_error(&diagram.data_part, "diagram object count exceeds range"))?;
        let z = state.take_z();
        state.objects.push(Object {
            numeric_id,
            parent_numeric_id,
            stable_id: format!("object:{numeric_id}"),
            parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
            kind: ObjectKind::Shape,
            unit_index,
            bounds: connector_bounds,
            z,
            text: None,
            source: SourceRef {
                part: diagram.data_part.clone(),
                mapping: MappingQuality::Derived,
                locator: SourceLocator::PptxShape {
                    shape_id: fallback_shape_id.saturating_add(*index as u32),
                    row: Some(u32::try_from(*index).unwrap_or(u32::MAX)),
                    column: None,
                    text_range: None,
                    metadata: metadata.clone(),
                },
            },
            visual: Visual::PaintedShape {
                geometry: Geometry::Path {
                    fill_rule: FillRule::NonZero,
                    commands: vec![
                        PathCommand::MoveTo {
                            x: start_x - connector_bounds.x,
                            y: start_y - connector_bounds.y,
                        },
                        PathCommand::LineTo {
                            x: end_x - connector_bounds.x,
                            y: end_y - connector_bounds.y,
                        },
                    ],
                },
                fill: Paint::None,
                stroke: Paint::Solid(0x5b9b_d5ff),
                stroke_width: 1.5,
            },
        });
    }
    if cycle {
        for (index, _, node_bounds) in &layout {
            let numeric_id = u32::try_from(state.objects.len()).map_err(|_| {
                format_error(&diagram.data_part, "diagram object count exceeds range")
            })?;
            let mut fill = theme
                .color(["accent2", "accent3", "accent4", "accent5"][index % 4])
                .unwrap_or(0x5b9b_d5ff);
            apply_drawingml_color_transform(&mut fill, "shade", 0.9);
            let visual = Visual::PaintedShape {
                geometry: Geometry::Path {
                    fill_rule: FillRule::NonZero,
                    commands: cycle_arrow_commands(*index, node_bounds.width, node_bounds.height),
                },
                fill: Paint::Solid(fill),
                stroke: Paint::None,
                stroke_width: 0.0,
            };
            let z = state.take_z();
            state.objects.push(Object {
                numeric_id,
                parent_numeric_id,
                stable_id: format!("object:{numeric_id}"),
                parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
                kind: ObjectKind::Shape,
                unit_index,
                bounds: *node_bounds,
                z,
                text: None,
                source: SourceRef {
                    part: diagram.data_part.clone(),
                    mapping: MappingQuality::Approximate,
                    locator: SourceLocator::PptxShape {
                        shape_id: fallback_shape_id.saturating_add(*index as u32),
                        row: Some(u32::try_from(*index).unwrap_or(u32::MAX)),
                        column: None,
                        text_range: None,
                        metadata: metadata.clone(),
                    },
                },
                visual: Visual::Effect {
                    shadow: Some(Shadow {
                        color: 0x0000_0059,
                        blur: 5.0,
                        offset_x: 2.0,
                        offset_y: 3.0,
                    }),
                    clip: None,
                    visual: Box::new(visual),
                },
            });
        }
    }
    for (index, node, node_bounds) in layout {
        let nested = node
            .parent_id
            .as_deref()
            .is_some_and(|parent_id| visible_ids.contains(parent_id));
        let process_label = (horizontal_process && !nested).then(|| node.text.clone());
        let custom_geometry = node.custom_geometry;
        if picture_list && custom_geometry {
            let picture_bounds = Rect {
                x: node_bounds.x,
                y: node_bounds.y - node_bounds.height * 0.65 / 0.35,
                width: node_bounds.width,
                height: node_bounds.height * 0.65 / 0.35,
            };
            let numeric_id = u32::try_from(state.objects.len()).map_err(|_| {
                format_error(&diagram.data_part, "diagram object count exceeds range")
            })?;
            let z = state.take_z();
            state.objects.push(Object {
                numeric_id,
                parent_numeric_id,
                stable_id: format!("object:{numeric_id}"),
                parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
                kind: ObjectKind::Shape,
                unit_index,
                bounds: picture_bounds,
                z,
                text: None,
                source: SourceRef {
                    part: diagram.data_part.clone(),
                    mapping: MappingQuality::Derived,
                    locator: SourceLocator::PptxShape {
                        shape_id: fallback_shape_id.saturating_add(index as u32),
                        row: Some(u32::try_from(index).unwrap_or(u32::MAX)),
                        column: None,
                        text_range: None,
                        metadata: metadata.clone(),
                    },
                },
                visual: Visual::PaintedShape {
                    geometry: Geometry::RoundedRectangle {
                        radius_x: picture_bounds.width.min(picture_bounds.height) / 6.0,
                        radius_y: picture_bounds.width.min(picture_bounds.height) / 6.0,
                    },
                    fill: Paint::Solid(theme.color("accent1").unwrap_or(0x4f81_bdff)),
                    stroke: Paint::Solid(0xffff_ffff),
                    stroke_width: 3.0,
                },
            });
        }
        let geometry = if cycle {
            let w = node_bounds.width;
            let h = node_bounds.height;
            let k = 0.552_284_8;
            let commands = match index % 4 {
                0 => vec![
                    PathCommand::MoveTo { x: 0.0, y: h },
                    PathCommand::LineTo { x: 0.0, y: 0.0 },
                    PathCommand::BezierCurveTo {
                        cp1x: w * k,
                        cp1y: 0.0,
                        cp2x: w,
                        cp2y: h * (1.0 - k),
                        x: w,
                        y: h,
                    },
                    PathCommand::ClosePath,
                ],
                1 => vec![
                    PathCommand::MoveTo { x: 0.0, y: 0.0 },
                    PathCommand::LineTo { x: w, y: 0.0 },
                    PathCommand::BezierCurveTo {
                        cp1x: w,
                        cp1y: h * k,
                        cp2x: w * k,
                        cp2y: h,
                        x: 0.0,
                        y: h,
                    },
                    PathCommand::ClosePath,
                ],
                2 => vec![
                    PathCommand::MoveTo { x: w, y: 0.0 },
                    PathCommand::LineTo { x: w, y: h },
                    PathCommand::BezierCurveTo {
                        cp1x: w * (1.0 - k),
                        cp1y: h,
                        cp2x: 0.0,
                        cp2y: h * k,
                        x: 0.0,
                        y: 0.0,
                    },
                    PathCommand::ClosePath,
                ],
                _ => vec![
                    PathCommand::MoveTo { x: w, y: h },
                    PathCommand::LineTo { x: 0.0, y: h },
                    PathCommand::BezierCurveTo {
                        cp1x: 0.0,
                        cp1y: h * (1.0 - k),
                        cp2x: w * (1.0 - k),
                        cp2y: 0.0,
                        x: w,
                        y: 0.0,
                    },
                    PathCommand::ClosePath,
                ],
            };
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands,
            }
        } else if custom_geometry {
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo {
                        x: 0.0,
                        y: node_bounds.height,
                    },
                    PathCommand::LineTo {
                        x: node_bounds.width,
                        y: node_bounds.height,
                    },
                    PathCommand::LineTo {
                        x: node_bounds.width * 0.463,
                        y: 0.0,
                    },
                    PathCommand::ClosePath,
                ],
            }
        } else if picture_list {
            node.preset_geometry
                .as_deref()
                .and_then(|preset| {
                    resolve_preset_geometry(preset, node_bounds, &HashMap::new(), 0.0, None, None)
                        .map(|(geometry, _, _)| geometry)
                })
                .unwrap_or(Geometry::RoundedRectangle {
                    radius_x: node_bounds.width.min(node_bounds.height) / 6.0,
                    radius_y: node_bounds.width.min(node_bounds.height) / 6.0,
                })
        } else {
            Geometry::RoundedRectangle {
                radius_x: 6.0,
                radius_y: 6.0,
            }
        };
        let inherited_fill = node
            .parent_id
            .as_deref()
            .and_then(|parent_id| explicit_fills.get(parent_id));
        let no_fill = node.no_fill
            || (horizontal_process
                && nested
                && node
                    .parent_id
                    .as_deref()
                    .is_some_and(|parent_id| no_fill_ids.contains(parent_id)));
        let fill = if cycle {
            theme
                .color(["accent2", "accent3", "accent4", "accent5"][index % 4])
                .unwrap_or(0x5b9b_d5ff)
        } else {
            let mut fill = node
                .fill_scheme
                .as_deref()
                .and_then(|value| theme.color(value))
                .unwrap_or(if picture_list {
                    theme.color("accent1").unwrap_or(0x4f81_bdff)
                } else if horizontal_process {
                    theme.color("accent1").unwrap_or(0x4f81_bdff)
                } else if nested {
                    0xeaf2_f8ff
                } else {
                    0xd9ea_f7ff
                });
            for (kind, ratio) in &node.fill_transforms {
                apply_drawingml_color_transform(&mut fill, kind, *ratio);
            }
            fill
        };
        let numeric_id = u32::try_from(state.objects.len())
            .map_err(|_| format_error(&diagram.data_part, "diagram object count exceeds range"))?;
        let text_length = u32::try_from(node.text.chars().count()).unwrap_or(u32::MAX);
        let z = state.take_z();
        let shape_fill = if no_fill || (horizontal_process && nested) {
            Paint::None
        } else if let Some(fill) = node.fill.as_ref().or(inherited_fill) {
            fill.paint(node_bounds)
        } else {
            Paint::Solid(fill)
        };
        let shape_fill = if diagram.scene_three_d {
            match shape_fill {
                Paint::Solid(color) => {
                    let channel = |shift: u32, light: f32| {
                        ((((color >> shift) & 0xff) as f32 * light).round() as u32).min(255)
                    };
                    Paint::Solid(
                        (channel(24, 1.076) << 24)
                            | (channel(16, 1.062) << 16)
                            | (channel(8, 1.048) << 8)
                            | (color & 0xff),
                    )
                }
                paint => paint,
            }
        } else {
            shape_fill
        };
        let visual = Visual::TextLayout {
            layout: TextLayout {
                vertical_align: if picture_list {
                    TextVerticalAlign::Top
                } else {
                    TextVerticalAlign::Center
                },
                ..TextLayout::default()
            },
            visual: Box::new(Visual::RichText {
                geometry,
                fill: shape_fill,
                stroke: if horizontal_process && nested {
                    Paint::None
                } else if cycle || (picture_list && !custom_geometry) {
                    Paint::Solid(0xffff_ffff)
                } else if picture_list {
                    Paint::None
                } else {
                    Paint::Solid(0x5b9b_d5ff)
                },
                stroke_width: if horizontal_process && nested {
                    0.0
                } else if cycle || (picture_list && !custom_geometry) {
                    3.0
                } else if picture_list {
                    0.0
                } else {
                    1.0
                },
                align: if picture_list || (horizontal_process && !nested) {
                    TextAlign::Start
                } else {
                    TextAlign::Center
                },
                line_height: 0.0,
                runs: vec![TextRun {
                    paint: None,
                    east_asian_line_breaks: true,
                    text: process_label
                        .as_ref()
                        .map_or_else(|| node.text.clone(), |_| String::new()),
                    font_family: "Arial".to_owned(),
                    font_size: if picture_list || (horizontal_process && !nested) {
                        5.0 * POINTS_TO_CSS_PIXELS
                    } else if horizontal_process {
                        36.0 * POINTS_TO_CSS_PIXELS
                    } else {
                        14.0 * POINTS_TO_CSS_PIXELS
                    },
                    color: if cycle || horizontal_process {
                        0xffff_ffff
                    } else {
                        0x1732_4dff
                    },
                    bold: !picture_list && !nested && !cycle && !horizontal_process,
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
        let shadow = node.shadow.or_else(|| {
            if cycle {
                Some(Shadow {
                    color: 0x0000_004d,
                    blur: 5.0,
                    offset_x: 2.0,
                    offset_y: 3.0,
                })
            } else {
                None
            }
        });
        let visual = if let Some(shadow) = shadow {
            Visual::Effect {
                shadow: Some(shadow),
                clip: None,
                visual: Box::new(visual),
            }
        } else {
            visual
        };
        let visual = if let Some(three_d) = node.three_d {
            Visual::AdvancedEffect {
                outer_shadow: None,
                inner_shadow: None,
                glow: None,
                reflection: None,
                soft_edge: None,
                three_d: Some(three_d),
                visual: Box::new(visual),
            }
        } else {
            visual
        };
        let visual = process_visual(visual);
        state.objects.push(Object {
            numeric_id,
            parent_numeric_id,
            stable_id: format!("object:{numeric_id}"),
            parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
            kind: if custom_geometry || process_label.is_some() {
                ObjectKind::Shape
            } else {
                ObjectKind::TextBox
            },
            unit_index,
            bounds: node_bounds,
            z,
            text: process_label.is_none().then(|| node.text.clone()),
            source: SourceRef {
                part: diagram.data_part.clone(),
                mapping: MappingQuality::Approximate,
                locator: SourceLocator::PptxShape {
                    shape_id: fallback_shape_id.saturating_add(index as u32),
                    row: Some(u32::try_from(index).unwrap_or(u32::MAX)),
                    column: None,
                    text_range: Some((0, text_length)),
                    metadata: metadata.clone(),
                },
            },
            visual,
        });
        if let Some(label) = process_label {
            let label_bounds = Rect {
                x: node_bounds.x,
                y: node_bounds.y,
                width: node_bounds.width * 0.2,
                height: node_bounds.height * 0.82,
            };
            let numeric_id = u32::try_from(state.objects.len()).map_err(|_| {
                format_error(&diagram.data_part, "diagram object count exceeds range")
            })?;
            let z = state.take_z();
            state.objects.push(Object {
                numeric_id,
                parent_numeric_id,
                stable_id: format!("object:{numeric_id}"),
                parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
                kind: ObjectKind::TextBox,
                unit_index,
                bounds: label_bounds,
                z,
                text: Some(label.clone()),
                source: SourceRef {
                    part: diagram.data_part.clone(),
                    mapping: MappingQuality::Approximate,
                    locator: SourceLocator::PptxShape {
                        shape_id: fallback_shape_id.saturating_add(index as u32),
                        row: Some(u32::try_from(index).unwrap_or(u32::MAX)),
                        column: None,
                        text_range: Some((0, text_length)),
                        metadata: metadata.clone(),
                    },
                },
                visual: process_visual(Visual::TextLayout {
                    layout: TextLayout {
                        orientation: TextOrientation::Rotated270,
                        vertical_align: TextVerticalAlign::Top,
                        wrap: false,
                        ..TextLayout::default()
                    },
                    visual: Box::new(Visual::RichText {
                        geometry: Geometry::Rectangle,
                        fill: Paint::None,
                        stroke: Paint::None,
                        stroke_width: 0.0,
                        align: TextAlign::End,
                        line_height: 0.0,
                        runs: vec![TextRun {
                            paint: None,
                            east_asian_line_breaks: true,
                            text: label,
                            font_family: "Arial".to_owned(),
                            font_size: 6.5 * POINTS_TO_CSS_PIXELS,
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
                }),
            });
        }
    }
    if horizontal_process {
        for (index, pair) in root_ids.windows(2).enumerate() {
            let Some(left) = bounds_by_id.get(&pair[0]) else {
                continue;
            };
            let size = (left.width * 0.15).max(1.0);
            let connector_bounds = Rect {
                x: left.x + left.width - size / 2.0,
                y: (left.y + left.width * 0.8).min(left.y + left.height - size),
                width: size,
                height: size,
            };
            let numeric_id = u32::try_from(state.objects.len()).map_err(|_| {
                format_error(&diagram.data_part, "diagram object count exceeds range")
            })?;
            let z = state.take_z();
            state.objects.push(Object {
                numeric_id,
                parent_numeric_id,
                stable_id: format!("object:{numeric_id}"),
                parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
                kind: ObjectKind::Shape,
                unit_index,
                bounds: connector_bounds,
                z,
                text: None,
                source: SourceRef {
                    part: diagram.data_part.clone(),
                    mapping: MappingQuality::Derived,
                    locator: SourceLocator::PptxShape {
                        shape_id: fallback_shape_id
                            .saturating_add(diagram_node_count as u32)
                            .saturating_add(index as u32),
                        row: Some(u32::try_from(index).unwrap_or(u32::MAX)),
                        column: None,
                        text_range: None,
                        metadata: metadata.clone(),
                    },
                },
                visual: process_visual(Visual::PaintedShape {
                    geometry: Geometry::Path {
                        fill_rule: FillRule::NonZero,
                        commands: vec![
                            PathCommand::MoveTo { x: 0.0, y: 0.0 },
                            PathCommand::LineTo {
                                x: connector_bounds.width,
                                y: connector_bounds.height / 2.0,
                            },
                            PathCommand::LineTo {
                                x: 0.0,
                                y: connector_bounds.height,
                            },
                            PathCommand::ClosePath,
                        ],
                    },
                    fill: Paint::Solid(theme.color("lt1").unwrap_or(0xffff_ffff)),
                    stroke: Paint::Solid(accent),
                    stroke_width: 1.5,
                }),
            });
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn try_push_pptx_diagram_drawing(
    package: &Package<'_>,
    diagram: &PptxDiagram,
    bounds: Rect,
    unit_index: u32,
    state: &mut PptxParseState,
    content_types: &ContentTypes,
    font_metrics: &FontMetricTable,
    theme: &PptxTheme,
    metadata: &PptxObjectMetadata,
) -> Result<bool, Diagnostic> {
    for (drawing_part, drawing_is_frame_relative) in &diagram.drawing_parts {
        if !package.has_part(drawing_part) {
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Parse,
                    Fidelity::Approximate,
                    "PPTX diagram drawing fallback is missing; the data-model fallback was used",
                )
                .in_part(drawing_part),
            );
            continue;
        }
        let object_count = state.objects.len();
        parse_presentation_part(
            package,
            drawing_part,
            unit_index,
            bounds,
            PartContent::SLIDE,
            state,
            content_types,
            font_metrics,
            theme,
            None,
            Some(&diagram.drawing_text_colors),
        )?;
        if state.objects.len() == object_count {
            continue;
        }
        for object in &mut state.objects[object_count..] {
            if *drawing_is_frame_relative {
                if let Visual::Group { children } = &mut object.visual {
                    for child in children {
                        child.bounds.x += bounds.x;
                        child.bounds.y += bounds.y;
                        if let Visual::Layer { transform, .. } = &mut child.visual {
                            transform.e +=
                                bounds.x - transform.a * bounds.x - transform.c * bounds.y;
                            transform.f +=
                                bounds.y - transform.b * bounds.x - transform.d * bounds.y;
                        }
                    }
                } else if let Visual::Layer { transform, .. } = &mut object.visual {
                    transform.e += bounds.x - transform.a * bounds.x - transform.c * bounds.y;
                    transform.f += bounds.y - transform.b * bounds.x - transform.d * bounds.y;
                }
                object.bounds.x += bounds.x;
                object.bounds.y += bounds.y;
            }
            if let SourceLocator::PptxShape {
                metadata: object_metadata,
                ..
            } = &mut object.source.locator
            {
                *object_metadata = metadata.clone();
            }
        }
        return Ok(true);
    }
    Ok(false)
}

#[derive(Clone, Debug)]
struct ResolvedTableBorder {
    border: TableCellBorder,
    source_row: usize,
    source_column: usize,
    order: usize,
}

fn table_border_is_visible(border: &TableCellBorder) -> bool {
    border.width > 0.0 && !matches!(border.paint, Paint::None)
}

fn table_border_replaces(candidate: &ResolvedTableBorder, existing: &ResolvedTableBorder) -> bool {
    if candidate.border.explicit != existing.border.explicit {
        return candidate.border.explicit;
    }
    let candidate_visible = table_border_is_visible(&candidate.border);
    let existing_visible = table_border_is_visible(&existing.border);
    if candidate_visible != existing_visible {
        return candidate_visible;
    }
    if (candidate.border.width - existing.border.width).abs() > f32::EPSILON {
        return candidate.border.width > existing.border.width;
    }
    candidate.order > existing.order
}

fn insert_table_border(
    borders: &mut BTreeMap<(u8, usize, usize), ResolvedTableBorder>,
    key: (u8, usize, usize),
    candidate: ResolvedTableBorder,
) {
    if borders
        .get(&key)
        .is_none_or(|existing| table_border_replaces(&candidate, existing))
    {
        borders.insert(key, candidate);
    }
}

#[allow(clippy::too_many_arguments)]
fn collect_table_cell_borders(
    borders: &mut BTreeMap<(u8, usize, usize), ResolvedTableBorder>,
    cell: &TableCellState,
    row: usize,
    column: usize,
    row_count: usize,
    column_count: usize,
    order: usize,
) {
    let row_end = row.saturating_add(cell.row_span).min(row_count);
    let column_end = column.saturating_add(cell.grid_span).min(column_count);
    if row >= row_end || column >= column_end {
        return;
    }
    let candidate = |side| ResolvedTableBorder {
        border: cell.border(side).clone(),
        source_row: row,
        source_column: column,
        order: order.saturating_mul(4).saturating_add(side.index()),
    };
    for segment in column..column_end {
        insert_table_border(borders, (0, row, segment), candidate(TableBorderSide::Top));
        insert_table_border(
            borders,
            (0, row_end, segment),
            candidate(TableBorderSide::Bottom),
        );
    }
    for segment in row..row_end {
        insert_table_border(
            borders,
            (1, column, segment),
            candidate(TableBorderSide::Left),
        );
        insert_table_border(
            borders,
            (1, column_end, segment),
            candidate(TableBorderSide::Right),
        );
    }
}

#[allow(clippy::too_many_arguments)]
fn push_table_border(
    key: (u8, usize, usize),
    resolved: ResolvedTableBorder,
    bounds: Rect,
    column_offsets: &[f32],
    row_offsets: &[f32],
    table_numeric_id: u32,
    part: &str,
    unit_index: u32,
    shape_id: u32,
    mapping: MappingQuality,
    state: &mut PptxParseState,
) -> Result<(), Diagnostic> {
    if !table_border_is_visible(&resolved.border) {
        return Ok(());
    }
    let (axis, boundary, segment) = key;
    let stroke_width = resolved.border.width;
    let (line_bounds, geometry) = if axis == 0 {
        let length = column_offsets[segment + 1] - column_offsets[segment];
        (
            Rect {
                x: bounds.x + column_offsets[segment],
                y: bounds.y + row_offsets[boundary] - stroke_width / 2.0,
                width: length,
                height: stroke_width,
            },
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo {
                        x: 0.0,
                        y: stroke_width / 2.0,
                    },
                    PathCommand::LineTo {
                        x: length,
                        y: stroke_width / 2.0,
                    },
                ],
            },
        )
    } else {
        let length = row_offsets[segment + 1] - row_offsets[segment];
        (
            Rect {
                x: bounds.x + column_offsets[boundary] - stroke_width / 2.0,
                y: bounds.y + row_offsets[segment],
                width: stroke_width,
                height: length,
            },
            Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo {
                        x: stroke_width / 2.0,
                        y: 0.0,
                    },
                    PathCommand::LineTo {
                        x: stroke_width / 2.0,
                        y: length,
                    },
                ],
            },
        )
    };
    let base = Visual::PaintedShape {
        geometry,
        fill: Paint::None,
        stroke: resolved.border.paint,
        stroke_width,
    };
    let dash = resolved
        .border
        .dash_pattern
        .lengths()
        .iter()
        .map(|value| value * stroke_width.max(1.0))
        .collect::<Vec<_>>();
    let visual = if dash.is_empty() && resolved.border.stroke_style == StrokeStyle::default() {
        base
    } else {
        Visual::StrokeStyle {
            style: StrokeStyle {
                dash,
                ..resolved.border.stroke_style
            },
            visual: Box::new(base),
        }
    };
    let numeric_id = u32::try_from(state.objects.len())
        .map_err(|_| format_error(part, "object count exceeds supported range"))?;
    let z = state.take_z();
    state.objects.push(Object {
        numeric_id,
        parent_numeric_id: Some(table_numeric_id),
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: Some(format!("object:{table_numeric_id}")),
        kind: ObjectKind::Shape,
        unit_index,
        bounds: line_bounds,
        z,
        text: None,
        source: SourceRef {
            part: part.to_owned(),
            mapping,
            locator: SourceLocator::PptxShape {
                shape_id,
                row: Some(u32::try_from(resolved.source_row).unwrap_or(u32::MAX)),
                column: Some(u32::try_from(resolved.source_column).unwrap_or(u32::MAX)),
                text_range: None,
                metadata: PptxObjectMetadata::default(),
            },
        },
        visual,
    });
    Ok(())
}

fn push_graphic_frame(
    mut frame: GraphicFrameState,
    part: &str,
    unit_index: u32,
    state: &mut PptxParseState,
    object_limit: usize,
    font_metrics: &FontMetricTable,
) -> Result<(), Diagnostic> {
    let metadata = frame.metadata.clone();
    let mut bounds = Rect {
        x: frame.x.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        y: frame.y.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        width: frame.width.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        height: frame.height.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
    };
    if !bounds.is_valid() {
        return Err(format_error(part, "graphic frame has invalid bounds"));
    }
    let mapping = if frame.shape_id.is_some() {
        MappingQuality::Exact
    } else {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Approximate,
                "graphic frame has no source ID; a session-local mapping was derived",
            )
            .in_part(part),
        );
        MappingQuality::Derived
    };
    let shape_id = frame
        .shape_id
        .unwrap_or_else(|| u32::try_from(state.objects.len()).unwrap_or(u32::MAX));
    if !frame.rows.is_empty() || !frame.columns.is_empty() {
        for row in &mut frame.rows {
            for cell in &mut row.cells {
                resolve_table_cell_paragraph_layouts(cell);
            }
        }
        let cell_count = frame.rows.iter().try_fold(0_usize, |count, row| {
            count
                .checked_add(row.cells.len())
                .ok_or_else(|| format_error(part, "table cell count exceeds supported range"))
        })?;
        let required = 1_usize
            .checked_add(cell_count)
            .and_then(|count| state.objects.len().checked_add(count))
            .ok_or_else(|| format_error(part, "table object count exceeds supported range"))?;
        if required > object_limit {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ObjectLimit,
                Phase::Parse,
                None,
                "document exceeds the configured object limit",
            )
            .in_part(part));
        }
        let column_count = frame.columns.len().max(
            frame
                .rows
                .iter()
                .map(|row| {
                    row.cells
                        .iter()
                        .filter(|cell| !cell.horizontal_merge)
                        .map(|cell| cell.grid_span.max(1))
                        .sum()
                })
                .max()
                .unwrap_or(0),
        );
        let column_sizes = normalized_table_sizes(&frame.columns, column_count, bounds.width);
        let row_sizes = table_row_sizes(&frame.rows, &column_sizes, bounds.height, font_metrics);
        bounds.height = bounds.height.max(row_sizes.iter().sum());
        let mut column_offsets = Vec::with_capacity(column_sizes.len() + 1);
        column_offsets.push(0.0);
        for size in &column_sizes {
            column_offsets.push(column_offsets.last().copied().unwrap_or(0.0) + size);
        }
        let mut row_offsets = Vec::with_capacity(row_sizes.len() + 1);
        row_offsets.push(0.0);
        for size in &row_sizes {
            row_offsets.push(row_offsets.last().copied().unwrap_or(0.0) + size);
        }
        let table_numeric_id = u32::try_from(state.objects.len())
            .map_err(|_| format_error(part, "object count exceeds supported range"))?;
        let z = state.take_z();
        let table_visual =
            frame
                .table_fill
                .take()
                .map_or(Visual::None, |fill| Visual::PaintedShape {
                    geometry: Geometry::Rectangle,
                    fill: fill.paint(bounds),
                    stroke: Paint::None,
                    stroke_width: 0.0,
                });
        state.objects.push(Object {
            numeric_id: table_numeric_id,
            parent_numeric_id: frame.parent_numeric_id,
            stable_id: format!("object:{table_numeric_id}"),
            parent_stable_id: frame
                .parent_numeric_id
                .map(|parent| format!("object:{parent}")),
            kind: ObjectKind::Table,
            unit_index,
            bounds,
            z,
            text: None,
            source: SourceRef {
                part: part.to_owned(),
                mapping,
                locator: SourceLocator::PptxShape {
                    shape_id,
                    row: None,
                    column: None,
                    text_range: None,
                    metadata: metadata.clone(),
                },
            },
            visual: frame
                .table_effects
                .take()
                .map_or(table_visual.clone(), |effects| {
                    effects.wrap(table_visual, None)
                }),
        });

        let mut table_borders = BTreeMap::new();
        let row_count = row_sizes.len();
        let mut cell_order = 0_usize;
        for (row_index, row) in frame.rows.into_iter().enumerate() {
            let mut logical_column = 0_usize;
            for cell in row.cells {
                // PowerPoint emits an hMerge continuation cell in addition to
                // the merge origin's gridSpan. The origin already consumed
                // those grid columns, so counting the continuation again
                // shifts every following cell one column to the right.
                if cell.horizontal_merge {
                    cell_order = cell_order.saturating_add(1);
                    continue;
                }
                let column_index = logical_column;
                logical_column = logical_column.saturating_add(cell.grid_span.max(1));
                cell_order = cell_order.saturating_add(1);
                if cell.vertical_merge {
                    continue;
                }
                if row_index >= row_count || column_index >= column_sizes.len() {
                    continue;
                }
                collect_table_cell_borders(
                    &mut table_borders,
                    &cell,
                    row_index,
                    column_index,
                    row_count,
                    column_sizes.len(),
                    cell_order,
                );
                let row_end = row_index.saturating_add(cell.row_span).min(row_count);
                let column_end = column_index
                    .saturating_add(cell.grid_span)
                    .min(column_sizes.len());
                let x = bounds.x + column_offsets[column_index];
                let y = bounds.y + row_offsets[row_index];
                let width = column_offsets[column_end] - column_offsets[column_index];
                let height = row_offsets[row_end] - row_offsets[row_index];
                let cell_bounds = Rect {
                    x,
                    y,
                    width,
                    height,
                };
                let numeric_id = u32::try_from(state.objects.len())
                    .map_err(|_| format_error(part, "object count exceeds supported range"))?;
                let has_text = !cell.text.is_empty();
                let text_length = u32::try_from(cell.text.chars().count()).unwrap_or(u32::MAX);
                let text_layout = TextLayout {
                    direction: cell.text_direction,
                    vertical_align: cell.vertical_align,
                    orientation: cell.text_orientation,
                    inset_left: cell.text_inset_left,
                    inset_right: cell.text_inset_right,
                    inset_top: cell.text_inset_top,
                    inset_bottom: cell.text_inset_bottom,
                    paragraphs: cell.paragraph_layouts,
                    ..TextLayout::default()
                };
                let mut run_effects = cell.run_effects;
                let runs = if cell.runs.is_empty() && has_text {
                    vec![TextRun {
                        paint: None,
                        east_asian_line_breaks: cell.east_asian_line_breaks,
                        text: cell.text.clone(),
                        font_family: cell.font_family,
                        font_size: cell.font_size,
                        color: cell.font_color,
                        bold: cell.bold,
                        italic: cell.italic,
                        underline: cell.underline,
                        strikethrough: cell.strikethrough,
                        highlight: 0,
                        baseline_shift: cell.baseline_shift,
                        letter_spacing: cell.letter_spacing,
                        horizontal_scale: 1.0,
                    }]
                } else {
                    cell.runs
                };
                run_effects.resize(runs.len(), TextEffect::default());
                let mut visual = if has_text {
                    Visual::RichText {
                        geometry: Geometry::Rectangle,
                        fill: cell
                            .style_fill
                            .as_ref()
                            .map_or_else(|| cell.fill.clone(), |fill| fill.paint(cell_bounds)),
                        stroke: Paint::None,
                        stroke_width: 0.0,
                        align: cell.align,
                        line_height: 0.0,
                        runs,
                    }
                } else {
                    Visual::PaintedShape {
                        geometry: Geometry::Rectangle,
                        fill: cell
                            .style_fill
                            .as_ref()
                            .map_or_else(|| cell.fill.clone(), |fill| fill.paint(cell_bounds)),
                        stroke: Paint::None,
                        stroke_width: 0.0,
                    }
                };
                if let Some(effects) = cell.effects {
                    visual = effects.wrap(visual, None);
                }
                if run_effects
                    .iter()
                    .any(|effect| *effect != TextEffect::default())
                {
                    visual = Visual::TextEffects {
                        effects: run_effects,
                        visual: Box::new(visual),
                    };
                }
                let visual = if has_text {
                    Visual::TextLayout {
                        layout: text_layout,
                        visual: Box::new(visual),
                    }
                } else {
                    visual
                };
                let z = state.take_z();
                state.objects.push(Object {
                    numeric_id,
                    parent_numeric_id: Some(table_numeric_id),
                    stable_id: format!("object:{numeric_id}"),
                    parent_stable_id: Some(format!("object:{table_numeric_id}")),
                    kind: ObjectKind::Cell,
                    unit_index,
                    bounds: cell_bounds,
                    z,
                    text: has_text.then_some(cell.text),
                    source: SourceRef {
                        part: part.to_owned(),
                        mapping,
                        locator: SourceLocator::PptxShape {
                            shape_id,
                            row: Some(u32::try_from(row_index).unwrap_or(u32::MAX)),
                            column: Some(u32::try_from(column_index).unwrap_or(u32::MAX)),
                            text_range: has_text.then_some((0, text_length)),
                            metadata: PptxObjectMetadata::default(),
                        },
                    },
                    visual,
                });
            }
        }
        for (key, border) in table_borders {
            if table_border_is_visible(&border.border) && state.objects.len() >= object_limit {
                return Err(Diagnostic::fatal(
                    DiagnosticCode::ObjectLimit,
                    Phase::Parse,
                    None,
                    "document exceeds the configured object limit",
                )
                .in_part(part));
            }
            push_table_border(
                key,
                border,
                bounds,
                &column_offsets,
                &row_offsets,
                table_numeric_id,
                part,
                unit_index,
                shape_id,
                mapping,
                state,
            )?;
        }
        return Ok(());
    }

    if let Some(chart) = frame.chart {
        push_basic_pptx_chart(
            chart,
            bounds,
            part,
            unit_index,
            frame.parent_numeric_id,
            shape_id,
            mapping,
            metadata.clone(),
            state,
            object_limit,
        )?;
        return Ok(());
    }

    if let Some((media_type, bytes, source_part)) = frame.embedded_raster {
        if state.objects.len() >= object_limit {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ObjectLimit,
                Phase::Parse,
                None,
                "document exceeds the configured object limit",
            )
            .in_part(part));
        }
        let numeric_id = u32::try_from(state.objects.len())
            .map_err(|_| format_error(part, "object count exceeds supported range"))?;
        let z = state.take_z();
        state.objects.push(Object {
            numeric_id,
            parent_numeric_id: frame.parent_numeric_id,
            stable_id: format!("object:{numeric_id}"),
            parent_stable_id: frame
                .parent_numeric_id
                .map(|parent| format!("object:{parent}")),
            kind: ObjectKind::Image,
            unit_index,
            bounds,
            z,
            text: None,
            source: SourceRef {
                part: source_part,
                mapping,
                locator: SourceLocator::PptxShape {
                    shape_id,
                    row: None,
                    column: None,
                    text_range: None,
                    metadata,
                },
            },
            visual: Visual::Image {
                media_type,
                bytes,
                crop: ImageCrop::default(),
            },
        });
        return Ok(());
    }

    if let Some((document, native_width, source_part)) = frame.embedded_text
        && native_width > 0.0
        && let Some(page_width) = document.units.first().map(|unit| unit.width)
        && page_width > 0.0
        && let Some(source) = document
            .objects
            .into_iter()
            .find(|object| object.text.as_deref().is_some_and(|text| !text.is_empty()))
    {
        if state.objects.len() >= object_limit {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ObjectLimit,
                Phase::Parse,
                None,
                "document exceeds the configured object limit",
            )
            .in_part(part));
        }
        let scale = bounds.width / native_width;
        let text = source.text.unwrap_or_default();
        let text_length = u32::try_from(text.chars().count()).unwrap_or(u32::MAX);
        let numeric_id = u32::try_from(state.objects.len())
            .map_err(|_| format_error(part, "object count exceeds supported range"))?;
        let z = state.take_z();
        state.objects.push(Object {
            numeric_id,
            parent_numeric_id: frame.parent_numeric_id,
            stable_id: format!("object:{numeric_id}"),
            parent_stable_id: frame
                .parent_numeric_id
                .map(|parent| format!("object:{parent}")),
            kind: ObjectKind::TextBox,
            unit_index,
            bounds: Rect {
                width: (bounds.width * source.bounds.width / page_width).min(bounds.width),
                ..bounds
            },
            z,
            text: Some(text),
            source: SourceRef {
                part: source_part,
                mapping,
                locator: SourceLocator::PptxShape {
                    shape_id,
                    row: None,
                    column: None,
                    text_range: Some((0, text_length)),
                    metadata,
                },
            },
            visual: scale_embedded_text_visual(source.visual, scale),
        });
        return Ok(());
    }

    if let Some((text, font_scale, source_part)) = frame.embedded_plain_text {
        if state.objects.len() >= object_limit {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ObjectLimit,
                Phase::Parse,
                None,
                "document exceeds the configured object limit",
            )
            .in_part(part));
        }
        let text_length = u32::try_from(text.chars().count()).unwrap_or(u32::MAX);
        let numeric_id = u32::try_from(state.objects.len())
            .map_err(|_| format_error(part, "object count exceeds supported range"))?;
        let z = state.take_z();
        state.objects.push(Object {
            numeric_id,
            parent_numeric_id: frame.parent_numeric_id,
            stable_id: format!("object:{numeric_id}"),
            parent_stable_id: frame
                .parent_numeric_id
                .map(|parent| format!("object:{parent}")),
            kind: ObjectKind::TextBox,
            unit_index,
            bounds,
            z,
            text: Some(text.clone()),
            source: SourceRef {
                part: source_part,
                mapping,
                locator: SourceLocator::PptxShape {
                    shape_id,
                    row: None,
                    column: None,
                    text_range: Some((0, text_length)),
                    metadata,
                },
            },
            visual: Visual::TextLayout {
                layout: TextLayout {
                    vertical_align: TextVerticalAlign::Center,
                    ..TextLayout::default()
                },
                visual: Box::new(Visual::RichText {
                    geometry: Geometry::Rectangle,
                    fill: Paint::None,
                    stroke: Paint::None,
                    stroke_width: 0.0,
                    align: TextAlign::Center,
                    line_height: 0.0,
                    runs: vec![TextRun {
                        paint: None,
                        east_asian_line_breaks: true,
                        text,
                        font_family: "Times New Roman".to_owned(),
                        font_size: bounds.height * font_scale,
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
                }),
            },
        });
        return Ok(());
    }

    if let Some((document, source_part)) = frame.embedded_document
        && let Some(source_unit) = document.units.first()
        && source_unit.width > 0.0
        && source_unit.height > 0.0
    {
        let text_objects: Vec<_> = document
            .objects
            .into_iter()
            .filter(|object| object.text.as_deref().is_some_and(|text| !text.is_empty()))
            .collect();
        let rendered = !text_objects.is_empty();
        if state.objects.len().saturating_add(text_objects.len()) > object_limit {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ObjectLimit,
                Phase::Parse,
                None,
                "document exceeds the configured object limit",
            )
            .in_part(part));
        }
        let scale_x = bounds.width / source_unit.width;
        let scale_y = bounds.height / source_unit.height;
        let text_scale = scale_x.min(scale_y);
        for source in text_objects {
            let text = source.text.unwrap_or_default();
            let text_length = u32::try_from(text.chars().count()).unwrap_or(u32::MAX);
            let numeric_id = u32::try_from(state.objects.len())
                .map_err(|_| format_error(part, "object count exceeds supported range"))?;
            let z = state.take_z();
            state.objects.push(Object {
                numeric_id,
                parent_numeric_id: frame.parent_numeric_id,
                stable_id: format!("object:{numeric_id}"),
                parent_stable_id: frame
                    .parent_numeric_id
                    .map(|parent| format!("object:{parent}")),
                kind: ObjectKind::TextBox,
                unit_index,
                bounds: Rect {
                    x: bounds.x + source.bounds.x * scale_x,
                    y: bounds.y + source.bounds.y * scale_y,
                    width: source.bounds.width * scale_x,
                    height: source.bounds.height * scale_y,
                },
                z,
                text: Some(text),
                source: SourceRef {
                    part: source_part.clone(),
                    mapping,
                    locator: SourceLocator::PptxShape {
                        shape_id,
                        row: None,
                        column: None,
                        text_range: Some((0, text_length)),
                        metadata: metadata.clone(),
                    },
                },
                visual: scale_embedded_text_visual(source.visual, text_scale),
            });
        }
        if rendered {
            return Ok(());
        }
    }

    let Some(label) = frame.placeholder else {
        return Ok(());
    };
    if state.objects.len() >= object_limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "document exceeds the configured object limit",
        )
        .in_part(part));
    }
    state.diagnostics.push(
        Diagnostic::warning(
            DiagnosticCode::UnsupportedFeature,
            Phase::Render,
            Fidelity::Approximate,
            format!("PPTX {label} content uses a source-mapped static placeholder"),
        )
        .in_part(part),
    );
    let numeric_id = u32::try_from(state.objects.len())
        .map_err(|_| format_error(part, "object count exceeds supported range"))?;
    let z = state.take_z();
    state.objects.push(Object {
        numeric_id,
        parent_numeric_id: frame.parent_numeric_id,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: frame
            .parent_numeric_id
            .map(|parent| format!("object:{parent}")),
        kind: ObjectKind::Unknown,
        unit_index,
        bounds,
        z,
        text: Some(label.to_owned()),
        source: SourceRef {
            part: part.to_owned(),
            mapping,
            locator: SourceLocator::PptxShape {
                shape_id,
                row: None,
                column: None,
                text_range: Some((0, u32::try_from(label.len()).unwrap_or(u32::MAX))),
                metadata,
            },
        },
        visual: Visual::RichText {
            geometry: Geometry::RoundedRectangle {
                radius_x: 6.0,
                radius_y: 6.0,
            },
            fill: Paint::Solid(0xf3f4_f6ff),
            stroke: Paint::Solid(0x9ca3_afff),
            stroke_width: 1.0,
            align: TextAlign::Center,
            line_height: 0.0,
            runs: vec![TextRun {
                paint: None,
                east_asian_line_breaks: true,
                text: label.to_owned(),
                font_family: "Arial".to_owned(),
                font_size: 14.0 * POINTS_TO_CSS_PIXELS,
                color: 0x3741_51ff,
                bold: true,
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

pub(super) fn scale_embedded_text_visual(visual: Visual, scale: f32) -> Visual {
    match visual {
        Visual::TextLayout { mut layout, visual } => {
            for value in [
                &mut layout.inset_left,
                &mut layout.inset_right,
                &mut layout.inset_top,
                &mut layout.inset_bottom,
                &mut layout.margin_left,
                &mut layout.margin_right,
                &mut layout.default_tab_stop,
                &mut layout.hanging_indent,
                &mut layout.paragraph_spacing,
                &mut layout.first_line_indent,
                &mut layout.column_spacing,
                &mut layout.text_stroke_width,
                &mut layout.text_baseline,
            ] {
                *value *= scale;
            }
            for paragraph in &mut layout.paragraphs {
                paragraph.margin_left *= scale;
                paragraph.margin_right *= scale;
                paragraph.first_line_indent *= scale;
                paragraph.default_tab_stop *= scale;
                paragraph.line_height *= scale;
                paragraph.space_before *= scale;
                paragraph.space_after *= scale;
            }
            Visual::TextLayout {
                layout,
                visual: Box::new(scale_embedded_text_visual(*visual, scale)),
            }
        }
        Visual::RichText {
            geometry,
            fill,
            stroke,
            stroke_width,
            align,
            line_height,
            mut runs,
        } => {
            for run in &mut runs {
                run.font_size *= scale;
                run.baseline_shift *= scale;
                run.letter_spacing *= scale;
            }
            Visual::RichText {
                geometry,
                fill,
                stroke,
                stroke_width: stroke_width * scale,
                align,
                line_height: line_height * scale,
                runs,
            }
        }
        Visual::Group { children } => Visual::Group {
            children: children
                .into_iter()
                .map(|mut child| {
                    child.visual = scale_embedded_text_visual(child.visual, scale);
                    child
                })
                .collect(),
        },
        visual => visual,
    }
}

#[allow(clippy::too_many_arguments)]
fn push_basic_pptx_chart(
    mut chart: BasicChart,
    target_bounds: Rect,
    slide_part: &str,
    unit_index: u32,
    parent_numeric_id: Option<u32>,
    shape_id: u32,
    mapping: MappingQuality,
    metadata: PptxObjectMetadata,
    state: &mut PptxParseState,
    object_limit: usize,
) -> Result<(), Diagnostic> {
    let (bounds, transform) = chart
        .native_size
        .filter(|(width, height)| {
            width.is_finite() && height.is_finite() && *width > 0.0 && *height > 0.0
        })
        .map_or((target_bounds, None), |(width, height)| {
            (
                Rect {
                    x: 0.0,
                    y: 0.0,
                    width,
                    height,
                },
                Some(AffineTransform {
                    a: target_bounds.width / width,
                    b: 0.0,
                    c: 0.0,
                    d: target_bounds.height / height,
                    e: target_bounds.x,
                    f: target_bounds.y,
                }),
            )
        });
    if let (Some(font_size), Some((basis_width, basis_height))) =
        (chart.font_size, chart.font_scale_basis)
    {
        let scale = (bounds.width * 15.0 / basis_width).min(bounds.height * 15.0 / basis_height);
        let twips = (font_size * 15.0 * scale).max(10.0);
        chart.font_size = Some((twips / 5.0).round() * 5.0 / 15.0);
    }
    chart.resolve_value_axis_layout(bounds);
    let point_count = chart
        .series
        .iter()
        .map(|series| series.values.len())
        .sum::<usize>();
    let data_table_object_count = chart.data_table.map_or(0, |_| {
        let categories = chart
            .series
            .iter()
            .map(|series| series.categories.len().max(series.values.len()))
            .max()
            .unwrap_or(0);
        categories
            .saturating_mul(chart.series.len().saturating_add(1))
            .saturating_mul(2)
            .saturating_add(categories)
            .saturating_add(chart.series.len().saturating_mul(3))
            .saturating_add(8)
    });
    let required = state
        .objects
        .len()
        .checked_add(point_count.saturating_mul(4))
        .and_then(|count| count.checked_add(chart.series.len()))
        .and_then(|count| count.checked_add(data_table_object_count))
        .and_then(|count| count.checked_add(8))
        .ok_or_else(|| format_error(slide_part, "chart object count exceeds supported range"))?;
    if required > object_limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "document exceeds the configured object limit",
        )
        .in_part(slide_part));
    }
    let chart_id = u32::try_from(state.objects.len())
        .map_err(|_| format_error(slide_part, "object count exceeds supported range"))?;
    let z = state.take_z();
    let (area_stroke, area_stroke_width) = chart.area_border(bounds, (Paint::None, 0.0));
    state.objects.push(Object {
        numeric_id: chart_id,
        parent_numeric_id,
        stable_id: format!("object:{chart_id}"),
        parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
        kind: ObjectKind::Group,
        unit_index,
        bounds,
        z,
        text: None,
        source: SourceRef {
            part: slide_part.to_owned(),
            mapping,
            locator: SourceLocator::PptxShape {
                shape_id,
                row: None,
                column: None,
                text_range: None,
                metadata,
            },
        },
        visual: transform.map_or_else(
            || Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: chart
                    .chart_area_fill
                    .as_ref()
                    .map_or_else(|| Paint::None, |fill| fill.paint(bounds)),
                stroke: area_stroke.clone(),
                stroke_width: area_stroke_width,
            },
            |transform| Visual::Layer {
                transform,
                opacity: 1.0,
                blend_mode: crate::model::BlendMode::Normal,
                visual: Box::new(Visual::PaintedShape {
                    geometry: Geometry::Rectangle,
                    fill: chart
                        .chart_area_fill
                        .as_ref()
                        .map_or_else(|| Paint::None, |fill| fill.paint(bounds)),
                    stroke: area_stroke.clone(),
                    stroke_width: area_stroke_width,
                }),
            },
        ),
    });

    if let Some(elements) = super::drawingml::chart_extended_elements(&chart, bounds, object_limit)?
    {
        if state.objects.len().saturating_add(elements.len()) > object_limit {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ObjectLimit,
                Phase::Parse,
                None,
                "chart elements exceed the configured object limit",
            )
            .in_part(&chart.source_part));
        }
        for element in elements {
            let numeric_id = state.objects.len() as u32;
            let z = state.take_z();
            state.objects.push(Object {
                numeric_id,
                parent_numeric_id: Some(chart_id),
                stable_id: format!("object:{numeric_id}"),
                parent_stable_id: Some(format!("object:{chart_id}")),
                kind: if element.text.is_some() {
                    ObjectKind::TextBox
                } else {
                    ObjectKind::Shape
                },
                unit_index,
                bounds: element.bounds,
                z,
                text: element.text,
                source: SourceRef {
                    part: chart.source_part.clone(),
                    mapping: MappingQuality::Derived,
                    locator: SourceLocator::PptxShape {
                        shape_id,
                        row: None,
                        column: None,
                        text_range: None,
                        metadata: PptxObjectMetadata::default(),
                    },
                },
                visual: element.visual,
            });
        }
        return Ok(());
    }

    let cartesian = chart.series.iter().any(|series| {
        matches!(
            series.kind,
            BasicChartKind::Bar
                | BasicChartKind::Line
                | BasicChartKind::Area
                | BasicChartKind::Scatter
                | BasicChartKind::Bubble
        )
    });
    let bar_of_pie = chart
        .series
        .iter()
        .any(|series| matches!(series.kind, BasicChartKind::BarOfPie(_)));
    if chart.show_title {
        let title_font_size = 18.0 * POINTS_TO_CSS_PIXELS;
        let title_bounds = chart.positioned_title_bounds(
            bounds,
            chart_title_bounds(
                bounds,
                bounds.y + 8.0,
                28.0,
                chart.title_text(),
                title_font_size,
            ),
        );
        if let Some(fill) = chart.title_fill.as_ref() {
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                title_bounds,
                Geometry::Rectangle,
                fill.paint(title_bounds),
                Paint::None,
                0.0,
                &chart.source_part,
                shape_id,
                None,
                None,
                MappingQuality::Exact,
            )?;
        }
        push_pptx_chart_text(
            state,
            chart_id,
            unit_index,
            title_bounds,
            chart.title_text().to_owned(),
            0x5959_59ff,
            &chart,
            shape_id,
            None,
            title_font_size,
            false,
            0.0,
            true,
            TextAlign::Center,
        )?;
    }
    let horizontal_bars = chart
        .series
        .iter()
        .any(|series| series.kind == BasicChartKind::Bar && series.bar_horizontal);
    let legend_bottom = chart.legend_position == ChartLegendPosition::Bottom;
    let legend_top = chart.legend_position == ChartLegendPosition::Top;
    let legend_left = chart.legend_position == ChartLegendPosition::Left;
    let legend_top_right = chart.legend_position == ChartLegendPosition::TopRight;
    let lateral_legend = chart.show_legend && !legend_bottom && !legend_top;
    let legacy_msgraph = chart.native_size.is_some();
    let legacy_pie = legacy_msgraph
        && chart
            .series
            .iter()
            .any(|series| matches!(series.kind, BasicChartKind::Pie | BasicChartKind::Doughnut));
    let has_data_table = chart.data_table.is_some();
    let plain_scatter = cartesian
        && chart
            .series
            .iter()
            .all(|series| series.kind == BasicChartKind::Scatter)
        && !chart.show_legend;
    let horizontal_category_font =
        chart.axis_label_font_size(chart.category_axis_options(), 8.0 * POINTS_TO_CSS_PIXELS);
    let horizontal_category_margin = chart
        .series
        .iter()
        .filter(|series| series.kind == BasicChartKind::Bar && series.bar_horizontal)
        .flat_map(|series| &series.categories)
        .map(|category| {
            category
                .chars()
                .map(|c| drawingml_fallback_character_width(c, horizontal_category_font))
                .sum::<f32>()
        })
        .reduce(f32::max)
        .map_or(0.18, |width| {
            ((width + 18.0) / bounds.width).clamp(0.18, 0.40)
        });
    let (plot_x, plot_y, plot_width, plot_height) = if legacy_msgraph && bar_of_pie {
        (0.111, 0.167, 0.647, 0.664)
    } else if legacy_msgraph && cartesian {
        (0.148, 0.08, 0.502, 0.70)
    } else {
        (
            if plain_scatter {
                0.08
            } else if legend_left {
                0.30
            } else if has_data_table {
                0.17
            } else if horizontal_bars {
                horizontal_category_margin
            } else if !chart.value_axis_options.title.trim().is_empty() {
                0.16
            } else {
                0.12
            },
            if plain_scatter {
                0.02
            } else if legend_top {
                0.20
            } else if chart.show_title {
                0.16
            } else if has_data_table {
                0.05
            } else {
                0.08
            },
            if plain_scatter {
                0.90
            } else if horizontal_bars && lateral_legend {
                (0.80 - horizontal_category_margin).clamp(0.50, 0.58)
            } else if lateral_legend {
                if chart.font_size.is_some() {
                    0.58
                } else if bar_of_pie {
                    0.70
                } else {
                    0.62
                }
            } else {
                if has_data_table { 0.80 } else { 0.82 }
            },
            if plain_scatter {
                0.86
            } else if cartesian
                && (!chart.horizontal_axis_options.title.trim().is_empty()
                    || chart
                        .horizontal_axis_options
                        .label_rotation_degrees
                        .is_some_and(|rotation| rotation.abs() >= 45.0))
            {
                0.58
            } else if cartesian {
                0.70
            } else {
                0.72
            },
        )
    };
    let plot = chart.plot_area_bounds(
        bounds,
        Rect {
            x: bounds.x + bounds.width * plot_x,
            y: bounds.y + bounds.height * plot_y,
            width: bounds.width * plot_width,
            height: bounds.height * plot_height,
        },
    );
    if let Some(fill) = chart.plot_area_fill.as_ref() {
        push_pptx_chart_shape(
            state,
            chart_id,
            unit_index,
            plot,
            Geometry::Rectangle,
            fill.paint(plot),
            Paint::None,
            0.0,
            &chart.source_part,
            shape_id,
            None,
            None,
            MappingQuality::Exact,
        )?;
    } else if let Some(color) = chart.plot_area_color {
        push_pptx_chart_shape(
            state,
            chart_id,
            unit_index,
            plot,
            Geometry::Rectangle,
            Paint::Solid(color),
            Paint::None,
            0.0,
            &chart.source_part,
            shape_id,
            None,
            None,
            MappingQuality::Derived,
        )?;
    }
    if let Some(color) = chart.plot_area_border_color.filter(|_| !legacy_pie) {
        push_pptx_chart_shape(
            state,
            chart_id,
            unit_index,
            plot,
            Geometry::Rectangle,
            Paint::None,
            Paint::Solid(color),
            chart.plot_area_border_width,
            &chart.source_part,
            shape_id,
            None,
            None,
            MappingQuality::Exact,
        )?;
    }
    let has_3d_bars = chart
        .series
        .iter()
        .any(|series| series.kind == BasicChartKind::Bar && series.three_d);
    let has_3d_area = chart
        .series
        .iter()
        .any(|series| series.kind == BasicChartKind::Area && series.three_d);
    let view_3d = chart.view_3d.unwrap_or_default();
    if cartesian {
        let scatter_only = chart
            .series
            .iter()
            .all(|series| series.kind == BasicChartKind::Scatter);
        let (minimum, maximum, major) = if scatter_only {
            chart.scatter_axis(false)
        } else {
            chart.value_axis()
        };
        if maximum > minimum && major > 0.0 {
            let numerical_axis = chart.numerical_axis_options();
            let axis_font_size =
                chart.axis_label_font_size(numerical_axis, 9.0 * POINTS_TO_CSS_PIXELS);
            let axis_label_height = axis_font_size * 1.25;
            let axis_bold = numerical_axis.label_bold.unwrap_or(false);
            for (value, ratio) in chart.value_axis_ticks() {
                let y = chart
                    .depth_axis_point(plot, 0.0, ratio, 0.0)
                    .map_or(plot.y + plot.height - plot.height * ratio, |p| p.1);
                let x = plot.x + plot.width * ratio;
                let (grid_bounds, grid_geometry) =
                    chart.bar_grid_line(plot, ratio, has_3d_bars, horizontal_bars);
                if chart.value_axis_grid_lines_visible()
                    && !(has_data_table
                        && !horizontal_bars
                        && (value - minimum).abs() <= major * 0.01)
                {
                    push_pptx_chart_shape(
                        state,
                        chart_id,
                        unit_index,
                        grid_bounds,
                        grid_geometry,
                        Paint::None,
                        Paint::Solid(chart.value_axis_grid_style(0xd9d9_d9ff, 0.75).0),
                        chart.value_axis_grid_style(0xd9d9_d9ff, 0.75).1,
                        &chart.source_part,
                        shape_id,
                        None,
                        None,
                        MappingQuality::Derived,
                    )?;
                }
                if let Some(tick_bounds) =
                    chart.value_axis_tick_bounds(plot, ratio, horizontal_bars)
                {
                    push_pptx_chart_shape(
                        state,
                        chart_id,
                        unit_index,
                        tick_bounds,
                        Geometry::Line,
                        Paint::None,
                        Paint::Solid(0x5959_59ff),
                        0.75,
                        &chart.source_part,
                        shape_id,
                        None,
                        None,
                        MappingQuality::Derived,
                    )?;
                }
                if !chart.axis_labels_visible(horizontal_bars) {
                    continue;
                }
                let axis_format = numerical_axis.number_format.as_deref();
                push_pptx_chart_text(
                    state,
                    chart_id,
                    unit_index,
                    if horizontal_bars {
                        Rect {
                            x: x - 20.0,
                            y: if chart.value_axis_at_top() {
                                plot.y - axis_label_height - 4.0
                            } else {
                                plot.y + plot.height + 4.0
                            },
                            width: 40.0,
                            height: axis_label_height,
                        }
                    } else {
                        let x = bounds.x
                            + if chart.value_axis_options.title.trim().is_empty() {
                                0.0
                            } else {
                                24.0
                            };
                        Rect {
                            x,
                            y: y - axis_label_height / 2.0,
                            width: (chart
                                .depth_axis_point(plot, 0.0, ratio, 0.0)
                                .map_or(plot.x, |p| p.0)
                                - x
                                - 6.0)
                                .max(1.0),
                            height: axis_label_height,
                        }
                    },
                    chart.value_axis_label(value),
                    if value < 0.0 && axis_format.is_some_and(|format| format.contains("[Red]")) {
                        0xff00_00ff
                    } else {
                        0x0000_00ff
                    },
                    &chart,
                    shape_id,
                    None,
                    axis_font_size,
                    axis_bold,
                    0.0,
                    false,
                    TextAlign::End,
                )?;
            }
        }
        for (horizontal, line) in [
            (
                false,
                Rect {
                    x: plot.x,
                    y: plot.y,
                    width: 0.01,
                    height: plot.height,
                },
            ),
            (
                true,
                Rect {
                    x: plot.x,
                    y: if horizontal_bars && chart.value_axis_at_top() {
                        plot.y
                    } else {
                        plot.y + plot.height
                    },
                    width: plot.width,
                    height: 0.01,
                },
            ),
        ] {
            if !chart.axis_line_visible(horizontal) {
                continue;
            }
            let (line, geometry) = chart
                .projected_axis_line(plot, horizontal)
                .unwrap_or((line, Geometry::Line));
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                line,
                geometry,
                Paint::None,
                Paint::Solid(0x5959_59ff),
                0.75,
                &chart.source_part,
                shape_id,
                None,
                None,
                MappingQuality::Derived,
            )?;
        }
        for (bounds, geometry, color, width) in chart.minor_axis_lines(plot) {
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                bounds,
                geometry,
                Paint::None,
                Paint::Solid(color),
                width,
                &chart.source_part,
                shape_id,
                None,
                None,
                MappingQuality::Derived,
            )?;
        }
        let extra_style = chart.axis_title_text_style(&chart.horizontal_axis_options);
        for (label_bounds, text) in chart.supplemental_axis_labels(plot, extra_style.font_size) {
            push_pptx_chart_text(
                state,
                chart_id,
                unit_index,
                label_bounds,
                text,
                0x0000_00ff,
                &chart,
                shape_id,
                None,
                extra_style.font_size,
                extra_style.bold,
                0.0,
                false,
                TextAlign::Center,
            )?;
        }
        if !chart.value_axis_options.title.trim().is_empty() {
            let axis = &chart.value_axis_options;
            let style = chart.axis_title_text_style(axis);
            push_pptx_chart_text(
                state,
                chart_id,
                unit_index,
                chart
                    .projected_value_title_bounds(plot, style.font_size)
                    .unwrap_or(Rect {
                        x: bounds.x + 20.0,
                        y: plot.y,
                        width: 24.0,
                        height: plot.height,
                    }),
                axis.title.clone(),
                0x0000_00ff,
                &chart,
                shape_id,
                None,
                style.font_size,
                style.bold,
                -90.0,
                false,
                TextAlign::Center,
            )?;
        }
        if !chart.horizontal_axis_options.title.trim().is_empty() {
            let axis = &chart.horizontal_axis_options;
            let style = chart.axis_title_text_style(axis);
            let font_size = style.font_size;
            push_pptx_chart_text(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: plot.x,
                    y: bounds.y + bounds.height - font_size * 1.75,
                    width: plot.width,
                    height: font_size * 1.25,
                },
                axis.title.clone(),
                0x0000_00ff,
                &chart,
                shape_id,
                None,
                font_size,
                style.bold,
                0.0,
                false,
                TextAlign::Center,
            )?;
        }
    }

    let mut bar_series = chart
        .series
        .iter()
        .enumerate()
        .filter(|(_, series)| series.kind == BasicChartKind::Bar)
        .collect::<Vec<_>>();
    if bar_series.iter().any(|(_, series)| series.bar_depth)
        && f32::from(chart.view_3d.unwrap_or_default().rot_y)
            .to_radians()
            .cos()
            >= 0.0
    {
        bar_series.reverse();
    }
    if has_3d_bars {
        for (wall_bounds, wall_geometry, wall_fill) in chart_bar_3d_walls(plot, view_3d) {
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                wall_bounds,
                wall_geometry,
                wall_fill,
                Paint::None,
                0.0,
                &chart.source_part,
                shape_id,
                None,
                None,
                MappingQuality::Derived,
            )?;
        }
        let (floor_bounds, floor_geometry, floor_fill) = chart.bar_floor(plot);
        push_pptx_chart_shape(
            state,
            chart_id,
            unit_index,
            floor_bounds,
            floor_geometry,
            floor_fill,
            Paint::Solid(0x2222_22ff),
            0.75,
            &chart.source_part,
            shape_id,
            None,
            None,
            MappingQuality::Derived,
        )?;
    }
    // Filled areas belong behind columns in combination charts, even though
    // their geometry is emitted later by the per-kind renderers.
    let area_layer_z = (!bar_series.is_empty()).then(|| state.take_z());
    let bar_category_count = bar_series
        .iter()
        .map(|(_, series)| series.values.len())
        .max()
        .unwrap_or(0);
    if bar_category_count > 0 {
        for (series_index, series) in &bar_series {
            let (minimum, maximum, _) = chart.value_axis();
            if maximum <= minimum {
                continue;
            }
            let color = series
                .color
                .unwrap_or_else(|| super::office_chart_palette_color(*series_index));
            for (category_index, value) in series.values.iter().copied().enumerate() {
                if !value.is_finite() || value == 0.0 {
                    continue;
                }
                let Some(mut bar_bounds) = chart_bar_segment_bounds(
                    &chart,
                    *series_index,
                    category_index,
                    plot,
                    (minimum, maximum),
                    chart.bar_gap_width_percent,
                ) else {
                    continue;
                };
                if series.three_d && !series.projected_column() {
                    bar_bounds =
                        chart_bar_3d_plane_bounds(bar_bounds, plot, view_3d, series.bar_horizontal);
                }
                push_pptx_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    bar_bounds,
                    if let Some([front, _, _]) =
                        chart.depth_box_faces(*series_index, category_index, plot, bar_bounds)
                    {
                        front
                    } else if series.bar_cone_to_max {
                        chart_bar_cone_to_max_geometry(
                            bar_bounds,
                            series.bar_horizontal,
                            value >= 0.0,
                            value.abs()
                                / if value >= 0.0 { maximum } else { minimum.abs() }
                                    .max(f32::MIN_POSITIVE),
                        )
                    } else {
                        chart_bar_geometry(
                            bar_bounds,
                            series.bar_horizontal,
                            series.bar_cone,
                            series.bar_cylinder,
                        )
                    },
                    series.bar_paint(category_index, bar_bounds, color),
                    series.point_stroke(category_index),
                    series.point_stroke_width(category_index),
                    &chart.source_part,
                    shape_id,
                    Some(u32::try_from(*series_index).unwrap_or(u32::MAX)),
                    Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                    MappingQuality::Exact,
                )?;
                series.apply_effects(state.objects.last_mut());
                if series.bar_cone_to_max
                    && let Some((cap_bounds, cap_geometry, cap_fill)) = chart_bar_cone_cap(
                        bar_bounds,
                        color,
                        series.bar_horizontal,
                        value >= 0.0,
                        value.abs()
                            / if value >= 0.0 { maximum } else { minimum.abs() }
                                .max(f32::MIN_POSITIVE),
                    )
                {
                    push_pptx_chart_shape(
                        state,
                        chart_id,
                        unit_index,
                        cap_bounds,
                        cap_geometry,
                        cap_fill,
                        Paint::None,
                        0.0,
                        &chart.source_part,
                        shape_id,
                        Some(u32::try_from(*series_index).unwrap_or(u32::MAX)),
                        Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                        MappingQuality::Derived,
                    )?;
                } else if series.bar_cylinder {
                    let (cap_bounds, cap_geometry, cap_fill) = chart_bar_cylinder_cap(
                        bar_bounds,
                        color,
                        series.bar_horizontal,
                        value >= 0.0,
                    );
                    push_pptx_chart_shape(
                        state,
                        chart_id,
                        unit_index,
                        cap_bounds,
                        cap_geometry,
                        cap_fill,
                        Paint::None,
                        0.0,
                        &chart.source_part,
                        shape_id,
                        Some(u32::try_from(*series_index).unwrap_or(u32::MAX)),
                        Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                        MappingQuality::Derived,
                    )?;
                } else if series.three_d && !series.bar_cone {
                    for (face_bounds, face_geometry, face_fill) in
                        chart.bar_faces(*series_index, category_index, plot, bar_bounds, color)
                    {
                        push_pptx_chart_shape(
                            state,
                            chart_id,
                            unit_index,
                            face_bounds,
                            face_geometry,
                            face_fill,
                            Paint::Solid(0x4444_44ff),
                            0.5,
                            &chart.source_part,
                            shape_id,
                            Some(u32::try_from(*series_index).unwrap_or(u32::MAX)),
                            Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                            MappingQuality::Derived,
                        )?;
                    }
                }
                if let Some((bounds, text, mut style, border)) =
                    chart_bar_data_label(&chart, series, category_index, bar_bounds, value)
                {
                    if horizontal_bars
                        && bar_category_count > 6
                        && series.data_label_font_size.is_none()
                        && !series
                            .data_labels
                            .iter()
                            .any(|label| label.index == category_index)
                    {
                        style.font_size = 8.0 * POINTS_TO_CSS_PIXELS;
                    }
                    if let Some((color, width)) = border {
                        push_pptx_chart_shape(
                            state,
                            chart_id,
                            unit_index,
                            bounds,
                            Geometry::Rectangle,
                            Paint::None,
                            Paint::Solid(color),
                            width,
                            &chart.source_part,
                            shape_id,
                            Some(*series_index as u32),
                            Some(category_index as u32),
                            MappingQuality::Derived,
                        )?;
                    }
                    push_pptx_chart_text(
                        state,
                        chart_id,
                        unit_index,
                        bounds,
                        text,
                        style.color,
                        &chart,
                        shape_id,
                        Some(category_index as u32),
                        style.font_size,
                        style.bold,
                        series.data_label_rotation_degrees.unwrap_or(0.0),
                        !series
                            .data_labels
                            .iter()
                            .any(|label| label.index == category_index && !label.text.is_empty()),
                        TextAlign::Center,
                    )?;
                }
            }
        }
    }
    let series_axis_style = chart.legend_text_style();
    for (series_index, label_bounds, label) in
        chart_3d_series_axis_labels(&chart, plot, series_axis_style.font_size)
    {
        push_pptx_chart_text(
            state,
            chart_id,
            unit_index,
            label_bounds,
            label,
            series_axis_style.color,
            &chart,
            shape_id,
            Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
            series_axis_style.font_size,
            series_axis_style.bold,
            0.0,
            false,
            TextAlign::End,
        )?;
    }
    let (series_labels, category_labels) =
        chart_area_3d_axis_labels(&chart, plot, series_axis_style.font_size);
    for (series_index, label_bounds, label, rotation) in series_labels {
        push_pptx_chart_text(
            state,
            chart_id,
            unit_index,
            label_bounds,
            label,
            series_axis_style.color,
            &chart,
            shape_id,
            Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
            series_axis_style.font_size,
            series_axis_style.bold,
            rotation,
            false,
            TextAlign::Center,
        )?;
    }
    for (category_index, label_bounds, label, rotation) in category_labels {
        push_pptx_chart_text(
            state,
            chart_id,
            unit_index,
            label_bounds,
            label,
            series_axis_style.color,
            &chart,
            shape_id,
            Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
            series_axis_style.font_size,
            series_axis_style.bold,
            rotation,
            false,
            TextAlign::Start,
        )?;
    }

    let (minimum, maximum, _) = chart.value_axis();
    for (index, (line_bounds, line_geometry)) in
        chart_series_line_geometries(&chart, plot, (minimum, maximum))
            .into_iter()
            .enumerate()
    {
        push_pptx_chart_shape(
            state,
            chart_id,
            unit_index,
            line_bounds,
            line_geometry,
            Paint::None,
            Paint::Solid(0x7f7f_7fff),
            0.75,
            &chart.source_part,
            shape_id,
            None,
            Some(u32::try_from(index).unwrap_or(u32::MAX)),
            MappingQuality::Derived,
        )?;
    }

    if cartesian
        && chart.axis_labels_visible(!horizontal_bars)
        && !has_3d_area
        && (!has_data_table || chart.series.iter().any(|s| s.projected_column()))
        && let Some(series) = chart
            .series
            .iter()
            .find(|series| !series.categories.is_empty() && !series.values.is_empty())
    {
        let category_font_size =
            chart.axis_label_font_size(chart.category_axis_options(), 8.0 * POINTS_TO_CSS_PIXELS);
        let category_font_size = chart.depth_category_font_size(series, plot, category_font_size);
        let category_count = series.categories.len();
        let visible_categories = series
            .categories
            .iter()
            .take(category_count)
            .enumerate()
            .filter(|(index, _)| chart.category_has_label(series, *index, plot, 18.0))
            .filter_map(|(index, category)| {
                chart
                    .category_x(series, index, plot, true)
                    .map(|x| (index, category, x))
            })
            .collect::<Vec<_>>();
        let category_width = plot.width / visible_categories.len().max(1) as f32;
        let rotate_categories = !horizontal_bars
            && !series.bar_depth
            && (visible_categories.len() > 6
                || chart_category_labels_overlap(
                    &series.categories,
                    category_width,
                    category_font_size,
                ));
        for index in 0..=category_count {
            let Some(tick_bounds) = chart.category_axis_tick_bounds(series, index, plot) else {
                continue;
            };
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                tick_bounds,
                Geometry::Line,
                Paint::None,
                Paint::Solid(0x5959_59ff),
                0.75,
                &chart.source_part,
                shape_id,
                None,
                Some(u32::try_from(index).unwrap_or(u32::MAX)),
                MappingQuality::Derived,
            )?;
        }
        for (index, category, x) in visible_categories {
            push_pptx_chart_text(
                state,
                chart_id,
                unit_index,
                if horizontal_bars {
                    let category_height = plot.height / category_count as f32;
                    Rect {
                        x: bounds.x,
                        y: plot.y
                            + category_count
                                .saturating_sub(chart.category_index(index, category_count) + 1)
                                as f32
                                * category_height,
                        width: plot.x - bounds.x - 6.0,
                        height: category_height,
                    }
                } else {
                    Rect {
                        x: x - category_width / 2.0,
                        y: chart.category_baseline(series, index, plot) + 8.0,
                        width: category_width,
                        height: (bounds.y + bounds.height - 28.0 - (plot.y + plot.height))
                            .max(18.0),
                    }
                },
                category.clone(),
                0x0000_00ff,
                &chart,
                shape_id,
                Some(u32::try_from(index).unwrap_or(u32::MAX)),
                category_font_size,
                false,
                chart
                    .horizontal_axis_options
                    .label_rotation_degrees
                    .unwrap_or(if rotate_categories { -45.0 } else { 0.0 }),
                false,
                TextAlign::Center,
            )?;
        }
    }

    for (series_index, series) in
        chart.series.iter().enumerate().filter(|(_, series)| {
            matches!(series.kind, BasicChartKind::Line | BasicChartKind::Area)
        })
    {
        let series_object_start = state.objects.len();
        let category_count = series.values.len();
        let (minimum, maximum, _) = chart.value_axis_for_series(series);
        if maximum <= minimum {
            continue;
        }
        if category_count < 2 {
            continue;
        }
        let color = series
            .color
            .unwrap_or_else(|| super::office_chart_palette_color(series_index));
        let (points, label_points) = if series.kind == BasicChartKind::Area && series.three_d {
            for (face_bounds, face_geometry, face_fill) in chart_area_3d_faces(
                &chart,
                series_index,
                plot,
                (minimum, maximum),
                view_3d,
                color,
            ) {
                push_pptx_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    face_bounds,
                    face_geometry,
                    face_fill,
                    Paint::Solid(0x0000_00ff),
                    series.stroke_width.unwrap_or(0.5),
                    &chart.source_part,
                    shape_id,
                    Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                    None,
                    MappingQuality::Derived,
                )?;
            }
            (Vec::new(), Vec::new())
        } else if series.kind == BasicChartKind::Area {
            let area_plot = chart.area_plot_bounds(series, plot).unwrap_or(plot);
            let Some((geometry, points, label_points)) =
                chart_area_geometry(&chart, series_index, area_plot, (minimum, maximum))
            else {
                continue;
            };
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                area_plot,
                geometry,
                Paint::Solid(color),
                series.point_stroke(0),
                series.point_stroke_width(0),
                &chart.source_part,
                shape_id,
                Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                None,
                MappingQuality::Derived,
            )?;
            series.apply_effects(state.objects.last_mut());
            (points, label_points)
        } else {
            let points = chart.line_points(
                series_index,
                plot,
                (minimum, maximum),
                has_data_table || chart.up_down_bars.is_some(),
            );
            (points.clone(), points)
        };
        if let Some((line_bounds, geometry)) = chart_line_geometry(&points) {
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                line_bounds,
                geometry,
                Paint::None,
                Paint::Solid(color),
                series.stroke_width.unwrap_or(2.0),
                &chart.source_part,
                shape_id,
                Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                None,
                MappingQuality::Exact,
            )?;
            series.apply_effects(state.objects.last_mut());
        }
        if let Some(symbol) = series.marker_symbol.as_deref() {
            let size = series.marker_size.unwrap_or(6.0 * POINTS_TO_CSS_PIXELS);
            for &(x, y, category_index) in &points {
                let marker_bounds = Rect {
                    x: x - size / 2.0,
                    y: y - size / 2.0,
                    width: size,
                    height: size,
                };
                let Some((geometry, filled)) = chart_marker_geometry(symbol, size) else {
                    continue;
                };
                push_pptx_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    marker_bounds,
                    geometry,
                    if filled {
                        Paint::Solid(color)
                    } else {
                        Paint::None
                    },
                    Paint::Solid(color),
                    1.0,
                    &chart.source_part,
                    shape_id,
                    Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                    Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                    MappingQuality::Exact,
                )?;
            }
        }
        for (x, y, category_index) in label_points {
            let Some(value) = series.values.get(category_index).copied() else {
                continue;
            };
            if let Some((label_bounds, text, style, border)) = chart_bar_data_label(
                &chart,
                series,
                category_index,
                Rect {
                    x,
                    y,
                    width: 0.0,
                    height: 0.0,
                },
                value,
            ) {
                if let Some((color, width)) = border {
                    push_pptx_chart_shape(
                        state,
                        chart_id,
                        unit_index,
                        label_bounds,
                        Geometry::Rectangle,
                        Paint::None,
                        Paint::Solid(color),
                        width,
                        &chart.source_part,
                        shape_id,
                        Some(series_index as u32),
                        Some(category_index as u32),
                        MappingQuality::Derived,
                    )?;
                }
                push_pptx_chart_text(
                    state,
                    chart_id,
                    unit_index,
                    label_bounds,
                    text,
                    style.color,
                    &chart,
                    shape_id,
                    Some(category_index as u32),
                    style.font_size,
                    style.bold,
                    series.data_label_rotation_degrees.unwrap_or(0.0),
                    false,
                    style.align,
                )?;
            }
        }
        if series.kind == BasicChartKind::Area
            && let Some(z) = area_layer_z
        {
            for object in &mut state.objects[series_object_start..] {
                if object.text.is_none() {
                    object.z = z;
                }
            }
        }
    }

    if let Some(options) = chart.up_down_bars.as_ref() {
        let (minimum, maximum, _) = chart.value_axis();
        for (category_index, bar_bounds, is_up) in
            chart.up_down_bar_bounds(plot, (minimum, maximum))
        {
            let fill = if is_up {
                options.up_fill.as_ref()
            } else {
                options.down_fill.as_ref()
            };
            let stroke = if is_up {
                options.up_stroke.as_ref()
            } else {
                options.down_stroke.as_ref()
            };
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                bar_bounds,
                Geometry::Rectangle,
                fill.map_or(
                    Paint::Solid(if is_up { 0xffff_ffff } else { 0x0000_00ff }),
                    |fill| fill.paint(bar_bounds),
                ),
                stroke.map_or(Paint::Solid(0x0000_00ff), |stroke| stroke.paint(bar_bounds)),
                1.0,
                &chart.source_part,
                shape_id,
                None,
                Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                MappingQuality::Exact,
            )?;
        }
    }

    let scatter_series = chart
        .series
        .iter()
        .enumerate()
        .filter(|(_, series)| series.kind == BasicChartKind::Scatter)
        .collect::<Vec<_>>();
    if !scatter_series.is_empty() {
        let (minimum, maximum, major) = chart.scatter_axis(true);
        let mut value = minimum;
        while value <= maximum + major * 0.01 {
            let x = plot.x + (value - minimum) / (maximum - minimum) * plot.width;
            push_pptx_chart_text(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: x - 24.0,
                    y: plot.y + plot.height + 4.0,
                    width: 48.0,
                    height: 18.0,
                },
                chart.horizontal_axis_options.format_label(value),
                0x0000_00ff,
                &chart,
                shape_id,
                None,
                9.0 * POINTS_TO_CSS_PIXELS,
                false,
                0.0,
                true,
                TextAlign::Center,
            )?;
            value += major;
        }
    }

    for (series_index, series) in scatter_series {
        let (minimum, maximum, _) = chart.scatter_axis(false);
        let (minimum_x, maximum_x, _) = chart.scatter_axis(true);
        let x_values = if series.x_values.len() == series.values.len() {
            series.x_values.clone()
        } else {
            (0..series.values.len()).map(|index| index as f32).collect()
        };
        let color = series
            .color
            .unwrap_or_else(|| super::office_chart_palette_color(series_index));
        if let Some((line_bounds, geometry)) = chart.scatter_line_geometry(series, plot) {
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                line_bounds,
                geometry,
                Paint::None,
                Paint::Solid(color),
                series.stroke_width.unwrap_or(2.0),
                &chart.source_part,
                shape_id,
                Some(series_index as u32),
                None,
                MappingQuality::Derived,
            )?;
        }
        for (point_index, element) in chart.scatter_error_bars(series, plot) {
            if let Visual::PaintedShape {
                geometry,
                fill,
                stroke,
                stroke_width,
            } = element.visual
            {
                push_pptx_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    element.bounds,
                    geometry,
                    fill,
                    stroke,
                    stroke_width,
                    &chart.source_part,
                    shape_id,
                    Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                    Some(u32::try_from(point_index).unwrap_or(u32::MAX)),
                    MappingQuality::Derived,
                )?;
            }
        }
        for (category_index, (x_value, y_value)) in x_values
            .iter()
            .copied()
            .zip(series.values.iter().copied())
            .enumerate()
        {
            let center_x = plot.x + (x_value - minimum_x) / (maximum_x - minimum_x) * plot.width;
            let center_y = chart_value_y(y_value, minimum, maximum, plot);
            let size = series.marker_size.unwrap_or(6.0 * POINTS_TO_CSS_PIXELS);
            let Some((geometry, filled)) =
                chart_marker_geometry(series.marker_symbol.as_deref().unwrap_or("none"), size)
            else {
                continue;
            };
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: center_x - size / 2.0,
                    y: center_y - size / 2.0,
                    width: size,
                    height: size,
                },
                geometry,
                if filled {
                    Paint::Solid(color)
                } else {
                    Paint::None
                },
                Paint::Solid(if filled { 0xffff_ffff } else { color }),
                1.0,
                &chart.source_part,
                shape_id,
                Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                Some(u32::try_from(category_index).unwrap_or(u32::MAX)),
                MappingQuality::Exact,
            )?;
        }
        if series.linear_trendline
            && let Some((slope, intercept, r_squared)) =
                super::linear_regression(&x_values, &series.values)
        {
            let raw_minimum_x = x_values.iter().copied().fold(f32::INFINITY, f32::min);
            let raw_maximum_x = x_values.iter().copied().fold(f32::NEG_INFINITY, f32::max);
            let x0 = plot.x + (raw_minimum_x - minimum_x) / (maximum_x - minimum_x) * plot.width;
            let x1 = plot.x + (raw_maximum_x - minimum_x) / (maximum_x - minimum_x) * plot.width;
            let y0 = chart_value_y(slope * raw_minimum_x + intercept, minimum, maximum, plot);
            let y1 = chart_value_y(slope * raw_maximum_x + intercept, minimum, maximum, plot);
            let line_bounds = Rect {
                x: x0.min(x1),
                y: y0.min(y1),
                width: (x1 - x0).abs(),
                height: (y1 - y0).abs(),
            };
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                line_bounds,
                Geometry::Path {
                    fill_rule: FillRule::NonZero,
                    commands: vec![
                        PathCommand::MoveTo {
                            x: x0 - line_bounds.x,
                            y: y0 - line_bounds.y,
                        },
                        PathCommand::LineTo {
                            x: x1 - line_bounds.x,
                            y: y1 - line_bounds.y,
                        },
                    ],
                },
                Paint::None,
                Paint::Solid(0x0000_00ff),
                1.5,
                &chart.source_part,
                shape_id,
                Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                None,
                MappingQuality::Derived,
            )?;
            if series.show_trendline_equation || series.show_trendline_r_squared {
                let mut lines = Vec::new();
                if series.show_trendline_equation {
                    let operator = if intercept.is_sign_negative() {
                        "−"
                    } else {
                        "+"
                    };
                    lines.push(format!("y = {slope:.4}x {operator} {:.4}", intercept.abs()));
                }
                if series.show_trendline_r_squared {
                    lines.push(format!("R² = {}", format_axis_value(r_squared)));
                }
                let (offset_x, offset_y) = series.trendline_label_offset.unwrap_or((0.0, 0.0));
                let label_width = plot.width * 0.24;
                let label_height = series
                    .trendline_label_font_size
                    .unwrap_or(9.0 * POINTS_TO_CSS_PIXELS)
                    .max(9.0 * POINTS_TO_CSS_PIXELS)
                    * 2.0;
                push_pptx_chart_text(
                    state,
                    chart_id,
                    unit_index,
                    Rect {
                        x: (plot.x + plot.width * (0.65 + offset_x))
                            .clamp(plot.x, plot.x + plot.width - label_width),
                        y: (plot.y + plot.height * (0.2 + offset_y))
                            .clamp(plot.y, plot.y + plot.height - label_height),
                        width: label_width,
                        height: label_height,
                    },
                    lines.join("\n"),
                    0x0000_00ff,
                    &chart,
                    shape_id,
                    None,
                    series
                        .trendline_label_font_size
                        .unwrap_or(9.0 * POINTS_TO_CSS_PIXELS),
                    false,
                    0.0,
                    true,
                    TextAlign::Center,
                )?;
            }
        }
    }

    for (series_index, series) in chart
        .series
        .iter()
        .enumerate()
        .filter(|(_, series)| series.kind == BasicChartKind::Bubble)
    {
        let color = series
            .color
            .unwrap_or_else(|| super::office_chart_palette_color(series_index));
        for (bubble, point_index) in chart_bubble_bounds(&chart, series_index, plot) {
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                bubble,
                Geometry::Ellipse,
                chart_bubble_paint(bubble, color, series.three_d),
                Paint::Solid(color),
                1.0,
                &chart.source_part,
                shape_id,
                Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                Some(u32::try_from(point_index).unwrap_or(u32::MAX)),
                MappingQuality::Exact,
            )?;
        }
    }

    // A mixed chart shares the host's Cartesian plot; standalone radar charts
    // use the common radial scene above, with their own grid and label layout.
    for (series_index, series) in chart
        .series
        .iter()
        .enumerate()
        .filter(|(_, s)| s.kind == BasicChartKind::Radar)
    {
        let (_, maximum, _) = chart.value_axis();
        if let Some(geometry) = radar_geometry(&series.values, maximum, plot) {
            let color = series
                .color
                .unwrap_or_else(|| super::office_chart_palette_color(series_index));
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                plot,
                geometry,
                series
                    .fill
                    .as_ref()
                    .map_or(Paint::None, |fill| fill.paint(plot)),
                Paint::Solid(color),
                series.stroke_width.unwrap_or(2.0),
                &chart.source_part,
                shape_id,
                Some(series_index as u32),
                None,
                MappingQuality::Derived,
            )?;
        }
    }

    if let Some((series_index, series)) = chart.series.iter().enumerate().find(|(_, series)| {
        matches!(
            series.kind,
            BasicChartKind::Pie | BasicChartKind::BarOfPie(_) | BasicChartKind::Doughnut
        )
    }) {
        let total = series
            .values
            .iter()
            .copied()
            .filter(|value| value.is_finite() && *value > 0.0)
            .sum::<f32>();
        if total <= 0.0 {
            return Ok(());
        }
        let doughnut = series.kind == BasicChartKind::Doughnut;
        let bar_of_pie_split = match series.kind {
            BasicChartKind::BarOfPie(split) => Some(usize::from(split).min(series.values.len())),
            _ => None,
        };
        let three_d = series.three_d && !doughnut;
        let pie = if bar_of_pie_split.is_some() {
            plot
        } else {
            let fallback = if !legacy_msgraph && !chart.show_title && !chart.show_legend {
                Rect {
                    x: bounds.x + bounds.width * 0.04,
                    y: bounds.y + bounds.height * 0.04,
                    width: bounds.width * 0.92,
                    height: bounds.height * 0.92,
                }
            } else {
                plot
            };
            chart.pie_bounds(bounds, fallback, three_d)
        };
        let radius = if bar_of_pie_split.is_some() {
            bounds.width * 0.206
        } else {
            pie.width / 2.0
        };
        let vertical_radius = if bar_of_pie_split.is_some() {
            plot.height * 0.46
        } else if three_d {
            radius * chart.view_3d.unwrap_or_default().pie_vertical_ratio()
        } else {
            radius
        };
        let depth = if three_d {
            radius * chart.view_3d.unwrap_or_default().pie_depth_ratio()
        } else {
            0.0
        };
        let center_x = pie.x
            + if bar_of_pie_split.is_some() {
                radius * 0.8
            } else {
                pie.width / 2.0
            };
        let center_y = pie.y + (pie.height - depth) / 2.0;
        if legacy_pie && let Some(color) = chart.plot_area_border_color {
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: center_x - radius,
                    y: center_y - vertical_radius,
                    width: radius * 2.0,
                    height: vertical_radius * 2.0 + depth,
                },
                Geometry::Rectangle,
                Paint::None,
                Paint::Solid(color),
                chart.plot_area_border_width,
                &chart.source_part,
                shape_id,
                None,
                None,
                MappingQuality::Exact,
            )?;
        }
        if bar_of_pie_split.is_some() {
            let left = center_x - radius;
            let right = plot.x + plot.width * 0.97;
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: left,
                    y: center_y - vertical_radius,
                    width: right - left,
                    height: vertical_radius * 2.0,
                },
                Geometry::Rectangle,
                Paint::None,
                Paint::Solid(0x2222_22ff),
                1.0,
                &chart.source_part,
                shape_id,
                None,
                None,
                MappingQuality::Derived,
            )?;
        }
        let (layout_total, shared_slices) = chart_pie_slices(
            series,
            (center_x, center_y),
            (radius, vertical_radius),
            depth,
            chart.view_3d.unwrap_or_default().pie_depth_perspective(),
            bar_of_pie_split,
        );
        debug_assert!((layout_total - total).abs() <= f32::EPSILON);
        for slice in &shared_slices {
            for (side_bounds, side_geometry) in slice.side.iter().chain(&slice.cut_side) {
                push_pptx_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    *side_bounds,
                    side_geometry.clone(),
                    chart_pie_side_paint(slice.color, *side_bounds),
                    Paint::None,
                    0.0,
                    &chart.source_part,
                    shape_id,
                    Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                    Some(u32::try_from(slice.index).unwrap_or(u32::MAX)),
                    MappingQuality::Derived,
                )?;
            }
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                slice.bounds,
                slice.geometry.clone(),
                series
                    .point_fills
                    .get(slice.index)
                    .and_then(Option::as_ref)
                    .map_or(Paint::Solid(slice.color), |fill| fill.paint(slice.bounds)),
                if bar_of_pie_split.is_some() {
                    Paint::Solid(0x2222_22ff)
                } else {
                    series
                        .point_border_colors
                        .get(slice.index)
                        .copied()
                        .flatten()
                        .map_or(Paint::None, Paint::Solid)
                },
                if bar_of_pie_split.is_some() {
                    1.0
                } else {
                    series
                        .point_border_widths
                        .get(slice.index)
                        .copied()
                        .unwrap_or(0.0)
                },
                &chart.source_part,
                shape_id,
                Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                Some(u32::try_from(slice.index).unwrap_or(u32::MAX)),
                MappingQuality::Exact,
            )?;
            series.apply_effects(state.objects.last_mut());
        }
        for mut label in chart_pie_label_layout(
            &chart,
            series,
            &shared_slices,
            bounds,
            radius,
            vertical_radius,
        ) {
            if !legacy_msgraph {
                label.text = label.text.replace('\n', ", ");
                if let Some((x, y)) = series
                    .data_labels
                    .iter()
                    .find(|point| point.index == label.index)
                    .filter(|point| point.position.as_deref() == Some("bestFit"))
                    .and_then(|point| point.manual_offset)
                {
                    label.bounds.x += bounds.width * x;
                    label.bounds.y += bounds.height * y;
                }
            }
            if let Some((leader_bounds, leader_geometry)) = label.leader {
                push_pptx_chart_shape(
                    state,
                    chart_id,
                    unit_index,
                    leader_bounds,
                    leader_geometry,
                    Paint::None,
                    Paint::Solid(label.style.color),
                    0.75,
                    &chart.source_part,
                    shape_id,
                    Some(u32::try_from(series_index).unwrap_or(u32::MAX)),
                    Some(u32::try_from(label.index).unwrap_or(u32::MAX)),
                    MappingQuality::Derived,
                )?;
            }
            push_pptx_chart_text(
                state,
                chart_id,
                unit_index,
                label.bounds,
                label.text,
                label.style.color,
                &chart,
                shape_id,
                Some(u32::try_from(label.index).unwrap_or(u32::MAX)),
                label.style.font_size,
                label.style.bold,
                0.0,
                true,
                label.style.align,
            )?;
        }
    }
    if chart.data_table.is_some() {
        push_pptx_chart_data_table(state, chart_id, unit_index, bounds, plot, &chart, shape_id)?;
    }
    if chart.show_legend {
        let legend_style = chart.legend_text_style();
        let legend_font_size = if legacy_msgraph {
            chart
                .font_size
                .map_or(legend_style.font_size, |size| size * 2.0 / 3.0)
        } else {
            legend_style.font_size
        };
        let constrained = chart.constrained_legend(bounds);
        let legend_line_height = constrained
            .as_ref()
            .map_or(legend_font_size * 1.45, |(_, step, _)| *step);
        let legend_entries = constrained
            .as_ref()
            .map_or_else(|| chart.legend_entries(), |(_, _, entries)| entries.clone());
        let long_legend = legend_entries
            .iter()
            .any(|(_, label, _)| label.chars().count() > 8);
        let horizontal_legend = chart.legend_position.horizontal();
        let legend_item_widths = legend_entries
            .iter()
            .map(|(_, label, _)| {
                (label.chars().count() as f32 * legend_font_size * 0.58
                    + legend_font_size * 0.65
                    + 12.0)
                    .max(56.0)
            })
            .collect::<Vec<_>>();
        let default_legend_width = if horizontal_legend {
            legend_item_widths
                .iter()
                .sum::<f32>()
                .min(bounds.width * 0.90)
        } else {
            legend_item_widths
                .iter()
                .copied()
                .fold(48.0_f32, f32::max)
                .min(bounds.width * if long_legend { 0.30 } else { 0.26 })
        };
        let default_legend_height = if horizontal_legend {
            legend_line_height + 8.0
        } else {
            legend_entries.len() as f32 * legend_line_height + 8.0
        };
        let legend_bounds = constrained
            .as_ref()
            .map(|(bounds, _, _)| *bounds)
            .unwrap_or_else(|| {
                chart.legend_bounds.map_or(
                    chart
                        .automatic_cartesian_layout(bounds)
                        .and_then(|(_, legend)| legend)
                        .unwrap_or(Rect {
                            x: if horizontal_legend {
                                bounds.x + (bounds.width - default_legend_width) / 2.0
                            } else if legend_left {
                                bounds.x + 4.0
                            } else {
                                bounds.x + bounds.width - default_legend_width - 4.0
                            },
                            y: if legend_bottom {
                                bounds.y + bounds.height - default_legend_height - 6.0
                            } else if legend_top {
                                bounds.y + bounds.height * 0.10
                            } else if legend_top_right {
                                bounds.y + 8.0
                            } else {
                                bounds.y + (bounds.height - default_legend_height) / 2.0
                            },
                            width: default_legend_width,
                            height: default_legend_height,
                        }),
                    |legend| Rect {
                        x: bounds.x + bounds.width * legend.x,
                        y: bounds.y + bounds.height * legend.y,
                        width: bounds.width * legend.width,
                        height: bounds.height * legend.height,
                    },
                )
            });
        let legend_x = legend_bounds.x + 4.0;
        let legend_width = legend_bounds.width;
        let legend_swatch = legend_font_size * 0.65;
        let legend_y = legend_bounds.y;
        if chart.legend_fill.is_some() || chart.legend_stroke.is_some() {
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                legend_bounds,
                Geometry::Rectangle,
                chart
                    .legend_fill
                    .as_ref()
                    .map_or(Paint::None, |fill| fill.paint(legend_bounds)),
                chart
                    .legend_stroke
                    .as_ref()
                    .map_or(Paint::None, |stroke| stroke.paint(legend_bounds)),
                chart.legend_stroke_width,
                &chart.source_part,
                shape_id,
                None,
                None,
                MappingQuality::Exact,
            )?;
        } else if chart.legend_bounds.is_some() {
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                legend_bounds,
                Geometry::Rectangle,
                Paint::Solid(0xffff_ffff),
                Paint::Solid(0x6666_66ff),
                0.75,
                &chart.source_part,
                shape_id,
                None,
                None,
                MappingQuality::Derived,
            )?;
        }
        let mut horizontal_offset = 0.0;
        for (legend_slot, (legend_index, label, color)) in legend_entries.into_iter().enumerate() {
            if label.trim().is_empty() {
                continue;
            }
            let item_width = legend_item_widths.get(legend_slot).copied().unwrap_or(56.0);
            let item_x = if horizontal_legend {
                legend_x + horizontal_offset
            } else {
                legend_x
            };
            let item_y = if horizontal_legend {
                legend_y + 7.0
            } else {
                legend_y + 7.0 + legend_slot as f32 * legend_line_height
            };
            let legend_key_is_line = chart.legend_key_is_line(legend_index);
            let key_width = chart.legend_key_width(legend_index, legend_swatch);
            push_pptx_chart_shape(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: item_x,
                    y: item_y
                        + if legend_key_is_line {
                            legend_swatch / 2.0
                        } else {
                            0.0
                        },
                    width: key_width,
                    height: if legend_key_is_line {
                        0.01
                    } else {
                        legend_swatch
                    },
                },
                if legend_key_is_line {
                    Geometry::Line
                } else {
                    Geometry::Rectangle
                },
                if legend_key_is_line {
                    Paint::None
                } else {
                    Paint::Solid(color)
                },
                if legend_key_is_line {
                    Paint::Solid(color)
                } else {
                    Paint::None
                },
                if legend_key_is_line { 2.0 } else { 0.0 },
                &chart.source_part,
                shape_id,
                Some(u32::try_from(legend_index).unwrap_or(u32::MAX)),
                None,
                MappingQuality::Derived,
            )?;
            if let Some((symbol, size)) = chart.legend_marker(legend_index) {
                let size = size.clamp(2.0, legend_swatch * 1.25);
                if let Some((geometry, filled)) = chart_marker_geometry(symbol, size) {
                    push_pptx_chart_shape(
                        state,
                        chart_id,
                        unit_index,
                        Rect {
                            x: item_x + (key_width - size) / 2.0,
                            y: item_y + (legend_swatch - size) / 2.0,
                            width: size,
                            height: size,
                        },
                        geometry,
                        if filled {
                            Paint::Solid(color)
                        } else {
                            Paint::None
                        },
                        Paint::Solid(color),
                        1.0,
                        &chart.source_part,
                        shape_id,
                        Some(u32::try_from(legend_index).unwrap_or(u32::MAX)),
                        None,
                        MappingQuality::Derived,
                    )?;
                }
            }
            push_pptx_chart_text(
                state,
                chart_id,
                unit_index,
                Rect {
                    x: item_x + key_width + 4.0,
                    y: item_y - 3.0,
                    width: if horizontal_legend {
                        item_width - key_width - 8.0
                    } else {
                        legend_width - key_width - 10.0
                    },
                    height: legend_line_height,
                },
                if chart.font_size.is_some() {
                    label.replace(' ', "\u{00a0}")
                } else {
                    label.to_owned()
                },
                legend_style.color,
                &chart,
                shape_id,
                Some(u32::try_from(legend_index).unwrap_or(u32::MAX)),
                legend_font_size,
                legend_style.bold,
                0.0,
                false,
                legend_style.align,
            )?;
            horizontal_offset += item_width;
        }
    }
    Ok(())
}

fn chart_category_labels_overlap(
    categories: &[String],
    category_width: f32,
    font_size: f32,
) -> bool {
    let available_width = category_width.max(1.0) * 0.9;
    categories.iter().any(|label| {
        label
            .chars()
            .map(|character| drawingml_fallback_character_width(character, font_size))
            .sum::<f32>()
            > available_width
    })
}

#[allow(clippy::too_many_arguments)]
fn push_pptx_chart_data_table(
    state: &mut PptxParseState,
    chart_id: u32,
    unit_index: u32,
    chart_bounds: Rect,
    plot: Rect,
    chart: &BasicChart,
    shape_id: u32,
) -> Result<(), Diagnostic> {
    let Some(layout) = chart_data_table_layout(chart, chart_bounds, plot) else {
        return Ok(());
    };
    for line in layout.lines {
        push_pptx_chart_shape(
            state,
            chart_id,
            unit_index,
            line.bounds,
            line.geometry.clone(),
            if matches!(line.geometry, Geometry::Rectangle) {
                Paint::Solid(line.color)
            } else {
                Paint::None
            },
            Paint::Solid(line.color),
            line.width,
            &chart.source_part,
            shape_id,
            line.series_index
                .map(|index| u32::try_from(index).unwrap_or(u32::MAX)),
            None,
            MappingQuality::Derived,
        )?;
    }
    for text in layout.texts {
        push_pptx_chart_text(
            state,
            chart_id,
            unit_index,
            text.bounds,
            text.text,
            0x0000_00ff,
            chart,
            shape_id,
            text.series_index
                .map(|index| u32::try_from(index).unwrap_or(u32::MAX)),
            10.0 * POINTS_TO_CSS_PIXELS,
            false,
            0.0,
            false,
            text.align,
        )?;
    }
    Ok(())
}

fn chart_value_y(value: f32, minimum: f32, maximum: f32, plot: Rect) -> f32 {
    let ratio = ((value - minimum) / (maximum - minimum)).clamp(0.0, 1.0);
    plot.y + plot.height - plot.height * ratio
}

#[allow(clippy::too_many_arguments)]
fn push_pptx_chart_shape(
    state: &mut PptxParseState,
    parent_numeric_id: u32,
    unit_index: u32,
    bounds: Rect,
    geometry: Geometry,
    fill: Paint,
    stroke: Paint,
    stroke_width: f32,
    source_part: &str,
    shape_id: u32,
    row: Option<u32>,
    column: Option<u32>,
    mapping: MappingQuality,
) -> Result<(), Diagnostic> {
    let numeric_id = u32::try_from(state.objects.len())
        .map_err(|_| format_error(source_part, "object count exceeds supported range"))?;
    let z = state.take_z();
    state.objects.push(Object {
        numeric_id,
        parent_numeric_id: Some(parent_numeric_id),
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: Some(format!("object:{parent_numeric_id}")),
        kind: ObjectKind::Shape,
        unit_index,
        bounds,
        z,
        text: None,
        source: SourceRef {
            part: source_part.to_owned(),
            mapping,
            locator: SourceLocator::PptxShape {
                shape_id,
                row,
                column,
                text_range: None,
                metadata: PptxObjectMetadata::default(),
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
fn push_pptx_chart_text(
    state: &mut PptxParseState,
    parent_numeric_id: u32,
    unit_index: u32,
    bounds: Rect,
    text: String,
    color: u32,
    chart: &BasicChart,
    shape_id: u32,
    row: Option<u32>,
    font_size: f32,
    bold: bool,
    rotation_degrees: f32,
    use_chart_font_size: bool,
    align: TextAlign,
) -> Result<(), Diagnostic> {
    let (font_size, bold) = chart_text_style(
        font_size,
        bold,
        chart.font_size,
        chart.font_bold,
        use_chart_font_size,
        chart.native_size.is_some(),
    );
    let source_part = chart.source_part.as_str();
    let numeric_id = u32::try_from(state.objects.len())
        .map_err(|_| format_error(source_part, "object count exceeds supported range"))?;
    let text_length = u32::try_from(text.chars().count()).unwrap_or(u32::MAX);
    let z = state.take_z();
    state.objects.push(Object {
        numeric_id,
        parent_numeric_id: Some(parent_numeric_id),
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: Some(format!("object:{parent_numeric_id}")),
        kind: ObjectKind::TextBox,
        unit_index,
        bounds,
        z,
        text: Some(text.clone()),
        source: SourceRef {
            part: source_part.to_owned(),
            mapping: MappingQuality::Derived,
            locator: SourceLocator::PptxShape {
                shape_id,
                row,
                column: None,
                text_range: Some((0, text_length)),
                metadata: PptxObjectMetadata::default(),
            },
        },
        visual: Visual::TextLayout {
            layout: TextLayout {
                vertical_align: TextVerticalAlign::Center,
                rotation_degrees,
                inset_left: 0.0,
                inset_right: 0.0,
                inset_top: 0.0,
                inset_bottom: 0.0,
                wrap: rotation_degrees.abs() < f32::EPSILON && !text.contains('\u{00a0}'),
                ..TextLayout::default()
            },
            visual: Box::new(Visual::RichText {
                geometry: Geometry::Rectangle,
                fill: Paint::None,
                stroke: Paint::None,
                stroke_width: 0.0,
                align,
                line_height: 0.0,
                runs: vec![TextRun {
                    paint: None,
                    east_asian_line_breaks: true,
                    text,
                    font_family: chart
                        .font_family
                        .clone()
                        .unwrap_or_else(|| "Arial".to_owned()),
                    font_size,
                    color,
                    bold,
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
    Ok(())
}

fn chart_text_style(
    font_size: f32,
    bold: bool,
    chart_font_size: Option<f32>,
    chart_bold: bool,
    use_chart_font_size: bool,
    legacy_msgraph: bool,
) -> (f32, bool) {
    if use_chart_font_size && !legacy_msgraph {
        (chart_font_size.unwrap_or(font_size), chart_bold || bold)
    } else {
        (font_size, bold)
    }
}

fn normalized_table_sizes(weights: &[u64], count: usize, available: f32) -> Vec<f32> {
    if count == 0 {
        return Vec::new();
    }
    let total = weights.iter().take(count).copied().sum::<u64>();
    if total == 0 {
        return vec![available / count as f32; count];
    }
    (0..count)
        .map(|index| weights.get(index).copied().unwrap_or(0) as f32 / total as f32 * available)
        .collect()
}

fn table_row_sizes(
    rows: &[TableRowState],
    column_sizes: &[f32],
    available: f32,
    font_metrics: &FontMetricTable,
) -> Vec<f32> {
    if rows.is_empty() {
        return Vec::new();
    }
    let authored = rows
        .iter()
        .map(|row| row.height as f32 / EMU_PER_CSS_PIXEL)
        .collect::<Vec<_>>();
    let mut sizes = authored.clone();
    let mut spanning_constraints = Vec::new();
    for (row_index, row) in rows.iter().enumerate() {
        let mut logical_column = 0_usize;
        for cell in &row.cells {
            // gridSpan on the merge origin already accounts for hMerge
            // continuation cells present in the XML row.
            if cell.horizontal_merge {
                continue;
            }
            let column_index = logical_column;
            logical_column = logical_column.saturating_add(cell.grid_span.max(1));
            if cell.vertical_merge {
                continue;
            }
            let width = column_sizes
                .iter()
                .skip(column_index)
                .take(cell.grid_span)
                .sum::<f32>();
            let content = estimated_table_cell_height(cell, width, font_metrics);
            let span = cell.row_span.min(rows.len() - row_index).max(1);
            if span == 1 {
                sizes[row_index] = sizes[row_index].max(content);
            } else {
                spanning_constraints.push((span, row_index, column_index, content));
            }
        }
    }
    spanning_constraints.sort_by_key(|&(span, row, column, _)| (span, row, column));
    for (span, row_index, _, required) in spanning_constraints {
        let end = row_index + span;
        let current = sizes[row_index..end].iter().sum::<f32>();
        let deficit = required - current;
        if deficit <= 0.0 {
            continue;
        }
        let automatic = (row_index..end)
            .filter(|&index| rows[index].height == 0)
            .collect::<Vec<_>>();
        let targets = if automatic.is_empty() {
            (row_index..end).collect::<Vec<_>>()
        } else {
            automatic
        };
        let extra = deficit / targets.len() as f32;
        for index in targets {
            sizes[index] += extra;
        }
    }

    let preferred_total = sizes.iter().sum::<f32>();
    // Row heights are minimums; the frame must grow when wrapped text needs more space.
    if preferred_total > available {
        return sizes;
    }

    let remaining = available - preferred_total;
    if remaining > 0.0 {
        let automatic = rows
            .iter()
            .enumerate()
            .filter_map(|(index, row)| (row.height == 0).then_some(index))
            .collect::<Vec<_>>();
        let targets = if automatic.is_empty() {
            (0..rows.len()).collect::<Vec<_>>()
        } else {
            automatic
        };
        let extra = remaining / targets.len() as f32;
        for index in targets {
            sizes[index] += extra;
        }
    }
    let residual = available - sizes.iter().sum::<f32>();
    if let Some(last) = sizes.last_mut() {
        *last += residual;
    }
    sizes
}

fn drawingml_cjk_character(character: char) -> bool {
    matches!(
        character as u32,
        0x1100..=0x11ff
            | 0x2e80..=0x2fff
            | 0x3000..=0x303f
            | 0x3040..=0x30ff
            | 0x3100..=0x312f
            | 0x3130..=0x318f
            | 0x31a0..=0x31bf
            | 0x31f0..=0x31ff
            | 0x3400..=0x4dbf
            | 0x4e00..=0x9fff
            | 0xa960..=0xa97f
            | 0xac00..=0xd7af
            | 0xd7b0..=0xd7ff
            | 0xf900..=0xfaff
            | 0x20000..=0x2fa1f
    )
}

fn drawingml_text_tokens(text: &str) -> Vec<&str> {
    let mut tokens = Vec::new();
    let mut offset = 0;
    while offset < text.len() {
        let remaining = &text[offset..];
        if remaining.starts_with("\r\n") {
            tokens.push(&text[offset..offset + 2]);
            offset += 2;
            continue;
        }
        let Some(character) = remaining.chars().next() else {
            break;
        };
        let character_bytes = character.len_utf8();
        if matches!(character, '\n' | '\t' | '\u{2028}') {
            tokens.push(&text[offset..offset + character_bytes]);
            offset += character_bytes;
            continue;
        }
        let whitespace = character.is_whitespace();
        let start = offset;
        offset += character_bytes;
        while offset < text.len() {
            let tail = &text[offset..];
            if tail.starts_with("\r\n") {
                break;
            }
            let Some(next) = tail.chars().next() else {
                break;
            };
            if matches!(next, '\n' | '\t' | '\u{2028}') || next.is_whitespace() != whitespace {
                break;
            }
            offset += next.len_utf8();
        }
        let token = &text[start..offset];
        if !whitespace && token.chars().any(drawingml_cjk_character) {
            let mut character_start = start;
            for character in token.chars() {
                let character_end = character_start + character.len_utf8();
                tokens.push(&text[character_start..character_end]);
                character_start = character_end;
            }
        } else {
            tokens.push(token);
        }
    }
    tokens
}

fn drawingml_latin_line_break_tokens(token: &str) -> Vec<&str> {
    let mut tokens = Vec::new();
    let mut start = 0_usize;
    for (index, character) in token.char_indices() {
        let end = index + character.len_utf8();
        if matches!(character, '-' | '\u{2010}') && end < token.len() {
            tokens.push(&token[start..end]);
            start = end;
        }
    }
    if start < token.len() {
        tokens.push(&token[start..]);
    }
    if tokens.is_empty() {
        tokens.push(token);
    }
    tokens
}

fn drawingml_hanging_punctuation(character: char) -> bool {
    matches!(
        character,
        '\u{3001}'
            | '\u{3002}'
            | '\u{3003}'
            | '\u{3009}'
            | '\u{300b}'
            | '\u{300d}'
            | '\u{300f}'
            | '\u{3011}'
            | '\u{3015}'
            | '\u{3017}'
            | '\u{3019}'
            | '\u{301b}'
            | '\u{301e}'
            | '\u{301f}'
            | '\u{303d}'
            | '\u{fe10}'
            | '\u{fe11}'
            | '\u{fe12}'
            | '\u{fe13}'
            | '\u{fe14}'
            | '\u{fe15}'
            | '\u{fe16}'
            | '\u{fe18}'
            | '\u{fe19}'
            | '\u{fe30}'
            | '\u{fe36}'
            | '\u{fe38}'
            | '\u{fe3a}'
            | '\u{fe3c}'
            | '\u{fe3e}'
            | '\u{fe40}'
            | '\u{fe42}'
            | '\u{fe44}'
            | '\u{fe45}'
            | '\u{fe46}'
            | '\u{fe48}'
            | '\u{fe49}'
            | '\u{fe4a}'
            | '\u{fe4b}'
            | '\u{fe4c}'
            | '\u{ff01}'
            | '\u{ff02}'
            | '\u{ff03}'
            | '\u{ff05}'
            | '\u{ff06}'
            | '\u{ff07}'
            | '\u{ff09}'
            | '\u{ff0a}'
            | '\u{ff0c}'
            | '\u{ff0e}'
            | '\u{ff0f}'
            | '\u{ff1a}'
            | '\u{ff1b}'
            | '\u{ff1f}'
            | '\u{ff20}'
            | '\u{ff3c}'
            | '\u{ff3d}'
            | '\u{ff5d}'
            | '\u{ff60}'
            | '\u{ff61}'
            | '\u{ff63}'
            | '\u{ff64}'
            | '\u{ff65}'
    )
}

fn estimated_drawingml_wrapping_width<F>(token: &str, width: f32, hanging: bool, width_of: F) -> f32
where
    F: Fn(char) -> f32,
{
    if !hanging {
        return width;
    }
    let punctuation_width = token
        .chars()
        .rev()
        .take_while(|character| drawingml_hanging_punctuation(*character))
        .map(width_of)
        .sum::<f32>();
    (width - punctuation_width).max(0.0)
}

fn fallback_drawingml_character_width(character: char, font_size: f32) -> f32 {
    drawingml_fallback_character_width(character, font_size)
}

fn estimated_drawingml_character_width(
    character: char,
    font_family: &str,
    font_size: f32,
    bold: bool,
    italic: bool,
    letter_spacing: f32,
    font_metrics: &FontMetricTable,
) -> f32 {
    let advance = font_metrics
        .advance_em_at_size(font_family, italic, bold, character, font_size)
        .map_or_else(
            || fallback_drawingml_character_width(character, font_size),
            |advance_em| advance_em * font_size,
        );
    (advance + letter_spacing).max(0.0)
}

#[allow(clippy::too_many_arguments)]
fn layout_estimated_drawingml_run(
    text: &str,
    font_family: &str,
    font_size: f32,
    bold: bool,
    italic: bool,
    letter_spacing: f32,
    usable_width: f32,
    wrap: bool,
    font_metrics: &FontMetricTable,
    lines: &mut usize,
    line_width: &mut f32,
    maximum_line_width: &mut f32,
) {
    let character_width = |character: char| {
        estimated_drawingml_character_width(
            character,
            font_family,
            font_size,
            bold,
            italic,
            letter_spacing,
            font_metrics,
        )
    };
    let new_line = |lines: &mut usize, line_width: &mut f32, maximum_line_width: &mut f32| {
        *maximum_line_width = maximum_line_width.max(*line_width);
        *lines = lines.saturating_add(1);
        *line_width = 0.0;
    };

    for token in drawingml_text_tokens(text) {
        if matches!(token, "\n" | "\r\n" | "\u{2028}") {
            new_line(lines, line_width, maximum_line_width);
            continue;
        }
        if token == "\t" {
            *line_width = ((*line_width / 36.0).floor() + 1.0) * 36.0;
            *maximum_line_width = maximum_line_width.max(*line_width);
            continue;
        }
        let token_width = token.chars().map(character_width).sum::<f32>();
        let oversized = wrap && token_width > usable_width && !token.trim().is_empty();
        if oversized {
            for character in token.chars() {
                let width = character_width(character);
                if *line_width > 0.0 && *line_width + width > usable_width {
                    new_line(lines, line_width, maximum_line_width);
                }
                *line_width += width;
                *maximum_line_width = maximum_line_width.max(*line_width);
            }
            continue;
        }
        if wrap && *line_width > 0.0 && *line_width + token_width > usable_width {
            new_line(lines, line_width, maximum_line_width);
        }
        *line_width += token_width;
        *maximum_line_width = maximum_line_width.max(*line_width);
    }
}

#[allow(clippy::too_many_arguments)]
fn estimated_drawingml_text_height(
    runs: &[TextRun],
    paragraphs: &[TextParagraphLayout],
    default_font_size: f32,
    usable_width: f32,
    wrap: bool,
    font_scale: f32,
    line_spacing_reduction: f32,
    font_metrics: &FontMetricTable,
) -> f32 {
    if runs.is_empty() {
        return 0.0;
    }
    let font_scale = font_scale.clamp(0.01, 1.0);
    let line_scale = (1.0 - line_spacing_reduction.clamp(0.0, 0.99)) * font_scale;
    let default_font_size = runs
        .iter()
        .map(|run| run.font_size)
        .reduce(f32::max)
        .unwrap_or(default_font_size)
        .max(1.0);
    let paragraph_metrics = |index: usize| {
        let layout = paragraphs.get(index);
        (
            layout.map_or(0.0, |layout| layout.margin_left.max(0.0)),
            layout.map_or(0.0, |layout| layout.margin_right.max(0.0)),
            layout.map_or(0.0, |layout| layout.first_line_indent),
            layout.map_or(36.0, |layout| layout.default_tab_stop.max(1.0)),
            layout.map_or(default_font_size * 1.2 * line_scale, |layout| {
                if layout.line_height > 0.0 {
                    layout.line_height * line_scale
                } else {
                    default_font_size * 1.2 * line_scale
                }
            }),
            layout.map_or(0.0, |layout| layout.space_before.max(0.0) * font_scale),
            layout.map_or(0.0, |layout| layout.space_after.max(0.0) * font_scale),
            layout.is_some_and(|layout| layout.latin_line_break),
            layout.is_none_or(|layout| layout.hanging_punctuation),
        )
    };
    let usable_width = usable_width.max(1.0);
    let mut paragraph_index = 0_usize;
    let mut soft_break_pending = false;
    let (
        mut margin_left,
        mut margin_right,
        mut first_line_indent,
        mut default_tab_stop,
        mut line_height,
        mut space_before,
        mut space_after,
        mut latin_line_break,
        mut hanging_punctuation,
    ) = paragraph_metrics(paragraph_index);
    let mut line_count = 1_usize;
    let mut line_width = (margin_left + first_line_indent).max(0.0);
    let mut line_start = line_width;
    let mut content_height = 0.0;

    for run in runs {
        let character_width = |character: char| {
            estimated_drawingml_character_width(
                character,
                &run.font_family,
                run.font_size * font_scale,
                run.bold,
                run.italic,
                run.letter_spacing * font_scale,
                font_metrics,
            )
        };
        for raw_token in drawingml_text_tokens(&run.text) {
            if raw_token == "\u{2028}" {
                soft_break_pending = false;
                line_count += 1;
                line_width = margin_left;
                line_start = line_width;
                continue;
            }
            if matches!(raw_token, "\n" | "\r\n") {
                soft_break_pending = false;
                content_height += space_before + line_count as f32 * line_height + space_after;
                paragraph_index = paragraph_index.saturating_add(1);
                (
                    margin_left,
                    margin_right,
                    first_line_indent,
                    default_tab_stop,
                    line_height,
                    space_before,
                    space_after,
                    latin_line_break,
                    hanging_punctuation,
                ) = paragraph_metrics(paragraph_index);
                line_count = 1;
                line_width = (margin_left + first_line_indent).max(0.0);
                line_start = line_width;
                continue;
            }
            let tokens = if latin_line_break {
                drawingml_latin_line_break_tokens(raw_token)
            } else {
                vec![raw_token]
            };
            for token in tokens {
                let soft_wrap_space =
                    !token.is_empty() && token.chars().all(|character| character == ' ');
                if wrap && soft_break_pending && !soft_wrap_space && token != "\t" {
                    line_count = line_count.saturating_add(1);
                    line_width = margin_left;
                    line_start = line_width;
                    soft_break_pending = false;
                }
                if token == "\t" {
                    line_width = if line_width < margin_left {
                        margin_left
                    } else {
                        ((line_width / default_tab_stop).floor() + 1.0) * default_tab_stop
                    };
                    continue;
                }
                let token_width = token.chars().map(character_width).sum::<f32>();
                let token_wrapping_width = estimated_drawingml_wrapping_width(
                    token,
                    token_width,
                    hanging_punctuation,
                    character_width,
                );
                let maximum_x = (usable_width - margin_right).max(margin_left + 1.0);
                let paragraph_width = (maximum_x - margin_left).max(1.0);
                let oversized =
                    wrap && token_wrapping_width > paragraph_width && !token.trim().is_empty();
                if wrap
                    && soft_wrap_space
                    && line_width > margin_left
                    && line_width + token_width > maximum_x
                {
                    soft_break_pending = true;
                    continue;
                }
                if wrap
                    && line_width > line_start
                    && line_width + token_wrapping_width > maximum_x
                    && !oversized
                {
                    line_count = line_count.saturating_add(1);
                    line_width = margin_left;
                    line_start = line_width;
                }
                if oversized {
                    for character in token.chars() {
                        let width = character_width(character);
                        let wrapping_width =
                            if hanging_punctuation && drawingml_hanging_punctuation(character) {
                                0.0
                            } else {
                                width
                            };
                        if line_width > line_start && line_width + wrapping_width > maximum_x {
                            line_count = line_count.saturating_add(1);
                            line_width = margin_left;
                            line_start = line_width;
                        }
                        line_width += width;
                    }
                } else {
                    line_width += token_width;
                }
            }
        }
    }
    content_height + space_before + line_count as f32 * line_height + space_after
}

fn estimated_table_cell_height(
    cell: &TableCellState,
    width: f32,
    font_metrics: &FontMetricTable,
) -> f32 {
    if cell.text.is_empty() {
        return 0.0;
    }
    let fallback_runs = cell.runs.is_empty().then(|| {
        vec![TextRun {
            paint: None,
            east_asian_line_breaks: true,
            text: cell.text.clone(),
            font_family: cell.font_family.clone(),
            font_size: cell.font_size,
            color: cell.font_color,
            bold: cell.bold,
            italic: cell.italic,
            underline: cell.underline,
            strikethrough: cell.strikethrough,
            highlight: 0,
            baseline_shift: cell.baseline_shift,
            letter_spacing: cell.letter_spacing,
            horizontal_scale: 1.0,
        }]
    });
    let runs = fallback_runs.as_deref().unwrap_or(&cell.runs);
    let font_size = runs
        .iter()
        .map(|run| run.font_size)
        .reduce(f32::max)
        .unwrap_or(cell.font_size)
        .max(1.0);
    let rotated = matches!(
        cell.text_orientation,
        TextOrientation::Rotated90 | TextOrientation::Rotated270
    );
    let usable_width = if rotated {
        f32::INFINITY
    } else {
        (width - cell.text_inset_left - cell.text_inset_right).max(1.0)
    };
    let mut lines = 1_usize;
    let mut line_width = 0.0;
    let mut maximum_line_width = 0.0;
    if rotated {
        for run in runs {
            layout_estimated_drawingml_run(
                &run.text,
                &run.font_family,
                run.font_size,
                run.bold,
                run.italic,
                run.letter_spacing,
                usable_width,
                false,
                font_metrics,
                &mut lines,
                &mut line_width,
                &mut maximum_line_width,
            );
        }
        return maximum_line_width + cell.text_inset_left + cell.text_inset_right;
    }

    estimated_drawingml_text_height(
        runs,
        &cell.paragraph_layouts,
        font_size,
        usable_width,
        true,
        1.0,
        0.0,
        font_metrics,
    ) + cell.text_inset_top
        + cell.text_inset_bottom
}

fn related_presentation_part(
    relationships: &[Relationship],
    type_suffix: &str,
    description: &str,
    source_part: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<Option<String>, Diagnostic> {
    let mut matches = relationships
        .iter()
        .filter(|relationship| relationship.type_uri.ends_with(type_suffix));
    let Some(relationship) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(format_error(
            source_part,
            format!("{source_part} has multiple {description} relationships"),
        ));
    }
    if relationship.external {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ExternalResourceBlocked,
                Phase::Security,
                Fidelity::Blocked,
                format!("external {description} relationship was blocked"),
            )
            .in_part(source_part),
        );
        return Ok(None);
    }
    Ok(Some(relationship.target.clone()))
}

fn presentation_fallback_theme_part(
    relationships: &[Relationship],
    source_part: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Option<String> {
    let mut count = 0_usize;
    let mut first_internal = None;
    let mut blocked_external = false;
    for relationship in relationships
        .iter()
        .filter(|relationship| relationship.type_uri.ends_with("/theme"))
    {
        count += 1;
        if relationship.external {
            blocked_external = true;
        } else if first_internal.is_none() {
            first_internal = Some(relationship.target.clone());
        }
    }
    if count > 1 {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::FormatInvalid,
                Phase::Parse,
                Fidelity::Approximate,
                format!(
                    "{source_part} has multiple presentation theme relationships; the first internal theme is used as the presentation fallback"
                ),
            )
            .in_part(source_part),
        );
    }
    if blocked_external {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ExternalResourceBlocked,
                Phase::Security,
                Fidelity::Blocked,
                "external presentation theme relationship was blocked",
            )
            .in_part(source_part),
        );
    }
    first_internal
}

fn resolve_shape_image_fill(
    fill: ShapeImageFillState,
    part: &str,
    relationships: &HashMap<&str, &Relationship>,
    package: &Package<'_>,
    state: &mut PptxParseState,
    content_types: &ContentTypes,
) -> Result<Paint, Diagnostic> {
    let relationship_id = fill
        .preferred_svg_relationship_id
        .as_deref()
        .or(fill.embedded_relationship_id.as_deref());
    let Some(relationship_id) = relationship_id else {
        if fill.linked_relationship_id.is_some() {
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::ExternalResourceBlocked,
                    Phase::Security,
                    Fidelity::Blocked,
                    "linked PPTX shape image fill was blocked",
                )
                .in_part(part)
                .with_detail("feature", "shape-image-fill"),
            );
        }
        return Ok(Paint::None);
    };
    let Some(relationship) = relationships.get(relationship_id) else {
        return Err(format_error(
            part,
            format!("shape image-fill relationship {relationship_id} does not exist"),
        ));
    };
    let mut paint = resolve_image_fill_relationship(
        relationship,
        part,
        fill.crop,
        fill.tile,
        package,
        state,
        content_types,
    )?;
    if let Paint::Image { mapping, .. } = &mut paint {
        *mapping = (fill.tile || fill.mapping != crate::model::ImageFillMapping::default())
            .then(|| Box::new(fill.mapping));
    }
    Ok(paint)
}

fn resolve_image_fill_relationship(
    relationship: &Relationship,
    part: &str,
    crop: ImageCrop,
    tile: bool,
    package: &Package<'_>,
    state: &mut PptxParseState,
    content_types: &ContentTypes,
) -> Result<Paint, Diagnostic> {
    if relationship.external {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ExternalResourceBlocked,
                Phase::Security,
                Fidelity::Blocked,
                "external PPTX shape image fill was blocked",
            )
            .in_part(part)
            .with_detail("feature", "shape-image-fill"),
        );
        return Ok(Paint::None);
    }
    if !relationship.type_uri.ends_with("/image") {
        return Err(format_error(
            part,
            format!(
                "PPTX image-fill relationship {} is not an image",
                relationship.id
            ),
        ));
    }
    let bytes = package.required_part(&relationship.target)?;
    let declared_content_type = content_types.for_part(&relationship.target);
    let media_type = pptx_image_media_type(declared_content_type, &relationship.target, &bytes);
    let media_type = match media_type {
        Ok(media_type) => media_type,
        Err(error) => {
            state.diagnostics.push(unsupported_image_diagnostic(
                &relationship.target,
                error,
                declared_content_type.is_some(),
            ));
            return Ok(Paint::None);
        }
    };
    reserve_materialized_image_bytes(
        &mut state.materialized_image_bytes,
        bytes.len(),
        package.limits().max_total_uncompressed_bytes,
        &relationship.target,
    )?;
    Ok(Paint::Image {
        mapping: None,
        media_type: media_type.to_owned(),
        bytes: bytes.into_vec(),
        crop,
        tile,
        tile_width: None,
        tile_height: None,
    })
}

fn materialize_picture_media(
    picture: &PictureState,
    part: &str,
    relationships: &HashMap<&str, &Relationship>,
    package: &Package<'_>,
    state: &mut PptxParseState,
    content_types: &ContentTypes,
) -> Result<Option<(MediaKind, String, Vec<u8>)>, Diagnostic> {
    let Some(relationship_id) = picture.media_relationship_id.as_deref() else {
        return Ok(None);
    };
    let Some(relationship) = relationships.get(relationship_id) else {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Omitted,
                format!("embedded media relationship {relationship_id} does not exist"),
            )
            .in_part(part),
        );
        return Ok(None);
    };
    if relationship.external {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ExternalResourceBlocked,
                Phase::Security,
                Fidelity::Blocked,
                "external audio/video relationship was blocked",
            )
            .in_part(part),
        );
        return Ok(None);
    }
    if !matches!(
        relationship.type_uri.rsplit('/').next(),
        Some("audio" | "video" | "media")
    ) {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Omitted,
                "embedded media relationship has an unsupported relationship type",
            )
            .in_part(part),
        );
        return Ok(None);
    }
    let Some(bytes) = package.part(&relationship.target)? else {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Container,
                Fidelity::Omitted,
                "embedded audio/video package part is missing",
            )
            .in_part(&relationship.target),
        );
        return Ok(None);
    };
    let relationship_hint = if relationship.type_uri.ends_with("/audio") {
        Some(MediaKind::Audio)
    } else if relationship.type_uri.ends_with("/video") {
        Some(MediaKind::Video)
    } else {
        picture.media_kind_hint
    };
    let identified = content_types.for_part(&relationship.target).map_or_else(
        || embedded_media_type(&relationship.target, &bytes, relationship_hint),
        |media_type| {
            embedded_media_type_from_mime(media_type, &bytes).or_else(|error| {
                if error == EmbeddedMediaError::UnsupportedFormat {
                    embedded_media_type(&relationship.target, &bytes, relationship_hint)
                } else {
                    Err(error)
                }
            })
        },
    );
    let (kind, media_type) = match identified {
        Ok(identified) => identified,
        Err(error) => {
            let message = match error {
                EmbeddedMediaError::UnsupportedFormat => {
                    "embedded audio/video format is outside the browser-playable allowlist"
                }
                EmbeddedMediaError::SignatureMismatch => {
                    "embedded audio/video bytes do not match the declared media format"
                }
            };
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Parse,
                    Fidelity::Omitted,
                    message,
                )
                .in_part(&relationship.target),
            );
            return Ok(None);
        }
    };
    reserve_materialized_image_bytes(
        &mut state.materialized_image_bytes,
        bytes.len(),
        package.limits().max_total_uncompressed_bytes,
        &relationship.target,
    )?;
    state.diagnostics.retain(|diagnostic| {
        let feature = diagnostic
            .details
            .iter()
            .find(|(key, _)| key == "feature")
            .map(|(_, value)| value.as_str());
        let target = diagnostic
            .details
            .iter()
            .find(|(key, _)| key == "target")
            .map(|(_, value)| value.as_str());
        feature != Some("media-playback") || target != Some(relationship.target.as_str())
    });
    Ok(Some((kind, media_type.to_owned(), bytes.into_vec())))
}

fn push_picture(
    picture: PictureState,
    part: &str,
    unit_index: u32,
    relationships: &HashMap<&str, &Relationship>,
    package: &Package<'_>,
    state: &mut PptxParseState,
    content_types: &ContentTypes,
) -> Result<(), Diagnostic> {
    let replacement_id = find_placeholder_object(
        state,
        unit_index,
        picture.placeholder_index,
        picture.placeholder_type.as_deref(),
    );
    let mut picture = picture;
    if picture.geometry.preset.is_none()
        && let Some((preset, adjustments)) =
            replacement_id.and_then(|id| state.placeholder_presets.get(&id))
    {
        picture.geometry.preset = Some(preset.clone());
        picture.geometry.adjustments.clone_from(adjustments);
    }
    let embedded_media =
        materialize_picture_media(&picture, part, relationships, package, state, content_types)?;
    if replacement_id.is_none() && state.objects.len() >= package.limits().max_document_objects {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "document exceeds the configured object limit",
        )
        .in_part(part));
    }

    let fallback_relationship = picture
        .embedded_relationship_id
        .as_deref()
        .map(|id| (id, false))
        .or_else(|| {
            picture
                .linked_relationship_id
                .as_deref()
                .map(|id| (id, true))
        });
    if picture.preferred_svg_relationship_id.is_none() && fallback_relationship.is_none() {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Omitted,
                "picture has no embedded image relationship and was omitted",
            )
            .in_part(part),
        );
        return Ok(());
    }

    let inherited_object = replacement_id
        .and_then(|numeric_id| usize::try_from(numeric_id).ok())
        .and_then(|index| state.objects.get(index));
    let inherited_bounds = inherited_object.map(|object| object.bounds);
    let inherited_transform = inherited_object
        .map(|object| leading_transform(&object.visual))
        .unwrap_or(AffineTransform::IDENTITY);
    let bounds = if let Some(bounds) = picture.explicit_bounds {
        bounds
    } else {
        Rect {
            x: picture
                .x
                .map(|value| value as f32 / EMU_PER_CSS_PIXEL)
                .or(inherited_bounds.map(|bounds| bounds.x))
                .ok_or_else(|| format_error(part, "picture transform is missing x"))?,
            y: picture
                .y
                .map(|value| value as f32 / EMU_PER_CSS_PIXEL)
                .or(inherited_bounds.map(|bounds| bounds.y))
                .ok_or_else(|| format_error(part, "picture transform is missing y"))?,
            width: picture
                .width
                .map(|value| value as f32 / EMU_PER_CSS_PIXEL)
                .or(inherited_bounds.map(|bounds| bounds.width))
                .ok_or_else(|| format_error(part, "picture transform is missing cx"))?,
            height: picture
                .height
                .map(|value| value as f32 / EMU_PER_CSS_PIXEL)
                .or(inherited_bounds.map(|bounds| bounds.height))
                .ok_or_else(|| format_error(part, "picture transform is missing cy"))?,
        }
    };
    if !bounds.is_valid() {
        return Err(format_error(part, "picture has invalid bounds"));
    }
    let numeric_id = if let Some(replacement_id) = replacement_id {
        replacement_id
    } else {
        u32::try_from(state.objects.len())
            .map_err(|_| format_error(part, "object count exceeds supported range"))?
    };
    let object_index = usize::try_from(numeric_id)
        .map_err(|_| format_error(part, "picture object ID exceeds addressable range"))?;
    let has_fallback = fallback_relationship.is_some();
    let candidates = [
        picture
            .preferred_svg_relationship_id
            .as_deref()
            .map(|id| (id, false, true)),
        fallback_relationship.map(|(id, is_linked)| (id, is_linked, false)),
    ];
    let mut selected_image = None;
    let mut raster_fallback = None;
    for (relationship_id, is_linked, is_preferred_svg) in candidates.into_iter().flatten() {
        if !is_preferred_svg
            && picture.preferred_svg_relationship_id.as_deref() == Some(relationship_id)
        {
            continue;
        }
        let Some(relationship) = relationships.get(relationship_id) else {
            if is_preferred_svg {
                state.diagnostics.push(preferred_svg_fallback_diagnostic(
                    part,
                    format!("relationship {relationship_id} does not exist"),
                    has_fallback,
                ));
                continue;
            }
            if selected_image.is_some() {
                state
                    .diagnostics
                    .push(runtime_fallback_unavailable_diagnostic(
                        part,
                        format!("relationship {relationship_id} does not exist"),
                    ));
                break;
            }
            return Err(format_error(
                part,
                format!("picture relationship {relationship_id} does not exist"),
            ));
        };
        if relationship.external {
            let message = if is_preferred_svg {
                if has_fallback {
                    "external preferred SVG relationship was blocked; raster fallback will be attempted"
                } else {
                    "external preferred SVG relationship was blocked; no raster fallback is available"
                }
            } else if selected_image.is_some() {
                "external raster fallback relationship was blocked; preferred SVG will be used without runtime fallback"
            } else {
                "external picture relationship was blocked"
            };
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::ExternalResourceBlocked,
                    Phase::Security,
                    Fidelity::Blocked,
                    message,
                )
                .in_part(part),
            );
            if is_preferred_svg {
                continue;
            }
            if selected_image.is_some() {
                break;
            }
            return Ok(());
        }
        if is_linked {
            if selected_image.is_some() {
                state
                    .diagnostics
                    .push(runtime_fallback_unavailable_diagnostic(
                        part,
                        "linked raster fallback is not an embedded package resource",
                    ));
                break;
            }
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Parse,
                    Fidelity::Omitted,
                    "linked picture is not an embedded package resource and was omitted",
                )
                .in_part(part),
            );
            return Ok(());
        }
        if !relationship.type_uri.ends_with("/image") {
            if is_preferred_svg {
                state.diagnostics.push(preferred_svg_fallback_diagnostic(
                    part,
                    format!("relationship {relationship_id} is not an image"),
                    has_fallback,
                ));
                continue;
            }
            if selected_image.is_some() {
                state
                    .diagnostics
                    .push(runtime_fallback_unavailable_diagnostic(
                        part,
                        format!("relationship {relationship_id} is not an image"),
                    ));
                break;
            }
            return Err(format_error(
                part,
                format!("relationship {relationship_id} is not an image"),
            ));
        }
        let declared_content_type = content_types.for_part(&relationship.target);
        if is_preferred_svg
            && declared_content_type.is_none()
            && !relationship
                .target
                .rsplit_once('.')
                .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("svg"))
        {
            state.diagnostics.push(preferred_svg_fallback_diagnostic(
                part,
                format!("relationship {relationship_id} does not target an SVG package part"),
                has_fallback,
            ));
            continue;
        }
        if !package.has_part(&relationship.target) {
            if is_preferred_svg {
                state.diagnostics.push(preferred_svg_fallback_diagnostic(
                    part,
                    format!("package part {} is missing", relationship.target),
                    has_fallback,
                ));
                continue;
            }
            if selected_image.is_some() {
                state
                    .diagnostics
                    .push(runtime_fallback_unavailable_diagnostic(
                        part,
                        format!("package part {} is missing", relationship.target),
                    ));
                break;
            }
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Container,
                    Fidelity::Omitted,
                    "picture package part is missing and was omitted",
                )
                .in_part(part)
                .with_detail("target", relationship.target.clone()),
            );
            return Ok(());
        }

        let materialized = match state.image_cache.get(&relationship.target).copied() {
            Some(ImageCacheEntry::Unsupported(error)) => {
                if is_preferred_svg {
                    state.diagnostics.push(preferred_svg_fallback_diagnostic(
                        part,
                        preferred_svg_error_reason(error),
                        has_fallback,
                    ));
                    continue;
                }
                if selected_image.is_some() {
                    state
                        .diagnostics
                        .push(runtime_fallback_unavailable_diagnostic(
                            part,
                            preferred_svg_error_reason(error),
                        ));
                    break;
                }
                state.diagnostics.push(unsupported_image_diagnostic(
                    &relationship.target,
                    error,
                    declared_content_type.is_some(),
                ));
                return Ok(());
            }
            Some(ImageCacheEntry::Object(index)) => {
                let Some((media_type, bytes)) = state
                    .objects
                    .get(index)
                    .and_then(|object| image_visual_data(&object.visual))
                else {
                    return Err(format_error(part, "embedded image cache is inconsistent"));
                };
                if is_preferred_svg && media_type != "image/svg+xml" {
                    state.diagnostics.push(preferred_svg_fallback_diagnostic(
                        part,
                        format!("relationship {relationship_id} resolves to {media_type}, not SVG"),
                        has_fallback,
                    ));
                    continue;
                }
                reserve_materialized_image_bytes(
                    &mut state.materialized_image_bytes,
                    bytes.len(),
                    package.limits().max_total_uncompressed_bytes,
                    &relationship.target,
                )?;
                let bytes = clone_image_bytes(bytes, &relationship.target)?;
                (media_type.to_owned(), bytes, false)
            }
            None => {
                let bytes = package.required_part(&relationship.target)?;
                let identified =
                    pptx_image_media_type(declared_content_type, &relationship.target, &bytes);
                let media_type = match identified {
                    Ok(media_type) => media_type.to_owned(),
                    Err(error) => {
                        state.image_cache.insert(
                            relationship.target.clone(),
                            ImageCacheEntry::Unsupported(error),
                        );
                        if is_preferred_svg {
                            state.diagnostics.push(preferred_svg_fallback_diagnostic(
                                part,
                                preferred_svg_error_reason(error),
                                has_fallback,
                            ));
                            continue;
                        }
                        if selected_image.is_some() {
                            state
                                .diagnostics
                                .push(runtime_fallback_unavailable_diagnostic(
                                    part,
                                    preferred_svg_error_reason(error),
                                ));
                            break;
                        }
                        state.diagnostics.push(unsupported_image_diagnostic(
                            &relationship.target,
                            error,
                            declared_content_type.is_some(),
                        ));
                        return Ok(());
                    }
                };
                if is_preferred_svg && media_type != "image/svg+xml" {
                    state.diagnostics.push(preferred_svg_fallback_diagnostic(
                        part,
                        format!("relationship {relationship_id} resolves to {media_type}, not SVG"),
                        has_fallback,
                    ));
                    continue;
                }
                reserve_materialized_image_bytes(
                    &mut state.materialized_image_bytes,
                    bytes.len(),
                    package.limits().max_total_uncompressed_bytes,
                    &relationship.target,
                )?;
                (media_type, bytes.into_vec(), true)
            }
        };
        if is_preferred_svg {
            selected_image = Some((
                relationship.target.clone(),
                materialized.0,
                materialized.1,
                materialized.2,
            ));
            if has_fallback {
                continue;
            }
        } else if selected_image.is_some() {
            raster_fallback = Some((materialized.0, materialized.1));
        } else {
            selected_image = Some((
                relationship.target.clone(),
                materialized.0,
                materialized.1,
                materialized.2,
            ));
        }
        break;
    }
    let Some((image_target, media_type, image_bytes, cache_new_object)) = selected_image else {
        return Ok(());
    };
    let z = state.take_z();
    let mapping = if picture.is_background {
        MappingQuality::Derived
    } else if picture.shape_id.is_some() {
        MappingQuality::Exact
    } else {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Approximate,
                "picture has no source ID; a session-local mapping was derived",
            )
            .in_part(part),
        );
        MappingQuality::Derived
    };
    let shape_id = picture.shape_id.unwrap_or(numeric_id);
    let parent_numeric_id = picture.parent_numeric_id;
    let transform = if picture.has_transform {
        shape_transform(
            bounds,
            picture.rotation_degrees,
            picture.flip_horizontal,
            picture.flip_vertical,
        )
    } else {
        inherited_transform
    };
    let geometry = if let Some(geometry) = picture.geometry.custom_geometry(bounds) {
        geometry
    } else if let Some(preset) = picture.geometry.preset.as_deref() {
        if let Some((geometry, _, fidelity)) = resolve_preset_geometry(
            preset,
            bounds,
            &picture.geometry.adjustments,
            picture.stroke_width,
            None,
            None,
        ) {
            if fidelity == Fidelity::Approximate {
                state.diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Render,
                        Fidelity::Approximate,
                        format!("PPTX picture preset `{preset}` uses an approximate geometry"),
                    )
                    .in_part(part),
                );
            }
            geometry
        } else {
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Render,
                    Fidelity::Unsupported,
                    format!("PPTX picture preset `{preset}` uses a rectangle fallback"),
                )
                .in_part(part),
            );
            Geometry::Rectangle
        }
    } else {
        Geometry::Rectangle
    };
    let has_outline = picture.stroke_width > 0.0 && !matches!(&picture.stroke, Paint::None);
    let has_picture_shape = geometry != Geometry::Rectangle
        || has_outline
        || picture.image_tile
        || picture.image_mapping != crate::model::ImageFillMapping::default();
    let can_use_painted_shape = has_picture_shape
        && raster_fallback.is_none()
        && picture.color_change_from.is_none()
        && picture.color_change_to.is_none();
    let visual = if can_use_painted_shape {
        let visual = Visual::PaintedShape {
            geometry: geometry.clone(),
            fill: Paint::Image {
                mapping: (picture.image_tile
                    || picture.image_mapping != crate::model::ImageFillMapping::default())
                .then(|| Box::new(picture.image_mapping.clone())),
                media_type,
                bytes: image_bytes,
                crop: picture.crop,
                tile: picture.image_tile,
                tile_width: None,
                tile_height: None,
            },
            stroke: picture.stroke.clone(),
            stroke_width: picture.stroke_width,
        };
        let style = StrokeStyle {
            cap: picture.line_cap,
            join: picture.line_join,
            compound: picture.line_compound,
            alignment: picture.line_alignment,
            miter_limit: picture.miter_limit.max(1.0),
            dash_offset: 0.0,
            dash: drawingml_dash_lengths(picture.dash_pattern.lengths(), picture.stroke_width),
        };
        if style == StrokeStyle::default() {
            visual
        } else {
            Visual::StrokeStyle {
                style,
                visual: Box::new(visual),
            }
        }
    } else {
        let visual = if let Some((fallback_media_type, fallback_bytes)) = raster_fallback {
            Visual::ImageWithFallback {
                media_type,
                bytes: image_bytes,
                fallback_media_type,
                fallback_bytes,
                crop: picture.crop,
            }
        } else {
            Visual::Image {
                media_type,
                bytes: image_bytes,
                crop: picture.crop,
            }
        };
        let visual = match (picture.color_change_from, picture.color_change_to) {
            (Some(from), Some(to)) => Visual::ImageColorChange {
                from,
                to,
                use_alpha: picture.color_change_use_alpha,
                visual: Box::new(visual),
            },
            (None, None) => visual,
            _ => {
                state.diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Parse,
                        Fidelity::Approximate,
                        "picture color-change effect is incomplete and was omitted",
                    )
                    .in_part(part),
                );
                visual
            }
        };
        if has_picture_shape {
            if has_outline {
                state.diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Render,
                        Fidelity::Approximate,
                        "picture outline was omitted because the image requires a runtime fallback or color-change effect",
                    )
                    .in_part(part),
                );
            }
            Visual::Effect {
                shadow: None,
                clip: Some(geometry),
                visual: Box::new(visual),
            }
        } else {
            visual
        }
    };
    let mut effects = picture.effects.finish();
    if picture.style_shadow && effects.outer_shadow.is_none() {
        effects.outer_shadow = Some(OuterShadow {
            shadow: Shadow {
                color: 0x0000_0059,
                blur: 4.0,
                offset_x: 2.0,
                offset_y: 2.0,
            },
            scale_x: 1.0,
            scale_y: 1.0,
            skew_x: 0.0,
            skew_y: 0.0,
            alignment: 7,
        });
    }
    let visual = if let Some((kind, media_type, bytes)) = embedded_media {
        Visual::Media {
            kind,
            media_type,
            bytes,
            poster: Box::new(visual),
        }
    } else {
        visual
    };
    let visual = effects.wrap(visual, None);
    let visual = if transform == AffineTransform::IDENTITY && picture.opacity == 1.0 {
        visual
    } else {
        Visual::Layer {
            transform,
            opacity: picture.opacity,
            blend_mode: crate::model::BlendMode::Normal,
            visual: Box::new(visual),
        }
    };
    let object = Object {
        numeric_id,
        parent_numeric_id,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
        kind: ObjectKind::Image,
        unit_index,
        bounds,
        z,
        text: None,
        source: SourceRef {
            part: part.to_owned(),
            mapping,
            locator: SourceLocator::PptxShape {
                shape_id,
                row: None,
                column: None,
                text_range: None,
                metadata: picture.metadata,
            },
        },
        visual,
    };
    if replacement_id.is_some() {
        invalidate_cached_image_object(&mut state.image_cache, object_index);
        *state
            .objects
            .get_mut(object_index)
            .ok_or_else(|| format_error(part, "placeholder object mapping is missing"))? = object;
    } else {
        state.objects.push(object);
    }
    if cache_new_object || replacement_id.is_some() {
        state
            .image_cache
            .entry(image_target)
            .or_insert(ImageCacheEntry::Object(object_index));
    }
    register_placeholder_object(
        state,
        unit_index,
        picture.placeholder_index,
        picture.placeholder_type.as_deref(),
        numeric_id,
        None,
        picture
            .geometry
            .preset
            .map(|preset| (preset, picture.geometry.adjustments)),
    );
    state.pending_placeholder_objects.remove(&numeric_id);
    Ok(())
}

fn image_visual_data(visual: &Visual) -> Option<(&str, &[u8])> {
    match visual {
        Visual::Image {
            media_type, bytes, ..
        }
        | Visual::ImageWithFallback {
            media_type, bytes, ..
        } => Some((media_type, bytes)),
        Visual::PaintedShape {
            fill: Paint::Image {
                media_type, bytes, ..
            },
            ..
        } => Some((media_type, bytes)),
        Visual::Layer { visual, .. }
        | Visual::Effect { visual, .. }
        | Visual::TextLayout { visual, .. }
        | Visual::StrokeStyle { visual, .. }
        | Visual::AdvancedEffect { visual, .. }
        | Visual::ImageColorChange { visual, .. } => image_visual_data(visual),
        Visual::Media { poster, .. } => image_visual_data(poster),
        _ => None,
    }
}

fn invalidate_cached_image_object(
    image_cache: &mut HashMap<String, ImageCacheEntry>,
    object_index: usize,
) {
    image_cache.retain(|_, entry| *entry != ImageCacheEntry::Object(object_index));
}

fn unsupported_image_diagnostic(
    part: &str,
    error: OfficeImageError,
    declared_by_content_types: bool,
) -> Diagnostic {
    let (phase, fidelity, message) = match error {
        OfficeImageError::UnsupportedFormat => (
            Phase::Parse,
            Fidelity::Omitted,
            if declared_by_content_types {
                "picture content type declared by [Content_Types].xml is outside the supported modern Office/Open XML image set"
            } else {
                "picture format is outside the supported modern Office/Open XML image set"
            },
        ),
        OfficeImageError::SignatureMismatch => (
            Phase::Parse,
            Fidelity::Omitted,
            if declared_by_content_types {
                "picture bytes do not match the image format declared by [Content_Types].xml"
            } else {
                "picture bytes do not match the image format declared by the package part name"
            },
        ),
        OfficeImageError::DisabledByOffice => (
            Phase::Security,
            Fidelity::Blocked,
            "DisabledByOffice: EPS and legacy PICT picture formats are disabled by current Microsoft Office",
        ),
    };
    Diagnostic::warning(DiagnosticCode::UnsupportedFeature, phase, fidelity, message).in_part(part)
}

fn pptx_image_media_type(
    declared_content_type: Option<&str>,
    part: &str,
    bytes: &[u8],
) -> Result<&'static str, OfficeImageError> {
    super::presentation_image::recover_office_image_signature(
        declared_content_type.map_or_else(
            || office_image_media_type(part, bytes),
            |media_type| office_image_media_type_from_mime(media_type, bytes),
        ),
        bytes,
    )
}

fn preferred_svg_fallback_diagnostic(
    part: &str,
    reason: impl AsRef<str>,
    has_fallback: bool,
) -> Diagnostic {
    let outcome = if has_fallback {
        "raster fallback will be attempted"
    } else {
        "picture was omitted because no raster fallback is available"
    };
    Diagnostic::warning(
        DiagnosticCode::UnsupportedFeature,
        Phase::Parse,
        if has_fallback {
            Fidelity::Approximate
        } else {
            Fidelity::Omitted
        },
        format!(
            "preferred SVG {} was unavailable; {outcome}",
            reason.as_ref()
        ),
    )
    .in_part(part)
}

fn runtime_fallback_unavailable_diagnostic(part: &str, reason: impl AsRef<str>) -> Diagnostic {
    Diagnostic::warning(
        DiagnosticCode::UnsupportedFeature,
        Phase::Parse,
        Fidelity::Approximate,
        format!(
            "raster fallback {} is unavailable; preferred SVG will be used without runtime fallback",
            reason.as_ref()
        ),
    )
    .in_part(part)
}

const fn preferred_svg_error_reason(error: OfficeImageError) -> &'static str {
    match error {
        OfficeImageError::UnsupportedFormat => "format is unsupported",
        OfficeImageError::SignatureMismatch => "bytes do not contain a valid SVG root",
        OfficeImageError::DisabledByOffice => "format is disabled by current Microsoft Office",
    }
}

struct PictureBullet {
    paragraph_index: usize,
    relationship_id: String,
    size: f32,
}

fn resolve_shape_paragraph_layouts(
    shape: &mut ShapeState,
) -> (Vec<TextParagraphLayout>, Vec<PictureBullet>) {
    let mut layouts = std::mem::take(&mut shape.paragraph_layouts);
    let styles = std::mem::take(&mut shape.paragraph_layout_styles);
    let run_starts = std::mem::take(&mut shape.paragraph_run_starts);
    let mut empty_prefixes = Vec::new();
    let mut picture_bullets = Vec::new();

    for (index, layout) in layouts.iter_mut().enumerate() {
        let style = styles.get(index);
        let start = run_starts
            .get(index)
            .copied()
            .unwrap_or(0)
            .min(shape.runs.len());
        let end = run_starts
            .get(index + 1)
            .copied()
            .unwrap_or(shape.runs.len())
            .min(shape.runs.len());
        let has_generated_prefix = style.and_then(|style| style.bullet.as_ref()).is_some()
            && shape
                .runs
                .get(start)
                .is_some_and(|run| run.text.ends_with('\t'));
        let content_start = start
            .saturating_add(usize::from(has_generated_prefix))
            .min(end);
        let has_content = shape.runs[content_start..end].iter().any(|run| {
            run.text
                .chars()
                .any(|character| !matches!(character, '\r' | '\n' | '\u{2028}'))
        });
        if has_generated_prefix && !has_content {
            empty_prefixes.push(start);
        }
        let font_size = shape.runs[content_start..end]
            .iter()
            .filter(|run| {
                run.text
                    .chars()
                    .any(|character| !matches!(character, '\r' | '\n' | '\u{2028}'))
            })
            .map(|run| run.font_size)
            .reduce(f32::max)
            .or_else(|| style.and_then(|style| style.font_size))
            .unwrap_or(shape.font_size)
            .max(1.0);

        if let Some(ParagraphBullet::Image(relationship_id)) =
            style.and_then(|style| style.bullet.as_ref())
        {
            picture_bullets.push(PictureBullet {
                paragraph_index: index,
                relationship_id: relationship_id.clone(),
                size: style
                    .and_then(|style| style.bullet_size.as_ref())
                    .map_or(font_size * 0.6, |size| size.resolve(font_size)),
            });
        }

        if has_generated_prefix
            && has_content
            && let Some(prefix) = shape.runs.get_mut(start)
        {
            // DrawingML's default buSzTx follows the paragraph text size, not the
            // list-level defRPr size that was available before the first run.
            prefix.font_size = style
                .and_then(|style| style.bullet_size.as_ref())
                .map_or(font_size, |size| size.resolve(font_size));
        }
        layout.line_height = style
            .and_then(|style| style.line_spacing.as_ref())
            .map_or(font_size * 1.2, |spacing| {
                spacing.resolve_for_layout(Some(font_size))
            });
        layout.space_before = style
            .and_then(|style| style.space_before.as_ref())
            .map_or(0.0, |spacing| spacing.resolve(font_size));
        layout.space_after = style
            .and_then(|style| style.space_after.as_ref())
            .map_or(0.0, |spacing| spacing.resolve(font_size));
    }

    if !empty_prefixes.is_empty() {
        for index in empty_prefixes.into_iter().rev() {
            shape.runs.remove(index);
        }
        shape.text = shape
            .runs
            .iter()
            .map(|run| run.text.as_str())
            .collect::<String>()
            .replace('\u{2028}', "\n");
    }

    if !shape.text_space_first_last_paragraph {
        if let Some(first) = layouts.first_mut() {
            first.space_before = 0.0;
        }
        if let Some(last) = layouts.last_mut() {
            last.space_after = 0.0;
        }
    }
    (layouts, picture_bullets)
}

fn picture_bullet_bounds(
    shape_bounds: Rect,
    layouts: &[TextParagraphLayout],
    bullets: &[PictureBullet],
    vertical_align: TextVerticalAlign,
    inset_left: f32,
    inset_right: f32,
    inset_top: f32,
    inset_bottom: f32,
) -> Vec<(String, Rect)> {
    let content_height = layouts.iter().fold(0.0, |height, paragraph| {
        height + paragraph.space_before + paragraph.line_height + paragraph.space_after
    });
    let available_height = (shape_bounds.height - inset_top - inset_bottom).max(0.0);
    let mut y = shape_bounds.y
        + inset_top
        + match vertical_align {
            TextVerticalAlign::Top => 0.0,
            TextVerticalAlign::Center => (available_height - content_height) / 2.0,
            TextVerticalAlign::Bottom => available_height - content_height,
        };
    let mut output = Vec::with_capacity(bullets.len());
    for (index, paragraph) in layouts.iter().enumerate() {
        y += paragraph.space_before;
        if let Some(bullet) = bullets
            .iter()
            .find(|bullet| bullet.paragraph_index == index)
        {
            let size = bullet.size.min(paragraph.line_height.max(1.0));
            let left =
                shape_bounds.x + inset_left + paragraph.margin_left + paragraph.first_line_indent;
            let right = shape_bounds.x + shape_bounds.width - inset_right - paragraph.margin_right;
            let x = match paragraph.align {
                TextAlign::Center => (left + right - size) / 2.0,
                TextAlign::End => right - size,
                _ => left,
            };
            output.push((
                bullet.relationship_id.clone(),
                Rect {
                    x,
                    y: y + (paragraph.line_height - size) / 2.0,
                    width: size,
                    height: size,
                },
            ));
        }
        y += paragraph.line_height + paragraph.space_after;
    }
    output
}

fn preset_text_height_overflows(
    shape: &ShapeState,
    paragraphs: &[TextParagraphLayout],
    available_width: f32,
    available_height: f32,
    font_metrics: &FontMetricTable,
) -> bool {
    if shape.text_orientation != TextOrientation::Horizontal
        || !shape.text_wrap
        || paragraphs.is_empty()
        || !available_width.is_finite()
        || available_width <= 0.0
        || !available_height.is_finite()
        || available_height <= 0.0
    {
        return false;
    }

    let font_scale = shape.text_font_scale.clamp(0.01, 1.0);
    let scale = if font_scale < 1.0 {
        font_scale
    } else if shape.text_auto_fit == TextAutoFit::Shrink {
        shape.text_min_scale.clamp(0.01, 1.0)
    } else {
        font_scale
    };
    let fallback_runs = shape.runs.is_empty().then(|| {
        vec![TextRun {
            paint: None,
            east_asian_line_breaks: true,
            text: shape.text.clone(),
            font_family: shape.font_family.clone(),
            font_size: shape.font_size,
            color: shape.font_color,
            bold: shape.bold,
            italic: shape.italic,
            underline: shape.underline,
            strikethrough: shape.strikethrough,
            highlight: shape.highlight,
            baseline_shift: shape.baseline_shift,
            letter_spacing: shape.letter_spacing,
            horizontal_scale: 1.0,
        }]
    });
    let runs = fallback_runs.as_deref().unwrap_or(&shape.runs);
    let required_height = estimated_drawingml_text_height(
        runs,
        paragraphs,
        shape.font_size,
        available_width,
        shape.text_wrap,
        scale,
        shape.text_line_spacing_reduction,
        font_metrics,
    );
    // Multi-column text can continue into the next column before it overflows
    // the preset's body area. Single-column shapes, by far the common case,
    // retain the exact renderer-aligned height comparison.
    required_height > available_height * shape.text_column_count.max(1) as f32 + f32::EPSILON
}

struct ShapePlacement<'a> {
    part: &'a str,
    unit_index: u32,
    parent_numeric_id: Option<u32>,
    group_fill: Option<Paint>,
    inherited: bool,
    object_limit: usize,
}

fn push_shape(
    shape: ShapeState,
    placement: ShapePlacement<'_>,
    state: &mut PptxParseState,
    font_metrics: &FontMetricTable,
    relationships: &HashMap<&str, &Relationship>,
    package: &Package<'_>,
    content_types: &ContentTypes,
) -> Result<(), Diagnostic> {
    let ShapePlacement {
        part,
        unit_index,
        parent_numeric_id,
        group_fill,
        inherited,
        object_limit,
    } = placement;
    let mut shape = shape;
    if shape.text_warp.is_some()
        && matches!(
            shape.text_orientation,
            TextOrientation::VerticalRl | TextOrientation::VerticalLr
        )
        && (shape.rotation_degrees.abs() % 180.0 - 90.0).abs() < 0.01
    {
        shape.text_orientation = TextOrientation::Horizontal;
    }
    if shape.math_present && shape.omit_empty_math && shape.text.trim().is_empty() {
        return Ok(());
    }
    let replacement_id = find_placeholder_object(
        state,
        unit_index,
        shape.placeholder_index,
        shape.placeholder_type.as_deref(),
    );
    if replacement_id.is_none() && state.objects.len() >= object_limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "document exceeds the configured object limit",
        )
        .in_part(part));
    }
    let bounds = resolved_shape_bounds(&shape, unit_index, state);
    if !bounds.is_valid() {
        return Err(format_error(part, "shape has invalid bounds"));
    }
    let text_transform_bounds = match (
        shape.text_x,
        shape.text_y,
        shape.text_width,
        shape.text_height,
    ) {
        (Some(x), Some(y), Some(width), Some(height)) => Some(Rect {
            x: x as f32 / EMU_PER_CSS_PIXEL,
            y: y as f32 / EMU_PER_CSS_PIXEL,
            width: width as f32 / EMU_PER_CSS_PIXEL,
            height: height as f32 / EMU_PER_CSS_PIXEL,
        }),
        _ => None,
    };
    if text_transform_bounds.is_some_and(|bounds| !bounds.is_valid()) {
        return Err(format_error(
            part,
            "shape text transform has invalid bounds",
        ));
    }
    if text_transform_bounds.is_some()
        && (shape.outer_shadow.is_some()
            || shape.inner_shadow.is_some()
            || shape.glow.is_some()
            || shape.reflection.is_some()
            || shape.soft_edge.is_some()
            || shape.three_d.is_some())
    {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Omitted,
                "PPTX shape effects were omitted because the shape uses an independent text transform",
            )
            .in_part(part),
        );
    }
    let mut preset_text_rectangle = None;
    if let Some(preset) = shape.preset.as_deref() {
        if let Some((geometry, text_rectangle, fidelity)) = resolve_preset_geometry(
            preset,
            bounds,
            &shape.preset_adjustments,
            shape.stroke_width,
            shape.head_arrow,
            shape.tail_arrow,
        ) {
            shape.geometry = geometry;
            preset_text_rectangle = text_rectangle;
            if fidelity == Fidelity::Approximate {
                state.diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::UnsupportedFeature,
                        Phase::Render,
                        Fidelity::Approximate,
                        format!("PPTX preset shape `{preset}` uses an approximate geometry"),
                    )
                    .in_part(part),
                );
            }
        } else if shape.is_connector {
            shape.geometry = connector_geometry(
                "straightConnector1",
                bounds,
                &shape.preset_adjustments,
                shape.stroke_width,
                shape.head_arrow,
                shape.tail_arrow,
            )
            .map(|(geometry, _)| geometry)
            .unwrap_or(Geometry::Line);
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Render,
                    Fidelity::Approximate,
                    format!("PPTX connector preset `{preset}` uses a straight connector fallback"),
                )
                .in_part(part),
            );
        } else {
            state.diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::UnsupportedFeature,
                    Phase::Render,
                    Fidelity::Unsupported,
                    format!("PPTX preset shape `{preset}` uses a rectangle fallback"),
                )
                .in_part(part),
            );
        }
    }
    let inherited_vertical_align = replacement_id
        .and_then(|numeric_id| usize::try_from(numeric_id).ok())
        .and_then(|index| state.objects.get(index))
        .and_then(|object| visual_text_vertical_align(&object.visual));
    let vertical_align = shape
        .vertical_align
        .or(inherited_vertical_align)
        .unwrap_or(TextVerticalAlign::Top);
    if shape.text_font_scale < 1.0 {
        let scale = shape.text_font_scale;
        // PowerPoint rounds persisted AutoFit sizes to whole points. Keep the
        // shared scale for paragraph spacing, expressing the rounded size in
        // its unscaled coordinate system before line measurement and wrapping.
        let rounded = |size: f32| {
            (size * scale / POINTS_TO_CSS_PIXELS).round().max(1.0) * POINTS_TO_CSS_PIXELS / scale
        };
        shape.font_size = rounded(shape.font_size);
        for run in &mut shape.runs {
            run.font_size = rounded(run.font_size);
        }
        for style in shape
            .paragraph_layout_styles
            .iter_mut()
            .chain(shape.paragraph_layout_style.iter_mut())
        {
            style.font_size = style.font_size.map(rounded);
        }
    }
    let (paragraph_layouts, picture_bullets) = resolve_shape_paragraph_layouts(&mut shape);
    // ponytail: parser-side placement follows authored paragraph rows; promote picture bullets
    // into TextLayout only if wrapped/autofit bullet paragraphs need runtime-font exactness.
    let picture_bullets = picture_bullet_bounds(
        bounds,
        &paragraph_layouts,
        &picture_bullets,
        vertical_align,
        shape.text_inset_left,
        shape.text_inset_right,
        shape.text_inset_top,
        shape.text_inset_bottom,
    );
    let discard_preset_vertical_insets = preset_text_rectangle.is_some_and(|rectangle| {
        let preset_inset_left = rectangle.x.max(0.0);
        let preset_inset_right = (bounds.width - rectangle.x - rectangle.width).max(0.0);
        let preset_inset_top = rectangle.y.max(0.0);
        let preset_inset_bottom = (bounds.height - rectangle.y - rectangle.height).max(0.0);
        let preset_available_width = bounds.width
            - shape.text_inset_left
            - shape.text_inset_right
            - preset_inset_left
            - preset_inset_right;
        let preset_available_height = bounds.height
            - shape.text_inset_top
            - shape.text_inset_bottom
            - preset_inset_top
            - preset_inset_bottom;
        preset_text_height_overflows(
            &shape,
            &paragraph_layouts,
            preset_available_width,
            preset_available_height,
            font_metrics,
        )
    });
    let (shadow, outer_shadow) = match shape.outer_shadow {
        Some(effect) if drawingml_outer_shadow_is_identity(&effect) => (Some(effect.shadow), None),
        effect @ Some(_) => (None, effect),
        None => (shape.shadow, None),
    };
    let inner_shadow = shape.inner_shadow;
    let glow = shape.glow;
    let reflection = shape.reflection;
    let soft_edge = shape.soft_edge;
    let three_d = shape.three_d;
    let numeric_id = replacement_id.unwrap_or(state.objects.len() as u32);
    let z = state.take_z();
    let mapping = if shape.shape_id.is_some() {
        MappingQuality::Exact
    } else {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Approximate,
                "shape has no source ID; a session-local mapping was derived",
            )
            .in_part(part),
        );
        MappingQuality::Derived
    };
    let shape_id = shape.shape_id.unwrap_or(numeric_id);
    let shape_depth = shape.depth;
    let placeholder_prompt = inherited
        && (shape.is_custom_prompt
            || is_content_placeholder_type(shape.placeholder_type.as_deref())
            || shape
                .placeholder_type
                .as_deref()
                .is_some_and(is_fixed_placeholder_type));
    let text = if placeholder_prompt {
        String::new()
    } else {
        shape.text.clone()
    };
    let text_length = u32::try_from(text.chars().count()).unwrap_or(u32::MAX);
    let has_text = !text.is_empty();
    let separate_shape_and_text = text_transform_bounds.is_some()
        || (has_text && (shape.flip_horizontal || shape.flip_vertical));
    let transform = shape_transform(
        bounds,
        shape.rotation_degrees,
        shape.flip_horizontal,
        shape.flip_vertical,
    );
    // DrawingML style references provide the base paint; explicit spPr values win per property.
    let fill = if shape.use_group_fill {
        group_fill.unwrap_or(Paint::None)
    } else if shape.explicit_fill {
        shape.fill.clone()
    } else {
        shape
            .style_fill
            .clone()
            .unwrap_or_else(|| shape.fill.clone())
    };
    let stroke = if shape.explicit_stroke {
        shape.stroke.clone()
    } else {
        shape
            .style_stroke
            .clone()
            .unwrap_or_else(|| shape.stroke.clone())
    };
    let stroke_width = if shape.explicit_stroke_width {
        shape.stroke_width
    } else {
        shape.style_stroke_width.unwrap_or(shape.stroke_width)
    };
    let fill = if shape.is_connector {
        Paint::None
    } else {
        fill
    };
    let text_font_size = shape
        .runs
        .iter()
        .filter(|run| !run.text.is_empty() && run.text != "\n")
        .map(|run| run.font_size)
        .reduce(f32::max)
        .unwrap_or(shape.font_size)
        .max(1.0);
    // ponytail: formula-only, untransformed shapes reuse the DOCX equation boxes;
    // mixed inline math needs PPTX paragraph composition before it can use them.
    let math_region = Rect {
        x: bounds.x + shape.text_inset_left.max(0.0),
        y: bounds.y + shape.text_inset_top.max(0.0),
        width: (bounds.width - shape.text_inset_left.max(0.0) - shape.text_inset_right.max(0.0))
            .max(0.0),
        height: (bounds.height - shape.text_inset_top.max(0.0) - shape.text_inset_bottom.max(0.0))
            .max(0.0),
    };
    let structured_math = shape.math_root.as_ref().is_some_and(|(root, formula)| {
        root.needs_box_layout() && shape.text.trim() == formula.trim()
    });
    let math_layout = if structured_math
        && !placeholder_prompt
        && !separate_shape_and_text
        && transform == AffineTransform::IDENTITY
        && shape.text_orientation == TextOrientation::Horizontal
        && shape.text_rotation_degrees == 0.0
        && shape.text_warp.is_none()
        && math_region.width > 0.0
        && math_region.height > 0.0
    {
        let (root, _) = shape.math_root.as_ref().unwrap();
        shape
            .runs
            .iter()
            .rev()
            .find(|run| !run.text.trim().is_empty())
            .and_then(|run| {
                match super::docx::layout_omml_for_shape(
                    root,
                    run,
                    font_metrics,
                    math_region.width,
                    math_region.height,
                    part,
                ) {
                    Ok((width, height, children))
                        if width.is_finite()
                            && height.is_finite()
                            && width <= math_region.width + 0.01
                            && height <= math_region.height + 0.01
                            && !children.is_empty()
                            && children.iter().all(|child| child.bounds.is_valid())
                            && children.len()
                                <= object_limit.saturating_sub(
                                    state.objects.len() + usize::from(replacement_id.is_none()),
                                ) =>
                    {
                        Some((width, height, children))
                    }
                    Ok(_) => None,
                    Err(error) => {
                        state.diagnostics.push(
                            Diagnostic::warning(
                                DiagnosticCode::UnsupportedFeature,
                                Phase::Render,
                                Fidelity::Approximate,
                                format!(
                                    "PPTX equation boxes fell back to linear text: {}",
                                    error.message
                                ),
                            )
                            .in_part(part),
                        );
                        None
                    }
                }
            })
    } else {
        None
    };
    if structured_math && math_layout.is_none() {
        state.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Render,
                Fidelity::Approximate,
                "PPTX structured formula uses linear text in this shape",
            )
            .in_part(part),
        );
    }
    let paragraph_style = shape.paragraph_layout_style.as_ref();
    let paragraph_font_size = paragraph_style
        .and_then(|style| style.font_size)
        .unwrap_or(text_font_size)
        .max(1.0);
    let line_height = paragraph_style
        .and_then(|style| style.line_spacing.as_ref())
        .map_or(text_font_size * 1.2, |spacing| {
            spacing.resolve_for_layout(Some(paragraph_font_size))
        });
    let paragraph_spacing = paragraph_style.map_or(0.0, |style| {
        style
            .space_before
            .as_ref()
            .map_or(0.0, |spacing| spacing.resolve(paragraph_font_size))
            + style
                .space_after
                .as_ref()
                .map_or(0.0, |spacing| spacing.resolve(paragraph_font_size))
    });
    let margin_left = paragraph_style
        .and_then(|style| style.margin_left)
        .unwrap_or(0.0)
        .max(0.0);
    let margin_right = paragraph_style
        .and_then(|style| style.margin_right)
        .unwrap_or(0.0)
        .max(0.0);
    let first_line_indent = paragraph_style
        .and_then(|style| style.indent)
        .unwrap_or(0.0);
    let default_tab_stop = paragraph_style
        .and_then(|style| style.default_tab_stop)
        .unwrap_or(36.0)
        .max(1.0);
    let uses_paragraph_layouts = !paragraph_layouts.is_empty();
    let visual = if math_layout.is_some() {
        Visual::PaintedShape {
            geometry: shape.geometry.clone(),
            fill: fill.clone(),
            stroke: stroke.clone(),
            stroke_width,
        }
    } else if has_text || placeholder_prompt {
        let mut run_effects = if shape.runs.is_empty() {
            vec![TextEffect::default()]
        } else {
            shape.run_effects
        };
        let mut runs = if shape.runs.is_empty() {
            vec![TextRun {
                paint: None,
                east_asian_line_breaks: true,
                text: text.clone(),
                font_family: shape.font_family,
                font_size: shape.font_size,
                color: shape.font_color,
                bold: shape.bold,
                italic: shape.italic,
                underline: shape.underline,
                strikethrough: shape.strikethrough,
                highlight: shape.highlight,
                baseline_shift: shape.baseline_shift,
                letter_spacing: shape.letter_spacing,
                horizontal_scale: 1.0,
            }]
        } else {
            shape.runs
        };
        if placeholder_prompt {
            for run in &mut runs {
                run.text.clear();
            }
        }
        run_effects.resize(runs.len(), TextEffect::default());
        let visual = Visual::RichText {
            geometry: if separate_shape_and_text {
                Geometry::Rectangle
            } else {
                shape.geometry.clone()
            },
            fill: if separate_shape_and_text {
                Paint::None
            } else {
                fill.clone()
            },
            stroke: if separate_shape_and_text {
                Paint::None
            } else {
                stroke.clone()
            },
            stroke_width: if separate_shape_and_text {
                0.0
            } else {
                stroke_width
            },
            align: shape.align,
            line_height,
            runs,
        };
        if run_effects
            .iter()
            .all(|effect| *effect == TextEffect::default())
        {
            visual
        } else {
            Visual::TextEffects {
                effects: run_effects,
                visual: Box::new(visual),
            }
        }
    } else {
        Visual::PaintedShape {
            geometry: shape.geometry.clone(),
            fill: fill.clone(),
            stroke: stroke.clone(),
            stroke_width,
        }
    };
    let visual = if (has_text || placeholder_prompt) && math_layout.is_none() {
        let (preset_inset_left, preset_inset_right, mut preset_inset_top, mut preset_inset_bottom) =
            preset_text_rectangle.map_or((0.0, 0.0, 0.0, 0.0), |rectangle| {
                let left = rectangle.x.max(0.0);
                let right = (bounds.width - rectangle.x - rectangle.width).max(0.0);
                let top = rectangle.y.max(0.0);
                let bottom = (bounds.height - rectangle.y - rectangle.height).max(0.0);
                match shape.text_orientation {
                    TextOrientation::Rotated90 => (top, bottom, right, left),
                    TextOrientation::Rotated270 => (bottom, top, left, right),
                    _ => (left, right, top, bottom),
                }
            });
        if discard_preset_vertical_insets {
            preset_inset_top = 0.0;
            preset_inset_bottom = 0.0;
        }
        Visual::TextLayout {
            layout: TextLayout {
                direction: shape.text_direction,
                orientation: shape.text_orientation,
                auto_fit: shape.text_auto_fit,
                vertical_align,
                min_scale: shape.text_min_scale,
                font_scale: shape.text_font_scale,
                line_spacing_reduction: shape.text_line_spacing_reduction,
                column_count: shape.text_column_count,
                column_spacing: shape.text_column_spacing,
                rotation_degrees: shape.text_rotation_degrees,
                horizontal_overflow: shape.text_horizontal_overflow,
                vertical_overflow: shape.text_vertical_overflow,
                wrap: shape.text_wrap,
                warp: shape.text_warp,
                default_tab_stop,
                hanging_indent: 0.0,
                paragraph_spacing: if uses_paragraph_layouts {
                    0.0
                } else {
                    paragraph_spacing
                },
                inset_left: shape.text_inset_left + preset_inset_left,
                inset_right: shape.text_inset_right + preset_inset_right,
                inset_top: shape.text_inset_top + preset_inset_top,
                inset_bottom: shape.text_inset_bottom + preset_inset_bottom,
                margin_left: if uses_paragraph_layouts {
                    0.0
                } else {
                    margin_left
                },
                margin_right: if uses_paragraph_layouts {
                    0.0
                } else {
                    margin_right
                },
                first_line_indent: if uses_paragraph_layouts {
                    0.0
                } else {
                    first_line_indent
                },
                paragraphs: paragraph_layouts,
                ..TextLayout::default()
            },
            visual: Box::new(visual),
        }
    } else {
        visual
    };
    let stroke_style = StrokeStyle {
        cap: shape.line_cap,
        join: shape.line_join,
        compound: shape.line_compound,
        alignment: shape.line_alignment,
        miter_limit: shape.miter_limit.max(1.0),
        dash_offset: 0.0,
        dash: drawingml_dash_lengths(shape.dash_pattern.lengths(), stroke_width),
    };
    let visual = if stroke_style == StrokeStyle::default() || separate_shape_and_text {
        visual
    } else {
        Visual::StrokeStyle {
            style: stroke_style.clone(),
            visual: Box::new(visual),
        }
    };
    let visual = if !separate_shape_and_text
        && (outer_shadow.is_some()
            || inner_shadow.is_some()
            || glow.is_some()
            || reflection.is_some()
            || soft_edge.is_some()
            || three_d.is_some())
    {
        Visual::AdvancedEffect {
            outer_shadow,
            inner_shadow,
            glow,
            reflection,
            soft_edge,
            three_d: three_d.clone(),
            visual: Box::new(visual),
        }
    } else {
        visual
    };
    let visual = if !separate_shape_and_text && let Some(shadow) = shadow {
        Visual::Effect {
            shadow: Some(shadow),
            clip: None,
            visual: Box::new(visual),
        }
    } else {
        visual
    };
    let visual = if transform == AffineTransform::IDENTITY || separate_shape_and_text {
        visual
    } else {
        Visual::Layer {
            transform,
            opacity: 1.0,
            blend_mode: crate::model::BlendMode::Normal,
            visual: Box::new(visual),
        }
    };
    let visual = if separate_shape_and_text {
        let text_bounds = text_transform_bounds.unwrap_or(bounds);
        let shadow_text_only = fill == Paint::None && stroke == Paint::None;
        let mut shape_visual = Visual::PaintedShape {
            geometry: shape.geometry,
            fill,
            stroke,
            stroke_width,
        };
        if stroke_style != StrokeStyle::default() {
            shape_visual = Visual::StrokeStyle {
                style: stroke_style,
                visual: Box::new(shape_visual),
            };
        }
        if text_transform_bounds.is_none()
            && (outer_shadow.is_some()
                || inner_shadow.is_some()
                || glow.is_some()
                || reflection.is_some()
                || soft_edge.is_some()
                || three_d.is_some())
        {
            shape_visual = Visual::AdvancedEffect {
                outer_shadow: outer_shadow.clone(),
                inner_shadow: inner_shadow.clone(),
                glow: glow.clone(),
                reflection: reflection.clone(),
                soft_edge,
                three_d: three_d.clone(),
                visual: Box::new(shape_visual),
            };
        }
        if let Some(shadow) = shadow {
            shape_visual = Visual::Effect {
                shadow: Some(shadow),
                clip: None,
                visual: Box::new(shape_visual),
            };
        }
        if transform != AffineTransform::IDENTITY {
            shape_visual = Visual::Layer {
                transform,
                opacity: 1.0,
                blend_mode: crate::model::BlendMode::Normal,
                visual: Box::new(shape_visual),
            };
        }
        let text_transform = if text_transform_bounds.is_none() {
            shape_transform(bounds, shape.rotation_degrees, false, false)
        } else {
            AffineTransform::IDENTITY
        };
        let mut text_visual = if text_transform == AffineTransform::IDENTITY {
            visual
        } else {
            Visual::Layer {
                transform: text_transform,
                opacity: 1.0,
                blend_mode: crate::model::BlendMode::Normal,
                visual: Box::new(visual),
            }
        };
        if shadow_text_only && let Some(shadow) = shadow {
            text_visual = Visual::Effect {
                shadow: Some(shadow),
                clip: None,
                visual: Box::new(text_visual),
            };
        }
        Visual::Group {
            children: vec![
                VisualBrushChild {
                    bounds,
                    visual: shape_visual,
                },
                VisualBrushChild {
                    bounds: text_bounds,
                    visual: text_visual,
                },
            ],
        }
    } else {
        visual
    };
    let object = Object {
        numeric_id,
        parent_numeric_id,
        stable_id: format!("object:{numeric_id}"),
        parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
        kind: if has_text {
            ObjectKind::TextBox
        } else {
            ObjectKind::Shape
        },
        unit_index,
        bounds,
        z,
        text: has_text.then_some(text),
        source: SourceRef {
            part: part.to_owned(),
            mapping,
            locator: SourceLocator::PptxShape {
                shape_id,
                row: None,
                column: None,
                text_range: has_text.then_some((0, text_length)),
                metadata: shape.metadata,
            },
        },
        visual,
    };
    if let Some(replacement_id) = replacement_id {
        let index = usize::try_from(replacement_id)
            .map_err(|_| format_error(part, "placeholder object ID exceeds addressable range"))?;
        invalidate_cached_image_object(&mut state.image_cache, index);
        *state
            .objects
            .get_mut(index)
            .ok_or_else(|| format_error(part, "placeholder object mapping is missing"))? = object;
    } else {
        state.objects.push(object);
    }
    if let Some((width, height, children)) = math_layout {
        let x = math_region.x
            + match shape.align {
                TextAlign::Center => (math_region.width - width).max(0.0) / 2.0,
                TextAlign::End => (math_region.width - width).max(0.0),
                _ => 0.0,
            };
        let y = math_region.y
            + match vertical_align {
                TextVerticalAlign::Center => (math_region.height - height).max(0.0) / 2.0,
                TextVerticalAlign::Bottom => (math_region.height - height).max(0.0),
                _ => 0.0,
            };
        for (index, mut child) in children.into_iter().enumerate() {
            child.bounds.x += x;
            child.bounds.y += y;
            let text = super::docx::equation_part_text(&child.visual);
            let child_id = state.objects.len() as u32;
            let child_z = state.take_z();
            state.objects.push(Object {
                numeric_id: child_id,
                parent_numeric_id: Some(numeric_id),
                stable_id: format!("object:{numeric_id}:math:{index}"),
                parent_stable_id: Some(format!("object:{numeric_id}")),
                kind: if text.is_some() {
                    ObjectKind::TextBox
                } else {
                    ObjectKind::Shape
                },
                unit_index,
                bounds: child.bounds,
                z: child_z,
                text,
                source: SourceRef {
                    part: part.to_owned(),
                    mapping: MappingQuality::Derived,
                    locator: SourceLocator::PptxShape {
                        shape_id,
                        row: None,
                        column: None,
                        text_range: Some((0, text_length)),
                        metadata: PptxObjectMetadata::default(),
                    },
                },
                visual: child.visual,
            });
        }
    }
    register_placeholder_object(
        state,
        unit_index,
        shape.placeholder_index,
        shape.placeholder_type.as_deref(),
        numeric_id,
        Some(&shape.paragraph_styles),
        shape
            .preset
            .map(|preset| (preset, shape.preset_adjustments)),
    );
    if placeholder_prompt {
        state.pending_placeholder_objects.insert(numeric_id);
    } else {
        state.pending_placeholder_objects.remove(&numeric_id);
    }
    for (relationship_id, bounds) in picture_bullets {
        let mut picture = PictureState::new(shape_depth, Some(numeric_id));
        picture.explicit_bounds = Some(bounds);
        picture.shape_id = Some(shape_id);
        picture.embedded_relationship_id = Some(relationship_id);
        push_picture(
            picture,
            part,
            unit_index,
            relationships,
            package,
            state,
            content_types,
        )?;
    }
    Ok(())
}

fn hide_uninstantiated_prompt_objects(unit_index: u32, state: &mut PptxParseState) {
    let object_ids: Vec<u32> = state
        .pending_placeholder_objects
        .iter()
        .copied()
        .filter(|numeric_id| {
            usize::try_from(*numeric_id)
                .ok()
                .and_then(|index| state.objects.get(index))
                .is_some_and(|object| object.unit_index == unit_index)
        })
        .collect();
    for numeric_id in object_ids {
        if let Ok(index) = usize::try_from(numeric_id)
            && let Some(object) = state.objects.get_mut(index)
        {
            object.visual = Visual::None;
        }
        state.pending_placeholder_objects.remove(&numeric_id);
    }
}

fn inherit_placeholder_style(shape: &mut ShapeState, unit_index: u32, state: &PptxParseState) {
    let Some(numeric_id) = find_placeholder_object(
        state,
        unit_index,
        shape.placeholder_index,
        shape.placeholder_type.as_deref(),
    ) else {
        return;
    };
    // Geometry, fill and text must inherit from the same resolved placeholder.
    if let Some((preset, adjustments)) = state.placeholder_presets.get(&numeric_id) {
        shape.preset = Some(preset.clone());
        shape.preset_adjustments.clone_from(adjustments);
    }
    if let Some(styles) = state.placeholder_text_styles.get(&numeric_id) {
        shape.paragraph_styles.clone_from(styles);
    }
    let Some(visual) = usize::try_from(numeric_id)
        .ok()
        .and_then(|index| state.objects.get(index))
        .map(|object| &object.visual)
    else {
        return;
    };
    inherit_visual_style(shape, visual);
}

fn inherit_visual_style(shape: &mut ShapeState, visual: &Visual) {
    match visual {
        Visual::PaintedShape { fill, .. } => {
            shape.fill = fill.clone();
        }
        Visual::RichText {
            align, runs, fill, ..
        } => {
            shape.fill = fill.clone();
            shape.align = *align;
            if let Some(run) = runs.first() {
                shape.font_family.clone_from(&run.font_family);
                shape.font_size = run.font_size;
                shape.font_color = run.color;
                shape.bold = run.bold;
                shape.italic = run.italic;
                shape.underline = run.underline;
                shape.strikethrough = run.strikethrough;
                shape.highlight = run.highlight;
                shape.baseline_shift = run.baseline_shift;
                shape.letter_spacing = run.letter_spacing;
            }
        }
        Visual::Text {
            font_family,
            font_size,
            color,
            bold,
            italic,
            align,
            ..
        } => {
            shape.font_family.clone_from(font_family);
            shape.font_size = *font_size;
            shape.font_color = *color;
            shape.bold = *bold;
            shape.italic = *italic;
            shape.align = *align;
        }
        Visual::AdvancedEffect {
            three_d: Some(three_d),
            visual,
            ..
        } => {
            shape.three_d = Some(three_d.clone());
            inherit_visual_style(shape, visual);
        }
        Visual::Layer { visual, .. }
        | Visual::Effect { visual, .. }
        | Visual::TextLayout { visual, .. }
        | Visual::StrokeStyle { visual, .. }
        | Visual::AdvancedEffect { visual, .. } => inherit_visual_style(shape, visual),
        _ => {}
    }
}

fn visual_text_vertical_align(visual: &Visual) -> Option<TextVerticalAlign> {
    match visual {
        Visual::TextLayout { layout, .. } => Some(layout.vertical_align),
        Visual::Layer { visual, .. }
        | Visual::Effect { visual, .. }
        | Visual::StrokeStyle { visual, .. }
        | Visual::AdvancedEffect { visual, .. } => visual_text_vertical_align(visual),
        _ => None,
    }
}

fn placeholder_keys(placeholder_index: Option<u32>, placeholder_type: Option<&str>) -> Vec<String> {
    let mut keys = Vec::with_capacity(3);
    if let (Some(index), Some(kind)) = (placeholder_index, placeholder_type) {
        keys.push(format!("index:{index}:type:{kind}"));
    }
    if let Some(kind) = placeholder_type {
        keys.push(format!("type:{kind}"));
    }
    if let Some(index) = placeholder_index {
        keys.push(format!("index:{index}"));
    }
    keys
}

fn placeholder_lookup_keys(
    placeholder_index: Option<u32>,
    placeholder_type: Option<&str>,
) -> Vec<String> {
    let mut keys = placeholder_keys(placeholder_index, placeholder_type);
    if placeholder_index.is_none() {
        match placeholder_type {
            Some("title") => keys.push("type:ctrTitle".to_owned()),
            Some("ctrTitle") => keys.push("type:title".to_owned()),
            _ => {}
        }
    }
    keys
}

fn is_fixed_placeholder_type(kind: &str) -> bool {
    matches!(
        kind,
        "title" | "ctrTitle" | "subTitle" | "dt" | "ftr" | "hdr" | "sldNum"
    )
}

fn is_content_placeholder_type(kind: Option<&str>) -> bool {
    matches!(
        kind,
        Some(
            "title"
                | "ctrTitle"
                | "subTitle"
                | "body"
                | "obj"
                | "pic"
                | "chart"
                | "tbl"
                | "dgm"
                | "media"
                | "clipArt"
        )
    )
}

fn find_placeholder_object(
    state: &PptxParseState,
    unit_index: u32,
    placeholder_index: Option<u32>,
    placeholder_type: Option<&str>,
) -> Option<u32> {
    if let (Some(index), Some(kind)) = (placeholder_index, placeholder_type)
        && !is_fixed_placeholder_type(kind)
    {
        let exact = state
            .placeholders
            .get(&(unit_index, format!("index:{index}:type:{kind}")))
            .copied();
        return exact.or_else(|| {
            if kind != "obj" {
                return None;
            }
            state
                .placeholders
                .get(&(unit_index, format!("index:{index}")))
                .copied()
        });
    }
    placeholder_lookup_keys(placeholder_index, placeholder_type)
        .into_iter()
        .find_map(|key| state.placeholders.get(&(unit_index, key)).copied())
}

fn register_placeholder_object(
    state: &mut PptxParseState,
    unit_index: u32,
    placeholder_index: Option<u32>,
    placeholder_type: Option<&str>,
    numeric_id: u32,
    paragraph_styles: Option<&[ParagraphStyle]>,
    preset: Option<(String, HashMap<String, f32>)>,
) {
    if placeholder_index.is_some() || placeholder_type.is_some() {
        if let Some(preset) = preset {
            state.placeholder_presets.insert(numeric_id, preset);
        } else {
            state.placeholder_presets.remove(&numeric_id);
        }
    }
    for key in placeholder_keys(placeholder_index, placeholder_type) {
        state.placeholders.insert((unit_index, key), numeric_id);
    }
    if let Some(styles) = paragraph_styles {
        state
            .placeholder_text_styles
            .insert(numeric_id, styles.to_vec());
    }
}

fn shape_bounds(shape: &ShapeState) -> Rect {
    Rect {
        x: shape.x.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        y: shape.y.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        width: shape.width.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        height: shape.height.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
    }
}

// Placeholder paint and geometry must use the same final layout extent.
fn resolved_shape_bounds(shape: &ShapeState, unit_index: u32, state: &PptxParseState) -> Rect {
    let bounds = shape_bounds(shape);
    if (bounds.width == 0.0 || bounds.height == 0.0)
        && let Some(existing) = find_placeholder_object(
            state,
            unit_index,
            shape.placeholder_index,
            shape.placeholder_type.as_deref(),
        )
        .and_then(|id| usize::try_from(id).ok())
        .and_then(|index| state.objects.get(index))
    {
        return existing.bounds;
    }
    bounds
}

fn picture_shape_bounds(picture: &PictureState) -> Rect {
    picture.explicit_bounds.unwrap_or(Rect {
        x: picture.x.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        y: picture.y.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        width: picture.width.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
        height: picture.height.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL,
    })
}

pub(super) fn resolve_preset_geometry(
    preset: &str,
    bounds: Rect,
    adjustments: &HashMap<String, f32>,
    stroke_width: f32,
    head_arrow: Option<DrawingMlLineEnd>,
    tail_arrow: Option<DrawingMlLineEnd>,
) -> Option<(Geometry, Option<Rect>, Fidelity)> {
    if preset == "rect" {
        return Some((Geometry::Rectangle, None, Fidelity::Exact));
    }
    connector_geometry(
        preset,
        bounds,
        adjustments,
        stroke_width,
        head_arrow,
        tail_arrow,
    )
    .map(|(geometry, fidelity)| (geometry, None, fidelity))
    .or_else(|| {
        drawingml_preset_geometry(preset, bounds, adjustments)
            .map(|(geometry, text_rectangle)| (geometry, text_rectangle, Fidelity::Exact))
    })
    .or_else(|| {
        adjusted_preset_geometry(
            preset,
            bounds,
            adjustments,
            stroke_width,
            head_arrow,
            tail_arrow,
        )
        .map(|geometry| (geometry, None, Fidelity::Exact))
    })
    .or_else(|| {
        preset_geometry(preset, bounds).map(|(geometry, fidelity)| (geometry, None, fidelity))
    })
}

fn preset_geometry(preset: &str, bounds: Rect) -> Option<(Geometry, Fidelity)> {
    let width = bounds.width.max(0.0);
    let height = bounds.height.max(0.0);
    let exact = |geometry| Some((geometry, Fidelity::Exact));
    let approximate = |geometry| Some((geometry, Fidelity::Approximate));
    match preset {
        "rect" => exact(Geometry::Rectangle),
        "ellipse" | "flowChartConnector" => exact(Geometry::Ellipse),
        "roundRect" | "round1Rect" | "round2SameRect" | "round2DiagRect" => {
            exact(Geometry::RoundedRectangle {
                radius_x: width.min(height) * 0.12,
                radius_y: width.min(height) * 0.12,
            })
        }
        "flowChartAlternateProcess" => exact(Geometry::RoundedRectangle {
            radius_x: width.min(height) * 0.18,
            radius_y: width.min(height) * 0.18,
        }),
        "triangle" => exact(polygon_geometry(&[
            (width / 2.0, 0.0),
            (width, height),
            (0.0, height),
        ])),
        "flowChartMerge" => exact(polygon_geometry(&[
            (0.0, 0.0),
            (width, 0.0),
            (width / 2.0, height),
        ])),
        "rtTriangle" => exact(polygon_geometry(&[
            (0.0, 0.0),
            (width, height),
            (0.0, height),
        ])),
        "diamond" => exact(diamond_geometry(width, height)),
        "parallelogram" => exact(polygon_geometry(&[
            (width * 0.2, 0.0),
            (width, 0.0),
            (width * 0.8, height),
            (0.0, height),
        ])),
        "trapezoid" => exact(polygon_geometry(&[
            (width * 0.2, 0.0),
            (width * 0.8, 0.0),
            (width, height),
            (0.0, height),
        ])),
        "hexagon" => exact(polygon_geometry(&[
            (width * 0.25, 0.0),
            (width * 0.75, 0.0),
            (width, height / 2.0),
            (width * 0.75, height),
            (width * 0.25, height),
            (0.0, height / 2.0),
        ])),
        "octagon" => exact(regular_polygon_geometry(
            width,
            height,
            8,
            -std::f32::consts::FRAC_PI_8,
        )),
        "pentagon" => exact(polygon_geometry(&[
            (width / 2.0, 0.0),
            (width, height * 0.38),
            (width * 0.82, height),
            (width * 0.18, height),
            (0.0, height * 0.38),
        ])),
        "star4" => exact(star_geometry(width, height, 4, 0.28)),
        "star5" => exact(star_geometry(width, height, 5, 0.381_966)),
        "star6" => exact(star_geometry(width, height, 6, 0.5)),
        "star7" => exact(star_geometry(width, height, 7, 0.45)),
        "star8" => exact(star_geometry(width, height, 8, 0.42)),
        "star10" => exact(star_geometry(width, height, 10, 0.42)),
        "star12" => exact(star_geometry(width, height, 12, 0.42)),
        "star16" => exact(star_geometry(width, height, 16, 0.68)),
        "star24" => exact(star_geometry(width, height, 24, 0.78)),
        "star32" => exact(star_geometry(width, height, 32, 0.82)),
        "chevron" => exact(polygon_geometry(&[
            (0.0, 0.0),
            (width * 0.72, 0.0),
            (width, height / 2.0),
            (width * 0.72, height),
            (0.0, height),
            (width * 0.28, height / 2.0),
        ])),
        "upArrow" => exact(polygon_geometry(&[
            (width / 2.0, 0.0),
            (width, height * 0.38),
            (width * 0.68, height * 0.38),
            (width * 0.68, height),
            (width * 0.32, height),
            (width * 0.32, height * 0.38),
            (0.0, height * 0.38),
        ])),
        "downArrow" => exact(polygon_geometry(&[
            (width * 0.32, 0.0),
            (width * 0.68, 0.0),
            (width * 0.68, height * 0.62),
            (width, height * 0.62),
            (width / 2.0, height),
            (0.0, height * 0.62),
            (width * 0.32, height * 0.62),
        ])),
        "rightArrow" => exact(polygon_geometry(&[
            (0.0, height * 0.28),
            (width * 0.62, height * 0.28),
            (width * 0.62, 0.0),
            (width, height / 2.0),
            (width * 0.62, height),
            (width * 0.62, height * 0.72),
            (0.0, height * 0.72),
        ])),
        "leftArrow" => exact(polygon_geometry(&[
            (width, height * 0.28),
            (width * 0.38, height * 0.28),
            (width * 0.38, 0.0),
            (0.0, height / 2.0),
            (width * 0.38, height),
            (width * 0.38, height * 0.72),
            (width, height * 0.72),
        ])),
        "leftRightArrow" => exact(polygon_geometry(&[
            (0.0, height / 2.0),
            (width * 0.25, 0.0),
            (width * 0.25, height * 0.28),
            (width * 0.75, height * 0.28),
            (width * 0.75, 0.0),
            (width, height / 2.0),
            (width * 0.75, height),
            (width * 0.75, height * 0.72),
            (width * 0.25, height * 0.72),
            (width * 0.25, height),
        ])),
        "upDownArrow" => exact(polygon_geometry(&[
            (width / 2.0, 0.0),
            (width, height * 0.25),
            (width * 0.72, height * 0.25),
            (width * 0.72, height * 0.75),
            (width, height * 0.75),
            (width / 2.0, height),
            (0.0, height * 0.75),
            (width * 0.28, height * 0.75),
            (width * 0.28, height * 0.25),
            (0.0, height * 0.25),
        ])),
        "stripedRightArrow" => approximate(polygon_geometry(&[
            (width * 0.18, height * 0.28),
            (width * 0.62, height * 0.28),
            (width * 0.62, 0.0),
            (width, height / 2.0),
            (width * 0.62, height),
            (width * 0.62, height * 0.72),
            (width * 0.18, height * 0.72),
        ])),
        "homePlate" => exact(polygon_geometry(&[
            (0.0, 0.0),
            (width * 0.75, 0.0),
            (width, height / 2.0),
            (width * 0.75, height),
            (0.0, height),
        ])),
        "flowChartPredefinedProcess" => exact(predefined_process_geometry(width, height)),
        "flowChartMagneticDisk" => approximate(cylinder_geometry(width, height)),
        "borderCallout1" => approximate(Geometry::Rectangle),
        "cloudCallout" | "cloud" => approximate(cloud_geometry(width, height)),
        "smileyFace" => approximate(super::smiley_geometry(width, height)),
        "wedgeRectCallout" => exact(wedge_rect_callout_geometry(
            width, height, -20_833.0, 62_500.0,
        )),
        "arc" => exact(arc_geometry(bounds, 270.0, 360.0, 0.0, None, None)),
        "plus" => exact(super::cross_geometry(
            width,
            height,
            width / 3.0,
            width * 2.0 / 3.0,
            height / 3.0,
            height * 2.0 / 3.0,
        )),
        _ => None,
    }
}

fn predefined_process_geometry(width: f32, height: f32) -> Geometry {
    let inset = (width * 0.12).min(height * 0.3);
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands: vec![
            PathCommand::MoveTo { x: 0.0, y: 0.0 },
            PathCommand::LineTo { x: width, y: 0.0 },
            PathCommand::LineTo {
                x: width,
                y: height,
            },
            PathCommand::LineTo { x: 0.0, y: height },
            PathCommand::ClosePath,
            PathCommand::MoveTo { x: inset, y: 0.0 },
            PathCommand::LineTo {
                x: inset,
                y: height,
            },
            PathCommand::MoveTo {
                x: width - inset,
                y: 0.0,
            },
            PathCommand::LineTo {
                x: width - inset,
                y: height,
            },
        ],
    }
}

fn cylinder_geometry(width: f32, height: f32) -> Geometry {
    let cap = (height * 0.18).min(width * 0.25);
    let kappa = 0.552_284_8;
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands: vec![
            PathCommand::MoveTo { x: 0.0, y: cap },
            PathCommand::BezierCurveTo {
                cp1x: 0.0,
                cp1y: cap * (1.0 - kappa),
                cp2x: width * (0.5 - kappa / 2.0),
                cp2y: 0.0,
                x: width / 2.0,
                y: 0.0,
            },
            PathCommand::BezierCurveTo {
                cp1x: width * (0.5 + kappa / 2.0),
                cp1y: 0.0,
                cp2x: width,
                cp2y: cap * (1.0 - kappa),
                x: width,
                y: cap,
            },
            PathCommand::LineTo {
                x: width,
                y: height - cap,
            },
            PathCommand::BezierCurveTo {
                cp1x: width,
                cp1y: height - cap * (1.0 - kappa),
                cp2x: width * (0.5 + kappa / 2.0),
                cp2y: height,
                x: width / 2.0,
                y: height,
            },
            PathCommand::BezierCurveTo {
                cp1x: width * (0.5 - kappa / 2.0),
                cp1y: height,
                cp2x: 0.0,
                cp2y: height - cap * (1.0 - kappa),
                x: 0.0,
                y: height - cap,
            },
            PathCommand::ClosePath,
            PathCommand::MoveTo { x: 0.0, y: cap },
            PathCommand::BezierCurveTo {
                cp1x: 0.0,
                cp1y: cap * (1.0 + kappa),
                cp2x: width * (0.5 - kappa / 2.0),
                cp2y: cap * 2.0,
                x: width / 2.0,
                y: cap * 2.0,
            },
            PathCommand::BezierCurveTo {
                cp1x: width * (0.5 + kappa / 2.0),
                cp1y: cap * 2.0,
                cp2x: width,
                cp2y: cap * (1.0 + kappa),
                x: width,
                y: cap,
            },
        ],
    }
}

fn cloud_geometry(width: f32, height: f32) -> Geometry {
    let points = [
        (0.08, 0.62),
        (0.02, 0.47),
        (0.14, 0.34),
        (0.16, 0.16),
        (0.34, 0.10),
        (0.47, 0.02),
        (0.62, 0.12),
        (0.80, 0.10),
        (0.88, 0.27),
        (0.99, 0.42),
        (0.92, 0.59),
        (0.94, 0.78),
        (0.76, 0.86),
        (0.61, 0.98),
        (0.45, 0.89),
        (0.27, 0.96),
        (0.18, 0.80),
    ]
    .map(|(x, y)| (width * x, height * y));
    polygon_geometry(&points)
}

fn adjusted_preset_geometry(
    preset: &str,
    bounds: Rect,
    adjustments: &HashMap<String, f32>,
    stroke_width: f32,
    head_arrow: Option<DrawingMlLineEnd>,
    tail_arrow: Option<DrawingMlLineEnd>,
) -> Option<Geometry> {
    let adjustment = |name: &str, default: f32| adjustments.get(name).copied().unwrap_or(default);
    match preset {
        "arc" => Some(arc_geometry(
            bounds,
            adjustment("adj1", 16_200_000.0) / 60_000.0,
            adjustment("adj2", 0.0) / 60_000.0,
            stroke_width,
            head_arrow,
            tail_arrow,
        )),
        "wedgeRectCallout" => Some(wedge_rect_callout_geometry(
            bounds.width.max(0.0),
            bounds.height.max(0.0),
            adjustment("adj1", -20_833.0),
            adjustment("adj2", 62_500.0),
        )),
        _ => None,
    }
}

fn wedge_rect_callout_geometry(
    width: f32,
    height: f32,
    horizontal_adjustment: f32,
    vertical_adjustment: f32,
) -> Geometry {
    let tip_x = width * (0.5 + horizontal_adjustment / 100_000.0);
    let tip_y = height * (0.5 + vertical_adjustment / 100_000.0);
    let half_base = width.min(height) * 0.1;
    let horizontal_offset = if width > 0.0 {
        (tip_x - width / 2.0) / width
    } else {
        0.0
    };
    let vertical_offset = if height > 0.0 {
        (tip_y - height / 2.0) / height
    } else {
        0.0
    };
    let edge = if horizontal_offset.abs() >= vertical_offset.abs() {
        if horizontal_offset < 0.0 { 3 } else { 1 }
    } else if vertical_offset < 0.0 {
        0
    } else {
        2
    };
    let mut points = vec![(0.0, 0.0)];
    match edge {
        0 => points.extend([
            ((tip_x - half_base).clamp(0.0, width), 0.0),
            (tip_x, tip_y),
            ((tip_x + half_base).clamp(0.0, width), 0.0),
            (width, 0.0),
            (width, height),
            (0.0, height),
        ]),
        1 => points.extend([
            (width, 0.0),
            (width, (tip_y - half_base).clamp(0.0, height)),
            (tip_x, tip_y),
            (width, (tip_y + half_base).clamp(0.0, height)),
            (width, height),
            (0.0, height),
        ]),
        2 => points.extend([
            (width, 0.0),
            (width, height),
            ((tip_x + half_base).clamp(0.0, width), height),
            (tip_x, tip_y),
            ((tip_x - half_base).clamp(0.0, width), height),
            (0.0, height),
        ]),
        _ => points.extend([
            (width, 0.0),
            (width, height),
            (0.0, height),
            (0.0, (tip_y + half_base).clamp(0.0, height)),
            (tip_x, tip_y),
            (0.0, (tip_y - half_base).clamp(0.0, height)),
        ]),
    }
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands: points
            .into_iter()
            .enumerate()
            .map(|(index, (x, y))| {
                if index == 0 {
                    PathCommand::MoveTo { x, y }
                } else {
                    PathCommand::LineTo { x, y }
                }
            })
            .chain(std::iter::once(PathCommand::ClosePath))
            .collect(),
    }
}

fn arc_geometry(
    bounds: Rect,
    start_degrees: f32,
    end_degrees: f32,
    stroke_width: f32,
    head_arrow: Option<DrawingMlLineEnd>,
    tail_arrow: Option<DrawingMlLineEnd>,
) -> Geometry {
    let center = (bounds.width.max(0.0) / 2.0, bounds.height.max(0.0) / 2.0);
    let radii = center;
    let start = start_degrees.to_radians();
    let mut sweep = (end_degrees - start_degrees).rem_euclid(360.0).to_radians();
    if sweep <= f32::EPSILON {
        sweep = std::f32::consts::TAU;
    }
    let segment_count = (sweep.abs() / std::f32::consts::FRAC_PI_2).ceil().max(1.0) as usize;
    let segment_sweep = sweep / segment_count as f32;
    let point = |angle: f32| {
        (
            center.0 + radii.0 * angle.cos(),
            center.1 + radii.1 * angle.sin(),
        )
    };
    let start_point = point(start);
    let mut commands = vec![PathCommand::MoveTo {
        x: start_point.0,
        y: start_point.1,
    }];
    let mut first_control = start_point;
    let mut last_control = start_point;
    let mut end_point = start_point;
    for index in 0..segment_count {
        let from = start + segment_sweep * index as f32;
        let to = from + segment_sweep;
        let [control_1, control_2, to_point] = ellipse_arc_bezier_points(center, radii, from, to);
        if index == 0 {
            first_control = control_1;
        }
        last_control = control_2;
        end_point = to_point;
        commands.push(PathCommand::BezierCurveTo {
            cp1x: control_1.0,
            cp1y: control_1.1,
            cp2x: control_2.0,
            cp2y: control_2.1,
            x: to_point.0,
            y: to_point.1,
        });
    }
    if let Some(line_end) = head_arrow {
        append_line_end(
            &mut commands,
            start_point,
            first_control,
            stroke_width,
            line_end,
        );
    }
    if let Some(line_end) = tail_arrow {
        append_line_end(
            &mut commands,
            end_point,
            last_control,
            stroke_width,
            line_end,
        );
    }
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    }
}

fn connector_geometry(
    preset: &str,
    bounds: Rect,
    adjustments: &HashMap<String, f32>,
    stroke_width: f32,
    head_arrow: Option<DrawingMlLineEnd>,
    tail_arrow: Option<DrawingMlLineEnd>,
) -> Option<(Geometry, Fidelity)> {
    if matches!(
        preset,
        "curvedConnector2" | "curvedConnector3" | "curvedConnector4" | "curvedConnector5"
    ) {
        return curved_connector_geometry(
            preset,
            bounds,
            adjustments,
            stroke_width,
            head_arrow,
            tail_arrow,
        )
        .map(|geometry| (geometry, Fidelity::Exact));
    }
    let points = match preset {
        "line" | "straightConnector1" => vec![(0.0, 0.0), (bounds.width, bounds.height)],
        "bentConnector2" => vec![
            (0.0, 0.0),
            (bounds.width, 0.0),
            (bounds.width, bounds.height),
        ],
        "bentConnector3" => {
            let bend_x =
                bounds.width * adjustments.get("adj1").copied().unwrap_or(50_000.0) / 100_000.0;
            vec![
                (0.0, 0.0),
                (bend_x, 0.0),
                (bend_x, bounds.height),
                (bounds.width, bounds.height),
            ]
        }
        "bentConnector4" => {
            let bend_x =
                bounds.width * adjustments.get("adj1").copied().unwrap_or(50_000.0) / 100_000.0;
            let bend_y =
                bounds.height * adjustments.get("adj2").copied().unwrap_or(50_000.0) / 100_000.0;
            vec![
                (0.0, 0.0),
                (bend_x, 0.0),
                (bend_x, bend_y),
                (bounds.width, bend_y),
                (bounds.width, bounds.height),
            ]
        }
        "bentConnector5" => {
            let first_x =
                bounds.width * adjustments.get("adj1").copied().unwrap_or(25_000.0) / 100_000.0;
            let bend_y =
                bounds.height * adjustments.get("adj2").copied().unwrap_or(50_000.0) / 100_000.0;
            let second_x =
                bounds.width * adjustments.get("adj3").copied().unwrap_or(75_000.0) / 100_000.0;
            vec![
                (0.0, 0.0),
                (first_x, 0.0),
                (first_x, bend_y),
                (second_x, bend_y),
                (second_x, bounds.height),
                (bounds.width, bounds.height),
            ]
        }
        _ => return None,
    };
    let mut commands = vec![PathCommand::MoveTo {
        x: points[0].0,
        y: points[0].1,
    }];
    commands.extend(
        points
            .iter()
            .skip(1)
            .map(|&(x, y)| PathCommand::LineTo { x, y }),
    );
    if let Some(line_end) = head_arrow {
        append_line_end(&mut commands, points[0], points[1], stroke_width, line_end);
    }
    if let Some(line_end) = tail_arrow {
        let last = points.len() - 1;
        append_line_end(
            &mut commands,
            points[last],
            points[last - 1],
            stroke_width,
            line_end,
        );
    }
    Some((
        Geometry::Path {
            fill_rule: FillRule::NonZero,
            commands,
        },
        Fidelity::Exact,
    ))
}

pub(super) fn drawingml_connector_geometry(
    preset: &str,
    bounds: Rect,
    stroke_width: f32,
    head_arrow: Option<DrawingMlLineEnd>,
    tail_arrow: Option<DrawingMlLineEnd>,
    flip_horizontal: bool,
    flip_vertical: bool,
    rotation_degrees: f32,
) -> Option<Geometry> {
    let (mut geometry, _) = connector_geometry(
        preset,
        bounds,
        &HashMap::new(),
        stroke_width,
        head_arrow,
        tail_arrow,
    )?;
    if (flip_horizontal || flip_vertical || rotation_degrees != 0.0)
        && let Geometry::Path { commands, .. } = &mut geometry
    {
        for command in commands {
            let flip = |x: &mut f32, y: &mut f32| {
                if flip_horizontal {
                    *x = bounds.width - *x;
                }
                if flip_vertical {
                    *y = bounds.height - *y;
                }
                let rotation = AffineTransform::rotation_about(
                    bounds.width / 2.0,
                    bounds.height / 2.0,
                    rotation_degrees,
                );
                (*x, *y) = (
                    rotation.a * *x + rotation.c * *y + rotation.e,
                    rotation.b * *x + rotation.d * *y + rotation.f,
                );
            };
            match command {
                PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => flip(x, y),
                PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => {
                    flip(cpx, cpy);
                    flip(x, y);
                }
                PathCommand::BezierCurveTo {
                    cp1x,
                    cp1y,
                    cp2x,
                    cp2y,
                    x,
                    y,
                } => {
                    flip(cp1x, cp1y);
                    flip(cp2x, cp2y);
                    flip(x, y);
                }
                PathCommand::ClosePath => {}
            }
        }
    }
    Some(geometry)
}

fn curved_connector_geometry(
    preset: &str,
    bounds: Rect,
    adjustments: &HashMap<String, f32>,
    stroke_width: f32,
    head_arrow: Option<DrawingMlLineEnd>,
    tail_arrow: Option<DrawingMlLineEnd>,
) -> Option<Geometry> {
    let fraction = |name: &str| adjustments.get(name).copied().unwrap_or(50_000.0) / 100_000.0;
    let width = bounds.width;
    let height = bounds.height;
    let start = (0.0, 0.0);
    let end = (width, height);
    let segments = match preset {
        "curvedConnector2" => vec![((width / 2.0, 0.0), (width, height / 2.0), end)],
        "curvedConnector3" => {
            let x2 = width * fraction("adj1");
            vec![
                ((x2 / 2.0, 0.0), (x2, height / 4.0), (x2, height / 2.0)),
                ((x2, height * 3.0 / 4.0), ((width + x2) / 2.0, height), end),
            ]
        }
        "curvedConnector4" => {
            let x2 = width * fraction("adj1");
            let x3 = (width + x2) / 2.0;
            let y4 = height * fraction("adj2");
            let y1 = y4 / 2.0;
            vec![
                ((x2 / 2.0, 0.0), (x2, y1 / 2.0), (x2, y1)),
                ((x2, (y1 + y4) / 2.0), ((x2 + x3) / 2.0, y4), (x3, y4)),
                (((x3 + width) / 2.0, y4), (width, (height + y4) / 2.0), end),
            ]
        }
        "curvedConnector5" => {
            let x3 = width * fraction("adj1");
            let x6 = width * fraction("adj3");
            let x1 = (x3 + x6) / 2.0;
            let y4 = height * fraction("adj2");
            let y1 = y4 / 2.0;
            let y5 = (height + y4) / 2.0;
            vec![
                ((x3 / 2.0, 0.0), (x3, y1 / 2.0), (x3, y1)),
                ((x3, (y1 + y4) / 2.0), ((x3 + x1) / 2.0, y4), (x1, y4)),
                (((x6 + x1) / 2.0, y4), (x6, (y5 + y4) / 2.0), (x6, y5)),
                ((x6, (y5 + height) / 2.0), ((x6 + width) / 2.0, height), end),
            ]
        }
        _ => return None,
    };
    let first_control = segments.first()?.0;
    let last_control = segments.last()?.1;
    let mut commands = vec![PathCommand::MoveTo {
        x: start.0,
        y: start.1,
    }];
    commands.extend(
        segments
            .into_iter()
            .map(
                |(first_control, second_control, endpoint)| PathCommand::BezierCurveTo {
                    cp1x: first_control.0,
                    cp1y: first_control.1,
                    cp2x: second_control.0,
                    cp2y: second_control.1,
                    x: endpoint.0,
                    y: endpoint.1,
                },
            ),
    );
    if let Some(line_end) = head_arrow {
        append_line_end(&mut commands, start, first_control, stroke_width, line_end);
    }
    if let Some(line_end) = tail_arrow {
        append_line_end(&mut commands, end, last_control, stroke_width, line_end);
    }
    Some(Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands,
    })
}

fn append_line_end(
    commands: &mut Vec<PathCommand>,
    tip: (f32, f32),
    adjacent: (f32, f32),
    stroke_width: f32,
    line_end: DrawingMlLineEnd,
) {
    let delta_x = adjacent.0 - tip.0;
    let delta_y = adjacent.1 - tip.1;
    let length = delta_x.hypot(delta_y);
    if length <= f32::EPSILON || stroke_width <= f32::EPSILON {
        return;
    }
    let direction_x = delta_x / length;
    let direction_y = delta_y / length;
    // PowerPoint keeps line ends legible on hairlines by sizing them from at
    // least a 2 pt pen. Above that floor the documented sm/med/lg stroke
    // multipliers remain proportional to the authored line width. In our CSS
    // pixel coordinate space 2 pt is 8/3 px.
    let line_end_stroke_width = stroke_width.max(8.0 / 3.0);
    let arrow_length = line_end_stroke_width * line_end.length.stroke_multiplier();
    let half_width = line_end_stroke_width * line_end.width.stroke_multiplier() / 2.0;
    let base_x = tip.0 + direction_x * arrow_length;
    let base_y = tip.1 + direction_y * arrow_length;
    let perpendicular_x = -direction_y * half_width;
    let perpendicular_y = direction_x * half_width;
    let left = (base_x + perpendicular_x, base_y + perpendicular_y);
    let right = (base_x - perpendicular_x, base_y - perpendicular_y);
    match line_end.kind {
        DrawingMlLineEndKind::Arrow => commands.extend([
            PathCommand::MoveTo {
                x: left.0,
                y: left.1,
            },
            PathCommand::LineTo { x: tip.0, y: tip.1 },
            PathCommand::LineTo {
                x: right.0,
                y: right.1,
            },
        ]),
        DrawingMlLineEndKind::Stealth => {
            let notch = (
                tip.0 + direction_x * arrow_length * 0.6,
                tip.1 + direction_y * arrow_length * 0.6,
            );
            commands.extend([
                PathCommand::MoveTo { x: tip.0, y: tip.1 },
                PathCommand::LineTo {
                    x: left.0,
                    y: left.1,
                },
                PathCommand::LineTo {
                    x: notch.0,
                    y: notch.1,
                },
                PathCommand::LineTo {
                    x: right.0,
                    y: right.1,
                },
                PathCommand::ClosePath,
            ]);
        }
        DrawingMlLineEndKind::Triangle => commands.extend([
            PathCommand::MoveTo { x: tip.0, y: tip.1 },
            PathCommand::LineTo {
                x: left.0,
                y: left.1,
            },
            PathCommand::LineTo {
                x: right.0,
                y: right.1,
            },
            PathCommand::ClosePath,
        ]),
        DrawingMlLineEndKind::Diamond => {
            let center = (
                tip.0 + direction_x * arrow_length / 2.0,
                tip.1 + direction_y * arrow_length / 2.0,
            );
            commands.extend([
                PathCommand::MoveTo { x: tip.0, y: tip.1 },
                PathCommand::LineTo {
                    x: center.0 + perpendicular_x,
                    y: center.1 + perpendicular_y,
                },
                PathCommand::LineTo {
                    x: tip.0 + direction_x * arrow_length,
                    y: tip.1 + direction_y * arrow_length,
                },
                PathCommand::LineTo {
                    x: center.0 - perpendicular_x,
                    y: center.1 - perpendicular_y,
                },
                PathCommand::ClosePath,
            ]);
        }
        DrawingMlLineEndKind::Oval => {
            let center_x = tip.0 + direction_x * arrow_length / 2.0;
            let center_y = tip.1 + direction_y * arrow_length / 2.0;
            let radius_x = direction_x * arrow_length / 2.0;
            let radius_y = direction_y * arrow_length / 2.0;
            let side_x = -direction_y * half_width;
            let side_y = direction_x * half_width;
            let kappa = 0.552_284_8;
            commands.extend([
                PathCommand::MoveTo {
                    x: center_x - radius_x,
                    y: center_y - radius_y,
                },
                PathCommand::BezierCurveTo {
                    cp1x: center_x - radius_x + side_x * kappa,
                    cp1y: center_y - radius_y + side_y * kappa,
                    cp2x: center_x + side_x - radius_x * kappa,
                    cp2y: center_y + side_y - radius_y * kappa,
                    x: center_x + side_x,
                    y: center_y + side_y,
                },
                PathCommand::BezierCurveTo {
                    cp1x: center_x + side_x + radius_x * kappa,
                    cp1y: center_y + side_y + radius_y * kappa,
                    cp2x: center_x + radius_x + side_x * kappa,
                    cp2y: center_y + radius_y + side_y * kappa,
                    x: center_x + radius_x,
                    y: center_y + radius_y,
                },
                PathCommand::BezierCurveTo {
                    cp1x: center_x + radius_x - side_x * kappa,
                    cp1y: center_y + radius_y - side_y * kappa,
                    cp2x: center_x - side_x + radius_x * kappa,
                    cp2y: center_y - side_y + radius_y * kappa,
                    x: center_x - side_x,
                    y: center_y - side_y,
                },
                PathCommand::BezierCurveTo {
                    cp1x: center_x - side_x - radius_x * kappa,
                    cp1y: center_y - side_y - radius_y * kappa,
                    cp2x: center_x - radius_x - side_x * kappa,
                    cp2y: center_y - radius_y - side_y * kappa,
                    x: center_x - radius_x,
                    y: center_y - radius_y,
                },
                PathCommand::ClosePath,
            ]);
        }
    }
}

fn relationship_id(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<Option<String>, Diagnostic> {
    attributes
        .iter()
        .find(|attribute| attribute.name.contains(':') && local_name(attribute.name) == "id")
        .map(|attribute| {
            decode_xml_text(attribute.value)
                .map(|value| value.into_owned())
                .map_err(|error| with_part(error, part))
        })
        .transpose()
}

fn color_transform_percentage(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<f32, Diagnostic> {
    Ok(optional_percentage_attribute(attributes, "val", part)?.unwrap_or(1.0))
}

fn optional_percentage_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<f32>, Diagnostic> {
    string_attribute(attributes, name, part)?
        .map(|value| {
            parse_percentage(&value)
                .map_err(|_| format_error(part, format!("attribute {name} is not a percentage")))
        })
        .transpose()
}

fn parse_percentage(value: &str) -> Result<f32, ()> {
    if let Some(value) = value.strip_suffix('%') {
        value
            .parse::<f32>()
            .map_err(|_| ())
            .and_then(|value| value.is_finite().then_some(value).ok_or(()))
            .map(|value| value / 100.0)
    } else {
        value
            .parse::<i64>()
            .map(|value| value as f32 / 100_000.0)
            .map_err(|_| ())
    }
}

fn boolean_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<bool>, Diagnostic> {
    let Some(value) = string_attribute(attributes, name, part)? else {
        return Ok(None);
    };
    match value.as_str() {
        "1" | "true" | "on" => Ok(Some(true)),
        "0" | "false" | "off" => Ok(Some(false)),
        _ => Err(format_error(
            part,
            format!("attribute {name} is not a boolean"),
        )),
    }
}

fn drawingml_preset_color(value: &str) -> Option<u32> {
    let rgb = match value {
        "aliceBlue" => 0xf0f8ff,
        "antiqueWhite" => 0xfaebd7,
        "aqua" | "cyan" => 0x00ffff,
        "aquamarine" => 0x7fffd4,
        "azure" => 0xf0ffff,
        "beige" => 0xf5f5dc,
        "bisque" => 0xffe4c4,
        "black" => 0x000000,
        "blanchedAlmond" => 0xffebcd,
        "blue" => 0x0000ff,
        "blueViolet" => 0x8a2be2,
        "brown" => 0xa52a2a,
        "burlyWood" => 0xdeb887,
        "cadetBlue" => 0x5f9ea0,
        "chartreuse" => 0x7fff00,
        "chocolate" => 0xd2691e,
        "coral" => 0xff7f50,
        "cornflowerBlue" => 0x6495ed,
        "cornsilk" => 0xfff8dc,
        "crimson" => 0xdc143c,
        "dkBlue" => 0x00008b,
        "dkCyan" => 0x008b8b,
        "dkGoldenrod" => 0xb8860b,
        "dkGray" | "dkGrey" => 0xa9a9a9,
        "dkGreen" => 0x006400,
        "dkKhaki" => 0xbdb76b,
        "dkMagenta" => 0x8b008b,
        "dkOliveGreen" => 0x556b2f,
        "dkOrange" => 0xff8c00,
        "dkOrchid" => 0x9932cc,
        "dkRed" => 0x8b0000,
        "dkSalmon" => 0xe9967a,
        "dkSeaGreen" => 0x8fbc8f,
        "dkSlateBlue" => 0x483d8b,
        "dkSlateGray" | "dkSlateGrey" => 0x2f4f4f,
        "dkTurquoise" => 0x00ced1,
        "dkViolet" => 0x9400d3,
        "deepPink" => 0xff1493,
        "deepSkyBlue" => 0x00bfff,
        "dimGray" | "dimGrey" => 0x696969,
        "dodgerBlue" => 0x1e90ff,
        "firebrick" => 0xb22222,
        "floralWhite" => 0xfffaf0,
        "forestGreen" => 0x228b22,
        "fuchsia" | "magenta" => 0xff00ff,
        "gainsboro" => 0xdcdcdc,
        "ghostWhite" => 0xf8f8ff,
        "gold" => 0xffd700,
        "goldenrod" => 0xdaa520,
        "gray" | "grey" => 0x808080,
        "green" => 0x008000,
        "greenYellow" => 0xadff2f,
        "honeydew" => 0xf0fff0,
        "hotPink" => 0xff69b4,
        "indianRed" => 0xcd5c5c,
        "indigo" => 0x4b0082,
        "ivory" => 0xfffff0,
        "khaki" => 0xf0e68c,
        "lavender" => 0xe6e6fa,
        "lavenderBlush" => 0xfff0f5,
        "lawnGreen" => 0x7cfc00,
        "lemonChiffon" => 0xfffacd,
        "ltBlue" => 0xadd8e6,
        "ltCoral" => 0xf08080,
        "ltCyan" => 0xe0ffff,
        "ltGoldenrodYellow" => 0xfafad2,
        "ltGray" | "ltGrey" => 0xd3d3d3,
        "ltGreen" => 0x90ee90,
        "ltPink" => 0xffb6c1,
        "ltSalmon" => 0xffa07a,
        "ltSeaGreen" => 0x20b2aa,
        "ltSkyBlue" => 0x87cefa,
        "ltSlateGray" | "ltSlateGrey" => 0x778899,
        "ltSteelBlue" => 0xb0c4de,
        "ltYellow" => 0xffffe0,
        "lime" => 0x00ff00,
        "limeGreen" => 0x32cd32,
        "linen" => 0xfaf0e6,
        "maroon" => 0x800000,
        "medAquamarine" => 0x66cdaa,
        "medBlue" => 0x0000cd,
        "medOrchid" => 0xba55d3,
        "medPurple" => 0x9370db,
        "medSeaGreen" => 0x3cb371,
        "medSlateBlue" => 0x7b68ee,
        "medSpringGreen" => 0x00fa9a,
        "medTurquoise" => 0x48d1cc,
        "medVioletRed" => 0xc71585,
        "midnightBlue" => 0x191970,
        "mintCream" => 0xf5fffa,
        "mistyRose" => 0xffe4e1,
        "moccasin" => 0xffe4b5,
        "navajoWhite" => 0xffdead,
        "navy" => 0x000080,
        "oldLace" => 0xfdf5e6,
        "olive" => 0x808000,
        "oliveDrab" => 0x6b8e23,
        "orange" => 0xffa500,
        "orangeRed" => 0xff4500,
        "orchid" => 0xda70d6,
        "paleGoldenrod" => 0xeee8aa,
        "paleGreen" => 0x98fb98,
        "paleTurquoise" => 0xafeeee,
        "paleVioletRed" => 0xdb7093,
        "papayaWhip" => 0xffefd5,
        "peachPuff" => 0xffdab9,
        "peru" => 0xcd853f,
        "pink" => 0xffc0cb,
        "plum" => 0xdda0dd,
        "powderBlue" => 0xb0e0e6,
        "purple" => 0x800080,
        "red" => 0xff0000,
        "rosyBrown" => 0xbc8f8f,
        "royalBlue" => 0x4169e1,
        "saddleBrown" => 0x8b4513,
        "salmon" => 0xfa8072,
        "sandyBrown" => 0xf4a460,
        "seaGreen" => 0x2e8b57,
        "seaShell" => 0xfff5ee,
        "sienna" => 0xa0522d,
        "silver" => 0xc0c0c0,
        "skyBlue" => 0x87ceeb,
        "slateBlue" => 0x6a5acd,
        "slateGray" | "slateGrey" => 0x708090,
        "snow" => 0xfffafa,
        "springGreen" => 0x00ff7f,
        "steelBlue" => 0x4682b4,
        "tan" => 0xd2b48c,
        "teal" => 0x008080,
        "thistle" => 0xd8bfd8,
        "tomato" => 0xff6347,
        "turquoise" => 0x40e0d0,
        "violet" => 0xee82ee,
        "wheat" => 0xf5deb3,
        "white" => 0xffffff,
        "whiteSmoke" => 0xf5f5f5,
        "yellow" => 0xffff00,
        "yellowGreen" => 0x9acd32,
        _ => return None,
    };
    Some((rgb << 8) | 0xff)
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
        "phClr" => 0x4472c4,
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
mod tests {
    use std::collections::HashMap;

    use super::{
        DrawingMlDashPattern, DrawingMlLineEnd, DrawingMlLineEndKind, DrawingMlLineEndSize,
        PptxTheme, ShapeState, SlideParseContext, TableBorderSide, TableCellStyleContext,
        TableStyleBorderSide, TextCapitalization, TextRunState, append_line_end,
        apply_master_text_styles, begin_shape_paragraph, chart_category_labels_overlap,
        chart_text_style, connector_geometry, curved_connector_geometry,
        drawingml_latin_line_break_tokens, drawingml_preset_geometry, drawingml_text_tokens,
        format_auto_number, parse_master_text_styles, parse_pptx_table_style, parse_pptx_theme,
        parse_slide, scale_embedded_text_visual, shape_transform,
    };
    use crate::diagnostic::{DiagnosticCode, Fidelity, Phase};
    use crate::font_metrics::FontMetricTable;
    use crate::format::presentation_image::stored_zip;
    use crate::limits::Limits;
    use crate::model::{
        Geometry, LineJoin, MappingQuality, MediaKind, ObjectKind, Paint, PathCommand,
        PathFillMode, Rect, SourceLocator, StrokeStyle, TextLayout, TextOrientation, TextRun,
        Visual,
    };
    use crate::package::Package;

    #[test]
    fn real_vml_linked_object_retains_its_drawingml_metafile() {
        let document = crate::format::detect_and_parse(
            include_bytes!("../../tests/fixtures/mce-vml-linked-emf.pptx"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        assert!(
            document
                .objects
                .iter()
                .any(|object| { format!("{:?}", object.visual).contains("image/x-emf") }),
            "the supported DrawingML fallback must retain the linked object image"
        );
    }

    #[test]
    fn real_ink2_renders_all_vml_handwriting() {
        let bytes = include_bytes!("../../tests/fixtures/Ink2.pptx");
        let document = crate::format::detect_and_parse(bytes, Limits::default())
            .unwrap()
            .unwrap();
        let ink: Vec<_> = document
            .objects
            .iter()
            .filter(|o| o.source.part == "ppt/drawings/vmlDrawing1.vml")
            .collect();
        assert_eq!(
            ink.len(),
            19,
            "all pen and highlighter shapes must render: {:?}",
            document.diagnostics
        );
        let mut colors = std::collections::BTreeMap::new();
        for object in &ink {
            let Visual::StrokeStyle { style, visual } = &object.visual else {
                panic!("ink needs its pen cap")
            };
            let Visual::Shape {
                geometry: Geometry::Path { commands, .. },
                fill,
                stroke: color,
                stroke_width,
                ..
            } = visual.as_ref()
            else {
                panic!("ink must be editable vector geometry, not a preview")
            };
            assert_eq!(*fill, 0);
            assert!(
                commands
                    .iter()
                    .any(|c| matches!(c, PathCommand::BezierCurveTo { .. }))
            );
            *colors.entry(*color).or_insert(0) += 1;
            if color >> 8 == 0xffff00 {
                assert_eq!(style.cap, crate::model::LineCap::Square);
                assert!((*stroke_width - 40.0 / 3.0).abs() < 0.001);
            } else {
                assert_eq!(style.cap, crate::model::LineCap::Round);
            }
        }
        assert_eq!(
            colors,
            [(0x7030a0ff, 2), (0xff0000ff, 3), (0xffff0055, 14)].into()
        );
        let first = ink[0];
        assert!((first.bounds.x - 38.375 * 4.0 / 3.0).abs() < 0.001);
        let Visual::StrokeStyle { visual, .. } = &first.visual else {
            unreachable!()
        };
        let Visual::Shape {
            geometry: Geometry::Path { commands, .. },
            ..
        } = visual.as_ref()
        else {
            unreachable!()
        };
        assert_eq!(
            commands
                .iter()
                .filter(|c| matches!(c, PathCommand::MoveTo { .. }))
                .count(),
            13,
            "do not stop at the first VML end command"
        );
        let PathCommand::MoveTo { x, y } = commands[0] else {
            unreachable!()
        };
        assert!((x - 16.0 / 5926.0 * first.bounds.width).abs() < 0.001);
        assert!((y - 1093.0 / 2514.0 * first.bounds.height).abs() < 0.001);
    }

    #[test]
    fn real_ink2_isolates_bad_paths_and_enforces_object_limits() {
        let original = include_bytes!("../../tests/fixtures/Ink2.pptx");
        let package = Package::open(original, Limits::default()).unwrap();
        let entries: Vec<_> = package
            .entry_names()
            .map(|name| {
                let mut data = package.required_part(name).unwrap().to_vec();
                if name == "ppt/drawings/vmlDrawing1.vml" {
                    data = String::from_utf8(data)
                        .unwrap()
                        .replacen("m1368,2468", "q1368,2468", 1)
                        .into_bytes();
                }
                (name.to_owned(), data)
            })
            .collect();
        let bytes = stored_zip(
            &entries
                .iter()
                .map(|(name, bytes)| (name.as_str(), bytes.as_slice()))
                .collect::<Vec<_>>(),
        );
        let document = crate::format::detect_and_parse(&bytes, Limits::default())
            .unwrap()
            .unwrap();
        assert_eq!(
            document
                .objects
                .iter()
                .filter(|o| o.source.part == "ppt/drawings/vmlDrawing1.vml")
                .count(),
            18
        );
        assert!(
            document
                .diagnostics
                .iter()
                .any(|d| d.code == DiagnosticCode::UnsupportedFeature
                    && d.location.part.as_deref() == Some("ppt/drawings/vmlDrawing1.vml"))
        );
        let error = crate::format::detect_and_parse(
            original,
            Limits {
                max_document_objects: 10,
                ..Limits::default()
            },
        )
        .unwrap_err();
        assert_eq!(error.code, DiagnosticCode::ObjectLimit);
    }

    const SLIDE_PART: &str = "ppt/slides/slide1.xml";
    const PNG: &[u8] = b"\x89PNG\r\n\x1a\nimage";

    #[test]
    fn next_twenty_real_chart_axis_fonts_and_bounds() {
        let mut failures = Vec::new();
        for entry in std::fs::read_dir("tests/fixtures/pptx-next20").unwrap() {
            let path = entry.unwrap().path();
            let bytes = std::fs::read(&path).unwrap();
            let package = crate::package::Package::open(&bytes, Limits::default()).unwrap();
            let part = package
                .entry_names()
                .filter(|n| n.starts_with("ppt/charts/chart") && n.ends_with(".xml"))
                .min()
                .unwrap();
            let chart = crate::format::drawingml::parse_chart(&package, part, |_| None)
                .unwrap()
                .unwrap();
            let document = crate::format::detect_and_parse(&bytes, Limits::default())
                .unwrap()
                .unwrap();
            if document
                .objects
                .iter()
                .any(|o| o.bounds.width < 0.0 || o.bounds.height < 0.0)
            {
                failures.push(format!("{}: invalid bounds", path.display()));
            }
            let axis = if chart.series.iter().any(|s| s.bar_horizontal) {
                &chart.value_axis_options
            } else {
                &chart.horizontal_axis_options
            };
            let Some(expected) = axis.label_font_size.or(chart.font_size) else {
                continue;
            };
            if chart.series.iter().all(|s| {
                matches!(
                    s.kind,
                    crate::format::drawingml::ChartKind::Pie
                        | crate::format::drawingml::ChartKind::Doughnut
                )
            }) || !chart.axis_labels_visible(!chart.series.iter().any(|s| s.bar_horizontal))
                || chart.data_table.is_some()
            {
                continue;
            }
            let Some(category) = chart.series.first().and_then(|s| s.categories.first()) else {
                continue;
            };
            if !document.objects.iter().filter(|o| o.text.as_ref() == Some(category)).any(|o| {
                let mut v = &o.visual;
                while let Visual::TextLayout { visual, .. } | Visual::Layer { visual, .. } = v { v = visual; }
                matches!(v, Visual::RichText { runs, .. } if runs.iter().any(|r| (r.font_size - expected).abs() < 0.01))
            }) { failures.push(format!("{}: category {:?} must use authored {}px", path.display(), category, expected)); }
        }
        assert!(failures.is_empty(), "{}", failures.join("\n"));
    }

    #[test]
    fn supplied_mixed_chart_area_is_behind_columns() {
        let document = crate::format::detect_and_parse(
            include_bytes!("../../tests/fixtures/corpus-stacked-mix.pptx"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let filled = document
            .objects
            .iter()
            .filter(|o| {
                o.unit_index == 2
                    && o.kind == ObjectKind::Shape
                    && o.bounds.width > 20.0
                    && o.bounds.height > 30.0
            })
            .filter_map(|o| {
                let mut visual = &o.visual;
                while let Visual::Layer { visual: inner, .. } = visual {
                    visual = inner;
                }
                match visual {
                    Visual::PaintedShape {
                        geometry,
                        fill: Paint::Solid(_),
                        ..
                    } => Some((o.z, geometry)),
                    _ => None,
                }
            })
            .collect::<Vec<_>>();
        let area = filled
            .iter()
            .filter(|(_, g)| matches!(g, Geometry::Path { .. }))
            .map(|(z, _)| *z)
            .min()
            .unwrap();
        let bar = filled
            .iter()
            .filter(|(_, g)| matches!(g, Geometry::Rectangle))
            .map(|(z, _)| *z)
            .max()
            .unwrap();
        assert!(area < bar, "area z={area} must be behind columns z={bar}");
    }

    #[test]
    fn supplied_pie_layout_keeps_all_nine_chart_borders() {
        let document = crate::format::detect_and_parse(
            include_bytes!("../../tests/fixtures/corpus-pie-layout.pptx"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let borders = document.objects.iter().filter(|o| o.kind == ObjectKind::Group).filter(|o| {
            let mut visual = &o.visual;
            while let Visual::Layer { visual: inner, .. } = visual { visual = inner; }
            matches!(visual, Visual::PaintedShape { stroke: Paint::Solid(0x000000ff), stroke_width, .. } if *stroke_width > 1.0)
        }).count();
        assert_eq!(borders, 9, "Office shows nine authored chart-area borders");
    }

    #[test]
    fn supplied_line_chart_keeps_only_authored_visible_labels() {
        let document = crate::format::detect_and_parse(
            include_bytes!("../../tests/fixtures/corpus-tdf105517.pptx"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let labels = document
            .objects
            .iter()
            .filter(|o| {
                o.text
                    .as_deref()
                    .is_some_and(|text| text.replace(',', "") == "220000")
            })
            .count();
        assert_eq!(
            labels, 1,
            "Office shows only the one undeleted 220,000 point label"
        );
        for text in ["1,200,000", "1,000,000", "220,000"] {
            let object = document
                .objects
                .iter()
                .find(|o| o.text.as_deref() == Some(text))
                .expect(text);
            let Visual::TextLayout { visual, .. } = &object.visual else {
                panic!("text layout");
            };
            let Visual::RichText { runs, .. } = visual.as_ref() else {
                panic!("text runs");
            };
            assert_eq!((runs[0].font_size, runs[0].bold), (32.0, true));
            assert!(
                object.bounds.width > 120.0,
                "numeric label must have room for all digits: {text}"
            );
        }
        let triangles = document.objects.iter().filter(|o| matches!(&o.visual,
            Visual::PaintedShape { geometry: Geometry::Path { commands, .. }, fill: Paint::Solid(0x226ca9ff), .. }
            if commands.len() == 4 && matches!(commands.last(), Some(PathCommand::ClosePath)))).count();
        assert_eq!(
            triangles, 5,
            "all five authored blue series points have automatic triangles"
        );
        for color in [0xf0ab00ff, 0x226ca9ff] {
            assert_eq!(document.objects.iter().filter(|o| matches!(&o.visual,
                Visual::PaintedShape { geometry: Geometry::Path { commands, .. }, stroke: Paint::Solid(c), .. }
                if *c == color && commands.len() == 5)).count(), 1,
                "each series preserves its five points in one stroked path");
        }
    }

    #[test]
    fn corpus_saved_autofit_does_not_shrink_again() {
        let document = crate::format::detect_and_parse(
            include_bytes!("../../tests/fixtures/corpus-TextFittingComparisonWithMSO_1.pptx"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let mut count = 0;
        for object in document
            .objects
            .iter()
            .filter(|o| o.text.as_deref().is_some_and(|t| t.contains("ABC abc")))
        {
            let mut visual = &object.visual;
            while let Visual::Layer { visual: inner, .. } = visual {
                visual = inner;
            }
            let Visual::TextLayout { layout, visual } = visual else {
                panic!("text layout")
            };
            assert_eq!(
                layout.auto_fit,
                crate::model::TextAutoFit::None,
                "slide {} already has a persisted final scale",
                object.unit_index + 1
            );
            assert!(layout.font_scale > 0.0 && layout.font_scale <= 1.0);
            let Visual::RichText { runs, .. } = visual.as_ref() else {
                panic!("rich text");
            };
            // Effective sizes measured in the cached Microsoft PowerPoint PDF, pages 1–16.
            let office_points = [
                36.0, 36.0, 33.0, 33.0, 33.0, 33.0, 31.0, 28.0, 25.0, 23.0, 20.0, 17.0, 14.0, 12.0,
                9.0, 9.0,
            ];
            let run = runs.iter().find(|r| r.text.contains("ABC")).unwrap();
            let actual_points = run.font_size * layout.font_scale / super::POINTS_TO_CSS_PIXELS;
            assert!(
                (actual_points - office_points[object.unit_index as usize]).abs() < 0.01,
                "slide {}: scaled size {}pt",
                object.unit_index + 1,
                actual_points
            );
            count += 1;
        }
        assert_eq!(count, 16);
    }

    #[test]
    fn table_style_header_border_overrides_whole_table() {
        let bytes = include_bytes!("../../tests/fixtures/bnc910045.pptx");
        let package = Package::open(bytes, Limits::default()).unwrap();
        let theme = parse_pptx_theme(&package, "ppt/theme/theme1.xml").unwrap();
        let style = parse_pptx_table_style(
            &package,
            "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
            &theme,
            Rect::default(),
        )
        .unwrap()
        .unwrap();
        let context = TableCellStyleContext {
            row: 0,
            column: 0,
            row_span: 1,
            column_span: 1,
            row_count: 3,
            column_count: 1,
            first_row: true,
            last_row: false,
            first_column: false,
            last_column: false,
            band_rows: true,
            band_columns: false,
        };
        assert_eq!(
            style.cell_border(TableBorderSide::Bottom, context).width,
            4.0
        );
    }

    #[test]
    fn table_style_fill_geometry_and_explicit_none_use_shared_semantics() {
        let source = Package::open(
            include_bytes!("../../tests/fixtures/table-style-cascade.pptx"),
            Limits::default(),
        )
        .unwrap();
        let theme = parse_pptx_theme(&source, "ppt/theme/theme1.xml").unwrap();
        let xml = source.required_part(super::TABLE_STYLES_PART).unwrap();
        let xml = std::str::from_utf8(&xml).unwrap()
            .replace("<a:solidFill><a:srgbClr val=\"EEEEEE\"/></a:solidFill>",
                "<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst><a:lin ang=\"2700000\" scaled=\"1\"/></a:gradFill>")
            .replace("<a:firstRow>", "<a:firstRow><a:tcStyle><a:fillRef idx=\"0\"/></a:tcStyle>");
        let style = super::parse_pptx_table_style_xml(
            xml.as_bytes(),
            &source,
            "{11111111-1111-1111-1111-111111111111}",
            &theme,
            Rect::default(),
        )
        .unwrap()
        .unwrap();
        let fill = style.background_fill.unwrap();
        let small = Rect {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 20.0,
        };
        let tall = Rect {
            height: 100.0,
            ..small
        };
        assert_ne!(
            fill.paint(small),
            fill.paint(tall),
            "gradient geometry is evaluated at the final cell dimensions"
        );
        assert_eq!(
            style.parts[&super::TableStyleRegion::FirstRow]
                .fill
                .as_ref()
                .unwrap()
                .paint(small),
            Paint::None
        );
    }

    #[test]
    fn built_in_light_style_is_available_without_package_definition() {
        assert!(
            super::built_in_pptx_table_style(
                "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}",
                &PptxTheme::default()
            )
            .is_some()
        );
    }

    #[test]
    fn supplied_chart_bg1_overrides_dark_slide_mapping() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/chart_pt_color_bg1.pptx"),
            Limits::default(),
        )
        .unwrap();
        let mut mapping = super::default_color_map();
        mapping.insert("bg1".to_owned(), "dk1".to_owned());
        let theme = parse_pptx_theme(&package, "ppt/theme/theme1.xml")
            .unwrap()
            .with_color_map(mapping);
        let chart = super::parse_basic_pptx_chart(&package, "ppt/charts/chart1.xml", &theme)
            .unwrap()
            .unwrap();
        assert_eq!(
            chart.series[0].point_colors[..2],
            [0x4472_c4ff, 0xd9d9_d9ff]
        );
    }

    #[test]
    fn supplied_bnc910045_resolves_cell_theme_fill_reference() {
        let bytes = include_bytes!("../../tests/fixtures/bnc910045.pptx");
        let package = Package::open(bytes, Limits::default()).unwrap();
        let theme = parse_pptx_theme(&package, "ppt/theme/theme1.xml").unwrap();
        let style = parse_pptx_table_style(
            &package,
            "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}",
            &theme,
            Rect::default(),
        )
        .unwrap()
        .unwrap();
        assert_eq!(
            style.parts[&super::TableStyleRegion::FirstRow]
                .text
                .font_color,
            Some(0xffff_ffff)
        );
        assert_eq!(
            style.parts[&super::TableStyleRegion::FirstRow]
                .fill
                .as_ref()
                .map(|fill| fill.paint(Rect::default())),
            Some(Paint::Solid(0x4f81_bdff))
        );
        assert_eq!(
            style.parts[&super::TableStyleRegion::WholeTable]
                .fill
                .as_ref()
                .map(|fill| fill.paint(Rect::default())),
            Some(Paint::None)
        );
        let document = crate::format::detect_and_parse(bytes, Limits::default())
            .unwrap()
            .unwrap();
        let cell = document
            .objects
            .iter()
            .find(|object| {
                object.text.as_deref() == Some("Table with blue background in the first row.")
            })
            .unwrap();
        assert!(
            cell.bounds.height >= 96.0,
            "three 18pt lines and cell insets must fit: {:?}",
            cell.bounds
        );
    }

    #[test]
    fn supplied_bitmap_picture_inherits_layout_round_rect_and_crop() {
        let document = crate::format::detect_and_parse(
            include_bytes!("../../tests/fixtures/customshape-bitmapfill-srcrect.pptx"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let picture = document
            .objects
            .iter()
            .find(|object| object.kind == ObjectKind::Image)
            .unwrap();
        fn painted(visual: &Visual) -> (&Geometry, &Paint) {
            match visual {
                Visual::PaintedShape { geometry, fill, .. } => (geometry, fill),
                Visual::Layer { visual, .. }
                | Visual::StrokeStyle { visual, .. }
                | Visual::Effect { visual, .. }
                | Visual::AdvancedEffect { visual, .. } => painted(visual),
                _ => panic!("picture must be clipped to its inherited round rectangle"),
            }
        }
        let (geometry, fill) = painted(&picture.visual);
        let expected = drawingml_preset_geometry(
            "roundRect",
            picture.bounds,
            &HashMap::from([("adj".to_owned(), 10813.0)]),
        )
        .unwrap()
        .0;
        assert_eq!(*geometry, expected);
        let Paint::Image { crop, .. } = fill else {
            panic!("image fill is retained")
        };
        assert!((crop.left - 0.04393).abs() < 0.00001);
        assert!((crop.right - 0.04393).abs() < 0.00001);
    }

    #[test]
    fn layout_preset_inheritance_preserves_local_geometry_and_size() {
        let source = include_bytes!("../../tests/fixtures/customshape-bitmapfill-srcrect.pptx");
        let archive = crate::zip::ZipArchive::parse(source, Limits::default()).unwrap();
        for (properties, is_shape, expected_preset) in [
            ("<a:prstGeom prst=\"rect\"/>", false, "rect"),
            (
                "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1905000\" cy=\"381000\"/></a:xfrm>",
                false,
                "roundRect",
            ),
            ("", true, "roundRect"),
            ("<a:prstGeom prst=\"rect\"/>", true, "rect"),
            (
                "<a:custGeom><a:pathLst><a:path w=\"1\" h=\"1\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo><a:lnTo><a:pt x=\"1\" y=\"1\"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>",
                true,
                "custom",
            ),
        ] {
            let parts: Vec<_> = archive
                .entries()
                .iter()
                .map(|entry| {
                    let mut bytes = archive.extract(entry).unwrap();
                    if entry.name() == SLIDE_PART {
                        let mut xml = String::from_utf8(bytes)
                            .unwrap()
                            .replace("<p:spPr/>", &format!("<p:spPr>{properties}</p:spPr>"));
                        if is_shape {
                            xml = xml
                                .replace("p:pic", "p:sp")
                                .replace("p:nvPicPr", "p:nvSpPr")
                                .replace("p:cNvPicPr", "p:cNvSpPr");
                        }
                        bytes = xml.into_bytes();
                    }
                    (entry.name(), bytes)
                })
                .collect();
            let bytes = stored_zip(
                &parts
                    .iter()
                    .map(|(name, bytes)| (*name, bytes.as_slice()))
                    .collect::<Vec<_>>(),
            );
            let document = crate::format::detect_and_parse(&bytes, Limits::default())
                .unwrap()
                .unwrap();
            let object = document
                .objects
                .iter()
                .find(|object| object.source.part == SLIDE_PART)
                .unwrap();
            fn geometry(visual: &Visual) -> Geometry {
                match visual {
                    Visual::PaintedShape { geometry, .. } | Visual::RichText { geometry, .. } => {
                        geometry.clone()
                    }
                    Visual::Image { .. } => Geometry::Rectangle,
                    Visual::Layer { visual, .. }
                    | Visual::StrokeStyle { visual, .. }
                    | Visual::TextLayout { visual, .. }
                    | Visual::Effect { visual, .. }
                    | Visual::AdvancedEffect { visual, .. } => geometry(visual),
                    _ => panic!("expected inherited or explicitly overridden geometry: {visual:?}"),
                }
            }
            if expected_preset == "custom" {
                let Geometry::Path { commands, .. } = geometry(&object.visual) else {
                    panic!("local custom path wins")
                };
                assert_eq!(commands.len(), 3);
                assert!(matches!(commands[1], PathCommand::LineTo { .. }));
                continue;
            }
            let expected = if expected_preset == "rect" {
                Geometry::Rectangle
            } else {
                drawingml_preset_geometry(
                    "roundRect",
                    object.bounds,
                    &HashMap::from([("adj".to_owned(), 10813.0)]),
                )
                .unwrap()
                .0
            };
            assert_eq!(
                geometry(&object.visual),
                expected,
                "{properties}, shape={is_shape}"
            );
        }
    }

    #[test]
    fn supplied_placeholder_priority_keeps_each_layout_text_color() {
        let document = crate::format::detect_and_parse(
            include_bytes!("../../tests/fixtures/placeholder-priority.pptx"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        fn check(visual: &Visual, found: &mut usize) {
            match visual {
                Visual::RichText { runs, .. } => {
                    for run in runs {
                        let expected = match run.text.as_str() {
                            "aaa" => 0xff6a_52ff,
                            "bbb" => 0x00b0_50ff,
                            _ => continue,
                        };
                        assert_eq!(
                            run.color, expected,
                            "{} must inherit its own layout placeholder",
                            run.text
                        );
                        *found += 1;
                    }
                }
                Visual::TextLayout { visual, .. }
                | Visual::Layer { visual, .. }
                | Visual::AdvancedEffect { visual, .. } => check(visual, found),
                _ => {}
            }
        }
        let mut found = 0;
        for object in &document.objects {
            check(&object.visual, &mut found);
        }
        assert_eq!(found, 2);
    }

    #[test]
    fn supplied_n819614_smartart_keeps_emu_connector_coordinates() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/n819614.pptx"),
            Limits::default(),
        )
        .unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();
        let paths: Vec<_> = objects
            .iter()
            .filter_map(|object| match &object.visual {
                Visual::PaintedShape {
                    geometry: Geometry::Path { commands, .. },
                    ..
                } => Some(commands),
                _ => None,
            })
            .collect();
        assert_eq!(
            paths.len(),
            55,
            "all original SmartArt connectors are present"
        );
        // drawing1.xml's first connector deliberately extends beyond its 91440-EMU height.
        let expected = [(563107.0, 45720.0), (563107.0, 121221.0), (0.0, 121221.0)];
        assert_eq!(paths[0].len(), expected.len());
        for (command, (x, y)) in paths[0].iter().zip(expected) {
            let (actual_x, actual_y) = match command {
                PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => (*x, *y),
                _ => panic!("expected an orthogonal connector"),
            };
            assert!((actual_x - x / 9525.0).abs() < 0.001, "x={actual_x}");
            assert!((actual_y - y / 9525.0).abs() < 0.001, "y={actual_y}");
        }
        let part = "ppt/diagrams/drawing1.xml";
        let mut authored_points = Vec::new();
        super::parse_xml(
            &package.required_part(part).unwrap(),
            package.limits(),
            |event| {
                if let super::XmlEvent::StartElement {
                    name, attributes, ..
                } = event
                    && super::local_name(name) == "pt"
                {
                    authored_points.push((
                        super::numeric_attribute(&attributes, "x", part)?.unwrap() as f32 / 9525.0,
                        super::numeric_attribute(&attributes, "y", part)?.unwrap() as f32 / 9525.0,
                    ));
                }
                Ok(())
            },
        )
        .unwrap();
        let commands: Vec<_> = paths.iter().flat_map(|commands| commands.iter()).collect();
        assert_eq!(commands.len(), authored_points.len());
        for (command, (expected_x, expected_y)) in commands.iter().zip(authored_points) {
            let (PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y }) = command else {
                panic!("expected an orthogonal connector");
            };
            assert!(
                (x - expected_x).abs() < 0.001 && (y - expected_y).abs() < 0.001,
                "connector point {command:?} differs from authored ({expected_x}, {expected_y})"
            );
        }
    }

    #[test]
    fn drawingml_auto_numbers_preserve_roman_and_east_asian_kinds() {
        assert_eq!(format_auto_number("romanUcParenBoth", 4), "(IV)");
        assert_eq!(format_auto_number("romanLcParenBoth", 4), "(iv)");
        assert_eq!(format_auto_number("ea1JpnKorPeriod", 1), "一.");
    }

    #[test]
    fn chart_category_rotation_uses_label_width_at_the_authored_font_size() {
        let dates = ["1/5/2002".to_owned(), "1/6/2002".to_owned()];
        assert!(chart_category_labels_overlap(&dates, 74.0, 24.0));
        assert!(!chart_category_labels_overlap(&dates, 74.0, 10.0));
    }

    #[test]
    fn msgraph_labels_keep_their_region_font_style() {
        assert_eq!(
            chart_text_style(12.0, false, Some(25.0), true, true, true),
            (12.0, false)
        );
    }

    #[test]
    fn embedded_text_metrics_scale_with_the_ole_frame() {
        let visual = scale_embedded_text_visual(
            Visual::TextLayout {
                layout: TextLayout {
                    inset_left: 4.0,
                    ..TextLayout::default()
                },
                visual: Box::new(Visual::RichText {
                    geometry: Geometry::Rectangle,
                    fill: Paint::None,
                    stroke: Paint::None,
                    stroke_width: 2.0,
                    align: crate::model::TextAlign::Start,
                    line_height: 24.0,
                    runs: vec![TextRun {
                        paint: None,
                        east_asian_line_breaks: true,
                        text: "OLE".to_owned(),
                        font_family: "Arial".to_owned(),
                        font_size: 20.0,
                        color: 0x0000_00ff,
                        bold: false,
                        italic: false,
                        underline: false,
                        strikethrough: false,
                        highlight: 0,
                        baseline_shift: 2.0,
                        letter_spacing: 1.0,
                        horizontal_scale: 1.0,
                    }],
                }),
            },
            0.5,
        );
        let Visual::TextLayout { layout, visual } = visual else {
            panic!("embedded text keeps its layout");
        };
        let Visual::RichText {
            line_height, runs, ..
        } = visual.as_ref()
        else {
            panic!("embedded text keeps its rich text");
        };
        assert_eq!(layout.inset_left, 2.0);
        assert_eq!(*line_height, 12.0);
        assert_eq!(runs[0].font_size, 10.0);
        assert_eq!(runs[0].baseline_shift, 1.0);
        assert_eq!(runs[0].letter_spacing, 0.5);
    }

    #[test]
    fn zero_fill_style_reference_keeps_shape_unfilled() {
        let bytes = stored_zip(&[(
            SLIDE_PART,
            br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
              <p:sp><p:nvSpPr><p:cNvPr id="2"/></p:nvSpPr><p:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                <a:custGeom><a:avLst/><a:pathLst><a:path w="1" h="1"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="1" y="1"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>
              </p:spPr><p:style><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef></p:style></p:sp>
            </p:spTree></p:cSld></p:sld>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).expect("shape parses");

        assert!(objects.iter().any(|object| matches!(
            object.visual,
            Visual::PaintedShape {
                fill: Paint::None,
                ..
            }
        )));
    }

    #[test]
    fn supplied_empty_chart_title_is_rendered_in_pptx() {
        let source = Package::open(
            include_bytes!("../../tests/fixtures/chart-empty-title-original.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let chart = source.required_part("xl/charts/chart1.xml").unwrap();
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:c="c" xmlns:r="r"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="15"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="2857500"/></p:xfrm><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", br#"<Relationships><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#),
            ("ppt/charts/chart1.xml", &chart),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();
        assert_eq!(
            objects
                .iter()
                .filter(|object| object.text.as_deref() == Some("Chart Title"))
                .count(),
            1
        );
    }

    #[test]
    fn area_chart_preserves_the_authored_series_opacity() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:c="c" xmlns:r="r"><p:cSld><p:spTree>
          <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="15"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="2857500"/></p:xfrm>
            <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let chart = br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:areaChart><c:ser>
          <c:spPr><a:solidFill><a:srgbClr val="4F81BD"/></a:solidFill></c:spPr>
          <c:val><c:numLit><c:pt idx="0"><c:v>32</c:v></c:pt><c:pt idx="1"><c:v>12</c:v></c:pt></c:numLit></c:val>
        </c:ser></c:areaChart></c:plotArea></c:chart></c:chartSpace>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            (
                "ppt/slides/_rels/slide1.xml.rels",
                br#"<Relationships><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
            ),
            ("ppt/charts/chart1.xml", chart),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).expect("area chart parses");
        assert!(objects.iter().any(|object| matches!(
            object.visual,
            Visual::PaintedShape {
                fill: Paint::Solid(0x4f81_bdff),
                ..
            }
        )));
    }

    #[test]
    fn pie_3d_chart_emits_depth_and_top_faces() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:c="c" xmlns:r="r"><p:cSld><p:spTree>
          <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="16"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="2857500"/></p:xfrm>
            <a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let chart = br#"<c:chartSpace xmlns:c="c"><c:chart><c:view3D><c:rotX val="30"/><c:rAngAx val="0"/><c:perspective val="30"/></c:view3D><c:plotArea><c:pie3DChart><c:varyColors val="1"/><c:ser><c:explosion val="25"/>
          <c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
        </c:ser></c:pie3DChart></c:plotArea></c:chart></c:chartSpace>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            (
                "ppt/slides/_rels/slide1.xml.rels",
                br#"<Relationships><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
            ),
            ("ppt/charts/chart1.xml", chart),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).expect("pie3D chart parses");

        assert!(objects.iter().any(|object| matches!(
            object.visual,
            Visual::PaintedShape {
                fill: Paint::LinearGradient { .. },
                ..
            }
        )));

        for column in 0..2 {
            let faces = objects
                .iter()
                .filter(|object| {
                    object.source.part == "ppt/charts/chart1.xml"
                        && matches!(
                            object.source.locator,
                            SourceLocator::PptxShape {
                                row: Some(0),
                                column: Some(value),
                                ..
                            } if value == column
                        )
                })
                .collect::<Vec<_>>();
            let side = faces
                .iter()
                .find(|object| {
                    matches!(
                        object.visual,
                        Visual::PaintedShape {
                            fill: Paint::LinearGradient { .. },
                            ..
                        }
                    )
                })
                .expect("3D pie slice has a shaded side face");
            let top = faces
                .iter()
                .find(|object| {
                    matches!(
                        object.visual,
                        Visual::PaintedShape {
                            fill: Paint::Solid(_),
                            ..
                        }
                    )
                })
                .expect("3D pie slice has a solid top face");
            assert!(side.bounds.width > 0.0 && side.bounds.height > 0.0);
            assert!(matches!(
                top.visual,
                Visual::PaintedShape {
                    stroke: Paint::None,
                    stroke_width: 0.0,
                    ..
                }
            ));
            let Visual::PaintedShape {
                geometry: Geometry::Path { commands, .. },
                ..
            } = &top.visual
            else {
                panic!("3D pie top face must be a path");
            };
            assert!(
                commands
                    .iter()
                    .filter(|command| matches!(command, PathCommand::BezierCurveTo { .. }))
                    .count()
                    >= 28,
                "3D pie arcs must use smooth five-degree curve segments"
            );
            assert_eq!(
                commands
                    .iter()
                    .filter(|command| matches!(command, PathCommand::LineTo { .. }))
                    .count(),
                1,
                "only the radial edge may remain straight"
            );
        }
        let top_position = |column| {
            objects
                .iter()
                .position(|object| {
                    matches!(
                        object.source.locator,
                        SourceLocator::PptxShape {
                            row: Some(0),
                            column: Some(value),
                            ..
                        } if value == column
                    ) && matches!(
                        object.visual,
                        Visual::PaintedShape {
                            fill: Paint::Solid(_),
                            ..
                        }
                    )
                })
                .unwrap()
        };
        let last_side_position = |column| {
            objects
                .iter()
                .rposition(|object| {
                    matches!(
                        object.source.locator,
                        SourceLocator::PptxShape {
                            row: Some(0),
                            column: Some(value),
                            ..
                        } if value == column
                    ) && matches!(
                        object.visual,
                        Visual::PaintedShape {
                            fill: Paint::LinearGradient { .. },
                            ..
                        }
                    )
                })
                .unwrap()
        };
        assert!(last_side_position(1) < top_position(0));
        assert!(top_position(1) < top_position(0));
    }

    #[test]
    fn pie_chart_legend_uses_categories() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:c="c" xmlns:r="r"><p:cSld><p:spTree>
          <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="17"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="2857500"/></p:xfrm>
            <a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let chart = br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:pieChart><c:ser>
          <c:tx><c:v>East</c:v></c:tx>
          <c:dLbls><c:dLbl><c:idx val="1"/><c:delete val="1"/></c:dLbl><c:showCatName val="1"/><c:dLblPos val="bestFit"/></c:dLbls>
          <c:cat><c:strLit><c:pt idx="0"><c:v>1st Qtr</c:v></c:pt><c:pt idx="1"><c:v>2nd Qtr</c:v></c:pt></c:strLit></c:cat>
          <c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
        </c:ser></c:pieChart></c:plotArea><c:legend/></c:chart></c:chartSpace>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            (
                "ppt/slides/_rels/slide1.xml.rels",
                br#"<Relationships><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
            ),
            ("ppt/charts/chart1.xml", chart),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).expect("pie chart parses");
        let labels = objects
            .iter()
            .filter_map(|object| object.text.as_deref())
            .collect::<Vec<_>>();

        assert!(labels.contains(&"1st Qtr"));
        assert!(labels.contains(&"2nd Qtr"));
        assert_eq!(
            labels.iter().filter(|label| **label == "1st Qtr").count(),
            2
        );
        assert_eq!(
            labels.iter().filter(|label| **label == "2nd Qtr").count(),
            1
        );
        assert!(!labels.contains(&"East"));
        let first_legend = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("1st Qtr"))
            .unwrap();
        assert!(
            first_legend.bounds.y > 100.0,
            "the default right legend must be vertically centered, not treated as top-right"
        );
        assert!(objects.iter().any(|object| matches!(
            (&object.source.locator, &object.visual),
            (
                SourceLocator::PptxShape {
                    row: Some(0),
                    column: None,
                    ..
                },
                Visual::PaintedShape {
                    geometry: Geometry::Rectangle,
                    stroke: Paint::None,
                    stroke_width: 0.0,
                    ..
                }
            )
        )));
    }

    #[test]
    fn chart_series_without_explicit_colors_use_the_chart_theme() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:areaChart>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
            </c:areaChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut theme = PptxTheme::default();
        theme.colors.insert("accent1".to_owned(), 0x4f81_bdff);
        theme.colors.insert("accent2".to_owned(), 0xc050_4dff);

        let chart = super::parse_basic_pptx_chart(&package, PART, &theme)
            .expect("chart parses")
            .expect("chart is supported");

        assert_eq!(
            chart
                .series
                .iter()
                .map(|series| series.color)
                .collect::<Vec<_>>(),
            [Some(0x4f81_bdff), Some(0xc050_4dff)]
        );
    }

    #[test]
    fn area_chart_category_labels_use_office_text_color() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:c="c" xmlns:r="r"><p:cSld><p:spTree>
          <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="6096000" cy="4064000"/></p:xfrm>
            <a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let chart = br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:areaChart><c:ser>
          <c:cat><c:strLit><c:pt idx="0"><c:v>1/5/2002</c:v></c:pt></c:strLit></c:cat>
          <c:val><c:numLit><c:pt idx="0"><c:v>32</c:v></c:pt></c:numLit></c:val>
        </c:ser></c:areaChart><c:catAx/><c:valAx/></c:plotArea></c:chart></c:chartSpace>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            (
                "ppt/slides/_rels/slide1.xml.rels",
                br#"<Relationships><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
            ),
            ("ppt/charts/chart1.xml", chart),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();

        let label = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("1/5/2002"))
            .expect("area chart category label");
        let Visual::TextLayout { visual, .. } = &label.visual else {
            panic!("category label keeps its text layout");
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("category label remains rich text");
        };
        assert_eq!(runs[0].color, 0x0000_00ff);
    }

    #[test]
    fn placeholder_lookup_keeps_index_collisions_type_safe() {
        assert_eq!(
            super::placeholder_lookup_keys(Some(2), Some("body")),
            ["index:2:type:body", "type:body", "index:2"]
        );
    }

    #[test]
    fn theme_ignores_unselected_extra_color_schemes() {
        const PART: &str = "ppt/theme/theme1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<a:theme xmlns:a="a"><a:themeElements>
              <a:clrScheme name="active"><a:accent1><a:srgbClr val="0099CC"/></a:accent1></a:clrScheme>
            </a:themeElements><a:extraClrSchemeLst><a:extraClrScheme>
              <a:clrScheme name="unused"><a:accent1><a:srgbClr val="FF9900"/></a:accent1></a:clrScheme>
            </a:extraClrScheme></a:extraClrSchemeLst></a:theme>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let theme = parse_pptx_theme(&package, PART).unwrap();

        assert_eq!(theme.color("accent1"), Some(0x0099_ccff));
    }

    #[test]
    fn theme_line_references_preserve_authored_widths() {
        const PART: &str = "ppt/theme/theme1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<a:theme xmlns:a="a"><a:themeElements><a:fmtScheme name="Office">
              <a:lnStyleLst>
                <a:ln w="9525"/><a:ln w="25400"/><a:ln w="38100"/>
              </a:lnStyleLst>
            </a:fmtScheme></a:themeElements></a:theme>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let theme = parse_pptx_theme(&package, PART).unwrap();

        assert_eq!(theme.line_width(1), Some(1.0));
        assert!(
            theme
                .line_width(2)
                .is_some_and(|width| (width - 8.0 / 3.0).abs() < 0.001)
        );
        assert_eq!(theme.line_width(3), Some(4.0));
    }

    #[test]
    fn theme_background_image_fill_preserves_tile_and_duotone() {
        const PART: &str = "ppt/theme/theme1.xml";
        let bytes = stored_zip(&[
            (PART, br#"<a:theme xmlns:a="a" xmlns:r="r"><a:themeElements>
              <a:fmtScheme name="Office"><a:bgFillStyleLst><a:solidFill/>
                <a:blipFill><a:blip r:embed="rId1"><a:duotone>
                  <a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr>
                  <a:schemeClr val="phClr"/>
                </a:duotone></a:blip><a:tile/>
                </a:blipFill>
              </a:bgFillStyleLst></a:fmtScheme>
            </a:themeElements></a:theme>"#),
            ("ppt/theme/_rels/theme1.xml.rels", br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/></Relationships>"#),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let theme = parse_pptx_theme(&package, PART).unwrap();
        let fill = &theme.background_image_fills[1];

        assert!(fill.tile);
        assert!(fill.relationship.is_some());
        assert!(fill.adjustment(0xff00_00ff, &theme).duotone.is_some());
    }

    fn text_run_state_for_font_routing(text: &str, east_asian: Option<&str>) -> TextRunState {
        TextRunState {
            stroke: None,
            stroke_width: 0.0,
            glow: None,
            paint: None,
            depth: 0,
            text: text.to_owned(),
            replacement_text: None,
            font_family: "Segoe UI".to_owned(),
            font_east_asian: east_asian.map(str::to_owned),
            font_complex_script: Some("Times New Roman".to_owned()),
            font_symbol: None,
            font_size: 20.0,
            color: 0xff000000,
            bold: false,
            italic: false,
            underline: false,
            wavy_underline: false,
            dotted_underline: false,
            heavy_underline: false,
            double_underline: false,
            dot_dash_underline: false,
            double_strikethrough: false,
            strikethrough: false,
            highlight: 0,
            baseline_shift: 0.0,
            letter_spacing: 0.0,
            east_asian_line_breaks: true,
            capitalization: TextCapitalization::None,
            shadow: None,
            inner_shadow: None,
            reflection: None,
            shadow_scale_x: 1.0,
            shadow_scale_y: 1.0,
            shadow_skew_x: 0.0,
            shadow_skew_y: 0.0,
            shadow_alignment: 7,
        }
    }

    #[test]
    fn drawingml_font_routing_segments_unicode_blocks_with_east_asian_font() {
        let runs = text_run_state_for_font_routing("Aé–◇☀✦عZ", Some("MS Mincho")).finish();

        assert_eq!(
            runs.iter()
                .map(|run| (run.text.as_str(), run.font_family.as_str()))
                .collect::<Vec<_>>(),
            [
                ("Aé–", "Segoe UI"),
                ("◇☀✦", "MS Mincho"),
                ("ع", "Times New Roman"),
                ("Z", "Segoe UI"),
            ]
        );
    }

    #[test]
    fn drawingml_font_routing_keeps_symbols_on_latin_without_east_asian_font() {
        let runs = text_run_state_for_font_routing("A☀Z", None).finish();

        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "A☀Z");
        assert_eq!(runs[0].font_family, "Segoe UI");
    }

    #[test]
    fn drawingml_preset_dashes_keep_standard_and_system_rhythms_distinct() {
        for (value, expected) in [
            ("solid", &[][..]),
            ("dot", &[1.0, 3.0][..]),
            ("sysDot", &[1.0, 1.0][..]),
            ("dash", &[4.0, 3.0][..]),
            ("sysDash", &[3.0, 1.0][..]),
            ("lgDash", &[8.0, 3.0][..]),
            ("dashDot", &[4.0, 3.0, 1.0, 3.0][..]),
            ("sysDashDot", &[3.0, 1.0, 1.0, 1.0][..]),
            ("lgDashDot", &[8.0, 3.0, 1.0, 3.0][..]),
            ("lgDashDotDot", &[8.0, 3.0, 1.0, 3.0, 1.0, 3.0][..]),
            ("sysDashDotDot", &[3.0, 1.0, 1.0, 1.0, 1.0, 1.0][..]),
        ] {
            assert_eq!(
                DrawingMlDashPattern::from_attribute(Some(value)).lengths(),
                expected
            );
        }
    }

    fn single_font_metric(
        family: &str,
        character: char,
        advance_em: f32,
        italic: bool,
        bold: bool,
    ) -> FontMetricTable {
        let family = family.as_bytes();
        let metrics_offset = 20 + 16 + family.len();
        let mut bytes = Vec::new();
        bytes.extend_from_slice(b"OVFM");
        bytes.extend_from_slice(&1_u16.to_le_bytes());
        bytes.extend_from_slice(&20_u16.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(&(metrics_offset as u32).to_le_bytes());
        bytes.extend_from_slice(&(family.len() as u16).to_le_bytes());
        bytes.push(u8::from(italic));
        bytes.push(4);
        bytes.extend_from_slice(&(if bold { 700_u16 } else { 400_u16 }).to_le_bytes());
        bytes.extend_from_slice(&0_u16.to_le_bytes());
        bytes.extend_from_slice(&0_u32.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(family);
        bytes.extend_from_slice(&(character as u32).to_le_bytes());
        bytes.extend_from_slice(&advance_em.to_bits().to_le_bytes());
        FontMetricTable::decode(&bytes, Limits::default()).unwrap()
    }

    #[test]
    fn drawingml_line_end_sizes_follow_powerpoint_stroke_multipliers() {
        for (size, multiplier) in [
            (DrawingMlLineEndSize::Small, 2.0),
            (DrawingMlLineEndSize::Medium, 3.0),
            (DrawingMlLineEndSize::Large, 5.0),
        ] {
            let mut commands = Vec::new();
            append_line_end(
                &mut commands,
                (0.0, 0.0),
                (100.0, 0.0),
                4.0,
                DrawingMlLineEnd {
                    kind: DrawingMlLineEndKind::Triangle,
                    width: size,
                    length: size,
                },
            );
            let extent = 4.0 * multiplier;
            assert!(matches!(
                commands.as_slice(),
                [
                    PathCommand::MoveTo { x: tip_x, y: tip_y },
                    PathCommand::LineTo { x: left_x, y: left_y },
                    PathCommand::LineTo { x: right_x, y: right_y },
                    PathCommand::ClosePath,
                ] if tip_x.abs() < 0.001
                    && tip_y.abs() < 0.001
                    && (*left_x - extent).abs() < 0.001
                    && (*right_x - extent).abs() < 0.001
                    && (*left_y - extent / 2.0).abs() < 0.001
                    && (*right_y + extent / 2.0).abs() < 0.001
            ));
        }
    }

    #[test]
    fn drawingml_line_ends_keep_powerpoint_hairline_size_floor() {
        let mut commands = Vec::new();
        append_line_end(
            &mut commands,
            (0.0, 0.0),
            (100.0, 0.0),
            1.0,
            DrawingMlLineEnd {
                kind: DrawingMlLineEndKind::Stealth,
                width: DrawingMlLineEndSize::Medium,
                length: DrawingMlLineEndSize::Medium,
            },
        );
        assert!(matches!(
            commands.as_slice(),
            [
                PathCommand::MoveTo { x: tip_x, y: tip_y },
                PathCommand::LineTo { x: left_x, y: left_y },
                PathCommand::LineTo { x: notch_x, y: notch_y },
                PathCommand::LineTo { x: right_x, y: right_y },
                PathCommand::ClosePath,
            ] if tip_x.abs() < 0.001
                && tip_y.abs() < 0.001
                && (*left_x - 8.0).abs() < 0.001
                && (*right_x - 8.0).abs() < 0.001
                && (*left_y - 4.0).abs() < 0.001
                && (*right_y + 4.0).abs() < 0.001
                && (*notch_x - 4.8).abs() < 0.001
                && notch_y.abs() < 0.001
        ));
    }

    #[test]
    fn drawingml_line_end_length_does_not_collapse_on_short_group_child_lines() {
        let mut commands = Vec::new();
        append_line_end(
            &mut commands,
            (0.0, 0.0),
            (0.02, 0.02),
            1.0,
            DrawingMlLineEnd {
                kind: DrawingMlLineEndKind::Triangle,
                width: DrawingMlLineEndSize::Medium,
                length: DrawingMlLineEndSize::Medium,
            },
        );
        assert!(matches!(
            commands.get(1),
            Some(PathCommand::LineTo { x, y }) if x.hypot(*y) > 7.9
        ));
    }

    #[test]
    fn bent_connector_keeps_out_of_bounds_adjustment_and_tail_arrow() {
        let (Geometry::Path { commands, .. }, _) = connector_geometry(
            "bentConnector4",
            Rect {
                x: 0.0,
                y: 0.0,
                width: 100.0,
                height: 100.0,
            },
            &HashMap::from([
                ("adj1".to_owned(), 17_394.0),
                ("adj2".to_owned(), 123_079.0),
            ]),
            1.0,
            None,
            Some(DrawingMlLineEnd {
                kind: DrawingMlLineEndKind::Triangle,
                width: DrawingMlLineEndSize::Medium,
                length: DrawingMlLineEndSize::Medium,
            }),
        )
        .expect("bentConnector4 has a DrawingML geometry") else {
            panic!("connector uses path geometry");
        };

        assert_eq!(commands.len(), 9, "tail arrow commands are retained");
        assert!(matches!(
            commands.as_slice(),
            [
                PathCommand::MoveTo { .. },
                PathCommand::LineTo { .. },
                PathCommand::LineTo { .. },
                PathCommand::LineTo { x: bend_x, y: bend_y },
                PathCommand::LineTo { x: tip_x, y: tip_y },
                PathCommand::MoveTo { x: arrow_tip_x, y: arrow_tip_y },
                PathCommand::LineTo { .. },
                PathCommand::LineTo { .. },
                PathCommand::ClosePath,
            ] if (*bend_x - 100.0).abs() < 0.001
                && (*bend_y - 123.079).abs() < 0.001
                && (*tip_x - 100.0).abs() < 0.001
                && (*tip_y - 100.0).abs() < 0.001
                && (*arrow_tip_x - 100.0).abs() < 0.001
                && (*arrow_tip_y - 100.0).abs() < 0.001
        ));
    }

    #[test]
    fn plus_preset_keeps_drawingml_and_fallback_adjustments_distinct() {
        let bounds = Rect {
            x: 0.0,
            y: 0.0,
            width: 120.0,
            height: 60.0,
        };
        assert_eq!(
            super::preset_geometry("plus", bounds).unwrap().0,
            crate::format::cross_geometry(120.0, 60.0, 40.0, 80.0, 20.0, 40.0),
        );
        // DrawingML uses 25% of the shorter side, not the fallback's thirds.
        let geometry = drawingml_preset_geometry("plus", bounds, &HashMap::new())
            .unwrap()
            .0;
        assert_ne!(geometry, super::preset_geometry("plus", bounds).unwrap().0);
    }

    #[test]
    fn preset_round_rectangle_uses_shape_aspect_ratio_and_zero_adjustment() {
        let bounds = Rect {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 40.0,
        };
        let square = drawingml_preset_geometry(
            "roundRect",
            bounds,
            &HashMap::from([("adj".to_owned(), 0.0)]),
        )
        .expect("zero-radius round rectangle is a valid rectangle path");
        let stadium = drawingml_preset_geometry(
            "roundRect",
            bounds,
            &HashMap::from([("adj".to_owned(), 50_000.0)]),
        )
        .expect("maximum round-rectangle adjustment is a valid stadium path");

        let Geometry::Path {
            commands: square, ..
        } = square.0
        else {
            panic!("zero adjustment must produce a single path");
        };
        assert!(square.iter().all(|command| !matches!(
            command,
            PathCommand::BezierCurveTo { .. } | PathCommand::QuadraticCurveTo { .. }
        )));
        assert!(square.iter().any(|command| matches!(
            command,
            PathCommand::LineTo { x, y } if (*x - 200.0).abs() < 0.001 && y.abs() < 0.001
        )));

        let Geometry::Path {
            commands: stadium, ..
        } = stadium.0
        else {
            panic!("maximum adjustment must produce a single path");
        };
        assert!(matches!(
            stadium.first(),
            Some(PathCommand::MoveTo { x, y })
                if x.abs() < 0.001 && (*y - 20.0).abs() < 0.001
        ));
        assert!(stadium.iter().any(|command| matches!(
            command,
            PathCommand::BezierCurveTo { x, y, .. }
                if (*x - 20.0).abs() < 0.001 && y.abs() < 0.001
        )));
    }

    #[test]
    fn action_button_back_previous_points_left() {
        let (geometry, _) = drawingml_preset_geometry(
            "actionButtonBackPrevious",
            Rect {
                x: 0.0,
                y: 0.0,
                width: 64.0,
                height: 64.0,
            },
            &HashMap::new(),
        )
        .expect("back action button geometry");
        let Geometry::LayeredPath { layers } = geometry else {
            panic!("back action button must retain its layered icon");
        };
        assert_eq!(layers[1].fill, PathFillMode::Darken);
        assert_eq!(
            layers[1].commands[0],
            PathCommand::MoveTo { x: 8.0, y: 32.0 }
        );
    }

    #[test]
    fn preset_arc_keeps_fill_and_stroke_paths_in_one_coordinate_space() {
        let geometry = drawingml_preset_geometry(
            "arc",
            Rect {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 204.0,
            },
            &HashMap::from([
                ("adj1".to_owned(), 16_200_000.0),
                ("adj2".to_owned(), 5_486_615.0),
            ]),
        )
        .expect("arc preset geometry");

        let Geometry::LayeredPath { layers } = geometry.0 else {
            panic!("arc must retain its separate fill and stroke paths");
        };
        assert_eq!(layers.len(), 2);
        assert_eq!(layers[0].commands[..4], layers[1].commands[..]);
        assert!(matches!(
            layers[1].commands.first(),
            Some(PathCommand::MoveTo { x, y })
                if (*x - 100.0).abs() < 1.0 && y.abs() < 0.001
        ));
        assert!(layers[1].commands.iter().any(|command| matches!(
            command,
            PathCommand::BezierCurveTo { x, .. } if *x > 185.0
        )));
    }

    #[test]
    fn preset_bevel_retains_office_fill_layers_and_outline() {
        let geometry = drawingml_preset_geometry(
            "bevel",
            Rect {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 180.0,
            },
            &HashMap::from([("adj".to_owned(), 25_198.0)]),
        )
        .expect("bevel preset geometry");

        let Geometry::LayeredPath { layers } = geometry.0 else {
            panic!("bevel must retain its independently painted faces");
        };
        assert_eq!(layers.len(), 6);
        assert_eq!(
            layers
                .iter()
                .map(|layer| (layer.fill, layer.stroke))
                .collect::<Vec<_>>(),
            [
                (PathFillMode::Normal, false),
                (PathFillMode::LightenLess, false),
                (PathFillMode::DarkenLess, false),
                (PathFillMode::Lighten, false),
                (PathFillMode::Darken, false),
                (PathFillMode::None, true),
            ]
        );
        assert!(matches!(
            layers[0].commands.first(),
            Some(PathCommand::MoveTo { x, y })
                if (*x - 45.3564).abs() < 0.001 && (*y - 45.3564).abs() < 0.001
        ));
    }

    #[test]
    fn asymmetric_no_smoking_uses_drawingml_at2_argument_order() {
        let geometry = drawingml_preset_geometry(
            "noSmoking",
            Rect {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 100.0,
            },
            &HashMap::from([("adj".to_owned(), 13_901.0)]),
        )
        .expect("no-smoking preset geometry");

        let Geometry::Path { commands, .. } = geometry.0 else {
            panic!("no-smoking must produce one compound path");
        };
        let inner_start = commands
            .iter()
            .filter_map(|command| match command {
                PathCommand::MoveTo { x, y } => Some((*x, *y)),
                _ => None,
            })
            .nth(1)
            .expect("no-smoking inner ring start");
        assert!(
            (inner_start.0 - 164.78525).abs() < 0.001 && (inner_start.1 - 73.77656).abs() < 0.001,
            "unexpected no-smoking inner start: {inner_start:?}"
        );
    }

    #[test]
    fn curved_right_arrow_uses_drawingml_at2_argument_order() {
        let geometry = drawingml_preset_geometry(
            "curvedRightArrow",
            Rect {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 100.0,
            },
            &HashMap::new(),
        )
        .expect("curved-right-arrow preset geometry");

        let Geometry::LayeredPath { layers } = geometry.0 else {
            panic!("curved-right-arrow must retain its fill and outline layers");
        };
        assert_eq!(layers.len(), 3);
        assert!(matches!(
            layers[0].commands.get(1),
            Some(PathCommand::BezierCurveTo { x, y, .. })
                // The authored arc ends at x1=w-ah=175 on the ellipse
                // centered at (200,31.25), not at the old parameter-angle endpoint.
                if (*x - 175.0).abs() < 0.001
                    && (*y - (31.25 + 31.25 * (1.0_f32 - (25.0_f32 / 200.0).powi(2)).sqrt())).abs() < 0.001
        ));
    }

    #[test]
    fn preset_text_rectangle_adds_to_authored_body_insets() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="9"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr>
            <p:txBody><a:bodyPr lIns="95250" rIns="190500" tIns="47625" bIns="95250" vert="vert270"/><a:lstStyle/><a:p><a:r><a:t>Inside</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("ellipse text rectangle produces valid text layout");

        let shape = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Inside"))
            .expect("text shape is present");
        let Visual::TextLayout { layout, .. } = &shape.visual else {
            panic!("shape text uses a text layout");
        };
        assert!((layout.inset_left - 24.6447).abs() < 0.001);
        assert!((layout.inset_right - 34.6447).abs() < 0.001);
        assert!((layout.inset_top - 34.2893).abs() < 0.001);
        assert!((layout.inset_bottom - 39.2893).abs() < 0.001);
    }

    #[test]
    fn cardinal_body_rotation_uses_rotated_text_layout_geometry() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="4"/></p:nvSpPr>
            <p:spPr><a:xfrm rot="5400000"><a:off x="5596459" y="-4614347"/><a:ext cx="905100" cy="10148572"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr rot="-5400000"/><a:lstStyle/><a:p><a:r><a:t>Only 1 URSK is derived and activated</a:t></a:r></a:p></p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="5"/></p:nvSpPr>
            <p:spPr><a:xfrm rot="-5400000"><a:off x="0" y="0"/><a:ext cx="905100" cy="10148572"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr rot="5400000"/><a:lstStyle/><a:p><a:r><a:t>Opposite quarter turn</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("counter-rotated DrawingML text parses");

        let Visual::Layer { visual, .. } = &objects[0].visual else {
            panic!("the authored shape rotation is preserved");
        };
        let Visual::TextLayout { layout, .. } = visual.as_ref() else {
            panic!("shape text uses a text layout");
        };
        assert_eq!(layout.orientation, TextOrientation::Rotated270);
        assert_eq!(layout.rotation_degrees, 0.0);
        let Visual::Layer { visual, .. } = &objects[1].visual else {
            panic!("the opposite authored shape rotation is preserved");
        };
        let Visual::TextLayout { layout, .. } = visual.as_ref() else {
            panic!("opposite shape text uses a text layout");
        };
        assert_eq!(layout.orientation, TextOrientation::Rotated90);
        assert_eq!(layout.rotation_degrees, 0.0);
    }

    #[test]
    fn shape_flip_does_not_mirror_its_text() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="5"/></p:nvSpPr>
            <p:spPr><a:xfrm rot="16200000" flipH="1"><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Readable</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new()).expect("flipped text parses");

        let Visual::Group { children } = &objects[0].visual else {
            panic!("flipped text must be separated from the mirrored shape");
        };
        let Visual::Layer {
            transform: shape, ..
        } = &children[0].visual
        else {
            panic!("shape keeps its mirror transform");
        };
        let Visual::Layer {
            transform: text, ..
        } = &children[1].visual
        else {
            panic!("text keeps only the authored rotation");
        };
        assert!(shape.a * shape.d - shape.b * shape.c < 0.0);
        assert!(text.a * text.d - text.b * text.c > 0.0);
    }

    #[test]
    fn preserves_negative_drawingml_body_insets() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="9"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
            <p:txBody><a:bodyPr lIns="-9000" rIns="-9000" tIns="-45000" bIns="-45000"/><a:lstStyle/><a:p><a:r><a:t>Inside</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("signed body insets are valid DrawingML coordinates");

        let Visual::TextLayout { layout, .. } = &objects[0].visual else {
            panic!("shape text uses a text layout");
        };
        assert!(layout.inset_left < 0.0);
        assert!(layout.inset_top < 0.0);
    }

    #[test]
    fn accepts_percentage_lexical_forms_and_signed_color_transforms() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="9"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="336699"><a:tint val="75%"/><a:satOff val="-1000"/></a:srgbClr></a:solidFill><a:ln><a:miter lim="800%"/></a:ln></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:lnSpc><a:spcPct val="90%"/></a:lnSpc></a:pPr><a:r><a:t>Text</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("DrawingML percentage lexical forms and signed offsets are valid");

        assert_eq!(objects.len(), 1);
    }

    #[test]
    fn overflowing_preset_text_rectangle_falls_back_to_authored_vertical_insets() {
        let paragraphs = (0..8)
            .map(|index| {
                format!("<a:p><a:r><a:rPr sz=\"1067\"/><a:t>Line {index}</a:t></a:r></a:p>")
            })
            .collect::<String>();
        let slide = format!(
            r#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
              <p:sp>
                <p:nvSpPr><p:cNvPr id="10"/></p:nvSpPr>
                <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1333500"/></a:xfrm><a:prstGeom prst="flowChartMagneticDisk"><a:avLst/></a:prstGeom></p:spPr>
                <p:txBody><a:bodyPr tIns="9525" bIns="19050" anchor="ctr"/><a:lstStyle/>{paragraphs}</p:txBody>
              </p:sp>
            </p:spTree></p:cSld></p:sld>"#
        );
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("overflowing preset text produces valid text layout");

        let shape = objects
            .iter()
            .find(|object| {
                object
                    .text
                    .as_deref()
                    .is_some_and(|text| text.starts_with("Line 0"))
            })
            .expect("text shape is present");
        let Visual::TextLayout { layout, .. } = &shape.visual else {
            panic!("shape text uses a text layout");
        };
        assert!((layout.inset_top - 1.0).abs() < 0.001);
        assert!((layout.inset_bottom - 2.0).abs() < 0.001);
    }

    #[test]
    fn wrapped_preset_text_uses_measured_line_count_for_vertical_inset_fallback() {
        let shape = |id: u32, x: u64, width: u64| {
            format!(
                r#"<p:sp>
                  <p:nvSpPr><p:cNvPr id="{id}"/></p:nvSpPr>
                  <p:spPr><a:xfrm><a:off x="{x}" y="0"/><a:ext cx="{width}" cy="824437"/></a:xfrm><a:prstGeom prst="round2DiagRect"><a:avLst/></a:prstGeom></p:spPr>
                  <p:txBody><a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720" anchor="ctr"/><a:lstStyle/>
                    <a:p><a:pPr latinLnBrk="0"/><a:r><a:rPr sz="1400"/><a:t>Phase A</a:t></a:r></a:p>
                    <a:p><a:pPr latinLnBrk="0"/><a:r><a:rPr sz="1400"/><a:t>Alpha Beta Gamma Delta Epsilon Zeta Eta Theta</a:t></a:r></a:p>
                  </p:txBody>
                </p:sp>"#
            )
        };
        let slide = format!(
            r#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>{}{}</p:spTree></p:cSld></p:sld>"#,
            shape(12, 0, 2_139_834),
            shape(13, 2_500_000, 5_715_000),
        );
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("wrapped preset text produces valid text layouts");

        let layouts = objects
            .iter()
            .filter(|object| {
                object
                    .text
                    .as_deref()
                    .is_some_and(|text| text.starts_with("Phase A"))
            })
            .map(|object| {
                let Visual::TextLayout { layout, .. } = &object.visual else {
                    panic!("shape text uses a text layout");
                };
                (object.bounds.width, layout)
            })
            .collect::<Vec<_>>();
        assert_eq!(layouts.len(), 2);
        let narrow = layouts
            .iter()
            .find(|(width, _)| *width < 300.0)
            .map(|(_, layout)| *layout)
            .expect("narrow shape is present");
        let wide = layouts
            .iter()
            .find(|(width, _)| *width > 500.0)
            .map(|(_, layout)| *layout)
            .expect("wide shape is present");

        assert!((narrow.inset_top - 4.8).abs() < 0.001);
        assert!((narrow.inset_bottom - 4.8).abs() < 0.001);
        assert!(wide.inset_top > 4.8);
        assert!(wide.inset_bottom > 4.8);
    }

    #[test]
    fn fitting_preset_text_rectangle_keeps_its_vertical_insets() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="11"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2139834" cy="824437"/></a:xfrm><a:prstGeom prst="round2DiagRect"><a:avLst/></a:prstGeom></p:spPr>
            <p:txBody><a:bodyPr tIns="45720" bIns="45720" anchor="ctr"/><a:lstStyle/>
              <a:p><a:r><a:rPr sz="1400"/><a:t>First line</a:t></a:r></a:p>
              <a:p><a:r><a:rPr sz="1400"/><a:t>Second line</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("fitting preset text produces valid text layout");

        let shape = objects
            .iter()
            .find(|object| {
                object
                    .text
                    .as_deref()
                    .is_some_and(|text| text.starts_with("First line"))
            })
            .expect("text shape is present");
        let Visual::TextLayout { layout, .. } = &shape.visual else {
            panic!("shape text uses a text layout");
        };
        assert!(layout.inset_top > 4.8);
        assert!(layout.inset_bottom > 4.8);
    }

    #[test]
    fn shape_flips_mirror_once_around_the_authored_bounds_center() {
        let bounds = Rect {
            x: 100.0,
            y: 200.0,
            width: 80.0,
            height: 40.0,
        };
        let horizontal = shape_transform(bounds, 0.0, true, false);
        let vertical = shape_transform(bounds, 0.0, false, true);

        assert_eq!(
            (horizontal.a, horizontal.b, horizontal.c, horizontal.d),
            (-1.0, -0.0, -0.0, 1.0)
        );
        assert!((horizontal.e - 280.0).abs() < 0.001);
        assert!(horizontal.f.abs() < 0.001);
        assert_eq!(
            (vertical.a, vertical.b, vertical.c, vertical.d),
            (1.0, 0.0, 0.0, -1.0)
        );
        assert!(vertical.e.abs() < 0.001);
        assert!((vertical.f - 440.0).abs() < 0.001);
    }

    #[test]
    fn parses_an_embedded_raster_picture_with_exact_source_mapping() {
        let bytes = picture_package_bytes("../media/image1.png", None, PNG);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics).expect("valid embedded picture");

        assert_eq!(objects.len(), 1);
        let picture = &objects[0];
        assert_eq!(picture.kind, ObjectKind::Image);
        assert_eq!((picture.bounds.x, picture.bounds.y), (-1.0, 2.0));
        assert_eq!((picture.bounds.width, picture.bounds.height), (3.0, 4.0));
        assert_eq!(picture.source.mapping, MappingQuality::Exact);
        assert!(matches!(
            picture.source.locator,
            SourceLocator::PptxShape { shape_id: 42, .. }
        ));
        assert!(matches!(
            &picture.visual,
            Visual::Image {
                media_type, bytes, ..
            }
                if media_type == "image/png" && bytes == PNG
        ));
        assert!(diagnostics.is_empty());
    }

    #[test]
    fn supplied_picture_preserves_heroic_perspective_camera() {
        let document = crate::format::detect_and_parse(
            include_bytes!("../../tests/fixtures/image-3d-rotation.pptx"),
            Limits::default(),
        )
        .unwrap()
        .unwrap();
        let picture = document
            .objects
            .iter()
            .find(|object| object.kind == ObjectKind::Image)
            .unwrap();
        assert!(
            matches!(&picture.visual,
                Visual::AdvancedEffect { three_d: Some(style), .. }
                    if style.camera_preset == "perspectiveHeroicExtremeLeftFacing"
            ),
            "authored picture camera must reach the shared renderer: {:?}",
            picture.visual
        );
    }

    #[test]
    fn picture_reflection_uses_the_shared_advanced_effect_model() {
        let slide = slide_xml().replace(
            "<a:extLst>",
            "<a:effectLst><a:reflection blurRad=\"6350\" stA=\"50000\" endA=\"300\" endPos=\"90000\" sy=\"-100000\"/></a:effectLst><a:extLst>",
        );
        let relationships = relationships_xml("../media/image1.png", None);
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", relationships.as_bytes()),
            ("ppt/media/image1.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();

        assert!(matches!(
            &objects[0].visual,
            Visual::AdvancedEffect { reflection: Some(reflection), .. }
                if (reflection.start_opacity - 0.5).abs() < 0.0001
                    && (reflection.end_opacity - 0.003).abs() < 0.0001
                    && (reflection.end_position - 0.9).abs() < 0.0001
                    && (reflection.blur - 2.0 / 3.0).abs() < 0.001
        ));
    }

    #[test]
    fn picture_style_line_reference_paints_the_image_border() {
        let slide = slide_xml().replace(
            "</p:pic>",
            "<p:style><a:lnRef idx=\"3\"><a:schemeClr val=\"lt1\"/></a:lnRef><a:effectRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:effectRef></p:style></p:pic>",
        );
        let relationships = relationships_xml("../media/image1.png", None);
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", relationships.as_bytes()),
            ("ppt/media/image1.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();

        assert!(matches!(
            &objects[0].visual,
            Visual::AdvancedEffect { outer_shadow: Some(_), visual, .. }
                if matches!(visual.as_ref(), Visual::PaintedShape {
                    stroke: Paint::Solid(0xffff_ffff), stroke_width, ..
                } if *stroke_width > 0.0)
        ));
    }

    #[test]
    fn shape_font_reference_wins_over_inherited_paragraph_color() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp>
          <p:nvSpPr><p:cNvPr id="13"/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
          <p:style><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>White</a:t></a:r></a:p></p:txBody>
        </p:sp></p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();

        let Visual::TextLayout { visual, .. } = &objects[0].visual else {
            panic!("shape text keeps its text layout");
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("shape text remains rich text");
        };
        assert_eq!(runs[0].color, 0xffff_ffff);
    }

    #[test]
    fn smartart_text_uses_diagram_color_definition() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:dgm="dgm" xmlns:r="r"><p:cSld><p:spTree>
          <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
            <a:graphic><a:graphicData uri="diagram"><dgm:relIds r:dm="rIdData" r:cs="rIdColors"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let relationships = br#"<Relationships>
          <Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
          <Relationship Id="rIdColors" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"/>
          <Relationship Id="rIdDrawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/>
        </Relationships>"#;
        let data = br#"<dgm:dataModel xmlns:dgm="dgm" xmlns:a="a" xmlns:dsp="dsp"><dgm:ptLst>
          <dgm:pt modelId="data1"><dgm:t><a:p><a:r><a:t>Visible</a:t></a:r></a:p></dgm:t></dgm:pt>
          <dgm:pt modelId="pres1" type="pres"><dgm:prSet presAssocID="data1" presStyleLbl="node0"/></dgm:pt>
        </dgm:ptLst><dgm:extLst><a:ext><dsp:dataModelExt relId="rIdDrawing"/></a:ext></dgm:extLst></dgm:dataModel>"#;
        let colors = br#"<dgm:colorsDef xmlns:dgm="dgm" xmlns:a="a"><dgm:styleLbl name="node0">
          <dgm:txFillClrLst><a:schemeClr val="dk1"/></dgm:txFillClrLst>
        </dgm:styleLbl></dgm:colorsDef>"#;
        let drawing = br#"<dsp:drawing xmlns:dsp="dsp" xmlns:a="a"><dsp:spTree><dsp:sp modelId="pres1">
          <dsp:nvSpPr><dsp:cNvPr id="1"/></dsp:nvSpPr>
          <dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"/></dsp:spPr>
          <dsp:style><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>
          <dsp:txBody><a:bodyPr/><a:p><a:r><a:t>Visible</a:t></a:r></a:p></dsp:txBody>
        </dsp:sp></dsp:spTree></dsp:drawing>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", relationships),
            ("ppt/diagrams/data1.xml", data),
            ("ppt/diagrams/colors1.xml", colors),
            ("ppt/diagrams/drawing1.xml", drawing),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();

        let object = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Visible"))
            .expect("SmartArt drawing text");
        let Visual::TextLayout { visual, .. } = &object.visual else {
            panic!("SmartArt text keeps its text layout");
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("SmartArt text remains rich text");
        };
        assert_eq!(runs[0].color, 0x0000_00ff);
    }

    #[test]
    fn accepts_a_fully_cropped_picture_without_drawing_source_pixels() {
        let slide = slide_xml().replace(
            "</p:blipFill>",
            "<a:srcRect l=\"60000\" t=\"70000\" r=\"40000\" b=\"30000\"/></p:blipFill>",
        );
        let relationships = relationships_xml("../media/image1.png", None);
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", relationships.as_bytes()),
            ("ppt/media/image1.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("a fully cropped Office picture is valid but paints no source pixels");

        assert_eq!(objects.len(), 1);
    }

    #[test]
    fn missing_smartart_drawing_falls_back_to_diagram_data() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:dgm="dgm" xmlns:r="r"><p:cSld><p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="7"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
            <a:graphic><a:graphicData uri="diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let slide_relationships = br#"<Relationships><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>"#;
        let data = br#"<dgm:dataModel xmlns:dgm="dgm"><dgm:ptLst><dgm:pt modelId="1"><dgm:t><a:p xmlns:a="a"><a:r><a:t>Root</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>"#;
        let data_relationships = br#"<Relationships><Relationship Id="rIdDrawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/></Relationships>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", slide_relationships),
            ("ppt/diagrams/data1.xml", data),
            ("ppt/diagrams/_rels/data1.xml.rels", data_relationships),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics)
            .expect("missing SmartArt drawing uses the bounded data-model fallback");

        assert!(
            objects
                .iter()
                .any(|object| object.text.as_deref() == Some("Root"))
        );
        assert!(diagnostics.iter().any(|diagnostic| {
            diagnostic.code == DiagnosticCode::UnsupportedFeature
                && diagnostic.message.contains("drawing fallback")
        }));
    }

    #[test]
    fn smartart_horizontal_process_keeps_layout_and_custom_geometry() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:dgm="dgm" xmlns:r="r"><p:cSld><p:spTree>
          <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/></p:xfrm>
            <a:graphic><a:graphicData uri="diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let relationships = br#"<Relationships><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>"#;
        let data = br#"<dgm:dataModel xmlns:dgm="dgm" xmlns:a="a"><dgm:ptLst>
          <dgm:pt modelId="doc" type="doc"><dgm:prSet loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/hProcess7#1"/></dgm:pt>
          <dgm:pt modelId="p1"><dgm:spPr><a:custGeom/><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></dgm:spPr><dgm:t><a:p><a:r><a:t>P1</a:t></a:r></a:p></dgm:t></dgm:pt>
          <dgm:pt modelId="p2"><dgm:spPr><a:effectLst><a:outerShdw blurRad="19050" dist="9525" dir="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst></dgm:spPr><dgm:t><a:p><a:r><a:t>P2</a:t></a:r></a:p></dgm:t></dgm:pt>
        </dgm:ptLst><dgm:cxnLst>
          <dgm:cxn type="parOf" srcId="doc" destId="p1"/><dgm:cxn type="parOf" srcId="doc" destId="p2"/>
        </dgm:cxnLst><dgm:bg><dgm:spPr><a:solidFill><a:srgbClr val="102030"/></a:solidFill></dgm:spPr></dgm:bg></dgm:dataModel>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", relationships),
            ("ppt/diagrams/data1.xml", data),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("SmartArt data fallback parses");
        let first = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("P1"))
            .expect("first process node");
        let second = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("P2"))
            .expect("second process node");

        assert_eq!(first.kind, ObjectKind::TextBox);
        assert!(second.bounds.x > first.bounds.x);
        assert!(matches!(
            &first.visual,
            Visual::Layer { visual, .. }
                if matches!(visual.as_ref(), Visual::TextLayout { layout, .. }
                    if layout.orientation == TextOrientation::Rotated270)
        ));
        assert!(objects.iter().any(|object| matches!(
            &object.visual,
            Visual::Layer { visual, .. }
                if matches!(visual.as_ref(), Visual::Effect { shadow: Some(_), .. })
        )));
        assert!(objects.iter().any(|object| {
            matches!(
                object.visual,
                Visual::PaintedShape {
                    fill: Paint::Solid(0x1020_30ff),
                    ..
                }
            )
        }));
        assert!(objects.iter().any(|object| matches!(
            &object.visual,
            Visual::Layer { visual, .. }
                if matches!(visual.as_ref(), Visual::TextLayout { visual, .. }
                    if matches!(visual.as_ref(), Visual::RichText {
                        geometry: Geometry::Path { .. },
                        fill: Paint::Solid(_),
                        ..
                    }))
        )));
        let connector = objects
            .iter()
            .find(|object| {
                matches!(
                    &object.visual,
                    Visual::Layer { visual, .. }
                        if matches!(visual.as_ref(), Visual::PaintedShape {
                            geometry: Geometry::Path { .. },
                            fill: Paint::Solid(0xffff_ffff),
                            stroke: Paint::Solid(_),
                            ..
                        })
                )
            })
            .expect("process connector");
        assert!(connector.z > second.z);
    }

    #[test]
    fn smartart_horizontal_list_reuses_shared_role_layout() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:dgm="dgm" xmlns:r="r"><p:cSld><p:spTree>
          <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="5715000" cy="3810000"/></p:xfrm>
            <a:graphic><a:graphicData uri="diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let relationships = br#"<Relationships><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>"#;
        let data = br#"<dgm:dataModel xmlns:dgm="dgm" xmlns:a="a"><dgm:ptLst>
          <dgm:pt modelId="doc" type="doc"><dgm:prSet loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/hList7#1"/></dgm:pt>
          <dgm:pt modelId="p1"><dgm:prSet phldr="1"/><dgm:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></dgm:spPr></dgm:pt>
          <dgm:pt modelId="p2"><dgm:prSet phldr="1"/><dgm:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></dgm:spPr></dgm:pt>
        </dgm:ptLst></dgm:dataModel>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", relationships),
            ("ppt/diagrams/data1.xml", data),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();

        assert_eq!(
            objects
                .iter()
                .filter(|object| object.text.as_deref() == Some("[Text]"))
                .count(),
            2,
        );
        assert_eq!(
            objects
                .iter()
                .filter(|object| matches!(
                    object.visual,
                    Visual::PaintedShape {
                        geometry: Geometry::Ellipse,
                        ..
                    }
                ))
                .count(),
            2,
        );
        assert!(objects.iter().any(|object| {
            matches!(
                object.visual,
                Visual::PaintedShape {
                    geometry: Geometry::Path { .. },
                    ..
                }
            )
        }));
    }

    #[test]
    fn smartart_picture_list_keeps_placeholder_shapes_and_text_region_geometry() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:dgm="dgm" xmlns:r="r"><p:cSld><p:spTree>
          <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="1524000" y="1397000"/><a:ext cx="7010400" cy="4927600"/></p:xfrm>
            <a:graphic><a:graphicData uri="diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let relationships = br#"<Relationships><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>"#;
        let data = br#"<dgm:dataModel xmlns:dgm="dgm" xmlns:a="a"><dgm:ptLst>
          <dgm:pt modelId="doc" type="doc"><dgm:prSet loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/pList1#1"/></dgm:pt>
          <dgm:pt modelId="p1"><dgm:prSet phldr="1"/><dgm:spPr><a:noFill/></dgm:spPr></dgm:pt>
          <dgm:pt modelId="p2"><dgm:prSet phldr="1"/></dgm:pt>
          <dgm:pt modelId="p3"><dgm:prSet phldr="1"/></dgm:pt>
          <dgm:pt modelId="p4"><dgm:spPr><a:custGeom/><a:solidFill><a:schemeClr val="accent6"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill></dgm:spPr><dgm:t><a:p><a:r><a:t>Blah</a:t></a:r></a:p></dgm:t></dgm:pt>
          <dgm:pt modelId="shape1" type="pres"><dgm:prSet presAssocID="p1" presName="pictRect"/><dgm:spPr><a:prstGeom prst="smileyFace"/></dgm:spPr></dgm:pt>
        </dgm:ptLst><dgm:cxnLst>
          <dgm:cxn srcId="doc" destId="p1"/><dgm:cxn srcId="doc" destId="p2"/>
          <dgm:cxn srcId="doc" destId="p3"/><dgm:cxn srcId="doc" destId="p4"/>
        </dgm:cxnLst></dgm:dataModel>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", relationships),
            ("ppt/diagrams/data1.xml", data),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();
        let nodes = objects
            .iter()
            .filter(|object| object.source.part == "ppt/diagrams/data1.xml")
            .collect::<Vec<_>>();

        assert_eq!(nodes.len(), 5);
        assert!(matches!(
            &nodes[0].visual,
            Visual::TextLayout { visual, .. }
                if matches!(visual.as_ref(), Visual::RichText {
                    geometry: Geometry::Path { .. } | Geometry::LayeredPath { .. },
                    ..
                })
        ));
        assert_eq!(nodes[3].bounds.height, nodes[2].bounds.height);
        assert!(nodes[4].bounds.y > nodes[3].bounds.y);
        assert!(nodes[4].bounds.height < nodes[3].bounds.height);
    }

    #[test]
    fn smartart_cycle_uses_arrow_segments_with_shadow() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:dgm="dgm" xmlns:r="r"><p:cSld><p:spTree>
          <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="3810000"/></p:xfrm>
            <a:graphic><a:graphicData uri="diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let relationships = br#"<Relationships><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>"#;
        let data = br#"<dgm:dataModel xmlns:dgm="dgm" xmlns:a="a"><dgm:ptLst>
          <dgm:pt modelId="doc" type="doc"><dgm:prSet loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/cycle8#1"/></dgm:pt>
          <dgm:pt modelId="p1"><dgm:t><a:p><a:r><a:t>Meetings</a:t></a:r></a:p></dgm:t></dgm:pt>
          <dgm:pt modelId="p2"><dgm:t><a:p><a:r><a:t>Phone Calls</a:t></a:r></a:p></dgm:t></dgm:pt>
          <dgm:pt modelId="p3"><dgm:t><a:p><a:r><a:t>E-mail</a:t></a:r></a:p></dgm:t></dgm:pt>
          <dgm:pt modelId="p4"><dgm:t><a:p><a:r><a:t>Documents</a:t></a:r></a:p></dgm:t></dgm:pt>
        </dgm:ptLst><dgm:cxnLst>
          <dgm:cxn srcId="doc" destId="p1"/><dgm:cxn srcId="doc" destId="p2"/>
          <dgm:cxn srcId="doc" destId="p3"/><dgm:cxn srcId="doc" destId="p4"/>
        </dgm:cxnLst></dgm:dataModel>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", relationships),
            ("ppt/diagrams/data1.xml", data),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();

        assert!(
            objects
                .iter()
                .any(|object| object.text.as_deref() == Some("Meetings"))
        );
        assert!(matches!(
            &objects.iter().find(|object| object.text.is_none() && matches!(object.visual, Visual::Effect { .. })).unwrap().visual,
            Visual::Effect { shadow: Some(_), visual, .. }
                if matches!(visual.as_ref(), Visual::PaintedShape {
                    geometry: Geometry::Path { commands, .. }, ..
                } if commands.len() > 6)
        ));
    }

    #[test]
    fn smartart_org_chart_does_not_invent_unstored_shadows() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:dgm="dgm" xmlns:r="r"><p:cSld><p:spTree>
          <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/></p:xfrm>
            <a:graphic><a:graphicData uri="diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let relationships = br#"<Relationships><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>"#;
        let data = br#"<dgm:dataModel xmlns:dgm="dgm" xmlns:a="a"><dgm:ptLst>
          <dgm:pt modelId="doc" type="doc"><dgm:prSet loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/orgChart1#1"/></dgm:pt>
          <dgm:pt modelId="fruit"><dgm:t><a:p><a:r><a:t>Fruit</a:t></a:r></a:p></dgm:t></dgm:pt>
        </dgm:ptLst><dgm:cxnLst><dgm:cxn srcId="doc" destId="fruit"/></dgm:cxnLst></dgm:dataModel>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", relationships),
            ("ppt/diagrams/data1.xml", data),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();

        let fruit = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Fruit"))
            .unwrap();
        assert!(matches!(fruit.visual, Visual::TextLayout { .. }));
    }

    #[test]
    fn preserves_embedded_video_with_its_authored_poster() {
        const MP4: &[u8] = b"\0\0\0\x18ftypisom\0\0\0\0";
        let slide = slide_xml().replace(
            "<p:nvPicPr><p:cNvPr id=\"42\"/></p:nvPicPr>",
            "<p:nvPicPr><p:cNvPr id=\"42\"/><p:nvPr><a:videoFile r:link=\"rIdMedia\"/></p:nvPr></p:nvPicPr>",
        );
        let relationships = r#"<Relationships>
          <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
          <Relationship Id="rIdMedia" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/media1.mp4"/>
        </Relationships>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", relationships.as_bytes()),
            ("ppt/media/image1.png", PNG),
            ("ppt/media/media1.mp4", MP4),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics).expect("valid embedded video");

        assert!(matches!(
            &objects[0].visual,
            Visual::Media {
                kind: MediaKind::Video,
                media_type,
                bytes,
                poster,
            } if media_type == "video/mp4"
                && bytes == MP4
                && matches!(poster.as_ref(), Visual::Image { media_type, bytes, .. }
                    if media_type == "image/png" && bytes == PNG)
        ));
        assert!(diagnostics.is_empty(), "{diagnostics:#?}");
    }

    #[test]
    fn preserves_picture_shape_geometry_and_outline() {
        let slide = slide_xml().replace(
            "<a:extLst><a:ext uri=\"extension\"/></a:extLst>",
            "<a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom><a:ln w=\"28575\"><a:solidFill><a:srgbClr val=\"D7D743\"/></a:solidFill><a:round/></a:ln>",
        );
        let relationships = relationships_xml("../media/image1.png", None);
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", relationships.as_bytes()),
            ("ppt/media/image1.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics)
            .expect("valid elliptical picture");

        assert!(
            matches!(
                &objects[0].visual,
                Visual::StrokeStyle {
                    style: StrokeStyle { join: LineJoin::Round, .. },
                    visual,
                } if matches!(
                visual.as_ref(),
                Visual::PaintedShape {
                    geometry: Geometry::Path { commands, .. },
                    fill: Paint::Image { media_type, bytes, .. },
                    stroke: Paint::Solid(0xd7d7_43ff),
                    stroke_width,
                } if media_type == "image/png"
                    && bytes == PNG
                    && matches!(commands.last(), Some(PathCommand::ClosePath))
                    && (*stroke_width - 3.0).abs() < 0.001
            )
            ),
            "{:#?}",
            objects[0].visual
        );
        assert!(diagnostics.is_empty());
    }

    #[test]
    fn preserves_drawingml_picture_color_change_for_rendering() {
        let slide = slide_xml().replace(
            "<a:blip r:embed=\"rIdImage\"/>",
            "<a:blip r:embed=\"rIdImage\"><a:clrChange><a:clrFrom><a:srgbClr val=\"FFFFFF\"/></a:clrFrom><a:clrTo><a:srgbClr val=\"FFFFFF\"><a:alpha val=\"0\"/></a:srgbClr></a:clrTo></a:clrChange></a:blip>",
        );
        let relationships = relationships_xml("../media/image1.png", None);
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", relationships.as_bytes()),
            ("ppt/media/image1.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics)
            .expect("valid DrawingML picture color change");

        assert!(matches!(
            &objects[0].visual,
            Visual::ImageColorChange {
                from: 0xffff_ffff,
                to: 0xffff_ff00,
                use_alpha: true,
                visual,
            } if matches!(visual.as_ref(), Visual::Image { media_type, bytes, .. }
                if media_type == "image/png" && bytes == PNG)
        ));
        assert!(diagnostics.is_empty());
    }

    #[test]
    fn preserves_drawingml_picture_fixed_alpha_for_rendering() {
        let slide = slide_xml().replace(
            "<a:blip r:embed=\"rIdImage\"/>",
            "<a:blip r:embed=\"rIdImage\"><a:alphaModFix amt=\"4000\"/></a:blip>",
        );
        let relationships = relationships_xml("../media/image1.png", None);
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", relationships.as_bytes()),
            ("ppt/media/image1.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics)
            .expect("valid DrawingML picture fixed alpha");

        assert!(matches!(
            &objects[0].visual,
            Visual::Layer {
                opacity,
                visual,
                ..
            } if (*opacity - 0.04).abs() < f32::EPSILON
                && matches!(visual.as_ref(), Visual::Image { media_type, bytes, .. }
                    if media_type == "image/png" && bytes == PNG)
        ));
        assert!(diagnostics.is_empty());
    }

    #[test]
    fn preserves_explicit_drawingml_picture_color_change_alpha_matching() {
        let slide = slide_xml().replace(
            "<a:blip r:embed=\"rIdImage\"/>",
            "<a:blip r:embed=\"rIdImage\"><a:clrChange useA=\"false\"><a:clrFrom><a:srgbClr val=\"FFFFFF\"/></a:clrFrom><a:clrTo><a:srgbClr val=\"336699\"><a:alpha val=\"50000\"/></a:srgbClr></a:clrTo></a:clrChange></a:blip>",
        );
        let relationships = relationships_xml("../media/image1.png", None);
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", relationships.as_bytes()),
            ("ppt/media/image1.png", PNG),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics)
            .expect("valid DrawingML picture color change with explicit useA");

        assert!(matches!(
            &objects[0].visual,
            Visual::ImageColorChange {
                from: 0xffff_ffff,
                to: 0x3366_9980,
                use_alpha: false,
                ..
            }
        ));
        assert!(diagnostics.is_empty());
    }

    #[test]
    fn blocks_external_pictures_without_loading_them() {
        let bytes =
            picture_package_bytes("https://example.invalid/image.png", Some("External"), b"");
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics)
            .expect("blocked resources are non-fatal");

        assert!(objects.is_empty());
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(diagnostics[0].code, DiagnosticCode::ExternalResourceBlocked);
    }

    #[test]
    fn parses_an_embedded_svg_picture() {
        let bytes = picture_package_bytes("../media/image1.svg", None, b"<svg/>");
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics).expect("valid SVG picture");

        assert_eq!(objects.len(), 1);
        assert!(matches!(
            &objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/svg+xml" && bytes == b"<svg/>"
        ));
        assert!(diagnostics.is_empty());
    }

    #[test]
    fn prefers_the_office_svg_extension_over_its_raster_fallback() {
        let bytes = svg_preferred_picture_package_bytes(true, b"<svg/>");
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics)
            .expect("valid preferred SVG picture");

        assert_eq!(objects.len(), 1);
        assert!(matches!(
            &objects[0].visual,
            Visual::ImageWithFallback {
                media_type,
                bytes,
                fallback_media_type,
                fallback_bytes,
                ..
            }
                if media_type == "image/svg+xml"
                    && bytes == b"<svg/>"
                    && fallback_media_type == "image/png"
                    && fallback_bytes == PNG
        ));
        assert!(diagnostics.is_empty());
    }

    #[test]
    fn uses_the_raster_fallback_when_the_preferred_svg_relationship_is_missing() {
        let bytes = svg_preferred_picture_package_bytes(false, b"<svg/>");
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics)
            .expect("missing preferred SVG relationship has a raster fallback");

        assert_eq!(objects.len(), 1);
        assert!(matches!(
            &objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/png" && bytes == PNG
        ));
        assert_eq!(diagnostics.len(), 1);
        assert!(diagnostics[0].message.contains("preferred SVG"));
        assert!(diagnostics[0].message.contains("fallback"));
    }

    #[test]
    fn uses_the_raster_fallback_when_the_preferred_svg_bytes_are_invalid() {
        let bytes = svg_preferred_picture_package_bytes(true, b"not svg");
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics)
            .expect("invalid preferred SVG bytes have a raster fallback");

        assert_eq!(objects.len(), 1);
        assert!(matches!(
            &objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/png" && bytes == PNG
        ));
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(diagnostics[0].fidelity, Fidelity::Approximate);
        assert!(diagnostics[0].message.contains("valid SVG root"));
        assert!(diagnostics[0].message.contains("fallback"));
    }

    #[test]
    fn explicitly_blocks_images_disabled_by_current_office() {
        let bytes = picture_package_bytes("../media/image1.eps", None, b"%!PS-Adobe");
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics)
            .expect("Office-disabled image is non-fatal");

        assert!(objects.is_empty());
        assert_eq!(diagnostics.len(), 1);
        assert_eq!(diagnostics[0].code, DiagnosticCode::UnsupportedFeature);
        assert_eq!(diagnostics[0].phase, Phase::Security);
        assert_eq!(diagnostics[0].fidelity, Fidelity::Blocked);
        assert!(diagnostics[0].message.starts_with("DisabledByOffice:"));
    }

    #[test]
    fn applies_the_object_limit_to_pictures() {
        let slide = slide_xml().replace(
            "</p:spTree>",
            &format!("{}</p:spTree>", picture_xml(43, "rIdImage")),
        );
        let rels = relationships_xml("../media/image1.png", None);
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", rels.as_bytes()),
            ("ppt/media/image1.png", PNG),
        ]);
        let package = Package::open(
            &bytes,
            Limits {
                max_document_objects: 1,
                ..Limits::default()
            },
        )
        .unwrap();
        let error = parse_test_slide(&package, &mut Vec::new(), &mut Vec::new())
            .expect_err("second picture exceeds the object budget");

        assert_eq!(error.code, DiagnosticCode::ObjectLimit);
    }

    #[test]
    fn bounds_materialized_bytes_when_a_part_is_reused() {
        let slide = slide_xml().replace(
            "</p:spTree>",
            &format!("{}</p:spTree>", picture_xml(43, "rIdImage")),
        );
        let rels = relationships_xml("../media/image1.png", None);
        let mut image = vec![0_u8; 4_096];
        image[..PNG.len()].copy_from_slice(PNG);
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", rels.as_bytes()),
            ("ppt/media/image1.png", &image),
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
        let error = parse_test_slide(&package, &mut Vec::new(), &mut Vec::new())
            .expect_err("reusing a part must not bypass the materialized-byte budget");

        assert_eq!(error.code, DiagnosticCode::ZipTotalSizeLimit);
    }

    #[test]
    fn diagnoses_remaining_unsupported_slide_objects_once_per_part() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:c="c"><p:cSld><p:spTree>
          <p:graphicFrame><a:graphic><a:graphicData><a:tbl/><a:tbl/></a:graphicData></a:graphic></p:graphicFrame>
          <p:graphicFrame><a:graphic><a:graphicData><c:chart/><c:chart/></a:graphicData></a:graphic></p:graphicFrame>
          <p:grpSp/><p:grpSp/>
          <p:sp><p:nvSpPr><p:cNvPr id="7"/></p:nvSpPr><p:spPr><a:prstGeom prst="hexagon"/><a:custGeom/></p:spPr></p:sp>
          <p:sp><p:nvSpPr><p:cNvPr id="8"/></p:nvSpPr><p:spPr><a:prstGeom prst="hexagon"/></p:spPr></p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut Vec::new(), &mut diagnostics)
            .expect("unsupported slide objects are diagnosed non-fatally");

        let unsupported = diagnostics
            .iter()
            .filter(|diagnostic| {
                diagnostic.code == DiagnosticCode::UnsupportedFeature
                    && diagnostic.phase != Phase::Parse
            })
            .collect::<Vec<_>>();
        assert_eq!(
            unsupported
                .iter()
                .filter(|diagnostic| diagnostic.message.to_ascii_lowercase().contains("chart"))
                .count(),
            1,
            "the remaining chart fallback diagnostic must be deduplicated"
        );
        assert!(
            unsupported
                .iter()
                .all(|diagnostic| diagnostic.message.to_ascii_lowercase().contains("chart"))
        );
        assert!(
            unsupported
                .iter()
                .all(|diagnostic| { diagnostic.location.part.as_deref() == Some(SLIDE_PART) })
        );
    }

    #[test]
    fn pptx_table_style_maps_outer_and_inside_borders_per_edge() {
        let style_id = "{0505E3EF-67EA-436B-97B2-0124C06EBD24}";
        let table_styles = format!(
            r#"<a:tblStyleLst xmlns:a="a"><a:tblStyle styleId="{style_id}">
              <a:wholeTbl><a:tcStyle><a:tcBdr>
                <a:left><a:ln w="9525"><a:solidFill><a:srgbClr val="110000"/></a:solidFill></a:ln></a:left>
                <a:right><a:ln w="19050"><a:solidFill><a:srgbClr val="220000"/></a:solidFill></a:ln></a:right>
                <a:top><a:ln w="28575"><a:solidFill><a:srgbClr val="330000"/></a:solidFill></a:ln></a:top>
                <a:bottom><a:ln w="38100"><a:solidFill><a:srgbClr val="440000"/></a:solidFill></a:ln></a:bottom>
                <a:insideH><a:ln w="47625"><a:solidFill><a:srgbClr val="550000"/></a:solidFill></a:ln></a:insideH>
                <a:insideV><a:ln w="57150"><a:solidFill><a:srgbClr val="660000"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></a:insideV>
              </a:tcBdr></a:tcStyle></a:wholeTbl>
            </a:tblStyle></a:tblStyleLst>"#
        );
        let bytes = stored_zip(&[(super::TABLE_STYLES_PART, table_styles.as_bytes())]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let style =
            parse_pptx_table_style(&package, style_id, &PptxTheme::default(), Rect::default())
                .expect("table style parses")
                .expect("referenced table style is found");
        let context = |row, column| TableCellStyleContext {
            row,
            column,
            row_span: 1,
            column_span: 1,
            row_count: 2,
            column_count: 2,
            first_row: false,
            last_row: false,
            first_column: false,
            last_column: false,
            band_rows: false,
            band_columns: false,
        };

        for (side, row, column, color, width) in [
            (TableBorderSide::Left, 0, 0, 0x1100_00ff, 1.0),
            (TableBorderSide::Top, 0, 0, 0x3300_00ff, 3.0),
            (TableBorderSide::Right, 0, 0, 0x6600_00ff, 6.0),
            (TableBorderSide::Bottom, 0, 0, 0x5500_00ff, 5.0),
            (TableBorderSide::Right, 1, 1, 0x2200_00ff, 2.0),
            (TableBorderSide::Bottom, 1, 1, 0x4400_00ff, 4.0),
        ] {
            let border = style.cell_border(side, context(row, column));
            assert_eq!(border.paint, Paint::Solid(color));
            assert!((border.width - width).abs() < 0.001);
        }
        assert_eq!(
            style
                .cell_border(TableBorderSide::Right, context(0, 0))
                .dash_pattern,
            DrawingMlDashPattern::SystemDash
        );
    }

    #[test]
    fn pptx_table_style_resolves_theme_line_references() {
        let style_id = "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}";
        let table_styles = format!(
            r#"<a:tblStyleLst xmlns:a="a"><a:tblStyle styleId="{style_id}">
              <a:wholeTbl><a:tcStyle><a:tcBdr><a:left><a:lnRef idx="2">
                <a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr>
              </a:lnRef></a:left></a:tcBdr></a:tcStyle></a:wholeTbl>
            </a:tblStyle></a:tblStyleLst>"#
        );
        let bytes = stored_zip(&[(super::TABLE_STYLES_PART, table_styles.as_bytes())]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let style =
            parse_pptx_table_style(&package, style_id, &PptxTheme::default(), Rect::default())
                .expect("table style parses")
                .expect("referenced table style is found");
        let border = style.parts[&super::TableStyleRegion::WholeTable].borders
            [TableStyleBorderSide::Left.index()]
        .as_ref()
        .unwrap();

        assert_eq!(border.paint, Paint::Solid(0xc0c9_e4ff));
        assert!((border.width - 8.0 / 3.0).abs() < 0.001);
    }

    #[test]
    fn pptx_table_style_resolves_theme_background_fill_reference() {
        const THEME_PART: &str = "ppt/theme/theme1.xml";
        let style_id = "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}";
        let table_styles = format!(
            r#"<a:tblStyleLst xmlns:a="a"><a:tblStyle styleId="{style_id}">
              <a:tblBg><a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef></a:tblBg>
            </a:tblStyle></a:tblStyleLst>"#
        );
        let bytes = stored_zip(&[
            (super::TABLE_STYLES_PART, table_styles.as_bytes()),
            (
                THEME_PART,
                br#"<a:theme xmlns:a="a"><a:themeElements>
                  <a:clrScheme><a:accent1><a:srgbClr val="5B9BD5"/></a:accent1></a:clrScheme>
                  <a:fmtScheme><a:fillStyleLst>
                    <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                    <a:gradFill><a:gsLst>
                      <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs>
                      <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/></a:schemeClr></a:gs>
                      <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/></a:schemeClr></a:gs>
                    </a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>
                  </a:fillStyleLst></a:fmtScheme>
                </a:themeElements></a:theme>"#,
            ),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let theme = parse_pptx_theme(&package, THEME_PART).unwrap();
        let style = parse_pptx_table_style(
            &package,
            style_id,
            &theme,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 100.0,
                height: 60.0,
            },
        )
        .expect("table style parses")
        .expect("referenced table style is found");

        let Some(Paint::LinearGradient { stops, .. }) = style.background_fill.map(|fill| {
            fill.paint(Rect {
                width: 100.0,
                height: 60.0,
                ..Rect::default()
            })
        }) else {
            panic!("table background must resolve the referenced theme gradient");
        };
        assert_eq!(stops.len(), 3);
        assert_eq!(
            stops.iter().map(|stop| stop.offset).collect::<Vec<_>>(),
            [0.0, 0.35, 1.0]
        );
        assert_ne!(stops[0].color, stops[2].color);
    }

    #[test]
    fn pptx_table_edges_keep_independent_styles_and_collapse_shared_borders() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="7"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
            <a:graphic><a:graphicData><a:tbl>
              <a:tblGrid><a:gridCol w="952500"/><a:gridCol w="952500"/></a:tblGrid>
              <a:tr h="952500">
                <a:tc><a:txBody><a:p/></a:txBody><a:tcPr>
                  <a:lnL w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnL>
                  <a:lnR w="38100"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:prstDash val="dash"/></a:lnR>
                  <a:lnT><a:noFill/></a:lnT>
                  <a:lnB w="28575"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:lnB>
                </a:tcPr></a:tc>
                <a:tc><a:txBody><a:p/></a:txBody><a:tcPr>
                  <a:lnL w="9525"><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:lnL>
                  <a:lnR><a:noFill/></a:lnR>
                  <a:lnT w="19050"><a:solidFill><a:srgbClr val="800080"/></a:solidFill></a:lnT>
                  <a:lnB><a:noFill/></a:lnB>
                </a:tcPr></a:tc>
              </a:tr>
            </a:tbl></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("independent table borders produce a valid slide");

        let cells = objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Cell)
            .collect::<Vec<_>>();
        assert_eq!(cells.len(), 2);
        for cell in cells {
            let Visual::PaintedShape {
                stroke,
                stroke_width,
                ..
            } = &cell.visual
            else {
                panic!("empty table cells use painted shapes");
            };
            assert_eq!(*stroke, Paint::None);
            assert_eq!(*stroke_width, 0.0);
        }

        let borders = objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Shape)
            .collect::<Vec<_>>();
        assert_eq!(
            borders.len(),
            4,
            "noFill edges are omitted and the shared edge is emitted once"
        );

        let shared = borders
            .iter()
            .find(|object| {
                (object.bounds.x - 98.0).abs() < 0.001
                    && (object.bounds.width - 4.0).abs() < 0.001
                    && (object.bounds.height - 100.0).abs() < 0.001
            })
            .expect("the shared vertical edge is present");
        let Visual::StrokeStyle { style, visual } = &shared.visual else {
            panic!("the shared dashed edge keeps its stroke style");
        };
        let Visual::PaintedShape {
            geometry,
            stroke,
            stroke_width,
            ..
        } = visual.as_ref()
        else {
            panic!("the shared edge uses line geometry");
        };
        assert_eq!(*stroke, Paint::Solid(0x0000_ffff));
        assert!((*stroke_width - 4.0).abs() < 0.001);
        assert_eq!(style.dash, vec![16.0, 12.0]);
        assert_eq!(
            *geometry,
            Geometry::Path {
                fill_rule: crate::model::FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo { x: 2.0, y: 0.0 },
                    PathCommand::LineTo { x: 2.0, y: 100.0 },
                ],
            }
        );
    }

    #[test]
    fn browser_font_metrics_change_pptx_table_cell_bounds() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="7"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="762000" cy="952500"/></p:xfrm>
            <a:graphic><a:graphicData><a:tbl>
              <a:tblGrid><a:gridCol w="762000"/></a:tblGrid>
              <a:tr h="0"><a:tc><a:txBody><a:p><a:pPr latinLnBrk="1"/><a:r>
                <a:rPr sz="1200" b="1" i="1" spc="100"><a:latin typeface="P0 Sans"/></a:rPr>
                <a:t>WWWWWWWW</a:t>
              </a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
              <a:tr h="0"><a:tc><a:txBody><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>
            </a:tbl></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let narrow = single_font_metric("P0 Sans", 'W', 0.3, true, true);
        let wide = single_font_metric("P0 Sans", 'W', 1.0, true, true);
        let mut narrow_objects = Vec::new();
        let mut wide_objects = Vec::new();

        parse_test_slide_with_metrics(&package, &mut narrow_objects, &mut Vec::new(), &narrow)
            .expect("narrow browser metrics produce a valid PPTX table");
        parse_test_slide_with_metrics(&package, &mut wide_objects, &mut Vec::new(), &wide)
            .expect("wide browser metrics produce a valid PPTX table");

        let narrow_cell = narrow_objects
            .iter()
            .find(|object| object.kind == ObjectKind::Cell)
            .expect("narrow layout has a table cell");
        let wide_cell = wide_objects
            .iter()
            .find(|object| object.kind == ObjectKind::Cell)
            .expect("wide layout has a table cell");
        assert!(
            wide_cell.bounds.height > narrow_cell.bounds.height,
            "wide measured advances must create more wrapped lines"
        );
    }

    #[test]
    fn pptx_table_row_height_accounts_for_emergency_latin_word_breaks() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="7"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="762000" cy="952500"/></p:xfrm>
            <a:graphic><a:graphicData><a:tbl>
              <a:tblGrid><a:gridCol w="762000"/></a:tblGrid>
              <a:tr h="0"><a:tc><a:txBody><a:p><a:r>
                <a:rPr sz="1200" b="1" i="1" spc="100"><a:latin typeface="P0 Sans"/></a:rPr>
                <a:t>WWWWWWWW</a:t>
              </a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
              <a:tr h="0"><a:tc><a:txBody><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>
            </a:tbl></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let narrow = single_font_metric("P0 Sans", 'W', 0.3, true, true);
        let wide = single_font_metric("P0 Sans", 'W', 1.0, true, true);
        let mut narrow_objects = Vec::new();
        let mut wide_objects = Vec::new();

        parse_test_slide_with_metrics(&package, &mut narrow_objects, &mut Vec::new(), &narrow)
            .expect("narrow browser metrics produce a valid PPTX table");
        parse_test_slide_with_metrics(&package, &mut wide_objects, &mut Vec::new(), &wide)
            .expect("wide browser metrics produce a valid PPTX table");

        let cell_height = |objects: &[crate::model::Object]| {
            objects
                .iter()
                .find(|object| object.kind == ObjectKind::Cell)
                .expect("layout has a table cell")
                .bounds
                .height
        };
        assert!(
            cell_height(&wide_objects) > cell_height(&narrow_objects) + 15.0,
            "latinLnBrk=false still wraps words wider than the cell body"
        );
    }

    #[test]
    fn pptx_table_height_collapses_an_overflowing_space_at_a_soft_wrap_boundary() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="7"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="219075" cy="952500"/></p:xfrm>
            <a:graphic><a:graphicData><a:tbl>
              <a:tblGrid><a:gridCol w="219075"/></a:tblGrid>
              <a:tr h="0"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>
                <a:rPr sz="750"/><a:t>AAAA BBBB</a:t>
              </a:r></a:p></a:txBody>
              <a:tcPr marL="0" marR="0" marT="0" marB="0"/></a:tc></a:tr>
              <a:tr h="0"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody>
              <a:tcPr marL="0" marR="0" marT="0" marB="0"/></a:tc></a:tr>
            </a:tbl></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("soft-wrapped table text produces a valid slide");

        let cells = objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Cell)
            .collect::<Vec<_>>();
        assert_eq!(cells.len(), 2);
        assert!((cells[0].bounds.height - 62.0).abs() < 0.001);
        assert!((cells[1].bounds.height - 38.0).abs() < 0.001);
    }

    #[test]
    fn table_paragraph_run_defaults_remain_scoped_to_their_list_level() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="8"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/></p:xfrm>
            <a:graphic><a:graphicData><a:tbl>
              <a:tblGrid><a:gridCol w="3810000"/></a:tblGrid>
              <a:tr h="1905000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle>
                <a:lvl1pPr><a:defRPr sz="1200" b="1" i="1" u="sng" strike="sngStrike" baseline="10000" spc="100">
                  <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                  <a:latin typeface="Level One"/>
                </a:defRPr></a:lvl1pPr>
                <a:lvl2pPr><a:defRPr sz="2000" b="0" i="0" u="none" strike="noStrike" baseline="0" spc="200">
                  <a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
                  <a:latin typeface="Level Two"/>
                </a:defRPr></a:lvl2pPr>
              </a:lstStyle>
              <a:p><a:pPr lvl="0"/><a:r><a:t>First</a:t></a:r></a:p>
              <a:p><a:pPr lvl="1"/><a:r><a:t>Second</a:t></a:r></a:p>
              </a:txBody><a:tcPr/></a:tc></a:tr>
            </a:tbl></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("list-level table defaults produce a valid slide");

        let cell = objects
            .iter()
            .find(|object| object.kind == ObjectKind::Cell)
            .expect("table cell is present");
        let Visual::TextLayout { visual, .. } = &cell.visual else {
            panic!("table cell text uses a text layout");
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("table cell text uses rich text runs");
        };
        let first = runs
            .iter()
            .find(|run| run.text == "First")
            .expect("first paragraph run is present");
        let second = runs
            .iter()
            .find(|run| run.text == "Second")
            .expect("second paragraph run is present");

        assert_eq!(first.font_family, "Level One");
        assert!((first.font_size - 16.0).abs() < 0.001);
        assert_eq!(first.color, 0xff00_00ff);
        assert!(first.bold && first.italic && first.underline && first.strikethrough);
        assert!((first.baseline_shift - 1.6).abs() < 0.001);
        assert!((first.letter_spacing - 4.0 / 3.0).abs() < 0.001);

        assert_eq!(second.font_family, "Level Two");
        assert!((second.font_size - 80.0 / 3.0).abs() < 0.001);
        assert_eq!(second.color, 0x0000_ffff);
        assert!(!second.bold && !second.italic && !second.underline && !second.strikethrough);
        assert!(second.baseline_shift.abs() < 0.001);
        assert!((second.letter_spacing - 8.0 / 3.0).abs() < 0.001);
    }

    #[test]
    fn text_run_effects_and_outline_do_not_mutate_the_containing_shape() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="9"/></p:nvSpPr>
            <p:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
              <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              <a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
              <a:effectLst><a:prstShdw prst="shdw13" blurRad="12700" dist="12700" dir="0"><a:srgbClr val="000000"/></a:prstShdw></a:effectLst>
            </p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>Shape effect</a:t></a:r></a:p></p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="10"/></p:nvSpPr>
            <p:spPr>
              <a:xfrm><a:off x="1905000" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
              <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
              <a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
            </p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>
              <a:rPr sz="1200">
                <a:ln><a:noFill/></a:ln>
                <a:effectLst>
                  <a:outerShdw blurRad="127000" dist="12700" dir="0" sy="23000" kx="1200000" algn="b"><a:srgbClr val="000000"/></a:outerShdw>
                  <a:innerShdw blurRad="127000" dist="12700" dir="0"><a:srgbClr val="000000"/></a:innerShdw>
                  <a:glow rad="127000"><a:srgbClr val="FFFFFF"/></a:glow>
                  <a:reflection blurRad="127000" dist="12700"/>
                  <a:softEdge rad="127000"/>
                </a:effectLst>
              </a:rPr>
              <a:t>Run effects</a:t>
            </a:r></a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("shape and text-run effects produce a valid slide");

        let shape_effect = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Shape effect"))
            .expect("shape-level effect shape is present");
        assert!(matches!(
            &shape_effect.visual,
            Visual::Effect {
                shadow: Some(_),
                ..
            }
        ));

        let run_effects = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Run effects"))
            .expect("run-level effect shape is present");
        let Visual::TextLayout { visual, .. } = &run_effects.visual else {
            panic!("run effects must not wrap the containing shape in an effect visual");
        };
        let Visual::TextEffects { effects, visual } = visual.as_ref() else {
            panic!("run effects stay aligned with their text runs");
        };
        assert_eq!(effects.len(), 1);
        assert!(effects[0].shadow.is_some());
        assert!(effects[0].inner_shadow.is_some());
        assert!(effects[0].reflection.is_some());
        assert!((effects[0].shadow_scale_y - 0.23).abs() < 0.001);
        assert!((effects[0].shadow_skew_x - 20.0).abs() < 0.001);
        assert_eq!(effects[0].shadow_alignment, 7);
        let Visual::RichText { stroke, .. } = visual.as_ref() else {
            panic!("shape text uses rich text");
        };
        assert_eq!(*stroke, Paint::Solid(0xff00_00ff));
    }

    #[test]
    fn shape_outer_shadow_preserves_drawingml_scale_and_alignment() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp><p:nvSpPr><p:cNvPr id="9"/></p:nvSpPr><p:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="6858000" cy="3429000"/></a:xfrm>
            <a:prstGeom prst="roundRect"/><a:solidFill><a:srgbClr val="4C91CF"/></a:solidFill>
            <a:effectLst><a:outerShdw blurRad="152400" dist="317500" dir="5400000" sx="90000" sy="-19000" rotWithShape="0">
              <a:prstClr val="black"><a:alpha val="15000"/></a:prstClr>
            </a:outerShdw></a:effectLst>
          </p:spPr></p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("transformed shape shadow produces a valid slide");

        let Visual::AdvancedEffect {
            outer_shadow: Some(effect),
            ..
        } = &objects[0].visual
        else {
            panic!("transformed outer shadow is preserved as an advanced effect");
        };
        assert!((effect.scale_x - 0.9).abs() < 0.001);
        assert!((effect.scale_y + 0.19).abs() < 0.001);
        assert_eq!(effect.alignment, 7);
        assert!((effect.shadow.blur - 16.0).abs() < 0.001);
        assert!((effect.shadow.offset_y - 100.0 / 3.0).abs() < 0.001);
        assert_eq!(effect.shadow.color, 0x0000_0026);
    }

    #[test]
    fn text_only_shape_shadow_follows_independent_text_transform() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:dsp="dsp"><p:cSld><p:spTree>
          <dsp:sp><dsp:nvSpPr><dsp:cNvPr id="9"/></dsp:nvSpPr><dsp:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm>
            <a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln>
            <a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000">
              <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>
            </a:outerShdw></a:effectLst>
          </dsp:spPr><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Shadowed text</a:t></a:r></a:p></dsp:txBody>
          <dsp:txXfrm><a:off x="1905000" y="0"/><a:ext cx="1905000" cy="952500"/></dsp:txXfrm></dsp:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("text-only shadow shape parses");

        let Visual::Group { children } = &objects[0].visual else {
            panic!("independent text transform keeps separate shape and text visuals");
        };
        assert!(matches!(
            &children[1].visual,
            Visual::Effect {
                shadow: Some(_),
                visual,
                ..
            } if matches!(visual.as_ref(), Visual::TextLayout { .. })
        ));
    }

    #[test]
    fn drawingml_symbol_fonts_keep_raw_glyph_codes_and_expose_unicode_semantics() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="9"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>
              <a:rPr sz="1200"><a:latin typeface="Main Sans"/><a:ea typeface="CJK Sans"/><a:sym typeface="Symbol"/></a:rPr>
              <a:t>A&#xF0A3;B&#xF0B3;&#x4E2D;</a:t>
            </a:r><a:r>
              <a:rPr sz="1200"><a:latin typeface="Symbol"/></a:rPr><a:t>AB</a:t>
            </a:r></a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("DrawingML symbol runs produce valid Unicode text");

        let shape = objects
            .iter()
            .find(|object| object.kind == ObjectKind::TextBox)
            .expect("text shape is present");
        assert_eq!(shape.text.as_deref(), Some("A≤B≥中ΑΒ"));
        let Visual::TextLayout { visual, .. } = &shape.visual else {
            panic!("shape text uses a text layout");
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("shape text uses rich text runs");
        };
        assert_eq!(
            runs.iter()
                .map(|run| (run.text.as_str(), run.font_family.as_str()))
                .collect::<Vec<_>>(),
            [
                ("A", "Main Sans"),
                ("\u{f0a3}", "Symbol"),
                ("B", "Main Sans"),
                ("\u{f0b3}", "Symbol"),
                ("中", "CJK Sans"),
                ("AB", "Symbol"),
            ]
        );
    }

    #[test]
    fn drawingml_bullets_render_semantic_glyphs_without_legacy_symbol_fonts() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="9"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p>
              <a:pPr><a:buFont typeface="Wingdings"/><a:buChar char="q"/></a:pPr>
              <a:r><a:rPr sz="1200"><a:latin typeface="Main Sans"/></a:rPr><a:t>Shape item</a:t></a:r>
            </a:p></p:txBody>
          </p:sp>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="10"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="952500"/><a:ext cx="3810000" cy="952500"/></p:xfrm>
            <a:graphic><a:graphicData><a:tbl>
              <a:tblGrid><a:gridCol w="3810000"/></a:tblGrid>
              <a:tr h="952500"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>
                <a:pPr><a:buFont typeface="Wingdings"/><a:buChar char="&#xA7;"/></a:pPr>
                <a:r><a:rPr sz="1200"><a:latin typeface="Main Sans"/></a:rPr><a:t>Cell item</a:t></a:r>
              </a:p></a:txBody><a:tcPr/></a:tc></a:tr>
            </a:tbl></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("legacy symbol-font bullets produce valid semantic text");

        for (kind, semantic_text, glyph) in [
            (ObjectKind::TextBox, "❑\tShape item", "❑\t"),
            (ObjectKind::Cell, "▪\tCell item", "▪\t"),
        ] {
            let object = objects
                .iter()
                .find(|object| object.kind == kind)
                .expect("bullet-bearing object is present");
            assert_eq!(object.text.as_deref(), Some(semantic_text));
            let Visual::TextLayout { visual, .. } = &object.visual else {
                panic!("bullet text uses a text layout");
            };
            let Visual::RichText { runs, .. } = visual.as_ref() else {
                panic!("bullet text uses rich text runs");
            };
            assert_eq!(runs[0].text, glyph);
            assert_eq!(runs[0].font_family, "Arial");
        }
    }

    #[test]
    fn drawingml_picture_bullets_materialize_their_embedded_images() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="9"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="504000" y="1769040"/><a:ext cx="9071640" cy="4384440"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>
              <a:p><a:pPr algn="ctr"><a:buBlip><a:blip r:embed="rId1"/></a:buBlip></a:pPr><a:r><a:rPr sz="3200"/><a:t> </a:t></a:r></a:p>
              <a:p><a:pPr algn="ctr"><a:buBlip><a:blip r:embed="rId2"/></a:buBlip></a:pPr><a:r><a:rPr sz="3200"/><a:t> </a:t></a:r></a:p>
              <a:p><a:pPr algn="ctr"><a:buBlip><a:blip r:embed="rId3"/></a:buBlip></a:pPr><a:r><a:rPr sz="3200"/><a:t> </a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            (
                "ppt/slides/_rels/slide1.xml.rels",
                br#"<Relationships>
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bullet1.gif"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bullet2.gif"/>
                  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bullet3.gif"/>
                </Relationships>"#,
            ),
            ("ppt/media/bullet1.gif", b"GIF89aembedded-bullet-1"),
            ("ppt/media/bullet2.gif", b"GIF89aembedded-bullet-2"),
            ("ppt/media/bullet3.gif", b"GIF89aembedded-bullet-3"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("picture bullet produces a valid slide");

        let bullets = objects
            .iter()
            .filter(|object| object.kind == ObjectKind::Image)
            .collect::<Vec<_>>();
        assert_eq!(bullets.len(), 3);
        assert!(bullets.iter().all(|object| matches!(
            &object.visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/gif" && bytes.starts_with(b"GIF89a")
        )));
        assert!(
            bullets
                .windows(2)
                .all(|pair| (pair[1].bounds.y - pair[0].bounds.y - 51.2).abs() < 0.01)
        );
        assert!(bullets.iter().all(|bullet| {
            (bullet.bounds.width - 25.6).abs() < 0.01
                && (bullet.bounds.x + bullet.bounds.width / 2.0 - 529.1123).abs() < 0.01
        }));
    }

    #[test]
    fn shape_autofit_preserves_persisted_transform_bounds() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="9"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1524000" cy="190500"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr><a:spAutoFit/></a:bodyPr><a:lstStyle/>
              <a:p><a:r><a:rPr sz="1200"/><a:t>First paragraph</a:t></a:r></a:p>
              <a:p><a:r><a:rPr sz="1200"/><a:t>Second paragraph</a:t></a:r></a:p>
              <a:p><a:r><a:rPr sz="1200"/><a:t>Third paragraph</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("shape autofit produces valid static slide geometry");

        let shape = objects
            .iter()
            .find(|object| object.kind == ObjectKind::TextBox)
            .expect("text shape is present");
        assert_eq!(
            shape.bounds,
            Rect {
                x: 100.0,
                y: 100.0,
                width: 160.0,
                height: 20.0,
            }
        );
        let Visual::TextLayout { layout, .. } = &shape.visual else {
            panic!("shape text uses a text layout");
        };
        assert_eq!(
            layout.vertical_overflow,
            crate::model::TextVerticalOverflow::Overflow
        );
        assert_eq!(layout.auto_fit, crate::model::TextAutoFit::None);
    }

    #[test]
    fn preserves_drawingml_caps_east_asian_breaks_and_centered_rtl_text() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp>
          <p:nvSpPr><p:cNvPr id="9"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr anchorCtr="1"/><a:lstStyle/><a:p><a:pPr rtl="1" eaLnBrk="0"/><a:r><a:rPr sz="1800" cap="all"/><a:t>Mixed Text</a:t></a:r></a:p></p:txBody>
        </p:sp><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="10"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="952500"/><a:ext cx="1905000" cy="952500"/></p:xfrm><a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="1905000"/></a:tblGrid><a:tr h="952500"><a:tc><a:txBody><a:p><a:pPr rtl="1" eaLnBrk="0"/><a:r><a:rPr cap="all"/><a:t>table text</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();
        let shape = objects
            .iter()
            .find(|object| object.kind == ObjectKind::TextBox)
            .expect("text shape");
        assert_eq!(shape.text.as_deref(), Some("MIXED TEXT"));
        let Visual::TextLayout { layout, visual } = &shape.visual else {
            panic!("shape text layout");
        };
        assert_eq!(layout.direction, crate::model::TextDirection::Rtl);
        let Visual::RichText { align, runs, .. } = visual.as_ref() else {
            panic!("shape rich text");
        };
        assert_eq!(*align, crate::model::TextAlign::Center);
        assert!(runs.iter().all(|run| !run.east_asian_line_breaks));
        assert_eq!(
            runs.iter().map(|run| run.text.as_str()).collect::<String>(),
            "MIXED TEXT"
        );
        let cell = objects
            .iter()
            .find(|object| object.kind == ObjectKind::Cell)
            .expect("table cell");
        assert_eq!(cell.text.as_deref(), Some("TABLE TEXT"));
        let Visual::TextLayout { layout, visual } = &cell.visual else {
            panic!("cell text layout");
        };
        assert_eq!(layout.direction, crate::model::TextDirection::Rtl);
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("cell rich text");
        };
        assert!(runs.iter().all(|run| !run.east_asian_line_breaks));
    }

    #[test]
    fn empty_paragraphs_use_the_end_marker_size_without_rendering_a_bullet() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="9"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:rPr sz="1400"/><a:t>Heading</a:t></a:r></a:p>
              <a:p><a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:buAutoNum type="arabicPeriod" startAt="2"/></a:pPr><a:r><a:rPr sz="100"/><a:t/></a:r><a:endParaRPr sz="100"/></a:p>
              <a:p><a:r><a:rPr sz="1000"/><a:t>Body</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="10"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="952500"/><a:ext cx="3810000" cy="952500"/></p:xfrm>
            <a:graphic><a:graphicData><a:tbl>
              <a:tblGrid><a:gridCol w="3810000"/></a:tblGrid>
              <a:tr h="952500"><a:tc><a:txBody>
                <a:p><a:r><a:rPr sz="1000"/><a:t>Item</a:t></a:r></a:p>
                <a:p><a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:buChar char="&#x2022;"/></a:pPr><a:endParaRPr sz="100"/></a:p>
              </a:txBody><a:tcPr/></a:tc></a:tr>
            </a:tbl></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("empty DrawingML paragraphs produce a valid slide");

        let shape = objects
            .iter()
            .find(|object| object.kind == ObjectKind::TextBox)
            .expect("shape text is present");
        assert_eq!(shape.text.as_deref(), Some("Heading\n\nBody"));
        let Visual::TextLayout { layout, .. } = &shape.visual else {
            panic!("shape text uses a text layout");
        };
        assert_eq!(layout.paragraphs.len(), 3);
        assert!((layout.paragraphs[1].line_height - 1.6).abs() < 0.001);

        let cell = objects
            .iter()
            .find(|object| object.kind == ObjectKind::Cell)
            .expect("table cell is present");
        assert_eq!(cell.text.as_deref(), Some("Item\n"));
        let Visual::TextLayout { layout, .. } = &cell.visual else {
            panic!("table-cell text uses a text layout");
        };
        assert_eq!(layout.paragraphs.len(), 2);
        assert!((layout.paragraphs[1].line_height - 1.6).abs() < 0.001);
    }

    #[test]
    fn bullet_indentation_keeps_shape_and_table_anchors_and_plain_text_boundaries() {
        // DrawingML bullets use a nonnegative marker anchor and an absolute gap;
        // plain paragraphs retain signed first-line indentation.
        let paragraphs = r#"<a:p><a:pPr indent="324000"><a:buChar char="x"/></a:pPr><a:r><a:t>Positive</a:t></a:r></a:p>
          <a:p><a:pPr indent="-324000"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Negative</a:t></a:r></a:p>
          <a:p><a:pPr marL="685800" indent="-324000"><a:buChar char="x"/></a:pPr><a:r><a:t>Hanging</a:t></a:r></a:p>
          <a:p><a:pPr marL="685800" indent="324000"><a:buChar char="x"/></a:pPr><a:r><a:t>Positive margin</a:t></a:r></a:p>
          <a:p><a:pPr marL="228600" indent="324000"><a:buNone/></a:pPr><a:r><a:t>Plain</a:t></a:r></a:p>"#;
        let body = format!(r#"<a:bodyPr/><a:lstStyle/>{paragraphs}"#);
        let slide = format!(
            r#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp><p:nvSpPr><p:cNvPr id="21"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="2857500"/></a:xfrm></p:spPr><p:txBody>{body}</p:txBody></p:sp>
          <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="22"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="2857500"/><a:ext cx="3810000" cy="2857500"/></p:xfrm>
          <a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="3810000"/></a:tblGrid><a:tr h="2857500"><a:tc><a:txBody>{body}</a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
          </p:spTree></p:cSld></p:sld>"#
        );
        let bytes = stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();
        let gap = 324000.0 / super::EMU_PER_CSS_PIXEL;
        let mut checked = 0;
        for object in objects
            .iter()
            .filter(|object| matches!(object.kind, ObjectKind::TextBox | ObjectKind::Cell))
        {
            let Visual::TextLayout { layout, .. } = &object.visual else {
                panic!("text layout");
            };
            for (paragraph, (margin, indent)) in layout.paragraphs.iter().zip([
                (gap, -gap),
                (gap, -gap),
                (72.0, -gap),
                (72.0 + gap, -gap),
                (24.0, gap),
            ]) {
                assert!((paragraph.margin_left - margin).abs() < 0.001);
                assert!((paragraph.first_line_indent - indent).abs() < 0.001);
            }
            assert_eq!(layout.paragraphs.len(), 5);
            checked += 1;
        }
        assert_eq!(checked, 2);
    }

    #[test]
    fn paragraph_line_break_properties_flow_through_shapes_and_table_cells() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="21"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle><a:defPPr latinLnBrk="1" hangingPunct="0"/></a:lstStyle>
              <a:p><a:pPr latinLnBrk="0" hangingPunct="1"/><a:r><a:t>Shape override</a:t></a:r></a:p>
              <a:p><a:r><a:t>Shape default</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="23"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="3810000" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/>
              <a:p><a:r><a:t>Shape omitted defaults</a:t></a:r></a:p>
            </p:txBody>
          </p:sp>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="22"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="952500"/><a:ext cx="3810000" cy="952500"/></p:xfrm>
            <a:graphic><a:graphicData><a:tbl>
              <a:tblGrid><a:gridCol w="3810000"/></a:tblGrid>
              <a:tr h="952500"><a:tc><a:txBody><a:bodyPr/><a:lstStyle><a:defPPr latinLnBrk="1" hangingPunct="0"/></a:lstStyle>
                <a:p><a:r><a:t>Cell paragraph</a:t></a:r></a:p>
              </a:txBody><a:tcPr/></a:tc></a:tr>
              <a:tr h="952500"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/>
                <a:p><a:r><a:t>Cell omitted defaults</a:t></a:r></a:p>
              </a:txBody><a:tcPr/></a:tc></a:tr>
            </a:tbl></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("DrawingML paragraph break properties produce valid text layouts");

        let shape = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Shape override\nShape default"))
            .expect("shape text is present");
        let Visual::TextLayout { layout, .. } = &shape.visual else {
            panic!("shape text uses a text layout");
        };
        assert!(!layout.paragraphs[0].latin_line_break);
        assert!(layout.paragraphs[0].hanging_punctuation);
        assert!(layout.paragraphs[1].latin_line_break);
        assert!(!layout.paragraphs[1].hanging_punctuation);

        let omitted_shape = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Shape omitted defaults"))
            .expect("shape with omitted paragraph properties is present");
        let Visual::TextLayout { layout, .. } = &omitted_shape.visual else {
            panic!("shape text uses a text layout");
        };
        assert!(!layout.paragraphs[0].latin_line_break);
        assert!(layout.paragraphs[0].hanging_punctuation);

        let cell = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Cell paragraph"))
            .expect("table cell text is present");
        let Visual::TextLayout { layout, .. } = &cell.visual else {
            panic!("table cell text uses a text layout");
        };
        assert!(layout.paragraphs[0].latin_line_break);
        assert!(!layout.paragraphs[0].hanging_punctuation);

        let omitted_cell = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("Cell omitted defaults"))
            .expect("table cell with omitted paragraph properties is present");
        let Visual::TextLayout { layout, .. } = &omitted_cell.visual else {
            panic!("table cell text uses a text layout");
        };
        assert!(!layout.paragraphs[0].latin_line_break);
        assert!(layout.paragraphs[0].hanging_punctuation);
    }

    #[test]
    fn presentation_default_text_styles_supply_shape_and_table_run_defaults() {
        let presentation = br#"<p:presentation xmlns:p="p" xmlns:a="a">
          <p:defaultTextStyle>
            <a:defPPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:defPPr>
            <a:lvl1pPr>
              <a:defRPr sz="1400" b="0">
                <a:latin typeface="Arial"/>
                <a:ea typeface="Microsoft YaHei"/>
              </a:defRPr>
            </a:lvl1pPr>
          </p:defaultTextStyle>
        </p:presentation>"#;
        let theme = PptxTheme::default();
        let styles = super::parse_default_text_styles(
            presentation,
            Limits::default(),
            "ppt/presentation.xml",
            &theme,
        )
        .expect("presentation default text styles parse");
        let level = &styles[0];
        assert!((level.font_size.unwrap() - 14.0 * super::POINTS_TO_CSS_PIXELS).abs() < 0.001);
        assert_eq!(level.font_family.as_deref(), Some("Arial"));
        assert_eq!(level.font_east_asian.as_deref(), Some("Microsoft YaHei"));
        assert_eq!(level.bold, Some(false));

        let defaults = super::TableCellRunDefaults::from_presentation_defaults(&theme, Some(level));
        assert!((defaults.font_size - 14.0 * super::POINTS_TO_CSS_PIXELS).abs() < 0.001);
        assert_eq!(defaults.font_family, "Arial");
        assert_eq!(defaults.font_east_asian.as_deref(), Some("Microsoft YaHei"));
        assert!(!defaults.bold);
    }

    #[test]
    fn text_run_hyperlinks_use_the_theme_link_style() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="9"><a:extLst/></p:cNvPr></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>
              <a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>https://example.com/</a:t>
            </a:r></a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("run hyperlink style produces valid rich text");
        let object = objects
            .iter()
            .find(|object| object.text.as_deref() == Some("https://example.com/"))
            .expect("hyperlinked text shape is present");
        let Visual::TextLayout { visual, .. } = &object.visual else {
            panic!("hyperlink uses a text layout");
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("hyperlink uses rich text");
        };
        assert_eq!(runs[0].color, 0x0563_c1ff);
        assert!(runs[0].underline);
    }

    #[test]
    fn omitted_master_paragraph_properties_keep_office_defaults() {
        const MASTER_PART: &str = "ppt/slideMasters/slideMaster1.xml";
        let master = br#"<p:sldMaster xmlns:p="p" xmlns:a="a"><p:txStyles>
          <p:bodyStyle><a:defPPr/><a:lvl1pPr/></p:bodyStyle>
        </p:txStyles></p:sldMaster>"#;
        let bytes = stored_zip(&[(MASTER_PART, master)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let styles = parse_master_text_styles(&package, MASTER_PART, &PptxTheme::default())
            .expect("master paragraph styles parse");
        let mut shape = ShapeState::new(0, &PptxTheme::default(), false);
        shape.placeholder_type = Some("body".to_owned());
        apply_master_text_styles(&mut shape, Some(&styles));
        shape.paragraph_pending = true;

        begin_shape_paragraph(&mut shape, false);

        assert!(!shape.paragraph_layouts[0].latin_line_break);
        assert!(shape.paragraph_layouts[0].hanging_punctuation);
    }

    #[test]
    fn master_title_color_uses_the_layout_color_map() {
        const MASTER_PART: &str = "ppt/slideMasters/slideMaster1.xml";
        let master = br#"<p:sldMaster xmlns:p="p" xmlns:a="a"><p:txStyles><p:titleStyle>
          <a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:defRPr></a:lvl1pPr>
        </p:titleStyle></p:txStyles></p:sldMaster>"#;
        let bytes = stored_zip(&[(MASTER_PART, master)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut color_map = super::default_color_map();
        color_map.insert("tx2".to_owned(), "lt2".to_owned());
        let theme = PptxTheme::default().with_color_map(color_map);
        let styles = parse_master_text_styles(&package, MASTER_PART, &theme).unwrap();

        assert_eq!(styles.title[0].font_color, theme.color("lt2"));
    }

    #[test]
    fn master_title_shadow_is_inherited_by_text_runs() {
        const MASTER_PART: &str = "ppt/slideMasters/slideMaster1.xml";
        let master = br#"<p:sldMaster xmlns:p="p" xmlns:a="a"><p:txStyles><p:titleStyle>
          <a:lvl1pPr><a:defRPr><a:effectLst><a:outerShdw blurRad="38100" dist="25500" dir="5400000">
            <a:srgbClr val="000000"><a:alpha val="75000"/></a:srgbClr>
          </a:outerShdw></a:effectLst></a:defRPr></a:lvl1pPr>
        </p:titleStyle></p:txStyles></p:sldMaster>"#;
        let bytes = stored_zip(&[(MASTER_PART, master)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let styles =
            parse_master_text_styles(&package, MASTER_PART, &PptxTheme::default()).unwrap();
        let mut shape = ShapeState::new(0, &PptxTheme::default(), false);
        shape.placeholder_type = Some("title".to_owned());
        apply_master_text_styles(&mut shape, Some(&styles));
        shape.paragraph_pending = true;

        begin_shape_paragraph(&mut shape, false);
        let run = TextRunState::from_shape(&shape, 1);

        let shadow = run.shadow.expect("master title shadow is inherited");
        assert_eq!(shadow.color, 0x0000_00bf);
        assert!((shadow.blur - 4.0).abs() < 0.001);
        assert!((shadow.offset_y - 25_500.0 / 9_525.0).abs() < 0.001);
    }

    #[test]
    fn text_body_three_d_is_preserved_through_placeholder_inheritance() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp>
          <p:nvSpPr><p:cNvPr id="2"/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr><a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="soft" dir="t"/></a:scene3d>
            <a:sp3d><a:bevelT w="19050" h="12700"/><a:extrusionClr><a:schemeClr val="accent2"/></a:extrusionClr></a:sp3d></a:bodyPr><a:lstStyle/><a:p><a:r><a:t>Beveled</a:t></a:r></a:p></p:txBody>
        </p:sp></p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();
        let Visual::AdvancedEffect {
            three_d: Some(style),
            ..
        } = &objects[0].visual
        else {
            panic!("text body 3D is preserved");
        };
        assert_eq!(style.bevel_top.as_ref().map(|bevel| bevel.width), Some(2.0));
        assert!(style.applies_to_text);
        assert_eq!(style.extrusion_color, Some(0xed7d_31ff));

        let mut inherited = ShapeState::new(0, &PptxTheme::default(), false);
        super::inherit_visual_style(&mut inherited, &objects[0].visual);
        assert_eq!(
            inherited
                .three_d
                .and_then(|style| style.bevel_top)
                .map(|bevel| bevel.height),
            Some(4.0 / 3.0)
        );
    }

    #[test]
    fn subtitle_uses_master_body_text_style() {
        let mut styles = super::MasterTextStyles::default();
        styles.title[0].font_color = Some(0x0033_66ff);
        styles.body[0].font_color = Some(0x0000_00ff);
        let mut shape = ShapeState::new(0, &PptxTheme::default(), false);
        shape.placeholder_type = Some("subTitle".to_owned());

        apply_master_text_styles(&mut shape, Some(&styles));

        assert_eq!(shape.paragraph_styles[0].font_color, Some(0x0000_00ff));
    }

    #[test]
    fn shape_explicit_line_break_stays_within_its_paragraph_and_preserves_run_style() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="11"/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/></a:xfrm></p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p>
              <a:pPr algn="ctr" marL="95250"><a:lnSpc><a:spcPct val="125000"/></a:lnSpc></a:pPr>
              <a:r><a:rPr sz="1200" b="1"><a:latin typeface="Segoe UI"/></a:rPr><a:t>SWE.1</a:t></a:r>
              <a:br/>
              <a:r><a:rPr sz="1000" i="1"><a:latin typeface="Arial"/></a:rPr><a:t>Software Engineering</a:t></a:r>
            </a:p></p:txBody>
          </p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("a shape manual line break produces valid text");

        let shape = objects
            .iter()
            .find(|object| object.kind == ObjectKind::TextBox)
            .expect("shape text is present");
        assert_eq!(shape.text.as_deref(), Some("SWE.1\nSoftware Engineering"));
        let Visual::TextLayout { layout, visual } = &shape.visual else {
            panic!("shape text uses a text layout");
        };
        assert_eq!(layout.paragraphs.len(), 1, "a:br is not a paragraph break");
        assert_eq!(layout.paragraphs[0].align, crate::model::TextAlign::Center);
        assert!((layout.paragraphs[0].margin_left - 10.0).abs() < 0.001);
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("shape text uses rich text runs");
        };
        assert_eq!(
            runs.iter().map(|run| run.text.as_str()).collect::<Vec<_>>(),
            ["SWE.1", "\u{2028}", "Software Engineering"]
        );
        assert_eq!(runs[1].font_family, runs[0].font_family);
        assert_eq!(runs[1].font_size, runs[0].font_size);
        assert_eq!(runs[1].bold, runs[0].bold);
        assert_eq!(runs[1].italic, runs[0].italic);
        assert_ne!(runs[2].font_family, runs[1].font_family);
        assert_ne!(runs[2].font_size, runs[1].font_size);
        let height = super::estimated_drawingml_text_height(
            runs,
            &layout.paragraphs,
            16.0,
            400.0,
            true,
            1.0,
            0.0,
            &FontMetricTable::default(),
        );
        assert!((height - 2.0 * layout.paragraphs[0].line_height).abs() < 0.001);
    }

    #[test]
    fn supplied_font_scale_estimates_emergency_wrapped_text_height() {
        let bytes = include_bytes!("../../tests/fixtures/font-scale.pptx");
        let package = Package::open(bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        parse_test_slide(&package, &mut objects, &mut Vec::new()).unwrap();
        let object = objects
            .iter()
            .find(|object| object.kind == ObjectKind::TextBox)
            .unwrap();
        let Visual::TextLayout { layout, visual } = &object.visual else {
            panic!("text layout");
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("rich text");
        };
        let height = |wrap| {
            super::estimated_drawingml_text_height(
                runs,
                &layout.paragraphs,
                42.666668,
                object.bounds.width - layout.inset_left - layout.inset_right,
                wrap,
                layout.font_scale,
                layout.line_spacing_reduction,
                &FontMetricTable::default(),
            )
        };
        assert!((layout.font_scale - 0.85).abs() < 0.001);
        assert!(!layout.paragraphs[0].latin_line_break);
        assert!(
            height(true) > height(false) + 150.0,
            "oversized words must add body lines even when latinLnBrk is false"
        );
    }

    #[test]
    fn text_no_shape_keeps_text_runs_unwarped() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:sp><p:nvSpPr><p:cNvPr id="2"/></p:nvSpPr><p:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr><p:txBody><a:bodyPr><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp></a:bodyPr>
            <a:lstStyle/><a:p><a:r><a:rPr b="1"/><a:t>Black text </a:t></a:r>
            <a:r><a:rPr b="1" err="1"/><a:t>which was imported as white.</a:t></a:r>
            <a:endParaRPr b="1" baseline="-25000"/></a:p>
          </p:txBody></p:sp>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();

        parse_test_slide(&package, &mut objects, &mut diagnostics).expect("text box parses");

        let shape = objects
            .iter()
            .find(|object| object.kind == ObjectKind::TextBox)
            .expect("text box is present");
        let Visual::TextLayout { layout, visual } = &shape.visual else {
            panic!("shape text uses a text layout");
        };
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("shape text uses rich text runs");
        };
        assert_eq!(layout.warp, None);
        assert!(
            runs.iter()
                .all(|run| !run.italic && run.baseline_shift == 0.0)
        );
        assert!(diagnostics.iter().all(|diagnostic| {
            diagnostic
                .details
                .iter()
                .all(|(key, value)| key != "feature" || value != "wordart-warp")
        }));
    }

    #[test]
    fn table_explicit_line_break_is_emitted_once_inside_the_current_paragraph() {
        let slide = br#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="12"/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/></p:xfrm>
            <a:graphic><a:graphicData><a:tbl>
              <a:tblGrid><a:gridCol w="3810000"/></a:tblGrid>
              <a:tr h="1905000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>
                <a:pPr algn="ctr"/><a:r><a:rPr sz="1200" b="1"><a:latin typeface="Segoe UI"/></a:rPr><a:t>SWE.1</a:t></a:r>
                <a:br/>
                <a:r><a:rPr sz="1000" i="1"><a:latin typeface="Arial"/></a:rPr><a:t>Software Engineering</a:t></a:r>
              </a:p></a:txBody><a:tcPr/></a:tc></a:tr>
            </a:tbl></a:graphicData></a:graphic>
          </p:graphicFrame>
        </p:spTree></p:cSld></p:sld>"#;
        let bytes = stored_zip(&[
            (SLIDE_PART, slide),
            ("ppt/slides/_rels/slide1.xml.rels", b"<Relationships/>"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let mut objects = Vec::new();

        parse_test_slide(&package, &mut objects, &mut Vec::new())
            .expect("a table-cell manual line break produces valid text");

        let cell = objects
            .iter()
            .find(|object| object.kind == ObjectKind::Cell)
            .expect("table cell is present");
        assert_eq!(cell.text.as_deref(), Some("SWE.1\nSoftware Engineering"));
        let Visual::TextLayout { layout, visual } = &cell.visual else {
            panic!("table-cell text uses a text layout");
        };
        assert_eq!(layout.paragraphs.len(), 1, "a:br is not a paragraph break");
        let Visual::RichText { runs, .. } = visual.as_ref() else {
            panic!("table-cell text uses rich text runs");
        };
        assert_eq!(
            runs.iter().map(|run| run.text.as_str()).collect::<Vec<_>>(),
            ["SWE.1", "\u{2028}", "Software Engineering"]
        );
        assert_eq!(runs[1].font_family, runs[0].font_family);
        assert_eq!(runs[1].font_size, runs[0].font_size);
        assert_eq!(runs[1].bold, runs[0].bold);
    }

    #[test]
    fn pptx_table_tokenization_preserves_breaks_spaces_words_and_cjk_boundaries() {
        assert_eq!(
            drawingml_text_tokens("word  中A\r\nheavy-duty co‐pilot a/b tail-\tword"),
            [
                "word",
                "  ",
                "中",
                "A",
                "\r\n",
                "heavy-duty",
                " ",
                "co‐pilot",
                " ",
                "a/b",
                " ",
                "tail-",
                "\t",
                "word"
            ]
        );
        assert_eq!(
            drawingml_latin_line_break_tokens("heavy-duty"),
            ["heavy-", "duty"]
        );
        assert_eq!(drawingml_latin_line_break_tokens("tail-"), ["tail-"]);
    }

    #[test]
    fn curved_connector_uses_drawingml_segment_formulas() {
        let geometry = curved_connector_geometry(
            "curvedConnector3",
            Rect {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 100.0,
            },
            &HashMap::from([("adj1".to_owned(), 25_000.0)]),
            1.0,
            None,
            None,
        )
        .expect("curvedConnector3 has a DrawingML geometry");

        assert_eq!(
            geometry,
            Geometry::Path {
                fill_rule: crate::model::FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo { x: 0.0, y: 0.0 },
                    PathCommand::BezierCurveTo {
                        cp1x: 25.0,
                        cp1y: 0.0,
                        cp2x: 50.0,
                        cp2y: 25.0,
                        x: 50.0,
                        y: 50.0,
                    },
                    PathCommand::BezierCurveTo {
                        cp1x: 50.0,
                        cp1y: 75.0,
                        cp2x: 125.0,
                        cp2y: 100.0,
                        x: 200.0,
                        y: 100.0,
                    },
                ],
            }
        );
    }

    fn picture_package_bytes(target: &str, target_mode: Option<&str>, image: &[u8]) -> Vec<u8> {
        let slide = slide_xml();
        let rels = relationships_xml(target, target_mode);
        let mut entries: Vec<(&str, &[u8])> = vec![
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", rels.as_bytes()),
        ];
        if target_mode != Some("External") {
            let part = match target {
                "../media/image1.png" => "ppt/media/image1.png",
                "../media/image1.svg" => "ppt/media/image1.svg",
                "../media/image1.eps" => "ppt/media/image1.eps",
                _ => panic!("unexpected test target"),
            };
            entries.push((part, image));
        }
        stored_zip(&entries)
    }

    fn svg_preferred_picture_package_bytes(include_svg_relationship: bool, svg: &[u8]) -> Vec<u8> {
        let slide = format!(
            "<p:sld xmlns:p=\"p\" xmlns:a=\"a\" xmlns:r=\"r\" xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\"><p:cSld><p:spTree>{}</p:spTree></p:cSld></p:sld>",
            svg_preferred_picture_xml(42, "rIdFallback", "rIdSvg")
        );
        let svg_relationship = include_svg_relationship.then_some(
            "<Relationship Id=\"rIdSvg\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.svg\"/>",
        );
        let relationships = format!(
            "<Relationships><Relationship Id=\"rIdFallback\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.png\"/>{}</Relationships>",
            svg_relationship.unwrap_or_default()
        );
        stored_zip(&[
            (SLIDE_PART, slide.as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", relationships.as_bytes()),
            ("ppt/media/image1.png", PNG),
            ("ppt/media/image1.svg", svg),
        ])
    }

    #[test]
    fn explicit_empty_backgrounds_do_not_report_unsupported_fills() {
        for background in [
            "<p:bgPr><a:noFill/></p:bgPr>",
            "<p:bgRef idx=\"0\"><a:srgbClr val=\"00FF00\"/></p:bgRef>",
        ] {
            let slide = format!(
                "<p:sld xmlns:p=\"p\" xmlns:a=\"a\"><p:cSld><p:bg>{background}</p:bg><p:spTree/></p:cSld></p:sld>"
            );
            let bytes = stored_zip(&[(SLIDE_PART, slide.as_bytes())]);
            let package = Package::open(&bytes, Limits::default()).unwrap();
            let mut objects = Vec::new();
            let mut diagnostics = Vec::new();
            parse_test_slide(&package, &mut objects, &mut diagnostics).unwrap();
            assert!(objects.is_empty());
            assert!(diagnostics.is_empty(), "{diagnostics:?}");
        }
    }

    fn parse_test_slide(
        package: &Package<'_>,
        objects: &mut Vec<crate::model::Object>,
        diagnostics: &mut Vec<crate::diagnostic::Diagnostic>,
    ) -> Result<(), crate::diagnostic::Diagnostic> {
        parse_test_slide_with_metrics(package, objects, diagnostics, &FontMetricTable::default())
    }

    fn parse_test_slide_with_metrics(
        package: &Package<'_>,
        objects: &mut Vec<crate::model::Object>,
        diagnostics: &mut Vec<crate::diagnostic::Diagnostic>,
        font_metrics: &FontMetricTable,
    ) -> Result<(), crate::diagnostic::Diagnostic> {
        let next_z = objects
            .iter()
            .map(|object| object.z)
            .max()
            .unwrap_or(-1)
            .saturating_add(1);
        let mut state = super::PptxParseState {
            objects: std::mem::take(objects),
            next_z,
            diagnostics: std::mem::take(diagnostics),
            image_cache: std::collections::HashMap::new(),
            materialized_image_bytes: 0,
            reported_unsupported_backgrounds: std::collections::HashSet::new(),
            placeholders: std::collections::HashMap::new(),
            placeholder_text_styles: std::collections::HashMap::new(),
            placeholder_presets: std::collections::HashMap::new(),
            pending_placeholder_objects: std::collections::HashSet::new(),
            default_theme: super::PptxTheme::default(),
            default_text_styles: vec![super::ParagraphStyle::default(); super::TEXT_LEVEL_COUNT],
            theme_cache: std::collections::HashMap::new(),
        };
        let result = parse_slide(
            package,
            SlideParseContext {
                part: SLIDE_PART,
                unit_index: 0,
                properties: super::presentation_part_properties(package, SLIDE_PART)?,
                bounds: crate::model::Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 960.0,
                    height: 720.0,
                },
            },
            &mut state,
            &crate::format::ContentTypes::default(),
            font_metrics,
        );
        *objects = state.objects;
        *diagnostics = state.diagnostics;
        result
    }

    fn slide_xml() -> String {
        format!(
            "<p:sld xmlns:p=\"p\" xmlns:a=\"a\" xmlns:r=\"r\"><p:cSld><p:spTree>{}</p:spTree></p:cSld></p:sld>",
            picture_xml(42, "rIdImage")
        )
    }

    fn picture_xml(shape_id: u32, relationship_id: &str) -> String {
        format!(
            "<p:pic><p:nvPicPr><p:cNvPr id=\"{shape_id}\"/></p:nvPicPr><p:blipFill><a:blip r:embed=\"{relationship_id}\"/></p:blipFill><p:spPr><a:xfrm><a:off x=\"-9525\" y=\"19050\"/><a:ext cx=\"28575\" cy=\"38100\"/></a:xfrm><a:extLst><a:ext uri=\"extension\"/></a:extLst></p:spPr></p:pic>"
        )
    }

    fn svg_preferred_picture_xml(
        shape_id: u32,
        fallback_relationship_id: &str,
        svg_relationship_id: &str,
    ) -> String {
        format!(
            "<p:pic><p:nvPicPr><p:cNvPr id=\"{shape_id}\"/></p:nvPicPr><p:blipFill><a:blip r:embed=\"{fallback_relationship_id}\"><a:extLst><a:ext uri=\"{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}\"><asvg:svgBlip r:embed=\"{svg_relationship_id}\"/></a:ext></a:extLst></a:blip></p:blipFill><p:spPr><a:xfrm><a:off x=\"-9525\" y=\"19050\"/><a:ext cx=\"28575\" cy=\"38100\"/></a:xfrm></p:spPr></p:pic>"
        )
    }

    fn relationships_xml(target: &str, target_mode: Option<&str>) -> String {
        let mode = target_mode.map_or(String::new(), |mode| format!(" TargetMode=\"{mode}\""));
        format!(
            "<Relationships><Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"{target}\"{mode}/></Relationships>"
        )
    }
}
