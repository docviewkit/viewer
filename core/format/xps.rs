//! XPS/OpenXPS fixed-page parser.

use std::collections::{HashMap, HashSet};

use skrifa::{
    FontRef, GlyphId, MetadataProvider,
    instance::{LocationRef, Size},
    outline::{DrawSettings, OutlinePen},
    raw::{
        TableProvider,
        tables::cmap::{CmapSubtable, PlatformId},
    },
};

use crate::RetainedInput;
use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
use crate::limits::Limits;
use crate::model::{
    AffineTransform, BlendMode, Document, DocumentFormat, DocumentKind, EmbeddedFont, FillRule,
    FontStyle, Geometry, GradientSpread, MappingQuality, Object, ObjectKind, Paint, PathCommand,
    Rect, SheetAxis, SourceLocator, SourceRef, StretchMode, StrokeStyle, TileMode, Unit, UnitKind,
    Visual, VisualBrushChild, XpsColor, XpsGradientStop,
};
use crate::package::{Package, PackageCache, resolve_internal_target};
use crate::xml::{XmlEvent, decode_xml_text, parse_xml};

use super::local_name;

const FIXED_REPRESENTATION: &str = "/fixedrepresentation";
const MAX_PAGE_DIMENSION: f32 = 1_000_000.0;
const FONT_CONTENT_TYPE: &str = "application/vnd.ms-opentype";
const OBFUSCATED_FONT_CONTENT_TYPE: &str = "application/vnd.ms-package.obfuscated-opentype";
const FIXED_PAGE_HEADER_BYTES: usize = 16 * 1024;

#[derive(Clone, Debug)]
struct Node {
    source_part: String,
    name: String,
    attributes: Vec<(String, String)>,
    children: Vec<Node>,
}

impl Node {
    fn local_name(&self) -> &str {
        local_name(&self.name)
    }

    fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|(candidate, _)| local_name(candidate) == name)
            .map(|(_, value)| value.as_str())
    }

    fn property(&self, name: &str) -> Option<&Node> {
        self.children
            .iter()
            .find(|child| child.local_name() == name)
    }
}

pub fn detect_and_parse(bytes: &[u8], limits: Limits) -> Result<Option<Document>, Diagnostic> {
    let Some(prepared) = prepare(bytes, limits)? else {
        return Ok(None);
    };
    prepared
        .materialize(None, &HashSet::new())
        .map(|(document, _)| Some(document))
}

pub(crate) struct PreparedXps {
    bytes: PreparedBytes,
    cache: PackageCache,
    diagnostics: Vec<Diagnostic>,
    sequence_part: String,
    page_parts: Vec<String>,
    page_dimensions: Vec<std::sync::OnceLock<(f32, f32)>>,
    limits: Limits,
}

enum PreparedBytes {
    Owned(Vec<u8>),
    Retained(std::sync::Arc<RetainedInput>),
}

impl std::ops::Deref for PreparedBytes {
    type Target = [u8];

    fn deref(&self) -> &Self::Target {
        match self {
            Self::Owned(bytes) => bytes,
            Self::Retained(bytes) => bytes,
        }
    }
}

pub(crate) fn prepare(bytes: &[u8], limits: Limits) -> Result<Option<PreparedXps>, Diagnostic> {
    prepare_impl(bytes, limits, None)
}

pub(crate) fn prepare_retained(
    bytes: std::sync::Arc<RetainedInput>,
    limits: Limits,
) -> Result<Option<PreparedXps>, Diagnostic> {
    let slice: &[u8] = &bytes;
    prepare_impl(slice, limits, Some(std::sync::Arc::clone(&bytes)))
}

fn prepare_impl(
    bytes: &[u8],
    limits: Limits,
    retained: Option<std::sync::Arc<RetainedInput>>,
) -> Result<Option<PreparedXps>, Diagnostic> {
    if !bytes.starts_with(b"PK") {
        return Ok(None);
    }
    let package = Package::open(bytes, limits)?;
    let mut diagnostics = Vec::new();
    let relationships = match package.relationships(None) {
        Ok(relationships) => relationships,
        Err(mut error) if error.code == DiagnosticCode::XmlInvalid => {
            let relationships = package.relationships_filtered(None, Some(FIXED_REPRESENTATION))?;
            error.severity = crate::diagnostic::Severity::Warning;
            error.fidelity = Fidelity::Omitted;
            diagnostics.push(error);
            relationships
        }
        Err(error) => return Err(error),
    };
    let sequence_part = relationships
        .into_iter()
        .find(|relationship| {
            !relationship.external
                && relationship
                    .type_uri
                    .to_ascii_lowercase()
                    .ends_with(FIXED_REPRESENTATION)
        })
        .map(|relationship| relationship.target);
    let mut sequence_part = match sequence_part {
        Some(part) => part,
        None => {
            let Some(part) = super::find_xps_sequence(&package, Some(&mut diagnostics))? else {
                return Ok(None);
            };
            diagnostics.push(Diagnostic::warning(DiagnosticCode::FormatInvalid, Phase::Parse,
                Fidelity::Approximate, "recovered missing fixedrepresentation relationship from the unique declared XPS sequence").in_part("_rels/.rels"));
            part
        }
    };
    if !package.has_part(&sequence_part)
        && super::find_xps_sequence(&package, Some(&mut diagnostics))?.is_none()
    {
        if let Some(part) = unique_fixed_document(&package, &sequence_part, &mut diagnostics)? {
            diagnostics.push(Diagnostic::warning(DiagnosticCode::FormatInvalid, Phase::Parse,
                Fidelity::Approximate, format!("recovered missing XPS sequence {sequence_part} through the unique FixedDocument {part}")).in_part(&sequence_part));
            sequence_part = part;
        }
    }
    let page_parts = collect_page_parts(&package, &sequence_part, &mut diagnostics)?;
    Ok(Some(PreparedXps {
        cache: package.cache(),
        diagnostics,
        bytes: retained.map_or_else(
            || PreparedBytes::Owned(bytes.to_vec()),
            PreparedBytes::Retained,
        ),
        sequence_part,
        page_dimensions: (0..page_parts.len())
            .map(|_| std::sync::OnceLock::new())
            .collect(),
        page_parts,
        limits,
    }))
}

fn unique_fixed_document(
    package: &Package<'_>,
    source: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<Option<String>, Diagnostic> {
    let types = super::parse_content_types_with_diagnostics(package, Some(diagnostics))?;
    let mut parts = package.entry_names().filter(|part| {
        matches!(
            types.for_part(part),
            Some(
                "application/vnd.ms-package.xps-fixeddocument+xml"
                    | "application/oxps-fixeddocument+xml"
            )
        )
    });
    let part = parts.next().map(str::to_owned);
    if parts.next().is_some() {
        return Err(format_error(
            source,
            "incomplete XPS sequence has multiple candidate FixedDocuments",
        ));
    }
    Ok(part)
}

fn collect_page_parts(
    package: &Package<'_>,
    sequence_part: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<Vec<String>, Diagnostic> {
    let sequence = parse_part(package, sequence_part)?;
    if !matches!(
        sequence.local_name(),
        "FixedDocumentSequence" | "FixedDocument" | "FixedPage"
    ) {
        return Err(format_error(
            sequence_part,
            "XPS fixed representation is not a FixedDocumentSequence",
        ));
    }
    let mut document_parts = Vec::new();
    if sequence.local_name() != "FixedDocumentSequence" {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::FormatInvalid,
                Phase::Parse,
                Fidelity::Approximate,
                format!(
                    "recovered direct {} as XPS fixed representation",
                    sequence.local_name()
                ),
            )
            .in_part(sequence_part),
        );
        document_parts.push(sequence_part.to_owned());
    }
    for reference in &sequence.children {
        if reference.local_name() != "DocumentReference" {
            continue;
        }
        let source = required_attribute(reference, "Source", sequence_part)?;
        let mut target = match resolve_part(sequence_part, source) {
            Ok(target) => target,
            Err(error) => {
                // Some producers put literal URI delimiters in ZIP member names.
                // Only an exact, archive-validated absolute member is recoverable.
                let Some(literal) = source
                    .strip_prefix('/')
                    .filter(|part| package.entry_names().any(|name| name == *part))
                else {
                    return Err(error);
                };
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::FormatInvalid,
                        Phase::Parse,
                        Fidelity::Approximate,
                        format!("recovered literal XPS part name {source}"),
                    )
                    .in_part(sequence_part),
                );
                literal.to_owned()
            }
        };
        if source.ends_with('/') && package.part(&target)?.is_some() {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::FormatInvalid,
                    Phase::Parse,
                    Fidelity::Approximate,
                    format!("ignored trailing slash on XPS part reference {source}"),
                )
                .in_part(sequence_part),
            );
        }
        if (source.ends_with('/') || source.ends_with("/.")) && package.part(&target)?.is_none() {
            let prefix = format!("{target}/");
            let content_types =
                super::parse_content_types_with_diagnostics(package, Some(diagnostics))?;
            let mut candidates = package.entry_names().filter(|part| {
                part.strip_prefix(&prefix)
                    .is_some_and(|name| !name.is_empty() && !name.contains('/'))
                    && content_types.for_part(part).is_some_and(|kind| {
                        kind == "application/vnd.ms-package.xps-fixeddocument+xml"
                            || kind == "application/oxps-fixeddocument+xml"
                    })
            });
            let candidate = candidates.next().ok_or_else(|| {
                format_error(
                    sequence_part,
                    format!("XPS directory reference {source} contains no FixedDocument"),
                )
            })?;
            if candidates.next().is_some() {
                return Err(format_error(
                    sequence_part,
                    format!("XPS directory reference {source} contains multiple FixedDocuments"),
                ));
            }
            target = candidate.to_owned();
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::FormatInvalid,
                    Phase::Parse,
                    Fidelity::Approximate,
                    format!("recovered XPS directory reference {source} as {target}"),
                )
                .in_part(sequence_part),
            );
        }
        document_parts.push(target);
    }
    if document_parts.is_empty()
        && let Some(part) = unique_fixed_document(package, sequence_part, diagnostics)?
    {
        diagnostics.push(Diagnostic::warning(DiagnosticCode::FormatInvalid, Phase::Parse,
            Fidelity::Approximate, format!("recovered missing XPS DocumentReference through the unique FixedDocument {part}")).in_part(sequence_part));
        document_parts.push(part);
    }
    if document_parts.is_empty() {
        return Err(format_error(
            sequence_part,
            "XPS document sequence contains no documents",
        ));
    }

    let mut page_parts = Vec::new();
    for document_part in document_parts {
        if !package.has_part(&document_part) {
            // Without the document index, only a unique page in this document's
            // directory has an unambiguous ownership and reading order.
            let prefix = document_part
                .rsplit_once('/')
                .map(|(directory, _)| format!("{directory}/"))
                .unwrap_or_default();
            let types = super::parse_content_types_with_diagnostics(package, Some(diagnostics))?;
            let mut candidates = package.entry_names().filter(|part| {
                part.starts_with(&prefix)
                    && matches!(
                        types.for_part(part),
                        Some(
                            "application/vnd.ms-package.xps-fixedpage+xml"
                                | "application/oxps-fixedpage+xml"
                        )
                    )
            });
            if let Some(page) = candidates.next() {
                if candidates.next().is_some() {
                    return Err(format_error(
                        &document_part,
                        "missing XPS document has multiple candidate pages with unknown reading order",
                    ));
                }
                if parse_part(package, page)?.local_name() != "FixedPage" {
                    return Err(format_error(
                        page,
                        "recovered XPS page part is not a FixedPage",
                    ));
                }
                diagnostics.push(Diagnostic::warning(DiagnosticCode::FormatInvalid, Phase::Parse,
                    Fidelity::Approximate, format!("recovered missing XPS document {document_part} through its unique declared page {page}")).in_part(&document_part));
                page_parts.push(page.to_owned());
                continue;
            }
        }
        let fixed_document = parse_part(package, &document_part)?;
        if fixed_document.local_name() == "FixedPage" {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::FormatInvalid,
                    Phase::Parse,
                    Fidelity::Approximate,
                    "recovered direct FixedPage document reference",
                )
                .in_part(&document_part),
            );
            page_parts.push(document_part);
            continue;
        }
        if fixed_document.local_name() != "FixedDocument" {
            return Err(format_error(
                &document_part,
                "XPS document reference is not a FixedDocument",
            ));
        }
        for page in &fixed_document.children {
            if page.local_name() != "PageContent" {
                continue;
            }
            let source = required_attribute(page, "Source", &document_part)?;
            page_parts.push(resolve_part(&document_part, source)?);
        }
    }
    if page_parts.is_empty() {
        return Err(format_error(
            sequence_part,
            "XPS document contains no fixed pages",
        ));
    }
    if page_parts.len() > package.limits().max_document_objects {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "XPS page count exceeds the configured object limit",
        ));
    }
    Ok(page_parts)
}

impl PreparedXps {
    pub(crate) fn materialize(
        &self,
        selected_unit: Option<usize>,
        emitted_fonts: &HashSet<String>,
    ) -> Result<(Document, HashSet<String>), Diagnostic> {
        let package = Package::open_with_cache(&self.bytes, self.limits, self.cache.clone())?;
        let (mut document, fonts) = parse(
            &package,
            &self.sequence_part,
            &self.page_parts,
            &self.page_dimensions,
            selected_unit,
            emitted_fonts,
        )?;
        if selected_unit.is_none_or(|index| index == 0) {
            document
                .diagnostics
                .extend(self.diagnostics.iter().cloned());
        }
        Ok((document, fonts))
    }
}

fn parse(
    package: &Package<'_>,
    sequence_part: &str,
    page_parts: &[String],
    page_dimensions: &[std::sync::OnceLock<(f32, f32)>],
    selected_unit: Option<usize>,
    emitted_fonts: &HashSet<String>,
) -> Result<(Document, HashSet<String>), Diagnostic> {
    let mut diagnostics = Vec::new();
    let content_types =
        super::parse_content_types_with_diagnostics(package, Some(&mut diagnostics))?;

    let mut document = Document {
        fatal: false,
        format: Some(DocumentFormat::Xps),
        kind: Some(DocumentKind::Text),
        units: Vec::with_capacity(page_parts.len()),
        outline: Vec::new(),
        objects: Vec::new(),
        embedded_fonts: Vec::new(),
        font_alternate_names: Vec::new(),
        diagnostics,
    };
    let mut loaded_fonts = HashSet::new();
    let mut font_bytes = 0;
    for (page_index, page_part) in page_parts.iter().enumerate() {
        let materialize_page = selected_unit.is_none_or(|selected| selected == page_index);
        let page = if materialize_page {
            Some(parse_part(package, page_part)?)
        } else if page_dimensions[page_index].get().is_none() {
            Some(parse_fixed_page_header(package, page_part)?)
        } else {
            None
        };
        let (width, height) = if let Some(page) = &page {
            if page.local_name() != "FixedPage" {
                return Err(format_error(page_part, "XPS page part is not a FixedPage"));
            }
            let width = finite_attribute(page, "Width", page_part)?;
            let height = finite_attribute(page, "Height", page_part)?;
            if width <= 0.0
                || height <= 0.0
                || width > MAX_PAGE_DIMENSION
                || height > MAX_PAGE_DIMENSION
            {
                return Err(format_error(page_part, "invalid XPS fixed-page dimensions"));
            }
            let dimensions = (width, height);
            let _ = page_dimensions[page_index].set(dimensions);
            dimensions
        } else {
            *page_dimensions[page_index].get().unwrap()
        };
        let unit_index = document.units.len() as u32;
        document.units.push(Unit {
            kind: UnitKind::Page,
            index: unit_index,
            id: format!("xps-page-{}", page_index + 1),
            name: format!("Page {}", page_index + 1),
            width,
            height,
            rows: 1,
            columns: 1,
            frozen_rows: 0,
            frozen_columns: 0,
            frozen_width: 0.0,
            frozen_height: 0.0,
            row_axis: SheetAxis::uniform(1, height),
            column_axis: SheetAxis::uniform(1, width),
            show_grid_lines: false,
            tab_color: None,
            sheet: None,
            slide: None,
        });
        if !materialize_page {
            continue;
        }
        let next_id = u32::try_from(document.objects.len())
            .map_err(|_| format_error(page_part, "XPS object count exceeds supported range"))?;
        let mut context = PageContext {
            package,
            content_types: &content_types,
            page_part,
            unit_index,
            page_bounds: Rect {
                x: 0.0,
                y: 0.0,
                width,
                height,
            },
            resources: HashMap::new(),
            resource_parts: HashSet::new(),
            objects: &mut document.objects,
            diagnostics: &mut document.diagnostics,
            fonts: &mut document.embedded_fonts,
            loaded_fonts: &mut loaded_fonts,
            font_bytes: &mut font_bytes,
            next_id,
            z: 0,
        };
        let page = page.as_ref().unwrap();
        context.collect_resources(page, page_part)?;
        context.render_children(page, None, "FixedPage")?;
    }
    if document.units.is_empty() {
        return Err(format_error(
            sequence_part,
            "XPS document contains no renderable fixed pages",
        ));
    }
    let mut next_emitted_fonts = emitted_fonts.clone();
    document.embedded_fonts.retain(|font| {
        if emitted_fonts.contains(&font.family) {
            false
        } else {
            next_emitted_fonts.insert(font.family.clone());
            true
        }
    });
    Ok((document, next_emitted_fonts))
}

struct PageContext<'a, 'b> {
    package: &'a Package<'a>,
    content_types: &'b super::ContentTypes,
    page_part: &'b str,
    unit_index: u32,
    page_bounds: Rect,
    resources: HashMap<String, Node>,
    resource_parts: HashSet<String>,
    objects: &'b mut Vec<Object>,
    diagnostics: &'b mut Vec<Diagnostic>,
    fonts: &'b mut Vec<EmbeddedFont>,
    loaded_fonts: &'b mut HashSet<String>,
    font_bytes: &'b mut usize,
    next_id: u32,
    z: i32,
}

impl PageContext<'_, '_> {
    fn collect_resources(&mut self, node: &Node, source_part: &str) -> Result<(), Diagnostic> {
        for child in &node.children {
            if child.local_name().ends_with(".Resources")
                || child.local_name() == "ResourceDictionary"
            {
                self.collect_resource_dictionary(child, source_part)?;
            }
        }
        Ok(())
    }

    fn collect_resource_dictionary(
        &mut self,
        node: &Node,
        source_part: &str,
    ) -> Result<(), Diagnostic> {
        if let Some(source) = node.attribute("Source") {
            let target = resolve_part(source_part, source)?;
            if self.resource_parts.insert(target.clone()) {
                let result = parse_part(self.package, &target)
                    .and_then(|dictionary| self.collect_resource_dictionary(&dictionary, &target));
                self.resource_parts.remove(&target);
                result?;
            }
        }
        for child in &node.children {
            if child.local_name() == "ResourceDictionary" {
                self.collect_resource_dictionary(child, source_part)?;
            }
            if let Some(key) = child.attribute("Key") {
                self.resources.insert(key.to_owned(), child.clone());
            }
        }
        Ok(())
    }

    fn render_children(
        &mut self,
        node: &Node,
        parent: Option<(u32, String)>,
        path: &str,
    ) -> Result<(), Diagnostic> {
        let mut visual_index = 0_usize;
        for child in &node.children {
            if child.local_name().contains('.') || child.local_name() == "ResourceDictionary" {
                continue;
            }
            visual_index += 1;
            let child_path = format!("{path}/{}[{visual_index}]", child.local_name());
            self.render_visual(child, parent.clone(), &child_path)?;
            if self.objects.len() > self.package.limits().max_document_objects {
                return Err(Diagnostic::fatal(
                    DiagnosticCode::ObjectLimit,
                    Phase::Parse,
                    None,
                    "XPS visual count exceeds the configured object limit",
                )
                .in_part(self.page_part));
            }
        }
        Ok(())
    }

    fn render_visual(
        &mut self,
        node: &Node,
        parent: Option<(u32, String)>,
        path: &str,
    ) -> Result<(), Diagnostic> {
        let start = self.objects.len();
        let result = match node.local_name() {
            "Canvas" => self.render_canvas(node, parent, path),
            "Path" => self.render_path(node, parent, path),
            "Glyphs" => self.render_glyphs(node, parent, path),
            _ => {
                self.omit_unknown_visual(node);
                Ok(())
            }
        };
        if let Err(mut error) = result {
            if error.code != DiagnosticCode::FormatInvalid {
                return Err(error);
            }
            self.objects.truncate(start);
            error.severity = crate::diagnostic::Severity::Warning;
            error.fidelity = Fidelity::Omitted;
            error.location.part = Some(node.source_part.clone());
            error.message = format!("omitted XPS {path}: {}", error.message);
            self.diagnostics.push(error);
        }
        Ok(())
    }

    fn render_canvas(
        &mut self,
        node: &Node,
        parent: Option<(u32, String)>,
        path: &str,
    ) -> Result<(), Diagnostic> {
        let inherited_resources = self.resources.clone();
        let result = (|| {
            self.collect_resources(node, self.page_part)?;
            let (numeric_id, stable_id) = self.ids(node, "canvas");
            let canvas_index = self.objects.len();
            self.push_object(
                numeric_id,
                stable_id.clone(),
                parent.clone(),
                ObjectKind::Group,
                self.page_bounds,
                None,
                "canvas",
                path,
                Visual::None,
            );
            self.render_children(node, Some((numeric_id, stable_id)), path)?;
            // Keep Canvas transforms, clips and opacity around the child visuals in
            // the render tree, not only in the source-object ancestry.
            let children = captured_visual_children(
                &self.objects[canvas_index + 1..],
                Some(numeric_id),
                self.page_part,
            )?;
            for object in &mut self.objects[canvas_index + 1..] {
                object.visual = Visual::None;
            }
            self.objects[canvas_index].visual =
                self.common_visual(node, Visual::Group { children }, self.page_bounds)?;
            Ok(())
        })();
        self.resources = inherited_resources;
        result
    }

    fn render_path(
        &mut self,
        node: &Node,
        parent: Option<(u32, String)>,
        path: &str,
    ) -> Result<(), Diagnostic> {
        let mut data = self
            .geometry(node, true)?
            .ok_or_else(|| format_error(self.page_part, "XPS Path has no geometry"))?;
        let mut fill_data = self.geometry(node, false)?.unwrap_or_else(|| data.clone());
        let fill = self.paint(node, "Fill")?;
        let stroke = self.paint(node, "Stroke")?;
        let stroke_width = optional_finite_attribute(node, "StrokeThickness")?
            .unwrap_or(1.0)
            .max(0.0);
        let mut bounds = geometry_bounds(&data)
            .ok_or_else(|| format_error(self.page_part, "XPS Path has empty geometry"))?;
        if !matches!(stroke, Paint::None) && stroke_width > 0.0 {
            let padding = stroke_width / 2.0;
            bounds.x -= padding;
            bounds.y -= padding;
            bounds.width += stroke_width;
            bounds.height += stroke_width;
        }
        rebase_geometry(&mut data, bounds.x, bounds.y);
        rebase_geometry(&mut fill_data, bounds.x, bounds.y);
        let cap_geometry = stroke_extensions(
            node,
            &data,
            stroke_width,
            self.package.limits().max_document_objects,
        )?;
        let fill = fill.localize_gradient(bounds.x, bounds.y);
        let stroke = stroke.localize_gradient(bounds.x, bounds.y);
        let cap_paint = stroke.clone();
        let mut visual = if fill_data == data {
            Visual::PaintedShape {
                geometry: data,
                fill,
                stroke,
                stroke_width,
            }
        } else {
            let local_bounds = bounds;
            Visual::Group {
                children: vec![
                    VisualBrushChild {
                        bounds: local_bounds,
                        visual: Visual::PaintedShape {
                            geometry: fill_data,
                            fill,
                            stroke: Paint::None,
                            stroke_width: 0.0,
                        },
                    },
                    VisualBrushChild {
                        bounds: local_bounds,
                        visual: Visual::PaintedShape {
                            geometry: data,
                            fill: Paint::None,
                            stroke,
                            stroke_width,
                        },
                    },
                ],
            }
        };
        if let Some(style) = stroke_style(node)? {
            visual = Visual::StrokeStyle {
                style,
                visual: Box::new(visual),
            };
        }
        if let Some(geometry) = cap_geometry {
            visual = Visual::Group {
                children: vec![
                    VisualBrushChild { bounds, visual },
                    VisualBrushChild {
                        bounds,
                        visual: Visual::PaintedShape {
                            geometry,
                            fill: cap_paint,
                            stroke: Paint::None,
                            stroke_width: 0.0,
                        },
                    },
                ],
            };
        }
        visual = self.common_visual(node, visual, bounds)?;
        let (numeric_id, stable_id) = self.ids(node, "path");
        self.push_object(
            numeric_id,
            stable_id,
            parent,
            ObjectKind::Shape,
            bounds,
            None,
            "path",
            path,
            visual,
        );
        Ok(())
    }

    fn render_glyphs(
        &mut self,
        node: &Node,
        parent: Option<(u32, String)>,
        path: &str,
    ) -> Result<(), Diagnostic> {
        let raw_text = node.attribute("UnicodeString").unwrap_or("");
        let text = raw_text.strip_prefix("{}").unwrap_or(raw_text).to_owned();
        let font_size = finite_attribute(node, "FontRenderingEmSize", self.page_part)?;
        let origin_x = finite_attribute(node, "OriginX", self.page_part)?;
        let origin_y = finite_attribute(node, "OriginY", self.page_part)?;
        if font_size < 0.0 {
            return Err(format_error(
                self.page_part,
                "XPS FontRenderingEmSize must be non-negative",
            ));
        }
        let bidi_level = node
            .attribute("BidiLevel")
            .unwrap_or("0")
            .parse::<u8>()
            .ok()
            .filter(|value| *value <= 61)
            .ok_or_else(|| format_error(self.page_part, "invalid XPS BidiLevel"))?;
        let sideways = match node.attribute("IsSideways") {
            None | Some("false") => false,
            Some("true") => true,
            Some(_) => return Err(format_error(self.page_part, "invalid XPS IsSideways value")),
        };
        if sideways && bidi_level % 2 == 1 {
            return Err(format_error(
                self.page_part,
                "XPS sideways glyphs cannot use an odd BidiLevel",
            ));
        }
        let style = GlyphStyle::parse(node.attribute("StyleSimulations"), self.page_part)?;
        let font_uri = required_attribute(node, "FontUri", self.page_part)?;
        let (font_family, font_face) = match self.load_font(font_uri, &node.source_part) {
            Ok(font) => font,
            Err(mut error)
                if error.code == DiagnosticCode::FormatInvalid
                    && !text.is_empty()
                    && !sideways
                    && bidi_level % 2 == 0 =>
            {
                let paint = self.paint(node, "Fill")?;
                let bounds = Rect {
                    x: origin_x,
                    y: origin_y - font_size,
                    width: font_size * (text.chars().count() as f32 + 1.0),
                    height: font_size * 1.2,
                };
                // Canvas/VisualBrush children retain visuals, not object text.
                let visual = self.common_visual(
                    node,
                    Visual::RichText {
                        geometry: Geometry::Rectangle,
                        fill: Paint::None,
                        stroke: Paint::None,
                        stroke_width: 0.0,
                        align: crate::model::TextAlign::Start,
                        line_height: font_size * 1.2,
                        runs: vec![crate::model::TextRun {
                            text: text.clone(),
                            font_family: "sans-serif".to_owned(),
                            font_size,
                            color: 0x000000ff,
                            paint: Some(Box::new(paint)),
                            east_asian_line_breaks: true,
                            bold: style.bold,
                            italic: style.italic,
                            underline: false,
                            strikethrough: false,
                            highlight: 0,
                            baseline_shift: 0.0,
                            letter_spacing: 0.0,
                            horizontal_scale: 1.0,
                        }],
                    },
                    bounds,
                )?;
                let (numeric_id, stable_id) = self.ids(node, "glyphs");
                self.push_object(
                    numeric_id,
                    stable_id,
                    parent,
                    ObjectKind::TextBox,
                    bounds,
                    Some(text),
                    "glyphs",
                    path,
                    visual,
                );
                error.severity = crate::diagnostic::Severity::Warning;
                error.fidelity = Fidelity::Approximate;
                error.message = format!(
                    "used fallback font for XPS UnicodeString: {}",
                    error.message
                );
                self.diagnostics.push(error);
                return Ok(());
            }
            Err(error) => return Err(error),
        };
        let font_bytes = self
            .fonts
            .iter()
            .find(|font| font.family == font_family)
            .map(|font| font.bytes.as_slice())
            .ok_or_else(|| format_error(self.page_part, "XPS embedded font was not retained"))?;
        let shape = |indices| {
            glyph_geometry(
                font_bytes,
                font_face,
                &text,
                indices,
                origin_x,
                origin_y,
                font_size,
                bidi_level % 2 == 1,
                sideways,
                style,
                self.package.limits().max_document_objects,
            )
        };
        let (geometry, bounds) = match shape(node.attribute("Indices")) {
            Err(mut error)
                if error.code == DiagnosticCode::FormatInvalid
                    && node.attribute("Indices").is_some()
                    && !text.is_empty() =>
            {
                let geometry = shape(None)?;
                error.severity = crate::diagnostic::Severity::Warning;
                error.fidelity = Fidelity::Approximate;
                error.location.part = Some(node.source_part.clone());
                error.message =
                    format!("recovered XPS glyphs from UnicodeString: {}", error.message);
                self.diagnostics.push(error);
                geometry
            }
            result => result?,
        };
        let fill = self
            .paint(node, "Fill")?
            .localize_gradient(bounds.x, bounds.y);
        let visual = self.common_visual(
            node,
            Visual::PaintedShape {
                geometry,
                fill: fill.clone(),
                stroke: if style.bold { fill } else { Paint::None },
                stroke_width: if style.bold { font_size * 0.02 } else { 0.0 },
            },
            bounds,
        )?;
        let (numeric_id, stable_id) = self.ids(node, "glyphs");
        self.push_object(
            numeric_id,
            stable_id,
            parent,
            ObjectKind::TextBox,
            bounds,
            (!text.is_empty()).then_some(text),
            "glyphs",
            path,
            visual,
        );
        Ok(())
    }

    fn geometry(&mut self, node: &Node, for_stroke: bool) -> Result<Option<Geometry>, Diagnostic> {
        let resource;
        let geometry = if let Some(value) = node.attribute("Data") {
            if let Some(key) = static_resource_key(value) {
                resource = self.resources.get(key).cloned().ok_or_else(|| {
                    format_error(self.page_part, format!("missing XPS resource {key}"))
                })?;
                Some(&resource)
            } else {
                return parse_path_geometry(value)
                    .map(Some)
                    .map_err(|message| format_error(self.page_part, message));
            }
        } else {
            node.property("Path.Data")
                .and_then(|property| property.children.first())
        };
        let Some(geometry) = geometry else {
            return Ok(None);
        };
        let mut result = if let Some(figures) = geometry.attribute("Figures") {
            Some(
                parse_path_geometry(figures)
                    .map_err(|message| format_error(self.page_part, message))?,
            )
        } else {
            parse_path_geometry_node(geometry, for_stroke)
        };
        if let Some(Geometry::Path {
            fill_rule,
            commands,
        }) = &mut result
        {
            if let Some(rule) = geometry.attribute("FillRule") {
                *fill_rule = if rule == "NonZero" {
                    FillRule::NonZero
                } else {
                    FillRule::EvenOdd
                };
            }
            let transform = self.brush_transform(geometry, "Transform")?;
            for command in commands {
                command.transform(transform);
            }
        }
        Ok(result)
    }

    fn paint(&mut self, node: &Node, property: &str) -> Result<Paint, Diagnostic> {
        if let Some(value) = node.attribute(property) {
            if let Some(key) = static_resource_key(value) {
                let resource = self.resources.get(key).cloned().ok_or_else(|| {
                    format_error(self.page_part, format!("missing XPS resource {key}"))
                })?;
                return self.brush(&resource);
            }
            return self
                .parse_color_value(value, &node.source_part)
                .map(xps_color_paint);
        }
        let property_name = format!("{}.{}", node.local_name(), property);
        let brush = node
            .property(&property_name)
            .and_then(|property_node| property_node.children.first());
        brush.map_or(Ok(Paint::None), |brush| self.brush(brush))
    }

    fn brush(&mut self, node: &Node) -> Result<Paint, Diagnostic> {
        match node.local_name() {
            "SolidColorBrush" => {
                let color = self.parse_color_value(
                    node.attribute("Color").ok_or_else(|| {
                        format_error(self.page_part, "XPS solid color is missing")
                    })?,
                    &node.source_part,
                )?;
                Ok(xps_color_paint(brush_color_opacity(node, color)?))
            }
            "LinearGradientBrush" => {
                let (x0, y0) = point(node.attribute("StartPoint").unwrap_or("0,0"))?;
                let (x1, y1) = point(node.attribute("EndPoint").unwrap_or("1,1"))?;
                self.gradient_brush(node, false, x0, y0, x1, y1, 0.0, 0.0)
            }
            "RadialGradientBrush" => {
                let (x1, y1) = point(node.attribute("Center").unwrap_or("0.5,0.5"))?;
                let (x0, y0) = point(node.attribute("GradientOrigin").unwrap_or("0.5,0.5"))?;
                let radius_x = optional_finite_attribute(node, "RadiusX")?
                    .unwrap_or(0.5)
                    .abs();
                self.gradient_brush(
                    node,
                    true,
                    x0,
                    y0,
                    x1,
                    y1,
                    radius_x,
                    optional_finite_attribute(node, "RadiusY")?
                        .unwrap_or(0.5)
                        .abs(),
                )
            }
            "ImageBrush" => self.image_brush(node),
            "VisualBrush" => self.visual_brush(node),
            _ => Err(format_error(
                self.page_part,
                format!("unsupported XPS brush {}", node.local_name()),
            )),
        }
    }

    fn image_brush(&mut self, node: &Node) -> Result<Paint, Diagnostic> {
        let Some(source) = node.attribute("ImageSource") else {
            return Ok(Paint::None);
        };
        let (source, source_profile, destination_profile) = color_converted_bitmap_source(source)?;
        let target = resolve_part(&node.source_part, source)?;
        let bytes = self
            .package
            .part(&target)?
            .ok_or_else(|| format_error(&target, "XPS image resource is missing"))?;
        let media_type = image_media_type(&target, &bytes)
            .ok_or_else(|| format_error(&target, "unsupported XPS image encoding"))?;
        let (width, height) = xps_image_dimensions(media_type, &bytes)
            .ok_or_else(|| format_error(&target, "XPS image dimensions are invalid"))?;
        if u64::from(width) * u64::from(height) > Limits::HARD_MAX.max_render_pixels as u64 {
            self.diagnostics.push(Diagnostic::warning(DiagnosticCode::ImageDimensionLimit,
                Phase::Security, Fidelity::Blocked,
                format!("blocked XPS image {width}x{height} before decoding: dimensions exceed the hard pixel limit")).in_part(&target));
            return Ok(Paint::None);
        }
        let (dpi_x, dpi_y) = xps_image_resolution(media_type, &bytes);
        let mut image = Visual::Image {
            media_type: media_type.to_owned(),
            bytes: bytes.into_vec(),
            crop: Default::default(),
        };
        if let Some(source_profile) = source_profile {
            let load_profile = |uri: &str| -> Result<Vec<u8>, Diagnostic> {
                let target = resolve_part(&node.source_part, uri)?;
                let bytes = self
                    .package
                    .part(&target)?
                    .ok_or_else(|| format_error(&target, "XPS bitmap color profile is missing"))?
                    .into_vec();
                icc_channel_count(&bytes)
                    .ok_or_else(|| format_error(&target, "XPS bitmap color profile is invalid"))?;
                Ok(bytes)
            };
            image = Visual::ColorManagedImage {
                source_profile: load_profile(source_profile)?,
                destination_profile: destination_profile.map(load_profile).transpose()?,
                visual: Box::new(image),
            };
        }
        Ok(Paint::Visual {
            viewbox: brush_rect(node.attribute("Viewbox").unwrap_or("0,0,1,1"), "Viewbox")?,
            viewport: brush_rect(node.attribute("Viewport").unwrap_or("0,0,1,1"), "Viewport")?,
            viewbox_relative: brush_units(node.attribute("ViewboxUnits"), "ViewboxUnits")?,
            viewport_relative: brush_units(node.attribute("ViewportUnits"), "ViewportUnits")?,
            tile_mode: tile_mode(node.attribute("TileMode"))?,
            stretch: stretch_mode(node.attribute("Stretch"))?,
            alignment_x: alignment_x(node.attribute("AlignmentX"))?,
            alignment_y: alignment_y(node.attribute("AlignmentY"))?,
            transform: self.brush_transform(node, "Transform")?,
            relative_transform: self.brush_transform(node, "RelativeTransform")?,
            opacity: brush_opacity(node)?,
            children: vec![VisualBrushChild {
                bounds: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: width as f32 * 96.0 / dpi_x,
                    height: height as f32 * 96.0 / dpi_y,
                },
                visual: image,
            }],
        })
    }

    #[allow(clippy::too_many_arguments)]
    fn gradient_brush(
        &self,
        node: &Node,
        radial: bool,
        start_x: f32,
        start_y: f32,
        end_x: f32,
        end_y: f32,
        radius_x: f32,
        radius_y: f32,
    ) -> Result<Paint, Diagnostic> {
        let relative = match node
            .attribute("MappingMode")
            .unwrap_or("RelativeToBoundingBox")
        {
            "RelativeToBoundingBox" => true,
            "Absolute" => false,
            _ => {
                return Err(format_error(
                    self.page_part,
                    "invalid XPS gradient MappingMode",
                ));
            }
        };
        let spread = match node.attribute("SpreadMethod").unwrap_or("Pad") {
            "Pad" => GradientSpread::Pad,
            "Reflect" => GradientSpread::Reflect,
            "Repeat" => GradientSpread::Repeat,
            _ => {
                return Err(format_error(
                    self.page_part,
                    "invalid XPS gradient SpreadMethod",
                ));
            }
        };
        let linear_rgb = match node
            .attribute("ColorInterpolationMode")
            .unwrap_or("SRgbLinearInterpolation")
        {
            "SRgbLinearInterpolation" => false,
            "ScRgbLinearInterpolation" => true,
            _ => {
                return Err(format_error(
                    self.page_part,
                    "invalid XPS gradient ColorInterpolationMode",
                ));
            }
        };
        Ok(Paint::XpsGradient {
            radial,
            start_x,
            start_y,
            end_x,
            end_y,
            radius_x,
            radius_y,
            relative,
            spread,
            linear_rgb,
            transform: self.brush_transform(node, "Transform")?,
            relative_transform: self.brush_transform(node, "RelativeTransform")?,
            stops: self.gradient_stops(node)?,
        })
    }

    fn visual_brush(&mut self, node: &Node) -> Result<Paint, Diagnostic> {
        let viewbox = brush_rect(node.attribute("Viewbox").unwrap_or("0,0,1,1"), "Viewbox")?;
        let viewport = brush_rect(node.attribute("Viewport").unwrap_or("0,0,1,1"), "Viewport")?;
        let viewbox_relative = brush_units(node.attribute("ViewboxUnits"), "ViewboxUnits")?;
        let viewport_relative = brush_units(node.attribute("ViewportUnits"), "ViewportUnits")?;
        let tile_mode = tile_mode(node.attribute("TileMode"))?;
        let stretch = stretch_mode(node.attribute("Stretch"))?;
        let alignment_x = alignment_x(node.attribute("AlignmentX"))?;
        let alignment_y = alignment_y(node.attribute("AlignmentY"))?;
        let visual = node
            .attribute("Visual")
            .and_then(static_resource_key)
            .map(|key| {
                self.resources.get(key).cloned().ok_or_else(|| {
                    format_error(self.page_part, format!("missing XPS resource {key}"))
                })
            })
            .transpose()?
            .or_else(|| {
                node.property("VisualBrush.Visual")
                    .and_then(|property| property.children.first())
                    .cloned()
            });
        let children = visual
            .as_ref()
            .map(|visual| self.capture_visual(visual))
            .transpose()?
            .unwrap_or_default();
        Ok(Paint::Visual {
            viewbox,
            viewport,
            viewbox_relative,
            viewport_relative,
            tile_mode,
            stretch,
            alignment_x,
            alignment_y,
            transform: self.brush_transform(node, "Transform")?,
            relative_transform: self.brush_transform(node, "RelativeTransform")?,
            opacity: brush_opacity(node)?,
            children,
        })
    }

    fn omit_unknown_visual(&mut self, node: &Node) {
        // Unknown wrappers may carry transforms or clipping, so omit their entire subtree.
        self.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::UnsupportedFeature,
                Phase::Parse,
                Fidelity::Omitted,
                format!(
                    "omitted unsupported XPS visual element {} and its subtree",
                    node.name
                ),
            )
            .in_part(&node.source_part),
        );
    }

    fn capture_visual(&mut self, node: &Node) -> Result<Vec<VisualBrushChild>, Diagnostic> {
        let start = self.objects.len();
        self.render_visual(node, None, "VisualBrush")?;
        let captured = self.objects.drain(start..).collect::<Vec<_>>();
        captured_visual_children(&captured, None, self.page_part)
    }

    fn brush_transform(&self, node: &Node, property: &str) -> Result<AffineTransform, Diagnostic> {
        if let Some(value) = node.attribute(property) {
            if let Some(key) = static_resource_key(value) {
                let resource = self.resources.get(key).ok_or_else(|| {
                    format_error(self.page_part, format!("missing XPS resource {key}"))
                })?;
                return matrix_transform(resource);
            }
            return matrix(value);
        }
        node.property(&format!("{}.{}", node.local_name(), property))
            .and_then(|property| property.children.first())
            .map_or(Ok(AffineTransform::IDENTITY), matrix_transform)
    }

    fn common_visual(
        &mut self,
        node: &Node,
        mut visual: Visual,
        bounds: Rect,
    ) -> Result<Visual, Diagnostic> {
        let clip_node = Node {
            source_part: node.source_part.clone(),
            name: "Path".to_owned(),
            attributes: node
                .attribute("Clip")
                .map(|value| vec![("Data".to_owned(), value.to_owned())])
                .unwrap_or_default(),
            children: node
                .property(&format!("{}.Clip", node.local_name()))
                .map(|property| {
                    vec![Node {
                        source_part: node.source_part.clone(),
                        name: "Path.Data".to_owned(),
                        attributes: Vec::new(),
                        children: property.children.clone(),
                    }]
                })
                .unwrap_or_default(),
        };
        if let Some(mut clip) = self.geometry(&clip_node, false)? {
            rebase_geometry(&mut clip, bounds.x, bounds.y);
            visual = Visual::Effect {
                shadow: None,
                clip: Some(clip),
                visual: Box::new(visual),
            };
        }
        if let Some(mask) = self.opacity_mask(node)? {
            visual = Visual::OpacityMask {
                mask: mask.localize_gradient(bounds.x, bounds.y),
                visual: Box::new(visual),
            };
        }
        let opacity = optional_finite_attribute(node, "Opacity")?
            .unwrap_or(1.0)
            .clamp(0.0, 1.0);
        if opacity != 1.0 {
            visual = Visual::OpacityMask {
                mask: Paint::Solid(0xffffff00 | (opacity * 255.0).round() as u32),
                visual: Box::new(visual),
            };
        }
        let transform = self.brush_transform(node, "RenderTransform")?;
        if transform != AffineTransform::IDENTITY {
            visual = Visual::Layer {
                transform,
                opacity: 1.0,
                blend_mode: BlendMode::Normal,
                visual: Box::new(visual),
            };
        }
        Ok(visual)
    }

    fn opacity_mask(&mut self, node: &Node) -> Result<Option<Paint>, Diagnostic> {
        let present = node.attribute("OpacityMask").is_some()
            || node
                .property(&format!("{}.OpacityMask", node.local_name()))
                .is_some();
        if !present {
            return Ok(None);
        }
        let paint = self.paint(node, "OpacityMask")?;
        if matches!(paint, Paint::None) {
            return Err(format_error(self.page_part, "XPS OpacityMask has no brush"));
        }
        Ok(Some(paint))
    }

    fn parse_color_value(&self, value: &str, source_part: &str) -> Result<XpsColor, Diagnostic> {
        if let Some(value) = parse_color(value) {
            return Ok(XpsColor::Rgba(value));
        }
        let rest = value
            .trim()
            .strip_prefix("ContextColor ")
            .ok_or_else(|| format_error(self.page_part, "invalid XPS color"))?;
        let (profile_uri, channels) = rest
            .split_once(char::is_whitespace)
            .ok_or_else(|| format_error(self.page_part, "invalid XPS ContextColor"))?;
        let values = channels
            .split(|character: char| character == ',' || character.is_ascii_whitespace())
            .filter(|value| !value.is_empty())
            .map(|value| {
                parse_finite(value)
                    .ok_or_else(|| format_error(self.page_part, "invalid XPS ContextColor channel"))
            })
            .collect::<Result<Vec<_>, _>>()?;
        if values.len() < 2 {
            return Err(format_error(
                self.page_part,
                "XPS ContextColor has no color channels",
            ));
        }
        let target = resolve_part(source_part, profile_uri)?;
        let bytes = self
            .package
            .part(&target)?
            .ok_or_else(|| format_error(&target, "XPS color profile is missing"))?;
        let channel_count = icc_channel_count(&bytes)
            .ok_or_else(|| format_error(&target, "XPS color profile is invalid or unsupported"))?;
        let source = &values[1..];
        if source.len() != channel_count {
            return Err(format_error(
                self.page_part,
                "XPS ContextColor channel count does not match its profile",
            ));
        }
        Ok(XpsColor::Context {
            alpha: values[0].clamp(0.0, 1.0),
            profile: bytes.into_vec(),
            channels: source.to_vec(),
        })
    }

    fn gradient_stops(&self, node: &Node) -> Result<Vec<XpsGradientStop>, Diagnostic> {
        let mut nodes = Vec::new();
        fn visit<'a>(node: &'a Node, nodes: &mut Vec<&'a Node>) {
            if node.local_name() == "GradientStop" {
                nodes.push(node);
            }
            for child in &node.children {
                visit(child, nodes);
            }
        }
        visit(node, &mut nodes);
        let opacity = brush_opacity(node)?;
        let mut stops = nodes
            .into_iter()
            .map(|node| {
                let offset = node
                    .attribute("Offset")
                    .and_then(parse_finite)
                    .ok_or_else(|| format_error("", "invalid XPS gradient stop offset"))?;
                let color = self.parse_color_value(
                    node.attribute("Color")
                        .ok_or_else(|| format_error("", "invalid XPS gradient stop color"))?,
                    &node.source_part,
                )?;
                Ok(XpsGradientStop {
                    offset,
                    color: xps_color_opacity(color, opacity),
                })
            })
            .collect::<Result<Vec<_>, Diagnostic>>()?;
        if stops.is_empty() {
            return Err(format_error("", "XPS gradient brush has no gradient stops"));
        }
        stops.sort_by(|left, right| left.offset.total_cmp(&right.offset));
        Ok(stops)
    }

    fn load_font(&mut self, uri: &str, source_part: &str) -> Result<(String, u32), Diagnostic> {
        let (raw, fragment) = uri
            .split_once('#')
            .map_or((uri, None), |(raw, fragment)| (raw, Some(fragment)));
        let face = fragment
            .map(|value| {
                value
                    .parse::<u32>()
                    .map_err(|_| format_error(self.page_part, "invalid XPS font face fragment"))
            })
            .transpose()?
            .unwrap_or(0);
        let family = format!(
            "xps-{}",
            raw.rsplit('/')
                .next()
                .unwrap_or("font")
                .replace(['.', '{', '}'], "-")
        );
        if raw.is_empty() || self.loaded_fonts.contains(&raw.to_ascii_lowercase()) {
            return Ok((family, face));
        }
        let target = resolve_part(source_part, raw)?;
        let bytes = self
            .package
            .part(&target)?
            .ok_or_else(|| format_error(&target, "XPS font resource is missing"))?;
        let mut bytes = bytes.into_vec();
        let content_type = self
            .content_types
            .for_part(&target)
            .ok_or_else(|| format_error(&target, "XPS font has no declared content type"))?;
        if content_type.contains(';') {
            return Err(format_error(
                &target,
                "XPS font content type cannot contain parameters",
            ));
        }
        if content_type == OBFUSCATED_FONT_CONTENT_TYPE {
            let key = font_obfuscation_key(&target)
                .ok_or_else(|| format_error(&target, "XPS obfuscated font URI has no GUID"))?;
            if bytes.len() < 32 {
                return Err(format_error(&target, "XPS obfuscated font is truncated"));
            }
            for index in 0..32 {
                bytes[index] ^= key[index % 16];
            }
        } else if content_type != FONT_CONTENT_TYPE {
            return Err(format_error(&target, "invalid XPS font content type"));
        }
        let total = self
            .font_bytes
            .checked_add(bytes.len())
            .ok_or_else(|| format_error(&target, "XPS embedded font byte count overflowed"))?;
        if total > self.package.limits().max_font_bytes {
            return Err(Diagnostic::fatal(
                DiagnosticCode::FontBytesLimit,
                Phase::Parse,
                None,
                "XPS embedded fonts exceed the configured font limit",
            )
            .in_part(&target));
        }
        *self.font_bytes = total;
        self.loaded_fonts.insert(raw.to_ascii_lowercase());
        self.fonts.push(EmbeddedFont {
            family: family.clone(),
            bytes,
            style: FontStyle::Normal,
            weight: 400,
        });
        Ok((family, face))
    }

    fn ids(&mut self, node: &Node, prefix: &str) -> (u32, String) {
        let numeric_id = self.next_id;
        self.next_id += 1;
        let stable_id = node
            .attribute("Name")
            .filter(|name| !name.is_empty())
            .map_or_else(
                || format!("xps-{}-{prefix}-{numeric_id}", self.unit_index + 1),
                |name| format!("xps-{}-{name}", self.unit_index + 1),
            );
        (numeric_id, stable_id)
    }

    #[allow(clippy::too_many_arguments)]
    fn push_object(
        &mut self,
        numeric_id: u32,
        stable_id: String,
        parent: Option<(u32, String)>,
        kind: ObjectKind,
        bounds: Rect,
        text: Option<String>,
        source_kind: &'static str,
        path: &str,
        visual: Visual,
    ) {
        let (parent_numeric_id, parent_stable_id) =
            parent.map_or((None, None), |(id, stable)| (Some(id), Some(stable)));
        self.objects.push(Object {
            numeric_id,
            parent_numeric_id,
            stable_id,
            parent_stable_id,
            kind,
            unit_index: self.unit_index,
            bounds,
            z: self.z,
            text,
            source: SourceRef {
                part: self.page_part.to_owned(),
                mapping: MappingQuality::Exact,
                locator: SourceLocator::Xps {
                    kind: source_kind,
                    path: path.to_owned(),
                },
            },
            visual,
        });
        self.z = self.z.saturating_add(1);
    }
}

fn parse_part(package: &Package<'_>, part: &str) -> Result<Node, Diagnostic> {
    let bytes = package.required_part(part)?;
    let xml = xps_xml_bytes(&bytes).map_err(|message| format_error(part, message))?;
    let mut stack = Vec::<Node>::new();
    let mut root = None;
    parse_xml(&xml, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let mut node = Node {
                    source_part: part.to_owned(),
                    name: name.to_owned(),
                    attributes: Vec::with_capacity(attributes.len()),
                    children: Vec::new(),
                };
                for attribute in attributes {
                    node.attributes.push((
                        attribute.name.to_owned(),
                        decode_xml_text(attribute.value)
                            .map_err(|error| error.in_part(part))?
                            .into_owned(),
                    ));
                }
                if empty {
                    attach_node(&mut stack, &mut root, node, part)?;
                } else {
                    stack.push(node);
                }
            }
            XmlEvent::EndElement { .. } => {
                let node = stack
                    .pop()
                    .ok_or_else(|| format_error(part, "unbalanced XPS XML element"))?;
                attach_node(&mut stack, &mut root, node, part)?;
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| {
        if error.location.part.is_some() {
            error
        } else {
            error.in_part(part)
        }
    })?;
    let root = root.ok_or_else(|| format_error(part, "XPS XML part has no root element"))?;
    let mut roots = compatible_nodes(root, &HashMap::new())?;
    if roots.len() != 1 {
        return Err(format_error(
            part,
            "XPS compatibility selection must produce one root",
        ));
    }
    Ok(roots.remove(0))
}

// XPS selects complete markup branches; DrawingML's chart-only streaming
// selection has a different fallback contract and cannot replace this traversal.
fn compatible_nodes(
    mut node: Node,
    inherited: &HashMap<String, String>,
) -> Result<Vec<Node>, Diagnostic> {
    let mut namespaces = inherited.clone();
    for (name, value) in &node.attributes {
        if let Some(prefix) = name.strip_prefix("xmlns:") {
            namespaces.insert(prefix.to_owned(), value.clone());
        }
    }
    let is_mc = |node: &Node, local: &str| {
        node.name.split_once(':').is_some_and(|(prefix, name)| {
            name == local
                && namespaces.get(prefix).is_some_and(|uri| {
                    uri == "http://schemas.openxmlformats.org/markup-compatibility/2006"
                })
        })
    };
    if is_mc(&node, "AlternateContent") {
        let branch = node
            .children
            .iter()
            .find(|child| {
                is_mc(child, "Choice")
                    && child.attribute("Requires").is_some_and(|requires| {
                        !requires.is_empty()
                            && requires.split_whitespace().all(|prefix| {
                                let uri = child
                                    .attributes
                                    .iter()
                                    .find(|(name, _)| name == &format!("xmlns:{prefix}"))
                                    .map(|(_, value)| value)
                                    .or_else(|| namespaces.get(prefix));
                                uri.is_some_and(|uri| {
                                    matches!(
                                        uri.as_str(),
                                        "http://schemas.microsoft.com/xps/2005/06"
                                            | "http://schemas.openxps.org/oxps/v1.0"
                                    )
                                })
                            })
                    })
            })
            .or_else(|| node.children.iter().find(|child| is_mc(child, "Fallback")));
        let Some(branch) = branch else {
            return Ok(Vec::new());
        };
        let selected = compatible_nodes(branch.clone(), &namespaces)?;
        return Ok(selected
            .into_iter()
            .flat_map(|branch| branch.children)
            .collect());
    }
    node.children = node
        .children
        .into_iter()
        .map(|child| compatible_nodes(child, &namespaces))
        .collect::<Result<Vec<_>, _>>()?
        .into_iter()
        .flatten()
        .collect();
    Ok(vec![node])
}

fn parse_fixed_page_header(package: &Package<'_>, part: &str) -> Result<Node, Diagnostic> {
    let mut bytes = package.required_part_prefix(part, FIXED_PAGE_HEADER_BYTES)?;
    if (bytes.starts_with(&[0xff, 0xfe]) || bytes.starts_with(&[0xfe, 0xff]))
        && !bytes.len().is_multiple_of(2)
    {
        bytes.pop();
    }
    let xml = xps_xml_bytes(&bytes).map_err(|message| format_error(part, message))?;
    let text = std::str::from_utf8(&xml)
        .map_err(|_| format_error(part, "invalid XPS fixed-page XML encoding"))?;
    if text.contains("AlternateContent") {
        let mut root = parse_part(package, part)?;
        root.children.clear();
        return Ok(root);
    }
    let name_at = text
        .find("FixedPage")
        .ok_or_else(|| format_error(part, "XPS fixed-page root exceeds the header scan limit"))?;
    let start = text[..name_at]
        .rfind('<')
        .ok_or_else(|| format_error(part, "invalid XPS fixed-page root"))?;
    let mut quote = None;
    let end = text[start..]
        .char_indices()
        .find_map(|(offset, character)| match character {
            '\'' | '"' if quote == Some(character) => {
                quote = None;
                None
            }
            '\'' | '"' if quote.is_none() => {
                quote = Some(character);
                None
            }
            '>' if quote.is_none() => Some(start + offset),
            _ => None,
        })
        .ok_or_else(|| format_error(part, "XPS fixed-page root exceeds the header scan limit"))?;
    let mut root = text[start..end].as_bytes().to_vec();
    root.extend_from_slice(b"/>");
    let mut parsed = None;
    parse_xml(&root, package.limits(), |event| {
        if let XmlEvent::StartElement {
            name, attributes, ..
        } = event
        {
            let mut node = Node {
                source_part: part.to_owned(),
                name: name.to_owned(),
                attributes: Vec::with_capacity(attributes.len()),
                children: Vec::new(),
            };
            for attribute in attributes {
                node.attributes.push((
                    attribute.name.to_owned(),
                    decode_xml_text(attribute.value)
                        .map_err(|error| error.in_part(part))?
                        .into_owned(),
                ));
            }
            parsed = Some(node);
        }
        Ok(())
    })
    .map_err(|error| error.in_part(part))?;
    parsed.ok_or_else(|| format_error(part, "XPS XML part has no root element"))
}

fn attach_node(
    stack: &mut [Node],
    root: &mut Option<Node>,
    node: Node,
    part: &str,
) -> Result<(), Diagnostic> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(format_error(
            part,
            "XPS XML part has multiple root elements",
        ));
    }
    Ok(())
}

fn xps_xml_bytes(bytes: &[u8]) -> Result<Vec<u8>, &'static str> {
    if bytes.starts_with(&[0xff, 0xfe]) {
        let chunks = bytes[2..].chunks_exact(2);
        if !chunks.remainder().is_empty() {
            return Err("truncated UTF-16LE XPS XML");
        }
        let units = chunks.map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
        return String::from_utf16(&units.collect::<Vec<_>>())
            .map(String::into_bytes)
            .map_err(|_| "invalid UTF-16LE XPS XML");
    }
    if bytes.starts_with(&[0xfe, 0xff]) {
        let chunks = bytes[2..].chunks_exact(2);
        if !chunks.remainder().is_empty() {
            return Err("truncated UTF-16BE XPS XML");
        }
        let units = chunks.map(|pair| u16::from_be_bytes([pair[0], pair[1]]));
        return String::from_utf16(&units.collect::<Vec<_>>())
            .map(String::into_bytes)
            .map_err(|_| "invalid UTF-16BE XPS XML");
    }
    Ok(bytes.to_vec())
}

fn resolve_part(source_part: &str, target: &str) -> Result<String, Diagnostic> {
    let target = target.trim();
    let (source, target) = if let Some(target) = target.strip_prefix('/') {
        (None, target)
    } else {
        (Some(source_part), target)
    };
    // Decode URI-unreserved bytes only. Encoded separators stay opaque and
    // decoded dot segments still pass through the shared traversal guard.
    let mut decoded = String::with_capacity(target.len());
    let mut chars = target.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '%' {
            let mut lookahead = chars.clone();
            if let Some(byte) = lookahead
                .next()
                .and_then(|c| c.to_digit(16))
                .zip(lookahead.next().and_then(|c| c.to_digit(16)))
                .map(|(high, low)| (high * 16 + low) as u8)
                .filter(|b| b.is_ascii_alphanumeric() || b"-._~".contains(b))
            {
                decoded.push(byte as char);
                chars = lookahead;
                continue;
            }
        }
        decoded.push(ch);
    }
    resolve_internal_target(source, &decoded)
        .map_err(|message| format_error(source_part, format!("invalid XPS part URI: {message}")))
}

fn required_attribute<'a>(node: &'a Node, name: &str, part: &str) -> Result<&'a str, Diagnostic> {
    node.attribute(name)
        .ok_or_else(|| format_error(part, format!("{} is missing {name}", node.local_name())))
}

fn finite_attribute(node: &Node, name: &str, part: &str) -> Result<f32, Diagnostic> {
    let value = required_attribute(node, name, part)?;
    parse_finite(value)
        .ok_or_else(|| format_error(part, format!("invalid {} {name}", node.local_name())))
}

fn optional_finite_attribute(node: &Node, name: &str) -> Result<Option<f32>, Diagnostic> {
    node.attribute(name)
        .map(|value| {
            parse_finite(value)
                .ok_or_else(|| format_error("", format!("invalid {} {name}", node.local_name())))
        })
        .transpose()
}

fn parse_finite(value: &str) -> Option<f32> {
    value
        .trim()
        .parse::<f32>()
        .ok()
        .filter(|value| value.is_finite())
}

fn static_resource_key(value: &str) -> Option<&str> {
    value
        .trim()
        .strip_prefix("{StaticResource")?
        .strip_suffix('}')
        .map(str::trim)
}

fn parse_color(value: &str) -> Option<u32> {
    let value = value.trim();
    if let Some(channels) = value.strip_prefix("sc#") {
        let channels = channels
            .split(',')
            .map(str::trim)
            .map(parse_finite)
            .collect::<Option<Vec<_>>>()?;
        let (alpha, red, green, blue) = match channels.as_slice() {
            [red, green, blue] => (1.0, *red, *green, *blue),
            [alpha, red, green, blue] => (*alpha, *red, *green, *blue),
            _ => return None,
        };
        let encode = |channel: f32| {
            let channel = channel.clamp(0.0, 1.0);
            let srgb = if channel <= 0.003_130_8 {
                channel * 12.92
            } else {
                1.055 * channel.powf(1.0 / 2.4) - 0.055
            };
            (srgb * 255.0).round() as u32
        };
        return Some(
            (encode(red) << 24)
                | (encode(green) << 16)
                | (encode(blue) << 8)
                | ((alpha.clamp(0.0, 1.0) * 255.0).round() as u32),
        );
    }
    let hex = value.strip_prefix('#')?;
    match hex.len() {
        6 => u32::from_str_radix(hex, 16)
            .ok()
            .map(|rgb| (rgb << 8) | 0xff),
        8 => u32::from_str_radix(hex, 16)
            .ok()
            .map(|argb| ((argb & 0x00ff_ffff) << 8) | (argb >> 24)),
        _ => None,
    }
}

fn point(value: &str) -> Result<(f32, f32), Diagnostic> {
    let values = number_list(value);
    if values.len() != 2 {
        return Err(format_error("", "invalid XPS point"));
    }
    Ok((values[0], values[1]))
}

fn brush_rect(value: &str, name: &str) -> Result<Rect, Diagnostic> {
    let values = number_list(value);
    if values.len() != 4 || values[2] <= 0.0 || values[3] <= 0.0 {
        return Err(format_error("", format!("invalid XPS VisualBrush {name}")));
    }
    Ok(Rect {
        x: values[0],
        y: values[1],
        width: values[2],
        height: values[3],
    })
}

fn brush_units(value: Option<&str>, name: &str) -> Result<bool, Diagnostic> {
    match value.unwrap_or("RelativeToBoundingBox") {
        "RelativeToBoundingBox" => Ok(true),
        "Absolute" => Ok(false),
        _ => Err(format_error("", format!("invalid XPS VisualBrush {name}"))),
    }
}

fn tile_mode(value: Option<&str>) -> Result<TileMode, Diagnostic> {
    match value.unwrap_or("None") {
        "None" => Ok(TileMode::None),
        "Tile" => Ok(TileMode::Tile),
        "FlipX" => Ok(TileMode::FlipX),
        "FlipY" => Ok(TileMode::FlipY),
        "FlipXY" => Ok(TileMode::FlipXY),
        _ => Err(format_error("", "invalid XPS tile brush TileMode")),
    }
}

fn stretch_mode(value: Option<&str>) -> Result<StretchMode, Diagnostic> {
    match value.unwrap_or("Fill") {
        "None" => Ok(StretchMode::None),
        "Fill" => Ok(StretchMode::Fill),
        "Uniform" => Ok(StretchMode::Uniform),
        "UniformToFill" => Ok(StretchMode::UniformToFill),
        _ => Err(format_error("", "invalid XPS tile brush Stretch")),
    }
}

fn alignment_x(value: Option<&str>) -> Result<f32, Diagnostic> {
    match value.unwrap_or("Center") {
        "Left" => Ok(0.0),
        "Center" => Ok(0.5),
        "Right" => Ok(1.0),
        _ => Err(format_error("", "invalid XPS tile brush AlignmentX")),
    }
}

fn alignment_y(value: Option<&str>) -> Result<f32, Diagnostic> {
    match value.unwrap_or("Center") {
        "Top" => Ok(0.0),
        "Center" => Ok(0.5),
        "Bottom" => Ok(1.0),
        _ => Err(format_error("", "invalid XPS tile brush AlignmentY")),
    }
}

fn matrix(value: &str) -> Result<AffineTransform, Diagnostic> {
    let values = number_list(value);
    if values.len() != 6 {
        return Err(format_error("", "invalid XPS brush transform"));
    }
    let transform = AffineTransform {
        a: values[0],
        b: values[1],
        c: values[2],
        d: values[3],
        e: values[4],
        f: values[5],
    };
    transform
        .is_valid()
        .then_some(transform)
        .ok_or_else(|| format_error("", "invalid XPS brush transform"))
}

fn matrix_transform(node: &Node) -> Result<AffineTransform, Diagnostic> {
    if node.local_name() != "MatrixTransform" {
        return Err(format_error(
            "",
            "XPS brush transform is not a MatrixTransform",
        ));
    }
    matrix(node.attribute("Matrix").unwrap_or("1,0,0,1,0,0"))
}

fn wrap_group_visual(wrapper: &Visual, child: Visual) -> Result<Visual, Diagnostic> {
    match wrapper {
        Visual::None => Ok(child),
        Visual::Layer {
            transform,
            opacity,
            blend_mode,
            visual,
        } => Ok(Visual::Layer {
            transform: *transform,
            opacity: *opacity,
            blend_mode: *blend_mode,
            visual: Box::new(wrap_group_visual(visual, child)?),
        }),
        Visual::Effect {
            shadow,
            clip,
            visual,
        } => Ok(Visual::Effect {
            shadow: *shadow,
            clip: clip.clone(),
            visual: Box::new(wrap_group_visual(visual, child)?),
        }),
        _ => Err(format_error("", "unsupported XPS Canvas visual wrapper")),
    }
}

fn captured_visual_children(
    objects: &[Object],
    parent: Option<u32>,
    part: &str,
) -> Result<Vec<VisualBrushChild>, Diagnostic> {
    let mut children = Vec::new();
    for object in objects
        .iter()
        .filter(|object| object.parent_numeric_id == parent)
    {
        if object.kind != ObjectKind::Group || visual_contains_group(&object.visual) {
            children.push(VisualBrushChild {
                bounds: object.bounds,
                visual: object.visual.clone(),
            });
            continue;
        }
        for child in captured_visual_children(objects, Some(object.numeric_id), part)? {
            children.push(VisualBrushChild {
                bounds: child.bounds,
                visual: wrap_group_visual(&object.visual, child.visual)
                    .map_err(|_| format_error(part, "invalid XPS visual hierarchy"))?,
            });
        }
    }
    Ok(children)
}

fn visual_contains_group(visual: &Visual) -> bool {
    match visual {
        Visual::Group { .. } | Visual::OpacityMask { .. } => true,
        Visual::Layer { visual, .. } | Visual::Effect { visual, .. } => {
            visual_contains_group(visual)
        }
        _ => false,
    }
}

fn number_list(value: &str) -> Vec<f32> {
    value
        .split(|character: char| character == ',' || character.is_ascii_whitespace())
        .filter(|value| !value.is_empty())
        .filter_map(parse_finite)
        .collect()
}

fn brush_opacity(node: &Node) -> Result<f32, Diagnostic> {
    optional_finite_attribute(node, "Opacity").map(|value| value.unwrap_or(1.0).clamp(0.0, 1.0))
}

fn xps_color_opacity(color: XpsColor, opacity: f32) -> XpsColor {
    match color {
        XpsColor::Rgba(color) => {
            XpsColor::Rgba((color & 0xffff_ff00) | ((color & 0xff) as f32 * opacity).round() as u32)
        }
        XpsColor::Context {
            alpha,
            profile,
            channels,
        } => XpsColor::Context {
            alpha: alpha * opacity,
            profile,
            channels,
        },
    }
}

fn brush_color_opacity(node: &Node, color: XpsColor) -> Result<XpsColor, Diagnostic> {
    Ok(xps_color_opacity(color, brush_opacity(node)?))
}

fn xps_color_paint(color: XpsColor) -> Paint {
    if let XpsColor::Rgba(color) = color {
        return Paint::Solid(color);
    }
    Paint::XpsGradient {
        radial: false,
        start_x: 0.0,
        start_y: 0.0,
        end_x: 1.0,
        end_y: 0.0,
        radius_x: 0.0,
        radius_y: 0.0,
        relative: true,
        spread: GradientSpread::Pad,
        linear_rgb: false,
        transform: AffineTransform::IDENTITY,
        relative_transform: AffineTransform::IDENTITY,
        stops: vec![XpsGradientStop { offset: 0.0, color }],
    }
}

fn icc_channel_count(bytes: &[u8]) -> Option<usize> {
    if bytes.len() < 128 || bytes.get(36..40)? != b"acsp" {
        return None;
    }
    let declared = u32::from_be_bytes(bytes.get(..4)?.try_into().ok()?) as usize;
    if !(128..=bytes.len()).contains(&declared) {
        return None;
    }
    match bytes.get(16..20)? {
        b"GRAY" => Some(1),
        b"RGB " | b"CMY " => Some(3),
        b"CMYK" => Some(4),
        [count @ b'2'..=b'9', b'C', b'L', b'R'] => Some(usize::from(count - b'0')),
        [count @ (b'A'..=b'F'), b'C', b'L', b'R'] => Some(usize::from(count - b'A' + 10)),
        [b'M', b'C', b'H', count @ b'2'..=b'9'] => Some(usize::from(count - b'0')),
        [b'M', b'C', b'H', count @ (b'A'..=b'F')] => Some(usize::from(count - b'A' + 10)),
        _ => None,
    }
}

fn color_converted_bitmap_source(
    value: &str,
) -> Result<(&str, Option<&str>, Option<&str>), Diagnostic> {
    let value = value.trim();
    if let Some(rest) = value.strip_prefix("{ColorConvertedBitmap") {
        let parts = rest
            .trim_end_matches('}')
            .split_ascii_whitespace()
            .collect::<Vec<_>>();
        if !(2..=3).contains(&parts.len()) {
            return Err(format_error("", "invalid XPS ColorConvertedBitmap markup"));
        }
        Ok((parts[0], Some(parts[1]), parts.get(2).copied()))
    } else {
        Ok((value, None, None))
    }
}

fn image_media_type(part: &str, bytes: &[u8]) -> Option<&'static str> {
    let extension = part
        .rsplit_once('.')
        .map(|(_, extension)| extension.to_ascii_lowercase());
    match extension.as_deref() {
        Some("png") if bytes.starts_with(b"\x89PNG\r\n\x1a\n") => Some("image/png"),
        Some("jpg" | "jpeg" | "jpe" | "jfif") if bytes.starts_with(b"\xff\xd8") => {
            Some("image/jpeg")
        }
        Some("tif" | "tiff") if bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*") => {
            Some("image/tiff")
        }
        Some("jxr" | "wdp" | "hdp") if bytes.starts_with(b"II\xbc\x01") => {
            Some("image/vnd.ms-photo")
        }
        _ => None,
    }
}

// Absolute image viewboxes use physical 1/96-inch units, not source pixels.
fn xps_image_resolution(media_type: &str, bytes: &[u8]) -> (f32, f32) {
    fn tiff(bytes: &[u8]) -> Option<(f32, f32)> {
        let little = bytes.starts_with(b"II");
        if !little && !bytes.starts_with(b"MM") {
            return None;
        }
        let u16_at = |at: usize| -> Option<u16> {
            let b = bytes.get(at..at.checked_add(2)?)?.try_into().ok()?;
            Some(if little {
                u16::from_le_bytes(b)
            } else {
                u16::from_be_bytes(b)
            })
        };
        let u32_at = |at: usize| -> Option<u32> {
            let b = bytes.get(at..at.checked_add(4)?)?.try_into().ok()?;
            Some(if little {
                u32::from_le_bytes(b)
            } else {
                u32::from_be_bytes(b)
            })
        };
        let ifd = u32_at(4)? as usize;
        let count = u16_at(ifd)? as usize;
        let mut x = None;
        let mut y = None;
        let mut unit = 2;
        for i in 0..count {
            let at = ifd.checked_add(2 + i * 12)?;
            match (u16_at(at)?, u16_at(at + 2)?) {
                (tag @ (282 | 283), 5) if u32_at(at + 4)? == 1 => {
                    let offset = u32_at(at + 8)? as usize;
                    let value = u32_at(offset)? as f32 / u32_at(offset.checked_add(4)?)? as f32;
                    if tag == 282 {
                        x = Some(value);
                    } else {
                        y = Some(value);
                    }
                }
                (296, 3) => unit = u16_at(at + 8)?,
                _ => {}
            }
        }
        let scale = match unit {
            2 => 1.0,
            3 => 2.54,
            _ => return None,
        };
        Some((x? * scale, y? * scale))
    }
    let resolution = (|| -> Option<(f32, f32)> {
        if media_type == "image/tiff" {
            return tiff(bytes);
        }
        if media_type == "image/png" {
            let mut at = 8usize;
            while let Some(header) = bytes.get(at..at.checked_add(8)?) {
                let length = u32::from_be_bytes(header[..4].try_into().ok()?) as usize;
                let data = bytes.get(at + 8..at.checked_add(8)?.checked_add(length)?)?;
                if &header[4..8] == b"pHYs" && data.len() == 9 && data[8] == 1 {
                    return Some((
                        u32::from_be_bytes(data[..4].try_into().ok()?) as f32 * 0.0254,
                        u32::from_be_bytes(data[4..8].try_into().ok()?) as f32 * 0.0254,
                    ));
                }
                at = at.checked_add(12)?.checked_add(length)?;
            }
        }
        if media_type == "image/jpeg" {
            let mut at = 2usize;
            let mut jfif = None;
            loop {
                while bytes.get(at) == Some(&255) {
                    at += 1;
                }
                let marker = *bytes.get(at)?;
                at += 1;
                if marker == 0xda || marker == 0xd9 {
                    return jfif;
                }
                if marker == 1 || (0xd0..=0xd7).contains(&marker) {
                    continue;
                }
                let length = u16::from_be_bytes(bytes.get(at..at + 2)?.try_into().ok()?) as usize;
                if length < 2 {
                    return jfif;
                }
                let data = bytes.get(at + 2..at.checked_add(length)?)?;
                if marker == 0xe1 && data.starts_with(b"Exif\0\0") {
                    if let Some(value) = tiff(&data[6..]) {
                        return Some(value);
                    }
                }
                if marker == 0xe0 && data.starts_with(b"JFIF\0") && data.len() >= 12 {
                    let scale = match data[7] {
                        1 => 1.0,
                        2 => 2.54,
                        _ => 0.0,
                    };
                    jfif = Some((
                        u16::from_be_bytes([data[8], data[9]]) as f32 * scale,
                        u16::from_be_bytes([data[10], data[11]]) as f32 * scale,
                    ));
                }
                at = at.checked_add(length)?;
            }
        }
        None
    })();
    resolution
        .filter(|(x, y)| x.is_finite() && y.is_finite() && *x > 0.0 && *y > 0.0)
        .unwrap_or((96.0, 96.0))
}

fn xps_image_dimensions(media_type: &str, bytes: &[u8]) -> Option<(u32, u32)> {
    let dimensions = match media_type {
        "image/png" => Some((
            u32::from_be_bytes(bytes.get(16..20)?.try_into().ok()?),
            u32::from_be_bytes(bytes.get(20..24)?.try_into().ok()?),
        )),
        "image/jpeg" => {
            let mut cursor = 2_usize;
            loop {
                while bytes.get(cursor) == Some(&0xff) {
                    cursor += 1;
                }
                let marker = *bytes.get(cursor)?;
                cursor += 1;
                if marker == 0xd9 || marker == 0xda {
                    return None;
                }
                if marker == 0x01 || (0xd0..=0xd7).contains(&marker) {
                    continue;
                }
                let length = usize::from(u16::from_be_bytes(
                    bytes.get(cursor..cursor + 2)?.try_into().ok()?,
                ));
                let segment = bytes.get(cursor + 2..cursor.checked_add(length)?)?;
                if matches!(
                    marker,
                    0xc0 | 0xc1
                        | 0xc2
                        | 0xc3
                        | 0xc5
                        | 0xc6
                        | 0xc7
                        | 0xc9
                        | 0xca
                        | 0xcb
                        | 0xcd
                        | 0xce
                        | 0xcf
                ) {
                    break Some((
                        u32::from(u16::from_be_bytes(segment.get(3..5)?.try_into().ok()?)),
                        u32::from(u16::from_be_bytes(segment.get(1..3)?.try_into().ok()?)),
                    ));
                }
                cursor = cursor.checked_add(length)?;
            }
        }
        "image/tiff" | "image/vnd.ms-photo" => {
            let little = bytes.starts_with(b"II");
            let read_u16 = |slice: &[u8]| {
                Some(if little {
                    u16::from_le_bytes(slice.try_into().ok()?)
                } else {
                    u16::from_be_bytes(slice.try_into().ok()?)
                })
            };
            let read_u32 = |slice: &[u8]| {
                Some(if little {
                    u32::from_le_bytes(slice.try_into().ok()?)
                } else {
                    u32::from_be_bytes(slice.try_into().ok()?)
                })
            };
            let offset = usize::try_from(read_u32(bytes.get(4..8)?)?).ok()?;
            let count = usize::from(read_u16(bytes.get(offset..offset + 2)?)?);
            let mut width = None;
            let mut height = None;
            for index in 0..count {
                let entry = bytes.get(offset + 2 + index * 12..offset + 14 + index * 12)?;
                let tag = read_u16(&entry[0..2])?;
                let kind = read_u16(&entry[2..4])?;
                let value = if kind == 3 {
                    u32::from(read_u16(&entry[8..10])?)
                } else {
                    read_u32(&entry[8..12])?
                };
                if tag == 256 || tag == 0xbc80 {
                    width = Some(value);
                }
                if tag == 257 || tag == 0xbc81 {
                    height = Some(value);
                }
            }
            Some((width?, height?))
        }
        _ => None,
    }?;
    (dimensions.0 > 0 && dimensions.1 > 0).then_some(dimensions)
}

fn geometry_bounds(geometry: &Geometry) -> Option<Rect> {
    let Geometry::Path { commands, .. } = geometry else {
        return None;
    };
    let mut bounds = None;
    let mut current = (0.0_f32, 0.0_f32);
    let mut start = current;
    for command in commands {
        match *command {
            PathCommand::MoveTo { x, y } => {
                include_point(&mut bounds, x, y);
                current = (x, y);
                start = current;
            }
            PathCommand::LineTo { x, y } => {
                include_point(&mut bounds, x, y);
                current = (x, y);
            }
            PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => {
                include_point(&mut bounds, x, y);
                for (p0, control, p1) in [(current.0, cpx, x), (current.1, cpy, y)] {
                    let denominator = p0 - 2.0 * control + p1;
                    if denominator != 0.0 {
                        let t = (p0 - control) / denominator;
                        if (0.0..1.0).contains(&t) {
                            include_point(
                                &mut bounds,
                                quadratic(current.0, cpx, x, t),
                                quadratic(current.1, cpy, y, t),
                            );
                        }
                    }
                }
                current = (x, y);
            }
            PathCommand::BezierCurveTo {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                include_point(&mut bounds, x, y);
                for t in cubic_extrema(current.0, cp1x, cp2x, x) {
                    include_point(
                        &mut bounds,
                        cubic(current.0, cp1x, cp2x, x, t),
                        cubic(current.1, cp1y, cp2y, y, t),
                    );
                }
                for t in cubic_extrema(current.1, cp1y, cp2y, y) {
                    include_point(
                        &mut bounds,
                        cubic(current.0, cp1x, cp2x, x, t),
                        cubic(current.1, cp1y, cp2y, y, t),
                    );
                }
                current = (x, y);
            }
            PathCommand::ClosePath => {
                include_point(&mut bounds, start.0, start.1);
                current = start;
            }
        }
    }
    bounds.map(|(min_x, min_y, max_x, max_y)| Rect {
        x: min_x,
        y: min_y,
        width: max_x - min_x,
        height: max_y - min_y,
    })
}

fn include_point(bounds: &mut Option<(f32, f32, f32, f32)>, x: f32, y: f32) {
    *bounds = Some(bounds.map_or((x, y, x, y), |(min_x, min_y, max_x, max_y)| {
        (min_x.min(x), min_y.min(y), max_x.max(x), max_y.max(y))
    }));
}

fn quadratic(p0: f32, p1: f32, p2: f32, t: f32) -> f32 {
    let inverse = 1.0 - t;
    inverse * inverse * p0 + 2.0 * inverse * t * p1 + t * t * p2
}

fn cubic(p0: f32, p1: f32, p2: f32, p3: f32, t: f32) -> f32 {
    let inverse = 1.0 - t;
    inverse.powi(3) * p0
        + 3.0 * inverse * inverse * t * p1
        + 3.0 * inverse * t * t * p2
        + t.powi(3) * p3
}

fn cubic_extrema(p0: f32, p1: f32, p2: f32, p3: f32) -> Vec<f32> {
    let a = -p0 + 3.0 * p1 - 3.0 * p2 + p3;
    let b = 2.0 * (p0 - 2.0 * p1 + p2);
    let c = p1 - p0;
    if a.abs() < f32::EPSILON {
        return (b.abs() >= f32::EPSILON)
            .then_some(-c / b)
            .filter(|t| (0.0..1.0).contains(t))
            .into_iter()
            .collect();
    }
    let discriminant = b * b - 4.0 * a * c;
    if discriminant < 0.0 {
        return Vec::new();
    }
    let root = discriminant.sqrt();
    [(-b + root) / (2.0 * a), (-b - root) / (2.0 * a)]
        .into_iter()
        .filter(|t| (0.0..1.0).contains(t))
        .collect()
}

fn rebase_geometry(geometry: &mut Geometry, x: f32, y: f32) {
    match geometry {
        Geometry::Path { commands, .. } => rebase_path(commands, x, y),
        Geometry::LayeredPath { layers } => {
            for layer in layers {
                rebase_path(&mut layer.commands, x, y);
            }
        }
        _ => {}
    }
}

#[derive(Clone, Copy)]
struct GlyphStyle {
    bold: bool,
    italic: bool,
}

impl GlyphStyle {
    fn parse(value: Option<&str>, part: &str) -> Result<Self, Diagnostic> {
        match value.unwrap_or("None") {
            "None" => Ok(Self {
                bold: false,
                italic: false,
            }),
            "BoldSimulation" => Ok(Self {
                bold: true,
                italic: false,
            }),
            "ItalicSimulation" => Ok(Self {
                bold: false,
                italic: true,
            }),
            "BoldItalicSimulation" => Ok(Self {
                bold: true,
                italic: true,
            }),
            _ => Err(format_error(part, "invalid XPS StyleSimulations value")),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
struct GlyphIndex {
    cluster: bool,
    code_units: usize,
    glyphs: usize,
    id: Option<u16>,
    advance: Option<f32>,
    u_offset: f32,
    v_offset: f32,
}

fn parse_glyph_indices(value: &str) -> Result<Vec<GlyphIndex>, Diagnostic> {
    if value.is_empty() {
        return Ok(Vec::new());
    }
    value
        .split(';')
        .map(|entry| {
            let mut entry = entry.trim();
            let mut has_cluster = false;
            let mut code_units = 1_usize;
            let mut glyphs = 1_usize;
            if let Some(cluster) = entry.strip_prefix('(') {
                has_cluster = true;
                let end = cluster
                    .find(')')
                    .ok_or_else(|| format_error("", "invalid XPS glyph cluster"))?;
                let (left, right) = cluster[..end]
                    .split_once(':')
                    .map_or((&cluster[..end], None), |(left, right)| (left, Some(right)));
                code_units = left
                    .parse()
                    .ok()
                    .filter(|value| *value > 0)
                    .ok_or_else(|| format_error("", "invalid XPS glyph cluster code-unit count"))?;
                if let Some(right) = right {
                    glyphs = right
                        .parse()
                        .ok()
                        .filter(|value| *value > 0)
                        .ok_or_else(|| format_error("", "invalid XPS glyph cluster glyph count"))?;
                }
                entry = &cluster[end + 1..];
            }
            let fields = entry.split(',').map(str::trim).collect::<Vec<_>>();
            if fields.len() > 4 {
                return Err(format_error("", "invalid XPS glyph index field count"));
            }
            let number = |index: usize| -> Result<Option<f32>, Diagnostic> {
                match fields.get(index).copied().unwrap_or("") {
                    "" => Ok(None),
                    value => parse_finite(value)
                        .map(Some)
                        .ok_or_else(|| format_error("", "invalid XPS glyph index number")),
                }
            };
            let advance = number(1)?;
            if advance.is_some_and(|value| value < 0.0) {
                return Err(format_error("", "XPS glyph advance must be non-negative"));
            }
            Ok(GlyphIndex {
                cluster: has_cluster,
                code_units,
                glyphs,
                id: match fields.first().copied().unwrap_or("") {
                    "" => None,
                    value => Some(
                        value
                            .parse::<u16>()
                            .map_err(|_| format_error("", "invalid XPS glyph ID"))?,
                    ),
                },
                advance,
                u_offset: number(2)?.unwrap_or(0.0),
                v_offset: number(3)?.unwrap_or(0.0),
            })
        })
        .collect()
}

struct XpsGlyphPen {
    commands: Vec<PathCommand>,
    scale: f32,
    x: f32,
    y: f32,
    sideways: bool,
    top_origin_x: f32,
    top_origin_y: f32,
    italic: bool,
    bounds: Option<(f32, f32, f32, f32)>,
}

impl XpsGlyphPen {
    fn map(&self, x: f32, y: f32) -> (f32, f32) {
        if self.sideways {
            let y = if self.italic {
                y - x * 20_f32.to_radians().tan()
            } else {
                y
            };
            (
                self.x + (self.top_origin_y - y) * self.scale,
                self.y + (self.top_origin_x - x) * self.scale,
            )
        } else {
            let x = if self.italic {
                x + y * 20_f32.to_radians().tan()
            } else {
                x
            };
            (self.x + x * self.scale, self.y - y * self.scale)
        }
    }

    fn point(&mut self, x: f32, y: f32) -> (f32, f32) {
        let point = self.map(x, y);
        self.bounds = Some(match self.bounds {
            Some((min_x, min_y, max_x, max_y)) => (
                min_x.min(point.0),
                min_y.min(point.1),
                max_x.max(point.0),
                max_y.max(point.1),
            ),
            None => (point.0, point.1, point.0, point.1),
        });
        point
    }
}

impl OutlinePen for XpsGlyphPen {
    fn move_to(&mut self, x: f32, y: f32) {
        let (x, y) = self.point(x, y);
        self.commands.push(PathCommand::MoveTo { x, y });
    }
    fn line_to(&mut self, x: f32, y: f32) {
        let (x, y) = self.point(x, y);
        self.commands.push(PathCommand::LineTo { x, y });
    }
    fn quad_to(&mut self, cpx: f32, cpy: f32, x: f32, y: f32) {
        let (cpx, cpy) = self.point(cpx, cpy);
        let (x, y) = self.point(x, y);
        self.commands
            .push(PathCommand::QuadraticCurveTo { cpx, cpy, x, y });
    }
    fn curve_to(&mut self, cp1x: f32, cp1y: f32, cp2x: f32, cp2y: f32, x: f32, y: f32) {
        let (cp1x, cp1y) = self.point(cp1x, cp1y);
        let (cp2x, cp2y) = self.point(cp2x, cp2y);
        let (x, y) = self.point(x, y);
        self.commands.push(PathCommand::BezierCurveTo {
            cp1x,
            cp1y,
            cp2x,
            cp2y,
            x,
            y,
        });
    }
    fn close(&mut self) {
        self.commands.push(PathCommand::ClosePath);
    }
}

fn rebase_path(commands: &mut [PathCommand], x: f32, y: f32) {
    let transform = AffineTransform {
        e: -x,
        f: -y,
        ..AffineTransform::IDENTITY
    };
    for command in commands {
        command.transform(transform);
    }
}

fn sideways_metrics(
    font: &FontRef<'_>,
    glyph_metrics: &skrifa::metrics::GlyphMetrics<'_>,
    glyph_id: GlyphId,
) -> Result<(f32, f32, f32), Diagnostic> {
    let top_origin_x = glyph_metrics
        .advance_width(glyph_id)
        .ok_or_else(|| format_error("", "XPS sideways glyph horizontal advance is missing"))?
        / 2.0;
    if let Ok(vorg) = font.vorg() {
        let top_origin_y = f32::from(vorg.vertical_origin_y(glyph_id));
        let descender = font
            .os2()
            .map(|table| f32::from(table.s_typo_descender()).abs())
            .or_else(|_| {
                font.hhea()
                    .map(|table| f32::from(table.descender().to_i16()).abs())
            })
            .map_err(|_| format_error("", "XPS font has no sideways descender metric"))?;
        return Ok((top_origin_x, top_origin_y, top_origin_y + descender));
    }
    if let Ok(vmtx) = font.vmtx() {
        let bounds = glyph_metrics
            .bounds(glyph_id)
            .ok_or_else(|| format_error("", "XPS sideways glyph bounds are missing"))?;
        let top_origin_y = bounds.y_max
            + f32::from(vmtx.side_bearing(glyph_id).ok_or_else(|| {
                format_error("", "XPS sideways glyph top side-bearing is missing")
            })?);
        let descender = vmtx
            .advance(glyph_id)
            .map(|advance| f32::from(advance) - top_origin_y)
            .filter(|value| *value >= 0.0)
            .ok_or_else(|| format_error("", "invalid XPS sideways glyph vertical metrics"))?;
        return Ok((top_origin_x, top_origin_y, top_origin_y + descender));
    }
    let (top_origin_y, descender) = if let Ok(os2) = font.os2() {
        (
            f32::from(os2.s_typo_ascender()),
            f32::from(os2.s_typo_descender()).abs(),
        )
    } else {
        let hhea = font
            .hhea()
            .map_err(|_| format_error("", "XPS font has no sideways line metrics"))?;
        (
            f32::from(hhea.ascender().to_i16()),
            f32::from(hhea.descender().to_i16()).abs(),
        )
    };
    Ok((top_origin_x, top_origin_y, top_origin_y + descender))
}

struct XpsCharmap<'a> {
    platform: PlatformId,
    encoding: u16,
    subtable: CmapSubtable<'a>,
}

impl XpsCharmap<'_> {
    fn map(&self, character: char) -> Option<GlyphId> {
        let codepoint = match (self.platform, self.encoding) {
            (PlatformId::Windows, 2) => encoded_codepoint(encoding_rs::SHIFT_JIS, character)?,
            (PlatformId::Windows, 3) => encoded_codepoint(encoding_rs::GBK, character)?,
            (PlatformId::Windows, 4) => encoded_codepoint(encoding_rs::BIG5, character)?,
            (PlatformId::Windows, 5) => encoded_codepoint(encoding_rs::EUC_KR, character)?,
            (PlatformId::Macintosh, 0) => encoded_codepoint(encoding_rs::MACINTOSH, character)?,
            _ => character as u32,
        };
        self.subtable.map_codepoint(codepoint).or_else(|| {
            (self.platform == PlatformId::Windows
                && self.encoding == 0
                && (0x20..=0xff).contains(&codepoint))
            .then(|| self.subtable.map_codepoint(codepoint + 0xf000))
            .flatten()
        })
    }
}

fn encoded_codepoint(encoding: &'static encoding_rs::Encoding, character: char) -> Option<u32> {
    let mut text = [0_u8; 4];
    let text = character.encode_utf8(&mut text);
    let (bytes, _, had_errors) = encoding.encode(text);
    if had_errors || bytes.is_empty() || bytes.len() > 2 {
        return None;
    }
    Some(
        bytes
            .iter()
            .fold(0_u32, |value, byte| (value << 8) | u32::from(*byte)),
    )
}

fn xps_charmap<'a>(font: &FontRef<'a>) -> Option<XpsCharmap<'a>> {
    let cmap = font.cmap().ok()?;
    let records = cmap.encoding_records();
    let data = cmap.offset_data();
    let find = |platform, encoding| {
        records.iter().find_map(|record| {
            (record.platform_id() == platform && record.encoding_id() == encoding)
                .then(|| record.subtable(data).ok())
                .flatten()
                .map(|subtable| XpsCharmap {
                    platform,
                    encoding,
                    subtable,
                })
        })
    };
    for encoding in [10, 1, 5, 4, 3, 2, 0] {
        if let Some(charmap) = find(PlatformId::Windows, encoding) {
            return Some(charmap);
        }
    }
    records
        .iter()
        .find_map(|record| {
            (record.platform_id() == PlatformId::Unicode)
                .then(|| record.subtable(data).ok())
                .flatten()
                .map(|subtable| XpsCharmap {
                    platform: PlatformId::Unicode,
                    encoding: record.encoding_id(),
                    subtable,
                })
        })
        .or_else(|| find(PlatformId::Macintosh, 0))
}

fn glyph_geometry(
    bytes: &[u8],
    face: u32,
    text: &str,
    indices: Option<&str>,
    origin_x: f32,
    origin_y: f32,
    em: f32,
    right_to_left: bool,
    sideways: bool,
    style: GlyphStyle,
    command_limit: usize,
) -> Result<(Geometry, Rect), Diagnostic> {
    let font = FontRef::from_index(bytes, face)
        .map_err(|_| format_error("", "invalid XPS font or font face index"))?;
    let metrics = font.metrics(Size::unscaled(), LocationRef::default());
    if metrics.units_per_em == 0 {
        return Err(format_error("", "XPS font has zero units-per-em"));
    }
    let glyph_metrics = font.glyph_metrics(Size::unscaled(), LocationRef::default());
    let outlines = font.outline_glyphs();
    let charmap = xps_charmap(&font);
    let utf16 = text.encode_utf16().collect::<Vec<_>>();
    let mut entries = match indices {
        Some(value) if !value.trim().is_empty() => parse_glyph_indices(value)?,
        _ => Vec::new(),
    };
    if entries.len() > command_limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "XPS glyph count exceeds the configured object limit",
        ));
    }
    let mut covered = 0_usize;
    let mut remaining = 0_usize;
    for entry in &entries {
        if remaining == 0 {
            covered = covered
                .checked_add(entry.code_units)
                .ok_or_else(|| format_error("", "XPS glyph cluster size overflow"))?;
            remaining = entry.glyphs;
        } else if entry.cluster {
            return Err(format_error("", "nested XPS glyph cluster mapping"));
        }
        remaining -= 1;
    }
    if remaining != 0 {
        return Err(format_error(
            "",
            "XPS glyph cluster has too few glyph specifications",
        ));
    }
    if !text.is_empty() && covered > utf16.len() {
        return Err(format_error("", "XPS glyph cluster exceeds UnicodeString"));
    }
    if !text.is_empty() {
        let suffix = String::from_utf16(&utf16[covered..])
            .map_err(|_| format_error("", "invalid XPS UnicodeString"))?;
        entries.extend(suffix.chars().map(|character| GlyphIndex {
            cluster: character.len_utf16() != 1,
            code_units: character.len_utf16(),
            glyphs: 1,
            id: None,
            advance: None,
            u_offset: 0.0,
            v_offset: 0.0,
        }));
    } else if entries.is_empty() {
        return Err(format_error(
            "",
            "XPS Glyphs has neither UnicodeString nor glyph indices",
        ));
    }
    if entries.len() > command_limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "XPS glyph count exceeds the configured object limit",
        ));
    }
    let mut pen = XpsGlyphPen {
        commands: Vec::new(),
        scale: em / f32::from(metrics.units_per_em),
        x: origin_x,
        y: origin_y,
        sideways,
        top_origin_x: 0.0,
        top_origin_y: 0.0,
        italic: style.italic,
        bounds: None,
    };
    let mut cursor = 0.0_f32;
    let mut text_offset = 0_usize;
    let mut cluster_remaining = 0_usize;
    let mut cluster_units = 0_usize;
    let mut cluster_glyphs = 0_usize;
    let mut cluster_character = None;
    for entry in entries {
        if cluster_remaining == 0 {
            cluster_remaining = entry.glyphs;
            cluster_units = entry.code_units;
            cluster_glyphs = entry.glyphs;
            if !text.is_empty() {
                if text_offset + cluster_units > utf16.len() {
                    return Err(format_error("", "XPS glyph cluster exceeds UnicodeString"));
                }
                let cluster = String::from_utf16(&utf16[text_offset..text_offset + cluster_units])
                    .map_err(|_| format_error("", "invalid XPS glyph cluster Unicode"))?;
                cluster_character = cluster.chars().next();
            }
        } else if entry.cluster {
            return Err(format_error("", "nested XPS glyph cluster mapping"));
        }
        if entry.id.is_none() && (cluster_units != 1 || cluster_glyphs != 1) {
            return Err(format_error(
                "",
                "XPS non-one-to-one glyph clusters require explicit glyph IDs",
            ));
        }
        let glyph_id = entry
            .id
            .map(|value| GlyphId::new(u32::from(value)))
            .or_else(|| cluster_character.and_then(|character| charmap.as_ref()?.map(character)))
            .unwrap_or_else(|| GlyphId::new(0));
        let glyph = outlines
            .get(glyph_id)
            .ok_or_else(|| format_error("", "XPS glyph outline is missing"))?;
        let (top_origin_x, top_origin_y, default_advance) = if sideways {
            sideways_metrics(&font, &glyph_metrics, glyph_id)?
        } else {
            (
                0.0,
                0.0,
                glyph_metrics
                    .advance_width(glyph_id)
                    .ok_or_else(|| format_error("", "XPS glyph advance is missing"))?,
            )
        };
        let advance = entry
            .advance
            .map(|value| value * em / 100.0)
            .unwrap_or((default_advance * pen.scale) + if style.bold { em * 0.02 } else { 0.0 });
        let direction = if right_to_left { -1.0 } else { 1.0 };
        pen.x = origin_x
            + direction
                * (cursor
                    + if right_to_left { advance } else { 0.0 }
                    + entry.u_offset * em / 100.0)
            + if style.bold { em * 0.01 } else { 0.0 };
        pen.y = origin_y - entry.v_offset * em / 100.0 - if style.bold { em * 0.01 } else { 0.0 };
        pen.top_origin_x = top_origin_x;
        pen.top_origin_y = top_origin_y;
        glyph
            .draw(
                DrawSettings::unhinted(Size::unscaled(), LocationRef::default()),
                &mut pen,
            )
            .map_err(|_| format_error("", "XPS glyph outline could not be drawn"))?;
        if pen.commands.len() > command_limit {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ObjectLimit,
                Phase::Parse,
                None,
                "XPS glyph outlines exceed the configured object limit",
            ));
        }
        cursor += advance;
        cluster_remaining -= 1;
        if cluster_remaining == 0 {
            text_offset += cluster_units;
            cluster_character = None;
        }
    }
    if cluster_remaining != 0 || (!text.is_empty() && text_offset != utf16.len()) {
        return Err(format_error(
            "",
            "XPS glyph indices do not cover UnicodeString",
        ));
    }
    let stroke_padding = if style.bold { em * 0.01 } else { 0.0 };
    let (min_x, min_y, max_x, max_y) = pen
        .bounds
        .unwrap_or((origin_x, origin_y, origin_x, origin_y));
    let bounds = Rect {
        x: min_x - stroke_padding,
        y: min_y - stroke_padding,
        width: max_x - min_x + stroke_padding * 2.0,
        height: max_y - min_y + stroke_padding * 2.0,
    };
    rebase_path(&mut pen.commands, bounds.x, bounds.y);
    Ok((
        Geometry::Path {
            fill_rule: FillRule::NonZero,
            commands: pen.commands,
        },
        bounds,
    ))
}

fn font_obfuscation_key(part: &str) -> Option<[u8; 16]> {
    let stem = part.rsplit('/').next()?.rsplit_once('.')?.0;
    let compact: String = stem.chars().filter(|character| *character != '-').collect();
    if compact.len() != 32 || !compact.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    let mut key = [0_u8; 16];
    for index in 0..16 {
        key[15 - index] = u8::from_str_radix(&compact[index * 2..index * 2 + 2], 16).ok()?;
    }
    Some(key)
}

fn stroke_extensions(
    node: &Node,
    geometry: &Geometry,
    width: f32,
    limit: usize,
) -> Result<Option<Geometry>, Diagnostic> {
    let dashed = node.attribute("StrokeDashArray").is_some();
    let start_cap = node.attribute(if dashed {
        "StrokeDashCap"
    } else {
        "StrokeStartLineCap"
    }) == Some("Triangle");
    let end_cap = node.attribute(if dashed {
        "StrokeDashCap"
    } else {
        "StrokeEndLineCap"
    }) == Some("Triangle");
    let miter = node.attribute("StrokeLineJoin").unwrap_or("Miter") == "Miter";
    if width <= 0.0
        || (!start_cap && !end_cap && !miter)
        || (node.attribute("Stroke").is_none() && node.property("Path.Stroke").is_none())
    {
        return Ok(None);
    }
    let Geometry::Path { commands, .. } = geometry else {
        return Ok(None);
    };
    let mut paths: Vec<Vec<(f32, f32)>> = Vec::new();
    let mut closed = Vec::new();
    let mut generated_points = 0;
    for command in commands {
        match *command {
            PathCommand::MoveTo { x, y } => {
                generated_points += 1;
                paths.push(vec![(x, y)]);
                closed.push(false);
            }
            PathCommand::ClosePath => {
                if let Some(points) = paths.last_mut() {
                    generated_points += 1;
                    points.push(points[0]);
                    *closed.last_mut().unwrap() = true;
                }
            }
            _ => {
                let Some(points) = paths.last_mut() else {
                    continue;
                };
                let previous_len = points.len();
                let from = *points.last().unwrap();
                match *command {
                    PathCommand::LineTo { x, y } => points.push((x, y)),
                    PathCommand::BezierCurveTo {
                        cp1x,
                        cp1y,
                        cp2x,
                        cp2y,
                        x,
                        y,
                    } => {
                        // Bound flattening by the control polygon; only cap placement uses this approximation.
                        let length = (cp1x - from.0).hypot(cp1y - from.1)
                            + (cp2x - cp1x).hypot(cp2y - cp1y)
                            + (x - cp2x).hypot(y - cp2y);
                        let steps = (length / 0.1).ceil().clamp(1.0, 4096.0) as usize;
                        for i in 1..=steps {
                            let t = i as f32 / steps as f32;
                            points.push((
                                cubic(from.0, cp1x, cp2x, x, t),
                                cubic(from.1, cp1y, cp2y, y, t),
                            ));
                        }
                    }
                    PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => {
                        let length = (cpx - from.0).hypot(cpy - from.1) + (x - cpx).hypot(y - cpy);
                        let steps = (length / 0.1).ceil().clamp(1.0, 4096.0) as usize;
                        for i in 1..=steps {
                            let t = i as f32 / steps as f32;
                            points
                                .push((quadratic(from.0, cpx, x, t), quadratic(from.1, cpy, y, t)));
                        }
                    }
                    _ => {}
                }
                generated_points += points.len() - previous_len;
            }
        }
        if generated_points > limit {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ObjectLimit,
                Phase::Parse,
                None,
                "XPS cap geometry exceeds the object limit",
            )
            .in_part(&node.source_part));
        }
    }
    let mut caps = Vec::new();
    let cap = |caps: &mut Vec<PathCommand>,
               point: (f32, f32),
               inside: (f32, f32)|
     -> Result<(), Diagnostic> {
        if caps.len() > limit.saturating_sub(4) {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ObjectLimit,
                Phase::Parse,
                None,
                "XPS cap count exceeds the object limit",
            )
            .in_part(&node.source_part));
        }
        let length = (point.0 - inside.0).hypot(point.1 - inside.1);
        if length == 0.0 {
            return Ok(());
        }
        let dx = (point.0 - inside.0) / length * width / 2.0;
        let dy = (point.1 - inside.1) / length * width / 2.0;
        caps.extend([
            PathCommand::MoveTo {
                x: point.0 - dy,
                y: point.1 + dx,
            },
            PathCommand::LineTo {
                x: point.0 + dx,
                y: point.1 + dy,
            },
            PathCommand::LineTo {
                x: point.0 + dy,
                y: point.1 - dx,
            },
            PathCommand::ClosePath,
        ]);
        Ok(())
    };
    let style = stroke_style(node)?.unwrap_or_default();
    let mut dash = style.dash;
    if dash.len() % 2 == 1 {
        dash.extend(dash.clone());
    }
    for (points, closed) in paths.iter().zip(&closed) {
        if !start_cap && !end_cap {
            continue;
        }
        if points.len() < 2 {
            continue;
        }
        if dashed && dash.iter().any(|value| *value > 0.0) {
            let period: f32 = dash.iter().sum();
            let length: f32 = points
                .windows(2)
                .map(|p| (p[1].0 - p[0].0).hypot(p[1].1 - p[0].1))
                .sum();
            if length / period * dash.len() as f32 > limit.saturating_sub(caps.len()) as f32 / 4.0 {
                return Err(Diagnostic::fatal(
                    DiagnosticCode::ObjectLimit,
                    Phase::Parse,
                    None,
                    "XPS dash cap count exceeds the object limit",
                )
                .in_part(&node.source_part));
            }
            let mut segments = Vec::new();
            crate::model::append_dashed_polyline(&mut segments, points, &dash, style.dash_offset);
            let mut run = Vec::new();
            for command in segments
                .into_iter()
                .chain([PathCommand::MoveTo { x: 0.0, y: 0.0 }])
            {
                match command {
                    PathCommand::MoveTo { x, y } => {
                        if run.len() > 1 {
                            cap(&mut caps, run[0], run[1])?;
                            cap(&mut caps, run[run.len() - 1], run[run.len() - 2])?;
                        }
                        run.clear();
                        run.push((x, y));
                    }
                    PathCommand::LineTo { x, y } => run.push((x, y)),
                    _ => {}
                }
            }
        } else if !*closed {
            if start_cap {
                cap(&mut caps, points[0], points[1])?;
            }
            if end_cap {
                cap(
                    &mut caps,
                    points[points.len() - 1],
                    points[points.len() - 2],
                )?;
            }
        }
    }
    // XPS clips a miter at its limit; Canvas changes it to a bevel. Add only
    // the area beyond that bevel, including the finite 180-degree cusp.
    if miter {
        let radius = width / 2.0;
        let reach = style.miter_limit.max(1.0) * radius;
        for (points, closed) in paths.iter().zip(&closed) {
            let mut points = points.clone();
            if *closed && points.len() > 2 {
                points.push(points[1]);
            }
            let mut distance = 0.0;
            for window in points.windows(3) {
                let [a, b, c] = [window[0], window[1], window[2]];
                let l1 = (b.0 - a.0).hypot(b.1 - a.1);
                let l2 = (c.0 - b.0).hypot(c.1 - b.1);
                distance += l1;
                if l1 == 0.0 || l2 == 0.0 {
                    continue;
                }
                if !dash.is_empty() {
                    let mut phase = (distance + style.dash_offset).rem_euclid(dash.iter().sum());
                    let mut drawn = false;
                    for (index, length) in dash.iter().enumerate() {
                        if phase < *length {
                            drawn = index % 2 == 0 && phase > 0.0;
                            break;
                        }
                        phase -= length;
                    }
                    if !drawn {
                        continue;
                    }
                }
                let u = ((b.0 - a.0) / l1, (b.1 - a.1) / l1);
                let v = ((c.0 - b.0) / l2, (c.1 - b.1) / l2);
                let dot = (u.0 * v.0 + u.1 * v.1).clamp(-1.0, 1.0);
                let sign = (u.0 * v.1 - u.1 * v.0).signum();
                let half_cos = ((1.0 + dot) / 2.0).sqrt();
                if half_cos > 0.0 && radius / half_cos <= reach {
                    continue;
                }
                let polygon = if half_cos < 1e-5 {
                    let normal = (u.1 * radius, -u.0 * radius);
                    [
                        (b.0 + normal.0, b.1 + normal.1),
                        (b.0 + normal.0 + u.0 * reach, b.1 + normal.1 + u.1 * reach),
                        (b.0 - normal.0 + u.0 * reach, b.1 - normal.1 + u.1 * reach),
                        (b.0 - normal.0, b.1 - normal.1),
                    ]
                } else {
                    let n1 = (sign * u.1, -sign * u.0);
                    let n2 = (sign * v.1, -sign * v.0);
                    let length = (n1.0 + n2.0).hypot(n1.1 + n2.1);
                    let direction = ((n1.0 + n2.0) / length, (n1.1 + n2.1) / length);
                    let p = (b.0 + n1.0 * radius, b.1 + n1.1 * radius);
                    let q = (b.0 + n2.0 * radius, b.1 + n2.1 * radius);
                    let t = (reach - radius * half_cos) / (direction.0 * u.0 + direction.1 * u.1);
                    let r = (reach - radius * half_cos) / (direction.0 * v.0 + direction.1 * v.1);
                    [
                        p,
                        (p.0 + u.0 * t, p.1 + u.1 * t),
                        (q.0 + v.0 * r, q.1 + v.1 * r),
                        q,
                    ]
                };
                caps.push(PathCommand::MoveTo {
                    x: polygon[0].0,
                    y: polygon[0].1,
                });
                caps.extend(
                    polygon[1..]
                        .iter()
                        .map(|&(x, y)| PathCommand::LineTo { x, y }),
                );
                caps.push(PathCommand::ClosePath);
                if caps.len() > limit {
                    return Err(Diagnostic::fatal(
                        DiagnosticCode::ObjectLimit,
                        Phase::Parse,
                        None,
                        "XPS stroke geometry exceeds the object limit",
                    )
                    .in_part(&node.source_part));
                }
            }
        }
    }
    if caps.is_empty() {
        return Ok(None);
    }
    Ok(Some(Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands: caps,
    }))
}

fn stroke_style(node: &Node) -> Result<Option<StrokeStyle>, Diagnostic> {
    let width = optional_finite_attribute(node, "StrokeThickness")?
        .unwrap_or(1.0)
        .max(0.0);
    let mut dash = node
        .attribute("StrokeDashArray")
        .map(number_list)
        .unwrap_or_default()
        .into_iter()
        .map(|value| value * width)
        .collect::<Vec<_>>();
    if dash.iter().any(|value| *value < 0.0) || !dash.iter().sum::<f32>().is_finite() {
        return Err(format_error(&node.source_part, "invalid XPS dash array"));
    }
    if width == 0.0 || dash.iter().all(|value| *value == 0.0) {
        dash.clear();
    }
    let line_cap = match node
        .attribute(if dash.is_empty() {
            "StrokeStartLineCap"
        } else {
            "StrokeDashCap"
        })
        .or_else(|| node.attribute("StrokeEndLineCap"))
    {
        Some("Round") => crate::model::LineCap::Round,
        Some("Square") => crate::model::LineCap::Square,
        _ => crate::model::LineCap::Flat,
    };
    let line_join = match node.attribute("StrokeLineJoin") {
        Some("Round") => crate::model::LineJoin::Round,
        Some("Bevel") => crate::model::LineJoin::Bevel,
        _ => crate::model::LineJoin::Miter,
    };
    let miter_limit = optional_finite_attribute(node, "StrokeMiterLimit")?
        .unwrap_or(10.0)
        .max(0.0);
    if dash.is_empty()
        && line_cap == crate::model::LineCap::Flat
        && line_join == crate::model::LineJoin::Miter
        && miter_limit == 10.0
    {
        return Ok(None);
    }
    Ok(Some(StrokeStyle {
        cap: line_cap,
        join: line_join,
        miter_limit,
        dash_offset: optional_finite_attribute(node, "StrokeDashOffset")?.unwrap_or(0.0) * width,
        dash,
        ..StrokeStyle::default()
    }))
}

fn parse_path_geometry_node(node: &Node, for_stroke: bool) -> Option<Geometry> {
    let mut commands = Vec::new();
    for figure in &node.children {
        if figure.local_name() != "PathFigure"
            || (!for_stroke && figure.attribute("IsFilled") == Some("false"))
        {
            continue;
        }
        let mut current = point(figure.attribute("StartPoint")?).ok()?;
        commands.push(PathCommand::MoveTo {
            x: current.0,
            y: current.1,
        });
        for segment in &figure.children {
            let mut next = Vec::new();
            if segment.local_name() == "ArcSegment" {
                let end = point(segment.attribute("Point")?).ok()?;
                let size = point(segment.attribute("Size")?).ok()?;
                arc_to_beziers(
                    current,
                    end,
                    size.0,
                    size.1,
                    segment
                        .attribute("RotationAngle")
                        .and_then(parse_finite)
                        .unwrap_or(0.0),
                    segment.attribute("IsLargeArc") == Some("true"),
                    segment.attribute("SweepDirection") == Some("Clockwise"),
                    &mut next,
                );
                current = end;
            } else {
                let (letter, points) = match segment.local_name() {
                    "PolyLineSegment" => ('L', segment.attribute("Points")?.to_owned()),
                    "PolyBezierSegment" => ('C', segment.attribute("Points")?.to_owned()),
                    "PolyQuadraticBezierSegment" => ('Q', segment.attribute("Points")?.to_owned()),
                    "LineSegment" => ('L', segment.attribute("Point")?.to_owned()),
                    "BezierSegment" => (
                        'C',
                        format!(
                            "{} {} {}",
                            segment.attribute("Point1")?,
                            segment.attribute("Point2")?,
                            segment.attribute("Point3")?
                        ),
                    ),
                    "QuadraticBezierSegment" => (
                        'Q',
                        format!(
                            "{} {}",
                            segment.attribute("Point1")?,
                            segment.attribute("Point2")?
                        ),
                    ),
                    _ => continue,
                };
                let Geometry::Path {
                    commands: parsed, ..
                } = parse_path_geometry(&format!("M{},{} {letter}{points}", current.0, current.1))
                    .ok()?
                else {
                    return None;
                };
                next.extend(parsed.into_iter().skip(1));
                current = match next.last()? {
                    PathCommand::LineTo { x, y }
                    | PathCommand::BezierCurveTo { x, y, .. }
                    | PathCommand::QuadraticCurveTo { x, y, .. } => (*x, *y),
                    _ => return None,
                };
            }
            if for_stroke && segment.attribute("IsStroked") == Some("false") {
                commands.push(PathCommand::MoveTo {
                    x: current.0,
                    y: current.1,
                });
            } else {
                commands.extend(next);
            }
        }
        if figure.attribute("IsClosed") == Some("true") {
            commands.push(PathCommand::ClosePath);
        }
    }
    Some(Geometry::Path {
        fill_rule: if node.attribute("FillRule") == Some("NonZero") {
            FillRule::NonZero
        } else {
            FillRule::EvenOdd
        },
        commands,
    })
}

#[derive(Clone, Copy)]
enum PathToken {
    Command(char),
    Number(f32),
}

fn path_tokens(value: &str) -> Result<Vec<PathToken>, String> {
    let bytes = value.as_bytes();
    let mut tokens = Vec::new();
    let mut cursor = 0_usize;
    while cursor < bytes.len() {
        let byte = bytes[cursor];
        if byte.is_ascii_whitespace() || byte == b',' {
            cursor += 1;
            continue;
        }
        if byte.is_ascii_alphabetic() {
            tokens.push(PathToken::Command(byte as char));
            cursor += 1;
            continue;
        }
        let start = cursor;
        if matches!(bytes[cursor], b'+' | b'-') {
            cursor += 1;
        }
        let mut decimal = false;
        while cursor < bytes.len()
            && (bytes[cursor].is_ascii_digit() || (!decimal && bytes[cursor] == b'.'))
        {
            decimal |= bytes[cursor] == b'.';
            cursor += 1;
        }
        if cursor < bytes.len() && matches!(bytes[cursor], b'e' | b'E') {
            cursor += 1;
            if cursor < bytes.len() && matches!(bytes[cursor], b'+' | b'-') {
                cursor += 1;
            }
            while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
                cursor += 1;
            }
        }
        let number = value[start..cursor]
            .parse::<f32>()
            .ok()
            .filter(|value| value.is_finite())
            .ok_or_else(|| "invalid XPS path number".to_owned())?;
        tokens.push(PathToken::Number(number));
    }
    Ok(tokens)
}

fn parse_path_geometry(value: &str) -> Result<Geometry, String> {
    let tokens = path_tokens(value)?;
    let mut index = 0_usize;
    let mut command = 'M';
    let mut current = (0.0_f32, 0.0_f32);
    let mut start = current;
    let mut last_control = None;
    let mut fill_rule = FillRule::EvenOdd;
    let mut commands = Vec::new();
    while index < tokens.len() {
        if let PathToken::Command(next) = tokens[index] {
            command = next;
            index += 1;
            if matches!(command, 'Z' | 'z') {
                commands.push(PathCommand::ClosePath);
                current = start;
                last_control = None;
                continue;
            }
            if command == 'F' {
                let mode = take_numbers(&tokens, &mut index, 1)?[0];
                fill_rule = if mode == 0.0 {
                    FillRule::EvenOdd
                } else {
                    FillRule::NonZero
                };
                continue;
            }
        }
        let relative = command.is_ascii_lowercase();
        let absolute = |point: (f32, f32), current: (f32, f32)| {
            if relative {
                (point.0 + current.0, point.1 + current.1)
            } else {
                point
            }
        };
        match command.to_ascii_uppercase() {
            'M' | 'L' => {
                let values = take_numbers(&tokens, &mut index, 2)?;
                let point = absolute((values[0], values[1]), current);
                if command.eq_ignore_ascii_case(&'M') {
                    commands.push(PathCommand::MoveTo {
                        x: point.0,
                        y: point.1,
                    });
                    start = point;
                    command = if relative { 'l' } else { 'L' };
                } else {
                    commands.push(PathCommand::LineTo {
                        x: point.0,
                        y: point.1,
                    });
                }
                current = point;
                last_control = None;
            }
            'H' => {
                let x = take_numbers(&tokens, &mut index, 1)?[0]
                    + if relative { current.0 } else { 0.0 };
                current.0 = x;
                commands.push(PathCommand::LineTo {
                    x: current.0,
                    y: current.1,
                });
                last_control = None;
            }
            'V' => {
                let y = take_numbers(&tokens, &mut index, 1)?[0]
                    + if relative { current.1 } else { 0.0 };
                current.1 = y;
                commands.push(PathCommand::LineTo {
                    x: current.0,
                    y: current.1,
                });
                last_control = None;
            }
            'C' => {
                let values = take_numbers(&tokens, &mut index, 6)?;
                let first = absolute((values[0], values[1]), current);
                let second = absolute((values[2], values[3]), current);
                let end = absolute((values[4], values[5]), current);
                commands.push(PathCommand::BezierCurveTo {
                    cp1x: first.0,
                    cp1y: first.1,
                    cp2x: second.0,
                    cp2y: second.1,
                    x: end.0,
                    y: end.1,
                });
                current = end;
                last_control = Some(second);
            }
            'S' => {
                let values = take_numbers(&tokens, &mut index, 4)?;
                let first = last_control.map_or(current, |control: (f32, f32)| {
                    (2.0 * current.0 - control.0, 2.0 * current.1 - control.1)
                });
                let second = absolute((values[0], values[1]), current);
                let end = absolute((values[2], values[3]), current);
                commands.push(PathCommand::BezierCurveTo {
                    cp1x: first.0,
                    cp1y: first.1,
                    cp2x: second.0,
                    cp2y: second.1,
                    x: end.0,
                    y: end.1,
                });
                current = end;
                last_control = Some(second);
            }
            'Q' => {
                let values = take_numbers(&tokens, &mut index, 4)?;
                let control = absolute((values[0], values[1]), current);
                let end = absolute((values[2], values[3]), current);
                commands.push(PathCommand::QuadraticCurveTo {
                    cpx: control.0,
                    cpy: control.1,
                    x: end.0,
                    y: end.1,
                });
                current = end;
                last_control = Some(control);
            }
            'A' => {
                let values = take_numbers(&tokens, &mut index, 7)?;
                let end = absolute((values[5], values[6]), current);
                arc_to_beziers(
                    current,
                    end,
                    values[0].abs(),
                    values[1].abs(),
                    values[2],
                    values[3] != 0.0,
                    values[4] != 0.0,
                    &mut commands,
                );
                current = end;
                last_control = None;
            }
            _ => return Err(format!("unsupported XPS path command {command}")),
        }
    }
    if commands.is_empty() {
        return Err("empty XPS path geometry".to_owned());
    }
    Ok(Geometry::Path {
        fill_rule,
        commands,
    })
}

fn take_numbers(tokens: &[PathToken], index: &mut usize, count: usize) -> Result<Vec<f32>, String> {
    let mut values = Vec::with_capacity(count);
    for _ in 0..count {
        match tokens.get(*index) {
            Some(PathToken::Number(value)) => {
                values.push(*value);
                *index += 1;
            }
            _ => return Err("truncated XPS path command".to_owned()),
        }
    }
    Ok(values)
}

#[allow(clippy::too_many_arguments)]
fn arc_to_beziers(
    start: (f32, f32),
    end: (f32, f32),
    mut rx: f32,
    mut ry: f32,
    rotation: f32,
    large: bool,
    sweep: bool,
    commands: &mut Vec<PathCommand>,
) {
    if rx == 0.0 || ry == 0.0 || start == end {
        commands.push(PathCommand::LineTo { x: end.0, y: end.1 });
        return;
    }
    let phi = rotation.to_radians();
    let (sin_phi, cos_phi) = phi.sin_cos();
    let dx = (start.0 - end.0) / 2.0;
    let dy = (start.1 - end.1) / 2.0;
    let x1 = cos_phi * dx + sin_phi * dy;
    let y1 = -sin_phi * dx + cos_phi * dy;
    let scale = (x1 * x1 / (rx * rx) + y1 * y1 / (ry * ry)).sqrt().max(1.0);
    rx *= scale;
    ry *= scale;
    let numerator = (rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1).max(0.0);
    let denominator = rx * rx * y1 * y1 + ry * ry * x1 * x1;
    let sign = if large == sweep { -1.0 } else { 1.0 };
    let coefficient = if denominator == 0.0 {
        0.0
    } else {
        sign * (numerator / denominator).sqrt()
    };
    let cx1 = coefficient * rx * y1 / ry;
    let cy1 = coefficient * -ry * x1 / rx;
    let center = (
        cos_phi * cx1 - sin_phi * cy1 + (start.0 + end.0) / 2.0,
        sin_phi * cx1 + cos_phi * cy1 + (start.1 + end.1) / 2.0,
    );
    let angle = |u: (f32, f32), v: (f32, f32)| (u.0 * v.1 - u.1 * v.0).atan2(u.0 * v.0 + u.1 * v.1);
    let unit_start = ((x1 - cx1) / rx, (y1 - cy1) / ry);
    let unit_end = ((-x1 - cx1) / rx, (-y1 - cy1) / ry);
    let theta = angle((1.0, 0.0), unit_start);
    let mut delta = angle(unit_start, unit_end);
    if sweep && delta < 0.0 {
        delta += std::f32::consts::TAU;
    }
    if !sweep && delta > 0.0 {
        delta -= std::f32::consts::TAU;
    }
    let segments = (delta.abs() / std::f32::consts::FRAC_PI_2).ceil() as usize;
    let segment_delta = delta / segments.max(1) as f32;
    let map = |point: (f32, f32)| {
        (
            center.0 + cos_phi * rx * point.0 - sin_phi * ry * point.1,
            center.1 + sin_phi * rx * point.0 + cos_phi * ry * point.1,
        )
    };
    for segment in 0..segments.max(1) {
        let a0 = theta + segment as f32 * segment_delta;
        let a1 = a0 + segment_delta;
        let alpha = 4.0 / 3.0 * ((a1 - a0) / 4.0).tan();
        let (s0, c0) = a0.sin_cos();
        let (s1, c1) = a1.sin_cos();
        let control1 = map((c0 - alpha * s0, s0 + alpha * c0));
        let control2 = map((c1 + alpha * s1, s1 - alpha * c1));
        let endpoint = map((c1, s1));
        commands.push(PathCommand::BezierCurveTo {
            cp1x: control1.0,
            cp1y: control1.1,
            cp2x: control2.0,
            cp2y: control2.1,
            x: endpoint.0,
            y: endpoint.1,
        });
    }
}

fn format_error(part: impl Into<String>, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message).in_part(part)
}

#[cfg(test)]
mod tests {
    use super::{
        GlyphStyle, font_obfuscation_key, geometry_bounds, glyph_geometry, icc_channel_count,
        parse_color, parse_glyph_indices, parse_path_geometry, xps_xml_bytes,
    };

    use crate::model::{FillRule, Geometry, PathCommand};

    #[test]
    #[ignore = "requires the downloaded SampleXpsDocuments_1_0 corpus in .cache/xps-tests"]
    fn real_tiff_corpus_opens_and_serializes_every_page() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join(".cache/xps-tests/SampleXpsDocuments_1_0/Handcrafted");
        let mut count = 0;
        for entry in std::fs::read_dir(root.join("RenderedDocs")).unwrap() {
            let path = entry.unwrap().path();
            if path.extension().is_none_or(|extension| extension != "tif") {
                continue;
            }
            let name = path.file_stem().unwrap().to_str().unwrap();
            let file = if name == "XPS_09_Examples" {
                "XPS_Examples"
            } else {
                name
            };
            let bytes = std::fs::read(root.join(format!("{file}.xps"))).unwrap();
            let prepared = super::prepare(&bytes, crate::limits::Limits::default())
                .unwrap()
                .unwrap();
            for index in 0..prepared.page_parts.len() {
                let (document, _) = prepared
                    .materialize(Some(index), &std::collections::HashSet::new())
                    .unwrap_or_else(|error| panic!("{file} page {}: {error:?}", index + 1));
                crate::protocol::encode(&document)
                    .unwrap_or_else(|error| panic!("{file} page {}: {error:?}", index + 1));
                // The immutable package must not rescan every other page on each lazy load.
                for (dimensions, unit) in prepared.page_dimensions.iter().zip(&document.units) {
                    assert_eq!(dimensions.get(), Some(&(unit.width, unit.height)));
                }
                count += 1;
            }
        }
        assert_eq!(count, 151);
    }

    #[test]
    #[ignore = "requires the downloaded SampleXpsDocuments_1_0 corpus in .cache/xps-tests"]
    fn real_unknown_visual_is_omitted_without_losing_siblings() {
        let bytes = std::fs::read(concat!(env!("CARGO_MANIFEST_DIR"),
            "/.cache/xps-tests/SampleXpsDocuments_1_0/ConformanceViolations/MarkupCompatibility/M1.1a.xps")).unwrap();
        let document = super::detect_and_parse(&bytes, crate::limits::Limits::default())
            .expect("an unknown visual must not fail the document")
            .unwrap();
        assert!(!document.fatal);
        assert_eq!(document.units.len(), 1);
        assert_eq!(document.objects.len(), 3);
        assert!(document.objects.iter().any(|object| object.text.as_deref()
            == Some("If you see this text, file a bug against this XPS Consumer.")));
        let warning = document
            .diagnostics
            .iter()
            .find(|d| d.message.contains("v2:Circle"))
            .unwrap();
        assert_eq!(warning.severity, crate::diagnostic::Severity::Warning);
        assert_eq!(warning.fidelity, crate::diagnostic::Fidelity::Omitted);
        assert_eq!(
            warning.location.part.as_deref(),
            Some("Documents/1/Pages/1.fpage")
        );
        crate::protocol::encode(&document).unwrap();
    }

    #[test]
    #[ignore = "requires the downloaded SampleXpsDocuments_1_0 corpus in .cache/xps-tests"]
    fn real_directory_document_reference_recovers_content() {
        let bytes = std::fs::read(concat!(env!("CARGO_MANIFEST_DIR"),
            "/.cache/xps-tests/SampleXpsDocuments_1_0/ConformanceViolations/OpenPackagingConventions/M1.1b.xps")).unwrap();
        let document = super::detect_and_parse(&bytes, crate::limits::Limits::default())
            .expect("a directory reference with one document should recover")
            .unwrap();
        assert_eq!(document.units.len(), 1);
        let text: String = document
            .objects
            .iter()
            .filter_map(|o| o.text.as_deref())
            .collect();
        assert_eq!(text.trim(), "This is a test");
        let warning = document
            .diagnostics
            .iter()
            .find(|d| d.message.contains("FixedDoc.fdoc"))
            .unwrap();
        assert_eq!(warning.severity, crate::diagnostic::Severity::Warning);
        assert_eq!(warning.location.part.as_deref(), Some("FixedDocSeq.fdseq"));
        crate::protocol::encode(&document).unwrap();
    }

    #[test]
    #[ignore = "requires the downloaded SampleXpsDocuments_1_0 corpus in .cache/xps-tests"]
    fn real_opc_metadata_and_uri_recovery() {
        for name in [
            "M1.5a", "M1.6b", "M1.8a", "M1.10a", "M1.26a", "M1.26c", "M1.27a", "M1.28a", "M2.5a",
            "M2.5b", "M2.6a", "M2.6b", "M2.7a", "M2.7b", "M4.2a",
        ] {
            let path = format!(
                "{}/.cache/xps-tests/SampleXpsDocuments_1_0/ConformanceViolations/OpenPackagingConventions/{name}.xps",
                env!("CARGO_MANIFEST_DIR")
            );
            let bytes = std::fs::read(path).unwrap();
            let document = super::detect_and_parse(&bytes, crate::limits::Limits::default())
                .unwrap_or_else(|error| panic!("{name}: {error:?}"))
                .unwrap();
            assert_eq!(document.units.len(), 1, "{name}");
            let text: String = document
                .objects
                .iter()
                .filter_map(|o| o.text.as_deref())
                .collect();
            assert_eq!(text.trim(), "This is a test", "{name}");
            if !["M1.8a", "M4.2a"].contains(&name) {
                assert!(
                    document
                        .diagnostics
                        .iter()
                        .any(|d| d.severity == crate::diagnostic::Severity::Warning),
                    "{name}"
                );
            }
        }
        assert!(super::resolve_part("FixedDocSeq.fdseq", "/%2e%2e/secret").is_err());
        assert_eq!(
            super::resolve_part("FixedDocSeq.fdseq", "/a%2fb").unwrap(),
            "a%2fb"
        );
    }

    #[test]
    #[ignore = "requires the downloaded SampleXpsDocuments_1_0 corpus in .cache/xps-tests"]
    fn real_xps_bad_objects_preserve_content_and_diagnostics() {
        let root = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/.cache/xps-tests/SampleXpsDocuments_1_0/ConformanceViolations/XPS/"
        );
        for name in [
            "M2.4b", "M2.4a", "M2.3a", "M2.18a", "M2.3b", "M2.13a", "M2.14a", "M3.2a", "M2.6a",
            "M2.6b", "M2.25a", "M5.2a", "M5.2b", "M5.4a", "M5.4b", "M5.15a", "M6.3a",
        ] {
            let bytes = std::fs::read(format!("{root}{name}.xps")).unwrap();
            let document = super::detect_and_parse(&bytes, crate::limits::Limits::default())
                .unwrap_or_else(|e| panic!("{name}: {e:?}"))
                .unwrap();
            assert_eq!(document.units.len(), 1, "{name}");
            assert!(!document.objects.is_empty(), "{name}");
            assert!(
                document
                    .diagnostics
                    .iter()
                    .any(|d| d.severity == crate::diagnostic::Severity::Warning),
                "{name}"
            );
            let text: String = document
                .objects
                .iter()
                .filter_map(|o| o.text.as_deref())
                .collect();
            if name == "M2.6a" {
                assert_eq!(text.trim(), "This is a test");
            }
            if name == "M2.6b" {
                assert_eq!(text.trim(), "his is a test");
            }
            if name == "M5.4a" {
                assert_eq!(text.trim(), "This is a red square");
            }
            if name == "M5.4b" {
                assert_eq!(text.trim(), "A Short String");
            }
            crate::protocol::encode(&document).unwrap();
        }
        for name in ["M2.71a", "M2.71b", "M2.71c"] {
            let bytes = std::fs::read(format!("{root}{name}.xps")).unwrap();
            assert_eq!(
                super::detect_and_parse(&bytes, crate::limits::Limits::default())
                    .unwrap_err()
                    .code,
                crate::diagnostic::DiagnosticCode::XmlDtdForbidden
            );
        }
    }

    #[test]
    fn real_cover_keeps_canvas_transform_around_child_visuals() {
        use crate::model::Visual;
        let document = super::detect_and_parse(
            include_bytes!("../../tests/fixtures/ecma-388-cover.xps"),
            crate::limits::Limits::default(),
        )
        .unwrap()
        .unwrap();
        let canvas = document
            .objects
            .iter()
            .find(|object| {
                matches!(object.visual, Visual::Layer { transform, .. }
                if (transform.a - 4.0 / 3.0).abs() < 0.0001)
            })
            .unwrap();
        let Visual::Layer { visual, .. } = &canvas.visual else {
            unreachable!()
        };
        assert!(
            matches!(visual.as_ref(), Visual::Group { children } if !children.is_empty()),
            "the cover scale must enclose its content in the render tree"
        );
        assert!(
            document
                .objects
                .iter()
                .filter(|object| object.text.is_some())
                .all(|object| matches!(object.visual, Visual::None)),
            "source text must remain searchable without drawing twice"
        );
        assert!(document.objects.iter().any(|object| {
            object
                .text
                .as_deref()
                .is_some_and(|text| text.contains("Open XML"))
        }));
    }

    #[test]
    fn parses_fixed_page_path_and_utf16_xml() {
        let Geometry::Path {
            fill_rule,
            commands,
        } = parse_path_geometry("F1 M 1,2 L 3,4 A 2,2 0 0 1 5,6 Z").unwrap()
        else {
            panic!("expected path");
        };
        assert_eq!(fill_rule, FillRule::NonZero);
        assert!(matches!(
            commands.first(),
            Some(PathCommand::MoveTo { x: 1.0, y: 2.0 })
        ));
        assert!(matches!(commands.last(), Some(PathCommand::ClosePath)));

        let text = "<FixedPage Width=\"1\" Height=\"1\"/>";
        let mut utf16 = vec![0xff, 0xfe];
        utf16.extend(text.encode_utf16().flat_map(u16::to_le_bytes));
        assert_eq!(xps_xml_bytes(&utf16).unwrap(), text.as_bytes());
    }

    #[test]
    fn reverses_xps_font_guid_bytes() {
        assert_eq!(
            font_obfuscation_key("Resources/00112233-4455-6677-8899-AABBCCDDEEFF.odttf").unwrap(),
            [
                0xff, 0xee, 0xdd, 0xcc, 0xbb, 0xaa, 0x99, 0x88, 0x77, 0x66, 0x55, 0x44, 0x33, 0x22,
                0x11, 0x00
            ],
        );
    }

    #[test]
    fn parses_every_xps_indices_mapping_entry() {
        let entries = parse_glyph_indices("(2:3)94;76;88;(2)162;,40,,10;").unwrap();
        assert_eq!(entries.len(), 6);
        assert_eq!((entries[0].code_units, entries[0].glyphs), (2, 3));
        assert_eq!(entries[3].id, Some(162));
        assert_eq!((entries[3].code_units, entries[3].glyphs), (2, 1));
        assert_eq!(entries[4].advance, Some(40.0));
        assert_eq!(entries[4].v_offset, 10.0);
        assert_eq!(entries[5].id, None);
        assert!(parse_glyph_indices("1,-1").is_err());
        assert!(parse_glyph_indices("70000").is_err());
    }

    #[test]
    fn lays_out_xps_glyph_outlines_without_browser_shaping() {
        let font = include_bytes!("../../tests/fixtures/LiberationSans-Regular.ttf");
        let plain = GlyphStyle {
            bold: false,
            italic: false,
        };
        let (_, ltr) = glyph_geometry(
            font,
            0,
            "AFQ",
            Some(";,100,30,10;"),
            100.0,
            100.0,
            48.0,
            false,
            false,
            plain,
            10_000,
        )
        .unwrap();
        let (_, rtl) = glyph_geometry(
            font,
            0,
            "AFQ",
            Some(";,100,30,10;"),
            100.0,
            100.0,
            48.0,
            true,
            false,
            plain,
            10_000,
        )
        .unwrap();
        assert!(rtl.x < ltr.x);
        assert!(rtl.x + rtl.width <= 100.0);

        let (explicit, _) = glyph_geometry(
            font,
            0,
            "",
            Some("36,100"),
            10.0,
            20.0,
            20.0,
            false,
            false,
            GlyphStyle {
                bold: true,
                italic: true,
            },
            10_000,
        )
        .unwrap();
        let crate::model::Geometry::Path { commands, .. } = explicit else {
            panic!("expected glyph outline path");
        };
        assert!(!commands.is_empty());

        let (_, sideways) = glyph_geometry(
            font, 0, "A", None, 100.0, 100.0, 48.0, false, true, plain, 10_000,
        )
        .unwrap();
        assert!(sideways.x >= 99.0);
        assert!(sideways.x + sideways.width > 100.0);
    }

    #[test]
    fn converts_scrgb_colors_without_dropping_alpha() {
        assert_eq!(parse_color("sc#0.5,1,0,0"), Some(0xff00_0080));
        assert_eq!(parse_color("sc#0,1,0"), Some(0x00ff_00ff));
    }

    #[test]
    fn validates_xps_icc_channel_counts_without_a_color_engine() {
        let mut profile = vec![0_u8; 128];
        profile[..4].copy_from_slice(&128_u32.to_be_bytes());
        profile[16..20].copy_from_slice(b"CMYK");
        profile[36..40].copy_from_slice(b"acsp");
        assert_eq!(icc_channel_count(&profile), Some(4));
        profile[16..20].copy_from_slice(b"5CLR");
        assert_eq!(icc_channel_count(&profile), Some(5));
        profile[36] = 0;
        assert_eq!(icc_channel_count(&profile), None);
    }

    #[test]
    fn computes_exact_xps_curve_bounds() {
        let geometry = parse_path_geometry("M 0,0 C 0,10 10,10 10,0").unwrap();
        let bounds = geometry_bounds(&geometry).unwrap();
        assert_eq!((bounds.x, bounds.y, bounds.width), (0.0, 0.0, 10.0));
        assert!((bounds.height - 7.5).abs() < 0.0001);
    }
}
