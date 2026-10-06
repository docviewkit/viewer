//! Versioned binary snapshot consumed by the TypeScript worker runtime.

use crate::diagnostic::{
    Diagnostic, DiagnosticCode, Fidelity as DiagnosticFidelity, Phase, Severity,
};
use crate::model::{
    Document, Geometry, Paint, PathCommand, PptxAction, SheetAxis, SourceLocator,
    SpeakerNoteParagraph, TextRun, Visual, XpsColor, XpsGradientStop,
};

const MAGIC: u32 = 0x3144_564f;
const VERSION: u16 = 76;
const NONE: u32 = u32::MAX;
const MAX_SNAPSHOT_BYTES: usize = 256 * 1024 * 1024;
const MAX_VISUAL_DEPTH: usize = 64;
const MAX_PATH_COMMANDS: usize = 65_536;
const MAX_GRADIENT_STOPS: usize = 4_096;
const MAX_VISUAL_BRUSH_CHILDREN: usize = 65_536;
pub(crate) const MAX_RICH_TEXT_RUNS: usize = 100_000;
const MAX_TAB_STOPS: usize = 256;
const MAX_TEXT_PARAGRAPHS: usize = 100_000;

#[derive(Debug)]
pub enum ProtocolError {
    SnapshotTooLarge,
    FieldTooLong,
    TooManyFields,
    AllocationFailed,
    InvalidValue,
}

impl core::fmt::Display for ProtocolError {
    fn fmt(&self, formatter: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        write!(formatter, "cannot encode core snapshot: {self:?}")
    }
}

impl std::error::Error for ProtocolError {}

struct Writer {
    bytes: Vec<u8>,
    expected_length: usize,
    overflowed: bool,
}

impl Writer {
    fn new(expected_length: usize) -> Result<Self, ProtocolError> {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(expected_length)
            .map_err(|_| ProtocolError::AllocationFailed)?;
        Ok(Self {
            bytes,
            expected_length,
            overflowed: false,
        })
    }

    fn append(&mut self, value: &[u8]) {
        let Some(end) = self.bytes.len().checked_add(value.len()) else {
            self.overflowed = true;
            return;
        };
        if end > self.expected_length {
            self.overflowed = true;
            return;
        }
        self.bytes.extend_from_slice(value);
    }

    fn u8(&mut self, value: u8) {
        self.append(&[value]);
    }

    fn u16(&mut self, value: u16) {
        self.append(&value.to_le_bytes());
    }

    fn u32(&mut self, value: u32) {
        self.append(&value.to_le_bytes());
    }

    fn i32(&mut self, value: i32) {
        self.append(&value.to_le_bytes());
    }

    fn f32(&mut self, value: f32) {
        self.append(&value.to_le_bytes());
    }

    fn string(&mut self, value: &str) -> Result<(), ProtocolError> {
        self.u32(
            value
                .len()
                .try_into()
                .map_err(|_| ProtocolError::FieldTooLong)?,
        );
        self.append(value.as_bytes());
        Ok(())
    }

    fn optional_string(&mut self, value: Option<&str>) -> Result<(), ProtocolError> {
        if let Some(value) = value {
            self.string(value)
        } else {
            self.u32(NONE);
            Ok(())
        }
    }

    fn byte_array(&mut self, value: &[u8]) -> Result<(), ProtocolError> {
        self.u32(
            value
                .len()
                .try_into()
                .map_err(|_| ProtocolError::FieldTooLong)?,
        );
        self.append(value);
        Ok(())
    }

    fn padding(&mut self, count: usize) {
        const ZEROES: [u8; 8] = [0; 8];
        let mut remaining = count;
        while remaining != 0 {
            let chunk = remaining.min(ZEROES.len());
            self.append(&ZEROES[..chunk]);
            remaining -= chunk;
        }
    }
}

enum Field<'a> {
    String(&'a str, &'a str),
    Unsigned(&'a str, u32),
}

fn write_fields(writer: &mut Writer, fields: &[Field<'_>]) -> Result<(), ProtocolError> {
    writer.u8(fields
        .len()
        .try_into()
        .map_err(|_| ProtocolError::TooManyFields)?);
    for field in fields {
        match field {
            Field::String(key, value) => {
                writer.string(key)?;
                writer.u8(0);
                writer.string(value)?;
            }
            Field::Unsigned(key, value) => {
                writer.string(key)?;
                writer.u8(1);
                writer.u32(*value);
            }
        }
    }
    Ok(())
}

fn code_name(code: DiagnosticCode) -> &'static str {
    match code {
        DiagnosticCode::None => "NONE",
        DiagnosticCode::InvalidLimits => "INVALID_LIMITS",
        DiagnosticCode::InputTooLarge => "INPUT_SIZE_LIMIT",
        DiagnosticCode::AllocationFailed => "ALLOCATION_FAILED",
        DiagnosticCode::UnsupportedFormat => "UNSUPPORTED_FORMAT",
        DiagnosticCode::ZipInvalid => "PACKAGE_ZIP_INVALID",
        DiagnosticCode::ZipMultiDiskForbidden => "PACKAGE_MULTI_DISK_FORBIDDEN",
        DiagnosticCode::Zip64Forbidden => "PACKAGE_ZIP64_UNSUPPORTED",
        DiagnosticCode::ZipEncrypted => "PACKAGE_ENCRYPTED",
        DiagnosticCode::ZipUnsupportedCompression => "PACKAGE_COMPRESSION_UNSUPPORTED",
        DiagnosticCode::ZipEntryLimit => "PACKAGE_ENTRY_LIMIT",
        DiagnosticCode::ZipEntryTooLarge => "PACKAGE_ENTRY_SIZE_LIMIT",
        DiagnosticCode::ZipTotalSizeLimit => "PACKAGE_INFLATED_SIZE_LIMIT",
        DiagnosticCode::ZipCompressionRatioLimit => "PACKAGE_COMPRESSION_RATIO_LIMIT",
        DiagnosticCode::ZipPathTraversal => "PACKAGE_PATH_TRAVERSAL",
        DiagnosticCode::ZipDuplicateEntry => "DUPLICATE_PACKAGE_ENTRY",
        DiagnosticCode::ZipCrcMismatch => "PACKAGE_CRC_MISMATCH",
        DiagnosticCode::ZipDeflateInvalid => "PACKAGE_DEFLATE_INVALID",
        DiagnosticCode::XmlInvalid => "XML_INVALID",
        DiagnosticCode::XmlEncodingUnsupported => "XML_ENCODING_UNSUPPORTED",
        DiagnosticCode::XmlDtdForbidden => "XML_DTD_FORBIDDEN",
        DiagnosticCode::XmlEntityForbidden => "XML_ENTITY_FORBIDDEN",
        DiagnosticCode::XmlDepthLimit => "XML_DEPTH_LIMIT",
        DiagnosticCode::XmlNodeLimit => "XML_NODE_LIMIT",
        DiagnosticCode::XmlAttributeLimit => "XML_ATTRIBUTE_LIMIT",
        DiagnosticCode::XmlSizeLimit => "XML_SIZE_LIMIT",
        DiagnosticCode::ActiveContentBlocked => "ACTIVE_CONTENT_BLOCKED",
        DiagnosticCode::ExternalResourceBlocked => "EXTERNAL_RESOURCE_BLOCKED",
        DiagnosticCode::UnsupportedFeature => "UNSUPPORTED_FEATURE",
        DiagnosticCode::FormatInvalid => "FORMAT_INVALID",
        DiagnosticCode::RelationshipLimit => "RELATIONSHIP_LIMIT",
        DiagnosticCode::ObjectLimit => "OBJECT_LIMIT",
        DiagnosticCode::ImageDimensionLimit => "IMAGE_DIMENSION_LIMIT",
        DiagnosticCode::LayoutBudgetExceeded => "LAYOUT_BUDGET_EXCEEDED",
        DiagnosticCode::EmbeddedFontInvalid => "EMBEDDED_FONT_INVALID",
        DiagnosticCode::FontBytesLimit => "FONT_BYTES_LIMIT",
        DiagnosticCode::ApproximateLayout => "APPROXIMATE_LAYOUT",
        DiagnosticCode::PdfPasswordRequired => "PDF_PASSWORD_REQUIRED",
        DiagnosticCode::PdfPasswordIncorrect => "PDF_PASSWORD_INCORRECT",
    }
}

fn write_diagnostic(writer: &mut Writer, diagnostic: &Diagnostic) -> Result<(), ProtocolError> {
    writer.string(code_name(diagnostic.code))?;
    writer.u8(match diagnostic.severity {
        Severity::Info => 0,
        Severity::Warning => 1,
        Severity::Error => 2,
        Severity::Fatal => 3,
    });
    writer.u8(match diagnostic.fidelity {
        DiagnosticFidelity::Exact => 0,
        DiagnosticFidelity::Approximate => 1,
        DiagnosticFidelity::Unsupported => 2,
        DiagnosticFidelity::Omitted | DiagnosticFidelity::Blocked => 3,
    });
    writer.u8(match diagnostic.phase {
        Phase::Identify => 0,
        Phase::Input => 0,
        Phase::Container => 1,
        Phase::Xml => 2,
        Phase::Parse => 2,
        Phase::Layout => 3,
        Phase::Render => 4,
        Phase::Security => 5,
    });
    writer.padding(1);
    writer.string(&diagnostic.message)?;
    writer.optional_string(diagnostic.location.part.as_deref())?;
    writer.optional_string(
        diagnostic
            .details
            .iter()
            .find_map(|(key, value)| (key == "objectId").then_some(value.as_str())),
    )?;
    let offset = diagnostic
        .location
        .byte_offset
        .and_then(|value| u32::try_from(value).ok());
    let mut fields = Vec::with_capacity(diagnostic.details.len() + usize::from(offset.is_some()));
    for (key, value) in &diagnostic.details {
        fields.push(Field::String(key, value));
    }
    if let Some(offset) = offset {
        fields.push(Field::Unsigned("byteOffset", offset));
    }
    write_fields(writer, &fields)
}

fn optional_u32_field<'a>(fields: &mut Vec<Field<'a>>, key: &'a str, value: Option<u32>) {
    if let Some(value) = value {
        fields.push(Field::Unsigned(key, value));
    }
}

fn push_action_fields<'a>(
    fields: &mut Vec<Field<'a>>,
    action: &'a PptxAction,
    kind_key: &'a str,
    action_key: &'a str,
    target_key: &'a str,
    tooltip_key: &'a str,
) {
    fields.push(Field::String(kind_key, &action.kind));
    if let Some(value) = action.action.as_deref() {
        fields.push(Field::String(action_key, value));
    }
    if let Some(value) = action.target.as_deref() {
        fields.push(Field::String(target_key, value));
    }
    if let Some(value) = action.tooltip.as_deref() {
        fields.push(Field::String(tooltip_key, value));
    }
}

fn add_action_fields(
    total: &mut usize,
    action: &PptxAction,
    kind_key: &str,
    action_key: &str,
    target_key: &str,
    tooltip_key: &str,
) -> Result<(), ProtocolError> {
    add_string_field(total, kind_key, &action.kind)?;
    if let Some(value) = action.action.as_deref() {
        add_string_field(total, action_key, value)?;
    }
    if let Some(value) = action.target.as_deref() {
        add_string_field(total, target_key, value)?;
    }
    if let Some(value) = action.tooltip.as_deref() {
        add_string_field(total, tooltip_key, value)?;
    }
    Ok(())
}

fn write_source(writer: &mut Writer, locator: &SourceLocator) -> Result<(), ProtocolError> {
    match locator {
        SourceLocator::PptxShape {
            shape_id,
            row,
            column,
            text_range,
            metadata,
        } => {
            writer.string("shape")?;
            let mut fields = vec![Field::Unsigned("shapeId", *shape_id)];
            optional_u32_field(&mut fields, "row", *row);
            optional_u32_field(&mut fields, "column", *column);
            if let Some((start, end)) = text_range {
                fields.push(Field::Unsigned("rangeStart", *start));
                fields.push(Field::Unsigned("rangeEnd", *end));
            }
            if let Some(value) = metadata.name.as_deref() {
                fields.push(Field::String("name", value));
            }
            if let Some(value) = metadata.title.as_deref() {
                fields.push(Field::String("title", value));
            }
            if let Some(value) = metadata.description.as_deref() {
                fields.push(Field::String("description", value));
            }
            if metadata.hidden {
                fields.push(Field::Unsigned("hidden", 1));
            }
            if let Some(action) = metadata.click_action.as_ref() {
                push_action_fields(
                    &mut fields,
                    action,
                    "clickKind",
                    "clickAction",
                    "clickTarget",
                    "clickTooltip",
                );
            }
            if let Some(action) = metadata.hover_action.as_ref() {
                push_action_fields(
                    &mut fields,
                    action,
                    "hoverKind",
                    "hoverAction",
                    "hoverTarget",
                    "hoverTooltip",
                );
            }
            write_fields(writer, &fields)
        }
        SourceLocator::OdpElement {
            element_id,
            path,
            row,
            column,
        } => {
            writer.string("element")?;
            let mut fields = vec![Field::String("path", path)];
            if let Some(element_id) = element_id {
                fields.push(Field::String("elementId", element_id));
            }
            optional_u32_field(&mut fields, "row", *row);
            optional_u32_field(&mut fields, "column", *column);
            write_fields(writer, &fields)
        }
        SourceLocator::Xlsx {
            kind,
            sheet_name,
            address,
            formula,
            drawing_id,
        } => {
            writer.string(kind)?;
            let mut fields = vec![Field::String("sheetName", sheet_name)];
            if let Some(address) = address {
                fields.push(Field::String("address", address));
            }
            if let Some(formula) = formula {
                fields.push(Field::String("formula", formula));
            }
            optional_u32_field(&mut fields, "drawingId", *drawing_id);
            write_fields(writer, &fields)
        }
        SourceLocator::Ods {
            kind,
            table_name,
            row,
            column,
            element_id,
            path,
        } => {
            writer.string(kind)?;
            let mut fields = vec![
                Field::String("tableName", table_name),
                Field::String("path", path),
            ];
            optional_u32_field(&mut fields, "row", *row);
            optional_u32_field(&mut fields, "column", *column);
            if let Some(element_id) = element_id {
                fields.push(Field::String("elementId", element_id));
            }
            write_fields(writer, &fields)
        }
        SourceLocator::Docx {
            kind,
            paragraph_id,
            paragraph_index,
            drawing_id,
            row,
            column,
            text_range,
            action,
        } => {
            writer.string(kind)?;
            let mut fields = Vec::new();
            if let Some(paragraph_id) = paragraph_id {
                fields.push(Field::String("paragraphId", paragraph_id));
            }
            optional_u32_field(&mut fields, "paragraphIndex", *paragraph_index);
            optional_u32_field(&mut fields, "drawingId", *drawing_id);
            optional_u32_field(&mut fields, "row", *row);
            optional_u32_field(&mut fields, "column", *column);
            if let Some((start, end)) = text_range {
                fields.push(Field::Unsigned("rangeStart", *start));
                fields.push(Field::Unsigned("rangeEnd", *end));
            }
            if let Some(action) = action {
                push_action_fields(
                    &mut fields,
                    action,
                    "clickKind",
                    "clickAction",
                    "clickTarget",
                    "clickTooltip",
                );
            }
            write_fields(writer, &fields)
        }
        SourceLocator::Odt {
            kind,
            element_id,
            path,
            row,
            column,
            text_range,
        } => {
            writer.string(kind)?;
            let mut fields = vec![Field::String("path", path)];
            if let Some(element_id) = element_id {
                fields.push(Field::String("elementId", element_id));
            }
            optional_u32_field(&mut fields, "row", *row);
            optional_u32_field(&mut fields, "column", *column);
            if let Some((start, end)) = text_range {
                fields.push(Field::Unsigned("rangeStart", *start));
                fields.push(Field::Unsigned("rangeEnd", *end));
            }
            write_fields(writer, &fields)
        }
        SourceLocator::Flat {
            kind,
            index,
            row,
            column,
            text_range,
        } => {
            writer.string(kind)?;
            let mut fields = Vec::new();
            optional_u32_field(&mut fields, "index", *index);
            optional_u32_field(&mut fields, "row", *row);
            optional_u32_field(&mut fields, "column", *column);
            if let Some((start, end)) = text_range {
                fields.push(Field::Unsigned("rangeStart", *start));
                fields.push(Field::Unsigned("rangeEnd", *end));
            }
            write_fields(writer, &fields)
        }
        SourceLocator::Legacy {
            kind,
            stream,
            record_offset,
            row,
            column,
            text_range,
        } => {
            writer.string(kind)?;
            let mut fields = vec![Field::String("stream", stream)];
            optional_u32_field(&mut fields, "recordOffset", *record_offset);
            optional_u32_field(&mut fields, "row", *row);
            optional_u32_field(&mut fields, "column", *column);
            if let Some((start, end)) = text_range {
                fields.push(Field::Unsigned("rangeStart", *start));
                fields.push(Field::Unsigned("rangeEnd", *end));
            }
            write_fields(writer, &fields)
        }
        SourceLocator::Iwork { kind, component } => {
            writer.string(kind)?;
            let fields = [Field::String("component", component)];
            write_fields(writer, &fields)
        }
        SourceLocator::Pdf {
            kind,
            object_number,
            byte_offset,
            action,
        } => {
            writer.string(kind)?;
            let mut fields = Vec::new();
            optional_u32_field(&mut fields, "objectNumber", *object_number);
            optional_u32_field(&mut fields, "byteOffset", *byte_offset);
            if let Some(action) = action {
                push_action_fields(
                    &mut fields,
                    action,
                    "clickKind",
                    "clickAction",
                    "clickTarget",
                    "clickTooltip",
                );
            }
            write_fields(writer, &fields)
        }
        SourceLocator::Xps { kind, path } => {
            writer.string(kind)?;
            let fields = [Field::String("path", path)];
            write_fields(writer, &fields)
        }
        SourceLocator::Ofd { kind, path } => {
            writer.string(kind)?;
            let fields = [Field::String("path", path)];
            write_fields(writer, &fields)
        }
    }
}

fn write_path_commands(writer: &mut Writer, commands: &[PathCommand]) -> Result<(), ProtocolError> {
    writer.u32(
        commands
            .len()
            .try_into()
            .map_err(|_| ProtocolError::TooManyFields)?,
    );
    for command in commands {
        writer.u8(command.code());
        writer.padding(3);
        match command {
            PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => {
                writer.f32(*x);
                writer.f32(*y);
            }
            PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => {
                writer.f32(*cpx);
                writer.f32(*cpy);
                writer.f32(*x);
                writer.f32(*y);
            }
            PathCommand::BezierCurveTo {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                writer.f32(*cp1x);
                writer.f32(*cp1y);
                writer.f32(*cp2x);
                writer.f32(*cp2y);
                writer.f32(*x);
                writer.f32(*y);
            }
            PathCommand::ClosePath => {}
        }
    }
    Ok(())
}

fn write_geometry(writer: &mut Writer, geometry: &Geometry) -> Result<(), ProtocolError> {
    writer.u8(geometry.code());
    match geometry {
        Geometry::Rectangle | Geometry::Ellipse | Geometry::Line => Ok(()),
        Geometry::RoundedRectangle { radius_x, radius_y } => {
            writer.f32(*radius_x);
            writer.f32(*radius_y);
            Ok(())
        }
        Geometry::Path {
            fill_rule,
            commands,
        } => {
            writer.u8(fill_rule.code());
            writer.padding(3);
            write_path_commands(writer, commands)
        }
        Geometry::LayeredPath { layers } => {
            writer.u32(
                layers
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for layer in layers {
                writer.u8(layer.fill_rule.code());
                writer.u8(layer.fill.code());
                writer.u8(u8::from(layer.stroke));
                writer.padding(1);
                write_path_commands(writer, &layer.commands)?;
            }
            Ok(())
        }
    }
}

fn write_paint(writer: &mut Writer, paint: &Paint, depth: usize) -> Result<(), ProtocolError> {
    if depth > MAX_VISUAL_DEPTH {
        return Err(ProtocolError::TooManyFields);
    }
    writer.u8(paint.code());
    writer.padding(3);
    match paint {
        Paint::None => Ok(()),
        Paint::Solid(color) => {
            writer.u32(*color);
            Ok(())
        }
        Paint::LinearGradient {
            x0,
            y0,
            x1,
            y1,
            stops,
        } => {
            writer.f32(*x0);
            writer.f32(*y0);
            writer.f32(*x1);
            writer.f32(*y1);
            writer.u32(
                stops
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for stop in stops {
                writer.f32(stop.offset);
                writer.u32(stop.color);
            }
            Ok(())
        }
        Paint::RadialGradient {
            x0,
            y0,
            r0,
            x1,
            y1,
            r1,
            stops,
        } => {
            writer.f32(*x0);
            writer.f32(*y0);
            writer.f32(*r0);
            writer.f32(*x1);
            writer.f32(*y1);
            writer.f32(*r1);
            writer.u32(
                stops
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for stop in stops {
                writer.f32(stop.offset);
                writer.u32(stop.color);
            }
            Ok(())
        }
        Paint::RectGradient {
            center_x,
            center_y,
            stops,
        } => {
            writer.f32(*center_x);
            writer.f32(*center_y);
            writer.u32(
                stops
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for stop in stops {
                writer.f32(stop.offset);
                writer.u32(stop.color);
            }
            Ok(())
        }
        Paint::CircleGradient {
            x0,
            y0,
            r0,
            x1,
            y1,
            r1,
            stops,
        } => {
            writer.f32(*x0);
            writer.f32(*y0);
            writer.f32(*r0);
            writer.f32(*x1);
            writer.f32(*y1);
            writer.f32(*r1);
            writer.u32(
                stops
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for stop in stops {
                writer.f32(stop.offset);
                writer.u32(stop.color);
            }
            Ok(())
        }
        Paint::MappedGradient {
            paint,
            tile,
            flip,
            rotate_with_shape,
        } => {
            for value in [tile.left, tile.top, tile.right, tile.bottom] {
                writer.f32(value);
            }
            writer.u8(flip.code());
            writer.u8(u8::from(*rotate_with_shape));
            writer.u16(0);
            write_paint(writer, paint, depth + 1)
        }
        Paint::ShapeGradient { focus, stops } => {
            for value in [focus.left, focus.top, focus.right, focus.bottom] {
                writer.f32(value);
            }
            writer.u32(
                stops
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for stop in stops {
                writer.f32(stop.offset);
                writer.u32(stop.color);
            }
            Ok(())
        }
        Paint::Pattern {
            preset,
            foreground,
            background,
        } => {
            writer.string(preset)?;
            writer.u32(*foreground);
            writer.u32(*background);
            Ok(())
        }
        Paint::Image {
            mapping,
            media_type,
            bytes,
            crop,
            tile,
            tile_width,
            tile_height,
        } => {
            writer.f32(crop.left);
            writer.f32(crop.top);
            writer.f32(crop.right);
            writer.f32(crop.bottom);
            writer.f32(tile_width.unwrap_or(0.0));
            writer.f32(tile_height.unwrap_or(0.0));
            writer.u8(u8::from(*tile) | (u8::from(mapping.is_some()) << 1));
            writer.padding(3);
            if let Some(mapping) = mapping {
                for value in [
                    mapping.scale_x,
                    mapping.scale_y,
                    mapping.offset_x,
                    mapping.offset_y,
                    mapping.alignment_x,
                    mapping.alignment_y,
                    mapping.fill_rectangle.left,
                    mapping.fill_rectangle.top,
                    mapping.fill_rectangle.right,
                    mapping.fill_rectangle.bottom,
                    mapping.dpi,
                ] {
                    writer.f32(value);
                }
                writer.u8(mapping.flip.code());
                writer.u8(u8::from(mapping.rotate_with_shape));
                writer.padding(2);
            }
            writer.string(media_type)?;
            writer.byte_array(bytes)
        }
        Paint::Visual {
            viewbox,
            viewport,
            viewbox_relative,
            viewport_relative,
            tile_mode,
            stretch,
            alignment_x,
            alignment_y,
            transform,
            relative_transform,
            opacity,
            children,
        } => {
            for value in [
                viewbox.x,
                viewbox.y,
                viewbox.width,
                viewbox.height,
                viewport.x,
                viewport.y,
                viewport.width,
                viewport.height,
                transform.a,
                transform.b,
                transform.c,
                transform.d,
                transform.e,
                transform.f,
                relative_transform.a,
                relative_transform.b,
                relative_transform.c,
                relative_transform.d,
                relative_transform.e,
                relative_transform.f,
                *opacity,
                *alignment_x,
                *alignment_y,
            ] {
                writer.f32(value);
            }
            writer.u8(tile_mode.code());
            writer.u8(stretch.code());
            writer.u8(u8::from(*viewbox_relative) | (u8::from(*viewport_relative) << 1));
            writer.padding(1);
            writer.u32(
                children
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for child in children {
                writer.f32(child.bounds.x);
                writer.f32(child.bounds.y);
                writer.f32(child.bounds.width);
                writer.f32(child.bounds.height);
                writer.u8(child.visual.code());
                writer.padding(3);
                write_visual_at_depth(writer, &child.visual, depth + 1)?;
            }
            Ok(())
        }
        Paint::XpsGradient {
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
            transform,
            relative_transform,
            stops,
        } => {
            for value in [
                *start_x,
                *start_y,
                *end_x,
                *end_y,
                *radius_x,
                *radius_y,
                transform.a,
                transform.b,
                transform.c,
                transform.d,
                transform.e,
                transform.f,
                relative_transform.a,
                relative_transform.b,
                relative_transform.c,
                relative_transform.d,
                relative_transform.e,
                relative_transform.f,
            ] {
                writer.f32(value);
            }
            writer
                .u8(u8::from(*radial) | (u8::from(*relative) << 1) | (u8::from(*linear_rgb) << 2));
            writer.u8(spread.code());
            writer.padding(2);
            writer.u32(
                stops
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for stop in stops {
                writer.f32(stop.offset);
                write_xps_color(writer, &stop.color)?;
            }
            Ok(())
        }
    }
}

fn write_xps_color(writer: &mut Writer, color: &XpsColor) -> Result<(), ProtocolError> {
    match color {
        XpsColor::Rgba(color) => {
            writer.u8(0);
            writer.padding(3);
            writer.u32(*color);
        }
        XpsColor::Context {
            alpha,
            profile,
            channels,
        } => {
            writer.u8(1);
            writer.padding(3);
            writer.f32(*alpha);
            writer.byte_array(profile)?;
            writer.u32(
                channels
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for channel in channels {
                writer.f32(*channel);
            }
        }
    }
    Ok(())
}

fn finite(values: &[f32]) -> bool {
    values.iter().all(|value| value.is_finite())
}

fn validate_sheet_axis(axis: &SheetAxis, count: u32) -> Result<(), ProtocolError> {
    if !axis.default_size.is_finite() || axis.default_size < 0.0 {
        return Err(ProtocolError::InvalidValue);
    }
    let mut next = 0_u32;
    for span in &axis.spans {
        if span.start < next
            || span.start > span.end
            || span.end >= count
            || !span.size.is_finite()
            || span.size < 0.0
        {
            return Err(ProtocolError::InvalidValue);
        }
        next = span.end.saturating_add(1);
    }
    Ok(())
}

fn write_sheet_axis(writer: &mut Writer, axis: &SheetAxis) -> Result<(), ProtocolError> {
    writer.f32(axis.default_size);
    writer.u32(
        axis.spans
            .len()
            .try_into()
            .map_err(|_| ProtocolError::TooManyFields)?,
    );
    for span in &axis.spans {
        writer.u32(span.start);
        writer.u32(span.end);
        writer.f32(span.size);
    }
    Ok(())
}

fn validate_speaker_note_paragraph(paragraph: &SpeakerNoteParagraph) -> Result<(), ProtocolError> {
    let layout = paragraph.layout;
    if !finite(&[
        layout.margin_left,
        layout.margin_right,
        layout.first_line_indent,
        layout.default_tab_stop,
        layout.line_height,
        layout.space_before,
        layout.space_after,
    ]) || layout.margin_left < 0.0
        || layout.margin_right < 0.0
        || layout.default_tab_stop <= 0.0
        || layout.line_height < 0.0
        || layout.space_before < 0.0
        || layout.space_after < 0.0
        || paragraph.runs.len() > MAX_RICH_TEXT_RUNS
        || paragraph.runs.iter().any(|run| {
            !finite(&[
                run.font_size,
                run.letter_spacing,
                run.baseline_shift,
                run.horizontal_scale,
            ]) || run.font_size <= 0.0
                || run.horizontal_scale <= 0.0
        })
    {
        return Err(ProtocolError::InvalidValue);
    }
    for run in &paragraph.runs {
        if let Some(paint) = &run.paint {
            validate_paint(paint, 1)?;
        }
    }
    Ok(())
}

fn speaker_note_paragraph_length(paragraph: &SpeakerNoteParagraph) -> Result<usize, ProtocolError> {
    validate_speaker_note_paragraph(paragraph)?;
    let _: u32 = paragraph
        .runs
        .len()
        .try_into()
        .map_err(|_| ProtocolError::TooManyFields)?;
    let mut total = 40_usize;
    for run in &paragraph.runs {
        add_length(&mut total, string_length(&run.text)?)?;
        add_length(&mut total, string_length(&run.font_family)?)?;
        add_length(&mut total, 28)?;
        add_length(
            &mut total,
            paint_length(run.paint.as_deref().unwrap_or(&Paint::None), 1)?,
        )?;
    }
    Ok(total)
}

fn write_text_run(writer: &mut Writer, run: &TextRun, depth: usize) -> Result<(), ProtocolError> {
    writer.string(&run.text)?;
    writer.string(&run.font_family)?;
    writer.f32(run.font_size);
    writer.u32(run.color);
    writer.u8(u8::from(run.bold)
        | (u8::from(run.italic) << 1)
        | (u8::from(run.underline) << 2)
        | (u8::from(run.strikethrough) << 3)
        | (u8::from(!run.east_asian_line_breaks) << 4)
        | (u8::from(run.paint.is_some()) << 5));
    writer.padding(3);
    writer.f32(run.letter_spacing);
    writer.u32(run.highlight);
    writer.f32(run.baseline_shift);
    writer.f32(run.horizontal_scale);
    write_paint(
        writer,
        run.paint.as_deref().unwrap_or(&Paint::None),
        depth + 1,
    )
}

fn write_speaker_note_paragraph(
    writer: &mut Writer,
    paragraph: &SpeakerNoteParagraph,
) -> Result<(), ProtocolError> {
    let layout = paragraph.layout;
    writer.u8(layout.align.code());
    writer.padding(3);
    writer.f32(layout.margin_left);
    writer.f32(layout.margin_right);
    writer.f32(layout.first_line_indent);
    writer.f32(layout.default_tab_stop);
    writer.f32(layout.line_height);
    writer.f32(layout.space_before);
    writer.f32(layout.space_after);
    writer.u8(u8::from(layout.latin_line_break) | (u8::from(layout.hanging_punctuation) << 1));
    writer.padding(3);
    writer.u32(
        paragraph
            .runs
            .len()
            .try_into()
            .map_err(|_| ProtocolError::TooManyFields)?,
    );
    for run in &paragraph.runs {
        write_text_run(writer, run, 0)?;
    }
    Ok(())
}

fn validate_path_commands(commands: &[PathCommand]) -> Result<(), ProtocolError> {
    if commands.len() > MAX_PATH_COMMANDS {
        return Err(ProtocolError::TooManyFields);
    }
    for command in commands {
        let valid = match command {
            PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => finite(&[*x, *y]),
            PathCommand::QuadraticCurveTo { cpx, cpy, x, y } => finite(&[*cpx, *cpy, *x, *y]),
            PathCommand::BezierCurveTo {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => finite(&[*cp1x, *cp1y, *cp2x, *cp2y, *x, *y]),
            PathCommand::ClosePath => true,
        };
        if !valid {
            return Err(ProtocolError::InvalidValue);
        }
    }
    Ok(())
}

fn validate_geometry(geometry: &Geometry) -> Result<(), ProtocolError> {
    match geometry {
        Geometry::Rectangle | Geometry::Ellipse | Geometry::Line => Ok(()),
        Geometry::RoundedRectangle { radius_x, radius_y } => {
            if finite(&[*radius_x, *radius_y]) && *radius_x >= 0.0 && *radius_y >= 0.0 {
                Ok(())
            } else {
                Err(ProtocolError::InvalidValue)
            }
        }
        Geometry::Path { commands, .. } => validate_path_commands(commands),
        Geometry::LayeredPath { layers } => {
            if layers.len() > MAX_PATH_COMMANDS {
                return Err(ProtocolError::TooManyFields);
            }
            let mut total = 0_usize;
            for layer in layers {
                validate_path_commands(&layer.commands)?;
                total = total
                    .checked_add(layer.commands.len())
                    .ok_or(ProtocolError::TooManyFields)?;
                if total > MAX_PATH_COMMANDS {
                    return Err(ProtocolError::TooManyFields);
                }
            }
            Ok(())
        }
    }
}

fn validate_stops(stops: &[crate::model::GradientStop]) -> Result<(), ProtocolError> {
    if stops.is_empty() {
        return Err(ProtocolError::InvalidValue);
    }
    if stops.len() > MAX_GRADIENT_STOPS {
        return Err(ProtocolError::TooManyFields);
    }
    let mut previous = -1.0;
    for stop in stops {
        if !stop.offset.is_finite() || !(0.0..=1.0).contains(&stop.offset) || stop.offset < previous
        {
            return Err(ProtocolError::InvalidValue);
        }
        previous = stop.offset;
    }
    Ok(())
}

fn validate_xps_stops(stops: &[XpsGradientStop]) -> Result<(), ProtocolError> {
    if stops.is_empty() || stops.len() > MAX_GRADIENT_STOPS {
        return Err(ProtocolError::InvalidValue);
    }
    let mut previous = f32::NEG_INFINITY;
    for stop in stops {
        if !stop.offset.is_finite() || stop.offset < previous {
            return Err(ProtocolError::InvalidValue);
        }
        if let XpsColor::Context {
            alpha,
            profile,
            channels,
        } = &stop.color
            && (!alpha.is_finite()
                || !(0.0..=1.0).contains(alpha)
                || profile.is_empty()
                || !(1..=15).contains(&channels.len())
                || !finite(channels))
        {
            return Err(ProtocolError::InvalidValue);
        }
        previous = stop.offset;
    }
    Ok(())
}

fn validate_paint(paint: &Paint, depth: usize) -> Result<(), ProtocolError> {
    if depth > MAX_VISUAL_DEPTH {
        return Err(ProtocolError::TooManyFields);
    }
    match paint {
        Paint::None | Paint::Solid(_) => Ok(()),
        Paint::LinearGradient {
            x0,
            y0,
            x1,
            y1,
            stops,
        } => {
            if !finite(&[*x0, *y0, *x1, *y1]) {
                return Err(ProtocolError::InvalidValue);
            }
            validate_stops(stops)
        }
        Paint::RadialGradient {
            x0,
            y0,
            r0,
            x1,
            y1,
            r1,
            stops,
        } => {
            if !finite(&[*x0, *y0, *r0, *x1, *y1, *r1]) || *r0 < 0.0 || *r1 < 0.0 {
                return Err(ProtocolError::InvalidValue);
            }
            validate_stops(stops)
        }
        Paint::RectGradient {
            center_x,
            center_y,
            stops,
        } => {
            if !finite(&[*center_x, *center_y]) {
                return Err(ProtocolError::InvalidValue);
            }
            validate_stops(stops)
        }
        Paint::CircleGradient {
            x0,
            y0,
            r0,
            x1,
            y1,
            r1,
            stops,
        } => {
            if !finite(&[*x0, *y0, *r0, *x1, *y1, *r1]) || *r0 < 0.0 || *r1 < 0.0 {
                return Err(ProtocolError::InvalidValue);
            }
            validate_stops(stops)
        }
        Paint::MappedGradient { paint, tile, .. } => {
            if !finite(&[tile.left, tile.top, tile.right, tile.bottom])
                || tile.left + tile.right >= 1.0
                || tile.top + tile.bottom >= 1.0
            {
                return Err(ProtocolError::InvalidValue);
            }
            validate_paint(paint, depth + 1)
        }
        Paint::ShapeGradient { focus, stops } => {
            if !finite(&[focus.left, focus.top, focus.right, focus.bottom]) {
                return Err(ProtocolError::InvalidValue);
            }
            validate_stops(stops)
        }
        Paint::Pattern { preset, .. } => {
            if preset.is_empty() || preset.len() > 64 {
                Err(ProtocolError::InvalidValue)
            } else {
                Ok(())
            }
        }
        Paint::Image {
            mapping,
            media_type,
            bytes,
            crop,
            tile_width,
            tile_height,
            ..
        } => {
            if mapping.as_ref().is_some_and(|m| {
                !finite(&[
                    m.scale_x,
                    m.scale_y,
                    m.offset_x,
                    m.offset_y,
                    m.alignment_x,
                    m.alignment_y,
                    m.fill_rectangle.left,
                    m.fill_rectangle.top,
                    m.fill_rectangle.right,
                    m.fill_rectangle.bottom,
                    m.dpi,
                ]) || m.scale_x <= 0.0
                    || m.scale_y <= 0.0
                    || m.dpi < 0.0
            }) || media_type.is_empty()
                || bytes.is_empty()
                || !crop.is_valid()
                || tile_width.is_some_and(|value| !value.is_finite() || value <= 0.0)
                || tile_height.is_some_and(|value| !value.is_finite() || value <= 0.0)
            {
                Err(ProtocolError::InvalidValue)
            } else {
                Ok(())
            }
        }
        Paint::Visual {
            viewbox,
            viewport,
            alignment_x,
            alignment_y,
            transform,
            relative_transform,
            opacity,
            children,
            ..
        } => {
            if !viewbox.is_valid()
                || !viewport.is_valid()
                || viewbox.width == 0.0
                || viewbox.height == 0.0
                || viewport.width == 0.0
                || viewport.height == 0.0
                || !alignment_x.is_finite()
                || !alignment_y.is_finite()
                || !(0.0..=1.0).contains(alignment_x)
                || !(0.0..=1.0).contains(alignment_y)
                || !transform.is_valid()
                || !relative_transform.is_valid()
                || !opacity.is_finite()
                || !(0.0..=1.0).contains(opacity)
                || children.len() > MAX_VISUAL_BRUSH_CHILDREN
            {
                return Err(ProtocolError::InvalidValue);
            }
            for child in children {
                if !child.bounds.is_valid() {
                    return Err(ProtocolError::InvalidValue);
                }
                validate_visual(&child.visual, depth + 1)?;
            }
            Ok(())
        }
        Paint::XpsGradient {
            start_x,
            start_y,
            end_x,
            end_y,
            radius_x,
            radius_y,
            transform,
            relative_transform,
            stops,
            ..
        } => {
            if !finite(&[*start_x, *start_y, *end_x, *end_y, *radius_x, *radius_y])
                || *radius_x < 0.0
                || *radius_y < 0.0
                || !transform.is_valid()
                || !relative_transform.is_valid()
            {
                return Err(ProtocolError::InvalidValue);
            }
            validate_xps_stops(stops)
        }
    }
}

fn validate_visual(visual: &Visual, depth: usize) -> Result<(), ProtocolError> {
    if depth > MAX_VISUAL_DEPTH {
        return Err(ProtocolError::TooManyFields);
    }
    match visual {
        Visual::None => Ok(()),
        Visual::Shape {
            geometry,
            stroke_width,
            ..
        } => {
            validate_geometry(geometry)?;
            if stroke_width.is_finite() && *stroke_width >= 0.0 {
                Ok(())
            } else {
                Err(ProtocolError::InvalidValue)
            }
        }
        Visual::Text {
            geometry,
            stroke_width,
            font_size,
            ..
        } => {
            validate_geometry(geometry)?;
            if finite(&[*stroke_width, *font_size]) && *stroke_width >= 0.0 && *font_size > 0.0 {
                Ok(())
            } else {
                Err(ProtocolError::InvalidValue)
            }
        }
        Visual::Image { crop, .. } | Visual::ImageWithFallback { crop, .. } => {
            if crop.is_valid() {
                Ok(())
            } else {
                Err(ProtocolError::InvalidValue)
            }
        }
        Visual::MaskedImage {
            media_type,
            bytes,
            resource_id,
            mask_resource_id,
            mask_width,
            mask_height,
            mask_media_type,
            alpha_mask,
            crop,
        } => {
            let pixels = usize::try_from(*mask_width).ok().and_then(|width| {
                usize::try_from(*mask_height)
                    .ok()
                    .and_then(|height| width.checked_mul(height))
            });
            if crop.is_valid()
                && !media_type.is_empty()
                && (!bytes.is_empty() || *resource_id != 0)
                && pixels.is_some_and(|pixels| pixels > 0)
                && !mask_media_type.is_empty()
                && (!alpha_mask.is_empty() || *mask_resource_id != 0)
            {
                Ok(())
            } else {
                Err(ProtocolError::InvalidValue)
            }
        }
        Visual::Group { children } => {
            if children.len() > MAX_VISUAL_BRUSH_CHILDREN {
                return Err(ProtocolError::TooManyFields);
            }
            for child in children {
                if !child.bounds.is_valid() {
                    return Err(ProtocolError::InvalidValue);
                }
                validate_visual(&child.visual, depth + 1)?;
            }
            Ok(())
        }
        Visual::OpacityMask { mask, visual } => {
            validate_paint(mask, depth)?;
            validate_visual(visual, depth + 1)
        }
        Visual::ColorManagedImage {
            source_profile,
            destination_profile,
            visual,
        } => {
            if source_profile.is_empty() || destination_profile.as_ref().is_some_and(Vec::is_empty)
            {
                return Err(ProtocolError::InvalidValue);
            }
            validate_visual(visual, depth + 1)
        }
        Visual::Media {
            media_type,
            bytes,
            poster,
            ..
        } => {
            if media_type.is_empty() || bytes.is_empty() {
                return Err(ProtocolError::InvalidValue);
            }
            validate_visual(poster, depth + 1)
        }
        Visual::ImageColorChange { visual, .. } => validate_visual(visual, depth + 1),
        Visual::ImageAdjustment { adjustment, visual } => {
            if !adjustment.is_valid() {
                return Err(ProtocolError::InvalidValue);
            }
            validate_visual(visual, depth + 1)
        }
        Visual::Layer {
            transform,
            opacity,
            visual,
            ..
        } => {
            if !transform.is_valid() || !opacity.is_finite() || !(0.0..=1.0).contains(opacity) {
                return Err(ProtocolError::InvalidValue);
            }
            validate_visual(visual, depth + 1)
        }
        Visual::PaintedShape {
            geometry,
            fill,
            stroke,
            stroke_width,
        } => {
            validate_geometry(geometry)?;
            validate_paint(fill, depth)?;
            validate_paint(stroke, depth)?;
            if stroke_width.is_finite() && *stroke_width >= 0.0 {
                Ok(())
            } else {
                Err(ProtocolError::InvalidValue)
            }
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
            validate_geometry(geometry)?;
            validate_paint(fill, depth)?;
            validate_paint(stroke, depth)?;
            if !finite(&[*stroke_width, *line_height])
                || *stroke_width < 0.0
                || *line_height < 0.0
                || runs.len() > MAX_RICH_TEXT_RUNS
            {
                return Err(ProtocolError::InvalidValue);
            }
            if runs.iter().any(|run| {
                !finite(&[
                    run.font_size,
                    run.letter_spacing,
                    run.baseline_shift,
                    run.horizontal_scale,
                ]) || run.font_size <= 0.0
                    || run.horizontal_scale <= 0.0
            }) {
                return Err(ProtocolError::InvalidValue);
            }
            for run in runs {
                if let Some(paint) = &run.paint {
                    validate_paint(paint, depth + 1)?;
                }
            }
            Ok(())
        }
        Visual::Effect {
            shadow,
            clip,
            visual,
        } => {
            if shadow.is_none() && clip.is_none() {
                return Err(ProtocolError::InvalidValue);
            }
            if shadow.is_some_and(|shadow| {
                !finite(&[shadow.blur, shadow.offset_x, shadow.offset_y]) || shadow.blur < 0.0
            }) {
                return Err(ProtocolError::InvalidValue);
            }
            if let Some(clip) = clip {
                validate_geometry(clip)?;
            }
            validate_visual(visual, depth + 1)
        }
        Visual::TextLayout { layout, visual } => {
            if layout
                .fill_character
                .is_some_and(|(offset, character)| offset == NONE || character.is_control())
            {
                return Err(ProtocolError::InvalidValue);
            }
            if layout.tab_stops.len() > MAX_TAB_STOPS
                || layout.paragraphs.len() > MAX_TEXT_PARAGRAPHS
            {
                return Err(ProtocolError::TooManyFields);
            }
            if !finite(&[
                layout.default_tab_stop,
                layout.hanging_indent,
                layout.paragraph_spacing,
                layout.inset_left,
                layout.inset_right,
                layout.inset_top,
                layout.inset_bottom,
                layout.margin_left,
                layout.margin_right,
                layout.first_line_indent,
                layout.column_spacing,
                layout.rotation_degrees,
                layout.font_scale,
                layout.line_spacing_reduction,
                layout.min_scale,
                layout.text_stroke_width,
                layout.text_baseline,
            ]) || layout.default_tab_stop <= 0.0
                || layout.hanging_indent < 0.0
                || layout.paragraph_spacing < 0.0
                || layout.inset_left < 0.0
                || layout.inset_right < 0.0
                || layout.inset_top < 0.0
                || layout.inset_bottom < 0.0
                || layout.margin_left < 0.0
                || layout.margin_right < 0.0
                || layout.column_count == 0
                || layout.column_count > 64
                || layout.column_spacing < 0.0
                || !(0.0..=1.0).contains(&layout.font_scale)
                || layout.font_scale == 0.0
                || !(0.0..1.0).contains(&layout.line_spacing_reduction)
                || !(0.0..=1.0).contains(&layout.min_scale)
                || layout.min_scale == 0.0
                || layout.text_stroke_width < 0.0
                || layout.text_baseline < 0.0
            {
                return Err(ProtocolError::InvalidValue);
            }
            if layout.wrap_regions.len() > 1024
                || layout.wrap_regions.iter().any(|rect| {
                    !finite(&[rect.x, rect.y, rect.width, rect.height])
                        || rect.width <= 0.0
                        || rect.height <= 0.0
                })
            {
                return Err(ProtocolError::InvalidValue);
            }
            let mut previous = 0.0;
            for stop in &layout.tab_stops {
                if !stop.position.is_finite() || stop.position <= previous || stop.align.code() > 2
                {
                    return Err(ProtocolError::InvalidValue);
                }
                previous = stop.position;
            }
            if layout.paragraphs.iter().any(|paragraph| {
                !finite(&[
                    paragraph.margin_left,
                    paragraph.margin_right,
                    paragraph.first_line_indent,
                    paragraph.default_tab_stop,
                    paragraph.line_height,
                    paragraph.space_before,
                    paragraph.space_after,
                ]) || paragraph.margin_left < 0.0
                    || paragraph.margin_right < 0.0
                    || paragraph.default_tab_stop <= 0.0
                    || paragraph.line_height < 0.0
                    || paragraph.space_before < 0.0
                    || paragraph.space_after < 0.0
                    || paragraph.drop_cap.is_some_and(|cap| {
                        !(1..=32).contains(&cap.characters)
                            || !(1..=32).contains(&cap.lines)
                            || cap.raised_lines > 32
                            || !finite(&[cap.padding, cap.outdent])
                            || cap.padding < 0.0
                            || cap.padding > 2048.0
                            || cap.outdent.abs() > 2048.0
                    })
                    || [paragraph.rule_above, paragraph.rule_below]
                        .into_iter()
                        .flatten()
                        .any(|rule| {
                            !finite(&[rule.stroke_width, rule.offset_x, rule.offset_y, rule.width])
                                || rule.stroke_width <= 0.0
                                || rule.width <= 0.0
                        })
            }) {
                return Err(ProtocolError::InvalidValue);
            }
            validate_paint(&layout.text_paint, depth)?;
            validate_paint(&layout.text_stroke_paint, depth)?;
            validate_visual(visual, depth + 1)
        }
        Visual::TextEffects { effects, visual } => {
            if effects.is_empty()
                || effects.len() > MAX_RICH_TEXT_RUNS
                || effects.iter().all(|effect| {
                    effect.shadow.is_none()
                        && effect.inner_shadow.is_none()
                        && effect.reflection.is_none()
                        && effect.stroke.is_none()
                        && effect.glow.is_none()
                        && !effect.fill_to_text
                        && !effect.wavy_underline
                        && !effect.dotted_underline
                        && !effect.heavy_underline
                        && !effect.double_underline
                        && !effect.dot_dash_underline
                        && !effect.double_strikethrough
                })
                || effects.iter().any(|effect| {
                    !finite(&[
                        effect.shadow_scale_x,
                        effect.shadow_scale_y,
                        effect.shadow_skew_x,
                        effect.shadow_skew_y,
                    ]) || effect.shadow_alignment > 8
                        || !effect.stroke_width.is_finite()
                        || effect.stroke_width < 0.0
                        || effect
                            .glow
                            .is_some_and(|glow| !glow.radius.is_finite() || glow.radius < 0.0)
                        || [effect.shadow, effect.inner_shadow]
                            .into_iter()
                            .flatten()
                            .any(|shadow| {
                                !finite(&[shadow.blur, shadow.offset_x, shadow.offset_y])
                                    || shadow.blur < 0.0
                            })
                        || effect.reflection.is_some_and(|reflection| {
                            !finite(&[
                                reflection.start_opacity,
                                reflection.end_opacity,
                                reflection.start_position,
                                reflection.end_position,
                                reflection.direction_degrees,
                                reflection.blur,
                                reflection.distance,
                                reflection.scale_x,
                                reflection.scale_y,
                            ]) || reflection.blur < 0.0
                        })
                })
            {
                return Err(ProtocolError::InvalidValue);
            }
            for effect in effects {
                if let Some(stroke) = &effect.stroke {
                    validate_paint(stroke, depth + 1)?;
                }
            }
            validate_visual(visual, depth + 1)
        }
        Visual::StrokeStyle { style, visual } => {
            if !style.dash_offset.is_finite()
                || !style.miter_limit.is_finite()
                || style.miter_limit <= 0.0
                || style.dash.len() > 256
                || style
                    .dash
                    .iter()
                    .any(|value| !value.is_finite() || *value < 0.0)
                || (!style.dash.is_empty() && !style.dash.iter().any(|value| *value > 0.0))
            {
                return Err(ProtocolError::InvalidValue);
            }
            validate_visual(visual, depth + 1)
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
            if outer_shadow.is_none()
                && inner_shadow.is_none()
                && glow.is_none()
                && reflection.is_none()
                && soft_edge.is_none()
                && three_d.is_none()
            {
                return Err(ProtocolError::InvalidValue);
            }
            if outer_shadow.is_some_and(|effect| {
                !finite(&[
                    effect.shadow.blur,
                    effect.shadow.offset_x,
                    effect.shadow.offset_y,
                    effect.scale_x,
                    effect.scale_y,
                    effect.skew_x,
                    effect.skew_y,
                ]) || effect.shadow.blur < 0.0
                    || effect.alignment > 8
            }) || inner_shadow.is_some_and(|shadow| {
                !finite(&[shadow.blur, shadow.offset_x, shadow.offset_y]) || shadow.blur < 0.0
            }) || glow.is_some_and(|glow| !glow.radius.is_finite() || glow.radius < 0.0)
                || reflection.is_some_and(|reflection| {
                    !finite(&[
                        reflection.start_opacity,
                        reflection.end_opacity,
                        reflection.start_position,
                        reflection.end_position,
                        reflection.direction_degrees,
                        reflection.blur,
                        reflection.distance,
                        reflection.scale_x,
                        reflection.scale_y,
                    ]) || !(0.0..=1.0).contains(&reflection.start_opacity)
                        || !(0.0..=1.0).contains(&reflection.end_opacity)
                        || !(0.0..=1.0).contains(&reflection.start_position)
                        || !(0.0..=1.0).contains(&reflection.end_position)
                        || reflection.start_position > reflection.end_position
                        || reflection.blur < 0.0
                        || reflection.scale_x <= 0.0
                        || reflection.scale_y <= 0.0
                })
                || soft_edge.is_some_and(|radius| !radius.is_finite() || radius < 0.0)
                || three_d.as_ref().is_some_and(|style| {
                    !finite(&[
                        style.camera_fov,
                        style.camera_zoom,
                        style.camera_latitude,
                        style.camera_longitude,
                        style.camera_revolution,
                        style.light_latitude,
                        style.light_longitude,
                        style.light_revolution,
                        style.z,
                        style.extrusion_height,
                        style.contour_width,
                        style.flat_text_z.unwrap_or(0.0),
                    ]) || style.camera_fov < 0.0
                        || style.camera_zoom <= 0.0
                        || style.extrusion_height < 0.0
                        || style.contour_width < 0.0
                        || style.camera_preset.is_empty()
                        || style.light_rig.is_empty()
                        || style.light_direction.is_empty()
                        || style.material.is_empty()
                        || style.backdrop.is_some_and(|backdrop| {
                            !finite(&[
                                backdrop.anchor_x,
                                backdrop.anchor_y,
                                backdrop.anchor_z,
                                backdrop.normal_x,
                                backdrop.normal_y,
                                backdrop.normal_z,
                                backdrop.up_x,
                                backdrop.up_y,
                                backdrop.up_z,
                            ]) || backdrop.normal_x.abs()
                                + backdrop.normal_y.abs()
                                + backdrop.normal_z.abs()
                                == 0.0
                                || backdrop.up_x.abs() + backdrop.up_y.abs() + backdrop.up_z.abs()
                                    == 0.0
                        })
                        || style.bevel_top.as_ref().is_some_and(|bevel| {
                            !finite(&[bevel.width, bevel.height])
                                || bevel.width < 0.0
                                || bevel.height < 0.0
                                || bevel.preset.is_empty()
                        })
                        || style.bevel_bottom.as_ref().is_some_and(|bevel| {
                            !finite(&[bevel.width, bevel.height])
                                || bevel.width < 0.0
                                || bevel.height < 0.0
                                || bevel.preset.is_empty()
                        })
                })
            {
                return Err(ProtocolError::InvalidValue);
            }
            validate_visual(visual, depth + 1)
        }
    }
}

fn write_visual(writer: &mut Writer, visual: &Visual) -> Result<(), ProtocolError> {
    validate_visual(visual, 0)?;
    write_visual_at_depth(writer, visual, 0)
}

fn write_visual_at_depth(
    writer: &mut Writer,
    visual: &Visual,
    depth: usize,
) -> Result<(), ProtocolError> {
    if depth > MAX_VISUAL_DEPTH {
        return Err(ProtocolError::TooManyFields);
    }
    match visual {
        Visual::None => Ok(()),
        Visual::Shape {
            geometry,
            fill,
            stroke,
            stroke_width,
        } => {
            write_geometry(writer, geometry)?;
            writer.u32(*fill);
            writer.u32(*stroke);
            writer.f32(*stroke_width);
            Ok(())
        }
        Visual::Text {
            geometry,
            fill,
            stroke,
            stroke_width,
            font_family,
            font_size,
            color,
            bold,
            italic,
            align,
        } => {
            write_geometry(writer, geometry)?;
            writer.u32(*fill);
            writer.u32(*stroke);
            writer.f32(*stroke_width);
            writer.string(font_family)?;
            writer.f32(*font_size);
            writer.u32(*color);
            writer.u8(u8::from(*bold) | (u8::from(*italic) << 1));
            writer.u8(align.code());
            writer.padding(2);
            Ok(())
        }
        Visual::Image {
            media_type,
            bytes,
            crop,
        } => {
            writer.f32(crop.left);
            writer.f32(crop.top);
            writer.f32(crop.right);
            writer.f32(crop.bottom);
            writer.string(media_type)?;
            writer.byte_array(bytes)
        }
        Visual::ImageWithFallback {
            media_type,
            bytes,
            fallback_media_type,
            fallback_bytes,
            crop,
        } => {
            writer.f32(crop.left);
            writer.f32(crop.top);
            writer.f32(crop.right);
            writer.f32(crop.bottom);
            writer.string(media_type)?;
            writer.byte_array(bytes)?;
            writer.string(fallback_media_type)?;
            writer.byte_array(fallback_bytes)
        }
        Visual::MaskedImage {
            media_type,
            bytes,
            resource_id,
            mask_resource_id,
            mask_width,
            mask_height,
            mask_media_type,
            alpha_mask,
            crop,
        } => {
            writer.f32(crop.left);
            writer.f32(crop.top);
            writer.f32(crop.right);
            writer.f32(crop.bottom);
            writer.string(media_type)?;
            writer.byte_array(bytes)?;
            writer.u32(*resource_id);
            writer.u32(*mask_resource_id);
            writer.u32(*mask_width);
            writer.u32(*mask_height);
            writer.string(mask_media_type)?;
            writer.byte_array(alpha_mask)
        }
        Visual::Group { children } => {
            writer.u32(
                children
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for child in children {
                writer.f32(child.bounds.x);
                writer.f32(child.bounds.y);
                writer.f32(child.bounds.width);
                writer.f32(child.bounds.height);
                writer.u8(child.visual.code());
                writer.padding(3);
                write_visual_at_depth(writer, &child.visual, depth + 1)?;
            }
            Ok(())
        }
        Visual::OpacityMask { mask, visual } => {
            write_paint(writer, mask, depth)?;
            writer.u8(visual.code());
            writer.padding(3);
            write_visual_at_depth(writer, visual, depth + 1)
        }
        Visual::ColorManagedImage {
            source_profile,
            destination_profile,
            visual,
        } => {
            writer.byte_array(source_profile)?;
            writer.u8(u8::from(destination_profile.is_some()));
            writer.padding(3);
            if let Some(profile) = destination_profile {
                writer.byte_array(profile)?;
            }
            writer.u8(visual.code());
            writer.padding(3);
            write_visual_at_depth(writer, visual, depth + 1)
        }
        Visual::Media {
            kind,
            media_type,
            bytes,
            poster,
        } => {
            writer.u8(kind.code());
            writer.padding(3);
            writer.string(media_type)?;
            writer.byte_array(bytes)?;
            writer.u8(poster.code());
            writer.padding(3);
            write_visual_at_depth(writer, poster, depth + 1)
        }
        Visual::Layer {
            transform,
            opacity,
            blend_mode,
            visual,
        } => {
            writer.f32(transform.a);
            writer.f32(transform.b);
            writer.f32(transform.c);
            writer.f32(transform.d);
            writer.f32(transform.e);
            writer.f32(transform.f);
            writer.f32(*opacity);
            writer.u8(visual.code());
            writer.u8(blend_mode.code());
            writer.padding(2);
            write_visual_at_depth(writer, visual, depth + 1)
        }
        Visual::ImageColorChange {
            from,
            to,
            use_alpha,
            visual,
        } => {
            writer.u32(*from);
            writer.u32(*to);
            writer.u8(u8::from(*use_alpha));
            writer.u8(visual.code());
            writer.padding(2);
            write_visual_at_depth(writer, visual, depth + 1)
        }
        Visual::ImageAdjustment { adjustment, visual } => {
            let flags = u8::from(adjustment.grayscale)
                | (u8::from(adjustment.bilevel_threshold.is_some()) << 1)
                | (u8::from(adjustment.duotone.is_some()) << 2);
            writer.u8(flags);
            writer.u8(visual.code());
            writer.padding(2);
            writer.f32(adjustment.bilevel_threshold.unwrap_or(0.0));
            writer.f32(adjustment.brightness);
            writer.f32(adjustment.contrast);
            let [first, second] = adjustment.duotone.unwrap_or([0, 0]);
            writer.u32(first);
            writer.u32(second);
            write_visual_at_depth(writer, visual, depth + 1)
        }
        Visual::PaintedShape {
            geometry,
            fill,
            stroke,
            stroke_width,
        } => {
            write_geometry(writer, geometry)?;
            write_paint(writer, fill, depth)?;
            write_paint(writer, stroke, depth)?;
            writer.f32(*stroke_width);
            Ok(())
        }
        Visual::RichText {
            geometry,
            fill,
            stroke,
            stroke_width,
            align,
            line_height,
            runs,
        } => {
            write_geometry(writer, geometry)?;
            write_paint(writer, fill, depth)?;
            write_paint(writer, stroke, depth)?;
            writer.f32(*stroke_width);
            writer.u8(align.code());
            writer.padding(3);
            writer.f32(*line_height);
            writer.u32(
                runs.len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for run in runs {
                write_text_run(writer, run, depth)?;
            }
            Ok(())
        }
        Visual::Effect {
            shadow,
            clip,
            visual,
        } => {
            writer.u8(u8::from(shadow.is_some()) | (u8::from(clip.is_some()) << 1));
            writer.padding(3);
            if let Some(shadow) = shadow {
                writer.u32(shadow.color);
                writer.f32(shadow.blur);
                writer.f32(shadow.offset_x);
                writer.f32(shadow.offset_y);
            }
            if let Some(clip) = clip {
                write_geometry(writer, clip)?;
            }
            writer.u8(visual.code());
            writer.padding(3);
            write_visual_at_depth(writer, visual, depth + 1)
        }
        Visual::TextLayout { layout, visual } => {
            writer.u8(layout.direction.code());
            writer.u8(layout.orientation.code());
            writer.u8(layout.auto_fit.code());
            writer.u8(u8::from(layout.prefix.is_some())
                | (layout.vertical_align.code() << 1)
                | (u8::from(layout.continues_after) << 3)
                | (u8::from(layout.compress_punctuation) << 4)
                | (u8::from(layout.fixed_line_height) << 5));
            writer.f32(layout.default_tab_stop);
            writer.f32(layout.hanging_indent);
            writer.f32(layout.min_scale);
            writer.f32(layout.paragraph_spacing);
            writer.f32(layout.inset_left);
            writer.f32(layout.inset_right);
            writer.f32(layout.inset_top);
            writer.f32(layout.inset_bottom);
            writer.f32(layout.margin_left);
            writer.f32(layout.margin_right);
            writer.f32(layout.first_line_indent);
            writer.u32(layout.column_count);
            writer.f32(layout.column_spacing);
            writer.f32(layout.rotation_degrees);
            writer.f32(layout.font_scale);
            writer.f32(layout.line_spacing_reduction);
            writer.u8(layout.horizontal_overflow.code());
            writer.u8(layout.vertical_overflow.code());
            writer.u8(u8::from(layout.wrap));
            writer.u8(u8::from(layout.warp.is_some()));
            writer.u8(u8::from(layout.text_fill)
                | (u8::from(layout.text_scale_to_fit) << 1)
                | (u8::from(layout.low_resolution_supersample) << 2)
                | (u8::from(layout.text_matrix_scale_to_fit) << 3));
            writer.padding(3);
            writer.u32(layout.text_stroke_color);
            writer.f32(layout.text_stroke_width);
            writer.f32(layout.text_baseline);
            write_paint(writer, &layout.text_paint, depth)?;
            write_paint(writer, &layout.text_stroke_paint, depth)?;
            writer.u32(
                layout
                    .tab_stops
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for stop in &layout.tab_stops {
                writer.f32(stop.position);
                writer.u8(stop.align.code());
                writer.u8(stop.leader.code());
                writer.padding(2);
            }
            if let Some(prefix) = &layout.prefix {
                writer.string(prefix)?;
            }
            if let Some(warp) = &layout.warp {
                writer.string(warp)?;
            }
            writer.u32(
                layout
                    .paragraphs
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for paragraph in &layout.paragraphs {
                writer.u8(paragraph.align.code());
                writer.padding(3);
                writer.f32(paragraph.margin_left);
                writer.f32(paragraph.margin_right);
                writer.f32(paragraph.first_line_indent);
                writer.f32(paragraph.default_tab_stop);
                writer.f32(paragraph.line_height);
                writer.f32(paragraph.space_before);
                writer.f32(paragraph.space_after);
                writer.u8(u8::from(paragraph.latin_line_break)
                    | (u8::from(paragraph.hanging_punctuation) << 1)
                    | (u8::from(paragraph.rule_above.is_some()) << 2)
                    | (u8::from(paragraph.rule_below.is_some()) << 3)
                    | (u8::from(paragraph.drop_cap.is_some()) << 4));
                writer.padding(3);
                for rule in [paragraph.rule_above, paragraph.rule_below] {
                    let rule = rule.unwrap_or(crate::model::TextParagraphRule {
                        color: 0,
                        stroke_width: 0.0,
                        offset_x: 0.0,
                        offset_y: 0.0,
                        width: 0.0,
                    });
                    writer.u32(rule.color);
                    writer.f32(rule.stroke_width);
                    writer.f32(rule.offset_x);
                    writer.f32(rule.offset_y);
                    writer.f32(rule.width);
                    writer.padding(4);
                }
                let cap = paragraph.drop_cap.unwrap_or(crate::model::TextDropCap {
                    characters: 0,
                    lines: 0,
                    raised_lines: 0,
                    padding: 0.0,
                    outdent: 0.0,
                });
                writer.u32(cap.characters);
                writer.u32(cap.lines);
                writer.u32(cap.raised_lines);
                writer.f32(cap.padding);
                writer.f32(cap.outdent);
            }
            writer.u32(
                layout
                    .wrap_regions
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for rect in &layout.wrap_regions {
                writer.f32(rect.x);
                writer.f32(rect.y);
                writer.f32(rect.width);
                writer.f32(rect.height);
            }
            writer.u32(layout.fill_character.map_or(NONE, |(offset, _)| offset));
            writer.u32(
                layout
                    .fill_character
                    .map_or(0, |(_, character)| character as u32),
            );
            writer.u8(visual.code());
            writer.padding(3);
            write_visual_at_depth(writer, visual, depth + 1)
        }
        Visual::TextEffects { effects, visual } => {
            writer.u32(
                effects
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for effect in effects {
                writer.u8(u8::from(effect.shadow.is_some())
                    | (u8::from(effect.inner_shadow.is_some()) << 1)
                    | (u8::from(effect.wavy_underline) << 2)
                    | (u8::from(effect.dotted_underline) << 3)
                    | (u8::from(effect.heavy_underline) << 4)
                    | (u8::from(effect.double_underline) << 5)
                    | (u8::from(effect.dot_dash_underline) << 6)
                    | (u8::from(effect.double_strikethrough) << 7));
                writer.padding(3);
                writer.f32(effect.shadow_scale_x);
                writer.f32(effect.shadow_scale_y);
                writer.f32(effect.shadow_skew_x);
                writer.f32(effect.shadow_skew_y);
                writer.u8(effect.shadow_alignment);
                writer.u8(u8::from(effect.reflection.is_some())
                    | (u8::from(effect.glow.is_some()) << 1)
                    | (u8::from(effect.stroke.is_some()) << 2)
                    | (u8::from(effect.fill_to_text) << 3));
                writer.padding(2);
                for shadow in [effect.shadow, effect.inner_shadow].into_iter().flatten() {
                    writer.u32(shadow.color);
                    writer.f32(shadow.blur);
                    writer.f32(shadow.offset_x);
                    writer.f32(shadow.offset_y);
                }
                if let Some(reflection) = effect.reflection {
                    writer.f32(reflection.start_opacity);
                    writer.f32(reflection.end_opacity);
                    writer.f32(reflection.start_position);
                    writer.f32(reflection.end_position);
                    writer.f32(reflection.direction_degrees);
                    writer.f32(reflection.blur);
                    writer.f32(reflection.distance);
                    writer.f32(reflection.scale_x);
                    writer.f32(reflection.scale_y);
                }
                if let Some(glow) = effect.glow {
                    writer.u32(glow.color);
                    writer.f32(glow.radius);
                }
                if let Some(stroke) = &effect.stroke {
                    writer.f32(effect.stroke_width);
                    write_paint(writer, stroke, depth + 1)?;
                }
            }
            writer.u8(visual.code());
            writer.padding(3);
            write_visual_at_depth(writer, visual, depth + 1)
        }
        Visual::StrokeStyle { style, visual } => {
            writer.u8(style.cap.code());
            writer.u8(style.join.code());
            writer.u8(style.compound.code());
            writer.u8(style.alignment.code());
            writer.f32(style.miter_limit);
            writer.f32(style.dash_offset);
            writer.u32(
                style
                    .dash
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::TooManyFields)?,
            );
            for value in &style.dash {
                writer.f32(*value);
            }
            writer.u8(visual.code());
            writer.padding(3);
            write_visual_at_depth(writer, visual, depth + 1)
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
            writer.u8(u8::from(inner_shadow.is_some())
                | (u8::from(glow.is_some()) << 1)
                | (u8::from(reflection.is_some()) << 2)
                | (u8::from(soft_edge.is_some()) << 3)
                | (u8::from(three_d.is_some()) << 4)
                | (u8::from(outer_shadow.is_some()) << 5));
            writer.padding(3);
            if let Some(effect) = outer_shadow {
                writer.u32(effect.shadow.color);
                writer.f32(effect.shadow.blur);
                writer.f32(effect.shadow.offset_x);
                writer.f32(effect.shadow.offset_y);
                writer.f32(effect.scale_x);
                writer.f32(effect.scale_y);
                writer.f32(effect.skew_x);
                writer.f32(effect.skew_y);
                writer.u8(effect.alignment);
                writer.padding(3);
            }
            if let Some(shadow) = inner_shadow {
                writer.u32(shadow.color);
                writer.f32(shadow.blur);
                writer.f32(shadow.offset_x);
                writer.f32(shadow.offset_y);
            }
            if let Some(glow) = glow {
                writer.u32(glow.color);
                writer.f32(glow.radius);
            }
            if let Some(reflection) = reflection {
                writer.f32(reflection.start_opacity);
                writer.f32(reflection.end_opacity);
                writer.f32(reflection.start_position);
                writer.f32(reflection.end_position);
                writer.f32(reflection.direction_degrees);
                writer.f32(reflection.blur);
                writer.f32(reflection.distance);
                writer.f32(reflection.scale_x);
                writer.f32(reflection.scale_y);
            }
            if let Some(radius) = soft_edge {
                writer.f32(*radius);
            }
            if let Some(style) = three_d {
                writer.string(&style.camera_preset)?;
                writer.string(&style.light_rig)?;
                writer.string(&style.light_direction)?;
                writer.string(&style.material)?;
                for value in [
                    style.camera_fov,
                    style.camera_zoom,
                    style.camera_latitude,
                    style.camera_longitude,
                    style.camera_revolution,
                    style.light_latitude,
                    style.light_longitude,
                    style.light_revolution,
                    style.z,
                    style.extrusion_height,
                    style.contour_width,
                ] {
                    writer.f32(value);
                }
                writer.u8(u8::from(style.bevel_top.is_some())
                    | (u8::from(style.bevel_bottom.is_some()) << 1)
                    | (u8::from(style.extrusion_color.is_some()) << 2)
                    | (u8::from(style.contour_color.is_some()) << 3)
                    | (u8::from(style.backdrop.is_some()) << 4)
                    | (u8::from(style.flat_text_z.is_some()) << 5)
                    | (u8::from(style.applies_to_text) << 6));
                writer.padding(3);
                for bevel in [&style.bevel_top, &style.bevel_bottom]
                    .into_iter()
                    .flatten()
                {
                    writer.f32(bevel.width);
                    writer.f32(bevel.height);
                    writer.string(&bevel.preset)?;
                }
                if let Some(color) = style.extrusion_color {
                    writer.u32(color);
                }
                if let Some(color) = style.contour_color {
                    writer.u32(color);
                }
                if let Some(backdrop) = style.backdrop {
                    for value in [
                        backdrop.anchor_x,
                        backdrop.anchor_y,
                        backdrop.anchor_z,
                        backdrop.normal_x,
                        backdrop.normal_y,
                        backdrop.normal_z,
                        backdrop.up_x,
                        backdrop.up_y,
                        backdrop.up_z,
                    ] {
                        writer.f32(value);
                    }
                }
                if let Some(z) = style.flat_text_z {
                    writer.f32(z);
                }
            }
            writer.u8(visual.code());
            writer.padding(3);
            write_visual_at_depth(writer, visual, depth + 1)
        }
    }
}

fn add_length(total: &mut usize, value: usize) -> Result<(), ProtocolError> {
    *total = total
        .checked_add(value)
        .filter(|length| *length <= MAX_SNAPSHOT_BYTES)
        .ok_or(ProtocolError::SnapshotTooLarge)?;
    Ok(())
}

fn string_length(value: &str) -> Result<usize, ProtocolError> {
    let _: u32 = value
        .len()
        .try_into()
        .map_err(|_| ProtocolError::FieldTooLong)?;
    4_usize
        .checked_add(value.len())
        .ok_or(ProtocolError::SnapshotTooLarge)
}

fn optional_string_length(value: Option<&str>) -> Result<usize, ProtocolError> {
    value.map_or(Ok(4), string_length)
}

fn byte_array_length(value: &[u8]) -> Result<usize, ProtocolError> {
    let _: u32 = value
        .len()
        .try_into()
        .map_err(|_| ProtocolError::FieldTooLong)?;
    4_usize
        .checked_add(value.len())
        .ok_or(ProtocolError::SnapshotTooLarge)
}

fn xps_color_length(color: &XpsColor) -> Result<usize, ProtocolError> {
    match color {
        XpsColor::Rgba(_) => Ok(8),
        XpsColor::Context {
            profile, channels, ..
        } => {
            let _: u32 = channels
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?;
            let mut total = 12;
            add_length(&mut total, byte_array_length(profile)?)?;
            add_length(
                &mut total,
                channels
                    .len()
                    .checked_mul(4)
                    .ok_or(ProtocolError::SnapshotTooLarge)?,
            )?;
            Ok(total)
        }
    }
}

fn add_string_field(total: &mut usize, key: &str, value: &str) -> Result<(), ProtocolError> {
    add_length(total, string_length(key)?)?;
    add_length(total, 1)?;
    add_length(total, string_length(value)?)
}

fn add_unsigned_field(total: &mut usize, key: &str) -> Result<(), ProtocolError> {
    add_length(total, string_length(key)?)?;
    add_length(total, 5)
}

fn source_length(locator: &SourceLocator) -> Result<usize, ProtocolError> {
    let mut total = 0;
    match locator {
        SourceLocator::PptxShape {
            row,
            column,
            text_range,
            metadata,
            ..
        } => {
            add_length(&mut total, string_length("shape")?)?;
            add_length(&mut total, 1)?;
            add_unsigned_field(&mut total, "shapeId")?;
            if row.is_some() {
                add_unsigned_field(&mut total, "row")?;
            }
            if column.is_some() {
                add_unsigned_field(&mut total, "column")?;
            }
            if text_range.is_some() {
                add_unsigned_field(&mut total, "rangeStart")?;
                add_unsigned_field(&mut total, "rangeEnd")?;
            }
            if let Some(value) = metadata.name.as_deref() {
                add_string_field(&mut total, "name", value)?;
            }
            if let Some(value) = metadata.title.as_deref() {
                add_string_field(&mut total, "title", value)?;
            }
            if let Some(value) = metadata.description.as_deref() {
                add_string_field(&mut total, "description", value)?;
            }
            if metadata.hidden {
                add_unsigned_field(&mut total, "hidden")?;
            }
            if let Some(action) = metadata.click_action.as_ref() {
                add_action_fields(
                    &mut total,
                    action,
                    "clickKind",
                    "clickAction",
                    "clickTarget",
                    "clickTooltip",
                )?;
            }
            if let Some(action) = metadata.hover_action.as_ref() {
                add_action_fields(
                    &mut total,
                    action,
                    "hoverKind",
                    "hoverAction",
                    "hoverTarget",
                    "hoverTooltip",
                )?;
            }
        }
        SourceLocator::OdpElement {
            element_id,
            path,
            row,
            column,
        } => {
            add_length(&mut total, string_length("element")?)?;
            add_length(&mut total, 1)?;
            add_string_field(&mut total, "path", path)?;
            if let Some(element_id) = element_id {
                add_string_field(&mut total, "elementId", element_id)?;
            }
            if row.is_some() {
                add_unsigned_field(&mut total, "row")?;
            }
            if column.is_some() {
                add_unsigned_field(&mut total, "column")?;
            }
        }
        SourceLocator::Xlsx {
            kind,
            sheet_name,
            address,
            formula,
            drawing_id,
        } => {
            add_length(&mut total, string_length(kind)?)?;
            add_length(&mut total, 1)?;
            add_string_field(&mut total, "sheetName", sheet_name)?;
            if let Some(address) = address {
                add_string_field(&mut total, "address", address)?;
            }
            if let Some(formula) = formula {
                add_string_field(&mut total, "formula", formula)?;
            }
            if drawing_id.is_some() {
                add_unsigned_field(&mut total, "drawingId")?;
            }
        }
        SourceLocator::Ods {
            kind,
            table_name,
            row,
            column,
            element_id,
            path,
        } => {
            add_length(&mut total, string_length(kind)?)?;
            add_length(&mut total, 1)?;
            add_string_field(&mut total, "tableName", table_name)?;
            add_string_field(&mut total, "path", path)?;
            if row.is_some() {
                add_unsigned_field(&mut total, "row")?;
            }
            if column.is_some() {
                add_unsigned_field(&mut total, "column")?;
            }
            if let Some(element_id) = element_id {
                add_string_field(&mut total, "elementId", element_id)?;
            }
        }
        SourceLocator::Docx {
            kind,
            paragraph_id,
            paragraph_index,
            drawing_id,
            row,
            column,
            text_range,
            action,
        } => {
            add_length(&mut total, string_length(kind)?)?;
            add_length(&mut total, 1)?;
            if let Some(paragraph_id) = paragraph_id {
                add_string_field(&mut total, "paragraphId", paragraph_id)?;
            }
            if paragraph_index.is_some() {
                add_unsigned_field(&mut total, "paragraphIndex")?;
            }
            if drawing_id.is_some() {
                add_unsigned_field(&mut total, "drawingId")?;
            }
            if row.is_some() {
                add_unsigned_field(&mut total, "row")?;
            }
            if column.is_some() {
                add_unsigned_field(&mut total, "column")?;
            }
            if text_range.is_some() {
                add_unsigned_field(&mut total, "rangeStart")?;
                add_unsigned_field(&mut total, "rangeEnd")?;
            }
            if let Some(action) = action {
                add_action_fields(
                    &mut total,
                    action,
                    "clickKind",
                    "clickAction",
                    "clickTarget",
                    "clickTooltip",
                )?;
            }
        }
        SourceLocator::Odt {
            kind,
            element_id,
            path,
            row,
            column,
            text_range,
        } => {
            add_length(&mut total, string_length(kind)?)?;
            add_length(&mut total, 1)?;
            add_string_field(&mut total, "path", path)?;
            if let Some(element_id) = element_id {
                add_string_field(&mut total, "elementId", element_id)?;
            }
            if row.is_some() {
                add_unsigned_field(&mut total, "row")?;
            }
            if column.is_some() {
                add_unsigned_field(&mut total, "column")?;
            }
            if text_range.is_some() {
                add_unsigned_field(&mut total, "rangeStart")?;
                add_unsigned_field(&mut total, "rangeEnd")?;
            }
        }
        SourceLocator::Flat {
            kind,
            index,
            row,
            column,
            text_range,
        } => {
            add_length(&mut total, string_length(kind)?)?;
            add_length(&mut total, 1)?;
            if index.is_some() {
                add_unsigned_field(&mut total, "index")?;
            }
            if row.is_some() {
                add_unsigned_field(&mut total, "row")?;
            }
            if column.is_some() {
                add_unsigned_field(&mut total, "column")?;
            }
            if text_range.is_some() {
                add_unsigned_field(&mut total, "rangeStart")?;
                add_unsigned_field(&mut total, "rangeEnd")?;
            }
        }
        SourceLocator::Legacy {
            kind,
            stream,
            record_offset,
            row,
            column,
            text_range,
        } => {
            add_length(&mut total, string_length(kind)?)?;
            add_length(&mut total, 1)?;
            add_string_field(&mut total, "stream", stream)?;
            if record_offset.is_some() {
                add_unsigned_field(&mut total, "recordOffset")?;
            }
            if row.is_some() {
                add_unsigned_field(&mut total, "row")?;
            }
            if column.is_some() {
                add_unsigned_field(&mut total, "column")?;
            }
            if text_range.is_some() {
                add_unsigned_field(&mut total, "rangeStart")?;
                add_unsigned_field(&mut total, "rangeEnd")?;
            }
        }
        SourceLocator::Iwork { kind, component } => {
            add_length(&mut total, string_length(kind)?)?;
            add_length(&mut total, 1)?;
            add_string_field(&mut total, "component", component)?;
        }
        SourceLocator::Pdf {
            kind,
            object_number,
            byte_offset,
            action,
        } => {
            add_length(&mut total, string_length(kind)?)?;
            add_length(&mut total, 1)?;
            if object_number.is_some() {
                add_unsigned_field(&mut total, "objectNumber")?;
            }
            if byte_offset.is_some() {
                add_unsigned_field(&mut total, "byteOffset")?;
            }
            if let Some(action) = action {
                add_action_fields(
                    &mut total,
                    action,
                    "clickKind",
                    "clickAction",
                    "clickTarget",
                    "clickTooltip",
                )?;
            }
        }
        SourceLocator::Xps { kind, path } => {
            add_length(&mut total, string_length(kind)?)?;
            add_length(&mut total, 1)?;
            add_string_field(&mut total, "path", path)?;
        }
        SourceLocator::Ofd { kind, path } => {
            add_length(&mut total, string_length(kind)?)?;
            add_length(&mut total, 1)?;
            add_string_field(&mut total, "path", path)?;
        }
    }
    Ok(total)
}

fn path_commands_length(commands: &[PathCommand]) -> Result<usize, ProtocolError> {
    let _: u32 = commands
        .len()
        .try_into()
        .map_err(|_| ProtocolError::TooManyFields)?;
    let mut total = 4;
    for command in commands {
        add_length(
            &mut total,
            match command {
                PathCommand::MoveTo { .. } | PathCommand::LineTo { .. } => 12,
                PathCommand::QuadraticCurveTo { .. } => 20,
                PathCommand::BezierCurveTo { .. } => 28,
                PathCommand::ClosePath => 4,
            },
        )?;
    }
    Ok(total)
}

fn geometry_length(geometry: &Geometry) -> Result<usize, ProtocolError> {
    match geometry {
        Geometry::Rectangle | Geometry::Ellipse | Geometry::Line => Ok(1),
        Geometry::RoundedRectangle { .. } => Ok(9),
        Geometry::Path { commands, .. } => {
            let mut total = 5;
            add_length(&mut total, path_commands_length(commands)?)?;
            Ok(total)
        }
        Geometry::LayeredPath { layers } => {
            let _: u32 = layers
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?;
            let mut total = 5;
            for layer in layers {
                add_length(&mut total, 4)?;
                add_length(&mut total, path_commands_length(&layer.commands)?)?;
            }
            Ok(total)
        }
    }
}

fn paint_length(paint: &Paint, depth: usize) -> Result<usize, ProtocolError> {
    if depth > MAX_VISUAL_DEPTH {
        return Err(ProtocolError::TooManyFields);
    }
    match paint {
        Paint::None => Ok(4),
        Paint::Solid(_) => Ok(8),
        Paint::LinearGradient { stops, .. } => {
            let _: u32 = stops
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?;
            stops
                .len()
                .checked_mul(8)
                .and_then(|length| length.checked_add(24))
                .filter(|length| *length <= MAX_SNAPSHOT_BYTES)
                .ok_or(ProtocolError::SnapshotTooLarge)
        }
        Paint::RadialGradient { stops, .. } => {
            let _: u32 = stops
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?;
            stops
                .len()
                .checked_mul(8)
                .and_then(|length| length.checked_add(32))
                .filter(|length| *length <= MAX_SNAPSHOT_BYTES)
                .ok_or(ProtocolError::SnapshotTooLarge)
        }
        Paint::RectGradient { stops, .. } => {
            let _: u32 = stops
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?;
            stops
                .len()
                .checked_mul(8)
                .and_then(|length| length.checked_add(16))
                .filter(|length| *length <= MAX_SNAPSHOT_BYTES)
                .ok_or(ProtocolError::SnapshotTooLarge)
        }
        Paint::CircleGradient { stops, .. } => {
            let _: u32 = stops
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?;
            stops
                .len()
                .checked_mul(8)
                .and_then(|length| length.checked_add(32))
                .filter(|length| *length <= MAX_SNAPSHOT_BYTES)
                .ok_or(ProtocolError::SnapshotTooLarge)
        }
        Paint::MappedGradient { paint, .. } => paint_length(paint, depth + 1)?
            .checked_add(24)
            .ok_or(ProtocolError::TooManyFields),
        Paint::ShapeGradient { stops, .. } => stops
            .len()
            .checked_mul(8)
            .and_then(|n| n.checked_add(24))
            .filter(|n| *n <= MAX_SNAPSHOT_BYTES)
            .ok_or(ProtocolError::SnapshotTooLarge),
        Paint::Pattern { preset, .. } => string_length(preset)?
            .checked_add(12)
            .filter(|length| *length <= MAX_SNAPSHOT_BYTES)
            .ok_or(ProtocolError::SnapshotTooLarge),
        Paint::Image {
            media_type,
            bytes,
            mapping,
            ..
        } => {
            let mut total = if mapping.is_some() { 80 } else { 32 };
            add_length(&mut total, string_length(media_type)?)?;
            add_length(&mut total, byte_array_length(bytes)?)?;
            Ok(total)
        }
        Paint::Visual { children, .. } => {
            let _: u32 = children
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?;
            let mut total = 104_usize;
            for child in children {
                add_length(&mut total, 20)?;
                add_length(
                    &mut total,
                    visual_length_at_depth(&child.visual, depth + 1)?,
                )?;
            }
            Ok(total)
        }
        Paint::XpsGradient { stops, .. } => {
            let _: u32 = stops
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?;
            let mut total = 84;
            for stop in stops {
                add_length(&mut total, 4)?;
                add_length(&mut total, xps_color_length(&stop.color)?)?;
            }
            Ok(total)
        }
    }
}

fn visual_length(visual: &Visual) -> Result<usize, ProtocolError> {
    validate_visual(visual, 0)?;
    visual_length_at_depth(visual, 0)
}

fn visual_length_at_depth(visual: &Visual, depth: usize) -> Result<usize, ProtocolError> {
    if depth > MAX_VISUAL_DEPTH {
        return Err(ProtocolError::TooManyFields);
    }
    match visual {
        Visual::None => Ok(0),
        Visual::Shape { geometry, .. } => geometry_length(geometry)?
            .checked_add(12)
            .ok_or(ProtocolError::SnapshotTooLarge),
        Visual::Text {
            geometry,
            font_family,
            ..
        } => {
            let family_length = string_length(font_family)?;
            geometry_length(geometry)?
                .checked_add(24)
                .and_then(|length| length.checked_add(family_length))
                .ok_or(ProtocolError::SnapshotTooLarge)
        }
        Visual::Image {
            media_type, bytes, ..
        } => {
            let media_type_length = string_length(media_type)?;
            let bytes_length = byte_array_length(bytes)?;
            16_usize
                .checked_add(media_type_length)
                .and_then(|length| length.checked_add(bytes_length))
                .ok_or(ProtocolError::SnapshotTooLarge)
        }
        Visual::ImageWithFallback {
            media_type,
            bytes,
            fallback_media_type,
            fallback_bytes,
            ..
        } => {
            let media_type_length = string_length(media_type)?;
            let bytes_length = byte_array_length(bytes)?;
            let fallback_media_type_length = string_length(fallback_media_type)?;
            let fallback_bytes_length = byte_array_length(fallback_bytes)?;
            16_usize
                .checked_add(media_type_length)
                .and_then(|length| length.checked_add(bytes_length))
                .and_then(|length| length.checked_add(fallback_media_type_length))
                .and_then(|length| length.checked_add(fallback_bytes_length))
                .ok_or(ProtocolError::SnapshotTooLarge)
        }
        Visual::MaskedImage {
            media_type,
            bytes,
            mask_media_type,
            alpha_mask,
            ..
        } => {
            let media_type_length = string_length(media_type)?;
            let bytes_length = byte_array_length(bytes)?;
            let mask_media_type_length = string_length(mask_media_type)?;
            let mask_length = byte_array_length(alpha_mask)?;
            32_usize
                .checked_add(media_type_length)
                .and_then(|length| length.checked_add(bytes_length))
                .and_then(|length| length.checked_add(mask_media_type_length))
                .and_then(|length| length.checked_add(mask_length))
                .ok_or(ProtocolError::SnapshotTooLarge)
        }
        Visual::Group { children } => {
            let mut total = 4_usize;
            for child in children {
                add_length(&mut total, 20)?;
                add_length(
                    &mut total,
                    visual_length_at_depth(&child.visual, depth + 1)?,
                )?;
            }
            Ok(total)
        }
        Visual::OpacityMask { mask, visual } => {
            let mut total = paint_length(mask, depth)?;
            add_length(&mut total, 4)?;
            add_length(&mut total, visual_length_at_depth(visual, depth + 1)?)?;
            Ok(total)
        }
        Visual::ColorManagedImage {
            source_profile,
            destination_profile,
            visual,
        } => {
            let mut total = byte_array_length(source_profile)?;
            add_length(&mut total, 8)?;
            if let Some(profile) = destination_profile {
                add_length(&mut total, byte_array_length(profile)?)?;
            }
            add_length(&mut total, visual_length_at_depth(visual, depth + 1)?)?;
            Ok(total)
        }
        Visual::Media {
            media_type,
            bytes,
            poster,
            ..
        } => {
            let mut total = 8_usize;
            add_length(&mut total, string_length(media_type)?)?;
            add_length(&mut total, byte_array_length(bytes)?)?;
            add_length(&mut total, visual_length_at_depth(poster, depth + 1)?)?;
            Ok(total)
        }
        Visual::Layer { visual, .. } => visual_length_at_depth(visual, depth + 1)?
            .checked_add(32)
            .ok_or(ProtocolError::SnapshotTooLarge),
        Visual::PaintedShape {
            geometry,
            fill,
            stroke,
            ..
        } => {
            let stroke_length = paint_length(stroke, depth)?;
            geometry_length(geometry)?
                .checked_add(paint_length(fill, depth)?)
                .and_then(|length| length.checked_add(stroke_length))
                .and_then(|length| length.checked_add(4))
                .ok_or(ProtocolError::SnapshotTooLarge)
        }
        Visual::RichText {
            geometry,
            fill,
            stroke,
            runs,
            ..
        } => {
            let _: u32 = runs
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?;
            let mut total = geometry_length(geometry)?;
            add_length(&mut total, paint_length(fill, depth)?)?;
            add_length(&mut total, paint_length(stroke, depth)?)?;
            add_length(&mut total, 16)?;
            for run in runs {
                add_length(&mut total, string_length(&run.text)?)?;
                add_length(&mut total, string_length(&run.font_family)?)?;
                add_length(&mut total, 28)?;
                add_length(
                    &mut total,
                    paint_length(run.paint.as_deref().unwrap_or(&Paint::None), depth + 1)?,
                )?;
            }
            Ok(total)
        }
        Visual::Effect {
            shadow,
            clip,
            visual,
        } => {
            let mut total = 8;
            if shadow.is_some() {
                add_length(&mut total, 16)?;
            }
            if let Some(clip) = clip {
                add_length(&mut total, geometry_length(clip)?)?;
            }
            add_length(&mut total, visual_length_at_depth(visual, depth + 1)?)?;
            Ok(total)
        }
        Visual::ImageColorChange { visual, .. } => visual_length_at_depth(visual, depth + 1)?
            .checked_add(12)
            .ok_or(ProtocolError::SnapshotTooLarge),
        Visual::ImageAdjustment { visual, .. } => visual_length_at_depth(visual, depth + 1)?
            .checked_add(24)
            .ok_or(ProtocolError::SnapshotTooLarge),
        Visual::TextLayout { layout, visual } => {
            let _: u32 = layout
                .tab_stops
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?;
            let _: u32 = layout
                .paragraphs
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?;
            let mut total = 112;
            add_length(
                &mut total,
                layout
                    .wrap_regions
                    .len()
                    .checked_mul(16)
                    .ok_or(ProtocolError::SnapshotTooLarge)?,
            )?;
            add_length(&mut total, paint_length(&layout.text_paint, depth)?)?;
            add_length(&mut total, paint_length(&layout.text_stroke_paint, depth)?)?;
            add_length(
                &mut total,
                layout
                    .tab_stops
                    .len()
                    .checked_mul(8)
                    .ok_or(ProtocolError::SnapshotTooLarge)?,
            )?;
            if let Some(prefix) = &layout.prefix {
                add_length(&mut total, string_length(prefix)?)?;
            }
            if let Some(warp) = &layout.warp {
                add_length(&mut total, string_length(warp)?)?;
            }
            add_length(
                &mut total,
                layout
                    .paragraphs
                    .len()
                    .checked_mul(104)
                    .ok_or(ProtocolError::SnapshotTooLarge)?,
            )?;
            add_length(&mut total, visual_length_at_depth(visual, depth + 1)?)?;
            Ok(total)
        }
        Visual::TextEffects { effects, visual } => {
            let mut total = 8_usize;
            add_length(
                &mut total,
                effects
                    .len()
                    .checked_mul(24)
                    .ok_or(ProtocolError::SnapshotTooLarge)?,
            )?;
            for effect in effects {
                for shadow in [effect.shadow, effect.inner_shadow].into_iter().flatten() {
                    let _ = shadow;
                    add_length(&mut total, 16)?;
                }
                if effect.reflection.is_some() {
                    add_length(&mut total, 36)?;
                }
                if effect.glow.is_some() {
                    add_length(&mut total, 8)?;
                }
                if let Some(stroke) = &effect.stroke {
                    add_length(&mut total, 4)?;
                    add_length(&mut total, paint_length(stroke, depth + 1)?)?;
                }
            }
            add_length(&mut total, visual_length_at_depth(visual, depth + 1)?)?;
            Ok(total)
        }
        Visual::StrokeStyle { style, visual } => {
            let mut total = 20;
            add_length(
                &mut total,
                style
                    .dash
                    .len()
                    .checked_mul(4)
                    .ok_or(ProtocolError::SnapshotTooLarge)?,
            )?;
            add_length(&mut total, visual_length_at_depth(visual, depth + 1)?)?;
            Ok(total)
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
            let mut total = 8;
            if outer_shadow.is_some() {
                add_length(&mut total, 36)?;
            }
            if inner_shadow.is_some() {
                add_length(&mut total, 16)?;
            }
            if glow.is_some() {
                add_length(&mut total, 8)?;
            }
            if reflection.is_some() {
                add_length(&mut total, 36)?;
            }
            if soft_edge.is_some() {
                add_length(&mut total, 4)?;
            }
            if let Some(style) = three_d {
                add_length(&mut total, 48)?;
                for value in [
                    &style.camera_preset,
                    &style.light_rig,
                    &style.light_direction,
                    &style.material,
                ] {
                    add_length(&mut total, string_length(value)?)?;
                }
                for bevel in [&style.bevel_top, &style.bevel_bottom]
                    .into_iter()
                    .flatten()
                {
                    add_length(&mut total, 8)?;
                    add_length(&mut total, string_length(&bevel.preset)?)?;
                }
                if style.extrusion_color.is_some() {
                    add_length(&mut total, 4)?;
                }
                if style.contour_color.is_some() {
                    add_length(&mut total, 4)?;
                }
                if style.backdrop.is_some() {
                    add_length(&mut total, 36)?;
                }
                if style.flat_text_z.is_some() {
                    add_length(&mut total, 4)?;
                }
            }
            add_length(&mut total, visual_length_at_depth(visual, depth + 1)?)?;
            Ok(total)
        }
    }
}

fn encoded_length(document: &Document) -> Result<usize, ProtocolError> {
    let mut total = 12;
    let _: u32 = document
        .diagnostics
        .len()
        .try_into()
        .map_err(|_| ProtocolError::SnapshotTooLarge)?;
    add_length(&mut total, 4)?;
    for diagnostic in &document.diagnostics {
        add_length(&mut total, string_length(code_name(diagnostic.code))?)?;
        add_length(&mut total, 4)?;
        add_length(&mut total, string_length(&diagnostic.message)?)?;
        add_length(
            &mut total,
            optional_string_length(diagnostic.location.part.as_deref())?,
        )?;
        add_length(
            &mut total,
            optional_string_length(
                diagnostic
                    .details
                    .iter()
                    .find_map(|(key, value)| (key == "objectId").then_some(value.as_str())),
            )?,
        )?;
        add_length(&mut total, 1)?;
        for (key, value) in &diagnostic.details {
            add_string_field(&mut total, key, value)?;
        }
        if diagnostic
            .location
            .byte_offset
            .and_then(|value| u32::try_from(value).ok())
            .is_some()
        {
            add_unsigned_field(&mut total, "byteOffset")?;
        }
    }

    let _: u32 = document
        .units
        .len()
        .try_into()
        .map_err(|_| ProtocolError::SnapshotTooLarge)?;
    add_length(&mut total, 4)?;
    for unit in &document.units {
        if ![
            unit.width,
            unit.height,
            unit.frozen_width,
            unit.frozen_height,
        ]
        .into_iter()
        .all(f32::is_finite)
            || unit.width <= 0.0
            || unit.height <= 0.0
            || unit.frozen_rows > unit.rows
            || unit.frozen_columns > unit.columns
            || !(0.0..=unit.width).contains(&unit.frozen_width)
            || !(0.0..=unit.height).contains(&unit.frozen_height)
        {
            return Err(ProtocolError::InvalidValue);
        }
        validate_sheet_axis(&unit.row_axis, unit.rows)?;
        validate_sheet_axis(&unit.column_axis, unit.columns)?;
        add_length(&mut total, 52)?;
        add_length(&mut total, 16)?;
        add_length(
            &mut total,
            unit.row_axis
                .spans
                .len()
                .checked_mul(12)
                .ok_or(ProtocolError::SnapshotTooLarge)?,
        )?;
        add_length(
            &mut total,
            unit.column_axis
                .spans
                .len()
                .checked_mul(12)
                .ok_or(ProtocolError::SnapshotTooLarge)?,
        )?;
        add_length(&mut total, string_length(&unit.id)?)?;
        add_length(&mut total, string_length(&unit.name)?)?;
        add_length(
            &mut total,
            optional_string_length(
                unit.slide
                    .as_ref()
                    .and_then(|slide| slide.source_id.as_deref()),
            )?,
        )?;
        add_length(
            &mut total,
            optional_string_length(
                unit.slide
                    .as_ref()
                    .and_then(|slide| slide.source_part.as_deref()),
            )?,
        )?;
        add_length(
            &mut total,
            optional_string_length(
                unit.slide
                    .as_ref()
                    .and_then(|slide| slide.speaker_notes.as_deref()),
            )?,
        )?;
        add_length(
            &mut total,
            optional_string_length(
                unit.slide
                    .as_ref()
                    .and_then(|slide| slide.speaker_notes_part.as_deref()),
            )?,
        )?;
        let note_paragraphs = unit
            .slide
            .as_ref()
            .map_or(&[][..], |slide| slide.speaker_note_paragraphs.as_slice());
        if note_paragraphs.len() > MAX_TEXT_PARAGRAPHS {
            return Err(ProtocolError::TooManyFields);
        }
        let _: u32 = note_paragraphs
            .len()
            .try_into()
            .map_err(|_| ProtocolError::TooManyFields)?;
        add_length(&mut total, 4)?;
        for paragraph in note_paragraphs {
            add_length(&mut total, speaker_note_paragraph_length(paragraph)?)?;
        }
        add_length(&mut total, 8)?;
        if unit.sheet.is_some() && unit.kind != crate::model::UnitKind::Sheet {
            return Err(ProtocolError::InvalidValue);
        }
        if let Some(sheet) = &unit.sheet
            && ![
                sheet.margins.left,
                sheet.margins.right,
                sheet.margins.top,
                sheet.margins.bottom,
                sheet.margins.header,
                sheet.margins.footer,
            ]
            .into_iter()
            .flatten()
            .all(|value| value.is_finite() && value >= 0.0)
        {
            return Err(ProtocolError::InvalidValue);
        }
        add_length(&mut total, 44)?;
        for value in [
            unit.sheet
                .as_ref()
                .and_then(|sheet| sheet.print_area.as_deref()),
            unit.sheet
                .as_ref()
                .and_then(|sheet| sheet.odd_header.as_deref()),
            unit.sheet
                .as_ref()
                .and_then(|sheet| sheet.odd_footer.as_deref()),
            unit.sheet
                .as_ref()
                .and_then(|sheet| sheet.even_header.as_deref()),
            unit.sheet
                .as_ref()
                .and_then(|sheet| sheet.even_footer.as_deref()),
            unit.sheet
                .as_ref()
                .and_then(|sheet| sheet.first_header.as_deref()),
            unit.sheet
                .as_ref()
                .and_then(|sheet| sheet.first_footer.as_deref()),
            unit.sheet
                .as_ref()
                .and_then(|sheet| sheet.print_titles.as_deref()),
        ] {
            add_length(&mut total, optional_string_length(value)?)?;
        }
        add_length(&mut total, 8)?;
        if let Some(sheet) = &unit.sheet {
            add_length(
                &mut total,
                (sheet.row_breaks.len() + sheet.column_breaks.len())
                    .checked_mul(12)
                    .ok_or(ProtocolError::SnapshotTooLarge)?,
            )?;
        }
    }

    let _: u32 = document
        .outline
        .len()
        .try_into()
        .map_err(|_| ProtocolError::SnapshotTooLarge)?;
    add_length(&mut total, 4)?;
    for item in &document.outline {
        if item.title.is_empty()
            || usize::try_from(item.unit_index)
                .ok()
                .is_none_or(|index| index >= document.units.len())
        {
            return Err(ProtocolError::InvalidValue);
        }
        add_length(&mut total, string_length(&item.title)?)?;
        add_length(&mut total, 8)?;
    }

    let _: u32 = document
        .embedded_fonts
        .len()
        .try_into()
        .map_err(|_| ProtocolError::SnapshotTooLarge)?;
    add_length(&mut total, 4)?;
    for font in &document.embedded_fonts {
        if font.family.is_empty()
            || font.family.len() > 256
            || font.bytes.is_empty()
            || !(1..=1000).contains(&font.weight)
        {
            return Err(ProtocolError::InvalidValue);
        }
        add_length(&mut total, 8)?;
        add_length(&mut total, string_length(&font.family)?)?;
        add_length(&mut total, byte_array_length(&font.bytes)?)?;
    }

    add_length(&mut total, 4)?;
    if document.font_alternate_names.len() > 1024 {
        return Err(ProtocolError::TooManyFields);
    }
    for (family, names) in &document.font_alternate_names {
        if family.is_empty() || family.len() > 128 || names.is_empty() || names.len() > 64 {
            return Err(ProtocolError::InvalidValue);
        }
        add_length(&mut total, string_length(family)?)?;
        add_length(&mut total, 4)?;
        for name in names {
            if name.is_empty() || name.len() > 128 {
                return Err(ProtocolError::InvalidValue);
            }
            add_length(&mut total, string_length(name)?)?;
        }
    }

    let _: u32 = document
        .objects
        .len()
        .try_into()
        .map_err(|_| ProtocolError::SnapshotTooLarge)?;
    add_length(&mut total, 4)?;
    for object in &document.objects {
        add_length(&mut total, 36)?;
        add_length(&mut total, string_length(&object.stable_id)?)?;
        add_length(
            &mut total,
            optional_string_length(object.parent_stable_id.as_deref())?,
        )?;
        add_length(&mut total, optional_string_length(object.text.as_deref())?)?;
        add_length(&mut total, string_length(&object.source.part)?)?;
        add_length(&mut total, source_length(&object.source.locator)?)?;
        add_length(&mut total, visual_length(&object.visual)?)?;
    }
    Ok(total)
}

pub fn encode(document: &Document) -> Result<Vec<u8>, ProtocolError> {
    let expected_length = encoded_length(document)?;
    let mut writer = Writer::new(expected_length)?;
    writer.u32(MAGIC);
    writer.u16(VERSION);
    writer.u8(u8::from(document.fatal));
    writer.u8(document.format.map_or(0, |format| format.code()));
    writer.u8(document.kind.map_or(0, |kind| kind.code()));
    writer.padding(3);

    writer.u32(
        document
            .diagnostics
            .len()
            .try_into()
            .map_err(|_| ProtocolError::SnapshotTooLarge)?,
    );
    for diagnostic in &document.diagnostics {
        write_diagnostic(&mut writer, diagnostic)?;
    }

    writer.u32(
        document
            .units
            .len()
            .try_into()
            .map_err(|_| ProtocolError::SnapshotTooLarge)?,
    );
    for unit in &document.units {
        writer.u8(unit.kind.code());
        writer.padding(3);
        writer.u32(unit.index);
        writer.string(&unit.id)?;
        writer.string(&unit.name)?;
        writer.f32(unit.width);
        writer.f32(unit.height);
        writer.u32(unit.rows);
        writer.u32(unit.columns);
        writer.u32(unit.frozen_rows);
        writer.u32(unit.frozen_columns);
        writer.f32(unit.frozen_width);
        writer.f32(unit.frozen_height);
        write_sheet_axis(&mut writer, &unit.row_axis)?;
        write_sheet_axis(&mut writer, &unit.column_axis)?;
        writer.u8(u8::from(unit.show_grid_lines));
        writer.padding(3);
        writer.u8(u8::from(unit.tab_color.is_some()));
        writer.padding(3);
        writer.u32(unit.tab_color.unwrap_or(0));
        writer.optional_string(
            unit.slide
                .as_ref()
                .and_then(|slide| slide.source_id.as_deref()),
        )?;
        writer.optional_string(
            unit.slide
                .as_ref()
                .and_then(|slide| slide.source_part.as_deref()),
        )?;
        writer.optional_string(
            unit.slide
                .as_ref()
                .and_then(|slide| slide.speaker_notes.as_deref()),
        )?;
        writer.optional_string(
            unit.slide
                .as_ref()
                .and_then(|slide| slide.speaker_notes_part.as_deref()),
        )?;
        let note_paragraphs = unit
            .slide
            .as_ref()
            .map_or(&[][..], |slide| slide.speaker_note_paragraphs.as_slice());
        writer.u32(
            note_paragraphs
                .len()
                .try_into()
                .map_err(|_| ProtocolError::TooManyFields)?,
        );
        for paragraph in note_paragraphs {
            write_speaker_note_paragraph(&mut writer, paragraph)?;
        }
        writer.u32(unit.slide.as_ref().map_or(0, |slide| slide.number));
        writer.u8(u8::from(
            unit.slide.as_ref().is_some_and(|slide| slide.hidden),
        ));
        writer.padding(3);
        let sheet = unit.sheet.as_ref();
        writer.u8(u8::from(sheet.is_some()));
        writer.u8(sheet
            .and_then(|sheet| sheet.orientation)
            .map_or(0, |orientation| orientation.code()));
        writer.u8(sheet.map_or(0, |sheet| {
            u8::from(sheet.fit_to_page)
                | (u8::from(sheet.different_odd_even) << 1)
                | (u8::from(sheet.different_first) << 2)
                | match sheet.view_mode {
                    Some(crate::model::SheetViewMode::PageLayout) => 1 << 3,
                    Some(crate::model::SheetViewMode::PageBreakPreview) => 1 << 4,
                    None => 0,
                }
        }));
        writer.padding(1);
        writer.u32(sheet.and_then(|sheet| sheet.paper_size).unwrap_or(NONE));
        writer.u32(sheet.and_then(|sheet| sheet.scale).unwrap_or(NONE));
        writer.u32(sheet.and_then(|sheet| sheet.fit_to_width).unwrap_or(NONE));
        writer.u32(sheet.and_then(|sheet| sheet.fit_to_height).unwrap_or(NONE));
        for value in [
            sheet.and_then(|sheet| sheet.margins.left),
            sheet.and_then(|sheet| sheet.margins.right),
            sheet.and_then(|sheet| sheet.margins.top),
            sheet.and_then(|sheet| sheet.margins.bottom),
            sheet.and_then(|sheet| sheet.margins.header),
            sheet.and_then(|sheet| sheet.margins.footer),
        ] {
            writer.f32(value.unwrap_or(f32::NAN));
        }
        writer.optional_string(sheet.and_then(|sheet| sheet.print_area.as_deref()))?;
        writer.optional_string(sheet.and_then(|sheet| sheet.odd_header.as_deref()))?;
        writer.optional_string(sheet.and_then(|sheet| sheet.odd_footer.as_deref()))?;
        writer.optional_string(sheet.and_then(|sheet| sheet.even_header.as_deref()))?;
        writer.optional_string(sheet.and_then(|sheet| sheet.even_footer.as_deref()))?;
        writer.optional_string(sheet.and_then(|sheet| sheet.first_header.as_deref()))?;
        writer.optional_string(sheet.and_then(|sheet| sheet.first_footer.as_deref()))?;
        writer.optional_string(sheet.and_then(|sheet| sheet.print_titles.as_deref()))?;
        for breaks in [
            sheet
                .map(|sheet| sheet.row_breaks.as_slice())
                .unwrap_or(&[]),
            sheet
                .map(|sheet| sheet.column_breaks.as_slice())
                .unwrap_or(&[]),
        ] {
            writer.u32(
                breaks
                    .len()
                    .try_into()
                    .map_err(|_| ProtocolError::SnapshotTooLarge)?,
            );
            for boundary in breaks {
                for value in boundary {
                    writer.u32(*value);
                }
            }
        }
    }

    writer.u32(
        document
            .outline
            .len()
            .try_into()
            .map_err(|_| ProtocolError::SnapshotTooLarge)?,
    );
    for item in &document.outline {
        writer.string(&item.title)?;
        writer.u32(item.unit_index);
        writer.u32(item.level);
    }

    writer.u32(
        document
            .embedded_fonts
            .len()
            .try_into()
            .map_err(|_| ProtocolError::SnapshotTooLarge)?,
    );
    for font in &document.embedded_fonts {
        writer.u8(font.style.code());
        writer.padding(3);
        writer.u32(u32::from(font.weight));
        writer.string(&font.family)?;
        writer.byte_array(&font.bytes)?;
    }

    writer.u32(document.font_alternate_names.len() as u32);
    for (family, names) in &document.font_alternate_names {
        writer.string(family)?;
        writer.u32(names.len() as u32);
        for name in names {
            writer.string(name)?;
        }
    }

    writer.u32(
        document
            .objects
            .len()
            .try_into()
            .map_err(|_| ProtocolError::SnapshotTooLarge)?,
    );
    for object in &document.objects {
        writer.u32(object.numeric_id);
        writer.i32(
            object
                .parent_numeric_id
                .map_or(-1, |value| i32::try_from(value).unwrap_or(i32::MAX)),
        );
        writer.u32(object.unit_index);
        writer.u8(object.kind.code());
        writer.u8(object.visual.code());
        writer.u8(object.source.mapping.code());
        writer.padding(1);
        writer.i32(object.z);
        writer.f32(object.bounds.x);
        writer.f32(object.bounds.y);
        writer.f32(object.bounds.width);
        writer.f32(object.bounds.height);
        writer.string(&object.stable_id)?;
        writer.optional_string(object.parent_stable_id.as_deref())?;
        writer.optional_string(object.text.as_deref())?;
        writer.string(&object.source.part)?;
        write_source(&mut writer, &object.source.locator)?;
        write_visual(&mut writer, &object.visual)?;
    }

    if writer.overflowed || writer.bytes.len() != expected_length {
        return Err(ProtocolError::SnapshotTooLarge);
    }
    Ok(writer.bytes)
}

#[cfg(test)]
mod tests {
    use super::{Writer, visual_length, write_visual};
    use crate::model::{
        AffineTransform, FillRule, Geometry, GradientSpread, GradientStop, ImageCrop, MediaKind,
        Paint, PathCommand, Rect, Shadow, StretchMode, TextAlign, TextAutoFit, TextDirection,
        TextLayout, TextOrientation, TextRun, TextVerticalAlign, TileMode, Visual,
        VisualBrushChild, XpsColor, XpsGradientStop,
    };

    #[test]
    fn masked_image_payload_has_an_exact_validated_length() {
        let visual = Visual::MaskedImage {
            media_type: "image/jpeg".to_owned(),
            bytes: vec![0xff, 0xd8, 0xff, 0xd9],
            resource_id: 5,
            mask_resource_id: 6,
            mask_width: 2,
            mask_height: 1,
            mask_media_type: "image/bmp".to_owned(),
            alpha_mask: vec![64, 192],
            crop: ImageCrop::default(),
        };
        let expected = visual_length(&visual).expect("masked image length");
        let mut writer = Writer::new(expected).expect("writer allocation");
        write_visual(&mut writer, &visual).expect("masked image payload");
        assert!(!writer.overflowed);
        assert_eq!(writer.bytes.len(), expected);

        let invalid = Visual::MaskedImage {
            media_type: "image/jpeg".to_owned(),
            bytes: vec![0xff, 0xd8, 0xff, 0xd9],
            resource_id: 5,
            mask_resource_id: 0,
            mask_width: 0,
            mask_height: 1,
            mask_media_type: "image/bmp".to_owned(),
            alpha_mask: vec![255],
            crop: ImageCrop::default(),
        };
        assert!(visual_length(&invalid).is_err());
    }

    #[test]
    fn opacity_masked_group_payload_has_an_exact_recursive_length() {
        let visual = Visual::OpacityMask {
            mask: Paint::LinearGradient {
                x0: 0.0,
                y0: 0.0,
                x1: 20.0,
                y1: 0.0,
                stops: vec![
                    GradientStop {
                        offset: 0.0,
                        color: 0x0000_0000,
                    },
                    GradientStop {
                        offset: 1.0,
                        color: 0xffff_ffff,
                    },
                ],
            },
            visual: Box::new(Visual::Group {
                children: vec![VisualBrushChild {
                    bounds: Rect {
                        x: 1.0,
                        y: 2.0,
                        width: 20.0,
                        height: 10.0,
                    },
                    visual: Visual::PaintedShape {
                        geometry: Geometry::Rectangle,
                        fill: Paint::Solid(0xff00_00ff),
                        stroke: Paint::None,
                        stroke_width: 0.0,
                    },
                }],
            }),
        };
        let expected = visual_length(&visual).expect("opacity-mask group length");
        let mut writer = Writer::new(expected).expect("writer allocation");
        write_visual(&mut writer, &visual).expect("opacity-mask group payload");
        assert!(!writer.overflowed);
        assert_eq!(writer.bytes.len(), expected);
    }

    #[test]
    fn color_managed_image_payload_has_an_exact_recursive_length() {
        let visual = Visual::ColorManagedImage {
            source_profile: vec![1, 2, 3, 4],
            destination_profile: Some(vec![5, 6, 7, 8]),
            visual: Box::new(Visual::Image {
                media_type: "image/png".to_owned(),
                bytes: vec![0x89, b'P', b'N', b'G'],
                crop: ImageCrop::default(),
            }),
        };
        let expected = visual_length(&visual).expect("color-managed image length");
        let mut writer = Writer::new(expected).expect("writer allocation");
        write_visual(&mut writer, &visual).expect("color-managed image payload");
        assert!(!writer.overflowed);
        assert_eq!(writer.bytes.len(), expected);
    }

    #[test]
    fn xps_gradient_payload_has_an_exact_validated_length() {
        let paint = Paint::XpsGradient {
            radial: true,
            start_x: 0.25,
            start_y: 0.5,
            end_x: 0.5,
            end_y: 0.5,
            radius_x: 0.5,
            radius_y: 0.25,
            relative: true,
            spread: GradientSpread::Reflect,
            linear_rgb: true,
            transform: AffineTransform::IDENTITY,
            relative_transform: AffineTransform::IDENTITY,
            stops: vec![
                XpsGradientStop {
                    offset: 0.0,
                    color: XpsColor::Rgba(0xff00_00ff),
                },
                XpsGradientStop {
                    offset: 1.0,
                    color: XpsColor::Context {
                        alpha: 1.0,
                        profile: vec![1, 2, 3, 4],
                        channels: vec![0.0, 1.0, 0.0],
                    },
                },
            ],
        };
        let visual = Visual::PaintedShape {
            geometry: Geometry::Rectangle,
            fill: paint,
            stroke: Paint::None,
            stroke_width: 0.0,
        };
        let expected = visual_length(&visual).expect("XPS gradient length");
        let mut writer = Writer::new(expected).expect("writer allocation");
        write_visual(&mut writer, &visual).expect("XPS gradient payload");
        assert!(!writer.overflowed);
        assert_eq!(writer.bytes.len(), expected);
    }

    #[test]
    fn visual_brush_payload_has_an_exact_recursive_length() {
        let visual = Visual::PaintedShape {
            geometry: Geometry::Rectangle,
            fill: Paint::Visual {
                viewbox: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 10.0,
                    height: 10.0,
                },
                viewport: Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 0.25,
                    height: 0.25,
                },
                viewbox_relative: false,
                viewport_relative: true,
                tile_mode: TileMode::FlipXY,
                stretch: StretchMode::Uniform,
                alignment_x: 0.5,
                alignment_y: 0.5,
                transform: AffineTransform::IDENTITY,
                relative_transform: AffineTransform::IDENTITY,
                opacity: 0.75,
                children: vec![VisualBrushChild {
                    bounds: Rect {
                        x: 0.0,
                        y: 0.0,
                        width: 10.0,
                        height: 10.0,
                    },
                    visual: Visual::PaintedShape {
                        geometry: Geometry::Rectangle,
                        fill: Paint::Solid(0x1234_56ff),
                        stroke: Paint::None,
                        stroke_width: 0.0,
                    },
                }],
            },
            stroke: Paint::None,
            stroke_width: 0.0,
        };
        let expected = visual_length(&visual).expect("visual brush length");
        let mut writer = Writer::new(expected).expect("writer allocation");
        write_visual(&mut writer, &visual).expect("visual brush payload");
        assert!(!writer.overflowed);
        assert_eq!(writer.bytes.len(), expected);
    }

    #[test]
    fn new_visuals_write_their_complete_v4_payload() {
        let visuals = [
            Visual::Layer {
                transform: AffineTransform {
                    a: 0.0,
                    b: 1.0,
                    c: -1.0,
                    d: 0.0,
                    e: 100.0,
                    f: 20.0,
                },
                opacity: 0.5,
                blend_mode: crate::model::BlendMode::Normal,
                visual: Box::new(Visual::PaintedShape {
                    geometry: Geometry::RoundedRectangle {
                        radius_x: 8.0,
                        radius_y: 6.0,
                    },
                    fill: Paint::LinearGradient {
                        x0: 0.0,
                        y0: 0.0,
                        x1: 80.0,
                        y1: 0.0,
                        stops: vec![
                            GradientStop {
                                offset: 0.0,
                                color: 0xff00_00ff,
                            },
                            GradientStop {
                                offset: 1.0,
                                color: 0x0000_ffff,
                            },
                        ],
                    },
                    stroke: Paint::Solid(0x1122_33ff),
                    stroke_width: 2.0,
                }),
            },
            Visual::RichText {
                geometry: Geometry::Path {
                    fill_rule: FillRule::EvenOdd,
                    commands: vec![
                        PathCommand::MoveTo { x: 0.0, y: 0.0 },
                        PathCommand::BezierCurveTo {
                            cp1x: 10.0,
                            cp1y: 0.0,
                            cp2x: 10.0,
                            cp2y: 20.0,
                            x: 0.0,
                            y: 20.0,
                        },
                        PathCommand::ClosePath,
                    ],
                },
                fill: Paint::None,
                stroke: Paint::None,
                stroke_width: 0.0,
                align: TextAlign::Center,
                line_height: 18.0,
                runs: vec![TextRun {
                    paint: None,
                    east_asian_line_breaks: true,
                    text: "Office".to_owned(),
                    font_family: "Aptos".to_owned(),
                    font_size: 12.0,
                    color: 0x0000_00ff,
                    bold: true,
                    italic: false,
                    underline: true,
                    strikethrough: false,
                    highlight: 0xffee_88ff,
                    baseline_shift: 2.0,
                    letter_spacing: 0.5,
                    horizontal_scale: 1.0,
                }],
            },
        ];

        for visual in visuals {
            let expected = visual_length(&visual).expect("visual length");
            let mut writer = Writer::new(expected).expect("writer allocation");
            write_visual(&mut writer, &visual).expect("v4 visual payload");
            assert!(!writer.overflowed);
            assert_eq!(writer.bytes.len(), expected);
        }
    }

    #[test]
    fn compositing_layers_write_their_blend_mode_into_the_reserved_wire_byte() {
        let visual = Visual::Layer {
            transform: AffineTransform::IDENTITY,
            opacity: 1.0,
            blend_mode: crate::model::BlendMode::ColorDodge,
            visual: Box::new(Visual::None),
        };
        let expected = visual_length(&visual).expect("visual length");
        let mut writer = Writer::new(expected).expect("writer allocation");
        write_visual(&mut writer, &visual).expect("layer visual payload");

        assert_eq!(writer.bytes.len(), 32);
        assert_eq!(writer.bytes[29], crate::model::BlendMode::ColorDodge.code());
    }

    #[test]
    fn v5_effects_and_text_layout_write_their_complete_payload() {
        let visual = Visual::TextLayout {
            layout: TextLayout {
                wrap_regions: vec![Rect {
                    x: -4.0,
                    y: 12.0,
                    width: 30.0,
                    height: 40.0,
                }],
                direction: TextDirection::Rtl,
                orientation: TextOrientation::VerticalRl,
                auto_fit: TextAutoFit::Shrink,
                vertical_align: TextVerticalAlign::Center,
                prefix: Some("1. ".to_owned()),
                tab_stops: vec![24.0.into(), 48.0.into()],
                default_tab_stop: 36.0,
                hanging_indent: 12.0,
                min_scale: 0.5,
                text_scale_to_fit: true,
                text_matrix_scale_to_fit: true,
                text_stroke_width: 2.0,
                text_stroke_paint: Paint::LinearGradient {
                    x0: 0.0,
                    y0: 0.0,
                    x1: 80.0,
                    y1: 0.0,
                    stops: vec![
                        GradientStop {
                            offset: 0.0,
                            color: 0xff00_00ff,
                        },
                        GradientStop {
                            offset: 1.0,
                            color: 0xaa00_00ff,
                        },
                    ],
                },
                ..TextLayout::default()
            },
            visual: Box::new(Visual::Effect {
                shadow: Some(Shadow {
                    color: 0x1122_3380,
                    blur: 6.0,
                    offset_x: 3.0,
                    offset_y: 4.0,
                }),
                clip: Some(Geometry::RoundedRectangle {
                    radius_x: 8.0,
                    radius_y: 6.0,
                }),
                visual: Box::new(Visual::PaintedShape {
                    geometry: Geometry::Ellipse,
                    fill: Paint::RadialGradient {
                        x0: 20.0,
                        y0: 20.0,
                        r0: 0.0,
                        x1: 20.0,
                        y1: 20.0,
                        r1: 40.0,
                        stops: vec![
                            GradientStop {
                                offset: 0.0,
                                color: 0xffff_ffff,
                            },
                            GradientStop {
                                offset: 1.0,
                                color: 0x0000_00ff,
                            },
                        ],
                    },
                    stroke: Paint::None,
                    stroke_width: 0.0,
                }),
            }),
        };

        let expected = visual_length(&visual).expect("visual length");
        let mut writer = Writer::new(expected).expect("writer allocation");
        write_visual(&mut writer, &visual).expect("v5 visual payload");
        assert!(!writer.overflowed);
        assert_eq!(writer.bytes.len(), expected);
    }

    #[test]
    fn v5_encoder_rejects_invalid_effect_and_text_layout_values() {
        let invalid_shadow = Visual::Effect {
            shadow: Some(Shadow {
                color: 0,
                blur: f32::NAN,
                offset_x: 0.0,
                offset_y: 0.0,
            }),
            clip: None,
            visual: Box::new(Visual::None),
        };
        assert!(visual_length(&invalid_shadow).is_err());

        let invalid_layout = Visual::TextLayout {
            layout: TextLayout {
                tab_stops: (1..=257).map(|value| (value as f32).into()).collect(),
                ..TextLayout::default()
            },
            visual: Box::new(Visual::None),
        };
        assert!(visual_length(&invalid_layout).is_err());

        let invalid_gradient = Visual::PaintedShape {
            geometry: Geometry::Rectangle,
            fill: Paint::RadialGradient {
                x0: 0.0,
                y0: 0.0,
                r0: -1.0,
                x1: 0.0,
                y1: 0.0,
                r1: 10.0,
                stops: vec![GradientStop {
                    offset: 0.0,
                    color: 0,
                }],
            },
            stroke: Paint::None,
            stroke_width: 0.0,
        };
        assert!(visual_length(&invalid_gradient).is_err());
    }

    #[test]
    fn image_color_change_writes_a_bounded_nested_visual_payload() {
        let visual = Visual::ImageColorChange {
            from: 0xffff_ffff,
            to: 0xffff_ff00,
            use_alpha: false,
            visual: Box::new(Visual::Image {
                media_type: "image/png".to_owned(),
                bytes: vec![1, 2, 3],
                crop: ImageCrop::default(),
            }),
        };
        let expected = visual_length(&visual).expect("color-change visual length");
        let mut writer = Writer::new(expected).expect("writer allocation");

        write_visual(&mut writer, &visual).expect("color-change visual payload");

        assert!(!writer.overflowed);
        assert_eq!(writer.bytes.len(), expected);
        assert_eq!(
            &writer.bytes[..8],
            &[0xff, 0xff, 0xff, 0xff, 0x00, 0xff, 0xff, 0xff]
        );
        assert_eq!(&writer.bytes[8..12], &[0, 3, 0, 0]);
    }

    #[test]
    fn embedded_media_writes_its_bytes_and_nested_poster() {
        let visual = Visual::Media {
            kind: MediaKind::Audio,
            media_type: "audio/mpeg".to_owned(),
            bytes: vec![0x49, 0x44, 0x33],
            poster: Box::new(Visual::None),
        };
        let expected = visual_length(&visual).expect("media visual length");
        let mut writer = Writer::new(expected).expect("writer allocation");

        write_visual(&mut writer, &visual).expect("media visual payload");

        assert!(!writer.overflowed);
        assert_eq!(writer.bytes.len(), expected);
        assert_eq!(&writer.bytes[..4], &[0, 0, 0, 0]);
    }
}
