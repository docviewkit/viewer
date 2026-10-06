//! Bounded, content-sniffed adapters for non-package text formats.

use std::collections::HashMap;

use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
use crate::font_metrics::FontMetricTable;
use crate::limits::Limits;
use crate::model::{
    Document, DocumentFormat, DocumentKind, Geometry, ImageCrop, MappingQuality, Object,
    ObjectKind, Paint, Rect, SourceLocator, SourceRef, TextAlign, TextLayout, TextRun, Unit,
    UnitKind, Visual,
};

use super::{
    presentation_image::{
        OfficeImageError, office_image_media_type_from_mime, reserve_materialized_image_bytes,
    },
    reserve_materialized_text_bytes,
};

const PAGE_WIDTH: f32 = 816.0;
const PAGE_HEIGHT: f32 = 1056.0;
const PAGE_MARGIN: f32 = 96.0;
const FONT_SIZE: f32 = 16.0;
const LINE_HEIGHT: f32 = 24.0;
const APPROXIMATE_CHARACTER_WIDTH: f32 = 8.0;

#[derive(Clone, Copy, Debug)]
struct RtfPageSpec {
    width: f32,
    height: f32,
    margin_left: f32,
    margin_right: f32,
    margin_top: f32,
    margin_bottom: f32,
}

impl Default for RtfPageSpec {
    fn default() -> Self {
        Self {
            width: PAGE_WIDTH,
            height: PAGE_HEIGHT,
            margin_left: PAGE_MARGIN,
            margin_right: PAGE_MARGIN,
            margin_top: PAGE_MARGIN,
            margin_bottom: PAGE_MARGIN,
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct RtfParagraphStyle {
    align: TextAlign,
    space_before: f32,
    space_after: f32,
    left_indent: f32,
    right_indent: f32,
    first_line_indent: f32,
    line_spacing: Option<i32>,
    line_spacing_multiple: bool,
    keep_next: bool,
}

impl Default for RtfParagraphStyle {
    fn default() -> Self {
        Self {
            align: TextAlign::Start,
            space_before: 0.0,
            space_after: 0.0,
            left_indent: 0.0,
            right_indent: 0.0,
            first_line_indent: 0.0,
            line_spacing: None,
            line_spacing_multiple: false,
            keep_next: false,
        }
    }
}

#[derive(Clone, Debug)]
struct RtfPicture {
    media_type: String,
    bytes: Vec<u8>,
    width: f32,
    height: f32,
    offset_x: f32,
    offset_y: f32,
    floating: bool,
    source_offset: u32,
}

#[derive(Clone, Debug, Default)]
struct RtfTableRow {
    left: f32,
    minimum_height: f32,
    cell_boundaries: Vec<f32>,
    cells: Vec<RichTextBlock>,
}

#[derive(Default, Clone, Debug)]
struct RichTextBlock {
    text: String,
    runs: Vec<TextRun>,
    style: RtfParagraphStyle,
    picture: Option<RtfPicture>,
    table_row: Option<RtfTableRow>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum RtfDestination {
    Visible,
    Footnote,
    FontTable,
    ColorTable,
    Skip,
}

#[derive(Clone, Debug)]
struct RtfState {
    destination: RtfDestination,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    baseline_shift: f32,
    font_size: f32,
    font_index: i32,
    color_index: usize,
    unicode_skip: usize,
    pending_ignorable: bool,
    paragraph: RtfParagraphStyle,
}

impl Default for RtfState {
    fn default() -> Self {
        Self {
            destination: RtfDestination::Visible,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            baseline_shift: 0.0,
            font_size: FONT_SIZE,
            font_index: 0,
            color_index: 0,
            unicode_skip: 1,
            pending_ignorable: false,
            paragraph: RtfParagraphStyle::default(),
        }
    }
}

pub fn detect_and_parse(
    bytes: &[u8],
    limits: Limits,
    font_metrics: &FontMetricTable,
) -> Result<Option<Document>, Diagnostic> {
    enforce_input_budget(bytes, limits)?;
    if is_known_unsupported_binary(bytes) {
        return Ok(None);
    }
    if looks_like_rtf(bytes) {
        return parse_rtf(bytes, limits, font_metrics).map(Some);
    }
    let Some(text) = decode_text(bytes)? else {
        return Ok(None);
    };
    if looks_like_html_input(&text) {
        return Ok(None);
    }
    if text.contains(',') && (text.contains('\n') || text.contains('\r')) {
        let rows = parse_csv_records(&text, limits)?;
        if rows.len() >= 2 && rows.iter().all(|row| row.len() >= 2) {
            return parse_csv(rows, limits).map(Some);
        }
    }
    Ok(None)
}

fn is_known_unsupported_binary(bytes: &[u8]) -> bool {
    bytes.starts_with(b"%PDF-")
        || bytes.starts_with(&[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
}

fn enforce_input_budget(bytes: &[u8], limits: Limits) -> Result<(), Diagnostic> {
    if bytes.len() > limits.max_entry_uncompressed_bytes
        || bytes.len() > limits.max_total_uncompressed_bytes
    {
        return Err(Diagnostic::fatal(
            DiagnosticCode::InputTooLarge,
            Phase::Input,
            None,
            "flat document exceeds the configured byte budget",
        ));
    }
    Ok(())
}

fn decode_text(bytes: &[u8]) -> Result<Option<String>, Diagnostic> {
    let decoded = if let Some(bytes) = bytes.strip_prefix(&[0xef, 0xbb, 0xbf]) {
        std::str::from_utf8(bytes)
            .map(str::to_owned)
            .map_err(|_| format_error("UTF-8 text after the byte-order mark is malformed"))?
    } else if let Some(bytes) = bytes.strip_prefix(&[0xff, 0xfe]) {
        decode_utf16(bytes, true)?
    } else if let Some(bytes) = bytes.strip_prefix(&[0xfe, 0xff]) {
        decode_utf16(bytes, false)?
    } else {
        let Ok(text) = std::str::from_utf8(bytes) else {
            return Ok(None);
        };
        text.to_owned()
    };
    if decoded.chars().any(|character| {
        character == '\0'
            || (character < '\u{20}' && !matches!(character, '\t' | '\n' | '\r' | '\u{0c}'))
    }) {
        return Ok(None);
    }
    Ok(Some(decoded))
}

fn decode_utf16(bytes: &[u8], little_endian: bool) -> Result<String, Diagnostic> {
    if !bytes.len().is_multiple_of(2) {
        return Err(format_error("UTF-16 text has a truncated code unit"));
    }
    let units = bytes.chunks_exact(2).map(|bytes| {
        if little_endian {
            u16::from_le_bytes([bytes[0], bytes[1]])
        } else {
            u16::from_be_bytes([bytes[0], bytes[1]])
        }
    });
    char::decode_utf16(units)
        .collect::<Result<String, _>>()
        .map_err(|_| format_error("UTF-16 text contains an unpaired surrogate"))
}

fn looks_like_rtf(bytes: &[u8]) -> bool {
    let bytes = bytes.strip_prefix(&[0xef, 0xbb, 0xbf]).unwrap_or(bytes);
    let position = bytes
        .iter()
        .position(|byte| !byte.is_ascii_whitespace())
        .unwrap_or(bytes.len());
    starts_with_ascii_case_insensitive(&bytes[position..], b"{\\rtf")
}

struct RtfParser<'a> {
    bytes: &'a [u8],
    limits: Limits,
    position: usize,
    nodes: usize,
    state: RtfState,
    stack: Vec<RtfState>,
    group_offsets: Vec<usize>,
    paragraphs: Vec<RichTextBlock>,
    current: RichTextBlock,
    footnotes: Vec<RichTextBlock>,
    current_footnote: Option<RichTextBlock>,
    footnote_number: u32,
    fonts: HashMap<i32, String>,
    colors: Vec<u32>,
    font_buffer: String,
    color_red: u8,
    color_green: u8,
    color_blue: u8,
    code_page: u32,
    fallback_skip: usize,
    pending_high_surrogate: Option<u16>,
    default_font: i32,
    diagnostics: Vec<Diagnostic>,
    omitted_picture_reported: bool,
    active_object_reported: bool,
    page: RtfPageSpec,
    materialized_image_bytes: usize,
    table_row: Option<RtfTableRow>,
}

pub(super) fn parse_rtf(
    bytes: &[u8],
    limits: Limits,
    font_metrics: &FontMetricTable,
) -> Result<Document, Diagnostic> {
    let bytes = bytes.strip_prefix(&[0xef, 0xbb, 0xbf]).unwrap_or(bytes);
    if !looks_like_rtf(bytes) {
        return Err(rtf_error("RTF signature is missing"));
    }
    RtfParser {
        bytes,
        limits,
        position: 0,
        nodes: 0,
        state: RtfState::default(),
        stack: Vec::new(),
        group_offsets: Vec::new(),
        paragraphs: Vec::new(),
        current: RichTextBlock::default(),
        footnotes: Vec::new(),
        current_footnote: None,
        footnote_number: 0,
        fonts: HashMap::new(),
        colors: Vec::new(),
        font_buffer: String::new(),
        color_red: 0,
        color_green: 0,
        color_blue: 0,
        code_page: 1252,
        fallback_skip: 0,
        pending_high_surrogate: None,
        default_font: 0,
        diagnostics: Vec::new(),
        omitted_picture_reported: false,
        active_object_reported: false,
        page: RtfPageSpec::default(),
        materialized_image_bytes: 0,
        table_row: None,
    }
    .parse(font_metrics)
}

impl RtfParser<'_> {
    fn parse(mut self, font_metrics: &FontMetricTable) -> Result<Document, Diagnostic> {
        while self.position < self.bytes.len() {
            let byte = self.bytes[self.position];
            self.position += 1;
            match byte {
                b'{' => self.start_group()?,
                b'}' => self.end_group()?,
                b'\\' => self.parse_control()?,
                b'\r' | b'\n' => {}
                _ => self.consume_plain_byte(byte)?,
            }
        }
        if !self.stack.is_empty() {
            return Err(rtf_error("RTF input ends before all groups close"));
        }
        if self.table_row.is_some() {
            return Err(rtf_error("RTF input ends before a table row closes"));
        }
        self.flush_pending_surrogate()?;
        self.flush_paragraph(false)?;
        layout_rtf(
            self.paragraphs,
            self.footnotes,
            self.diagnostics,
            self.page,
            self.limits,
            font_metrics,
        )
    }

    fn start_group(&mut self) -> Result<(), Diagnostic> {
        self.consume_node()?;
        if self.stack.len() >= self.limits.max_xml_depth {
            return Err(Diagnostic::fatal(
                DiagnosticCode::XmlDepthLimit,
                Phase::Parse,
                Some(self.position.saturating_sub(1)),
                "RTF group nesting exceeds the configured depth limit",
            )
            .in_part("input.rtf"));
        }
        self.stack.push(self.state.clone());
        self.group_offsets.push(self.position.saturating_sub(1));
        Ok(())
    }

    fn end_group(&mut self) -> Result<(), Diagnostic> {
        self.consume_node()?;
        let Some(previous) = self.stack.pop() else {
            return Err(rtf_error("RTF input contains an unmatched closing brace"));
        };
        if self.state.destination == RtfDestination::Footnote
            && previous.destination != RtfDestination::Footnote
        {
            self.finish_footnote()?;
        }
        self.group_offsets.pop();
        self.state = previous;
        Ok(())
    }

    fn parse_control(&mut self) -> Result<(), Diagnostic> {
        self.consume_node()?;
        let Some(&next) = self.bytes.get(self.position) else {
            return Err(rtf_error("RTF input ends after a control escape"));
        };
        self.position += 1;
        match next {
            b'\r' | b'\n' => self.handle_control_word("par", None, self.position.saturating_sub(2)),
            b'\\' | b'{' | b'}' => self.consume_character_control(next as char),
            b'~' => self.consume_character_control('\u{a0}'),
            b'_' => self.consume_character_control('\u{2011}'),
            b'-' => self.consume_character_control('\u{00ad}'),
            b'*' => {
                self.state.pending_ignorable = true;
                Ok(())
            }
            b'\'' => self.parse_hex_escape(),
            byte if byte.is_ascii_alphabetic() => self.parse_control_word(byte),
            _ => Ok(()),
        }
    }

    fn parse_hex_escape(&mut self) -> Result<(), Diagnostic> {
        let end = self
            .position
            .checked_add(2)
            .ok_or_else(|| rtf_error("RTF hex escape overflows the input range"))?;
        let Some(value) = self.bytes.get(self.position..end) else {
            return Err(rtf_error("RTF hex escape is truncated"));
        };
        let value = std::str::from_utf8(value)
            .ok()
            .and_then(|value| u8::from_str_radix(value, 16).ok())
            .ok_or_else(|| rtf_error("RTF hex escape is malformed"))?;
        self.position = end;
        if self.fallback_skip > 0 {
            self.fallback_skip -= 1;
            return Ok(());
        }
        self.consume_decoded_character(decode_ansi_byte(value, self.code_page))
    }

    fn parse_control_word(&mut self, first: u8) -> Result<(), Diagnostic> {
        let start = self.position - 1;
        while self
            .bytes
            .get(self.position)
            .is_some_and(u8::is_ascii_alphabetic)
        {
            self.position += 1;
        }
        if self.position - start > self.limits.max_xml_name_bytes {
            return Err(rtf_error(
                "RTF control word exceeds the configured name byte limit",
            ));
        }
        let mut word = String::with_capacity(self.position - start);
        word.push((first as char).to_ascii_lowercase());
        for byte in &self.bytes[start + 1..self.position] {
            word.push((*byte as char).to_ascii_lowercase());
        }
        let sign = if self.bytes.get(self.position) == Some(&b'-') {
            self.position += 1;
            -1_i64
        } else {
            1_i64
        };
        let number_start = self.position;
        while self
            .bytes
            .get(self.position)
            .is_some_and(u8::is_ascii_digit)
        {
            self.position += 1;
        }
        let parameter = if self.position > number_start {
            let value = std::str::from_utf8(&self.bytes[number_start..self.position])
                .ok()
                .and_then(|value| value.parse::<i64>().ok())
                .and_then(|value| value.checked_mul(sign))
                .ok_or_else(|| rtf_error("RTF control parameter is out of range"))?;
            Some(
                i32::try_from(value)
                    .map_err(|_| rtf_error("RTF control parameter is out of range"))?,
            )
        } else {
            None
        };
        if self.bytes.get(self.position) == Some(&b' ') {
            self.position += 1;
        }
        self.handle_control_word(&word, parameter, start.saturating_sub(1))
    }

    fn handle_control_word(
        &mut self,
        word: &str,
        parameter: Option<i32>,
        control_offset: usize,
    ) -> Result<(), Diagnostic> {
        if self.state.pending_ignorable {
            self.state.pending_ignorable = false;
            self.state.destination = RtfDestination::Skip;
        }
        match word {
            "footnote" if self.state.destination == RtfDestination::Visible => {
                self.footnote_number = self.footnote_number.max(1);
                self.current_footnote = Some(RichTextBlock::default());
                self.state.destination = RtfDestination::Footnote;
                return Ok(());
            }
            "fonttbl" => {
                self.state.destination = RtfDestination::FontTable;
                return Ok(());
            }
            "colortbl" => {
                self.state.destination = RtfDestination::ColorTable;
                return Ok(());
            }
            "stylesheet" | "info" | "filetbl" | "listtable" | "listoverridetable" | "generator"
            | "xmlnstbl" | "datastore" | "themedata" | "colorschememapping" | "latentstyles"
            | "fldinst" | "listtext" => {
                self.state.destination = RtfDestination::Skip;
                return Ok(());
            }
            "pict" => {
                if rtf_picture_is_fallback(self.bytes, &self.group_offsets, control_offset) {
                    self.state.destination = RtfDestination::Skip;
                    return Ok(());
                }
                if let Some(picture) = parse_rtf_picture(
                    self.bytes,
                    &self.group_offsets,
                    control_offset,
                    self.limits,
                    &mut self.materialized_image_bytes,
                    &mut self.diagnostics,
                )? {
                    self.push_picture(picture)?;
                } else if !self.omitted_picture_reported {
                    self.diagnostics.push(
                        Diagnostic::warning(
                            DiagnosticCode::UnsupportedFeature,
                            Phase::Parse,
                            Fidelity::Omitted,
                            "RTF picture has no safe supported raster or metafile payload",
                        )
                        .in_part("input.rtf"),
                    );
                    self.omitted_picture_reported = true;
                }
                self.state.destination = RtfDestination::Skip;
                return Ok(());
            }
            "shppict" => return Ok(()),
            "nonshppict" => {
                self.state.destination = RtfDestination::Skip;
                return Ok(());
            }
            "object" | "objdata" => {
                self.state.destination = RtfDestination::Skip;
                if !self.active_object_reported {
                    self.diagnostics.push(
                        Diagnostic::warning(
                            DiagnosticCode::ActiveContentBlocked,
                            Phase::Security,
                            Fidelity::Blocked,
                            "embedded RTF objects are blocked and never executed",
                        )
                        .in_part("input.rtf"),
                    );
                    self.active_object_reported = true;
                }
                return Ok(());
            }
            _ => {}
        }

        if self.state.destination == RtfDestination::FontTable {
            if word == "f"
                && let Some(index) = parameter
            {
                self.state.font_index = index;
                self.font_buffer.clear();
            }
            return Ok(());
        }
        if self.state.destination == RtfDestination::ColorTable {
            let value = parameter.unwrap_or_default().clamp(0, 255) as u8;
            match word {
                "red" => self.color_red = value,
                "green" => self.color_green = value,
                "blue" => self.color_blue = value,
                _ => {}
            }
            return Ok(());
        }
        if !matches!(
            self.state.destination,
            RtfDestination::Visible | RtfDestination::Footnote
        ) {
            return Ok(());
        }

        match word {
            "ansi" => self.code_page = 1252,
            "ansicpg" => {
                if let Some(code_page) = parameter.and_then(|value| u32::try_from(value).ok()) {
                    self.code_page = code_page;
                }
            }
            "deff" => {
                if let Some(index) = parameter {
                    self.default_font = index;
                    self.state.font_index = index;
                }
            }
            "f" => {
                if let Some(index) = parameter {
                    self.state.font_index = index;
                }
            }
            "cf" => {
                if let Some(index) = parameter.and_then(|value| usize::try_from(value).ok()) {
                    self.state.color_index = index;
                }
            }
            "fs" => {
                let size = parameter.unwrap_or(24);
                if !(1..=32_767).contains(&size) {
                    return Err(rtf_error("RTF font size is outside the supported range"));
                }
                self.state.font_size = size as f32 / 2.0 * 4.0 / 3.0;
            }
            "b" => self.state.bold = parameter != Some(0),
            "i" => self.state.italic = parameter != Some(0),
            "ul" => self.state.underline = parameter != Some(0),
            "ulnone" => self.state.underline = false,
            "strike" => self.state.strikethrough = parameter != Some(0),
            "super" => self.state.baseline_shift = self.state.font_size * 0.35,
            "sub" => self.state.baseline_shift = -self.state.font_size * 0.2,
            "nosupersub" => self.state.baseline_shift = 0.0,
            "plain" => {
                let destination = self.state.destination;
                let paragraph = self.state.paragraph;
                let unicode_skip = self.state.unicode_skip;
                self.state = RtfState {
                    destination,
                    font_index: self.default_font,
                    unicode_skip,
                    paragraph,
                    ..RtfState::default()
                };
            }
            "pard" => self.state.paragraph = RtfParagraphStyle::default(),
            "ql" => self.state.paragraph.align = TextAlign::Start,
            "qc" => self.state.paragraph.align = TextAlign::Center,
            "qr" => self.state.paragraph.align = TextAlign::End,
            "qj" => self.state.paragraph.align = TextAlign::Justify,
            "sb" => self.state.paragraph.space_before = twips(parameter.unwrap_or_default()),
            "sa" => self.state.paragraph.space_after = twips(parameter.unwrap_or_default()),
            "li" | "lin" => self.state.paragraph.left_indent = twips(parameter.unwrap_or_default()),
            "ri" | "rin" => {
                self.state.paragraph.right_indent = twips(parameter.unwrap_or_default())
            }
            "fi" => self.state.paragraph.first_line_indent = twips(parameter.unwrap_or_default()),
            "sl" => self.state.paragraph.line_spacing = parameter.filter(|value| *value != 0),
            "slmult" => self.state.paragraph.line_spacing_multiple = parameter == Some(1),
            "keepn" => self.state.paragraph.keep_next = parameter != Some(0),
            "paperw" | "pgwsxn" => {
                if let Some(value) = positive_twips(parameter) {
                    self.page.width = value;
                }
            }
            "paperh" | "pghsxn" => {
                if let Some(value) = positive_twips(parameter) {
                    self.page.height = value;
                }
            }
            "margl" | "marglsxn" => {
                if let Some(value) = nonnegative_twips(parameter) {
                    self.page.margin_left = value;
                }
            }
            "margr" | "margrsxn" => {
                if let Some(value) = nonnegative_twips(parameter) {
                    self.page.margin_right = value;
                }
            }
            "margt" | "margtsxn" => {
                if let Some(value) = nonnegative_twips(parameter) {
                    self.page.margin_top = value;
                }
            }
            "margb" | "margbsxn" => {
                if let Some(value) = nonnegative_twips(parameter) {
                    self.page.margin_bottom = value;
                }
            }
            "trowd" => self.start_table_row()?,
            "trleft" => {
                if let Some(row) = self.table_row.as_mut() {
                    row.left = twips(parameter.unwrap_or_default());
                }
            }
            "trrh" => {
                if let Some(row) = self.table_row.as_mut() {
                    row.minimum_height = twips(parameter.unwrap_or_default().saturating_abs());
                }
            }
            "cellx" => self.push_table_boundary(parameter)?,
            "cell" => self.flush_table_cell()?,
            "row" => self.finish_table_row()?,
            "uc" => {
                if let Some(skip) = parameter.and_then(|value| usize::try_from(value).ok()) {
                    self.state.unicode_skip = skip.min(32);
                }
            }
            "u" => {
                let value =
                    parameter.ok_or_else(|| rtf_error("RTF Unicode escape has no value"))?;
                let unit = (i64::from(value) & 0xffff) as u16;
                self.append_unicode_unit(unit)?;
                self.fallback_skip = self.state.unicode_skip;
            }
            "chftn" => {
                if self.state.destination == RtfDestination::Visible {
                    self.footnote_number = self
                        .footnote_number
                        .checked_add(1)
                        .ok_or_else(object_limit)?;
                }
                let number = self.footnote_number.max(1).to_string();
                self.append_visible(&number)?;
            }
            "par" => {
                if self.state.destination == RtfDestination::Footnote {
                    self.append_visible("\n")?;
                } else if self.table_row.is_some() {
                    if !self.current.text.is_empty() {
                        self.append_visible("\n")?;
                    }
                } else {
                    self.flush_paragraph(true)?;
                }
            }
            "line" => self.append_visible("\n")?,
            "tab" => self.append_visible("\t")?,
            "bin" => {
                let length = parameter
                    .and_then(|value| usize::try_from(value).ok())
                    .ok_or_else(|| rtf_error("RTF binary payload has an invalid length"))?;
                let end = self
                    .position
                    .checked_add(length)
                    .filter(|end| *end <= self.bytes.len())
                    .ok_or_else(|| rtf_error("RTF binary payload is truncated"))?;
                self.position = end;
            }
            _ => {}
        }
        Ok(())
    }

    fn consume_plain_byte(&mut self, byte: u8) -> Result<(), Diagnostic> {
        match self.state.destination {
            RtfDestination::FontTable => {
                if byte == b';' {
                    let name = self.font_buffer.trim().to_owned();
                    if !name.is_empty() {
                        if self.fonts.len() >= self.limits.max_document_objects {
                            return Err(object_limit());
                        }
                        self.fonts.insert(self.state.font_index, name);
                    }
                    self.font_buffer.clear();
                } else {
                    self.font_buffer
                        .push(decode_ansi_byte(byte, self.code_page));
                }
                Ok(())
            }
            RtfDestination::ColorTable => {
                if byte == b';' {
                    if self.colors.len() >= self.limits.max_document_objects {
                        return Err(object_limit());
                    }
                    self.colors.push(
                        (u32::from(self.color_red) << 24)
                            | (u32::from(self.color_green) << 16)
                            | (u32::from(self.color_blue) << 8)
                            | 0xff,
                    );
                    self.color_red = 0;
                    self.color_green = 0;
                    self.color_blue = 0;
                }
                Ok(())
            }
            RtfDestination::Skip => Ok(()),
            RtfDestination::Visible | RtfDestination::Footnote => {
                if self.fallback_skip > 0 {
                    self.fallback_skip -= 1;
                    Ok(())
                } else {
                    self.consume_decoded_character(decode_ansi_byte(byte, self.code_page))
                }
            }
        }
    }

    fn consume_character_control(&mut self, character: char) -> Result<(), Diagnostic> {
        if self.fallback_skip > 0 {
            self.fallback_skip -= 1;
            return Ok(());
        }
        self.consume_decoded_character(character)
    }

    fn consume_decoded_character(&mut self, character: char) -> Result<(), Diagnostic> {
        match self.state.destination {
            RtfDestination::Visible | RtfDestination::Footnote => {
                let mut encoded = [0_u8; 4];
                self.append_visible(character.encode_utf8(&mut encoded))
            }
            RtfDestination::FontTable => {
                self.font_buffer.push(character);
                Ok(())
            }
            RtfDestination::ColorTable | RtfDestination::Skip => Ok(()),
        }
    }

    fn append_unicode_unit(&mut self, unit: u16) -> Result<(), Diagnostic> {
        if (0xd800..=0xdbff).contains(&unit) {
            self.flush_pending_surrogate()?;
            self.pending_high_surrogate = Some(unit);
            return Ok(());
        }
        if (0xdc00..=0xdfff).contains(&unit) {
            let Some(high) = self.pending_high_surrogate.take() else {
                return self.consume_decoded_character('\u{fffd}');
            };
            let codepoint =
                0x1_0000 + ((u32::from(high) - 0xd800) << 10) + (u32::from(unit) - 0xdc00);
            return self.consume_decoded_character(char::from_u32(codepoint).unwrap_or('\u{fffd}'));
        }
        self.flush_pending_surrogate()?;
        self.consume_decoded_character(char::from_u32(u32::from(unit)).unwrap_or('\u{fffd}'))
    }

    fn flush_pending_surrogate(&mut self) -> Result<(), Diagnostic> {
        if self.pending_high_surrogate.take().is_some() {
            self.consume_decoded_character('\u{fffd}')?;
        }
        Ok(())
    }

    fn append_visible(&mut self, text: &str) -> Result<(), Diagnostic> {
        if !matches!(
            self.state.destination,
            RtfDestination::Visible | RtfDestination::Footnote
        ) || text.is_empty()
        {
            return Ok(());
        }
        let font_family = self
            .fonts
            .get(&self.state.font_index)
            .map(String::as_str)
            .unwrap_or("Arial");
        let text = super::normalize_symbol_font_character(text, Some(font_family));
        let run = self.rtf_run(text.clone());
        let current = if self.state.destination == RtfDestination::Footnote {
            self.current_footnote
                .as_mut()
                .ok_or_else(|| rtf_error("RTF footnote text has no active destination"))?
        } else {
            &mut self.current
        };
        current.text.push_str(&text);
        if let Some(last) = current.runs.last_mut()
            && last.same_style(&run)
        {
            last.text.push_str(&text);
        } else {
            current.runs.push(run);
        }
        Ok(())
    }

    fn rtf_run(&self, text: String) -> TextRun {
        TextRun {
            paint: None,
            east_asian_line_breaks: true,
            text,
            font_family: self
                .fonts
                .get(&self.state.font_index)
                .cloned()
                .unwrap_or_else(|| "Arial".to_owned()),
            font_size: self.state.font_size,
            color: self
                .colors
                .get(self.state.color_index)
                .copied()
                .unwrap_or(0x0000_00ff),
            bold: self.state.bold,
            italic: self.state.italic,
            underline: self.state.underline,
            strikethrough: self.state.strikethrough,
            highlight: 0,
            baseline_shift: self.state.baseline_shift,
            letter_spacing: 0.0,
            horizontal_scale: 1.0,
        }
    }

    fn finish_footnote(&mut self) -> Result<(), Diagnostic> {
        self.flush_pending_surrogate()?;
        let mut footnote = self
            .current_footnote
            .take()
            .ok_or_else(|| rtf_error("RTF footnote destination has no content state"))?;
        while footnote.text.ends_with('\n') {
            footnote.text.pop();
        }
        while let Some(run) = footnote.runs.last_mut() {
            while run.text.ends_with('\n') {
                run.text.pop();
            }
            if run.text.is_empty() {
                footnote.runs.pop();
            } else {
                break;
            }
        }
        if footnote.runs.is_empty() {
            footnote.runs.push(self.rtf_run(String::new()));
        }
        footnote.style = self.state.paragraph;
        if self.footnotes.len() >= self.limits.max_document_objects {
            return Err(object_limit());
        }
        self.footnotes.push(footnote);
        Ok(())
    }

    fn flush_paragraph(&mut self, include_empty: bool) -> Result<(), Diagnostic> {
        self.flush_pending_surrogate()?;
        if self.current.text.is_empty() && !include_empty {
            return Ok(());
        }
        if self.paragraphs.len() >= self.limits.max_document_objects {
            return Err(object_limit());
        }
        if self.current.runs.is_empty() {
            self.current.runs.push(self.rtf_run(String::new()));
        }
        self.current.style = self.state.paragraph;
        self.paragraphs.push(std::mem::take(&mut self.current));
        Ok(())
    }

    fn start_table_row(&mut self) -> Result<(), Diagnostic> {
        self.flush_paragraph(false)?;
        if self.table_row.is_some() {
            return Err(rtf_error(
                "RTF table starts a new row before the previous row ends",
            ));
        }
        self.table_row = Some(RtfTableRow::default());
        Ok(())
    }

    fn flush_table_cell(&mut self) -> Result<(), Diagnostic> {
        self.flush_pending_surrogate()?;
        if self.current.runs.is_empty() {
            self.current.runs.push(self.rtf_run(String::new()));
        }
        let Some(row) = self.table_row.as_mut() else {
            return Err(rtf_error("RTF table cell appears outside a row"));
        };
        if row.cells.len() >= self.limits.max_document_objects {
            return Err(object_limit());
        }
        self.current.style = self.state.paragraph;
        row.cells.push(std::mem::take(&mut self.current));
        Ok(())
    }

    fn push_table_boundary(&mut self, parameter: Option<i32>) -> Result<(), Diagnostic> {
        let Some(boundary) = parameter.map(twips) else {
            return Err(rtf_error("RTF table cell boundary has no position"));
        };
        let Some(row) = self.table_row.as_mut() else {
            return Err(rtf_error("RTF table cell boundary appears outside a row"));
        };
        if row.cell_boundaries.len() >= self.limits.max_document_objects {
            return Err(object_limit());
        }
        if row
            .cell_boundaries
            .last()
            .is_some_and(|previous| boundary <= *previous)
        {
            return Err(rtf_error("RTF table cell boundaries are not increasing"));
        }
        row.cell_boundaries.push(boundary);
        Ok(())
    }

    fn finish_table_row(&mut self) -> Result<(), Diagnostic> {
        if !self.current.text.is_empty() {
            self.flush_table_cell()?;
        }
        let Some(mut row) = self.table_row.take() else {
            return Err(rtf_error("RTF table row terminator appears outside a row"));
        };
        if row.cell_boundaries.is_empty() {
            return Err(rtf_error("RTF table row has no cell boundaries"));
        }
        if row.cells.len() > row.cell_boundaries.len() {
            return Err(rtf_error("RTF table row has more cells than boundaries"));
        }
        while row.cells.len() < row.cell_boundaries.len() {
            row.cells.push(RichTextBlock::default());
        }
        if self.paragraphs.len() >= self.limits.max_document_objects {
            return Err(object_limit());
        }
        self.paragraphs.push(RichTextBlock {
            table_row: Some(row),
            ..RichTextBlock::default()
        });
        Ok(())
    }

    fn push_picture(&mut self, picture: RtfPicture) -> Result<(), Diagnostic> {
        self.flush_paragraph(false)?;
        if self.paragraphs.len() >= self.limits.max_document_objects {
            return Err(object_limit());
        }
        self.paragraphs.push(RichTextBlock {
            style: self.state.paragraph,
            picture: Some(picture),
            runs: vec![self.rtf_run(String::new())],
            ..RichTextBlock::default()
        });
        Ok(())
    }

    fn consume_node(&mut self) -> Result<(), Diagnostic> {
        self.nodes = self.nodes.checked_add(1).ok_or_else(object_limit)?;
        if self.nodes > self.limits.max_xml_nodes {
            return Err(Diagnostic::fatal(
                DiagnosticCode::XmlNodeLimit,
                Phase::Parse,
                Some(self.position.saturating_sub(1)),
                "RTF token count exceeds the configured parser budget",
            )
            .in_part("input.rtf"));
        }
        Ok(())
    }
}

fn twips(value: i32) -> f32 {
    value as f32 / 15.0
}

fn positive_twips(value: Option<i32>) -> Option<f32> {
    value.filter(|value| *value > 0).map(twips)
}

fn nonnegative_twips(value: Option<i32>) -> Option<f32> {
    value.filter(|value| *value >= 0).map(twips)
}

fn parse_rtf_picture(
    bytes: &[u8],
    group_offsets: &[usize],
    control_offset: usize,
    limits: Limits,
    materialized_image_bytes: &mut usize,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<Option<RtfPicture>, Diagnostic> {
    let Some(&picture_group_start) = group_offsets.last() else {
        return Err(rtf_error("RTF picture appears outside a group"));
    };
    let picture_group_end = rtf_group_end(bytes, picture_group_start)?;
    let picture_prefix = bytes
        .get(picture_group_start..picture_group_end)
        .ok_or_else(|| rtf_error("RTF picture group is outside the input"))?;
    let Some((declared_media_type, marker_offset)) = [
        (b"pngblip".as_slice(), "image/png"),
        (b"jpegblip".as_slice(), "image/jpeg"),
        (b"emfblip".as_slice(), "image/x-emf"),
        (b"wmetafile".as_slice(), "image/x-wmf"),
    ]
    .into_iter()
    .find_map(|(control, media_type)| {
        find_rtf_control(picture_prefix, control)
            .map(|offset| (media_type, picture_group_start + offset))
    }) else {
        return Ok(None);
    };
    let data_start = rtf_control_end(bytes, marker_offset, picture_group_end)?;
    let payload = decode_rtf_picture_hex(bytes, data_start, picture_group_end, limits)?;
    let media_type = match office_image_media_type_from_mime(declared_media_type, &payload) {
        Ok(media_type) => media_type.to_owned(),
        Err(OfficeImageError::SignatureMismatch) => {
            diagnostics.push(
                Diagnostic::warning(
                    DiagnosticCode::FormatInvalid,
                    Phase::Parse,
                    Fidelity::Omitted,
                    "RTF picture payload does not match its declared format",
                )
                .in_part("input.rtf"),
            );
            return Ok(None);
        }
        Err(OfficeImageError::UnsupportedFormat | OfficeImageError::DisabledByOffice) => {
            return Ok(None);
        }
    };
    reserve_materialized_image_bytes(
        materialized_image_bytes,
        payload.len(),
        limits.max_total_uncompressed_bytes,
        "input.rtf",
    )?;

    let picture_width = rtf_control_parameter(picture_prefix, b"picwgoal")
        .map(twips)
        .or_else(|| rtf_control_parameter(picture_prefix, b"picw").map(|value| value as f32));
    let picture_height = rtf_control_parameter(picture_prefix, b"pichgoal")
        .map(twips)
        .or_else(|| rtf_control_parameter(picture_prefix, b"pich").map(|value| value as f32));
    let scale_x = rtf_control_parameter(picture_prefix, b"picscalex")
        .unwrap_or(100)
        .max(1) as f32
        / 100.0;
    let scale_y = rtf_control_parameter(picture_prefix, b"picscaley")
        .unwrap_or(100)
        .max(1) as f32
        / 100.0;
    let mut width = picture_width.unwrap_or(1.0) * scale_x;
    let mut height = picture_height.unwrap_or(1.0) * scale_y;
    let mut offset_x = 0.0;
    let mut offset_y = 0.0;
    let mut floating = false;

    for group_start in group_offsets.iter().rev().skip(1) {
        let prefix = bytes
            .get(*group_start..control_offset)
            .ok_or_else(|| rtf_error("RTF picture ancestry is outside the input"))?;
        let shape_left = rtf_control_parameter(prefix, b"shpleft");
        let shape_right = rtf_control_parameter(prefix, b"shpright");
        let shape_top = rtf_control_parameter(prefix, b"shptop");
        let shape_bottom = rtf_control_parameter(prefix, b"shpbottom");
        if let (Some(left), Some(right), Some(top), Some(bottom)) =
            (shape_left, shape_right, shape_top, shape_bottom)
            && right > left
            && bottom > top
        {
            offset_x = twips(left);
            offset_y = twips(top);
            width = twips(right - left);
            height = twips(bottom - top);
            floating = true;
            break;
        }
    }
    if ![width, height, offset_x, offset_y]
        .into_iter()
        .all(f32::is_finite)
        || width <= 0.0
        || height <= 0.0
    {
        return Err(rtf_error("RTF picture has invalid display geometry"));
    }
    let source_offset = u32::try_from(picture_group_start)
        .map_err(|_| rtf_error("RTF picture offset overflows"))?;
    Ok(Some(RtfPicture {
        media_type,
        bytes: payload,
        width,
        height,
        offset_x,
        offset_y,
        floating,
        source_offset,
    }))
}

fn rtf_group_end(bytes: &[u8], group_start: usize) -> Result<usize, Diagnostic> {
    if bytes.get(group_start) != Some(&b'{') {
        return Err(rtf_error("RTF picture group has no opening brace"));
    }
    let mut depth = 0_usize;
    let mut position = group_start;
    while let Some(&byte) = bytes.get(position) {
        if byte == b'\\' {
            if matches!(bytes.get(position + 1), Some(b'\\' | b'{' | b'}')) {
                position = position.saturating_add(2);
                continue;
            }
        } else if byte == b'{' {
            depth = depth
                .checked_add(1)
                .ok_or_else(|| rtf_error("RTF picture nesting overflows"))?;
        } else if byte == b'}' {
            depth = depth
                .checked_sub(1)
                .ok_or_else(|| rtf_error("RTF picture has an unmatched closing brace"))?;
            if depth == 0 {
                return Ok(position);
            }
        }
        position = position.saturating_add(1);
    }
    Err(rtf_error("RTF picture group is truncated"))
}

fn rtf_picture_is_fallback(bytes: &[u8], group_offsets: &[usize], _: usize) -> bool {
    group_offsets
        .iter()
        .rev()
        .skip(1)
        .any(|group_start| rtf_group_starts_with_control(bytes, *group_start, b"nonshppict"))
}

fn rtf_group_starts_with_control(bytes: &[u8], group_start: usize, control: &[u8]) -> bool {
    let mut position = group_start.saturating_add(1);
    while bytes.get(position).is_some_and(u8::is_ascii_whitespace) {
        position += 1;
    }
    if bytes.get(position..position.saturating_add(2)) == Some(b"\\*") {
        position += 2;
        while bytes.get(position).is_some_and(u8::is_ascii_whitespace) {
            position += 1;
        }
    }
    bytes.get(position) == Some(&b'\\')
        && bytes
            .get(position + 1..position + 1 + control.len())
            .is_some_and(|candidate| candidate.eq_ignore_ascii_case(control))
        && !bytes
            .get(position + 1 + control.len())
            .is_some_and(u8::is_ascii_alphabetic)
}

fn find_rtf_control(bytes: &[u8], control: &[u8]) -> Option<usize> {
    bytes
        .windows(control.len().saturating_add(1))
        .position(|window| {
            window.first() == Some(&b'\\')
                && window
                    .get(1..)
                    .is_some_and(|candidate| candidate.eq_ignore_ascii_case(control))
        })
}

fn rtf_control_end(bytes: &[u8], offset: usize, limit: usize) -> Result<usize, Diagnostic> {
    if bytes.get(offset) != Some(&b'\\') {
        return Err(rtf_error("RTF picture format control is malformed"));
    }
    let mut position = offset + 1;
    while position < limit && bytes.get(position).is_some_and(u8::is_ascii_alphabetic) {
        position += 1;
    }
    if bytes.get(position) == Some(&b'-') {
        position += 1;
    }
    while position < limit && bytes.get(position).is_some_and(u8::is_ascii_digit) {
        position += 1;
    }
    if bytes.get(position) == Some(&b' ') {
        position += 1;
    }
    Ok(position)
}

fn decode_rtf_picture_hex(
    bytes: &[u8],
    start: usize,
    end: usize,
    limits: Limits,
) -> Result<Vec<u8>, Diagnostic> {
    let encoded_length = end.saturating_sub(start);
    if encoded_length / 2 > limits.max_entry_uncompressed_bytes {
        return Err(Diagnostic::fatal(
            DiagnosticCode::InputTooLarge,
            Phase::Parse,
            Some(start),
            "RTF picture exceeds the configured byte budget",
        )
        .in_part("input.rtf"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact((encoded_length / 2).min(limits.max_entry_uncompressed_bytes))
        .map_err(|_| {
            Diagnostic::fatal(
                DiagnosticCode::AllocationFailed,
                Phase::Parse,
                Some(start),
                "unable to allocate the decoded RTF picture",
            )
            .in_part("input.rtf")
        })?;
    let mut high = None;
    let mut position = start;
    let mut started = false;
    while position < end {
        let byte = bytes[position];
        if !started && byte == b'\\' {
            position = rtf_control_end(bytes, position, end)?;
            continue;
        }
        if byte.is_ascii_whitespace() || matches!(byte, b'{' | b'}') {
            position += 1;
            continue;
        }
        let Some(nibble) = hex_nibble(byte) else {
            return Err(rtf_error("RTF picture contains malformed hexadecimal data"));
        };
        started = true;
        if let Some(previous) = high.take() {
            output.push((previous << 4) | nibble);
            if output.len() > limits.max_entry_uncompressed_bytes {
                return Err(Diagnostic::fatal(
                    DiagnosticCode::InputTooLarge,
                    Phase::Parse,
                    Some(start),
                    "RTF picture exceeds the configured byte budget",
                )
                .in_part("input.rtf"));
            }
        } else {
            high = Some(nibble);
        }
        position += 1;
    }
    if high.is_some() || output.is_empty() {
        return Err(rtf_error("RTF picture hexadecimal data is truncated"));
    }
    Ok(output)
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn rtf_control_parameter(bytes: &[u8], control: &[u8]) -> Option<i32> {
    let mut result = None;
    let mut search_start = 0;
    while let Some(relative) = find_rtf_control(&bytes[search_start..], control) {
        let offset = search_start + relative + 1 + control.len();
        let mut end = offset;
        if bytes.get(end) == Some(&b'-') {
            end += 1;
        }
        while bytes.get(end).is_some_and(u8::is_ascii_digit) {
            end += 1;
        }
        if end > offset {
            result = std::str::from_utf8(&bytes[offset..end])
                .ok()
                .and_then(|value| value.parse::<i32>().ok());
        }
        search_start = offset;
    }
    result
}

fn decode_ansi_byte(byte: u8, code_page: u32) -> char {
    if byte < 0x80 || code_page != 1252 {
        return char::from(byte);
    }
    const WINDOWS_1252: [char; 32] = [
        '\u{20ac}', '\u{0081}', '\u{201a}', '\u{0192}', '\u{201e}', '\u{2026}', '\u{2020}',
        '\u{2021}', '\u{02c6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{008d}',
        '\u{017d}', '\u{008f}', '\u{0090}', '\u{2018}', '\u{2019}', '\u{201c}', '\u{201d}',
        '\u{2022}', '\u{2013}', '\u{2014}', '\u{02dc}', '\u{2122}', '\u{0161}', '\u{203a}',
        '\u{0153}', '\u{009d}', '\u{017e}', '\u{0178}',
    ];
    if byte < 0xa0 {
        WINDOWS_1252[(byte - 0x80) as usize]
    } else {
        char::from_u32(u32::from(byte)).unwrap_or('\u{fffd}')
    }
}

fn rtf_error(message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message)
        .in_part("input.rtf")
}

fn rtf_line_height(block: &RichTextBlock, font_metrics: &FontMetricTable) -> f32 {
    let maximum_font_size = block
        .runs
        .iter()
        .map(|run| run.font_size)
        .reduce(f32::max)
        .unwrap_or(FONT_SIZE);
    let natural = block
        .runs
        .iter()
        .map(|run| {
            run.font_size
                * font_metrics
                    .line_height_em(&run.font_family, run.italic, run.bold, 1.2)
                    .unwrap_or(1.2)
        })
        .reduce(f32::max)
        .unwrap_or(FONT_SIZE * 1.2);
    let line_height = match block.style.line_spacing {
        Some(spacing) if block.text.is_empty() && block.style.line_spacing_multiple => {
            maximum_font_size * (spacing.unsigned_abs() as f32 / 240.0)
        }
        Some(spacing) if block.style.line_spacing_multiple => {
            natural * (spacing.unsigned_abs() as f32 / 240.0)
        }
        Some(spacing) if spacing < 0 => twips(spacing.saturating_abs()),
        Some(spacing) => twips(spacing).max(natural),
        None => natural,
    };
    line_height.max(1.0)
}

fn rtf_table_row_height(row: &RtfTableRow, font_metrics: &FontMetricTable) -> f32 {
    let mut left = row.left;
    let mut height = row.minimum_height.max(1.0);
    for (cell, right) in row.cells.iter().zip(&row.cell_boundaries) {
        let width = (*right - left - twips(110)).max(1.0);
        let text_height = rich_text_line_count(&cell.text, width, &cell.runs, font_metrics) as f32
            * rtf_line_height(cell, font_metrics);
        height = height
            .max(cell.style.space_before + text_height + cell.style.space_after + twips(51 + 55));
        left = *right;
    }
    height
}

fn layout_rtf(
    paragraphs: Vec<RichTextBlock>,
    footnotes: Vec<RichTextBlock>,
    diagnostics: Vec<Diagnostic>,
    page: RtfPageSpec,
    limits: Limits,
    font_metrics: &FontMetricTable,
) -> Result<Document, Diagnostic> {
    if paragraphs.len() > limits.max_document_objects {
        return Err(object_limit());
    }
    if ![
        page.width,
        page.height,
        page.margin_left,
        page.margin_right,
        page.margin_top,
        page.margin_bottom,
    ]
    .into_iter()
    .all(f32::is_finite)
        || page.width <= page.margin_left + page.margin_right
        || page.height <= page.margin_top + page.margin_bottom
    {
        return Err(rtf_error("RTF page size or margins are invalid"));
    }
    let content_width = page.width - page.margin_left - page.margin_right;
    let content_bottom = page.height - page.margin_bottom;
    let mut document = Document {
        fatal: false,
        format: Some(DocumentFormat::Rtf),
        kind: Some(DocumentKind::Text),
        units: vec![page_unit(0, page)],
        outline: Vec::new(),
        objects: Vec::new(),
        embedded_fonts: Vec::new(),
        font_alternate_names: Vec::new(),
        diagnostics,
    };
    let mut unit_index = 0_u32;
    let mut y = page.margin_top;
    let mut materialized_text_bytes = 0_usize;
    let mut picture_index = 0_u32;
    let mut table_row_index = 0_u32;
    for (paragraph_index, paragraph) in paragraphs.into_iter().enumerate() {
        let paragraph_index = u32::try_from(paragraph_index).map_err(|_| object_limit())?;
        let line_height = rtf_line_height(&paragraph, font_metrics);
        let paragraph_width =
            (content_width - paragraph.style.left_indent - paragraph.style.right_indent).max(1.0);
        let line_count = if paragraph.table_row.is_none() && paragraph.picture.is_none() {
            rich_text_line_count(
                &paragraph.text,
                paragraph_width,
                &paragraph.runs,
                font_metrics,
            )
        } else {
            0
        };
        let content_height = if let Some(row) = &paragraph.table_row {
            rtf_table_row_height(row, font_metrics)
        } else if let Some(picture) = &paragraph.picture {
            picture.offset_y.max(0.0) + picture.height
        } else {
            line_count as f32 * line_height
        };
        let mut pagination_height =
            paragraph.style.space_before + content_height + paragraph.style.space_after;
        pagination_height += paragraph
            .picture
            .as_ref()
            .filter(|picture| picture.floating)
            .map_or(0.0, |_| line_height);
        if pagination_height > content_bottom - page.margin_top {
            return Err(Diagnostic::fatal(
                DiagnosticCode::LayoutBudgetExceeded,
                Phase::Layout,
                None,
                "one RTF block exceeds the supported page layout height",
            )
            .in_part("input.rtf"));
        }
        if y + pagination_height > content_bottom && y > page.margin_top {
            push_rtf_page(&mut document, &mut unit_index, page, limits)?;
            y = page.margin_top;
        }
        let starts_at_page_top = (y - page.margin_top).abs() < f32::EPSILON;
        y += paragraph.style.space_before;

        if let Some(row) = paragraph.table_row {
            let mut cell_left = row.left;
            for (column, (cell, cell_right)) in
                row.cells.into_iter().zip(row.cell_boundaries).enumerate()
            {
                let column = u32::try_from(column).map_err(|_| object_limit())?;
                let cell_width = (cell_right - cell_left).max(1.0);
                let cell_line_height = rtf_line_height(&cell, font_metrics);
                let text_length =
                    u32::try_from(cell.text.chars().count()).map_err(|_| object_limit())?;
                reserve_materialized_text_bytes(
                    &mut materialized_text_bytes,
                    cell.text.len(),
                    2,
                    limits.max_total_uncompressed_bytes,
                    "input.rtf",
                )?;
                if document.objects.len() >= limits.max_document_objects {
                    return Err(object_limit());
                }
                let numeric_id =
                    u32::try_from(document.objects.len()).map_err(|_| object_limit())?;
                document.objects.push(Object {
                    numeric_id,
                    parent_numeric_id: None,
                    stable_id: format!("rtf:table-row:{table_row_index}:cell:{column}"),
                    parent_stable_id: None,
                    kind: ObjectKind::Cell,
                    unit_index,
                    bounds: Rect {
                        x: page.margin_left + cell_left,
                        y,
                        width: cell_width,
                        height: content_height,
                    },
                    z: i32::try_from(paragraph_index).unwrap_or(i32::MAX),
                    text: Some(cell.text.clone()),
                    source: SourceRef {
                        part: "input.rtf".to_owned(),
                        mapping: MappingQuality::Exact,
                        locator: SourceLocator::Flat {
                            kind: "paragraph",
                            index: Some(paragraph_index),
                            row: Some(table_row_index),
                            column: Some(column),
                            text_range: Some((0, text_length)),
                        },
                    },
                    visual: Visual::TextLayout {
                        layout: TextLayout {
                            inset_left: twips(55),
                            inset_right: twips(55),
                            inset_top: twips(51),
                            inset_bottom: twips(55),
                            first_line_indent: cell.style.first_line_indent,
                            ..TextLayout::default()
                        },
                        visual: Box::new(Visual::RichText {
                            geometry: Geometry::Rectangle,
                            fill: Paint::Solid(0xffff_ffff),
                            stroke: Paint::Solid(0x0000_00ff),
                            stroke_width: twips(5),
                            align: cell.style.align,
                            line_height: cell_line_height,
                            runs: cell.runs,
                        }),
                    },
                });
                cell_left = cell_right;
            }
            table_row_index = table_row_index.checked_add(1).ok_or_else(object_limit)?;
            y += content_height + paragraph.style.space_after;
            continue;
        }

        if let Some(picture) = paragraph.picture {
            if document.objects.len() >= limits.max_document_objects {
                return Err(object_limit());
            }
            let numeric_id = u32::try_from(document.objects.len()).map_err(|_| object_limit())?;
            document.objects.push(Object {
                numeric_id,
                parent_numeric_id: None,
                stable_id: format!("rtf:picture:{}", picture.source_offset),
                parent_stable_id: None,
                kind: ObjectKind::Image,
                unit_index,
                bounds: Rect {
                    x: page.margin_left + picture.offset_x,
                    y: y + picture.offset_y
                        + if picture.floating && starts_at_page_top {
                            line_height / 2.0
                        } else {
                            line_height / 3.0
                        },
                    width: picture.width,
                    height: picture.height,
                },
                z: i32::try_from(paragraph_index).unwrap_or(i32::MAX),
                text: None,
                source: SourceRef {
                    part: "input.rtf".to_owned(),
                    mapping: MappingQuality::Exact,
                    locator: SourceLocator::Flat {
                        kind: "picture",
                        index: Some(picture_index),
                        row: None,
                        column: None,
                        text_range: None,
                    },
                },
                visual: Visual::Image {
                    media_type: picture.media_type,
                    bytes: picture.bytes,
                    crop: ImageCrop::default(),
                },
            });
            picture_index = picture_index.checked_add(1).ok_or_else(object_limit)?;
            y += content_height + paragraph.style.space_after;
            continue;
        }

        reserve_materialized_text_bytes(
            &mut materialized_text_bytes,
            paragraph.text.len(),
            2,
            limits.max_total_uncompressed_bytes,
            "input.rtf",
        )?;
        if document.objects.len() >= limits.max_document_objects {
            return Err(object_limit());
        }
        let numeric_id = u32::try_from(document.objects.len()).map_err(|_| object_limit())?;
        let text_length =
            u32::try_from(paragraph.text.chars().count()).map_err(|_| object_limit())?;
        let text = paragraph.text;
        document.objects.push(Object {
            numeric_id,
            parent_numeric_id: None,
            stable_id: format!("rtf:paragraph:{paragraph_index}"),
            parent_stable_id: None,
            kind: ObjectKind::Paragraph,
            unit_index,
            bounds: Rect {
                x: page.margin_left + paragraph.style.left_indent,
                y,
                width: paragraph_width,
                height: content_height,
            },
            z: i32::try_from(paragraph_index).unwrap_or(i32::MAX),
            text: Some(text.clone()),
            source: SourceRef {
                part: "input.rtf".to_owned(),
                mapping: MappingQuality::Exact,
                locator: SourceLocator::Flat {
                    kind: "paragraph",
                    index: Some(paragraph_index),
                    row: None,
                    column: None,
                    text_range: Some((0, text_length)),
                },
            },
            visual: Visual::TextLayout {
                layout: TextLayout {
                    inset_left: 0.0,
                    inset_right: 0.0,
                    inset_top: 0.0,
                    inset_bottom: 0.0,
                    first_line_indent: paragraph.style.first_line_indent,
                    ..TextLayout::default()
                },
                visual: Box::new(Visual::RichText {
                    geometry: Geometry::Rectangle,
                    fill: Paint::None,
                    stroke: Paint::None,
                    stroke_width: 0.0,
                    align: paragraph.style.align,
                    line_height,
                    runs: paragraph.runs,
                }),
            },
        });
        y += content_height + paragraph.style.space_after;
    }
    if !footnotes.is_empty() {
        let heights = footnotes
            .iter()
            .map(|footnote| {
                let line_height = rtf_line_height(footnote, font_metrics);
                let line_count = rich_text_line_count(
                    &footnote.text,
                    content_width,
                    &footnote.runs,
                    font_metrics,
                );
                (line_height, line_count as f32 * line_height)
            })
            .collect::<Vec<_>>();
        let total_height = heights
            .iter()
            .zip(&footnotes)
            .map(|((_, height), footnote)| {
                footnote.style.space_before + height + footnote.style.space_after
            })
            .sum::<f32>();
        let mut footnote_y = content_bottom - total_height;
        if footnote_y < y + 8.0 && y > page.margin_top {
            push_rtf_page(&mut document, &mut unit_index, page, limits)?;
            footnote_y = content_bottom - total_height;
        }
        document.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ApproximateLayout,
                Phase::Layout,
                Fidelity::Approximate,
                "RTF footnotes are kept in a page-bottom note flow with approximate pagination",
            )
            .in_part("input.rtf"),
        );
        for (footnote_index, (footnote, (line_height, height))) in
            footnotes.into_iter().zip(heights).enumerate()
        {
            footnote_y += footnote.style.space_before;
            reserve_materialized_text_bytes(
                &mut materialized_text_bytes,
                footnote.text.len(),
                2,
                limits.max_total_uncompressed_bytes,
                "input.rtf",
            )?;
            if document.objects.len() >= limits.max_document_objects {
                return Err(object_limit());
            }
            let numeric_id = u32::try_from(document.objects.len()).map_err(|_| object_limit())?;
            let text_length =
                u32::try_from(footnote.text.chars().count()).map_err(|_| object_limit())?;
            document.objects.push(Object {
                numeric_id,
                parent_numeric_id: None,
                stable_id: format!("rtf:footnote:{}", footnote_index + 1),
                parent_stable_id: None,
                kind: ObjectKind::Paragraph,
                unit_index,
                bounds: Rect {
                    x: page.margin_left + footnote.style.left_indent,
                    y: footnote_y,
                    width: (content_width
                        - footnote.style.left_indent
                        - footnote.style.right_indent)
                        .max(1.0),
                    height,
                },
                z: i32::MAX,
                text: Some(footnote.text.clone()),
                source: SourceRef {
                    part: "input.rtf".to_owned(),
                    mapping: MappingQuality::Exact,
                    locator: SourceLocator::Flat {
                        kind: "paragraph",
                        index: u32::try_from(footnote_index).ok(),
                        row: None,
                        column: None,
                        text_range: Some((0, text_length)),
                    },
                },
                visual: Visual::TextLayout {
                    layout: TextLayout {
                        first_line_indent: footnote.style.first_line_indent,
                        ..TextLayout::default()
                    },
                    visual: Box::new(Visual::RichText {
                        geometry: Geometry::Rectangle,
                        fill: Paint::None,
                        stroke: Paint::None,
                        stroke_width: 0.0,
                        align: footnote.style.align,
                        line_height,
                        runs: footnote.runs,
                    }),
                },
            });
            footnote_y += height + footnote.style.space_after;
        }
    }
    Ok(document)
}

fn push_rtf_page(
    document: &mut Document,
    unit_index: &mut u32,
    page: RtfPageSpec,
    limits: Limits,
) -> Result<(), Diagnostic> {
    if document.units.len() >= limits.max_document_objects {
        return Err(Diagnostic::fatal(
            DiagnosticCode::LayoutBudgetExceeded,
            Phase::Layout,
            None,
            "RTF pagination exceeds the configured document budget",
        )
        .in_part("input.rtf"));
    }
    *unit_index = unit_index.checked_add(1).ok_or_else(object_limit)?;
    document.units.push(page_unit(*unit_index, page));
    Ok(())
}

fn looks_like_html_input(text: &str) -> bool {
    let prefix = text.trim_start().as_bytes();
    if is_html_doctype(prefix) || prefix.starts_with(b"<!--") {
        return true;
    }
    let prefix = if starts_with_ascii_case_insensitive(prefix, b"<?xml") {
        let Some(end) = prefix.windows(2).position(|window| window == b"?>") else {
            return false;
        };
        trim_ascii_start(&prefix[end + 2..])
    } else {
        prefix
    };
    if is_html_doctype(prefix) || prefix.starts_with(b"<!--") {
        return true;
    }
    let Some(tag) = prefix.strip_prefix(b"<") else {
        return false;
    };
    let name_length = tag
        .iter()
        .position(|byte| !byte.is_ascii_alphanumeric() && *byte != b'-')
        .unwrap_or(tag.len());
    if name_length == 0 || name_length == tag.len() {
        return false;
    }
    let boundary = tag[name_length];
    if !boundary.is_ascii_whitespace() && !matches!(boundary, b'>' | b'/') {
        return false;
    }
    const HTML_TAGS: &str = concat!(
        "a abbr address article aside audio b blockquote body br button canvas caption code ",
        "col colgroup data datalist dd del details dfn dialog div dl dt em fieldset figcaption ",
        "figure footer form h1 h2 h3 h4 h5 h6 head header hgroup hr html i iframe img input ",
        "ins kbd label legend li link main map mark menu meta meter nav noscript object ol ",
        "optgroup option output p picture pre progress q s samp script search section select ",
        "slot small source span strong style sub summary sup svg table tbody td template ",
        "textarea tfoot th thead time title tr u ul var video wbr",
    );
    HTML_TAGS
        .split_ascii_whitespace()
        .any(|candidate| tag[..name_length].eq_ignore_ascii_case(candidate.as_bytes()))
}

fn is_html_doctype(prefix: &[u8]) -> bool {
    let Some(remainder) = strip_ascii_case_insensitive(prefix, b"<!doctype") else {
        return false;
    };
    let remainder = trim_ascii_start(remainder);
    let Some(remainder) = strip_ascii_case_insensitive(remainder, b"html") else {
        return false;
    };
    remainder
        .first()
        .is_none_or(|byte| byte.is_ascii_whitespace() || matches!(byte, b'>' | b'['))
}

fn trim_ascii_start(mut value: &[u8]) -> &[u8] {
    while value.first().is_some_and(u8::is_ascii_whitespace) {
        value = &value[1..];
    }
    value
}

fn strip_ascii_case_insensitive<'a>(value: &'a [u8], prefix: &[u8]) -> Option<&'a [u8]> {
    starts_with_ascii_case_insensitive(value, prefix).then(|| &value[prefix.len()..])
}

fn starts_with_ascii_case_insensitive(value: &[u8], prefix: &[u8]) -> bool {
    value
        .get(..prefix.len())
        .is_some_and(|value| value.eq_ignore_ascii_case(prefix))
}

fn rich_text_line_count(
    text: &str,
    width: f32,
    runs: &[TextRun],
    font_metrics: &FontMetricTable,
) -> usize {
    let width = width.max(1.0);
    let mut lines = 1_usize;
    let mut line_width = 0.0_f32;
    let mut word_width = 0.0_f32;
    let mut consume = |character: char, run: &TextRun| {
        if character == '\n' {
            place_rtf_word(&mut lines, &mut line_width, &mut word_width, width);
            lines += 1;
            line_width = 0.0;
            return;
        }
        let advance = font_metrics
            .advance_em_at_size(
                &run.font_family,
                run.italic,
                run.bold,
                character,
                run.font_size,
            )
            .map(|advance| advance * run.font_size)
            .unwrap_or_else(|| {
                let factor = if character.is_whitespace() {
                    0.28
                } else if character.is_ascii_uppercase() {
                    0.6
                } else if character.is_ascii_punctuation() {
                    0.32
                } else if character.is_ascii() {
                    0.5
                } else {
                    1.0
                };
                run.font_size * factor
            })
            .max(0.0);
        if character.is_whitespace() {
            place_rtf_word(&mut lines, &mut line_width, &mut word_width, width);
            if line_width > 0.0 && line_width + advance > width {
                lines += 1;
                line_width = 0.0;
            } else {
                line_width += advance;
            }
        } else {
            word_width += advance;
        }
    };
    if runs.is_empty() {
        let fallback = TextRun {
            paint: None,
            east_asian_line_breaks: true,
            text: String::new(),
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
            horizontal_scale: 1.0,
        };
        for character in text.chars() {
            consume(character, &fallback);
        }
    } else {
        for run in runs {
            for character in run.text.chars() {
                consume(character, run);
            }
        }
    }
    place_rtf_word(&mut lines, &mut line_width, &mut word_width, width);
    lines.max(1)
}

fn place_rtf_word(lines: &mut usize, line_width: &mut f32, word_width: &mut f32, width: f32) {
    if *word_width <= 0.0 {
        return;
    }
    if *line_width > 0.0 && *line_width + *word_width > width {
        *lines += 1;
        *line_width = 0.0;
    }
    *line_width += *word_width;
    *word_width = 0.0;
}

fn parse_csv_records(text: &str, limits: Limits) -> Result<Vec<Vec<String>>, Diagnostic> {
    let mut rows = Vec::new();
    let mut row = Vec::new();
    let mut field = String::new();
    let mut characters = text.chars().peekable();
    let mut quoted = false;
    let mut after_quote = false;
    let mut at_field_start = true;
    let mut source_objects = 0_usize;
    while let Some(character) = characters.next() {
        if quoted {
            if character == '"' {
                if characters.peek() == Some(&'"') {
                    characters.next();
                    field.push('"');
                } else {
                    quoted = false;
                    after_quote = true;
                }
            } else if character == '\r' {
                if characters.peek() == Some(&'\n') {
                    characters.next();
                }
                field.push('\n');
            } else {
                field.push(character);
            }
            continue;
        }
        if after_quote && !matches!(character, ',' | '\r' | '\n') {
            return Err(csv_error(
                "characters after a closing CSV quote are not permitted",
            ));
        }
        match character {
            '"' if at_field_start => {
                quoted = true;
                at_field_start = false;
            }
            '"' => return Err(csv_error("a CSV quote must begin at the start of a field")),
            ',' => {
                push_csv_field(&mut row, &mut field, &mut source_objects, limits)?;
                at_field_start = true;
                after_quote = false;
            }
            '\r' | '\n' => {
                if character == '\r' && characters.peek() == Some(&'\n') {
                    characters.next();
                }
                push_csv_field(&mut row, &mut field, &mut source_objects, limits)?;
                rows.push(std::mem::take(&mut row));
                at_field_start = true;
                after_quote = false;
            }
            _ => {
                field.push(character);
                at_field_start = false;
            }
        }
    }
    if quoted {
        return Err(csv_error("CSV input ends inside a quoted field"));
    }
    if !row.is_empty() || !field.is_empty() || !at_field_start || after_quote {
        push_csv_field(&mut row, &mut field, &mut source_objects, limits)?;
        rows.push(row);
    }
    Ok(rows)
}

fn push_csv_field(
    row: &mut Vec<String>,
    field: &mut String,
    source_objects: &mut usize,
    limits: Limits,
) -> Result<(), Diagnostic> {
    *source_objects = source_objects.checked_add(1).ok_or_else(object_limit)?;
    if *source_objects > limits.max_document_objects {
        return Err(object_limit());
    }
    row.push(std::mem::take(field));
    Ok(())
}

fn parse_csv(rows: Vec<Vec<String>>, limits: Limits) -> Result<Document, Diagnostic> {
    let row_count = u32::try_from(rows.len()).map_err(|_| object_limit())?;
    let column_count = rows.iter().map(Vec::len).max().unwrap_or(0);
    let column_count_u32 = u32::try_from(column_count).map_err(|_| object_limit())?;
    let actual_cells = rows.iter().try_fold(0_usize, |total, row| {
        total.checked_add(row.len()).ok_or_else(object_limit)
    })?;
    if actual_cells > limits.max_document_objects {
        return Err(object_limit());
    }
    let column_widths = csv_column_widths(&rows, column_count);
    let row_heights = csv_row_heights(&rows, &column_widths);
    let width = column_widths.iter().copied().sum::<f32>().max(1.0);
    let height = row_heights.iter().copied().sum::<f32>().max(1.0);
    let row_axis = crate::model::SheetAxis::from_sizes(&row_heights, LINE_HEIGHT);
    let column_axis = crate::model::SheetAxis::from_sizes(&column_widths, 64.0);
    let mut document = Document {
        fatal: false,
        format: Some(DocumentFormat::Csv),
        kind: Some(DocumentKind::Spreadsheet),
        units: vec![Unit {
            kind: UnitKind::Sheet,
            index: 0,
            id: "unit:0".to_owned(),
            name: "Sheet1".to_owned(),
            width,
            height,
            rows: row_count,
            columns: column_count_u32,
            frozen_rows: 0,
            frozen_columns: 0,
            frozen_width: 0.0,
            frozen_height: 0.0,
            row_axis,
            column_axis,
            show_grid_lines: true,
            tab_color: None,
            sheet: None,
            slide: None,
        }],
        outline: Vec::new(),
        objects: Vec::new(),
        embedded_fonts: Vec::new(),
        font_alternate_names: Vec::new(),
        diagnostics: Vec::new(),
    };
    let mut materialized_text_bytes = 0_usize;
    let mut y = 0.0;
    for (row_index, row) in rows.into_iter().enumerate() {
        let row_index = u32::try_from(row_index).map_err(|_| object_limit())?;
        let mut x = 0.0;
        for (column_index, text) in row.into_iter().enumerate() {
            let width = column_widths[column_index];
            let height = row_heights[row_index as usize];
            reserve_materialized_text_bytes(
                &mut materialized_text_bytes,
                text.len(),
                2,
                limits.max_total_uncompressed_bytes,
                "input.csv",
            )?;
            let text_length = u32::try_from(text.chars().count()).map_err(|_| object_limit())?;
            let column_index_u32 = u32::try_from(column_index).map_err(|_| object_limit())?;
            let numeric_id = u32::try_from(document.objects.len()).map_err(|_| object_limit())?;
            document.objects.push(Object {
                numeric_id,
                parent_numeric_id: None,
                stable_id: format!("csv:cell:{row_index}:{column_index_u32}"),
                parent_stable_id: None,
                kind: ObjectKind::Cell,
                unit_index: 0,
                bounds: Rect {
                    x,
                    y,
                    width,
                    height,
                },
                z: 0,
                text: Some(text.clone()),
                source: SourceRef {
                    part: "input.csv".to_owned(),
                    mapping: MappingQuality::Exact,
                    locator: SourceLocator::Flat {
                        kind: "cell",
                        index: None,
                        row: Some(row_index),
                        column: Some(column_index_u32),
                        text_range: Some((0, text_length)),
                    },
                },
                visual: csv_cell_visual(text),
            });
            x += width;
        }
        y += row_heights[row_index as usize];
    }
    Ok(document)
}

fn csv_column_widths(rows: &[Vec<String>], column_count: usize) -> Vec<f32> {
    (0..column_count)
        .map(|column| {
            let characters = rows
                .iter()
                .filter_map(|row| row.get(column))
                .flat_map(|value| value.split('\n'))
                .map(|line| line.chars().count())
                .max()
                .unwrap_or(0);
            (characters as f32 * APPROXIMATE_CHARACTER_WIDTH + 16.0).clamp(64.0, 320.0)
        })
        .collect()
}

fn csv_row_heights(rows: &[Vec<String>], column_widths: &[f32]) -> Vec<f32> {
    rows.iter()
        .map(|row| {
            row.iter()
                .enumerate()
                .map(|(column, value)| {
                    let characters_per_line = ((column_widths[column] - 16.0)
                        / APPROXIMATE_CHARACTER_WIDTH)
                        .floor()
                        .max(1.0) as usize;
                    value
                        .split('\n')
                        .map(|line| line.chars().count().max(1).div_ceil(characters_per_line))
                        .sum::<usize>()
                })
                .max()
                .unwrap_or(1)
                .max(1) as f32
                * LINE_HEIGHT
        })
        .collect()
}

fn page_unit(index: u32, page: RtfPageSpec) -> Unit {
    Unit {
        kind: UnitKind::Page,
        index,
        id: format!("unit:{index}"),
        name: format!("Page {}", index + 1),
        width: page.width,
        height: page.height,
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

fn csv_cell_visual(text: String) -> Visual {
    Visual::RichText {
        geometry: Geometry::Rectangle,
        fill: Paint::Solid(0xffff_ffff),
        stroke: Paint::Solid(0xd0d0_d0ff),
        stroke_width: 1.0,
        align: TextAlign::Start,
        line_height: LINE_HEIGHT,
        runs: vec![TextRun {
            paint: None,
            east_asian_line_breaks: true,
            text,
            font_family: "Arial".to_owned(),
            font_size: 14.0,
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
    }
}

fn csv_error(message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message)
        .in_part("input.csv")
}

fn format_error(message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message).in_part("input")
}

fn object_limit() -> Diagnostic {
    Diagnostic::fatal(
        DiagnosticCode::ObjectLimit,
        Phase::Parse,
        None,
        "flat document exceeds the configured object limit",
    )
    .in_part("input")
}

#[cfg(test)]
mod tests {
    use super::{RichTextBlock, RtfParagraphStyle, detect_and_parse, parse_rtf, rtf_line_height};
    use crate::font_metrics::FontMetricTable;
    use crate::limits::Limits;
    use crate::model::{SourceLocator, TextRun};

    #[test]
    fn rejects_html_documents_and_common_fragments() {
        let html_inputs: &[&[u8]] = &[
            b"<!doctype html><html><body>document</body></html>",
            b"\xef\xbb\xbf \n<HTML><BODY>upper-case document</BODY></HTML>",
            b"<?xml version=\"1.0\"?><html><body>XHTML document</body></html>",
            b"  <BODY class=\"page\">body fragment</BODY>",
            b"<p>paragraph fragment</p>",
            b"<div data-kind=\"card\">division fragment</div>",
            b"<span>inline fragment</span>",
            b"<strong>inline formatting fragment</strong>",
            b"<h1>heading fragment</h1>",
            b"<table><tr><td>cell</td></tr></table>",
            b"<!-- generated fragment --><section>section fragment</section>",
            b"<div>first,second</div>\n<span>third,fourth</span>",
        ];

        for bytes in html_inputs {
            let parsed = detect_and_parse(bytes, Limits::default(), &FontMetricTable::default())
                .expect("HTML rejection should not be a parse failure");
            assert!(
                parsed.is_none(),
                "HTML bytes must remain unsupported: {}",
                String::from_utf8_lossy(bytes)
            );
        }
    }

    #[test]
    fn keeps_plain_text_unsupported() {
        let plain_text_inputs: &[&[u8]] = &[
            b"Plain title\r\nsecond line",
            b"\xef\xbb\xbfUTF-8 text with a byte-order mark",
            &[0xff, 0xfe, b'H', 0, b'i', 0],
            &[0xfe, 0xff, 0, b'H', 0, b'i'],
            b"2 < 3 and 5 > 4",
            b"<proposal> is a domain-specific marker",
            b"<htmlbook> is not an HTML element",
        ];

        for bytes in plain_text_inputs {
            let parsed = detect_and_parse(bytes, Limits::default(), &FontMetricTable::default())
                .expect("plain-text rejection should not be a parse failure");
            assert!(parsed.is_none(), "plain text must remain unsupported");
        }
    }

    #[test]
    fn rejects_pdf_bytes_that_resemble_csv() {
        let parsed = detect_and_parse(
            b"%PDF-1.7,header\n1,2",
            Limits::default(),
            &FontMetricTable::default(),
        )
        .expect("PDF rejection should not be a parse failure");

        assert!(parsed.is_none(), "PDF bytes must remain unsupported");
    }

    #[test]
    fn applies_rtf_multiple_line_spacing_to_the_natural_font_line_box() {
        let run = TextRun {
            paint: None,
            east_asian_line_breaks: true,
            text: "Text".to_owned(),
            font_family: "Arial".to_owned(),
            font_size: 14.0,
            color: 0x0000_00ff,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            highlight: 0,
            baseline_shift: 0.0,
            letter_spacing: 0.0,
            horizontal_scale: 1.0,
        };
        let multiple = RichTextBlock {
            text: "Text".to_owned(),
            runs: vec![run.clone()],
            style: RtfParagraphStyle {
                line_spacing: Some(288),
                line_spacing_multiple: true,
                ..RtfParagraphStyle::default()
            },
            ..RichTextBlock::default()
        };
        let at_least = RichTextBlock {
            text: "Text".to_owned(),
            runs: vec![run],
            style: RtfParagraphStyle {
                line_spacing: Some(288),
                ..RtfParagraphStyle::default()
            },
            ..RichTextBlock::default()
        };

        let metrics = FontMetricTable::default();
        assert!((rtf_line_height(&multiple, &metrics) - 20.16).abs() < 0.001);
        assert!((rtf_line_height(&at_least, &metrics) - 19.2).abs() < 0.001);
    }

    #[test]
    fn escaped_rtf_newlines_use_paragraph_semantics_in_each_destination() {
        let source = concat!(
            "{\\rtf1{\\fonttbl{\\f0 Arial;}}",
            "\\fs48 Title\\par \\par \\fs24 Body\r\n text",
            "{\\*\\unknown hidden\\par ignored}",
            "\\chftn{\\footnote Note\\par continued}\\par ",
            "\\trowd\\cellx2000\\intbl Cell\\par continued\\cell\\row}"
        );
        let parse = |source: &str| {
            parse_rtf(
                source.as_bytes(),
                Limits::default(),
                &FontMetricTable::default(),
            )
            .unwrap()
        };
        let expected = parse(source);
        assert_eq!(expected.objects[0].text.as_deref(), Some("Title"));
        assert_eq!(expected.objects[1].text.as_deref(), Some(""));
        assert_eq!(expected.objects[2].text.as_deref(), Some("Body text1"));
        for newline in ["\\\n", "\\\r", "\\\r\n"] {
            let actual = parse(&source.replace("\\par ", newline));
            assert_eq!(actual.objects, expected.objects);
        }
    }

    #[test]
    fn renders_automatic_rtf_footnote_markers_without_merging_the_note_into_the_body() {
        let document = parse_rtf(
            br"{\rtf1 Body\super\chftn{\footnote\pard\plain\fi-720\li720\super\chftn{} \nosupersub Note text\par}}",
            Limits::default(),
            &FontMetricTable::default(),
        )
        .unwrap();

        assert_eq!(document.units.len(), 1);
        assert_eq!(document.objects[0].text.as_deref(), Some("Body1"));
        assert_eq!(document.objects[1].text.as_deref(), Some("1 Note text"));
        assert!(matches!(
            document.objects[1].source.locator,
            SourceLocator::Flat {
                kind: "paragraph",
                ..
            }
        ));
        assert!(document.objects[1].bounds.y > document.objects[0].bounds.y);
    }
}
