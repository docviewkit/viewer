//! Bounded, namespace-preserving XML tokenizer for OOXML and ODF parts.
//!
//! It intentionally has no DTD or general-entity support. The five predefined
//! XML entities and numeric character references are accepted without expansion.

use std::borrow::Cow;

use crate::diagnostic::{Diagnostic, DiagnosticCode, Phase};
use crate::limits::Limits;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct XmlAttribute<'a> {
    pub name: &'a str,
    /// Raw XML value. Use [`decode_xml_text`] when decoded text is required.
    pub value: &'a str,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum XmlEvent<'a> {
    StartElement {
        name: &'a str,
        attributes: Vec<XmlAttribute<'a>>,
        empty: bool,
    },
    EndElement {
        name: &'a str,
    },
    Text(&'a str),
    Cdata(&'a str),
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct XmlSummary {
    pub elements: usize,
    pub text_nodes: usize,
    pub max_depth: usize,
}

/// Validates an XML part and returns bounded structural counts.
pub fn validate_xml(input: &[u8], limits: Limits) -> Result<XmlSummary, Diagnostic> {
    parse_xml(input, limits, |_| Ok(()))
}

/// Tokenizes an XML part. Events are valid only during the callback. Namespace
/// prefixes are preserved; namespace resolution belongs to the format parser.
pub fn parse_xml<F>(input: &[u8], limits: Limits, mut on_event: F) -> Result<XmlSummary, Diagnostic>
where
    F: for<'a> FnMut(XmlEvent<'a>) -> Result<(), Diagnostic>,
{
    let input = xml_utf8_input(input, limits)?;
    parse_xml_shared(&input, limits, &mut on_event)
}

/// Select MCE alternatives before format adapters consume the XML. Wrappers are
/// removed, so every consumer (including text-box extraction) sees the same tree.
pub fn parse_ooxml<F>(input: &[u8], limits: Limits, on_event: F) -> Result<XmlSummary, Diagnostic>
where
    F: for<'a> FnMut(XmlEvent<'a>) -> Result<(), Diagnostic>,
{
    parse_ooxml_excluding(input, limits, &[], on_event)
}

/// Some format adapters cannot consume every namespace supported by Core.
pub fn parse_ooxml_excluding<F>(
    input: &[u8],
    limits: Limits,
    excluded: &[&str],
    mut on_event: F,
) -> Result<XmlSummary, Diagnostic>
where
    F: for<'a> FnMut(XmlEvent<'a>) -> Result<(), Diagnostic>,
{
    let input = xml_utf8_input(input, limits)?;
    parse_ooxml_shared(&input, limits, excluded, &mut on_event)
}

fn parse_ooxml_shared<'a>(
    input: &'a [u8],
    limits: Limits,
    excluded: &[&str],
    on_event: &mut dyn FnMut(XmlEvent<'a>) -> Result<(), Diagnostic>,
) -> Result<XmlSummary, Diagnostic> {
    struct Frame {
        namespaces: usize,
        emit: bool,
        hidden: bool,
        chosen: Option<bool>,
    }
    let mut namespaces = Vec::<(&str, Cow<'_, str>)>::new();
    let mut frames = Vec::<Frame>::new();
    parse_xml_shared(input, limits, &mut |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let saved = namespaces.len();
                for attribute in &attributes {
                    if let Some(prefix) = attribute.name.strip_prefix("xmlns:") {
                        namespaces.push((prefix, decode_xml_text(attribute.value)?));
                    } else if attribute.name == "xmlns" {
                        namespaces.push(("", decode_xml_text(attribute.value)?));
                    }
                }
                let resolve = |prefix: &str| {
                    namespaces
                        .iter()
                        .rev()
                        .find(|(key, _)| *key == prefix)
                        .map(|(_, value)| value.as_ref())
                };
                let (prefix, local) = name.split_once(':').unwrap_or(("", name));
                let mce = resolve(prefix)
                    == Some("http://schemas.openxmlformats.org/markup-compatibility/2006");
                let mut hidden = frames.last().is_some_and(|frame| frame.hidden);
                let wrapper = mce && matches!(local, "AlternateContent" | "Choice" | "Fallback");
                if mce && matches!(local, "Choice" | "Fallback") && !hidden {
                    if let Some(chosen) = frames.last_mut().and_then(|frame| frame.chosen.as_mut())
                    {
                        let supported = if local == "Fallback" {
                            true
                        } else {
                            attributes
                                .iter()
                                .find(|a| a.name == "Requires")
                                .map(|a| decode_xml_text(a.value))
                                .transpose()?
                                .is_some_and(|value| {
                                    !value.trim().is_empty()
                                        && value.split_whitespace().all(|prefix| {
                                            resolve(prefix).is_some_and(|uri| {
                                                ooxml_namespace_supported(uri)
                                                    && !excluded.contains(&uri)
                                            })
                                        })
                                })
                        };
                        hidden = *chosen || !supported;
                        *chosen |= supported;
                    }
                }
                let emit = !hidden && !wrapper;
                if !empty {
                    frames.push(Frame {
                        namespaces: saved,
                        emit,
                        hidden,
                        chosen: (mce && local == "AlternateContent").then_some(false),
                    });
                }
                if emit {
                    on_event(XmlEvent::StartElement {
                        name,
                        attributes,
                        empty,
                    })?;
                }
                if empty {
                    namespaces.truncate(saved);
                }
            }
            XmlEvent::EndElement { name } => {
                if let Some(frame) = frames.pop() {
                    namespaces.truncate(frame.namespaces);
                    if frame.emit {
                        on_event(XmlEvent::EndElement { name })?;
                    }
                }
            }
            event => {
                if !frames.last().is_some_and(|frame| frame.hidden) {
                    on_event(event)?;
                }
            }
        }
        Ok(())
    })
}

fn xml_utf8_input<'a>(input: &'a [u8], limits: Limits) -> Result<Cow<'a, [u8]>, Diagnostic> {
    if input.starts_with(&[0x00, 0x00, 0xfe, 0xff])
        || input.starts_with(&[0xff, 0xfe, 0x00, 0x00])
        || input.starts_with(&[0x00, 0x00, 0x00, b'<'])
        || input.starts_with(&[b'<', 0x00, 0x00, 0x00])
    {
        return Err(xml_error(
            DiagnosticCode::XmlEncodingUnsupported,
            Some(0),
            "UTF-32 XML is not supported",
        ));
    }
    let (bytes, little_endian) = if let Some(bytes) = input.strip_prefix(&[0xff, 0xfe]) {
        (bytes, true)
    } else if let Some(bytes) = input.strip_prefix(&[0xfe, 0xff]) {
        (bytes, false)
    } else if input.starts_with(&[b'<', 0x00]) {
        (input, true)
    } else if input.starts_with(&[0x00, b'<']) {
        (input, false)
    } else {
        return Ok(Cow::Borrowed(input));
    };
    limits
        .validate()
        .map_err(|error| xml_error(DiagnosticCode::InvalidLimits, None, error.to_string()))?;
    if input.len() > limits.max_xml_bytes {
        return Err(xml_error(
            DiagnosticCode::XmlSizeLimit,
            None,
            "XML part exceeds the configured byte limit",
        ));
    }
    if !bytes.len().is_multiple_of(2) {
        return Err(xml_error(
            DiagnosticCode::XmlEncodingUnsupported,
            Some(input.len() - 1),
            "UTF-16 XML has a truncated code unit",
        ));
    }
    let units = bytes.chunks_exact(2).map(|pair| {
        if little_endian {
            u16::from_le_bytes([pair[0], pair[1]])
        } else {
            u16::from_be_bytes([pair[0], pair[1]])
        }
    });
    let mut text = String::new();
    for character in char::decode_utf16(units) {
        let character = character.map_err(|_| {
            xml_error(
                DiagnosticCode::XmlEncodingUnsupported,
                None,
                "UTF-16 XML contains an unpaired surrogate",
            )
        })?;
        if text.len() + character.len_utf8() > limits.max_xml_bytes {
            return Err(xml_error(
                DiagnosticCode::XmlSizeLimit,
                None,
                "decoded XML part exceeds the configured byte limit",
            ));
        }
        text.push(character);
    }
    Ok(Cow::Owned(text.into_bytes()))
}

fn ooxml_namespace_supported(uri: &str) -> bool {
    // Only namespaces with native consumers, never arbitrary Office versions.
    matches!(
        uri,
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            | "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            | "http://schemas.openxmlformats.org/presentationml/2006/main"
            | "http://schemas.openxmlformats.org/drawingml/2006/main"
            | "http://schemas.openxmlformats.org/drawingml/2006/chart"
            | "http://schemas.openxmlformats.org/drawingml/2006/diagram"
            | "http://schemas.openxmlformats.org/drawingml/2006/picture"
            | "http://purl.oclc.org/ooxml/wordprocessingml/main"
            | "http://purl.oclc.org/ooxml/spreadsheetml/main"
            | "http://purl.oclc.org/ooxml/presentationml/main"
            | "http://purl.oclc.org/ooxml/drawingml/main"
            | "http://purl.oclc.org/ooxml/drawingml/chart"
            | "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            | "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            | "http://schemas.microsoft.com/office/word/2010/wordml"
            | "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
            | "http://schemas.microsoft.com/office/powerpoint/2010/main"
            | "http://schemas.microsoft.com/office/drawing/2010/main"
            | "http://schemas.microsoft.com/office/drawing/2007/8/2/chart"
            | "http://schemas.microsoft.com/office/drawing/2014/chartex"
            | "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
            | "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
            | "urn:schemas-microsoft-com:vml"
    )
}

// Share the tokenizer across format callbacks without changing event ownership or errors.
#[inline(never)]
fn parse_xml_shared<'a>(
    input: &'a [u8],
    limits: Limits,
    on_event: &mut dyn FnMut(XmlEvent<'a>) -> Result<(), Diagnostic>,
) -> Result<XmlSummary, Diagnostic> {
    limits
        .validate()
        .map_err(|error| xml_error(DiagnosticCode::InvalidLimits, None, error.to_string()))?;
    if input.len() > limits.max_xml_bytes {
        return Err(xml_error(
            DiagnosticCode::XmlSizeLimit,
            None,
            "XML part exceeds the configured byte limit",
        ));
    }
    let bom_bytes = usize::from(input.starts_with(&[0xef, 0xbb, 0xbf])) * 3;
    let text = std::str::from_utf8(&input[bom_bytes..]).map_err(|error| {
        xml_error(
            DiagnosticCode::XmlEncodingUnsupported,
            Some(bom_bytes + error.valid_up_to()),
            "XML is not valid UTF-8",
        )
    })?;
    let mut parser = Parser {
        text,
        source_offset: bom_bytes,
        cursor: 0,
        limits,
        stack: Vec::new(),
        root_seen: false,
        root_closed: false,
        nodes_seen: 0,
        summary: XmlSummary::default(),
    };

    while parser.cursor < parser.text.len() {
        if parser.remaining().starts_with('<') {
            if parser.remaining().starts_with("<!--") {
                parser.parse_comment()?;
            } else if parser.remaining().starts_with("<![CDATA[") {
                let event = parser.parse_cdata()?;
                on_event(event)?;
            } else if parser.remaining().starts_with("<?") {
                parser.parse_processing_instruction()?;
            } else if parser.remaining().starts_with("</") {
                let event = parser.parse_end_element()?;
                on_event(event)?;
            } else if parser.remaining().starts_with("<!") {
                return Err(parser.error(
                    DiagnosticCode::XmlDtdForbidden,
                    "DTD, ENTITY, and other XML declarations are forbidden",
                ));
            } else {
                let event = parser.parse_start_element()?;
                on_event(event)?;
            }
        } else {
            let event = parser.parse_text()?;
            if let Some(event) = event {
                on_event(event)?;
            }
        }
    }

    if !parser.stack.is_empty() {
        return Err(parser.error(
            DiagnosticCode::XmlInvalid,
            "XML ended before all elements were closed",
        ));
    }
    if !parser.root_seen {
        return Err(parser.error(DiagnosticCode::XmlInvalid, "XML has no root element"));
    }
    Ok(parser.summary)
}

struct Parser<'a> {
    text: &'a str,
    source_offset: usize,
    cursor: usize,
    limits: Limits,
    stack: Vec<&'a str>,
    root_seen: bool,
    root_closed: bool,
    nodes_seen: usize,
    summary: XmlSummary,
}

impl<'a> Parser<'a> {
    fn remaining(&self) -> &'a str {
        &self.text[self.cursor..]
    }

    fn source_position(&self) -> usize {
        self.source_offset + self.cursor
    }

    fn error(&self, code: DiagnosticCode, message: impl Into<String>) -> Diagnostic {
        xml_error(code, Some(self.source_position()), message)
    }

    fn bump_node(&mut self, text_node: bool) -> Result<(), Diagnostic> {
        self.consume_node_budget()?;
        if text_node {
            self.summary.text_nodes += 1;
        } else {
            self.summary.elements += 1;
        }
        Ok(())
    }

    fn consume_node_budget(&mut self) -> Result<(), Diagnostic> {
        self.nodes_seen = self
            .nodes_seen
            .checked_add(1)
            .ok_or_else(|| self.error(DiagnosticCode::XmlNodeLimit, "XML node count overflow"))?;
        if self.nodes_seen > self.limits.max_xml_nodes {
            return Err(self.error(
                DiagnosticCode::XmlNodeLimit,
                "XML exceeds the configured node limit",
            ));
        }
        Ok(())
    }

    fn parse_comment(&mut self) -> Result<(), Diagnostic> {
        let body_start = self.cursor + 4;
        let relative_end = self.text[body_start..]
            .find("-->")
            .ok_or_else(|| self.error(DiagnosticCode::XmlInvalid, "unterminated XML comment"))?;
        let body_end = body_start + relative_end;
        if self.text[body_start..body_end].contains("--") {
            return Err(self.error(
                DiagnosticCode::XmlInvalid,
                "XML comment contains a forbidden double hyphen",
            ));
        }
        validate_xml_chars(
            &self.text[body_start..body_end],
            self.source_offset + body_start,
        )?;
        self.cursor = body_end + 3;
        self.consume_node_budget()?;
        Ok(())
    }

    fn parse_cdata(&mut self) -> Result<XmlEvent<'a>, Diagnostic> {
        if self.stack.is_empty() {
            return Err(self.error(
                DiagnosticCode::XmlInvalid,
                "CDATA is only allowed inside the root element",
            ));
        }
        let body_start = self.cursor + 9;
        let relative_end = self.text[body_start..]
            .find("]]>")
            .ok_or_else(|| self.error(DiagnosticCode::XmlInvalid, "unterminated CDATA section"))?;
        let body_end = body_start + relative_end;
        validate_xml_chars(
            &self.text[body_start..body_end],
            self.source_offset + body_start,
        )?;
        self.cursor = body_end + 3;
        self.bump_node(true)?;
        Ok(XmlEvent::Cdata(&self.text[body_start..body_end]))
    }

    fn parse_processing_instruction(&mut self) -> Result<(), Diagnostic> {
        let body_start = self.cursor + 2;
        let relative_end = self.text[body_start..].find("?>").ok_or_else(|| {
            self.error(
                DiagnosticCode::XmlInvalid,
                "unterminated XML processing instruction",
            )
        })?;
        let body_end = body_start + relative_end;
        let mut target_end = body_start;
        while target_end < body_end && !self.text.as_bytes()[target_end].is_ascii_whitespace() {
            target_end += 1;
        }
        if target_end == body_start {
            return Err(self.error(
                DiagnosticCode::XmlInvalid,
                "processing instruction has no target",
            ));
        }
        validate_name(
            &self.text[body_start..target_end],
            self.limits.max_xml_name_bytes,
            self.source_offset + body_start,
        )?;
        validate_xml_chars(
            &self.text[body_start..body_end],
            self.source_offset + body_start,
        )?;
        self.cursor = body_end + 2;
        self.consume_node_budget()?;
        Ok(())
    }

    fn parse_start_element(&mut self) -> Result<XmlEvent<'a>, Diagnostic> {
        self.cursor += 1;
        let name = self.parse_name()?;
        if self.stack.is_empty() {
            if self.root_seen || self.root_closed {
                return Err(self.error(
                    DiagnosticCode::XmlInvalid,
                    "XML contains more than one root element",
                ));
            }
            self.root_seen = true;
        }

        let mut attributes = Vec::new();
        let mut empty = false;
        loop {
            let had_whitespace = self.skip_whitespace();
            if self.remaining().starts_with("/>") {
                self.cursor += 2;
                empty = true;
                break;
            }
            if self.remaining().starts_with('>') {
                self.cursor += 1;
                break;
            }
            if !had_whitespace {
                return Err(self.error(
                    DiagnosticCode::XmlInvalid,
                    "XML attributes must be separated by whitespace",
                ));
            }
            if attributes.len() >= self.limits.max_xml_attributes_per_element {
                return Err(self.error(
                    DiagnosticCode::XmlAttributeLimit,
                    "element exceeds the configured attribute limit",
                ));
            }
            let attribute_name = self.parse_name()?;
            if attributes
                .iter()
                .any(|attribute: &XmlAttribute<'_>| attribute.name == attribute_name)
            {
                return Err(self.error(
                    DiagnosticCode::XmlInvalid,
                    "element contains a duplicate attribute",
                ));
            }
            self.skip_whitespace();
            if !self.remaining().starts_with('=') {
                return Err(self.error(DiagnosticCode::XmlInvalid, "XML attribute is missing '='"));
            }
            self.cursor += 1;
            self.skip_whitespace();
            let quote = self
                .remaining()
                .as_bytes()
                .first()
                .copied()
                .filter(|byte| *byte == b'\'' || *byte == b'"')
                .ok_or_else(|| {
                    self.error(
                        DiagnosticCode::XmlInvalid,
                        "XML attribute value is not quoted",
                    )
                })?;
            self.cursor += 1;
            let value_start = self.cursor;
            let relative_end = self
                .remaining()
                .as_bytes()
                .iter()
                .position(|byte| *byte == quote)
                .ok_or_else(|| {
                    self.error(
                        DiagnosticCode::XmlInvalid,
                        "unterminated XML attribute value",
                    )
                })?;
            let value_end = value_start + relative_end;
            let value = &self.text[value_start..value_end];
            if value.contains('<') {
                return Err(self.error(
                    DiagnosticCode::XmlInvalid,
                    "XML attribute value contains '<'",
                ));
            }
            validate_xml_chars(value, self.source_offset + value_start)?;
            validate_entity_references(value, self.source_offset + value_start)?;
            attributes.try_reserve(1).map_err(|_| {
                self.error(
                    DiagnosticCode::AllocationFailed,
                    "unable to grow XML attribute list",
                )
            })?;
            attributes.push(XmlAttribute {
                name: attribute_name,
                value,
            });
            self.cursor = value_end + 1;
        }

        let depth = self
            .stack
            .len()
            .checked_add(1)
            .ok_or_else(|| self.error(DiagnosticCode::XmlDepthLimit, "XML depth overflow"))?;
        if depth > self.limits.max_xml_depth {
            return Err(self.error(
                DiagnosticCode::XmlDepthLimit,
                "XML exceeds the configured depth limit",
            ));
        }
        self.summary.max_depth = self.summary.max_depth.max(depth);
        self.bump_node(false)?;
        if empty {
            if self.stack.is_empty() {
                self.root_closed = true;
            }
        } else {
            self.stack.try_reserve(1).map_err(|_| {
                self.error(
                    DiagnosticCode::AllocationFailed,
                    "unable to grow XML element stack",
                )
            })?;
            self.stack.push(name);
        }
        Ok(XmlEvent::StartElement {
            name,
            attributes,
            empty,
        })
    }

    fn parse_end_element(&mut self) -> Result<XmlEvent<'a>, Diagnostic> {
        self.cursor += 2;
        let name = self.parse_name()?;
        self.skip_whitespace();
        if !self.remaining().starts_with('>') {
            return Err(self.error(DiagnosticCode::XmlInvalid, "closing element is missing '>'"));
        }
        self.cursor += 1;
        let expected = self.stack.pop().ok_or_else(|| {
            self.error(
                DiagnosticCode::XmlInvalid,
                "closing element has no matching start element",
            )
        })?;
        if expected != name {
            return Err(self.error(
                DiagnosticCode::XmlInvalid,
                format!("closing element '{name}' does not match '{expected}'"),
            ));
        }
        if self.stack.is_empty() {
            self.root_closed = true;
        }
        Ok(XmlEvent::EndElement { name })
    }

    fn parse_text(&mut self) -> Result<Option<XmlEvent<'a>>, Diagnostic> {
        let start = self.cursor;
        let length = self.remaining().find('<').unwrap_or(self.remaining().len());
        self.cursor += length;
        let value = &self.text[start..self.cursor];
        validate_xml_chars(value, self.source_offset + start)?;
        validate_entity_references(value, self.source_offset + start)?;
        if value.contains("]]>") {
            return Err(self.error(
                DiagnosticCode::XmlInvalid,
                "text contains the reserved CDATA terminator",
            ));
        }
        if self.stack.is_empty() {
            if !value.bytes().all(|byte| byte.is_ascii_whitespace()) {
                return Err(self.error(
                    DiagnosticCode::XmlInvalid,
                    "non-whitespace text appears outside the root element",
                ));
            }
            return Ok(None);
        }
        if value.is_empty() {
            return Ok(None);
        }
        self.bump_node(true)?;
        Ok(Some(XmlEvent::Text(value)))
    }

    fn parse_name(&mut self) -> Result<&'a str, Diagnostic> {
        let start = self.cursor;
        while let Some(&byte) = self.text.as_bytes().get(self.cursor) {
            if byte.is_ascii_whitespace()
                || matches!(byte, b'/' | b'>' | b'=' | b'?' | b'\'' | b'"' | b'<')
            {
                break;
            }
            self.cursor += 1;
        }
        let name = &self.text[start..self.cursor];
        validate_name(
            name,
            self.limits.max_xml_name_bytes,
            self.source_offset + start,
        )?;
        Ok(name)
    }

    fn skip_whitespace(&mut self) -> bool {
        let start = self.cursor;
        while self
            .text
            .as_bytes()
            .get(self.cursor)
            .is_some_and(u8::is_ascii_whitespace)
        {
            self.cursor += 1;
        }
        self.cursor != start
    }
}

fn validate_name(name: &str, max_bytes: usize, offset: usize) -> Result<(), Diagnostic> {
    if name.is_empty() || name.len() > max_bytes {
        return Err(xml_error(
            DiagnosticCode::XmlInvalid,
            Some(offset),
            "XML name is empty or too long",
        ));
    }
    let mut chars = name.chars();
    let Some(first) = chars.next() else {
        return Err(xml_error(
            DiagnosticCode::XmlInvalid,
            Some(offset),
            "XML name is empty or too long",
        ));
    };
    if !is_name_start(first) || !chars.all(is_name_continue) {
        return Err(xml_error(
            DiagnosticCode::XmlInvalid,
            Some(offset),
            "XML name contains invalid characters",
        ));
    }
    Ok(())
}

fn is_name_start(character: char) -> bool {
    character == ':'
        || character == '_'
        || character.is_ascii_alphabetic()
        || (!character.is_ascii() && !character.is_control() && !character.is_whitespace())
}

fn is_name_continue(character: char) -> bool {
    is_name_start(character)
        || character.is_ascii_digit()
        || matches!(character, '-' | '.' | '\u{b7}')
}

fn validate_xml_chars(value: &str, base_offset: usize) -> Result<(), Diagnostic> {
    for (relative, character) in value.char_indices() {
        let code = character as u32;
        let valid = matches!(code, 0x9 | 0xa | 0xd)
            || (0x20..=0xd7ff).contains(&code)
            || (0xe000..=0xfffd).contains(&code)
            || (0x10000..=0x10ffff).contains(&code);
        if !valid {
            return Err(xml_error(
                DiagnosticCode::XmlInvalid,
                Some(base_offset + relative),
                "XML contains a forbidden character",
            ));
        }
    }
    Ok(())
}

fn validate_entity_references(value: &str, base_offset: usize) -> Result<(), Diagnostic> {
    let bytes = value.as_bytes();
    let mut cursor = 0;
    while cursor < bytes.len() {
        if bytes[cursor] != b'&' {
            cursor += 1;
            continue;
        }
        let end = value[cursor + 1..]
            .find(';')
            .map(|relative| cursor + 1 + relative)
            .ok_or_else(|| {
                xml_error(
                    DiagnosticCode::XmlEntityForbidden,
                    Some(base_offset + cursor),
                    "unterminated XML entity reference",
                )
            })?;
        validate_entity(&value[cursor + 1..end], base_offset + cursor)?;
        cursor = end + 1;
    }
    Ok(())
}

fn validate_entity(entity: &str, offset: usize) -> Result<char, Diagnostic> {
    match entity {
        "amp" => Ok('&'),
        "lt" => Ok('<'),
        "gt" => Ok('>'),
        "quot" => Ok('"'),
        "apos" => Ok('\''),
        value if value.starts_with("#x") && value.len() > 2 => {
            let code = u32::from_str_radix(&value[2..], 16).map_err(|_| {
                xml_error(
                    DiagnosticCode::XmlEntityForbidden,
                    Some(offset),
                    "invalid hexadecimal character reference",
                )
            })?;
            validate_character_reference(code, offset)
        }
        value if value.starts_with('#') && value.len() > 1 => {
            let code = value[1..].parse::<u32>().map_err(|_| {
                xml_error(
                    DiagnosticCode::XmlEntityForbidden,
                    Some(offset),
                    "invalid decimal character reference",
                )
            })?;
            validate_character_reference(code, offset)
        }
        _ => Err(xml_error(
            DiagnosticCode::XmlEntityForbidden,
            Some(offset),
            "general XML entities are forbidden",
        )),
    }
}

fn validate_character_reference(code: u32, offset: usize) -> Result<char, Diagnostic> {
    let character = char::from_u32(code).ok_or_else(|| {
        xml_error(
            DiagnosticCode::XmlEntityForbidden,
            Some(offset),
            "character reference is not a Unicode scalar",
        )
    })?;
    validate_xml_chars(character.encode_utf8(&mut [0; 4]), offset)?;
    Ok(character)
}

/// Decodes predefined and numeric XML references without supporting general entities.
pub fn decode_xml_text(value: &str) -> Result<Cow<'_, str>, Diagnostic> {
    if !value.contains('&') {
        return Ok(Cow::Borrowed(value));
    }
    validate_entity_references(value, 0)?;
    let mut output = String::new();
    output.try_reserve(value.len()).map_err(|_| {
        xml_error(
            DiagnosticCode::AllocationFailed,
            None,
            "unable to allocate decoded XML text",
        )
    })?;
    let mut cursor = 0;
    while let Some(relative) = value[cursor..].find('&') {
        let start = cursor + relative;
        output.push_str(&value[cursor..start]);
        let end = value[start + 1..]
            .find(';')
            .map(|relative| start + 1 + relative)
            .ok_or_else(|| {
                xml_error(
                    DiagnosticCode::XmlEntityForbidden,
                    Some(start),
                    "unterminated XML entity reference",
                )
            })?;
        output.push(validate_entity(&value[start + 1..end], start)?);
        cursor = end + 1;
    }
    output.push_str(&value[cursor..]);
    Ok(Cow::Owned(output))
}

fn xml_error(
    code: DiagnosticCode,
    offset: Option<usize>,
    message: impl Into<String>,
) -> Diagnostic {
    Diagnostic::fatal(code, Phase::Xml, offset, message)
}

#[cfg(test)]
mod tests {
    use super::{XmlEvent, decode_xml_text, parse_xml, validate_xml};
    use crate::diagnostic::DiagnosticCode;
    use crate::limits::Limits;

    #[test]
    fn real_ooxml_alternate_content_is_selected_once() {
        let package = crate::package::Package::open(
            include_bytes!("../tests/fixtures/ooxml-nested-alternate.docx"),
            Limits::default(),
        )
        .unwrap();
        let xml = package.required_part("word/document.xml").unwrap();
        let mut text = Vec::new();
        super::parse_ooxml(&xml, Limits::default(), |event| {
            if let XmlEvent::Text(value) = event {
                if matches!(value, "ABC" | "PQR") {
                    text.push(value.to_owned());
                }
            }
            Ok(())
        })
        .unwrap();
        assert_eq!(text, ["ABC", "PQR"]);
    }

    #[test]
    fn real_utf16_ooxml_part_is_tokenized() {
        // customXml/item18.xml from LibreOffice's tdf167689_x15_namespace.xlsx.
        let xml = include_bytes!("../tests/fixtures/ooxml-utf16-custom.xml");
        let mut content = Vec::new();
        super::parse_ooxml(xml, Limits::default(), |event| {
            if let XmlEvent::Cdata(value) = event {
                content.push(value.to_owned());
            }
            Ok(())
        })
        .unwrap();
        assert_eq!(content, ["False"]);

        let mut big_endian = vec![0xfe, 0xff];
        for pair in xml[2..].chunks_exact(2) {
            big_endian.extend_from_slice(&[pair[1], pair[0]]);
        }
        let mut content = Vec::new();
        super::parse_ooxml(&big_endian, Limits::default(), |event| {
            if let XmlEvent::Cdata(value) = event {
                content.push(value.to_owned());
            }
            Ok(())
        })
        .unwrap();
        assert_eq!(content, ["False"]);
        assert_eq!(
            parse_xml(&[0xff, 0xfe, b'<', 0, 0, 0xd8], Limits::default(), |_| Ok(
                ()
            ))
            .unwrap_err()
            .code,
            DiagnosticCode::XmlEncodingUnsupported
        );
    }

    #[test]
    fn ooxml_choices_use_scoped_namespace_uris_and_first_supported_branch() {
        let xml = br#"<root xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:u="urn:unknown">
            <m:AlternateContent><m:Choice Requires="a u"><t>unknown</t></m:Choice>
              <m:Choice Requires="a"><t>first</t><m:AlternateContent>
                <m:Choice xmlns:a="urn:unknown" Requires="a"><t>shadowed</t></m:Choice>
                <m:Fallback><t>fallback</t></m:Fallback></m:AlternateContent></m:Choice>
              <m:Choice Requires="a"><t>second</t></m:Choice><m:Fallback><t>duplicate</t></m:Fallback>
            </m:AlternateContent><t>after</t></root>"#;
        let mut text = Vec::new();
        super::parse_ooxml(xml, Limits::default(), |event| {
            if let XmlEvent::Text(value) = event {
                if !value.trim().is_empty() {
                    text.push(value.to_owned());
                }
            }
            Ok(())
        })
        .unwrap();
        assert_eq!(text, ["first", "fallback", "after"]);
    }

    #[test]
    fn ooxml_choices_respect_adapter_namespace_support() {
        let xml = br#"<root xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml"><m:AlternateContent><m:Choice Requires="v"><t>vml</t></m:Choice><m:Fallback><t>drawingml</t></m:Fallback></m:AlternateContent></root>"#;
        for (excluded, expected) in [
            (&[][..], "vml"),
            (&["urn:schemas-microsoft-com:vml"][..], "drawingml"),
        ] {
            let mut text = String::new();
            super::parse_ooxml_excluding(xml, Limits::default(), excluded, |event| {
                if let XmlEvent::Text(value) = event {
                    text.push_str(value);
                }
                Ok(())
            })
            .unwrap();
            assert_eq!(text, expected);
        }
    }

    #[test]
    fn parses_common_namespaced_ooxml_and_odf() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
            <w:document xmlns:w="urn:word"><w:p w:id="1">A &amp; B</w:p></w:document>"#;
        let mut names = Vec::new();
        let summary = parse_xml(xml, Limits::default(), |event| {
            if let XmlEvent::StartElement { name, .. } = event {
                names.push(name.to_owned());
            }
            Ok(())
        })
        .unwrap();
        assert_eq!(names, ["w:document", "w:p"]);
        assert_eq!(summary.elements, 2);
        assert_eq!(summary.max_depth, 2);

        let odf =
            "<office:document xmlns:office=\"urn:o\"><text:p>中文 Ω</text:p></office:document>";
        validate_xml(odf.as_bytes(), Limits::default()).unwrap();
        assert_eq!(decode_xml_text("A &amp; &#x3a9;").unwrap(), "A & Ω");
    }

    #[test]
    fn rejects_dtd_and_general_entities() {
        let dtd = br#"<!DOCTYPE a [<!ENTITY x SYSTEM "file:///etc/passwd">]><a>&x;</a>"#;
        assert_eq!(
            validate_xml(dtd, Limits::default()).unwrap_err().code,
            DiagnosticCode::XmlDtdForbidden
        );
        assert_eq!(
            validate_xml(b"<a>&custom;</a>", Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::XmlEntityForbidden
        );
    }

    #[test]
    fn preserves_mutable_callback_state_nested_parsing_and_early_errors() {
        let expected = validate_xml(b"<", Limits::default()).unwrap_err();
        let mut calls = 0;
        let actual = parse_xml(b"<root><child/></root>", Limits::default(), |_| {
            calls += 1;
            validate_xml(b"<nested/>", Limits::default())?;
            if calls == 2 {
                Err(expected.clone())
            } else {
                Ok(())
            }
        });
        assert_eq!(actual.unwrap_err(), expected);
        assert_eq!(calls, 2);
    }

    #[test]
    fn enforces_depth_node_and_attribute_limits() {
        let limits = Limits {
            max_xml_depth: 2,
            ..Limits::default()
        };
        assert_eq!(
            validate_xml(b"<a><b><c/></b></a>", limits)
                .unwrap_err()
                .code,
            DiagnosticCode::XmlDepthLimit
        );

        let limits = Limits {
            max_xml_nodes: 2,
            ..Limits::default()
        };
        assert_eq!(
            validate_xml(b"<a><b/><c/></a>", limits).unwrap_err().code,
            DiagnosticCode::XmlNodeLimit
        );

        let limits = Limits {
            max_xml_attributes_per_element: 1,
            ..Limits::default()
        };
        assert_eq!(
            validate_xml(b"<a x='1' y='2'/>", limits).unwrap_err().code,
            DiagnosticCode::XmlAttributeLimit
        );
    }

    #[test]
    fn rejects_malformed_structure_and_accepts_cdata() {
        assert_eq!(
            validate_xml(b"<a><b></a>", Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::XmlInvalid
        );
        validate_xml(
            b"<!--before--><a><![CDATA[x < y && z]]></a><!--after-->",
            Limits::default(),
        )
        .unwrap();
    }
}
