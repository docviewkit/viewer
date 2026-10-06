//! Shared OMML capture. DOCX and formula-only PPTX shapes reuse one geometric
//! layout; mixed or transformed PPTX math keeps the linear-text fallback.

use crate::diagnostic::{Diagnostic, DiagnosticCode, Phase};
use crate::xml::XmlAttribute;

fn format_error(part: &str, message: &str) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message).in_part(part)
}

#[derive(Debug)]
pub(super) struct OmmlNode {
    depth: usize,
    pub(super) name: String,
    pub(super) text: String,
    pub(super) value: Option<String>,
    pub(super) attributes: Vec<(String, String)>,
    pub(super) children: Vec<OmmlNode>,
}

#[derive(Debug)]
pub(super) struct OmmlCapture {
    pub(super) depth: usize,
    stack: Vec<OmmlNode>,
    approximate: bool,
}

#[derive(Debug)]
pub(super) struct OmmlResult {
    pub(super) root: OmmlNode,
    pub(super) text: String,
    pub(super) approximate: bool,
}

impl OmmlCapture {
    pub(super) fn new(depth: usize, local: &str) -> Self {
        Self {
            depth,
            stack: vec![OmmlNode {
                depth,
                name: local.to_owned(),
                text: String::new(),
                value: None,
                attributes: Vec::new(),
                children: Vec::new(),
            }],
            approximate: false,
        }
    }

    pub(super) fn start(
        &mut self,
        local: &str,
        attributes: &[XmlAttribute<'_>],
        depth: usize,
        empty: bool,
        part: &str,
    ) -> Result<(), Diagnostic> {
        if !matches!(
            local,
            "r" | "t"
                | "f"
                | "num"
                | "den"
                | "sSup"
                | "sSub"
                | "sSubSup"
                | "e"
                | "sup"
                | "sub"
                | "rad"
                | "deg"
                | "nary"
                | "naryPr"
                | "chr"
                | "d"
                | "dPr"
                | "begChr"
                | "endChr"
                | "sepChr"
                | "m"
                | "mPr"
                | "mr"
                | "eqArr"
                | "func"
                | "fName"
                | "acc"
                | "accPr"
                | "bar"
                | "barPr"
                | "pos"
                | "limLow"
                | "limUpp"
                | "lim"
                | "groupChr"
                | "groupChrPr"
                | "box"
                | "borderBox"
                | "sPre"
                | "ctrlPr"
                | "rPr"
                | "p"
                | "pPr"
        ) {
            self.approximate = true;
        }
        let node = OmmlNode {
            depth,
            name: local.to_owned(),
            text: String::new(),
            value: super::optional_xml_attribute(attributes, "val", part)?,
            attributes: attributes
                .iter()
                .map(|a| (a.name.to_owned(), a.value.to_owned()))
                .collect(),
            children: Vec::new(),
        };
        if empty {
            self.stack
                .last_mut()
                .ok_or_else(|| format_error(part, "OMML parser stack is empty"))?
                .children
                .push(node);
        } else {
            self.stack.push(node);
        }
        Ok(())
    }

    pub(super) fn text(&mut self, text: &str) {
        if let Some(node) = self.stack.last_mut().filter(|node| node.name == "t") {
            node.text.push_str(text);
        }
    }

    pub(super) fn end(&mut self, depth: usize, part: &str) -> Result<(), Diagnostic> {
        if self.stack.len() <= 1 || self.stack.last().is_none_or(|node| node.depth != depth) {
            return Ok(());
        }
        let node = self
            .stack
            .pop()
            .ok_or_else(|| format_error(part, "OMML parser stack is empty"))?;
        self.stack
            .last_mut()
            .ok_or_else(|| format_error(part, "OMML node has no parent"))?
            .children
            .push(node);
        Ok(())
    }

    pub(super) fn finish(mut self, part: &str) -> Result<OmmlResult, Diagnostic> {
        if self.stack.len() != 1 {
            return Err(format_error(part, "OMML parser ended with unclosed nodes"));
        }
        let root = self
            .stack
            .pop()
            .ok_or_else(|| format_error(part, "OMML parser has no root"))?;
        Ok(OmmlResult {
            text: render_omml_node(&root),
            approximate: self.approximate,
            root,
        })
    }
}

fn omml_child(node: &OmmlNode, name: &str) -> String {
    node.children
        .iter()
        .find(|child| child.name == name)
        .map(render_omml_node)
        .unwrap_or_default()
}

impl OmmlNode {
    pub(super) fn needs_box_layout(&self) -> bool {
        matches!(
            self.name.as_str(),
            "f" | "sSup" | "sSub" | "sSubSup" | "rad" | "nary" | "d" | "m" | "eqArr"
        ) || self.children.iter().any(Self::needs_box_layout)
    }

    pub(super) fn child(&self, name: &str) -> Option<&Self> {
        self.children.iter().find(|child| child.name == name)
    }

    pub(super) fn property(&self, name: &str) -> Option<&str> {
        self.children
            .iter()
            .filter(|child| child.name.ends_with("Pr"))
            .find_map(|properties| properties.child(name))
            .and_then(|property| property.value.as_deref())
    }
}

// Unicode mathematical alphabets have deliberate holes occupied by Letterlike
// Symbols. Do not substitute lookalikes (e.g. U+2113 for script small l).
// https://www.unicode.org/charts/nameslist/n_1D400.html
pub(super) fn math_alphabet(text: &str, script: &str, bold: bool) -> String {
    text.chars()
        .map(|c| {
            let (upper, lower, digits) = match (script, bold) {
                ("script", false) => (0x1d49c, 0x1d4b6, 0),
                ("script", true) => (0x1d4d0, 0x1d4ea, 0),
                ("fraktur", false) => (0x1d504, 0x1d51e, 0),
                ("fraktur", true) => (0x1d56c, 0x1d586, 0),
                ("double-struck", _) => (0x1d538, 0x1d552, 0x1d7d8),
                ("sans-serif", false) => (0x1d5a0, 0x1d5ba, 0x1d7e2),
                ("sans-serif", true) => (0x1d5d4, 0x1d5ee, 0x1d7ec),
                ("monospace", _) => (0x1d670, 0x1d68a, 0x1d7f6),
                _ => return c,
            };
            let point = match c {
                'A'..='Z' => upper + c as u32 - 'A' as u32,
                'a'..='z' => lower + c as u32 - 'a' as u32,
                '0'..='9' if digits != 0 => digits + c as u32 - '0' as u32,
                _ => return c,
            };
            char::from_u32(match point {
                0x1d49d => 0x212c,
                0x1d4a0 => 0x2130,
                0x1d4a1 => 0x2131,
                0x1d4a3 => 0x210b,
                0x1d4a4 => 0x2110,
                0x1d4a7 => 0x2112,
                0x1d4a8 => 0x2133,
                0x1d4ad => 0x211b,
                0x1d4ba => 0x212f,
                0x1d4bc => 0x210a,
                0x1d4c4 => 0x2134,
                0x1d506 => 0x212d,
                0x1d50b => 0x210c,
                0x1d50c => 0x2111,
                0x1d515 => 0x211c,
                0x1d51d => 0x2128,
                0x1d53a => 0x2102,
                0x1d53f => 0x210d,
                0x1d545 => 0x2115,
                0x1d547 => 0x2119,
                0x1d548 => 0x211a,
                0x1d549 => 0x211d,
                0x1d551 => 0x2124,
                value => value,
            })
            .unwrap_or(c)
        })
        .collect()
}

fn omml_children(node: &OmmlNode, name: &str) -> Vec<String> {
    node.children
        .iter()
        .filter(|child| child.name == name)
        .map(render_omml_node)
        .collect()
}

fn omml_descendant_value(node: &OmmlNode, name: &str) -> Option<String> {
    if node.name == name {
        return node.value.clone().or_else(|| {
            let text = render_omml_node(node);
            (!text.is_empty()).then_some(text)
        });
    }
    node.children
        .iter()
        .find_map(|child| omml_descendant_value(child, name))
}

fn render_omml_node(node: &OmmlNode) -> String {
    let joined = || {
        let mut text = node.text.clone();
        for child in &node.children {
            text.push_str(&render_omml_node(child));
        }
        text
    };
    match node.name.as_str() {
        "r" => math_alphabet(
            &joined(),
            node.property("scr").unwrap_or("roman"),
            matches!(node.property("sty"), Some("b" | "bi")),
        ),
        "f" => format!(
            "({})/({})",
            omml_child(node, "num"),
            omml_child(node, "den")
        ),
        "sSup" => format!("{}^({})", omml_child(node, "e"), omml_child(node, "sup")),
        "sSub" => format!("{}_({})", omml_child(node, "e"), omml_child(node, "sub")),
        "sSubSup" => format!(
            "{}_({})^({})",
            omml_child(node, "e"),
            omml_child(node, "sub"),
            omml_child(node, "sup")
        ),
        "rad" => {
            let degree = omml_child(node, "deg");
            let body = omml_child(node, "e");
            if degree.is_empty() {
                format!("√({body})")
            } else {
                format!("√[{degree}]({body})")
            }
        }
        "nary" => {
            let operator = omml_descendant_value(node, "chr").unwrap_or_else(|| "∑".to_owned());
            let subscript = omml_child(node, "sub");
            let superscript = omml_child(node, "sup");
            let body = omml_child(node, "e");
            format!("{operator}_({subscript})^({superscript})({body})")
        }
        "d" => {
            let begin = omml_descendant_value(node, "begChr").unwrap_or_else(|| "(".to_owned());
            let end = omml_descendant_value(node, "endChr").unwrap_or_else(|| ")".to_owned());
            format!("{begin}{}{end}", omml_child(node, "e"))
        }
        "m" => omml_children(node, "mr")
            .into_iter()
            .map(|row| format!("[{row}]"))
            .collect::<Vec<_>>()
            .join("; "),
        "mr" => omml_children(node, "e").join(", "),
        "eqArr" => omml_children(node, "e").join("; "),
        "func" => format!("{}({})", omml_child(node, "fName"), omml_child(node, "e")),
        "acc" => format!(
            "{}{}",
            omml_child(node, "e"),
            omml_descendant_value(node, "chr").unwrap_or_else(|| "ˆ".to_owned())
        ),
        "bar" => {
            let body = omml_child(node, "e");
            if omml_descendant_value(node, "pos").as_deref() == Some("bot") {
                format!("_({body})")
            } else {
                format!("‾({body})")
            }
        }
        "limLow" => format!("{}_({})", omml_child(node, "e"), omml_child(node, "lim")),
        "limUpp" => format!("{}^({})", omml_child(node, "e"), omml_child(node, "lim")),
        "groupChr" => format!(
            "{}({})",
            omml_descendant_value(node, "chr").unwrap_or_else(|| "⏞".to_owned()),
            omml_child(node, "e")
        ),
        "box" | "borderBox" => format!("□({})", omml_child(node, "e")),
        "sPre" => format!(
            "_({})^({}){}",
            omml_child(node, "sub"),
            omml_child(node, "sup"),
            omml_child(node, "e")
        ),
        _ => joined(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        limits::Limits,
        xml::{XmlEvent, decode_xml_text, parse_xml},
    };

    #[test]
    fn capture_preserves_nested_properties_empty_nodes_styles_and_unknown_content() {
        let xml = r#"<m:oMath><m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:e><m:nary><m:naryPr><m:chr m:val="∏"/></m:naryPr><m:e><m:r><m:rPr><m:scr m:val="script"/></m:rPr><w:rPr><w:color w:val="FF0000"/></w:rPr><m:t>l&amp;T</m:t></m:r></m:e></m:nary></m:e></m:nary><m:unknown><m:r><m:t>x</m:t></m:r></m:unknown></m:oMath>"#;
        let mut capture = OmmlCapture::new(0, "oMath");
        let mut depth = 0;
        parse_xml(xml.as_bytes(), Limits::default(), |event| {
            match event {
                XmlEvent::StartElement {
                    name,
                    attributes,
                    empty,
                } => {
                    if depth > 0 {
                        capture.start(
                            super::super::local_name(name),
                            &attributes,
                            depth,
                            empty,
                            "math.xml",
                        )?;
                    }
                    if !empty {
                        depth += 1;
                    }
                }
                XmlEvent::EndElement { .. } => {
                    depth -= 1;
                    capture.end(depth, "math.xml")?;
                }
                XmlEvent::Text(text) => capture.text(&decode_xml_text(text)?),
                XmlEvent::Cdata(text) => capture.text(text),
            }
            Ok(())
        })
        .unwrap();
        let result = capture.finish("math.xml").unwrap();
        assert!(result.approximate);
        assert_eq!(
            result.root.child("nary").unwrap().property("chr"),
            Some("∑")
        );
        assert!(
            result.text.contains("∏") && result.text.contains("𝓁&𝒯") && result.text.ends_with('x')
        );
        assert!(
            OmmlCapture::new(0, "oMath")
                .finish("math.xml")
                .unwrap()
                .text
                .is_empty()
        );
    }

    #[test]
    fn alphabets_keep_symbols_and_use_unicode_holes_not_lookalikes() {
        assert_eq!(math_alphabet("lTeB", "script", false), "𝓁𝒯ℯℬ");
        assert_eq!(math_alphabet("Ce⇏↻⋩α", "fraktur", false), "ℭ𝔢⇏↻⋩α");
        assert_eq!(
            math_alphabet("CHNPQRZe09", "double-struck", false),
            "ℂℍℕℙℚℝℤ𝕖𝟘𝟡"
        );
        assert_eq!(math_alphabet("Az", "script", true), "𝓐𝔃");
        assert_eq!(math_alphabet("x1", "unknown", true), "x1");
    }
}
