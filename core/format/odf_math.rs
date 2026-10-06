//! Shared ODF embedded MathML text semantics; host adapters retain their own layout.
use super::local_name;
use crate::diagnostic::{Diagnostic, DiagnosticCode, Phase};
use crate::limits::Limits;
use crate::xml::{XmlEvent, decode_xml_text, parse_xml};

fn format_error(part: &str, message: &str) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message).in_part(part)
}
fn with_part(error: Diagnostic, part: &str) -> Diagnostic {
    error.in_part(part)
}

#[derive(Debug)]
pub(super) struct OdfMath {
    pub source_part: String,
    pub text: String,
    pub approximate: bool,
}

#[derive(Debug)]
struct OdfMathNode {
    name: String,
    text: String,
    children: Vec<OdfMathNode>,
}

pub(super) fn parse_odf_math(
    bytes: &[u8],
    limits: Limits,
    part: &str,
) -> Result<Option<OdfMath>, Diagnostic> {
    // Only the inert, canonical OpenOffice declaration is accepted. No DTD is loaded
    // or entity expanded; internal subsets and all other declarations remain forbidden.
    const DECLARATION: &[u8] = br#"<!DOCTYPE math:math PUBLIC "-//OpenOffice.org//DTD Modified W3C MathML 1.01//EN" "math.dtd">"#;
    let mut prolog = bytes
        .strip_prefix(b"\xef\xbb\xbf")
        .unwrap_or(bytes)
        .trim_ascii_start();
    if prolog.starts_with(b"<?xml ") {
        if let Some(end) = prolog.windows(2).position(|w| w == b"?>") {
            prolog = prolog[end + 2..].trim_ascii_start();
        }
    }
    let mut normalized = None;
    if bytes.len() <= limits.max_xml_bytes && prolog.starts_with(DECLARATION) {
        let start = bytes.len() - prolog.len();
        let mut copy = bytes.to_vec();
        copy[start..start + DECLARATION.len()].fill(b' ');
        normalized = Some(copy);
    }
    let bytes = normalized.as_deref().unwrap_or(bytes);
    let mut found_math = false;
    let mut stack: Vec<OdfMathNode> = Vec::new();
    let mut root = None;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes: _,
                empty,
            } => {
                let local = local_name(name);
                if stack.is_empty() {
                    if !found_math && local == "math" {
                        found_math = true;
                        let node = OdfMathNode {
                            name: local.to_owned(),
                            text: String::new(),
                            children: Vec::new(),
                        };
                        if empty {
                            root = Some(node);
                        } else {
                            stack.push(node);
                        }
                    }
                } else {
                    let node = OdfMathNode {
                        name: local.to_owned(),
                        text: String::new(),
                        children: Vec::new(),
                    };
                    if empty {
                        if let Some(parent) = stack.last_mut() {
                            parent.children.push(node);
                        }
                    } else {
                        stack.push(node);
                    }
                }
            }
            XmlEvent::EndElement { name } => {
                let local = local_name(name);
                if stack.last().is_some_and(|node| node.name == local) {
                    let node = stack.pop().ok_or_else(|| {
                        format_error(part, "MathML parser state ended unexpectedly")
                    })?;
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(node);
                    } else {
                        root = Some(node);
                    }
                }
            }
            XmlEvent::Text(text) => {
                if let Some(node) = stack.last_mut() {
                    let text = decode_xml_text(text).map_err(|error| with_part(error, part))?;
                    if !text.trim().is_empty() {
                        node.text.push_str(&text);
                    }
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(node) = stack.last_mut()
                    && !text.trim().is_empty()
                {
                    node.text.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    let Some(root) = root else {
        return Ok(None);
    };
    let (text, approximate) = render_odf_math_node(&root);
    Ok(Some(OdfMath {
        source_part: part.to_owned(),
        text,
        approximate,
    }))
}

fn render_odf_math_node(node: &OdfMathNode) -> (String, bool) {
    match node.name.as_str() {
        // Only the presentation branch is visible; annotations are alternate encodings.
        "semantics" => {
            return node
                .children
                .first()
                .map(render_odf_math_node)
                .unwrap_or_default();
        }
        "annotation" | "annotation-xml" => return (String::new(), false),
        _ => {}
    }
    let mut approximate = false;
    let children = node
        .children
        .iter()
        .map(|child| {
            let (text, child_approximate) = render_odf_math_node(child);
            approximate |= child_approximate;
            text
        })
        .collect::<Vec<_>>();
    let own_text = node.text.trim();
    let joined = || {
        let mut text = own_text.to_owned();
        for child in &children {
            text.push_str(child);
        }
        text
    };
    let text = match node.name.as_str() {
        "math" | "mrow" | "mi" | "mn" | "mo" => joined(),
        "mfrac" if children.len() >= 2 => format!("({})/({})", children[0], children[1]),
        "msup" if children.len() >= 2 => format!("{}^({})", children[0], children[1]),
        "msub" if children.len() >= 2 => format!("{}_({})", children[0], children[1]),
        "msqrt" if !children.is_empty() => format!("√({})", children.concat()),
        _ => {
            approximate = true;
            joined()
        }
    };
    (text, approximate)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shares_math_semantics_and_keeps_unsafe_declarations_blocked() {
        for (xml, text, approximate) in [
            (
                "<math><semantics><mrow><mn>1</mn><mo>+</mo><mn>1</mn></mrow><annotation>1 + 1</annotation></semantics></math>",
                "1+1",
                false,
            ),
            (
                "<math><msqrt><mfrac><mi>a</mi><mi>b</mi></mfrac></msqrt></math>",
                "√((a)/(b))",
                false,
            ),
            ("<math><unknown><mi>x</mi></unknown></math>", "x", true),
        ] {
            let math = parse_odf_math(xml.as_bytes(), Limits::default(), "math.xml")
                .unwrap()
                .unwrap();
            assert_eq!((math.text.as_str(), math.approximate), (text, approximate));
        }
        for xml in [
            r#"<!DOCTYPE math:math SYSTEM "https://example.com/math.dtd"><math:math/>"#,
            r#"<!DOCTYPE math:math PUBLIC "-//OpenOffice.org//DTD Modified W3C MathML 1.01//EN" "math.dtd" [<!ENTITY x "boom">]><math:math>&x;</math:math>"#,
        ] {
            assert_eq!(
                parse_odf_math(xml.as_bytes(), Limits::default(), "math.xml")
                    .unwrap_err()
                    .code,
                DiagnosticCode::XmlDtdForbidden
            );
        }
    }
}
