//! Bounded extraction of document-scoped OOXML and OpenDocument font parts.

#[cfg(feature = "native-formats")]
use std::collections::HashMap;
use std::collections::HashSet;

use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
use crate::model::{EmbeddedFont, FontStyle};
use crate::package::Package;
use crate::xml::{XmlAttribute, XmlEvent, decode_xml_text, parse_xml};

use super::local_name;

#[derive(Clone, Debug)]
struct FontReference {
    family: String,
    relationship_id: String,
    #[cfg(feature = "native-formats")]
    key: Option<[u8; 16]>,
    style: FontStyle,
    weight: u16,
}

#[cfg(feature = "native-formats")]
pub(super) fn extract_docx(
    package: &Package<'_>,
    main_part: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> (Vec<EmbeddedFont>, Vec<(String, Vec<String>)>) {
    let Ok(relationships) = package.relationships(Some(main_part)) else {
        return (Vec::new(), Vec::new());
    };
    let font_tables: Vec<_> = relationships
        .iter()
        .filter(|relationship| relationship.type_uri.ends_with("/fontTable"))
        .collect();
    let Some(font_table) = font_tables.first() else {
        return (Vec::new(), Vec::new());
    };
    if font_tables.len() > 1 || font_table.external {
        warn(
            diagnostics,
            main_part,
            "DOCX embedded font table relationship is ambiguous or external; embedded fonts were ignored",
        );
        return (Vec::new(), Vec::new());
    }
    let font_table_part = font_table.target.as_str();
    let Ok(Some(bytes)) = package.part(font_table_part) else {
        warn(
            diagnostics,
            font_table_part,
            "DOCX embedded font table could not be read; browser fonts will be used",
        );
        return (Vec::new(), Vec::new());
    };
    let (references, alternate_names) =
        match docx_references(&bytes, package, font_table_part, diagnostics) {
            Ok(references) => references,
            Err(error) => {
                warn(
                    diagnostics,
                    font_table_part,
                    format!("DOCX embedded font metadata is invalid: {}", error.message),
                );
                return (Vec::new(), Vec::new());
            }
        };
    (
        materialize(package, font_table_part, references, true, diagnostics),
        alternate_names,
    )
}

#[cfg(feature = "native-formats")]
pub(super) fn extract_pptx(
    package: &Package<'_>,
    main_part: &str,
    presentation: &[u8],
    diagnostics: &mut Vec<Diagnostic>,
) -> Vec<EmbeddedFont> {
    let references = match pptx_references(presentation, package, main_part) {
        Ok(references) => references,
        Err(error) => {
            warn(
                diagnostics,
                main_part,
                format!("PPTX embedded font metadata is invalid: {}", error.message),
            );
            return Vec::new();
        }
    };
    materialize(package, main_part, references, false, diagnostics)
}

#[cfg(feature = "odf-formats")]
pub(super) fn extract_odf(
    package: &Package<'_>,
    diagnostics: &mut Vec<Diagnostic>,
) -> Vec<EmbeddedFont> {
    let mut references = Vec::new();
    for part in ["styles.xml", "content.xml"] {
        let Ok(Some(bytes)) = package.part(part) else {
            continue;
        };
        match odf_references(&bytes, package, part) {
            Ok(mut parsed) => references.append(&mut parsed),
            Err(error) => warn(
                diagnostics,
                part,
                format!(
                    "OpenDocument embedded font metadata is invalid: {}",
                    error.message
                ),
            ),
        }
    }
    materialize_odf(package, references, diagnostics)
}

#[cfg(feature = "native-formats")]
fn docx_references(
    bytes: &[u8],
    package: &Package<'_>,
    part: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> Result<(Vec<FontReference>, Vec<(String, Vec<String>)>), Diagnostic> {
    let mut family: Option<String> = None;
    let mut references = Vec::new();
    let mut alternate_names = Vec::new();
    let mut alternate_names_truncated = false;
    parse_xml(bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name, attributes, ..
            } => match local_name(name) {
                "font" => family = optional_attribute(&attributes, "name")?,
                "altName" => {
                    if let Some(family) = &family {
                        let value = optional_attribute(&attributes, "val")?.unwrap_or_default();
                        let mut names = Vec::new();
                        let mut invalid =
                            family.is_empty() || family.len() > 128 || has_control(family);
                        for name in value
                            .split(',')
                            .map(str::trim)
                            .filter(|name| !name.is_empty())
                        {
                            if name.len() > 128 || has_control(name) || names.len() >= 64 {
                                invalid = true;
                                continue;
                            }
                            if !names
                                .iter()
                                .any(|existing: &String| existing.eq_ignore_ascii_case(name))
                            {
                                names.push(name.to_owned());
                            }
                        }
                        if invalid {
                            warn(
                                diagnostics,
                                part,
                                "Invalid or excessive DOCX font alternate names were ignored",
                            );
                        }
                        if !names.is_empty()
                            && !family.is_empty()
                            && family.len() <= 128
                            && !has_control(family)
                        {
                            if alternate_names.len() < 1024 {
                                alternate_names.push((family.clone(), names));
                            } else if !alternate_names_truncated {
                                alternate_names_truncated = true;
                                warn(
                                    diagnostics,
                                    part,
                                    "Excessive DOCX font alternate-name declarations were ignored",
                                );
                            }
                        }
                    }
                }
                variant @ ("embedRegular" | "embedBold" | "embedItalic" | "embedBoldItalic") => {
                    let Some(family) = family.as_ref() else {
                        return Ok(());
                    };
                    let Some(relationship_id) = optional_attribute(&attributes, "id")? else {
                        return Ok(());
                    };
                    let key = optional_attribute(&attributes, "fontKey")?
                        .as_deref()
                        .and_then(parse_guid_key);
                    let (style, weight) = match variant {
                        "embedBold" => (FontStyle::Normal, 700),
                        "embedItalic" => (FontStyle::Italic, 400),
                        "embedBoldItalic" => (FontStyle::Italic, 700),
                        _ => (FontStyle::Normal, 400),
                    };
                    references.push(FontReference {
                        family: family.clone(),
                        relationship_id,
                        key,
                        style,
                        weight,
                    });
                }
                _ => {}
            },
            XmlEvent::EndElement { name } if local_name(name) == "font" => family = None,
            _ => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok((references, alternate_names))
}

#[cfg(feature = "native-formats")]
fn pptx_references(
    bytes: &[u8],
    package: &Package<'_>,
    part: &str,
) -> Result<Vec<FontReference>, Diagnostic> {
    let mut in_embedded_font = false;
    let mut family: Option<String> = None;
    let mut variants = Vec::<(String, FontStyle, u16)>::new();
    let mut references = Vec::new();
    parse_xml(bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => match local_name(name) {
                "embeddedFont" => {
                    in_embedded_font = true;
                    family = None;
                    variants.clear();
                    if empty {
                        in_embedded_font = false;
                    }
                }
                "font" if in_embedded_font => {
                    family = optional_attribute(&attributes, "typeface")?;
                }
                variant @ ("regular" | "bold" | "italic" | "boldItalic") if in_embedded_font => {
                    if let Some(id) = optional_attribute(&attributes, "id")? {
                        let (style, weight) = match variant {
                            "bold" => (FontStyle::Normal, 700),
                            "italic" => (FontStyle::Italic, 400),
                            "boldItalic" => (FontStyle::Italic, 700),
                            _ => (FontStyle::Normal, 400),
                        };
                        variants.push((id, style, weight));
                    }
                }
                _ => {}
            },
            XmlEvent::EndElement { name } if local_name(name) == "embeddedFont" => {
                if let Some(family) = family.take() {
                    references.extend(variants.drain(..).map(
                        |(relationship_id, style, weight)| FontReference {
                            family: family.clone(),
                            relationship_id,
                            key: None,
                            style,
                            weight,
                        },
                    ));
                }
                variants.clear();
                in_embedded_font = false;
            }
            _ => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(references)
}

#[cfg(feature = "odf-formats")]
fn odf_references(
    bytes: &[u8],
    package: &Package<'_>,
    part: &str,
) -> Result<Vec<FontReference>, Diagnostic> {
    let mut family: Option<String> = None;
    let mut style = FontStyle::Normal;
    let mut weight = 400_u16;
    let mut references = Vec::new();
    parse_xml(bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name, attributes, ..
            } => match local_name(name) {
                "font-face" => {
                    family = optional_attribute(&attributes, "font-family")?
                        .or(optional_attribute(&attributes, "name")?)
                        .map(|value| value.trim_matches(['\'', '"']).to_owned());
                    style = match optional_attribute(&attributes, "font-style")?.as_deref() {
                        Some("italic") => FontStyle::Italic,
                        Some("oblique") => FontStyle::Oblique,
                        _ => FontStyle::Normal,
                    };
                    weight = parse_odf_weight(
                        optional_attribute(&attributes, "font-weight")?.as_deref(),
                    );
                }
                "font-face-uri" => {
                    if let (Some(family), Some(href)) =
                        (family.as_ref(), optional_attribute(&attributes, "href")?)
                    {
                        references.push(FontReference {
                            family: family.clone(),
                            relationship_id: href,
                            #[cfg(feature = "native-formats")]
                            key: None,
                            style,
                            weight,
                        });
                    }
                }
                _ => {}
            },
            XmlEvent::EndElement { name } if local_name(name) == "font-face" => {
                family = None;
                style = FontStyle::Normal;
                weight = 400;
            }
            _ => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(references)
}

#[cfg(feature = "odf-formats")]
fn parse_odf_weight(value: Option<&str>) -> u16 {
    match value.map(str::trim) {
        Some("bold" | "bolder") => 700,
        Some(value) => value
            .parse::<u16>()
            .ok()
            .filter(|weight| (1..=1000).contains(weight))
            .unwrap_or(400),
        None => 400,
    }
}

#[cfg(feature = "odf-formats")]
fn materialize_odf(
    package: &Package<'_>,
    references: Vec<FontReference>,
    diagnostics: &mut Vec<Diagnostic>,
) -> Vec<EmbeddedFont> {
    let mut fonts = Vec::new();
    let mut seen = HashSet::new();
    let mut total_bytes = 0_usize;
    let mut budget_reported = false;
    for reference in references {
        let family = reference.family.trim();
        let Some(target) = safe_odf_font_target(&reference.relationship_id) else {
            warn(
                diagnostics,
                "styles.xml",
                "external or unsafe OpenDocument embedded font URI was ignored",
            );
            continue;
        };
        let key = (
            family.to_lowercase(),
            reference.style.code(),
            reference.weight,
        );
        if family.is_empty() || family.len() > 256 || has_control(family) {
            continue;
        }
        let Ok(Some(bytes)) = package.part(&target) else {
            warn(
                diagnostics,
                &target,
                "OpenDocument embedded font data could not be read; browser fonts will be used",
            );
            continue;
        };
        if bytes.is_empty() {
            warn(
                diagnostics,
                &target,
                "OpenDocument embedded font data is empty",
            );
            continue;
        }
        if seen.contains(&key) {
            continue;
        }
        let Some(next_total) = total_bytes.checked_add(bytes.len()) else {
            continue;
        };
        if next_total > package.limits().max_font_bytes {
            if !budget_reported {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::FontBytesLimit,
                        Phase::Parse,
                        Fidelity::Omitted,
                        "embedded fonts exceed the configured font byte limit; excess fonts were ignored",
                    )
                    .in_part("styles.xml"),
                );
                budget_reported = true;
            }
            continue;
        }
        total_bytes = next_total;
        if !seen.insert(key) {
            continue;
        }
        fonts.push(EmbeddedFont {
            family: family.to_owned(),
            bytes: bytes.into_vec(),
            style: reference.style,
            weight: reference.weight,
        });
    }
    fonts
}

#[cfg(feature = "odf-formats")]
fn safe_odf_font_target(value: &str) -> Option<String> {
    let value = value.trim();
    if value.is_empty() || value.starts_with('/') || value.contains(['\\', '\0', '?', '#', ':']) {
        return None;
    }
    let mut segments = Vec::new();
    for segment in value.split('/') {
        match segment {
            "" | "." => {}
            ".." => return None,
            _ => segments.push(segment),
        }
    }
    (!segments.is_empty()).then(|| segments.join("/"))
}

#[cfg(feature = "native-formats")]
fn materialize(
    package: &Package<'_>,
    relationship_part: &str,
    references: Vec<FontReference>,
    obfuscated: bool,
    diagnostics: &mut Vec<Diagnostic>,
) -> Vec<EmbeddedFont> {
    if references.is_empty() {
        return Vec::new();
    }
    let relationships = match package.relationships(Some(relationship_part)) {
        Ok(relationships) => relationships,
        Err(error) => {
            warn(
                diagnostics,
                relationship_part,
                format!(
                    "embedded font relationships could not be read: {}",
                    error.message
                ),
            );
            return Vec::new();
        }
    };
    let by_id: HashMap<_, _> = relationships
        .iter()
        .map(|relationship| (relationship.id.as_str(), relationship))
        .collect();
    let mut fonts = Vec::new();
    let mut seen = HashSet::new();
    let mut total_bytes = 0_usize;
    let mut budget_reported = false;
    for reference in references {
        let family = reference.family.trim();
        if family.is_empty() || family.len() > 256 || has_control(family) {
            warn(
                diagnostics,
                relationship_part,
                "embedded font has an invalid family name",
            );
            continue;
        }
        let key = (
            family.to_lowercase(),
            reference.style.code(),
            reference.weight,
        );
        let Some(relationship) = by_id.get(reference.relationship_id.as_str()).copied() else {
            warn(
                diagnostics,
                relationship_part,
                format!(
                    "embedded font relationship {} is missing",
                    reference.relationship_id
                ),
            );
            continue;
        };
        if relationship.external || !relationship.type_uri.ends_with("/font") {
            warn(
                diagnostics,
                relationship_part,
                "embedded font relationship is external or has an invalid type",
            );
            continue;
        }
        let Ok(Some(bytes)) = package.part(&relationship.target) else {
            warn(
                diagnostics,
                &relationship.target,
                "embedded font data could not be read; browser fonts will be used",
            );
            continue;
        };
        if bytes.is_empty() {
            warn(
                diagnostics,
                &relationship.target,
                "embedded font data is empty",
            );
            continue;
        }
        let mut bytes = bytes.into_vec();
        if obfuscated {
            let Some(key) = reference.key else {
                warn(
                    diagnostics,
                    &relationship.target,
                    "obfuscated DOCX font has no valid fontKey and was ignored",
                );
                continue;
            };
            if bytes.len() < 32 {
                warn(
                    diagnostics,
                    &relationship.target,
                    "obfuscated DOCX font is shorter than 32 bytes and was ignored",
                );
                continue;
            }
            for index in 0..32 {
                bytes[index] ^= key[index % key.len()];
            }
        }
        if seen.contains(&key) {
            continue;
        }
        let Some(next_total) = total_bytes.checked_add(bytes.len()) else {
            continue;
        };
        if next_total > package.limits().max_font_bytes {
            if !budget_reported {
                diagnostics.push(
                    Diagnostic::warning(
                        DiagnosticCode::FontBytesLimit,
                        Phase::Parse,
                        Fidelity::Omitted,
                        "embedded fonts exceed the configured font byte limit; excess fonts were ignored",
                    )
                    .in_part(relationship_part),
                );
                budget_reported = true;
            }
            continue;
        }
        total_bytes = next_total;
        if !seen.insert(key) {
            continue;
        }
        fonts.push(EmbeddedFont {
            family: family.to_owned(),
            bytes,
            style: reference.style,
            weight: reference.weight,
        });
    }
    fonts
}

#[cfg(feature = "native-formats")]
fn parse_guid_key(value: &str) -> Option<[u8; 16]> {
    let compact: String = value
        .trim()
        .trim_matches(['{', '}'])
        .chars()
        .filter(|character| *character != '-')
        .collect();
    if compact.len() != 32 || !compact.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    let mut key = [0_u8; 16];
    for index in 0..16 {
        let offset = index * 2;
        key[15 - index] = u8::from_str_radix(&compact[offset..offset + 2], 16).ok()?;
    }
    Some(key)
}

fn optional_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
) -> Result<Option<String>, Diagnostic> {
    attributes
        .iter()
        .find(|attribute| local_name(attribute.name) == name)
        .map(|attribute| decode_xml_text(attribute.value).map(|value| value.into_owned()))
        .transpose()
}

fn has_control(value: &str) -> bool {
    value.chars().any(|character| character.is_control())
}

fn with_part(mut error: Diagnostic, part: &str) -> Diagnostic {
    if error.location.part.is_none() {
        error.location.part = Some(part.to_owned());
    }
    error
}

fn warn(diagnostics: &mut Vec<Diagnostic>, part: &str, message: impl Into<String>) {
    diagnostics.push(
        Diagnostic::warning(
            DiagnosticCode::EmbeddedFontInvalid,
            Phase::Parse,
            Fidelity::Omitted,
            message,
        )
        .in_part(part),
    );
}

#[cfg(all(test, feature = "native-formats"))]
mod tests {
    use super::{docx_references, parse_guid_key};
    use crate::{limits::Limits, package::Package};

    #[test]
    fn docx_alternate_names_skip_empty_duplicates_and_invalid_candidates() {
        let package = Package::open(
            include_bytes!("../../tests/fixtures/complex0.docx"),
            Limits::default(),
        )
        .unwrap();
        let xml = format!(
            r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:font w:name="Authored"><w:altName w:val=" , Courier New,,, courier new, {}, Arial, "/></w:font>
            <w:font w:name="No aliases"/>
        </w:fonts>"#,
            "x".repeat(129)
        );
        let mut diagnostics = Vec::new();
        let (references, names) = docx_references(
            xml.as_bytes(),
            &package,
            "word/fontTable.xml",
            &mut diagnostics,
        )
        .unwrap();
        assert!(references.is_empty());
        assert_eq!(
            names,
            vec![(
                "Authored".to_owned(),
                vec!["Courier New".to_owned(), "Arial".to_owned()]
            )]
        );
        assert_eq!(diagnostics.len(), 1);
    }

    #[test]
    fn reverses_docx_font_obfuscation_guid_bytes() {
        assert_eq!(
            parse_guid_key("{001B70DC-AA60-4AD5-90EC-18A0948E1EAE}"),
            Some([
                0xae, 0x1e, 0x8e, 0x94, 0xa0, 0x18, 0xec, 0x90, 0xd5, 0x4a, 0x60, 0xaa, 0xdc, 0x70,
                0x1b, 0x00,
            ]),
        );
        assert_eq!(parse_guid_key("not-a-guid"), None);
    }
}
