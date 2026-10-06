//! OFD fixed-document parser.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use sm2::dsa::{Signature as Sm2Signature, VerifyingKey, signature::Verifier};
use sm3::{Digest, Sm3};

use crate::RetainedInput;
use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
use crate::limits::Limits;
use crate::model::{
    AffineTransform, BlendMode, Document, DocumentFormat, DocumentKind, EmbeddedFont, FillRule,
    FontStyle, Geometry, GradientStop, LineCap, LineJoin, MappingQuality, Object, ObjectKind,
    OutlineItem, Paint, PathCommand, Rect, SheetAxis, SourceLocator, SourceRef, StretchMode,
    StrokeStyle, TextAlign, TextLayout, TextRun, TileMode, Unit, UnitKind, Visual,
    VisualBrushChild,
};
use crate::package::{Package, resolve_internal_target};
use crate::xml::{XmlEvent, decode_xml_text, parse_xml};

const MM_TO_CSS_PX: f32 = 96.0 / 25.4;
const MAX_PAGE_DIMENSION: f32 = 1_000_000.0;
const SM2_WITH_SM3_OID: &[u8] = &[0x2a, 0x81, 0x1c, 0xcf, 0x55, 0x01, 0x83, 0x75];

#[derive(Clone, Debug)]
struct Node {
    name: String,
    attributes: Vec<(String, String)>,
    children: Vec<Node>,
    text: String,
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

    fn descendants<'a>(&'a self, name: &str) -> Vec<&'a Node> {
        let mut found = Vec::new();
        self.collect_descendants(name, &mut found);
        found
    }

    fn collect_descendants<'a>(&'a self, name: &str, found: &mut Vec<&'a Node>) {
        for child in &self.children {
            if child.local_name() == name {
                found.push(child);
            }
            child.collect_descendants(name, found);
        }
    }
}

#[derive(Clone, Copy)]
struct DerValue<'a> {
    tag: u8,
    full: &'a [u8],
    content: &'a [u8],
}

fn der_value(bytes: &[u8]) -> Option<(DerValue<'_>, &[u8])> {
    let (&tag, rest) = bytes.split_first()?;
    let (&first_length, rest) = rest.split_first()?;
    let (length, header_tail) = if first_length & 0x80 == 0 {
        (usize::from(first_length), 0)
    } else {
        let count = usize::from(first_length & 0x7f);
        if count == 0 || count > std::mem::size_of::<usize>() || rest.len() < count {
            return None;
        }
        let mut length = 0_usize;
        for byte in &rest[..count] {
            length = length.checked_mul(256)?.checked_add(usize::from(*byte))?;
        }
        (length, count)
    };
    let header = 2 + header_tail;
    let end = header.checked_add(length)?;
    let full = bytes.get(..end)?;
    Some((
        DerValue {
            tag,
            full,
            content: &full[header..],
        },
        &bytes[end..],
    ))
}

fn der_children(value: DerValue<'_>) -> Option<Vec<DerValue<'_>>> {
    if value.tag & 0x20 == 0 {
        return None;
    }
    let mut bytes = value.content;
    let mut children = Vec::new();
    while !bytes.is_empty() {
        let (child, rest) = der_value(bytes)?;
        children.push(child);
        bytes = rest;
    }
    Some(children)
}

fn find_der_content<'a>(
    bytes: &'a [u8],
    depth: u8,
    predicate: &impl Fn(u8, &[u8]) -> bool,
) -> Option<&'a [u8]> {
    if depth > 32 {
        return None;
    }
    let mut rest = bytes;
    while !rest.is_empty() {
        let (value, tail) = der_value(rest)?;
        if predicate(value.tag, value.content) {
            return Some(value.content);
        }
        if value.tag & 0x20 != 0
            && let Some(found) = find_der_content(value.content, depth + 1, predicate)
        {
            return Some(found);
        }
        rest = tail;
    }
    None
}

fn find_embedded_ofd(bytes: &[u8]) -> Option<&[u8]> {
    find_der_content(bytes, 0, &|tag, content| {
        tag == 0x04 && content.starts_with(b"PK")
    })
}

fn find_embedded_image(bytes: &[u8]) -> Option<(&'static str, &[u8])> {
    let bytes = find_der_content(bytes, 0, &|tag, content| {
        tag == 0x04
            && super::presentation_image::office_image_media_type_from_signature(content).is_ok()
    })?;
    super::presentation_image::office_image_media_type_from_signature(bytes)
        .ok()
        .map(|media_type| (media_type, bytes))
}

fn sm3_digest(bytes: &[u8]) -> [u8; 32] {
    Sm3::digest(bytes).into()
}

fn base64_digest(bytes: [u8; 32]) -> String {
    const ALPHABET: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut encoded = String::with_capacity(44);
    for chunk in bytes.chunks(3) {
        let bits = (u32::from(chunk[0]) << 16)
            | (u32::from(*chunk.get(1).unwrap_or(&0)) << 8)
            | u32::from(*chunk.get(2).unwrap_or(&0));
        encoded.push(ALPHABET[((bits >> 18) & 63) as usize] as char);
        encoded.push(ALPHABET[((bits >> 12) & 63) as usize] as char);
        encoded.push(if chunk.len() > 1 {
            ALPHABET[((bits >> 6) & 63) as usize] as char
        } else {
            '='
        });
        encoded.push(if chunk.len() > 2 {
            ALPHABET[(bits & 63) as usize] as char
        } else {
            '='
        });
    }
    encoded
}

enum SignatureVerification {
    Valid,
    Invalid,
    Unsupported,
}

fn inspect_signatures(
    package: &Package<'_>,
    path: &str,
    page_ids: &[String],
    limits: Limits,
    diagnostics: &mut Vec<Diagnostic>,
) -> Vec<SealAppearance> {
    let index_part = match resolve_internal_target(None, path) {
        Ok(part) => part,
        Err(message) => {
            diagnostics.push(signature_warning("OFD.xml", Fidelity::Omitted, message));
            return Vec::new();
        }
    };
    let index = match parse_part(package, &index_part, limits) {
        Ok(index) => index,
        Err(error) => {
            diagnostics.push(signature_warning(
                &index_part,
                Fidelity::Omitted,
                error.message,
            ));
            return Vec::new();
        }
    };
    let mut seals = Vec::new();
    for entry in index.descendants("Signature") {
        let Some(base_loc) = entry.attribute("BaseLoc") else {
            diagnostics.push(signature_warning(
                &index_part,
                Fidelity::Omitted,
                "OFD signature entry has no BaseLoc",
            ));
            continue;
        };
        let signature_part = match resolve_internal_target(Some(&index_part), base_loc) {
            Ok(part) => part,
            Err(message) => {
                diagnostics.push(signature_warning(&index_part, Fidelity::Omitted, message));
                continue;
            }
        };
        let signature_bytes = match package.part(&signature_part) {
            Ok(Some(bytes)) => bytes.into_vec(),
            Ok(None) => {
                diagnostics.push(signature_warning(
                    &signature_part,
                    Fidelity::Omitted,
                    "OFD signature description is missing",
                ));
                continue;
            }
            Err(error) => {
                diagnostics.push(signature_warning(
                    &signature_part,
                    Fidelity::Omitted,
                    error.message,
                ));
                continue;
            }
        };
        let signature = match parse_xml_node(&signature_bytes, &signature_part, limits) {
            Ok(signature) => signature,
            Err(error) => {
                diagnostics.push(signature_warning(
                    &signature_part,
                    Fidelity::Omitted,
                    error.message,
                ));
                continue;
            }
        };
        let references_valid =
            verify_signature_references(package, &signature, &signature_part, diagnostics);
        let signed_value = signature
            .descendants("SignedValue")
            .into_iter()
            .find_map(|node| {
                let path = node.text.trim();
                (!path.is_empty()).then_some(path)
            })
            .and_then(|path| resolve_internal_target(Some(&signature_part), path).ok())
            .and_then(|part| {
                package
                    .part(&part)
                    .ok()
                    .flatten()
                    .map(|bytes| bytes.into_vec())
            });
        if references_valid {
            match signed_value
                .as_deref()
                .map(|bytes| verify_ses_sm2(bytes, &signature_bytes))
                .unwrap_or(SignatureVerification::Unsupported)
            {
                SignatureVerification::Valid => diagnostics.push(signature_warning(
                    &signature_part,
                    Fidelity::Unsupported,
                    "OFD_SIGNATURE_CERTIFICATE_TRUST_UNAVAILABLE: the SM2/SM3 signature is valid, but certificate trust, revocation, and timestamp status were not established",
                )),
                SignatureVerification::Invalid => diagnostics.push(signature_warning(
                    &signature_part,
                    Fidelity::Blocked,
                    "OFD_SIGNATURE_VALUE_INVALID: the SM2/SM3 electronic signature does not match the signed description",
                )),
                SignatureVerification::Unsupported => diagnostics.push(signature_warning(
                    &signature_part,
                    Fidelity::Unsupported,
                    "OFD_SIGNATURE_VALUE_UNVERIFIABLE: the signature value has no supported SES SM2 certificate container",
                )),
            }
        }
        let seal_bytes = signature
            .descendants("Seal")
            .into_iter()
            .flat_map(|seal| seal.descendants("BaseLoc"))
            .find_map(|node| {
                let path = node.text.trim();
                resolve_internal_target(Some(&signature_part), path)
                    .ok()
                    .and_then(|part| package.part(&part).ok().flatten())
                    .map(|bytes| bytes.into_vec())
            })
            .or_else(|| signed_value.clone());
        let payload = seal_bytes.as_deref().and_then(|bytes| {
            bytes
                .starts_with(b"PK")
                .then_some(bytes)
                .or_else(|| find_embedded_ofd(bytes))
                .map(|bytes| SealPayload::Ofd(Arc::from(bytes)))
                .or_else(|| {
                    find_embedded_image(bytes).map(|(media_type, bytes)| SealPayload::Image {
                        media_type,
                        bytes: Arc::from(bytes),
                    })
                })
        });
        let Some(payload) = payload else {
            continue;
        };
        for stamp in signature.descendants("StampAnnot") {
            let (Some(unit_index), Some(bounds)) = (
                stamp
                    .attribute("PageRef")
                    .and_then(|id| page_ids.iter().position(|candidate| candidate == id)),
                stamp.attribute("Boundary").and_then(parse_box),
            ) else {
                continue;
            };
            let clip = match stamp.attribute("Clip") {
                Some(value) => match parse_box(value) {
                    Some(clip) => Some(clip),
                    None => {
                        diagnostics.push(signature_warning(
                            &signature_part,
                            Fidelity::Omitted,
                            "OFD stamp annotation has an invalid Clip",
                        ));
                        continue;
                    }
                },
                None => None,
            };
            seals.push(SealAppearance {
                unit_index,
                bounds,
                clip,
                payload: payload.clone(),
                source_part: signature_part.clone(),
            });
        }
    }
    seals
}

fn verify_signature_references(
    package: &Package<'_>,
    signature: &Node,
    signature_part: &str,
    diagnostics: &mut Vec<Diagnostic>,
) -> bool {
    let Some(references) = signature.descendants("References").into_iter().next() else {
        diagnostics.push(signature_warning(
            signature_part,
            Fidelity::Blocked,
            "OFD_SIGNATURE_REFERENCE_INVALID: SignedInfo has no References",
        ));
        return false;
    };
    if !matches!(
        references.attribute("CheckMethod"),
        Some("SM3" | "1.2.156.10197.1.401")
    ) {
        diagnostics.push(signature_warning(
            signature_part,
            Fidelity::Unsupported,
            "OFD_SIGNATURE_REFERENCE_UNSUPPORTED: only SM3 reference digests are supported",
        ));
        return false;
    }
    let mut valid = true;
    for reference in references
        .children
        .iter()
        .filter(|child| child.local_name() == "Reference")
    {
        let expected = reference
            .descendants("CheckValue")
            .into_iter()
            .next()
            .map(|node| node.text.split_ascii_whitespace().collect::<String>());
        let actual = reference
            .attribute("FileRef")
            .and_then(|path| resolve_internal_target(Some(signature_part), path).ok())
            .and_then(|part| package.part(&part).ok().flatten())
            .map(|bytes| base64_digest(sm3_digest(&bytes)));
        if expected.is_none() || expected != actual {
            valid = false;
            diagnostics.push(signature_warning(
                signature_part,
                Fidelity::Blocked,
                "OFD_SIGNATURE_REFERENCE_INVALID: an SM3 protected-file digest does not match",
            ));
        }
    }
    valid
}

fn verify_ses_sm2(bytes: &[u8], signature_xml: &[u8]) -> SignatureVerification {
    let Some((outer, tail)) = der_value(bytes) else {
        return SignatureVerification::Unsupported;
    };
    if outer.tag != 0x30 || !tail.is_empty() {
        return SignatureVerification::Unsupported;
    }
    let Some(children) = der_children(outer) else {
        return SignatureVerification::Unsupported;
    };
    let Some(tbs) = children.first().copied().filter(|value| value.tag == 0x30) else {
        return SignatureVerification::Unsupported;
    };
    if find_der_content(bytes, 0, &|tag, content| {
        tag == 0x06 && content == SM2_WITH_SM3_OID
    })
    .is_none()
    {
        return SignatureVerification::Unsupported;
    }
    let digest = sm3_digest(signature_xml);
    if find_der_content(tbs.full, 0, &|tag, content| {
        (tag == 0x03 && content.first() == Some(&0) && content.get(1..) == Some(digest.as_slice()))
            || (tag == 0x04 && content == digest)
    })
    .is_none()
    {
        return SignatureVerification::Invalid;
    }
    let signature = children
        .iter()
        .rev()
        .filter(|value| value.tag == 0x03 && value.content.first() == Some(&0))
        .find_map(|value| parse_sm2_signature(&value.content[1..]));
    let certificate = children
        .get(1)
        .filter(|value| value.tag == 0x04 && find_sm2_public_key(value.content).is_some())
        .map(|value| value.content)
        .or_else(|| {
            der_children(tbs)?.into_iter().rev().find_map(|value| {
                (value.tag == 0x04 && find_sm2_public_key(value.content).is_some())
                    .then_some(value.content)
            })
        });
    let (Some(signature), Some(public_key)) =
        (signature, certificate.and_then(find_sm2_public_key))
    else {
        return SignatureVerification::Unsupported;
    };
    let Ok(verifying_key) = VerifyingKey::from_sec1_bytes("1234567812345678", public_key) else {
        return SignatureVerification::Unsupported;
    };
    if verifying_key.verify(tbs.full, &signature).is_ok() {
        SignatureVerification::Valid
    } else {
        SignatureVerification::Invalid
    }
}

fn find_sm2_public_key(bytes: &[u8]) -> Option<&[u8]> {
    find_der_content(bytes, 0, &|tag, content| {
        tag == 0x03 && content.len() == 66 && content[..2] == [0, 4]
    })
    .and_then(|content| content.get(1..))
}

fn parse_sm2_signature(bytes: &[u8]) -> Option<Sm2Signature> {
    let (sequence, tail) = der_value(bytes)?;
    if sequence.tag != 0x30 || !tail.is_empty() {
        return None;
    }
    let integers = der_children(sequence)?;
    if integers.len() != 2 || integers.iter().any(|integer| integer.tag != 0x02) {
        return None;
    }
    let mut raw = [0_u8; 64];
    for (index, integer) in integers.iter().enumerate() {
        let positive_prefix = integer.content.first() == Some(&0);
        let value = integer.content.iter().position(|byte| *byte != 0).map_or(
            &integer.content[integer.content.len().saturating_sub(1)..],
            |start| &integer.content[start..],
        );
        if value.is_empty() || value.len() > 32 || (!positive_prefix && value[0] & 0x80 != 0) {
            return None;
        }
        let end = (index + 1) * 32;
        raw[end - value.len()..end].copy_from_slice(value);
    }
    Sm2Signature::from_slice(&raw).ok()
}

fn signature_warning(part: &str, fidelity: Fidelity, message: impl Into<String>) -> Diagnostic {
    Diagnostic::warning(
        DiagnosticCode::UnsupportedFeature,
        Phase::Security,
        fidelity,
        message,
    )
    .in_part(part)
}

pub(crate) struct PreparedOfd {
    bytes: PreparedBytes,
    page_parts: Vec<String>,
    templates: HashMap<String, String>,
    annotation_parts: Vec<Vec<String>>,
    outline: Vec<OutlineItem>,
    page_sizes: Vec<Rect>,
    fonts: HashMap<String, String>,
    images: HashMap<String, String>,
    draw_params: HashMap<String, Node>,
    composites: HashMap<String, (String, Node)>,
    seals: Vec<SealAppearance>,
    embedded_fonts: Vec<EmbeddedFont>,
    diagnostics: Vec<Diagnostic>,
    limits: Limits,
}

struct SealAppearance {
    unit_index: usize,
    bounds: Rect,
    clip: Option<Rect>,
    payload: SealPayload,
    source_part: String,
}

#[derive(Clone)]
enum SealPayload {
    Ofd(Arc<[u8]>),
    Image {
        media_type: &'static str,
        bytes: Arc<[u8]>,
    },
}

enum PreparedBytes {
    Owned(Vec<u8>),
    Retained(Arc<RetainedInput>),
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

pub fn detect_and_parse(bytes: &[u8], limits: Limits) -> Result<Option<Document>, Diagnostic> {
    let Some(prepared) = prepare(bytes, limits)? else {
        return Ok(None);
    };
    prepared.materialize(None).map(Some)
}

pub(crate) fn prepare(bytes: &[u8], limits: Limits) -> Result<Option<PreparedOfd>, Diagnostic> {
    prepare_impl(bytes, limits, None)
}

pub(crate) fn prepare_retained(
    bytes: Arc<RetainedInput>,
    limits: Limits,
) -> Result<Option<PreparedOfd>, Diagnostic> {
    let slice: &[u8] = &bytes;
    prepare_impl(slice, limits, Some(Arc::clone(&bytes)))
}

fn prepare_impl(
    bytes: &[u8],
    limits: Limits,
    retained: Option<Arc<RetainedInput>>,
) -> Result<Option<PreparedOfd>, Diagnostic> {
    if !bytes.starts_with(b"PK") {
        return Ok(None);
    }
    let package = Package::open(bytes, limits)?;
    let Some(root_bytes) = package.part("OFD.xml")? else {
        return Ok(None);
    };
    let root = parse_xml_node(&root_bytes, "OFD.xml", limits)?;
    if root.local_name() != "OFD" {
        return Ok(None);
    }
    let doc_root = root
        .descendants("DocRoot")
        .into_iter()
        .next()
        .map(|node| node.text.trim())
        .filter(|value| !value.is_empty())
        .ok_or_else(|| format_error("OFD.xml", "OFD package has no document root"))?;
    let document_part = resolve_internal_target(None, doc_root)
        .map_err(|message| format_error("OFD.xml", message))?;
    let document = parse_part(&package, &document_part, limits)?;
    if document.local_name() != "Document" {
        return Err(format_error(
            &document_part,
            "OFD document root is not Document",
        ));
    }
    let common_page_size = document
        .descendants("PhysicalBox")
        .into_iter()
        .next()
        .and_then(|node| parse_box(&node.text))
        .filter(|bounds| validate_page_size(*bounds, &document_part).is_ok());

    let mut diagnostics = Vec::new();
    if !matches!(root.attribute("Version"), Some("1.0" | "1.1")) {
        diagnostics.push(preparation_warning(
            "OFD.xml",
            "OFD version is not explicitly supported; compatible content will be rendered best-effort",
        ));
    }
    let mut page_ids = Vec::new();
    let mut page_parts = Vec::new();
    for page in document.descendants("Page") {
        let Some(path) = page.attribute("BaseLoc") else {
            diagnostics.push(preparation_warning(
                &document_part,
                "OFD page has no BaseLoc",
            ));
            continue;
        };
        match resolve_internal_target(Some(&document_part), path) {
            Ok(part) => {
                page_ids.push(page.attribute("ID").unwrap_or("").to_owned());
                page_parts.push(part);
            }
            Err(message) => diagnostics.push(preparation_warning(&document_part, message)),
        }
    }
    if page_parts.is_empty() {
        return Err(format_error(
            &document_part,
            "OFD document contains no pages",
        ));
    }
    if page_parts.len() > limits.max_document_objects {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "OFD page count exceeds the configured object limit",
        ));
    }
    let seals = root
        .descendants("Signatures")
        .into_iter()
        .find_map(|node| (!node.text.trim().is_empty()).then_some(node.text.trim()))
        .map(|path| inspect_signatures(&package, path, &page_ids, limits, &mut diagnostics))
        .unwrap_or_default();
    let page_size = if let Some(page_size) = common_page_size {
        page_size
    } else {
        let first_page = parse_part(&package, &page_parts[0], limits)?;
        first_page
            .descendants("PhysicalBox")
            .into_iter()
            .next()
            .and_then(|node| parse_box(&node.text))
            .filter(|bounds| validate_page_size(*bounds, &page_parts[0]).is_ok())
            .ok_or_else(|| format_error(&document_part, "OFD document has no valid PhysicalBox"))?
    };
    let mut templates = HashMap::new();
    for template in document.descendants("TemplatePage") {
        let (Some(id), Some(path)) = (template.attribute("ID"), template.attribute("BaseLoc"))
        else {
            diagnostics.push(preparation_warning(
                &document_part,
                "OFD template has no ID or BaseLoc",
            ));
            continue;
        };
        match resolve_internal_target(Some(&document_part), path) {
            Ok(part) => {
                templates.insert(id.to_owned(), part);
            }
            Err(message) => diagnostics.push(preparation_warning(&document_part, message)),
        }
    }
    let mut outline = Vec::new();
    if let Some(root) = document.descendants("Outlines").into_iter().next() {
        collect_outline(root, 0, &page_ids, &mut outline);
    }
    let mut annotation_parts = vec![Vec::new(); page_parts.len()];
    if let Some(path) = document
        .descendants("Annotations")
        .into_iter()
        .next()
        .map(|node| node.text.trim())
        .filter(|value| !value.is_empty())
    {
        let index_part = match resolve_internal_target(Some(&document_part), path) {
            Ok(part) => part,
            Err(message) => {
                diagnostics.push(preparation_warning(&document_part, message));
                String::new()
            }
        };
        let index = if index_part.is_empty() {
            None
        } else {
            match parse_part(&package, &index_part, limits) {
                Ok(index) => Some(index),
                Err(error) => {
                    diagnostics.push(preparation_warning(&index_part, error.message));
                    None
                }
            }
        };
        for page in index.iter().flat_map(|index| index.descendants("Page")) {
            let Some(page_index) = page
                .attribute("PageID")
                .and_then(|id| page_ids.iter().position(|candidate| candidate == id))
            else {
                continue;
            };
            for file in page.descendants("FileLoc") {
                let path = file.text.trim();
                if path.is_empty() {
                    continue;
                }
                match resolve_internal_target(Some(&index_part), path) {
                    Ok(part) => annotation_parts[page_index].push(part),
                    Err(message) => diagnostics.push(preparation_warning(&index_part, message)),
                }
            }
        }
    }

    let mut fonts = HashMap::new();
    let mut images = HashMap::new();
    let mut draw_params = HashMap::new();
    let mut composites = HashMap::new();
    let mut embedded_fonts = Vec::new();
    let mut embedded_font_bytes = 0usize;
    for resource in document
        .descendants("PublicRes")
        .into_iter()
        .chain(document.descendants("DocumentRes"))
    {
        let path = resource.text.trim();
        if path.is_empty() {
            continue;
        }
        let part = match resolve_internal_target(Some(&document_part), path) {
            Ok(part) => part,
            Err(message) => {
                diagnostics.push(preparation_warning(&document_part, message));
                continue;
            }
        };
        let resource = match parse_part(&package, &part, limits) {
            Ok(resource) => resource,
            Err(error) => {
                diagnostics.push(preparation_warning(&part, error.message));
                continue;
            }
        };
        let base = resource.attribute("BaseLoc").unwrap_or("");
        for draw_param in resource.descendants("DrawParam") {
            if let Some(id) = draw_param.attribute("ID") {
                draw_params.insert(id.to_owned(), draw_param.clone());
            }
        }
        for composite in resource.descendants("CompositeGraphicUnit") {
            if let Some(id) = composite.attribute("ID") {
                composites.insert(id.to_owned(), (part.clone(), composite.clone()));
            }
        }
        for font in resource.descendants("Font") {
            if let (Some(id), Some(name)) = (font.attribute("ID"), font.attribute("FontName")) {
                fonts.insert(id.to_owned(), name.to_owned());
                let Some(file) = font
                    .descendants("FontFile")
                    .into_iter()
                    .next()
                    .map(|node| node.text.trim())
                    .filter(|value| !value.is_empty())
                else {
                    continue;
                };
                let target = if base.is_empty() {
                    file.to_owned()
                } else {
                    format!("{base}/{file}")
                };
                let target = match resolve_internal_target(Some(&part), &target) {
                    Ok(target) => target,
                    Err(message) => {
                        diagnostics.push(preparation_warning(&part, message));
                        continue;
                    }
                };
                let bytes = match package.part(&target) {
                    Ok(Some(bytes)) => bytes.into_vec(),
                    Ok(None) => {
                        diagnostics.push(preparation_warning(
                            &target,
                            "OFD embedded font part is missing",
                        ));
                        continue;
                    }
                    Err(error) => {
                        diagnostics.push(preparation_warning(&target, error.message));
                        continue;
                    }
                };
                if !bytes
                    .get(..4)
                    .is_some_and(|signature| matches!(signature, b"\0\x01\0\0" | b"OTTO" | b"true"))
                {
                    diagnostics.push(preparation_warning(
                        &target,
                        "OFD embedded font has an unsupported browser font container",
                    ));
                    continue;
                }
                let Some(total) = embedded_font_bytes.checked_add(bytes.len()) else {
                    diagnostics.push(preparation_warning(
                        &target,
                        "OFD embedded font byte count overflowed",
                    ));
                    continue;
                };
                if total > limits.max_font_bytes {
                    diagnostics.push(preparation_warning(
                        &target,
                        "OFD embedded fonts exceed the configured font limit",
                    ));
                    continue;
                }
                embedded_font_bytes = total;
                embedded_fonts.push(EmbeddedFont {
                    family: name.to_owned(),
                    bytes,
                    style: FontStyle::Normal,
                    weight: 400,
                });
            }
        }
        for media in resource.descendants("MultiMedia") {
            let Some(id) = media.attribute("ID") else {
                continue;
            };
            let Some(file) = media
                .descendants("MediaFile")
                .into_iter()
                .next()
                .map(|node| node.text.trim())
                .filter(|value| !value.is_empty())
            else {
                continue;
            };
            let target = if base.is_empty() {
                file.to_owned()
            } else {
                format!("{base}/{file}")
            };
            let target = resolve_internal_target(Some(&part), &target)
                .map_err(|message| format_error(&part, message))?;
            images.insert(id.to_owned(), target);
        }
    }

    let page_sizes = page_parts
        .iter()
        .map(|part| {
            parse_part(&package, part, limits)
                .ok()
                .and_then(|page| {
                    page.descendants("PhysicalBox")
                        .into_iter()
                        .next()
                        .and_then(|node| parse_box(&node.text))
                })
                .filter(|bounds| validate_page_size(*bounds, part).is_ok())
                .unwrap_or(page_size)
        })
        .collect();
    Ok(Some(PreparedOfd {
        bytes: retained.map_or_else(
            || PreparedBytes::Owned(bytes.to_vec()),
            PreparedBytes::Retained,
        ),
        page_parts,
        templates,
        annotation_parts,
        outline,
        page_sizes,
        fonts,
        images,
        draw_params,
        composites,
        seals,
        embedded_fonts,
        diagnostics,
        limits,
    }))
}

impl PreparedOfd {
    pub(crate) fn materialize(&self, selected_unit: Option<usize>) -> Result<Document, Diagnostic> {
        let package = Package::open(&self.bytes, self.limits)?;
        let include_global = selected_unit.is_none_or(|unit| unit == 0);
        let mut document = Document {
            fatal: false,
            format: Some(DocumentFormat::Ofd),
            kind: Some(DocumentKind::Text),
            units: Vec::with_capacity(self.page_parts.len()),
            outline: if include_global {
                self.outline.clone()
            } else {
                Vec::new()
            },
            objects: Vec::new(),
            font_alternate_names: Vec::new(),
            embedded_fonts: if include_global {
                self.embedded_fonts.clone()
            } else {
                Vec::new()
            },
            diagnostics: if include_global {
                self.diagnostics.clone()
            } else {
                Vec::new()
            },
        };
        for (page_index, page_part) in self.page_parts.iter().enumerate() {
            document
                .units
                .push(page_unit(page_index, self.page_sizes[page_index]));
            if selected_unit.is_some_and(|selected| selected != page_index) {
                continue;
            }
            let page = match parse_part(&package, page_part, self.limits) {
                Ok(page) if page.local_name() == "Page" => page,
                Ok(_) => {
                    document
                        .diagnostics
                        .push(unsupported_page(page_part, "root is not Page"));
                    continue;
                }
                Err(error) => {
                    document
                        .diagnostics
                        .push(unsupported_page(page_part, &error.message));
                    continue;
                }
            };
            for template in page
                .descendants("Template")
                .into_iter()
                .filter(|template| template.attribute("ZOrder") != Some("Foreground"))
            {
                render_template(
                    &package,
                    template,
                    page_part,
                    page_index as u32,
                    &self.templates,
                    &self.fonts,
                    &self.images,
                    &self.draw_params,
                    &self.composites,
                    &mut document,
                );
            }
            render_page(
                &package,
                &page,
                page_part,
                page_index as u32,
                &self.fonts,
                &self.images,
                &self.draw_params,
                &self.composites,
                &mut document,
            );
            for template in page
                .descendants("Template")
                .into_iter()
                .filter(|template| template.attribute("ZOrder") == Some("Foreground"))
            {
                render_template(
                    &package,
                    template,
                    page_part,
                    page_index as u32,
                    &self.templates,
                    &self.fonts,
                    &self.images,
                    &self.draw_params,
                    &self.composites,
                    &mut document,
                );
            }
            for part in &self.annotation_parts[page_index] {
                match parse_part(&package, part, self.limits) {
                    Ok(annotation) => render_page(
                        &package,
                        &annotation,
                        part,
                        page_index as u32,
                        &self.fonts,
                        &self.images,
                        &self.draw_params,
                        &self.composites,
                        &mut document,
                    ),
                    Err(error) => document
                        .diagnostics
                        .push(unsupported_object(part, &error.message)),
                }
            }
            for seal in self
                .seals
                .iter()
                .filter(|seal| seal.unit_index == page_index)
            {
                render_seal(seal, page_index as u32, self.limits, &mut document);
            }
        }
        Ok(document)
    }
}

fn render_seal(seal: &SealAppearance, unit_index: u32, limits: Limits, document: &mut Document) {
    if object_limit_reached(document, limits.max_document_objects, &seal.source_part) {
        return;
    }
    let (kind, visual) = match &seal.payload {
        SealPayload::Ofd(bytes) => {
            let embedded = match prepare(bytes, limits).and_then(|prepared| {
                let mut prepared = prepared.ok_or_else(|| {
                    format_error(&seal.source_part, "electronic seal is not an OFD package")
                })?;
                // A seal appearance is content, not another signed-document trust boundary.
                // Dropping nested stamps also prevents recursively embedded seal packages.
                prepared.seals.clear();
                prepared.materialize(Some(0))
            }) {
                Ok(document) => document,
                Err(error) => {
                    document.diagnostics.push(preparation_warning(
                        &seal.source_part,
                        format!("OFD vector seal was omitted: {}", error.message),
                    ));
                    return;
                }
            };
            let Some(page) = embedded.units.first() else {
                return;
            };
            if page.width <= 0.0 || page.height <= 0.0 || embedded.objects.is_empty() {
                return;
            }
            let transform = AffineTransform {
                a: seal.bounds.width / page.width,
                b: 0.0,
                c: 0.0,
                d: seal.bounds.height / page.height,
                e: seal.bounds.x,
                f: seal.bounds.y,
            };
            let children = embedded
                .objects
                .into_iter()
                .filter(|object| object.unit_index == 0)
                .map(|object| VisualBrushChild {
                    bounds: object.bounds,
                    visual: object.visual,
                })
                .collect();
            (
                ObjectKind::Shape,
                Visual::Layer {
                    transform,
                    opacity: 1.0,
                    blend_mode: BlendMode::Normal,
                    visual: Box::new(Visual::Group { children }),
                },
            )
        }
        SealPayload::Image { media_type, bytes } => (
            ObjectKind::Image,
            // Raster seals are ink overlays (as in OFDRW), not opaque page images.
            // Keep the signed image bytes intact and use the shared compositor.
            Visual::Layer {
                transform: AffineTransform::IDENTITY,
                opacity: 1.0,
                blend_mode: BlendMode::Multiply,
                visual: Box::new(Visual::Image {
                    media_type: (*media_type).to_owned(),
                    bytes: bytes.to_vec(),
                    crop: Default::default(),
                }),
            },
        ),
    };
    let visual = match seal.clip {
        Some(clip) => Visual::Effect {
            shadow: None,
            clip: Some(Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo {
                        x: clip.x,
                        y: clip.y,
                    },
                    PathCommand::LineTo {
                        x: clip.x + clip.width,
                        y: clip.y,
                    },
                    PathCommand::LineTo {
                        x: clip.x + clip.width,
                        y: clip.y + clip.height,
                    },
                    PathCommand::LineTo {
                        x: clip.x,
                        y: clip.y + clip.height,
                    },
                    PathCommand::ClosePath,
                ],
            }),
            visual: Box::new(visual),
        },
        None => visual,
    };
    let numeric_id = document.objects.len() as u32;
    document.objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: object_stable_id(&seal.source_part, unit_index, "seal", "", numeric_id),
        parent_stable_id: None,
        kind,
        unit_index,
        bounds: seal.bounds,
        z: numeric_id as i32,
        text: None,
        source: SourceRef {
            part: seal.source_part.clone(),
            mapping: MappingQuality::Exact,
            locator: SourceLocator::Ofd {
                kind: "seal",
                path: "/Signature/SignedInfo/StampAnnot".to_owned(),
            },
        },
        visual,
    });
}

fn collect_outline(node: &Node, level: u32, page_ids: &[String], outline: &mut Vec<OutlineItem>) {
    for item in node
        .children
        .iter()
        .filter(|child| child.local_name() == "OutlineElem")
    {
        let destination = item
            .descendants("Dest")
            .into_iter()
            .find_map(|destination| destination.attribute("PageID"))
            .and_then(|id| page_ids.iter().position(|candidate| candidate == id));
        if let (Some(title), Some(unit_index)) = (item.attribute("Title"), destination)
            && !title.is_empty()
        {
            outline.push(OutlineItem {
                title: title.to_owned(),
                unit_index: unit_index as u32,
                level,
            });
        }
        collect_outline(item, level.saturating_add(1), page_ids, outline);
    }
}

#[allow(clippy::too_many_arguments)]
fn render_template(
    package: &Package<'_>,
    reference: &Node,
    page_part: &str,
    unit_index: u32,
    templates: &HashMap<String, String>,
    fonts: &HashMap<String, String>,
    images: &HashMap<String, String>,
    draw_params: &HashMap<String, Node>,
    composites: &HashMap<String, (String, Node)>,
    document: &mut Document,
) {
    let Some(id) = reference.attribute("TemplateID") else {
        document.diagnostics.push(unsupported_object(
            page_part,
            "OFD template reference has no TemplateID",
        ));
        return;
    };
    let Some(part) = templates.get(id) else {
        document.diagnostics.push(unsupported_object(
            page_part,
            "OFD template resource is missing",
        ));
        return;
    };
    match parse_part(package, part, package.limits()) {
        Ok(template) if template.local_name() == "Page" => {
            render_page(
                package,
                &template,
                part,
                unit_index,
                fonts,
                images,
                draw_params,
                composites,
                document,
            );
        }
        Ok(_) => document
            .diagnostics
            .push(unsupported_object(part, "OFD template root is not Page")),
        Err(error) => document
            .diagnostics
            .push(unsupported_object(part, &error.message)),
    }
}

const OBJECT_LIMIT_MESSAGE: &str = "OFD objects beyond the configured object limit were omitted";

fn object_limit_reached(document: &mut Document, limit: usize, part: &str) -> bool {
    if document.objects.len() < limit {
        return false;
    }
    if !document
        .diagnostics
        .iter()
        .any(|diagnostic| diagnostic.message == OBJECT_LIMIT_MESSAGE)
    {
        document.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ObjectLimit,
                Phase::Render,
                Fidelity::Omitted,
                OBJECT_LIMIT_MESSAGE,
            )
            .in_part(part),
        );
    }
    true
}

// Keep page-space culling bounds outside the transform and authored drawing bounds inside it.
fn transform_container_child(object: &mut Object, transform: AffineTransform) {
    let child = VisualBrushChild {
        bounds: object.bounds,
        visual: std::mem::replace(&mut object.visual, Visual::None),
    };
    object.bounds = transform_rect(transform, object.bounds);
    object.visual = Visual::Group {
        children: vec![VisualBrushChild {
            bounds: object.bounds,
            visual: Visual::Layer {
                transform,
                opacity: 1.0,
                blend_mode: BlendMode::Normal,
                visual: Box::new(Visual::Group {
                    children: vec![child],
                }),
            },
        }],
    };
}

fn render_page(
    package: &Package<'_>,
    page: &Node,
    page_part: &str,
    unit_index: u32,
    fonts: &HashMap<String, String>,
    images: &HashMap<String, String>,
    draw_params: &HashMap<String, Node>,
    composites: &HashMap<String, (String, Node)>,
    document: &mut Document,
) {
    fn visit(
        package: &Package<'_>,
        node: &Node,
        page_part: &str,
        unit_index: u32,
        fonts: &HashMap<String, String>,
        images: &HashMap<String, String>,
        draw_params: &HashMap<String, Node>,
        composites: &HashMap<String, (String, Node)>,
        inherited_draw_param: Option<&str>,
        composite_stack: &mut Vec<String>,
        document: &mut Document,
    ) {
        if object_limit_reached(document, package.limits().max_document_objects, page_part) {
            return;
        }
        let draw_param = node.attribute("DrawParam").or(inherited_draw_param);
        match node.local_name() {
            "TextObject" => render_text(
                node,
                page_part,
                unit_index,
                fonts,
                draw_params,
                inherited_draw_param,
                document,
            ),
            "ImageObject" => render_image(package, node, page_part, unit_index, images, document),
            "PathObject" => {
                let mut patterns = [None, None];
                for (index, name) in ["FillColor", "StrokeColor"].into_iter().enumerate() {
                    let Some(color) =
                        style_paint_node(node, inherited_draw_param, draw_params, name)
                    else {
                        continue;
                    };
                    let Some(pattern) = color
                        .children
                        .iter()
                        .find(|child| child.local_name() == "Pattern")
                    else {
                        continue;
                    };
                    // An explicit pattern must never fall back to the color's solid Value.
                    patterns[index] = Some(Paint::None);
                    if composite_stack.len() >= 64 {
                        document.diagnostics.push(unsupported_object(
                            page_part,
                            "OFD pattern nesting is too deep",
                        ));
                        continue;
                    }
                    let start = document.objects.len();
                    composite_stack.push(String::new());
                    for cell in pattern
                        .children
                        .iter()
                        .filter(|child| child.local_name() == "CellContent")
                    {
                        visit(
                            package,
                            cell,
                            page_part,
                            unit_index,
                            fonts,
                            images,
                            draw_params,
                            composites,
                            None,
                            composite_stack,
                            document,
                        );
                    }
                    composite_stack.pop();
                    let children = document
                        .objects
                        .drain(start..)
                        .map(|object| VisualBrushChild {
                            bounds: object.bounds,
                            visual: object.visual,
                        })
                        .collect();
                    match pattern_paint(pattern, color, node, children) {
                        Ok(paint) => patterns[index] = Some(paint),
                        Err(message) => document
                            .diagnostics
                            .push(unsupported_object(page_part, message)),
                    }
                }
                render_path(
                    node,
                    page_part,
                    unit_index,
                    draw_params,
                    inherited_draw_param,
                    patterns,
                    document,
                );
            }
            "CompositeObject" => {
                let Some(resource_id) = node.attribute("ResourceID") else {
                    document.diagnostics.push(unsupported_object(
                        page_part,
                        "OFD CompositeObject has no ResourceID",
                    ));
                    return;
                };
                let Some((resource_part, resource)) = composites.get(resource_id) else {
                    document.diagnostics.push(unsupported_object(
                        page_part,
                        "OFD composite graphic resource is missing",
                    ));
                    return;
                };
                let Some(bounds) = node.attribute("Boundary").and_then(parse_box) else {
                    document.diagnostics.push(unsupported_object(
                        page_part,
                        "OFD CompositeObject has no valid Boundary",
                    ));
                    return;
                };
                if composite_stack.len() >= 64
                    || composite_stack
                        .iter()
                        .any(|candidate| candidate == resource_id)
                {
                    document.diagnostics.push(unsupported_object(
                        page_part,
                        "OFD composite graphic nesting is cyclic or too deep",
                    ));
                    return;
                }
                let start = document.objects.len();
                composite_stack.push(resource_id.to_owned());
                for child in &resource.children {
                    visit(
                        package,
                        child,
                        resource_part,
                        unit_index,
                        fonts,
                        images,
                        draw_params,
                        composites,
                        draw_param,
                        composite_stack,
                        document,
                    );
                }
                composite_stack.pop();
                let source_id = node.attribute("ID").unwrap_or(resource_id);
                // Resource children are local to the composite, not already offset by its Boundary.
                let mut transform =
                    object_transform(node, Rect::default()).unwrap_or(AffineTransform::IDENTITY);
                transform.e += bounds.x;
                transform.f += bounds.y;
                let end = document.objects.len();
                for index in start..end {
                    transform_container_child(&mut document.objects[index], transform);
                    let visual =
                        std::mem::replace(&mut document.objects[index].visual, Visual::None);
                    let visual = apply_object_opacity(node, visual);
                    let visual = apply_object_clips(node, visual, page_part, document);
                    let object = &mut document.objects[index];
                    object.visual = visual;
                    object.stable_id = format!("ofd-composite-{source_id}-{}", object.stable_id);
                }
            }
            _ => {
                let start = document.objects.len();
                for child in &node.children {
                    visit(
                        package,
                        child,
                        page_part,
                        unit_index,
                        fonts,
                        images,
                        draw_params,
                        composites,
                        draw_param,
                        composite_stack,
                        document,
                    );
                }
                if node.local_name() == "Appearance" {
                    if let Some(bounds) = node.attribute("Boundary").and_then(parse_box) {
                        let transform = AffineTransform {
                            e: bounds.x,
                            f: bounds.y,
                            ..AffineTransform::IDENTITY
                        };
                        for object in &mut document.objects[start..] {
                            transform_container_child(object, transform);
                        }
                    } else {
                        document.diagnostics.push(unsupported_object(
                            page_part,
                            "OFD Appearance has no valid Boundary; content retained without translation",
                        ));
                    }
                }
            }
        }
    }
    let mut composite_stack = Vec::new();
    visit(
        package,
        page,
        page_part,
        unit_index,
        fonts,
        images,
        draw_params,
        composites,
        None,
        &mut composite_stack,
        document,
    );
}

// OFD cells map to the shared visual brush after adapter-specific step clamping
// and cell clipping. XPS viewbox/stretch defaults are not OFD pattern semantics.
fn pattern_paint(
    pattern: &Node,
    color: &Node,
    object: &Node,
    children: Vec<VisualBrushChild>,
) -> Result<Paint, &'static str> {
    if children.is_empty() {
        return Err("OFD pattern cell has no renderable content");
    }
    let dimension = |name: &str| {
        pattern
            .attribute(name)
            .and_then(|value| value.parse::<f32>().ok())
            .filter(|value| value.is_finite() && *value > 0.0 && *value <= MAX_PAGE_DIMENSION)
            .map(|value| value * MM_TO_CSS_PX)
    };
    let width = dimension("Width").ok_or("OFD pattern Width is invalid")?;
    let height = dimension("Height").ok_or("OFD pattern Height is invalid")?;
    let step = |name, minimum: f32| {
        if pattern.attribute(name).is_none() {
            return Ok(minimum);
        }
        pattern
            .attribute(name)
            .and_then(|value| value.parse::<f32>().ok())
            .filter(|value| value.is_finite() && *value <= MAX_PAGE_DIMENSION)
            .map(|value| (value * MM_TO_CSS_PX).max(minimum))
            .ok_or("OFD pattern step is invalid")
    };
    let viewport = Rect {
        x: 0.0,
        y: 0.0,
        width: step("XStep", width)?,
        height: step("YStep", height)?,
    };
    let bounds = object
        .attribute("Boundary")
        .and_then(parse_box)
        .ok_or("OFD pattern object Boundary is invalid")?;
    let mut transform = if pattern.attribute("CTM").is_some() {
        object_transform(pattern, Rect::default()).ok_or("OFD pattern CTM is invalid")?
    } else {
        AffineTransform::IDENTITY
    };
    match pattern.attribute("RelativeTo").unwrap_or("Object") {
        "Object" => {
            transform.e += bounds.x;
            transform.f += bounds.y;
        }
        // Page anchoring through a transformed container needs its inverse transform.
        _ => return Err("OFD page-relative pattern is unsupported; paint omitted"),
    }
    let tile_mode = match pattern.attribute("ReflectMethod").unwrap_or("Normal") {
        "Normal" => TileMode::Tile,
        "Column" => TileMode::FlipX,
        "Row" => TileMode::FlipY,
        "RowAndColumn" => TileMode::FlipXY,
        _ => return Err("OFD pattern ReflectMethod is invalid"),
    };
    Ok(Paint::Visual {
        viewbox: viewport,
        viewport,
        viewbox_relative: false,
        viewport_relative: false,
        tile_mode,
        stretch: StretchMode::None,
        alignment_x: 0.0,
        alignment_y: 0.0,
        transform,
        relative_transform: AffineTransform::IDENTITY,
        opacity: color
            .attribute("Alpha")
            .and_then(|value| value.parse::<u8>().ok())
            .unwrap_or(255) as f32
            / 255.0,
        children: vec![VisualBrushChild {
            bounds: Rect {
                x: 0.0,
                y: 0.0,
                width,
                height,
            },
            visual: Visual::Effect {
                shadow: None,
                clip: Some(Geometry::Rectangle),
                visual: Box::new(Visual::Group { children }),
            },
        }],
    })
}

fn render_path(
    node: &Node,
    page_part: &str,
    unit_index: u32,
    draw_params: &HashMap<String, Node>,
    inherited_draw_param: Option<&str>,
    patterns: [Option<Paint>; 2],
    document: &mut Document,
) {
    let Some(bounds) = node.attribute("Boundary").and_then(parse_box) else {
        document.diagnostics.push(unsupported_object(
            page_part,
            "OFD PathObject has no valid Boundary",
        ));
        return;
    };
    let Some(data) = node
        .children
        .iter()
        .find(|child| child.local_name() == "AbbreviatedData")
        .map(|value| value.text.as_str())
    else {
        document.diagnostics.push(unsupported_object(
            page_part,
            "OFD PathObject has no AbbreviatedData",
        ));
        return;
    };
    let commands = match parse_path(data) {
        Ok(commands) => commands,
        Err(message) => {
            document
                .diagnostics
                .push(unsupported_object(page_part, message));
            return;
        }
    };
    let paint = |name: &str, default| {
        patterns[usize::from(name == "StrokeColor")]
            .clone()
            .or_else(|| style_paint(node, inherited_draw_param, draw_params, name))
            .unwrap_or(Paint::Solid(default))
    };
    let (fill, stroke) = path_paint_flags(node);
    let numeric_id = document.objects.len() as u32;
    let source_id = node.attribute("ID").unwrap_or("");
    let fill_rule = if node.attribute("Rule") == Some("Even-Odd") {
        FillRule::EvenOdd
    } else {
        FillRule::NonZero
    };
    let stroke_width = style_value(node, inherited_draw_param, draw_params, "LineWidth")
        .as_deref()
        .and_then(|value| value.parse::<f32>().ok())
        .filter(|value| value.is_finite() && *value >= 0.0)
        .unwrap_or(0.353)
        * MM_TO_CSS_PX;
    let mesh = fill
        .then(|| style_paint_node(node, inherited_draw_param, draw_params, "FillColor"))
        .flatten()
        .and_then(|color| mesh_visual(color, bounds));
    let visual = if let Some(mesh) = mesh {
        let clipped = Visual::Effect {
            shadow: None,
            clip: Some(Geometry::Path {
                fill_rule,
                commands: commands.clone(),
            }),
            visual: Box::new(mesh),
        };
        if stroke {
            Visual::Group {
                children: vec![
                    VisualBrushChild {
                        bounds,
                        visual: clipped,
                    },
                    VisualBrushChild {
                        bounds,
                        visual: Visual::PaintedShape {
                            geometry: Geometry::Path {
                                fill_rule,
                                commands,
                            },
                            fill: Paint::None,
                            stroke: paint("StrokeColor", 0x0000_00ff),
                            stroke_width,
                        },
                    },
                ],
            }
        } else {
            clipped
        }
    } else {
        Visual::PaintedShape {
            geometry: Geometry::Path {
                fill_rule,
                commands,
            },
            fill: if fill {
                paint("FillColor", 0x0000_00ff)
            } else {
                Paint::None
            },
            stroke: if stroke {
                paint("StrokeColor", 0x0000_00ff)
            } else {
                Paint::None
            },
            stroke_width,
        }
    };
    let visual = apply_object_transform(
        node,
        bounds,
        apply_stroke_style(node, inherited_draw_param, draw_params, visual),
    );
    let visual = apply_object_clips(node, visual, page_part, document);
    document.objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: object_stable_id(page_part, unit_index, "path", source_id, numeric_id),
        parent_stable_id: None,
        kind: ObjectKind::Shape,
        unit_index,
        bounds,
        z: numeric_id as i32,
        text: None,
        source: SourceRef {
            part: page_part.to_owned(),
            mapping: MappingQuality::Exact,
            locator: SourceLocator::Ofd {
                kind: "path",
                path: format!("/Page/Content/PathObject[{source_id}]"),
            },
        },
        visual,
    });
}

fn path_paint_flags(node: &Node) -> (bool, bool) {
    (
        node.attribute("Fill") == Some("true"),
        node.attribute("Stroke") != Some("false"),
    )
}

fn draw_param_chain<'a>(
    node: &'a Node,
    inherited_draw_param: Option<&'a str>,
    draw_params: &'a HashMap<String, Node>,
) -> Vec<&'a Node> {
    let mut chain = Vec::new();
    let mut next = node.attribute("DrawParam").or(inherited_draw_param);
    let mut visited = HashSet::new();
    while let Some(id) = next {
        if !visited.insert(id.to_owned()) {
            break;
        }
        let Some(draw_param) = draw_params.get(id) else {
            break;
        };
        chain.push(draw_param);
        next = draw_param.attribute("Relative");
    }
    chain
}

fn style_value(
    node: &Node,
    inherited_draw_param: Option<&str>,
    draw_params: &HashMap<String, Node>,
    name: &str,
) -> Option<String> {
    node.attribute(name).map(str::to_owned).or_else(|| {
        draw_param_chain(node, inherited_draw_param, draw_params)
            .into_iter()
            .find_map(|draw_param| draw_param.attribute(name).map(str::to_owned))
    })
}

fn color_child(node: &Node, name: &str) -> Option<u32> {
    let color = node
        .children
        .iter()
        .find(|child| child.local_name() == name)?;
    parse_color(
        color.attribute("Value")?,
        color
            .attribute("Alpha")
            .and_then(|value| value.parse::<u8>().ok())
            .unwrap_or(255),
    )
}

fn gradient_stops(shading: &Node) -> Vec<GradientStop> {
    let mut stops = shading
        .children
        .iter()
        .filter(|child| child.local_name() == "Segment")
        .filter_map(|segment| {
            let offset = segment.attribute("Position")?.parse::<f32>().ok()?;
            if !offset.is_finite() {
                return None;
            }
            let color = segment
                .children
                .iter()
                .find(|child| child.local_name() == "Color")?;
            Some(GradientStop {
                offset: offset.clamp(0.0, 1.0),
                color: parse_color(
                    color.attribute("Value")?,
                    color
                        .attribute("Alpha")
                        .and_then(|value| value.parse::<u8>().ok())
                        .unwrap_or(255),
                )?,
            })
        })
        .collect::<Vec<_>>();
    stops.sort_by(|left, right| left.offset.total_cmp(&right.offset));
    if stops.len() == 1 {
        stops.push(GradientStop {
            offset: 1.0,
            color: stops[0].color,
        });
    }
    stops
}

fn color_node_paint(color: &Node) -> Option<Paint> {
    for shading in &color.children {
        let stops = gradient_stops(shading);
        if stops.len() < 2 {
            continue;
        }
        match shading.local_name() {
            "AxialShd" | "LinearShd" => {
                let [x0, y0] = parse_numbers::<2>(shading.attribute("StartPoint")?)?;
                let [x1, y1] = parse_numbers::<2>(shading.attribute("EndPoint")?)?;
                return Some(Paint::LinearGradient {
                    x0: x0 * MM_TO_CSS_PX,
                    y0: y0 * MM_TO_CSS_PX,
                    x1: x1 * MM_TO_CSS_PX,
                    y1: y1 * MM_TO_CSS_PX,
                    stops,
                });
            }
            "RadialShd" => {
                let [x0, y0] = parse_numbers::<2>(shading.attribute("StartPoint")?)?;
                let [x1, y1] = parse_numbers::<2>(shading.attribute("EndPoint")?)?;
                let r0 = shading
                    .attribute("StartRadius")
                    .and_then(|value| value.parse::<f32>().ok())
                    .filter(|value| value.is_finite() && *value >= 0.0)
                    .unwrap_or(0.0);
                let r1 = shading
                    .attribute("EndRadius")?
                    .parse::<f32>()
                    .ok()
                    .filter(|value| value.is_finite() && *value >= 0.0)?;
                return Some(Paint::RadialGradient {
                    x0: x0 * MM_TO_CSS_PX,
                    y0: y0 * MM_TO_CSS_PX,
                    r0: r0 * MM_TO_CSS_PX,
                    x1: x1 * MM_TO_CSS_PX,
                    y1: y1 * MM_TO_CSS_PX,
                    r1: r1 * MM_TO_CSS_PX,
                    stops,
                });
            }
            _ => {}
        }
    }
    Some(Paint::Solid(parse_color(
        color.attribute("Value")?,
        color
            .attribute("Alpha")
            .and_then(|value| value.parse::<u8>().ok())
            .unwrap_or(255),
    )?))
}

fn paint_child(node: &Node, name: &str) -> Option<Paint> {
    node.children
        .iter()
        .find(|child| child.local_name() == name)
        .and_then(color_node_paint)
}

fn style_paint(
    node: &Node,
    inherited_draw_param: Option<&str>,
    draw_params: &HashMap<String, Node>,
    name: &str,
) -> Option<Paint> {
    paint_child(node, name).or_else(|| {
        draw_param_chain(node, inherited_draw_param, draw_params)
            .into_iter()
            .find_map(|draw_param| paint_child(draw_param, name))
    })
}

fn style_paint_node<'a>(
    node: &'a Node,
    inherited_draw_param: Option<&'a str>,
    draw_params: &'a HashMap<String, Node>,
    name: &str,
) -> Option<&'a Node> {
    node.children
        .iter()
        .find(|child| child.local_name() == name)
        .or_else(|| {
            draw_param_chain(node, inherited_draw_param, draw_params)
                .into_iter()
                .find_map(|draw_param| {
                    draw_param
                        .children
                        .iter()
                        .find(|child| child.local_name() == name)
                })
        })
}

#[derive(Clone, Copy)]
struct MeshPoint {
    x: f32,
    y: f32,
    color: u32,
    edge: u8,
}

fn mesh_visual(color: &Node, bounds: Rect) -> Option<Visual> {
    let shading = color
        .children
        .iter()
        .find(|child| matches!(child.local_name(), "GouraudShd" | "LaGouraudShd"))?;
    let points = shading
        .children
        .iter()
        .filter(|child| child.local_name() == "Point")
        .take(4096)
        .filter_map(|point| {
            let color = point
                .children
                .iter()
                .find(|child| child.local_name() == "Color")?;
            Some(MeshPoint {
                x: point.attribute("X")?.parse::<f32>().ok()? * MM_TO_CSS_PX,
                y: point.attribute("Y")?.parse::<f32>().ok()? * MM_TO_CSS_PX,
                color: parse_color(
                    color.attribute("Value")?,
                    color
                        .attribute("Alpha")
                        .and_then(|value| value.parse::<u8>().ok())
                        .unwrap_or(255),
                )?,
                edge: point
                    .attribute("EdgeFlag")
                    .and_then(|value| value.parse::<u8>().ok())
                    .unwrap_or(0),
            })
        })
        .filter(|point| point.x.is_finite() && point.y.is_finite())
        .collect::<Vec<_>>();
    let triangles = if shading.local_name() == "LaGouraudShd" {
        lattice_mesh_triangles(
            &points,
            shading.attribute("VerticesPerRow")?.parse::<usize>().ok()?,
        )
    } else {
        gouraud_triangles(&points)
    };
    if triangles.is_empty() {
        return None;
    }
    let depth = if triangles.len() <= 64 {
        3
    } else if triangles.len() <= 512 {
        2
    } else if triangles.len() <= 4096 {
        1
    } else {
        0
    };
    let mut children = Vec::new();
    for [a, b, c] in triangles.into_iter().take(8192) {
        tessellate_mesh_triangle(a, b, c, depth, bounds, &mut children);
    }
    Some(Visual::Group { children })
}

fn lattice_mesh_triangles(points: &[MeshPoint], columns: usize) -> Vec<[MeshPoint; 3]> {
    if columns < 2 || points.len() < columns * 2 {
        return Vec::new();
    }
    let rows = points.len() / columns;
    let mut triangles = Vec::new();
    for row in 0..rows - 1 {
        for column in 0..columns - 1 {
            let top_left = points[row * columns + column];
            let top_right = points[row * columns + column + 1];
            let bottom_left = points[(row + 1) * columns + column];
            let bottom_right = points[(row + 1) * columns + column + 1];
            triangles.push([top_left, top_right, bottom_left]);
            triangles.push([top_right, bottom_right, bottom_left]);
        }
    }
    triangles
}

fn gouraud_triangles(points: &[MeshPoint]) -> Vec<[MeshPoint; 3]> {
    let mut triangles = Vec::new();
    let mut index = 0;
    while index + 2 < points.len() {
        let mut triangle = [points[index], points[index + 1], points[index + 2]];
        triangles.push(triangle);
        index += 3;
        while let Some(point) = points.get(index).copied().filter(|point| point.edge != 0) {
            triangle = if point.edge == 2 {
                [triangle[0], triangle[2], point]
            } else {
                [triangle[1], triangle[2], point]
            };
            triangles.push(triangle);
            index += 1;
        }
    }
    triangles
}

fn tessellate_mesh_triangle(
    a: MeshPoint,
    b: MeshPoint,
    c: MeshPoint,
    depth: u8,
    bounds: Rect,
    children: &mut Vec<VisualBrushChild>,
) {
    if depth > 0 {
        let ab = midpoint(a, b);
        let bc = midpoint(b, c);
        let ca = midpoint(c, a);
        for [a, b, c] in [[a, ab, ca], [ab, b, bc], [ca, bc, c], [ab, bc, ca]] {
            tessellate_mesh_triangle(a, b, c, depth - 1, bounds, children);
        }
        return;
    }
    children.push(VisualBrushChild {
        bounds,
        visual: Visual::PaintedShape {
            geometry: Geometry::Path {
                fill_rule: FillRule::NonZero,
                commands: vec![
                    PathCommand::MoveTo { x: a.x, y: a.y },
                    PathCommand::LineTo { x: b.x, y: b.y },
                    PathCommand::LineTo { x: c.x, y: c.y },
                    PathCommand::ClosePath,
                ],
            },
            fill: Paint::Solid(average_color(a.color, b.color, c.color)),
            stroke: Paint::None,
            stroke_width: 0.0,
        },
    });
}

fn midpoint(a: MeshPoint, b: MeshPoint) -> MeshPoint {
    MeshPoint {
        x: (a.x + b.x) / 2.0,
        y: (a.y + b.y) / 2.0,
        color: blend_color(a.color, b.color),
        edge: 0,
    }
}

fn blend_color(a: u32, b: u32) -> u32 {
    [24_u32, 16, 8, 0]
        .into_iter()
        .map(|shift| (((a >> shift) & 0xff) + ((b >> shift) & 0xff)) / 2)
        .zip([24_u32, 16, 8, 0])
        .fold(0, |color, (channel, shift)| color | channel << shift)
}

fn average_color(a: u32, b: u32, c: u32) -> u32 {
    [24_u32, 16, 8, 0]
        .into_iter()
        .map(|shift| (((a >> shift) & 0xff) + ((b >> shift) & 0xff) + ((c >> shift) & 0xff)) / 3)
        .zip([24_u32, 16, 8, 0])
        .fold(0, |color, (channel, shift)| color | channel << shift)
}

fn style_color(
    node: &Node,
    inherited_draw_param: Option<&str>,
    draw_params: &HashMap<String, Node>,
    name: &str,
) -> Option<u32> {
    color_child(node, name).or_else(|| {
        draw_param_chain(node, inherited_draw_param, draw_params)
            .into_iter()
            .find_map(|draw_param| color_child(draw_param, name))
    })
}

fn apply_stroke_style(
    node: &Node,
    inherited_draw_param: Option<&str>,
    draw_params: &HashMap<String, Node>,
    visual: Visual,
) -> Visual {
    let mut style = StrokeStyle::default();
    style.cap = match style_value(node, inherited_draw_param, draw_params, "Cap").as_deref() {
        Some("Round") => LineCap::Round,
        Some("Square") => LineCap::Square,
        _ => LineCap::Flat,
    };
    style.join = match style_value(node, inherited_draw_param, draw_params, "Join").as_deref() {
        Some("Round") => LineJoin::Round,
        Some("Bevel") => LineJoin::Bevel,
        _ => LineJoin::Miter,
    };
    style.miter_limit = style_value(node, inherited_draw_param, draw_params, "MiterLimit")
        .as_deref()
        .and_then(|value| value.parse::<f32>().ok())
        .filter(|value| value.is_finite() && *value >= 0.0)
        .unwrap_or(style.miter_limit);
    style.dash = style_value(node, inherited_draw_param, draw_params, "DashPattern")
        .as_deref()
        .map(parse_deltas)
        .unwrap_or_default();
    if style == StrokeStyle::default() {
        visual
    } else {
        Visual::StrokeStyle {
            style,
            visual: Box::new(visual),
        }
    }
}

fn parse_path(data: &str) -> Result<Vec<PathCommand>, &'static str> {
    let tokens = data.split_ascii_whitespace().collect::<Vec<_>>();
    let mut commands = Vec::new();
    let mut index = 0;
    let number = |tokens: &[&str], index: &mut usize| -> Result<f32, &'static str> {
        let value = tokens
            .get(*index)
            .ok_or("OFD path command is truncated")?
            .parse::<f32>()
            .map_err(|_| "OFD path coordinate is invalid")?;
        *index += 1;
        if !value.is_finite() {
            return Err("OFD path coordinate is not finite");
        }
        Ok(value * MM_TO_CSS_PX)
    };
    while index < tokens.len() {
        let command = tokens[index];
        index += 1;
        commands.push(match command {
            "M" | "S" => PathCommand::MoveTo {
                x: number(&tokens, &mut index)?,
                y: number(&tokens, &mut index)?,
            },
            "L" => PathCommand::LineTo {
                x: number(&tokens, &mut index)?,
                y: number(&tokens, &mut index)?,
            },
            "Q" => PathCommand::QuadraticCurveTo {
                cpx: number(&tokens, &mut index)?,
                cpy: number(&tokens, &mut index)?,
                x: number(&tokens, &mut index)?,
                y: number(&tokens, &mut index)?,
            },
            "B" => PathCommand::BezierCurveTo {
                cp1x: number(&tokens, &mut index)?,
                cp1y: number(&tokens, &mut index)?,
                cp2x: number(&tokens, &mut index)?,
                cp2y: number(&tokens, &mut index)?,
                x: number(&tokens, &mut index)?,
                y: number(&tokens, &mut index)?,
            },
            "C" => PathCommand::ClosePath,
            _ => return Err("OFD path command is unsupported"),
        });
    }
    if commands.is_empty() {
        return Err("OFD path has no commands");
    }
    Ok(commands)
}

fn parse_color(value: &str, alpha: u8) -> Option<u32> {
    let channels = value
        .split_ascii_whitespace()
        .map(|channel| {
            channel
                .strip_prefix('#')
                .map_or_else(|| channel.parse::<u8>(), |hex| u8::from_str_radix(hex, 16))
        })
        .collect::<Result<Vec<_>, _>>()
        .ok()?;
    let [red, green, blue] = channels.as_slice() else {
        return None;
    };
    Some((*red as u32) << 24 | (*green as u32) << 16 | (*blue as u32) << 8 | alpha as u32)
}

fn render_text(
    node: &Node,
    page_part: &str,
    unit_index: u32,
    fonts: &HashMap<String, String>,
    draw_params: &HashMap<String, Node>,
    inherited_draw_param: Option<&str>,
    document: &mut Document,
) {
    let Some(bounds) = node.attribute("Boundary").and_then(parse_box) else {
        document.diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::FormatInvalid,
                Phase::Parse,
                Fidelity::Omitted,
                "OFD TextObject has no valid Boundary",
            )
            .in_part(page_part),
        );
        return;
    };
    let text = node
        .descendants("TextCode")
        .into_iter()
        .map(|code| code.text.as_str())
        .collect::<String>();
    if text.is_empty() {
        return;
    }
    let numeric_id = document.objects.len() as u32;
    let source_id = node.attribute("ID").unwrap_or("");
    let font_family = node
        .attribute("Font")
        .and_then(|id| fonts.get(id))
        .cloned()
        .unwrap_or_else(|| "sans-serif".to_owned());
    let font_size = node
        .attribute("Size")
        .and_then(|value| value.parse::<f32>().ok())
        .filter(|value| value.is_finite() && *value > 0.0)
        .unwrap_or(3.0)
        * MM_TO_CSS_PX;
    let color =
        style_color(node, inherited_draw_param, draw_params, "FillColor").unwrap_or(0x0000_00ff);
    let bold = node
        .attribute("Weight")
        .and_then(|value| value.parse::<u16>().ok())
        .is_some_and(|weight| weight >= 600);
    let italic = node.attribute("Italic") == Some("true");
    let mut children =
        positioned_text_children(node, &font_family, font_size, color, bold, italic, bounds);
    let layout = TextLayout {
        inset_left: 0.0,
        inset_right: 0.0,
        inset_top: 0.0,
        inset_bottom: 0.0,
        wrap: false,
        text_baseline: font_size,
        ..TextLayout::default()
    };
    if let Some(paint @ (Paint::LinearGradient { .. } | Paint::RadialGradient { .. })) =
        style_paint(node, inherited_draw_param, draw_params, "FillColor")
    {
        // Positioned glyphs share the TextObject's gradient, not one gradient per glyph.
        for child in &mut children {
            let mut glyph_layout = layout.clone();
            glyph_layout.text_paint = paint
                .clone()
                .localize_gradient(child.bounds.x - bounds.x, child.bounds.y - bounds.y);
            child.visual = Visual::TextLayout {
                layout: glyph_layout,
                visual: Box::new(std::mem::replace(&mut child.visual, Visual::None)),
            };
        }
    }
    let visual = apply_object_transform(
        node,
        bounds,
        if children.is_empty() {
            Visual::Text {
                geometry: Geometry::Rectangle,
                fill: 0x0000_0000,
                stroke: 0x0000_0000,
                stroke_width: 0.0,
                font_family,
                font_size,
                color,
                bold,
                italic,
                align: TextAlign::Start,
            }
        } else {
            Visual::TextLayout {
                layout,
                visual: Box::new(Visual::Group { children }),
            }
        },
    );
    let visual = apply_object_clips(node, visual, page_part, document);
    document.objects.push(Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: object_stable_id(page_part, unit_index, "text", source_id, numeric_id),
        parent_stable_id: None,
        kind: ObjectKind::TextBox,
        unit_index,
        bounds,
        z: numeric_id as i32,
        text: Some(text),
        source: SourceRef {
            part: page_part.to_owned(),
            mapping: MappingQuality::Exact,
            locator: SourceLocator::Ofd {
                kind: "text",
                path: format!("/Page/Content/TextObject[{source_id}]"),
            },
        },
        visual,
    });
}

fn positioned_text_children(
    node: &Node,
    font_family: &str,
    font_size: f32,
    color: u32,
    bold: bool,
    italic: bool,
    bounds: Rect,
) -> Vec<VisualBrushChild> {
    let mut children = Vec::new();
    for code in node.descendants("TextCode") {
        let x = code
            .attribute("X")
            .and_then(|value| value.parse::<f32>().ok())
            .filter(|value| value.is_finite())
            .unwrap_or(0.0)
            * MM_TO_CSS_PX;
        let baseline = code
            .attribute("Y")
            .and_then(|value| value.parse::<f32>().ok())
            .filter(|value| value.is_finite())
            .unwrap_or(font_size / MM_TO_CSS_PX)
            * MM_TO_CSS_PX;
        let advances = parse_deltas(code.attribute("DeltaX").unwrap_or(""));
        let vertical_advances = parse_deltas(code.attribute("DeltaY").unwrap_or(""));
        let mut cursor = x;
        let mut cursor_y = baseline;
        for (index, character) in code.text.chars().enumerate() {
            children.push(VisualBrushChild {
                bounds: Rect {
                    x: bounds.x + cursor,
                    y: bounds.y + cursor_y - font_size,
                    width: font_size,
                    height: font_size * 1.2,
                },
                visual: Visual::RichText {
                    geometry: Geometry::Rectangle,
                    fill: Paint::None,
                    stroke: Paint::None,
                    stroke_width: 0.0,
                    align: TextAlign::Start,
                    line_height: font_size * 1.2,
                    runs: vec![TextRun {
                        paint: None,
                        east_asian_line_breaks: true,
                        text: character.to_string(),
                        font_family: font_family.to_owned(),
                        font_size,
                        color,
                        bold,
                        italic,
                        underline: false,
                        strikethrough: false,
                        highlight: 0,
                        baseline_shift: 0.0,
                        letter_spacing: 0.0,
                        horizontal_scale: 1.0,
                    }],
                },
            });
            cursor += advances
                .get(index)
                .or_else(|| advances.last())
                .copied()
                .unwrap_or(0.0);
            cursor_y += vertical_advances
                .get(index)
                .or_else(|| vertical_advances.last())
                .copied()
                .unwrap_or(0.0);
        }
    }
    children
}

fn parse_deltas(value: &str) -> Vec<f32> {
    let tokens = value.split_ascii_whitespace().collect::<Vec<_>>();
    let mut values = Vec::new();
    let mut index = 0;
    while index < tokens.len() {
        if tokens[index] == "g" {
            let Some(count) = tokens
                .get(index + 1)
                .and_then(|value| value.parse::<usize>().ok())
            else {
                break;
            };
            let Some(value) = tokens
                .get(index + 2)
                .and_then(|value| value.parse::<f32>().ok())
            else {
                break;
            };
            if !value.is_finite() || count > 100_000usize.saturating_sub(values.len()) {
                break;
            }
            values.extend(std::iter::repeat_n(value * MM_TO_CSS_PX, count));
            index += 3;
        } else {
            let Ok(value) = tokens[index].parse::<f32>() else {
                break;
            };
            if !value.is_finite() {
                break;
            }
            values.push(value * MM_TO_CSS_PX);
            index += 1;
        }
    }
    values
}

fn render_image(
    package: &Package<'_>,
    node: &Node,
    page_part: &str,
    unit_index: u32,
    images: &HashMap<String, String>,
    document: &mut Document,
) {
    let Some(authored_bounds) = node.attribute("Boundary").and_then(parse_box) else {
        document.diagnostics.push(unsupported_object(
            page_part,
            "OFD ImageObject has no valid Boundary",
        ));
        return;
    };
    let transform = object_transform(node, authored_bounds);
    let bounds = if transform.is_some() {
        Rect {
            x: authored_bounds.x,
            y: authored_bounds.y,
            width: MM_TO_CSS_PX,
            height: MM_TO_CSS_PX,
        }
    } else {
        authored_bounds
    };
    let Some(resource_id) = node.attribute("ResourceID") else {
        document.diagnostics.push(unsupported_object(
            page_part,
            "OFD ImageObject has no ResourceID",
        ));
        return;
    };
    let Some(part) = images.get(resource_id) else {
        document.diagnostics.push(unsupported_object(
            page_part,
            "OFD image resource is missing",
        ));
        return;
    };
    let Ok(Some(bytes)) = package.part(part) else {
        document.diagnostics.push(unsupported_object(
            part,
            "OFD image part is missing or invalid",
        ));
        return;
    };
    let Ok(media_type) = super::presentation_image::office_image_media_type(part, &bytes) else {
        document.diagnostics.push(unsupported_object(
            part,
            "OFD image encoding is unsupported",
        ));
        return;
    };
    let numeric_id = document.objects.len() as u32;
    let source_id = node.attribute("ID").unwrap_or("");
    let visual = apply_object_opacity(
        node,
        Visual::Image {
            media_type: media_type.to_owned(),
            bytes: bytes.into_vec(),
            crop: Default::default(),
        },
    );
    let mut object = Object {
        numeric_id,
        parent_numeric_id: None,
        stable_id: object_stable_id(page_part, unit_index, "image", source_id, numeric_id),
        parent_stable_id: None,
        kind: ObjectKind::Image,
        unit_index,
        bounds,
        z: numeric_id as i32,
        text: None,
        source: SourceRef {
            part: page_part.to_owned(),
            mapping: MappingQuality::Exact,
            locator: SourceLocator::Ofd {
                kind: "image",
                path: format!("/Page/Content/ImageObject[{source_id}]"),
            },
        },
        visual,
    };
    if let Some(transform) = transform {
        transform_container_child(&mut object, transform);
    }
    // Clips are relative to the authored Boundary, outside the image's unit-square CTM.
    let visual = apply_object_clips(node, object.visual, page_part, document);
    object.visual = Visual::Group {
        children: vec![VisualBrushChild {
            bounds: authored_bounds,
            visual,
        }],
    };
    document.objects.push(object);
}

fn apply_object_clips(
    node: &Node,
    mut visual: Visual,
    page_part: &str,
    document: &mut Document,
) -> Visual {
    let Some(clips) = node
        .children
        .iter()
        .find(|child| child.local_name() == "Clips")
    else {
        return visual;
    };
    let mut invalid = false;
    for clip in clips
        .children
        .iter()
        .filter(|child| child.local_name() == "Clip")
    {
        let mut commands = Vec::new();
        for area in clip
            .children
            .iter()
            .filter(|child| child.local_name() == "Area")
        {
            let transform = area
                .attribute("CTM")
                .and_then(parse_numbers::<6>)
                .unwrap_or([1.0, 0.0, 0.0, 1.0, 0.0, 0.0]);
            for path in area
                .children
                .iter()
                .filter(|child| child.local_name() == "Path")
            {
                let Some(bounds) = path.attribute("Boundary").and_then(parse_box) else {
                    invalid = true;
                    continue;
                };
                let Some(data) = path
                    .children
                    .iter()
                    .find(|child| child.local_name() == "AbbreviatedData")
                    .map(|child| child.text.as_str())
                else {
                    invalid = true;
                    continue;
                };
                match parse_path(data) {
                    Ok(path_commands) => commands.extend(transform_path_commands(
                        path_commands,
                        transform[0] * bounds.x + transform[2] * bounds.y,
                        transform[1] * bounds.x + transform[3] * bounds.y,
                        transform,
                    )),
                    Err(_) => invalid = true,
                }
            }
        }
        if clips.attribute("TransFlag") == Some("true") {
            if let Some(transform) = node.attribute("CTM").and_then(parse_numbers::<6>) {
                commands = transform_path_commands(commands, 0.0, 0.0, transform);
            }
        }
        if !commands.is_empty() {
            visual = Visual::Effect {
                shadow: None,
                clip: Some(Geometry::Path {
                    fill_rule: FillRule::NonZero,
                    commands,
                }),
                visual: Box::new(visual),
            };
        } else {
            invalid = true;
        }
    }
    if invalid {
        document.diagnostics.push(unsupported_object(
            page_part,
            "OFD clip contains an invalid or unsupported area",
        ));
    }
    visual
}

fn transform_path_commands(
    commands: Vec<PathCommand>,
    offset_x: f32,
    offset_y: f32,
    [a, b, c, d, e, f]: [f32; 6],
) -> Vec<PathCommand> {
    let transform = AffineTransform {
        a,
        b,
        c,
        d,
        e: offset_x + e * MM_TO_CSS_PX,
        f: offset_y + f * MM_TO_CSS_PX,
    };
    commands
        .into_iter()
        .map(|mut command| {
            command.transform(transform);
            command
        })
        .collect()
}

fn apply_object_transform(node: &Node, bounds: Rect, visual: Visual) -> Visual {
    let visual = apply_object_opacity(node, visual);
    let Some(transform) = object_transform(node, bounds) else {
        return visual;
    };
    // OFD Boundary is already in page coordinates; CTM affects drawing, not hit-test bounds.
    Visual::Group {
        children: vec![VisualBrushChild {
            bounds,
            visual: Visual::Layer {
                transform,
                opacity: 1.0,
                blend_mode: BlendMode::Normal,
                visual: Box::new(visual),
            },
        }],
    }
}

fn object_transform(node: &Node, bounds: Rect) -> Option<AffineTransform> {
    let [a, b, c, d, e, f] = node.attribute("CTM").and_then(parse_numbers::<6>)?;
    Some(AffineTransform {
        a,
        b,
        c,
        d,
        e: bounds.x + e * MM_TO_CSS_PX - a * bounds.x - c * bounds.y,
        f: bounds.y + f * MM_TO_CSS_PX - b * bounds.x - d * bounds.y,
    })
}

fn transform_rect(transform: AffineTransform, bounds: Rect) -> Rect {
    let corners = [
        (bounds.x, bounds.y),
        (bounds.x + bounds.width, bounds.y),
        (bounds.x, bounds.y + bounds.height),
        (bounds.x + bounds.width, bounds.y + bounds.height),
    ];
    let points = corners.map(|(x, y)| {
        (
            transform.a * x + transform.c * y + transform.e,
            transform.b * x + transform.d * y + transform.f,
        )
    });
    let min_x = points
        .iter()
        .map(|point| point.0)
        .fold(f32::INFINITY, f32::min);
    let max_x = points
        .iter()
        .map(|point| point.0)
        .fold(f32::NEG_INFINITY, f32::max);
    let min_y = points
        .iter()
        .map(|point| point.1)
        .fold(f32::INFINITY, f32::min);
    let max_y = points
        .iter()
        .map(|point| point.1)
        .fold(f32::NEG_INFINITY, f32::max);
    Rect {
        x: min_x,
        y: min_y,
        width: max_x - min_x,
        height: max_y - min_y,
    }
}

fn apply_object_opacity(node: &Node, visual: Visual) -> Visual {
    let opacity = node
        .attribute("Alpha")
        .and_then(|value| value.parse::<u8>().ok())
        .map_or(1.0, |value| value as f32 / 255.0);
    if opacity >= 1.0 {
        visual
    } else {
        Visual::Layer {
            transform: AffineTransform::IDENTITY,
            opacity,
            blend_mode: BlendMode::Normal,
            visual: Box::new(visual),
        }
    }
}

fn parse_numbers<const N: usize>(value: &str) -> Option<[f32; N]> {
    let values = value
        .split_ascii_whitespace()
        .map(str::parse::<f32>)
        .collect::<Result<Vec<_>, _>>()
        .ok()?;
    let values: [f32; N] = values.try_into().ok()?;
    values
        .iter()
        .all(|value| value.is_finite())
        .then_some(values)
}

fn page_unit(index: usize, bounds: Rect) -> Unit {
    Unit {
        kind: UnitKind::Page,
        index: index as u32,
        id: format!("ofd-page-{}", index + 1),
        name: format!("Page {}", index + 1),
        width: bounds.width,
        height: bounds.height,
        rows: 1,
        columns: 1,
        frozen_rows: 0,
        frozen_columns: 0,
        frozen_width: 0.0,
        frozen_height: 0.0,
        row_axis: SheetAxis::uniform(1, bounds.height),
        column_axis: SheetAxis::uniform(1, bounds.width),
        show_grid_lines: false,
        tab_color: None,
        sheet: None,
        slide: None,
    }
}

fn object_stable_id(
    part: &str,
    unit_index: u32,
    kind: &str,
    source_id: &str,
    numeric_id: u32,
) -> String {
    format!(
        "ofd-{}-{part}-{kind}-{}-{numeric_id}",
        unit_index + 1,
        if source_id.is_empty() {
            "anonymous"
        } else {
            source_id
        }
    )
}

fn parse_box(value: &str) -> Option<Rect> {
    let values = value
        .split_ascii_whitespace()
        .map(str::parse::<f32>)
        .collect::<Result<Vec<_>, _>>()
        .ok()?;
    let [x, y, width, height] = values.as_slice() else {
        return None;
    };
    let bounds = Rect {
        x: x * MM_TO_CSS_PX,
        y: y * MM_TO_CSS_PX,
        width: width * MM_TO_CSS_PX,
        height: height * MM_TO_CSS_PX,
    };
    bounds.is_valid().then_some(bounds)
}

fn validate_page_size(bounds: Rect, part: &str) -> Result<(), Diagnostic> {
    if bounds.width <= 0.0
        || bounds.height <= 0.0
        || bounds.width > MAX_PAGE_DIMENSION
        || bounds.height > MAX_PAGE_DIMENSION
    {
        return Err(format_error(part, "invalid OFD page dimensions"));
    }
    Ok(())
}

fn parse_part(package: &Package<'_>, part: &str, limits: Limits) -> Result<Node, Diagnostic> {
    let bytes = package.required_part(part)?;
    parse_xml_node(&bytes, part, limits)
}

fn parse_xml_node(bytes: &[u8], part: &str, limits: Limits) -> Result<Node, Diagnostic> {
    std::str::from_utf8(bytes).map_err(|_| format_error(part, "OFD XML encoding is not UTF-8"))?;
    let mut stack = Vec::<Node>::new();
    let mut root = None;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let mut node = Node {
                    name: name.to_owned(),
                    attributes: Vec::with_capacity(attributes.len()),
                    children: Vec::new(),
                    text: String::new(),
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
                    .ok_or_else(|| format_error(part, "unbalanced OFD XML element"))?;
                attach_node(&mut stack, &mut root, node, part)?;
            }
            XmlEvent::Text(value) => {
                if let Some(node) = stack.last_mut() {
                    node.text
                        .push_str(&decode_xml_text(value).map_err(|error| error.in_part(part))?);
                }
            }
            XmlEvent::Cdata(value) => {
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(value);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| error.in_part(part))?;
    root.ok_or_else(|| format_error(part, "OFD XML part has no root element"))
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
            "OFD XML contains multiple root elements",
        ));
    }
    Ok(())
}

fn unsupported_page(part: &str, detail: &str) -> Diagnostic {
    Diagnostic::warning(
        DiagnosticCode::FormatInvalid,
        Phase::Parse,
        Fidelity::Omitted,
        format!("OFD page was skipped: {detail}"),
    )
    .in_part(part)
}

fn preparation_warning(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::warning(
        DiagnosticCode::FormatInvalid,
        Phase::Parse,
        Fidelity::Omitted,
        message,
    )
    .in_part(part)
}

fn unsupported_object(part: &str, message: &str) -> Diagnostic {
    Diagnostic::warning(
        DiagnosticCode::UnsupportedFeature,
        Phase::Render,
        Fidelity::Omitted,
        message,
    )
    .in_part(part)
}

fn format_error(part: impl Into<String>, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message).in_part(part)
}

fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;

    use super::{
        MM_TO_CSS_PX, Node, SM2_WITH_SM3_OID, SignatureVerification, base64_digest,
        color_node_paint, detect_and_parse, find_embedded_ofd, object_limit_reached, parse_box,
        parse_color, parse_path, path_paint_flags, positioned_text_children, prepare, render_path,
        sm3_digest, style_value, verify_ses_sm2,
    };

    #[test]
    fn converts_ofd_millimetres_to_css_pixels() {
        let bounds = parse_box("0 0 210 297").unwrap();
        assert!((bounds.width - 210.0 * MM_TO_CSS_PX).abs() < 0.001);
        assert!((bounds.height - 297.0 * MM_TO_CSS_PX).abs() < 0.001);
        assert!(parse_box("0 0 -1 297").is_none());
    }

    #[test]
    fn keeps_each_ofd_page_size_in_lazy_unit_metadata() {
        let mut prepared = prepare(
            include_bytes!("../../tests/fixtures/ofdrw-h.ofd"),
            crate::limits::Limits::default(),
        )
        .unwrap()
        .unwrap();
        prepared.page_sizes[1] = parse_box("0 0 320 240").unwrap();

        let document = prepared.materialize(Some(0)).unwrap();

        assert_eq!(document.units[1].width, 320.0 * MM_TO_CSS_PX);
        assert_eq!(document.units[1].height, 240.0 * MM_TO_CSS_PX);
    }

    #[test]
    fn parses_ofd_path_commands_in_local_millimetres() {
        let commands = parse_path("M 0 0 L 20 0 L 20 10 L 0 10 C").unwrap();
        assert_eq!(commands.len(), 5);
        assert!(matches!(
            commands.last(),
            Some(crate::model::PathCommand::ClosePath)
        ));
    }

    #[test]
    fn parses_revised_ofd_path_start_and_hex_rgb_channels() {
        let commands = parse_path("S 0 0 L 156 0 C").unwrap();
        assert!(matches!(
            commands.first(),
            Some(crate::model::PathCommand::MoveTo { .. })
        ));
        assert_eq!(parse_color("#ee #20 #25", 255), Some(0xee20_25ff));
    }

    #[test]
    fn keeps_ofd_path_boundary_outside_ctm_for_scene_culling() {
        let path = Node {
            name: "PathObject".to_owned(),
            attributes: vec![
                ("Boundary".to_owned(), "155.41 105.63 12.8 14.73".to_owned()),
                (
                    "CTM".to_owned(),
                    "0.353 0 0 0.353 166.072 127.507".to_owned(),
                ),
                ("Fill".to_owned(), "true".to_owned()),
                ("Stroke".to_owned(), "false".to_owned()),
            ],
            children: vec![Node {
                name: "AbbreviatedData".to_owned(),
                attributes: Vec::new(),
                children: Vec::new(),
                text: "M -470.76 -361.44 L -434.47 -319.67 C".to_owned(),
            }],
            text: String::new(),
        };
        let mut document = crate::model::Document {
            fatal: false,
            format: None,
            kind: None,
            units: Vec::new(),
            outline: Vec::new(),
            objects: Vec::new(),
            embedded_fonts: Vec::new(),
            font_alternate_names: Vec::new(),
            diagnostics: Vec::new(),
        };
        render_path(
            &path,
            "Doc_0/Pages/Page_0/Content.xml",
            0,
            &HashMap::new(),
            None,
            [None, None],
            &mut document,
        );

        let crate::model::Visual::Group { children } = &document.objects[0].visual else {
            panic!("OFD path Boundary must remain the scene-culling bound");
        };
        assert_eq!(children.len(), 1);
        assert_eq!(children[0].bounds, document.objects[0].bounds);
        assert!(matches!(
            children[0].visual,
            crate::model::Visual::Layer { .. }
        ));
    }

    #[test]
    fn resolves_draw_param_child_before_relative_parent() {
        let node = Node {
            name: "PathObject".to_owned(),
            attributes: vec![("DrawParam".to_owned(), "child".to_owned())],
            children: Vec::new(),
            text: String::new(),
        };
        let mut draw_params = HashMap::new();
        draw_params.insert(
            "parent".to_owned(),
            Node {
                name: "DrawParam".to_owned(),
                attributes: vec![("LineWidth".to_owned(), "1".to_owned())],
                children: Vec::new(),
                text: String::new(),
            },
        );
        draw_params.insert(
            "child".to_owned(),
            Node {
                name: "DrawParam".to_owned(),
                attributes: vec![
                    ("Relative".to_owned(), "parent".to_owned()),
                    ("LineWidth".to_owned(), "2".to_owned()),
                ],
                children: Vec::new(),
                text: String::new(),
            },
        );
        assert_eq!(
            style_value(&node, None, &draw_params, "LineWidth").as_deref(),
            Some("2")
        );
    }

    #[test]
    fn maps_ofd_axial_shading_to_shared_gradient_paint() {
        let segment = |position: &str, value: &str| Node {
            name: "Segment".to_owned(),
            attributes: vec![("Position".to_owned(), position.to_owned())],
            children: vec![Node {
                name: "Color".to_owned(),
                attributes: vec![("Value".to_owned(), value.to_owned())],
                children: Vec::new(),
                text: String::new(),
            }],
            text: String::new(),
        };
        let color = Node {
            name: "FillColor".to_owned(),
            attributes: vec![("Value".to_owned(), "0 0 0".to_owned())],
            children: vec![Node {
                name: "AxialShd".to_owned(),
                attributes: vec![
                    ("StartPoint".to_owned(), "0 0".to_owned()),
                    ("EndPoint".to_owned(), "10 0".to_owned()),
                ],
                children: vec![segment("0", "255 0 0"), segment("1", "0 0 255")],
                text: String::new(),
            }],
            text: String::new(),
        };
        assert!(matches!(
            color_node_paint(&color),
            Some(crate::model::Paint::LinearGradient { ref stops, .. }) if stops.len() == 2
        ));
    }

    #[test]
    fn uses_ofd_path_fill_and_stroke_defaults() {
        let path = Node {
            name: "PathObject".to_owned(),
            attributes: Vec::new(),
            children: Vec::new(),
            text: String::new(),
        };
        assert_eq!(path_paint_flags(&path), (false, true));
    }

    #[test]
    fn keeps_delta_y_only_ofd_text_on_one_vertical_axis() {
        let text = Node {
            name: "TextObject".to_owned(),
            attributes: Vec::new(),
            children: vec![Node {
                name: "TextCode".to_owned(),
                attributes: vec![
                    ("X".to_owned(), "0.25".to_owned()),
                    ("Y".to_owned(), "3.15".to_owned()),
                    ("DeltaY".to_owned(), "6.35 6.35".to_owned()),
                ],
                children: Vec::new(),
                text: "购买方".to_owned(),
            }],
            text: String::new(),
        };
        let children = positioned_text_children(
            &text,
            "KaiTi",
            3.175 * MM_TO_CSS_PX,
            0x9c52_23ff,
            false,
            false,
            crate::model::Rect {
                x: 6.6 * MM_TO_CSS_PX,
                y: 32.8 * MM_TO_CSS_PX,
                width: 4.11 * MM_TO_CSS_PX,
                height: 17.2 * MM_TO_CSS_PX,
            },
        );
        assert_eq!(children.len(), 3);
        assert!(children.windows(2).all(|pair| {
            (pair[0].bounds.x - pair[1].bounds.x).abs() < 0.001
                && pair[1].bounds.y > pair[0].bounds.y
        }));
    }

    #[test]
    fn reports_the_ofd_object_limit_once() {
        let mut document = crate::model::Document {
            fatal: false,
            format: None,
            kind: None,
            units: Vec::new(),
            outline: Vec::new(),
            objects: Vec::new(),
            embedded_fonts: Vec::new(),
            font_alternate_names: Vec::new(),
            diagnostics: Vec::new(),
        };
        assert!(object_limit_reached(
            &mut document,
            0,
            "Doc_0/Pages/Page_0/Content.xml"
        ));
        assert!(object_limit_reached(
            &mut document,
            0,
            "Doc_0/Pages/Page_0/Content.xml"
        ));
        assert_eq!(document.diagnostics.len(), 1);
        assert_eq!(
            document.diagnostics[0].code,
            crate::diagnostic::DiagnosticCode::ObjectLimit
        );
    }

    #[test]
    fn computes_standard_sm3_and_ofd_reference_value() {
        let digest = sm3_digest(b"abc");
        assert_eq!(
            digest,
            [
                0x66, 0xc7, 0xf0, 0xf4, 0x62, 0xee, 0xed, 0xd9, 0xd1, 0xf2, 0xd4, 0x6b, 0xdc, 0x10,
                0xe4, 0xe2, 0x41, 0x67, 0xc4, 0x87, 0x5c, 0xf2, 0xf7, 0xa2, 0x29, 0x7d, 0xa0, 0x2b,
                0x8f, 0x4b, 0xa8, 0xe0,
            ]
        );
        assert_eq!(
            base64_digest(digest),
            "Zsfw9GLu7dnR8tRr3BDk4kFnxIdc8veiKX2gK49LqOA="
        );
    }

    #[test]
    fn extracts_an_ofd_vector_seal_from_bounded_der() {
        let der = [0x30, 0x06, 0x04, 0x04, b'P', b'K', 3, 4];
        assert_eq!(find_embedded_ofd(&der), Some(&der[4..]));
    }

    #[test]
    fn pattern_steps_and_cell_clip_preserve_ofd_boundaries() {
        use crate::model::{Geometry, Paint, Rect, TileMode, Visual, VisualBrushChild};
        let parse = |xml: &str| {
            super::parse_xml_node(xml.as_bytes(), "test", crate::limits::Limits::default()).unwrap()
        };
        let color = parse("<FillColor Value='0 0 0' Alpha='128'/>");
        let object = parse("<PathObject Boundary='5 7 50 50'/>");
        let children = || {
            vec![VisualBrushChild {
                bounds: Rect::default(),
                visual: Visual::None,
            }]
        };
        for (method, expected) in [
            ("Normal", TileMode::Tile),
            ("Column", TileMode::FlipX),
            ("Row", TileMode::FlipY),
            ("RowAndColumn", TileMode::FlipXY),
        ] {
            let pattern = parse(&format!(
                "<Pattern Width='10' Height='20' XStep='-1' YStep='30' ReflectMethod='{method}'/>"
            ));
            let Paint::Visual {
                viewport,
                transform,
                tile_mode,
                opacity,
                children,
                ..
            } = super::pattern_paint(&pattern, &color, &object, children()).unwrap()
            else {
                panic!("visual brush");
            };
            assert_eq!(viewport.width, 10.0 * MM_TO_CSS_PX);
            assert_eq!(viewport.height, 30.0 * MM_TO_CSS_PX);
            assert_eq!(transform.e, 5.0 * MM_TO_CSS_PX);
            assert_eq!(transform.f, 7.0 * MM_TO_CSS_PX);
            assert_eq!(tile_mode, expected);
            assert_eq!(opacity, 128.0 / 255.0);
            assert_eq!(children[0].bounds.height, 20.0 * MM_TO_CSS_PX);
            assert!(matches!(
                children[0].visual,
                Visual::Effect {
                    clip: Some(Geometry::Rectangle),
                    ..
                }
            ));
        }
        for attributes in [
            "Width='0' Height='20'",
            "Width='10' Height='20' XStep='NaN'",
            "Width='10' Height='20' CTM='broken'",
            "Width='10' Height='20' RelativeTo='Page'",
        ] {
            let pattern = parse(&format!("<Pattern {attributes}/>"));
            assert!(super::pattern_paint(&pattern, &color, &object, children()).is_err());
        }
    }

    #[test]
    fn supplied_pattern_watermark_is_a_translucent_text_tile_not_a_black_rectangle() {
        use crate::model::{Paint, Visual};
        fn fill(visual: &Visual) -> &Paint {
            match visual {
                Visual::Group { children } => fill(&children[0].visual),
                Visual::Layer { visual, .. } | Visual::StrokeStyle { visual, .. } => fill(visual),
                Visual::PaintedShape { fill, .. } => fill,
                _ => panic!("unexpected watermark visual"),
            }
        }
        let document = detect_and_parse(
            include_bytes!("../../tests/fixtures/ofdrw-pattern.ofd"),
            crate::limits::Limits::default(),
        )
        .unwrap()
        .unwrap();
        let object = document
            .objects
            .iter()
            .find(|o| o.stable_id.contains("-path-99-"))
            .unwrap();
        let Paint::Visual {
            viewport,
            transform,
            children,
            ..
        } = fill(&object.visual)
        else {
            panic!("Pattern watermark became a solid black rectangle");
        };
        assert!((viewport.width - 127.0079 * MM_TO_CSS_PX).abs() < 0.001);
        assert!((viewport.height - 66.2257 * MM_TO_CSS_PX).abs() < 0.001);
        assert!((transform.b + std::f32::consts::FRAC_1_SQRT_2).abs() < 0.0001);
        assert!(format!("{children:?}").contains("严"));
        assert!(format!("{children:?}").contains("opacity: 0.2"));
    }

    #[test]
    fn preserves_supplied_signout_pen_coordinates_through_composites() {
        use crate::model::{Geometry, PathCommand, Rect, Visual};
        fn first_point(visual: &Visual, bounds: Rect) -> (f32, f32) {
            match visual {
                Visual::Group { children } => first_point(&children[0].visual, children[0].bounds),
                Visual::Layer {
                    transform: t,
                    visual,
                    ..
                } => {
                    let (x, y) = first_point(visual, bounds);
                    (t.a * x + t.c * y + t.e, t.b * x + t.d * y + t.f)
                }
                Visual::StrokeStyle { visual, .. } => first_point(visual, bounds),
                Visual::PaintedShape {
                    geometry: Geometry::Path { commands, .. },
                    ..
                } => {
                    let PathCommand::MoveTo { x, y } = commands[0] else {
                        panic!("pen start");
                    };
                    (bounds.x + x, bounds.y + y)
                }
                _ => panic!("unexpected pen visual"),
            }
        }
        let document = detect_and_parse(
            include_bytes!("../../tests/fixtures/ofdrw-signout.ofd"),
            crate::limits::Limits::default(),
        )
        .unwrap()
        .unwrap();
        // Authored Appearance offset + 0.3528 * (Path Boundary + first MoveTo).
        for (id, x, y) in [
            (396, 75.74452, 148.30722),
            (477, 102.46608, 134.0024),
            (562, 70.85318, 133.93484),
        ] {
            let object = document
                .objects
                .iter()
                .find(|o| o.stable_id.contains(&format!("-path-{id}-")))
                .unwrap();
            let point = first_point(&object.visual, object.bounds);
            assert!(
                (point.0 - x * MM_TO_CSS_PX).abs() < 0.002,
                "pen {id}: {point:?}"
            );
            assert!(
                (point.1 - y * MM_TO_CSS_PX).abs() < 0.002,
                "pen {id}: {point:?}"
            );
        }
    }

    #[test]
    fn positions_supplied_annotation_content_in_page_coordinates() {
        let document = detect_and_parse(
            include_bytes!("../../tests/fixtures/ofdrw-20240531141733.ofd"),
            crate::limits::Limits::default(),
        )
        .unwrap()
        .unwrap();
        let annotations = document
            .objects
            .iter()
            .filter(|object| object.source.part == "Doc_0/Annots/Page_0/Annotation.xml")
            .collect::<Vec<_>>();
        assert_eq!(annotations.len(), 2);
        for (object, x, y) in [
            (annotations[0], 90.0, 8.0),
            (annotations[1], 206.4583, 29.5417),
        ] {
            assert!(
                (object.bounds.x - x * MM_TO_CSS_PX).abs() < 0.001,
                "annotation x: {}",
                object.bounds.x
            );
            assert!((object.bounds.y - y * MM_TO_CSS_PX).abs() < 0.001);
            let crate::model::Visual::Group { children } = &object.visual else {
                panic!("annotation page bounds must stay outside the transform");
            };
            assert_eq!(children[0].bounds, object.bounds);
            let crate::model::Visual::Layer {
                transform, visual, ..
            } = &children[0].visual
            else {
                panic!("annotation must translate its visual as well as its bounds");
            };
            let crate::model::Visual::Group { children } = visual.as_ref() else {
                panic!("annotation must preserve child-local bounds");
            };
            assert_eq!(children.len(), 1);
            assert_eq!(children[0].bounds.x, 0.0);
            assert_eq!(children[0].bounds.y, 0.0);
            assert_eq!(transform.e, x * MM_TO_CSS_PX);
            assert_eq!(transform.f, y * MM_TO_CSS_PX);
        }
    }

    #[test]
    fn supplied_timeline_clip_transforms_its_path_boundary_with_the_area() {
        use crate::model::{Geometry, PathCommand, Visual};
        let document = detect_and_parse(
            include_bytes!("../../tests/fixtures/ofdrw-intro.ofd"),
            crate::limits::Limits::default(),
        )
        .unwrap()
        .unwrap();
        let line = document
            .objects
            .iter()
            .find(|object| object.stable_id.contains("-path-273-"))
            .unwrap();
        let Visual::Effect {
            clip: Some(Geometry::Path { commands, .. }),
            ..
        } = &line.visual
        else {
            panic!("timeline clip missing");
        };
        // The Area CTM scales both the Path Boundary origin and its path coordinates.
        // With the owning object's Boundary, the clip covers the 320 x 240 mm page.
        let points: Vec<_> = commands
            .iter()
            .filter_map(|command| match command {
                PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } => Some((
                    (line.bounds.x + x) / MM_TO_CSS_PX,
                    (line.bounds.y + y) / MM_TO_CSS_PX,
                )),
                _ => None,
            })
            .collect();
        for (actual, expected) in [
            (
                points.iter().map(|p| p.0).fold(f32::INFINITY, f32::min),
                0.0,
            ),
            (
                points.iter().map(|p| p.1).fold(f32::INFINITY, f32::min),
                0.0,
            ),
            (
                points.iter().map(|p| p.0).fold(f32::NEG_INFINITY, f32::max),
                320.0,
            ),
            (
                points.iter().map(|p| p.1).fold(f32::NEG_INFINITY, f32::max),
                240.0,
            ),
        ] {
            assert!(
                (actual - expected).abs() < 0.05,
                "clip edge {actual}, expected {expected}"
            );
        }
    }

    #[test]
    fn supplied_image_clip_stays_in_boundary_space_unless_transformed() {
        use crate::model::{Geometry, PathCommand, Visual};
        let bytes = include_bytes!("../../tests/fixtures/ofdrw-y.ofd");
        let limits = crate::limits::Limits::default();
        let package = super::Package::open(bytes, limits).unwrap();
        let resource = super::parse_part(&package, "Doc_0/DocumentRes.xml", limits).unwrap();
        let image = resource
            .descendants("ImageObject")
            .into_iter()
            .find(|node| node.attribute("ID") == Some("1690"))
            .unwrap();
        let mut document = detect_and_parse(bytes, limits).unwrap().unwrap();
        for flag in [None, Some("false"), Some("true")] {
            let mut node = image.clone();
            let clips = node
                .children
                .iter_mut()
                .find(|child| child.local_name() == "Clips")
                .unwrap();
            clips.attributes.retain(|(name, _)| name != "TransFlag");
            if let Some(flag) = flag {
                clips.attributes.push(("TransFlag".into(), flag.into()));
            }
            let visual = super::apply_object_clips(&node, Visual::None, "test", &mut document);
            let Visual::Effect {
                clip: Some(Geometry::Path { commands, .. }),
                ..
            } = visual
            else {
                panic!("image clip missing");
            };
            let PathCommand::MoveTo { x, y } = commands[0] else {
                panic!("clip origin");
            };
            let (sx, sy) = if flag == Some("true") {
                (346.8615, 256.3842)
            } else {
                (1.0, 1.0)
            };
            assert!((x / MM_TO_CSS_PX - 1.0348 * sx).abs() < 0.01);
            assert!((y / MM_TO_CSS_PX - 255.4348 * sy).abs() < 0.02);
        }
    }

    #[test]
    fn applies_supplied_ofd_image_ctm_instead_of_stretching_to_its_boundary() {
        let document = detect_and_parse(
            include_bytes!("../../tests/fixtures/ofdrw-testImageOverridePage.ofd"),
            crate::limits::Limits::default(),
        )
        .unwrap()
        .unwrap();
        let image = document
            .objects
            .iter()
            .find(|object| object.stable_id.contains("-image-20-"))
            .unwrap();

        assert!(matches!(image.visual, crate::model::Visual::Group { .. }));
        assert!((image.bounds.width / MM_TO_CSS_PX - 29.982586).abs() < 0.001);
        assert!((image.bounds.height / MM_TO_CSS_PX - 29.982586).abs() < 0.001);
    }

    #[test]
    fn supplied_intro_keeps_gradient_text_visible_on_black() {
        use crate::model::{Paint, Rect, Visual};
        fn gradient_origins(visual: &Visual, bounds: Rect, origins: &mut Vec<(f32, f32)>) {
            match visual {
                Visual::TextLayout { layout, visual } => {
                    if let Paint::LinearGradient { x0, y0, .. } = &layout.text_paint {
                        origins.push((bounds.x + x0, bounds.y + y0));
                    }
                    gradient_origins(visual, bounds, origins);
                }
                Visual::Layer { visual, .. } => gradient_origins(visual, bounds, origins),
                Visual::Group { children } => {
                    for child in children {
                        gradient_origins(&child.visual, child.bounds, origins);
                    }
                }
                _ => {}
            }
        }
        let prepared = super::prepare(
            include_bytes!("../../tests/fixtures/ofdrw-intro.ofd"),
            crate::limits::Limits::default(),
        )
        .unwrap()
        .unwrap();
        let document = prepared.materialize(Some(4)).unwrap();
        let background = document
            .objects
            .iter()
            .find(|o| o.stable_id.contains("-path-127-"))
            .unwrap();
        assert!(
            (background.bounds.width - document.units[4].width).abs() < 0.01,
            "black background must cover the authored page, not be transformed twice: {:?}",
            background.bounds
        );

        for id in [187, 188, 189] {
            let object = document
                .objects
                .iter()
                .find(|o| o.stable_id.contains(&format!("-text-{id}-")))
                .unwrap();
            let mut origins = Vec::new();
            gradient_origins(&object.visual, object.bounds, &mut origins);
            assert_eq!(
                origins.len(),
                object.text.as_ref().unwrap().chars().count(),
                "text {id} lost its gradient"
            );
            assert!(
                origins
                    .iter()
                    .all(|&(x, y)| (x - origins[0].0).abs() < 0.001
                        && (y - origins[0].1).abs() < 0.001),
                "gradient must remain continuous across positioned glyphs"
            );
        }
    }

    #[test]
    fn supplied_opaque_seal_uses_multiply_without_changing_image_bytes() {
        use crate::model::{BlendMode, Visual};
        let prepared = super::prepare(
            include_bytes!("../../tests/fixtures/ofdrw-resource-path.ofd"),
            crate::limits::Limits::default(),
        )
        .unwrap()
        .unwrap();
        let super::SealPayload::Image { bytes, .. } = &prepared.seals[0].payload else {
            panic!("raster seal");
        };
        let original = bytes.clone();
        let document = prepared.materialize(Some(0)).unwrap();
        let seal = document
            .objects
            .iter()
            .find(|o| {
                matches!(
                    o.source.locator,
                    crate::model::SourceLocator::Ofd { kind: "seal", .. }
                )
            })
            .unwrap();
        let Visual::Effect { visual, .. } = &seal.visual else {
            panic!("seal clip");
        };
        let Visual::Layer {
            blend_mode: BlendMode::Multiply,
            visual,
            ..
        } = visual.as_ref()
        else {
            panic!("opaque white seal background covers the document instead of multiplying");
        };
        let Visual::Image { bytes, .. } = visual.as_ref() else {
            panic!("seal image");
        };
        assert_eq!(bytes.as_slice(), original.as_ref());
        fn ordinary_image(visual: &Visual) -> bool {
            match visual {
                Visual::Image { .. } => true,
                Visual::Layer {
                    blend_mode: BlendMode::Normal,
                    visual,
                    ..
                } => ordinary_image(visual),
                Visual::Group { children } => {
                    children.iter().any(|child| ordinary_image(&child.visual))
                }
                _ => false,
            }
        }
        assert!(document.objects.iter().any(|o| ordinary_image(&o.visual)));
    }

    #[test]
    fn renders_raster_and_clipped_seals_from_supplied_ofdrw_document() {
        let document = detect_and_parse(
            include_bytes!("../../tests/fixtures/ofdrw-h.ofd"),
            crate::limits::Limits::default(),
        )
        .unwrap()
        .unwrap();
        let seals = document
            .objects
            .iter()
            .filter(|object| {
                matches!(
                    &object.source.locator,
                    crate::model::SourceLocator::Ofd { kind, .. } if *kind == "seal"
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(seals.len(), 6);
        assert_eq!(seals.iter().filter(|seal| seal.unit_index == 0).count(), 2);
        assert_eq!(
            seals
                .iter()
                .filter(|seal| matches!(
                    &seal.visual,
                    crate::model::Visual::Effect {
                        clip: Some(crate::model::Geometry::Path { .. }),
                        ..
                    }
                ))
                .count(),
            5
        );
        assert_eq!(
            seals
                .iter()
                .filter(|seal| matches!(&seal.visual, crate::model::Visual::Layer { blend_mode: crate::model::BlendMode::Multiply, visual, .. } if matches!(visual.as_ref(), crate::model::Visual::Image { .. })))
                .count(),
            1
        );
    }

    #[test]
    fn verifies_and_rejects_a_complete_ses_sm2_container() {
        use sm2::dsa::{Signature, SigningKey, signature::Signer};
        use sm2::elliptic_curve::sec1::ToEncodedPoint;

        fn der(tag: u8, content: &[u8]) -> Vec<u8> {
            let length = if content.len() < 128 {
                vec![content.len() as u8]
            } else {
                vec![0x82, (content.len() >> 8) as u8, content.len() as u8]
            };
            [vec![tag], length, content.to_vec()].concat()
        }
        fn sequence(values: &[Vec<u8>]) -> Vec<u8> {
            der(0x30, &values.concat())
        }
        fn positive_integer(value: &[u8]) -> Vec<u8> {
            let first = value
                .iter()
                .position(|byte| *byte != 0)
                .unwrap_or(value.len() - 1);
            let value = &value[first..];
            if value[0] & 0x80 == 0 {
                der(0x02, value)
            } else {
                der(0x02, &[vec![0], value.to_vec()].concat())
            }
        }

        let signature_xml = b"<Signature/>";
        let digest = sm3_digest(signature_xml);
        let tbs = sequence(&[
            der(0x02, &[1]),
            der(0x06, SM2_WITH_SM3_OID),
            der(0x04, &digest),
        ]);
        let signing_key = SigningKey::from_slice("1234567812345678", &[1; 32]).unwrap();
        let signature: Signature = signing_key.sign(&tbs);
        let signature = sequence(&[
            positive_integer(&signature.r_bytes()),
            positive_integer(&signature.s_bytes()),
        ]);
        let public_key = signing_key
            .verifying_key()
            .as_affine()
            .to_encoded_point(false);
        let certificate = sequence(&[der(
            0x03,
            &[vec![0], public_key.as_bytes().to_vec()].concat(),
        )]);
        let signed_value = sequence(&[
            tbs,
            der(0x04, &certificate),
            der(0x06, SM2_WITH_SM3_OID),
            der(0x03, &[vec![0], signature].concat()),
        ]);
        assert!(matches!(
            verify_ses_sm2(&signed_value, signature_xml),
            SignatureVerification::Valid
        ));
        assert!(matches!(
            verify_ses_sm2(&signed_value, b"<tampered/>"),
            SignatureVerification::Invalid
        ));
    }
}
