//! OPC package access built on the bounded ZIP and XML primitives.

use std::collections::{HashMap, HashSet};
use std::ops::Deref;
use std::sync::{Arc, Mutex};

use crate::diagnostic::{Diagnostic, DiagnosticCode, Phase};
use crate::limits::Limits;
use crate::xml::{XmlAttribute, XmlEvent, decode_xml_text, parse_xml};
use crate::zip::ZipArchive;

/// A resolved OPC relationship. External targets remain opaque and are never
/// used as package-part names.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Relationship {
    pub id: String,
    pub type_uri: String,
    pub target: String,
    pub external: bool,
}

#[derive(Clone, Debug)]
pub enum PartBytes {
    Owned(Vec<u8>),
    Shared(Arc<[u8]>),
}

impl PartBytes {
    pub fn into_vec(self) -> Vec<u8> {
        match self {
            Self::Owned(bytes) => bytes,
            Self::Shared(bytes) => bytes.as_ref().to_vec(),
        }
    }
}

impl Deref for PartBytes {
    type Target = [u8];

    fn deref(&self) -> &Self::Target {
        match self {
            Self::Owned(bytes) => bytes,
            Self::Shared(bytes) => bytes,
        }
    }
}

/// Lazily extracts bounded parts from an untrusted OPC ZIP package.
#[derive(Debug)]
pub struct Package<'a> {
    archive: ZipArchive<'a>,
    limits: Limits,
    cache: PackageCache,
}

#[derive(Clone, Debug, Default)]
pub(crate) struct PackageCache {
    parts: Arc<Mutex<HashMap<String, Arc<[u8]>>>>,
    bytes: Arc<Mutex<usize>>,
}

impl<'a> Package<'a> {
    pub fn open(bytes: &'a [u8], limits: Limits) -> Result<Self, Diagnostic> {
        Ok(Self {
            archive: ZipArchive::parse(bytes, limits)?,
            limits,
            cache: PackageCache::default(),
        })
    }

    #[cfg(any(feature = "native-formats", feature = "xps-formats"))]
    pub(crate) fn open_with_cache(
        bytes: &'a [u8],
        limits: Limits,
        cache: PackageCache,
    ) -> Result<Self, Diagnostic> {
        Ok(Self {
            archive: ZipArchive::parse(bytes, limits)?,
            limits,
            cache,
        })
    }

    #[cfg(any(feature = "native-formats", feature = "xps-formats"))]
    pub(crate) fn cache(&self) -> PackageCache {
        self.cache.clone()
    }

    /// Opens an iWork ZIP with its path compatibility and expanded-byte budget.
    /// Container-level ZIP64 handling is shared with ordinary packages.
    #[cfg(feature = "iwork-formats")]
    pub(crate) fn open_iwork(bytes: &'a [u8], limits: Limits) -> Result<Self, Diagnostic> {
        Ok(Self {
            archive: ZipArchive::parse_iwork(bytes, limits)?,
            limits,
            cache: PackageCache::default(),
        })
    }

    pub const fn limits(&self) -> Limits {
        self.limits
    }

    pub fn has_part(&self, name: &str) -> bool {
        self.package_entry(name).is_some()
            || self.package_entry(&format!("{name}/[0].piece")).is_some()
            || self
                .package_entry(&format!("{name}/[0].last.piece"))
                .is_some()
    }

    pub fn entry_names(&self) -> impl Iterator<Item = &str> {
        self.archive.entries().iter().map(|entry| entry.name())
    }

    pub fn part(&self, name: &str) -> Result<Option<PartBytes>, Diagnostic> {
        if let Some(bytes) = self
            .cache
            .parts
            .lock()
            .ok()
            .and_then(|parts| parts.get(name).cloned())
        {
            return Ok(Some(PartBytes::Shared(bytes)));
        }
        let bytes = if let Some(entry) = self.package_entry(name) {
            self.archive
                .extract(entry)
                .map_err(|error| error.in_part(name))?
        } else {
            let prefix = format!("{name}/[");
            let mut pieces = self
                .archive
                .entries()
                .iter()
                .filter_map(|entry| {
                    let suffix = entry.name().get(prefix.len()..)?;
                    if !entry
                        .name()
                        .get(..prefix.len())?
                        .eq_ignore_ascii_case(&prefix)
                    {
                        return None;
                    }
                    let (index, ending) = suffix.split_once(']')?;
                    let last = match ending {
                        ".piece" => false,
                        ".last.piece" => true,
                        _ => return None,
                    };
                    Some((index.parse::<usize>().ok()?, last, entry))
                })
                .collect::<Vec<_>>();
            if pieces.is_empty() {
                return Ok(None);
            }
            pieces.sort_by_key(|piece| piece.0);
            let mut bytes = Vec::new();
            for (expected, (index, last, entry)) in pieces.iter().enumerate() {
                if *index != expected || *last != (expected + 1 == pieces.len()) {
                    return Err(xml_error(name, "invalid OPC piece sequence"));
                }
                let piece = self
                    .archive
                    .extract(entry)
                    .map_err(|error| error.in_part(name))?;
                if piece.len()
                    > self
                        .limits
                        .max_entry_uncompressed_bytes
                        .saturating_sub(bytes.len())
                {
                    return Err(Diagnostic::fatal(
                        DiagnosticCode::ZipEntryTooLarge,
                        Phase::Container,
                        None,
                        "OPC part exceeds the configured byte limit",
                    )
                    .in_part(name));
                }
                bytes.extend_from_slice(&piece);
            }
            bytes
        };
        let cache_limit = self
            .limits
            .max_total_uncompressed_bytes
            .min(64 * Limits::MIB);
        if cacheable_part(name)
            && let (Ok(mut parts), Ok(mut cached_bytes)) =
                (self.cache.parts.lock(), self.cache.bytes.lock())
            && bytes.len() <= cache_limit.saturating_sub(*cached_bytes)
        {
            let bytes: Arc<[u8]> = bytes.into();
            parts.insert(name.to_owned(), Arc::clone(&bytes));
            *cached_bytes += bytes.len();
            return Ok(Some(PartBytes::Shared(bytes)));
        }
        Ok(Some(PartBytes::Owned(bytes)))
    }

    fn package_entry(&self, name: &str) -> Option<&crate::zip::ZipEntry> {
        self.archive
            .entry(name)
            .filter(|entry| !entry.is_directory())
            .or_else(|| {
                let mut matches = self.archive.entries().iter().filter(|entry| {
                    !entry.is_directory() && entry.name().eq_ignore_ascii_case(name)
                });
                let entry = matches.next()?;
                matches.next().is_none().then_some(entry)
            })
    }

    pub fn required_part(&self, name: &str) -> Result<PartBytes, Diagnostic> {
        self.part(name)?.ok_or_else(|| {
            Diagnostic::fatal(
                DiagnosticCode::ZipInvalid,
                Phase::Container,
                None,
                format!("required package part is missing: {name}"),
            )
            .in_part(name)
        })
    }

    #[cfg(feature = "xps-formats")]
    pub(crate) fn required_part_prefix(
        &self,
        name: &str,
        max_bytes: usize,
    ) -> Result<Vec<u8>, Diagnostic> {
        if let Some(entry) = self.package_entry(name) {
            self.archive.extract_prefix(entry, max_bytes)
        } else {
            let bytes = self.required_part(name)?;
            Ok(bytes[..bytes.len().min(max_bytes)].to_vec())
        }
    }

    #[cfg(feature = "iwork-formats")]
    pub(crate) fn remaining_iwork_budget(&self) -> (usize, usize) {
        self.archive.remaining_nested_budget()
    }

    #[cfg(feature = "iwork-formats")]
    pub(crate) fn consume_iwork_budget(&self, expanded_bytes: usize, operations: usize) -> bool {
        self.archive
            .consume_nested_budget(expanded_bytes, operations)
    }

    /// Reads relationships for the package root (`None`) or a source part.
    pub fn relationships(
        &self,
        source_part: Option<&str>,
    ) -> Result<Vec<Relationship>, Diagnostic> {
        self.relationships_filtered(source_part, None)
    }

    // Typed consumers do not address edges by Id; unrelated malformed metadata
    // must not prevent them from following an otherwise valid typed edge.
    pub(crate) fn relationships_filtered(
        &self,
        source_part: Option<&str>,
        type_suffix: Option<&str>,
    ) -> Result<Vec<Relationship>, Diagnostic> {
        let rels_part = relationship_part_name(source_part);
        let Some(bytes) = self.part(&rels_part)? else {
            return Ok(Vec::new());
        };
        let mut relationships = Vec::new();
        let mut ids = HashSet::new();
        let mut edge_count = 0;
        parse_xml(&bytes, self.limits, |event| {
            let XmlEvent::StartElement {
                name, attributes, ..
            } = event
            else {
                return Ok(());
            };
            if local_name(name) != "Relationship" {
                return Ok(());
            }

            edge_count += 1;
            if edge_count > self.limits.max_relationship_edges {
                return Err(Diagnostic::fatal(
                    DiagnosticCode::RelationshipLimit,
                    Phase::Container,
                    None,
                    "OPC package exceeds the configured relationship limit",
                )
                .in_part(&rels_part));
            }
            if let Some(suffix) = type_suffix {
                let kind = optional_attribute(&attributes, "Type", &rels_part)?;
                if !kind.is_some_and(|kind| kind.to_ascii_lowercase().ends_with(suffix)) {
                    return Ok(());
                }
            }
            let id = if type_suffix.is_some() {
                optional_attribute(&attributes, "Id", &rels_part)?.unwrap_or_default()
            } else {
                required_attribute(&attributes, "Id", &rels_part)?
            };
            if type_suffix.is_none() && !ids.insert(id.clone()) {
                return Err(xml_error(
                    &rels_part,
                    format!("duplicate OPC relationship id: {id}"),
                ));
            }
            let type_uri = required_attribute(&attributes, "Type", &rels_part)?;
            let raw_target = required_attribute(&attributes, "Target", &rels_part)?;
            let target_mode = optional_attribute(&attributes, "TargetMode", &rels_part)?;
            let explicit_external = match target_mode.as_deref() {
                None | Some("Internal") => false,
                Some("External") => true,
                Some(_) => {
                    return Err(xml_error(
                        &rels_part,
                        "OPC relationship TargetMode must be Internal or External",
                    ));
                }
            };
            let external = explicit_external || looks_external(&raw_target);
            let target =
                if external || (type_uri.ends_with("/hyperlink") && raw_target.starts_with('#')) {
                    raw_target
                } else {
                    resolve_internal_target(source_part, &raw_target).map_err(|message| {
                        xml_error(
                            &rels_part,
                            format!("invalid OPC relationship target: {message}"),
                        )
                    })?
                };
            relationships.push(Relationship {
                id,
                type_uri,
                target,
                external,
            });
            Ok(())
        })
        .map_err(|error| {
            if error.location.part.is_some() {
                error
            } else {
                error.in_part(&rels_part)
            }
        })?;
        Ok(relationships)
    }
}

fn cacheable_part(name: &str) -> bool {
    name.ends_with(".xml")
        || name.ends_with(".rels")
        || name == "[Content_Types].xml"
        || name == "mimetype"
        || name.rsplit_once('.').is_some_and(|(_, extension)| {
            ["odttf", "ttf", "otf"]
                .iter()
                .any(|font| extension.eq_ignore_ascii_case(font))
        })
}

pub fn relationship_part_name(source_part: Option<&str>) -> String {
    let Some(source_part) = source_part else {
        return "_rels/.rels".to_owned();
    };
    match source_part.rsplit_once('/') {
        Some((directory, file)) => format!("{directory}/_rels/{file}.rels"),
        None => format!("_rels/{source_part}.rels"),
    }
}

pub(crate) fn resolve_internal_target(
    source_part: Option<&str>,
    target: &str,
) -> Result<String, &'static str> {
    if target.is_empty() || target.contains('\0') || target.contains('\\') {
        return Err("empty target or forbidden path character");
    }
    if target.contains('?') || target.contains('#') {
        return Err("query and fragment components are not package part names");
    }

    let mut segments = Vec::new();
    if !target.starts_with('/')
        && let Some((directory, _)) = source_part.and_then(|part| part.rsplit_once('/'))
    {
        segments.extend(directory.split('/').filter(|segment| !segment.is_empty()));
    }
    for segment in target.trim_start_matches('/').split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                if segments.pop().is_none() {
                    return Err("target escapes the package root");
                }
            }
            _ => segments.push(segment),
        }
    }
    if segments.is_empty() {
        return Err("target resolves to the package root");
    }
    Ok(segments.join("/"))
}

fn looks_external(target: &str) -> bool {
    if target.starts_with("//") {
        return true;
    }
    let Some(colon) = target.find(':') else {
        return false;
    };
    let scheme = &target[..colon];
    !scheme.is_empty()
        && scheme.bytes().enumerate().all(|(index, byte)| {
            matches!(
                (index, byte),
                (0, b'a'..=b'z' | b'A'..=b'Z')
                    | (_, b'a'..=b'z' | b'A'..=b'Z' | b'0'..=b'9' | b'+' | b'-' | b'.')
            )
        })
}

fn required_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<String, Diagnostic> {
    optional_attribute(attributes, name, part)?
        .ok_or_else(|| xml_error(part, format!("Relationship is missing {name}")))
}

fn optional_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<String>, Diagnostic> {
    attributes
        .iter()
        .find(|attribute| attribute.name == name)
        .map(|attribute| {
            decode_xml_text(attribute.value)
                .map(|value| value.into_owned())
                .map_err(|error| error.in_part(part))
        })
        .transpose()
}

fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

fn xml_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::XmlInvalid, Phase::Xml, None, message).in_part(part)
}

#[cfg(test)]
mod tests {
    use super::{looks_external, relationship_part_name, resolve_internal_target};

    #[cfg(feature = "xps-formats")]
    #[test]
    #[ignore = "requires the downloaded SampleXpsDocuments_1_0 corpus in .cache/xps-tests"]
    fn real_xps_font_cache_reuses_validated_bytes_with_a_bounded_budget() {
        let bytes = std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/.cache/xps-tests/SampleXpsDocuments_1_0/Handcrafted/XPS_Examples.xps"
        ))
        .unwrap();
        let limits = crate::limits::Limits::default();
        let package = super::Package::open(&bytes, limits).unwrap();
        let name = package
            .entry_names()
            .find(|name| name.ends_with(".ODTTF"))
            .unwrap();
        let super::PartBytes::Shared(first) = package.required_part(name).unwrap() else {
            panic!("XPS font bytes must enter the bounded package cache");
        };
        let reopened = super::Package::open_with_cache(&bytes, limits, package.cache()).unwrap();
        let super::PartBytes::Shared(second) = reopened.required_part(name).unwrap() else {
            panic!("lazy pages must reuse the shared font allocation");
        };
        assert!(std::sync::Arc::ptr_eq(&first, &second));
        assert_eq!(*package.cache.bytes.lock().unwrap(), first.len());
        let uncached = super::Package::open(&bytes, limits).unwrap();
        *uncached.cache.bytes.lock().unwrap() = 64 * crate::limits::Limits::MIB;
        let super::PartBytes::Owned(fallback) = uncached.required_part(name).unwrap() else {
            panic!("a full cache must not retain another font");
        };
        assert_eq!(fallback.as_slice(), first.as_ref());
        assert!(uncached.cache.parts.lock().unwrap().is_empty());
    }

    #[test]
    fn resolves_relationship_targets_without_allowing_root_escape() {
        assert_eq!(
            relationship_part_name(Some("ppt/slides/slide1.xml")),
            "ppt/slides/_rels/slide1.xml.rels"
        );
        assert_eq!(
            resolve_internal_target(Some("ppt/slides/slide1.xml"), "../media/image1.png"),
            Ok("ppt/media/image1.png".to_owned())
        );
        assert!(resolve_internal_target(None, "../outside.xml").is_err());
        assert!(looks_external("https://example.invalid/image.png"));
        assert!(looks_external("//example.invalid/image.png"));
        assert!(!looks_external("ppt/slides/slide1.xml"));
    }
}
