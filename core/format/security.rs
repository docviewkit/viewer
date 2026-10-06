//! Format-level security inspection shared by the enabled package parsers.
//!
//! The inspector never evaluates package content. It only reports parts that
//! could carry executable behavior. The ODF build also rejects encryption,
//! which this engine intentionally does not decrypt.

use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
use crate::limits::Limits;
use crate::package::Package;
#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
use crate::xml::{XmlEvent, parse_xml};

#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
use super::local_name;

#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
const ODF_MIMETYPE_PREFIX: &[u8] = b"application/vnd.oasis.opendocument.";
#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
const ODF_MANIFEST_PART: &str = "META-INF/manifest.xml";
#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
const ODF_ACTIVE_XML_PARTS: [&str; 2] = ["content.xml", "settings.xml"];

/// Rejects encrypted ODF packages and returns de-duplicated diagnostics for
/// executable or embedded content that the engine will not run.
///
/// Call this once after [`Package::open`] and append the returned diagnostics
/// to the parsed document. Any returned error is fatal and must stop parsing.
pub fn inspect(package: &Package<'_>) -> Result<Vec<Diagnostic>, Diagnostic> {
    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    let is_odf = package
        .part("mimetype")?
        .as_deref()
        .is_some_and(|value| value.starts_with(ODF_MIMETYPE_PREFIX));

    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    if is_odf && let Some(manifest) = package.part(ODF_MANIFEST_PART)? {
        reject_odf_encryption(&manifest, package.limits())?;
    }

    // There are only four finding classes, so a fixed array avoids a set and
    // bounds retained path data independently of package entry count.
    let mut findings: [Option<String>; 4] = [None, None, None, None];
    for name in package.entry_names() {
        record_entry_findings(name, &mut findings);
    }

    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    if is_odf {
        for part in ODF_ACTIVE_XML_PARTS {
            let Some(bytes) = package.part(part)? else {
                continue;
            };
            if contains_active_odf_element(&bytes, package.limits())
                .map_err(|error| with_part(error, part))?
            {
                record(&mut findings[3], part);
            }
        }
    }

    Ok(finding_diagnostics(findings))
}

#[cfg(feature = "odf-formats")]
pub(crate) fn inspect_flat_odf(
    bytes: &[u8],
    limits: Limits,
) -> Result<Vec<Diagnostic>, Diagnostic> {
    let mut findings: [Option<String>; 4] = [None, None, None, None];
    if contains_active_odf_element(bytes, limits)
        .map_err(|error| with_part(error, "content.xml"))?
    {
        record(&mut findings[3], "content.xml");
    }
    Ok(finding_diagnostics(findings))
}

#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
fn reject_odf_encryption(bytes: &[u8], limits: Limits) -> Result<(), Diagnostic> {
    parse_xml(bytes, limits, |event| {
        if matches!(
            event,
            XmlEvent::StartElement { name, .. } if local_name(name) == "encryption-data"
        ) {
            return Err(Diagnostic::fatal(
                DiagnosticCode::ZipEncrypted,
                Phase::Security,
                None,
                "ODF package contains encrypted content; encrypted packages are not supported",
            )
            .in_part(ODF_MANIFEST_PART));
        }
        Ok(())
    })
    .map_err(|error| with_part(error, ODF_MANIFEST_PART))?;
    Ok(())
}

#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
fn contains_active_odf_element(bytes: &[u8], limits: Limits) -> Result<bool, Diagnostic> {
    let mut found = false;
    parse_xml(bytes, limits, |event| {
        let XmlEvent::StartElement { name, .. } = event else {
            return Ok(());
        };
        if matches!(
            local_name(name),
            "script" | "scripts" | "event-listener" | "event-listeners"
        ) {
            found = true;
        }
        Ok(())
    })?;
    Ok(found)
}

fn record_entry_findings(name: &str, findings: &mut [Option<String>; 4]) {
    let segments = name.split('/');
    for segment in segments {
        if segment.eq_ignore_ascii_case("vbaProject.bin")
            || segment.eq_ignore_ascii_case("macrosheets")
            || segment.eq_ignore_ascii_case("dialogsheets")
        {
            record(&mut findings[0], name);
        }
        if segment.eq_ignore_ascii_case("activeX") {
            record(&mut findings[1], name);
        }
        if segment.eq_ignore_ascii_case("embeddings")
            || contains_ascii_case_insensitive(segment, "oleobject")
        {
            record(&mut findings[2], name);
        }
        #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
        if segment.eq_ignore_ascii_case("Scripts") || segment.eq_ignore_ascii_case("Basic") {
            record(&mut findings[3], name);
        }
    }
}

fn contains_ascii_case_insensitive(value: &str, needle: &str) -> bool {
    value
        .as_bytes()
        .windows(needle.len())
        .any(|window| window.eq_ignore_ascii_case(needle.as_bytes()))
}

fn record(slot: &mut Option<String>, part: &str) {
    if slot.is_none() {
        *slot = Some(part.to_owned());
    }
}

fn finding_diagnostics(findings: [Option<String>; 4]) -> Vec<Diagnostic> {
    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    const MESSAGES: [&str; 4] = [
        "macro content was blocked and will not execute",
        "ActiveX content was blocked and will not execute",
        "embedded or OLE content was blocked and will not execute",
        "ODF scripts or event bindings were blocked and will not execute",
    ];
    #[cfg(not(any(feature = "odf-formats", feature = "legacy-office-formats")))]
    const MESSAGES: [&str; 4] = [
        "macro content was blocked and will not execute",
        "ActiveX content was blocked and will not execute",
        "embedded or OLE content was blocked and will not execute",
        "",
    ];

    findings
        .into_iter()
        .zip(MESSAGES)
        .filter_map(|(part, message)| {
            part.map(|part| {
                Diagnostic::warning(
                    DiagnosticCode::ActiveContentBlocked,
                    Phase::Security,
                    Fidelity::Blocked,
                    message,
                )
                .in_part(part)
            })
        })
        .collect()
}

#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
fn with_part(error: Diagnostic, part: &str) -> Diagnostic {
    if error.location.part.is_some() {
        error
    } else {
        error.in_part(part)
    }
}

#[cfg(test)]
mod tests {
    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    use super::{contains_active_odf_element, reject_odf_encryption};
    use super::{finding_diagnostics, record_entry_findings};
    use crate::diagnostic::DiagnosticCode;
    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    use crate::diagnostic::Severity;
    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    use crate::limits::Limits;

    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    #[test]
    fn rejects_odf_manifest_encryption_with_package_encrypted_code() {
        let manifest = br#"<?xml version="1.0"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:full-path="content.xml">
    <manifest:encryption-data manifest:checksum-type="SHA256/1K"/>
  </manifest:file-entry>
</manifest:manifest>"#;

        let error = reject_odf_encryption(manifest, Limits::default()).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::ZipEncrypted);
        assert_eq!(error.severity, Severity::Fatal);
        assert_eq!(
            error.location.part.as_deref(),
            Some("META-INF/manifest.xml")
        );
    }

    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    #[test]
    fn accepts_an_unencrypted_odf_manifest() {
        let manifest = br#"<manifest:manifest xmlns:manifest="urn:test">
  <manifest:file-entry manifest:full-path="content.xml"/>
</manifest:manifest>"#;

        reject_odf_encryption(manifest, Limits::default()).unwrap();
    }

    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    #[test]
    fn recognizes_odf_scripts_and_event_bindings_streamingly() {
        let content = br#"<office:document-content xmlns:office="urn:office" xmlns:script="urn:script">
  <office:scripts><office:script script:language="ooo:Basic"/></office:scripts>
  <script:event-listeners><script:event-listener script:event-name="dom:click"/></script:event-listeners>
</office:document-content>"#;

        assert!(contains_active_odf_element(content, Limits::default()).unwrap());
        assert!(!contains_active_odf_element(
            br#"<office:document-content xmlns:office="urn:office"><office:body/></office:document-content>"#,
            Limits::default(),
        )
        .unwrap());
    }

    #[test]
    fn entry_findings_are_case_insensitive_and_deduplicated_by_class() {
        let mut findings = [None, None, None, None];
        for name in [
            "word/vbaProject.bin",
            "xl/VBAPROJECT.BIN",
            "ppt/activeX/activeX1.bin",
            "word/embeddings/oleObject1.bin",
            "Basic/Standard/Module1.xml",
            "Scripts/python/run.py",
        ] {
            record_entry_findings(name, &mut findings);
        }

        let diagnostics = finding_diagnostics(findings);
        assert_eq!(
            diagnostics.len(),
            if cfg!(any(
                feature = "odf-formats",
                feature = "legacy-office-formats"
            )) {
                4
            } else {
                3
            }
        );
        assert!(
            diagnostics
                .iter()
                .all(|diagnostic| diagnostic.code == DiagnosticCode::ActiveContentBlocked)
        );
        assert_eq!(
            diagnostics[0].location.part.as_deref(),
            Some("word/vbaProject.bin")
        );
    }

    #[test]
    fn ordinary_media_and_object_replacements_are_not_active_content() {
        let mut findings = [None, None, None, None];
        for name in [
            "ppt/media/image1.png",
            "ObjectReplacements/Object 1",
            "word/media/ole-logo.png",
        ] {
            record_entry_findings(name, &mut findings);
        }

        assert!(findings.iter().all(Option::is_none));
    }
}
