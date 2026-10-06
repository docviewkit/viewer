#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
use crate::diagnostic::{Diagnostic, DiagnosticCode, Phase};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(in crate::format) enum OfficeImageError {
    UnsupportedFormat,
    SignatureMismatch,
    DisabledByOffice,
}

pub(in crate::format) fn office_image_media_type(
    part: &str,
    bytes: &[u8],
) -> Result<&'static str, OfficeImageError> {
    let extension = part
        .rsplit('/')
        .next()
        .and_then(|name| name.rsplit_once('.').map(|(_, extension)| extension))
        .map(str::to_ascii_lowercase);
    let Some(extension) = extension.as_deref() else {
        return Err(OfficeImageError::UnsupportedFormat);
    };
    let format = match extension {
        "png" => ImageFormat::Png,
        "jpg" | "jpeg" | "jpe" | "jfif" => ImageFormat::Jpeg,
        "gif" => ImageFormat::Gif,
        "webp" => ImageFormat::Webp,
        "bmp" => ImageFormat::BitmapFile,
        "dib" | "rle" => ImageFormat::Dib,
        "svg" => ImageFormat::Svg,
        "tif" | "tiff" => ImageFormat::Tiff,
        "ico" => ImageFormat::Icon,
        "pcx" => ImageFormat::Pcx,
        "jp2" | "jpx" => ImageFormat::Jp2,
        "j2k" | "jpc" => ImageFormat::J2k,
        "emf" => ImageFormat::Emf,
        "wmf" => ImageFormat::Wmf,
        "emz" => ImageFormat::Emz,
        "wmz" => ImageFormat::Wmz,
        "eps" | "pict" | "pct" | "pic" | "pcz" => {
            return Err(OfficeImageError::DisabledByOffice);
        }
        _ => return Err(OfficeImageError::UnsupportedFormat),
    };
    validate_image_format(format, bytes)
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats",
    test
))]
pub(in crate::format) fn office_image_media_type_from_mime(
    media_type: &str,
    bytes: &[u8],
) -> Result<&'static str, OfficeImageError> {
    let media_type = media_type
        .split_once(';')
        .map_or(media_type, |(essence, _)| essence)
        .trim()
        .to_ascii_lowercase();
    let format = match media_type.as_str() {
        "image/png" | "image/x-png" => ImageFormat::Png,
        "image/jpeg" | "image/jpg" | "image/pjpeg" => ImageFormat::Jpeg,
        "image/gif" => ImageFormat::Gif,
        "image/webp" => ImageFormat::Webp,
        "image/bmp" | "image/x-bmp" | "image/x-ms-bmp" | "image/x-dib" | "image/x-rle" => {
            ImageFormat::Dib
        }
        "image/svg+xml" | "image/svg" => ImageFormat::Svg,
        "image/tiff" | "image/tif" => ImageFormat::Tiff,
        "image/x-icon" | "image/vnd.microsoft.icon" | "image/ico" => ImageFormat::Icon,
        "image/x-pcx" | "image/pcx" | "image/vnd.zbrush.pcx" => ImageFormat::Pcx,
        "image/jp2" | "image/jpx" => ImageFormat::Jp2,
        "image/j2k" | "image/jpc" => ImageFormat::J2k,
        "image/emf" | "image/x-emf" | "application/x-emf" => ImageFormat::Emf,
        "image/wmf" | "image/x-wmf" | "application/x-wmf" | "application/x-msmetafile" => {
            ImageFormat::Wmf
        }
        "image/emz" | "image/x-emz" => ImageFormat::Emz,
        "image/wmz" | "image/x-wmz" => ImageFormat::Wmz,
        "application/postscript"
        | "application/x-pict"
        | "image/eps"
        | "image/x-eps"
        | "image/pict"
        | "image/x-pict" => {
            return Err(OfficeImageError::DisabledByOffice);
        }
        _ => return Err(OfficeImageError::UnsupportedFormat),
    };
    validate_image_format(format, bytes)
}

/// Recover mislabeled supported images; never bypass disabled-format policy.
#[cfg(any(feature = "native-formats", feature = "odf-formats"))]
pub(in crate::format) fn recover_office_image_signature(
    declared: Result<&'static str, OfficeImageError>,
    bytes: &[u8],
) -> Result<&'static str, OfficeImageError> {
    declared.or_else(|error| match error {
        OfficeImageError::SignatureMismatch => office_image_media_type_from_signature(bytes)
            .map_err(|_| OfficeImageError::SignatureMismatch),
        _ => Err(error),
    })
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "ofd-formats",
    feature = "legacy-office-formats"
))]
pub(in crate::format) fn office_image_media_type_from_signature(
    bytes: &[u8],
) -> Result<&'static str, OfficeImageError> {
    for format in [
        ImageFormat::Png,
        ImageFormat::Jpeg,
        ImageFormat::Gif,
        ImageFormat::Webp,
        ImageFormat::BitmapFile,
        ImageFormat::Svg,
        ImageFormat::Tiff,
        ImageFormat::Icon,
        ImageFormat::Pcx,
        ImageFormat::Jp2,
        ImageFormat::J2k,
        ImageFormat::Emf,
        ImageFormat::Wmf,
    ] {
        if let Ok(media_type) = validate_image_format(format, bytes) {
            return Ok(media_type);
        }
    }
    Err(OfficeImageError::UnsupportedFormat)
}

#[derive(Clone, Copy)]
enum ImageFormat {
    Png,
    Jpeg,
    Gif,
    Webp,
    BitmapFile,
    Dib,
    Svg,
    Tiff,
    Icon,
    Pcx,
    Jp2,
    J2k,
    Emf,
    Wmf,
    Emz,
    Wmz,
}

fn validate_image_format(
    format: ImageFormat,
    bytes: &[u8],
) -> Result<&'static str, OfficeImageError> {
    let (media_type, signature_matches) = match format {
        ImageFormat::Png => ("image/png", bytes.starts_with(b"\x89PNG\r\n\x1a\n")),
        ImageFormat::Jpeg => ("image/jpeg", bytes.starts_with(b"\xff\xd8\xff")),
        ImageFormat::Gif => (
            "image/gif",
            bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a"),
        ),
        ImageFormat::Webp => (
            "image/webp",
            bytes.len() >= 12 && &bytes[..4] == b"RIFF" && &bytes[8..12] == b"WEBP",
        ),
        ImageFormat::BitmapFile => ("image/bmp", is_bitmap_file(bytes)),
        ImageFormat::Dib => ("image/bmp", is_bitmap_file(bytes) || is_dib(bytes)),
        ImageFormat::Svg => ("image/svg+xml", has_svg_root(bytes)),
        ImageFormat::Tiff => ("image/tiff", is_tiff(bytes)),
        ImageFormat::Icon => ("image/x-icon", is_icon(bytes)),
        ImageFormat::Pcx => ("image/x-pcx", is_pcx(bytes)),
        ImageFormat::Jp2 => ("image/jp2", is_jp2(bytes)),
        ImageFormat::J2k => ("image/j2k", bytes.starts_with(b"\xff\x4f\xff\x51")),
        ImageFormat::Emf => ("image/x-emf", is_emf(bytes)),
        ImageFormat::Wmf => ("image/x-wmf", is_wmf(bytes)),
        ImageFormat::Emz => ("image/x-emz", is_gzip(bytes)),
        ImageFormat::Wmz => ("image/x-wmz", is_gzip(bytes)),
    };
    signature_matches
        .then_some(media_type)
        .ok_or(OfficeImageError::SignatureMismatch)
}

fn is_bitmap_file(bytes: &[u8]) -> bool {
    bytes.starts_with(b"BM")
}

fn is_dib(bytes: &[u8]) -> bool {
    let Some(header) = bytes.get(..4) else {
        return false;
    };
    matches!(
        u32::from_le_bytes(header.try_into().expect("four-byte DIB header")),
        12 | 16 | 40 | 52 | 56 | 64 | 108 | 124
    )
}

fn has_svg_root(bytes: &[u8]) -> bool {
    let bytes = bytes.strip_prefix(b"\xef\xbb\xbf").unwrap_or(bytes);
    let Ok(mut source) = std::str::from_utf8(bytes) else {
        return false;
    };
    loop {
        source = source.trim_start_matches(char::is_whitespace);
        if source.starts_with("<?") {
            let Some(end) = source.find("?>") else {
                return false;
            };
            source = &source[end + 2..];
        } else if source.starts_with("<!--") {
            let Some(end) = source.find("-->") else {
                return false;
            };
            source = &source[end + 3..];
        } else if source.starts_with("<!DOCTYPE") {
            let Some(end) = xml_declaration_end(source) else {
                return false;
            };
            source = &source[end + 1..];
        } else {
            break;
        }
    }

    let Some(tag) = source.strip_prefix('<') else {
        return false;
    };
    if tag.starts_with(['/', '!', '?']) {
        return false;
    }
    let name_end = tag
        .find(|character: char| character.is_whitespace() || matches!(character, '/' | '>'))
        .unwrap_or(tag.len());
    let qualified_name = &tag[..name_end];
    !qualified_name.is_empty()
        && qualified_name
            .rsplit_once(':')
            .map_or(qualified_name == "svg", |(_, local)| local == "svg")
}

fn xml_declaration_end(source: &str) -> Option<usize> {
    let mut quote = None;
    let mut subset_depth = 0_u32;
    for (index, character) in source.char_indices() {
        if let Some(delimiter) = quote {
            if character == delimiter {
                quote = None;
            }
            continue;
        }
        match character {
            '\'' | '"' => quote = Some(character),
            '[' => subset_depth = subset_depth.saturating_add(1),
            ']' => subset_depth = subset_depth.saturating_sub(1),
            '>' if subset_depth == 0 => return Some(index),
            _ => {}
        }
    }
    None
}

fn is_tiff(bytes: &[u8]) -> bool {
    bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*")
}

fn is_icon(bytes: &[u8]) -> bool {
    bytes.len() >= 6 && bytes[..4] == [0, 0, 1, 0] && u16::from_le_bytes([bytes[4], bytes[5]]) != 0
}

fn is_pcx(bytes: &[u8]) -> bool {
    bytes.len() >= 128
        && bytes[0] == 0x0a
        && matches!(bytes[1], 0 | 2 | 3 | 4 | 5)
        && matches!(bytes[2], 0 | 1)
        && matches!(bytes[3], 1 | 2 | 4 | 8)
}

fn is_jp2(bytes: &[u8]) -> bool {
    bytes.starts_with(b"\0\0\0\x0cjP  \r\n\x87\n")
}

fn is_emf(bytes: &[u8]) -> bool {
    if bytes.len() < 88 || bytes[..4] != 1_u32.to_le_bytes() || bytes[40..44] != *b" EMF" {
        return false;
    }
    let header_size = u32::from_le_bytes(bytes[4..8].try_into().expect("EMF header size"));
    header_size >= 88
        && header_size.is_multiple_of(4)
        && usize::try_from(header_size).is_ok_and(|size| size <= bytes.len())
}

fn is_wmf(bytes: &[u8]) -> bool {
    if bytes.starts_with(b"\xd7\xcd\xc6\x9a") {
        return bytes.len() >= 40
            && placeable_wmf_checksum_matches(bytes)
            && has_wmf_header(&bytes[22..]);
    }
    has_wmf_header(bytes)
}

fn placeable_wmf_checksum_matches(bytes: &[u8]) -> bool {
    let Some(header) = bytes.get(..22) else {
        return false;
    };
    let expected = u16::from_le_bytes([header[20], header[21]]);
    let actual = header[..20]
        .chunks_exact(2)
        .map(|word| u16::from_le_bytes([word[0], word[1]]))
        .fold(0_u16, |checksum, word| checksum ^ word);
    actual == expected
}

fn has_wmf_header(bytes: &[u8]) -> bool {
    bytes.len() >= 18
        && matches!(u16::from_le_bytes([bytes[0], bytes[1]]), 1 | 2)
        && u16::from_le_bytes([bytes[2], bytes[3]]) == 9
        && matches!(u16::from_le_bytes([bytes[4], bytes[5]]), 0x0100 | 0x0300)
}

fn is_gzip(bytes: &[u8]) -> bool {
    bytes.len() >= 10
        && bytes[..3] == [0x1f, 0x8b, 8]
        && bytes[3] & 0xe0 == 0
        && gzip_optional_header_is_complete(bytes)
}

fn gzip_optional_header_is_complete(bytes: &[u8]) -> bool {
    let flags = bytes[3];
    let mut offset = 10_usize;
    if flags & 0x04 != 0 {
        let Some(length) = bytes.get(offset..offset + 2) else {
            return false;
        };
        offset += 2 + usize::from(u16::from_le_bytes([length[0], length[1]]));
        if offset > bytes.len() {
            return false;
        }
    }
    for flag in [0x08, 0x10] {
        if flags & flag != 0 {
            let Some(end) = bytes
                .get(offset..)
                .and_then(|tail| tail.iter().position(|byte| *byte == 0))
            else {
                return false;
            };
            offset += end + 1;
        }
    }
    flags & 0x02 == 0 || offset.checked_add(2).is_some_and(|end| end <= bytes.len())
}

#[derive(Clone, Debug, Eq, PartialEq)]
#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
pub(in crate::format) enum OdfImageTarget {
    Embedded(String),
    External,
}

#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
pub(in crate::format) fn resolve_odf_image_target(
    href: &str,
    max_path_bytes: usize,
) -> Result<OdfImageTarget, &'static str> {
    if href.is_empty() || href.len() > max_path_bytes {
        return Err("image reference is empty or exceeds the configured path limit");
    }
    if looks_external(href) || href.starts_with("../") {
        return Ok(OdfImageTarget::External);
    }
    if href.contains('\0')
        || href.contains('\\')
        || href.contains('?')
        || href.contains('#')
        || href.starts_with('/')
    {
        return Err("image reference contains a forbidden path component");
    }

    let mut segments = Vec::new();
    for segment in href.split('/') {
        match segment {
            "." => {}
            "" | ".." => return Err("image reference is not a canonical package path"),
            _ => segments.push(segment),
        }
    }
    if segments.is_empty() {
        return Err("image reference resolves to the package root");
    }
    Ok(OdfImageTarget::Embedded(segments.join("/")))
}

#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
fn looks_external(target: &str) -> bool {
    if target.starts_with("//") {
        return true;
    }
    let Some(colon) = target.find(':') else {
        return false;
    };
    target[..colon].bytes().enumerate().all(|(index, byte)| {
        matches!(
            (index, byte),
            (0, b'a'..=b'z' | b'A'..=b'Z')
                | (_, b'a'..=b'z' | b'A'..=b'Z' | b'0'..=b'9' | b'+' | b'-' | b'.')
        )
    })
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(in crate::format) enum ImageCacheEntry {
    Object(usize),
    Unsupported(OfficeImageError),
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(in crate::format) fn clone_image_bytes(
    bytes: &[u8],
    part: &str,
) -> Result<Vec<u8>, Diagnostic> {
    let mut clone = Vec::new();
    clone.try_reserve_exact(bytes.len()).map_err(|_| {
        Diagnostic::fatal(
            DiagnosticCode::AllocationFailed,
            Phase::Parse,
            None,
            "unable to allocate storage for a repeated embedded image",
        )
        .in_part(part)
    })?;
    clone.extend_from_slice(bytes);
    Ok(clone)
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(in crate::format) fn reserve_materialized_image_bytes(
    total: &mut usize,
    additional: usize,
    limit: usize,
    part: &str,
) -> Result<(), Diagnostic> {
    let next = total.checked_add(additional).ok_or_else(|| {
        Diagnostic::fatal(
            DiagnosticCode::ZipTotalSizeLimit,
            Phase::Parse,
            None,
            "materialized embedded image bytes overflow the document budget",
        )
        .in_part(part)
    })?;
    if next > limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ZipTotalSizeLimit,
            Phase::Parse,
            None,
            "materialized embedded images exceed the configured uncompressed-byte budget",
        )
        .in_part(part));
    }
    *total = next;
    Ok(())
}

#[cfg(test)]
#[cfg(any(feature = "native-formats", feature = "odf-formats"))]
pub(in crate::format) fn stored_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut bytes = Vec::new();
    let mut central = Vec::new();

    for (name, content) in entries {
        let name = name.as_bytes();
        let offset = u32::try_from(bytes.len()).expect("test ZIP offset fits u32");
        let crc = crate::zip::crc32(content);

        push_u32(&mut bytes, 0x0403_4b50);
        push_u16(&mut bytes, 20);
        push_u16(&mut bytes, 1 << 11);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u16(&mut bytes, 0);
        push_u32(&mut bytes, crc);
        push_u32(&mut bytes, content.len() as u32);
        push_u32(&mut bytes, content.len() as u32);
        push_u16(&mut bytes, name.len() as u16);
        push_u16(&mut bytes, 0);
        bytes.extend_from_slice(name);
        bytes.extend_from_slice(content);

        push_u32(&mut central, 0x0201_4b50);
        push_u16(&mut central, 20);
        push_u16(&mut central, 20);
        push_u16(&mut central, 1 << 11);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, crc);
        push_u32(&mut central, content.len() as u32);
        push_u32(&mut central, content.len() as u32);
        push_u16(&mut central, name.len() as u16);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u16(&mut central, 0);
        push_u32(&mut central, 0);
        push_u32(&mut central, offset);
        central.extend_from_slice(name);
    }

    let central_offset = bytes.len() as u32;
    let central_size = central.len() as u32;
    bytes.extend_from_slice(&central);
    push_u32(&mut bytes, 0x0605_4b50);
    push_u16(&mut bytes, 0);
    push_u16(&mut bytes, 0);
    push_u16(&mut bytes, entries.len() as u16);
    push_u16(&mut bytes, entries.len() as u16);
    push_u32(&mut bytes, central_size);
    push_u32(&mut bytes, central_offset);
    push_u16(&mut bytes, 0);
    bytes
}

#[cfg(test)]
#[cfg(any(feature = "native-formats", feature = "odf-formats"))]
fn push_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

#[cfg(test)]
#[cfg(any(feature = "native-formats", feature = "odf-formats"))]
fn push_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

#[cfg(test)]
mod tests {
    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    use super::{OdfImageTarget, resolve_odf_image_target};
    use super::{OfficeImageError, office_image_media_type, office_image_media_type_from_mime};

    #[cfg(any(feature = "native-formats", feature = "odf-formats"))]
    #[test]
    fn recovers_supplied_mislabeled_metafiles_without_weakening_format_checks() {
        let package = crate::package::Package::open(
            include_bytes!("../../tests/fixtures/word-columns-tab-leaders.odt"),
            crate::limits::Limits::default(),
        )
        .unwrap();
        for part in ["media/image6.emf", "media/image7.emf"] {
            let bytes = package.part(part).unwrap().unwrap();
            let declared = office_image_media_type(part, &bytes);
            assert_eq!(declared, Err(OfficeImageError::SignatureMismatch));
            assert_eq!(
                super::recover_office_image_signature(declared, &bytes),
                Ok("image/x-wmf")
            );
            let mut corrupt = bytes.to_vec();
            corrupt[20] ^= 1; // Placeable WMF checksum must still be validated.
            assert_eq!(
                super::recover_office_image_signature(declared, &corrupt),
                declared
            );
        }
        for error in [
            OfficeImageError::DisabledByOffice,
            OfficeImageError::UnsupportedFormat,
        ] {
            assert_eq!(
                super::recover_office_image_signature(Err(error), b"\x89PNG\r\n\x1a\n"),
                Err(error)
            );
        }
    }

    #[test]
    fn accepts_office_image_extensions_with_matching_signatures() {
        let mut dib = vec![0_u8; 40];
        dib[..4].copy_from_slice(&40_u32.to_le_bytes());
        let mut pcx = vec![0_u8; 128];
        pcx[..4].copy_from_slice(&[0x0a, 5, 1, 8]);
        let mut emf = vec![0_u8; 88];
        emf[..4].copy_from_slice(&1_u32.to_le_bytes());
        emf[4..8].copy_from_slice(&88_u32.to_le_bytes());
        emf[40..44].copy_from_slice(b" EMF");
        let mut wmf = vec![0_u8; 18];
        wmf[..2].copy_from_slice(&1_u16.to_le_bytes());
        wmf[2..4].copy_from_slice(&9_u16.to_le_bytes());
        wmf[4..6].copy_from_slice(&0x0300_u16.to_le_bytes());

        let cases: Vec<(&str, &[u8], &str)> = vec![
            ("Pictures/a.PNG", b"\x89PNG\r\n\x1a\nrest", "image/png"),
            ("Pictures/a.jpeg", b"\xff\xd8\xffrest", "image/jpeg"),
            ("Pictures/a.jfif", b"\xff\xd8\xffrest", "image/jpeg"),
            ("Pictures/a.gif", b"GIF89arest", "image/gif"),
            ("Pictures/a.webp", b"RIFFxxxxWEBPrest", "image/webp"),
            ("Pictures/a.bmp", b"BMrest", "image/bmp"),
            ("Pictures/a.dib", &dib, "image/bmp"),
            ("Pictures/a.rle", &dib, "image/bmp"),
            (
                "Pictures/a.svg",
                b"\xef\xbb\xbf <?xml version='1.0'?><!--x--><svg xmlns='http://www.w3.org/2000/svg'/>",
                "image/svg+xml",
            ),
            ("Pictures/a.tif", b"II*\0rest", "image/tiff"),
            ("Pictures/a.tiff", b"MM\0*rest", "image/tiff"),
            ("Pictures/a.ico", b"\0\0\x01\0\x01\0rest", "image/x-icon"),
            ("Pictures/a.pcx", &pcx, "image/x-pcx"),
            (
                "Pictures/a.jp2",
                b"\0\0\0\x0cjP  \r\n\x87\nrest",
                "image/jp2",
            ),
            (
                "Pictures/a.jpx",
                b"\0\0\0\x0cjP  \r\n\x87\nrest",
                "image/jp2",
            ),
            ("Pictures/a.j2k", b"\xff\x4f\xff\x51rest", "image/j2k"),
            ("Pictures/a.jpc", b"\xff\x4f\xff\x51rest", "image/j2k"),
            ("Pictures/a.emf", &emf, "image/x-emf"),
            ("Pictures/a.wmf", &wmf, "image/x-wmf"),
            (
                "Pictures/a.emz",
                b"\x1f\x8b\x08\0\0\0\0\0\0\xff",
                "image/x-emz",
            ),
            (
                "Pictures/a.wmz",
                b"\x1f\x8b\x08\0\0\0\0\0\0\xff",
                "image/x-wmz",
            ),
        ];
        for (part, bytes, media_type) in cases {
            assert_eq!(
                office_image_media_type(part, bytes),
                Ok(media_type),
                "{part}"
            );
        }
    }

    #[test]
    fn rejects_spoofed_or_truncated_image_signatures() {
        assert_eq!(
            office_image_media_type("Pictures/a.png", b"BMrest"),
            Err(OfficeImageError::SignatureMismatch)
        );
        assert_eq!(
            office_image_media_type("Pictures/a.svg", b"<svgx/>"),
            Err(OfficeImageError::SignatureMismatch)
        );
        assert_eq!(
            office_image_media_type("Pictures/a.pcx", b"\x0a\x05\x01\x08"),
            Err(OfficeImageError::SignatureMismatch)
        );
        assert_eq!(
            office_image_media_type("Pictures/a.emz", b"\x1f\x8b\x08"),
            Err(OfficeImageError::SignatureMismatch)
        );
        assert_eq!(
            office_image_media_type("Pictures/a.avif", b"image"),
            Err(OfficeImageError::UnsupportedFormat)
        );
    }

    #[test]
    fn validates_declared_image_media_types_independently_of_file_names() {
        assert_eq!(
            office_image_media_type_from_mime(
                "IMAGE/PNG; charset=binary",
                b"\x89PNG\r\n\x1a\nrest"
            ),
            Ok("image/png")
        );
        assert_eq!(
            office_image_media_type_from_mime("image/x-png", b"\x89PNG\r\n\x1a\nrest"),
            Ok("image/png")
        );
        assert_eq!(
            office_image_media_type_from_mime("image/svg", b"<svg/>"),
            Ok("image/svg+xml")
        );
        assert_eq!(
            office_image_media_type_from_mime("image/png", b"GIF89arest"),
            Err(OfficeImageError::SignatureMismatch)
        );
        assert_eq!(
            office_image_media_type_from_mime("image/avif", b"image"),
            Err(OfficeImageError::UnsupportedFormat)
        );
        assert_eq!(
            office_image_media_type_from_mime("image/x-pict", b"image"),
            Err(OfficeImageError::DisabledByOffice)
        );
    }

    #[test]
    fn classifies_formats_disabled_by_current_office() {
        for extension in ["eps", "pict", "pct", "pic", "pcz"] {
            assert_eq!(
                office_image_media_type(&format!("Pictures/a.{extension}"), b"image"),
                Err(OfficeImageError::DisabledByOffice),
                "{extension}",
            );
        }
    }

    #[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
    #[test]
    fn resolves_only_bounded_picture_parts_for_odf() {
        assert_eq!(
            resolve_odf_image_target("./Pictures/a.png", 128),
            Ok(OdfImageTarget::Embedded("Pictures/a.png".to_owned()))
        );
        assert_eq!(
            resolve_odf_image_target("https://example.invalid/a.png", 128),
            Ok(OdfImageTarget::External)
        );
        assert_eq!(
            resolve_odf_image_target("../Pictures/a.png", 128),
            Ok(OdfImageTarget::External)
        );
        assert_eq!(
            resolve_odf_image_target("Thumbnails/a.png", 128),
            Ok(OdfImageTarget::Embedded("Thumbnails/a.png".to_owned()))
        );
        assert_eq!(
            resolve_odf_image_target("image.png", 128),
            Ok(OdfImageTarget::Embedded("image.png".to_owned()))
        );
        assert!(resolve_odf_image_target("Pictures/a.png#x", 128).is_err());
    }
}
