use crate::model::MediaKind;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum EmbeddedMediaError {
    UnsupportedFormat,
    SignatureMismatch,
}

pub(super) fn embedded_media_type(
    target: &str,
    bytes: &[u8],
    kind_hint: Option<MediaKind>,
) -> Result<(MediaKind, &'static str), EmbeddedMediaError> {
    let extension = target
        .rsplit_once('.')
        .map(|(_, extension)| extension.to_ascii_lowercase())
        .unwrap_or_default();
    let (kind, media_type) = match extension.as_str() {
        "mp4" | "m4v" => (kind_hint.unwrap_or(MediaKind::Video), "video/mp4"),
        "m4a" => (MediaKind::Audio, "audio/mp4"),
        "mov" => (MediaKind::Video, "video/quicktime"),
        "webm" => (kind_hint.unwrap_or(MediaKind::Video), "video/webm"),
        "weba" => (MediaKind::Audio, "audio/webm"),
        "ogv" => (MediaKind::Video, "video/ogg"),
        "ogg" | "oga" => (kind_hint.unwrap_or(MediaKind::Audio), "audio/ogg"),
        "mp3" => (MediaKind::Audio, "audio/mpeg"),
        "wav" | "wave" => (MediaKind::Audio, "audio/wav"),
        "aac" => (MediaKind::Audio, "audio/aac"),
        _ => return Err(EmbeddedMediaError::UnsupportedFormat),
    };
    validate_signature(media_type, bytes)?;
    Ok((kind, media_type))
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(super) fn embedded_media_type_from_mime(
    media_type: &str,
    bytes: &[u8],
) -> Result<(MediaKind, &'static str), EmbeddedMediaError> {
    let normalized = media_type
        .split(';')
        .next()
        .unwrap_or_default()
        .trim()
        .to_ascii_lowercase();
    let (kind, canonical) = match normalized.as_str() {
        "video/mp4" | "application/mp4" => (MediaKind::Video, "video/mp4"),
        "audio/mp4" | "audio/x-m4a" => (MediaKind::Audio, "audio/mp4"),
        "video/quicktime" => (MediaKind::Video, "video/quicktime"),
        "video/webm" => (MediaKind::Video, "video/webm"),
        "audio/webm" => (MediaKind::Audio, "audio/webm"),
        "video/ogg" => (MediaKind::Video, "video/ogg"),
        "audio/ogg" | "application/ogg" => (MediaKind::Audio, "audio/ogg"),
        "audio/mpeg" | "audio/mp3" => (MediaKind::Audio, "audio/mpeg"),
        "audio/wav" | "audio/wave" | "audio/x-wav" => (MediaKind::Audio, "audio/wav"),
        "audio/aac" => (MediaKind::Audio, "audio/aac"),
        _ => return Err(EmbeddedMediaError::UnsupportedFormat),
    };
    validate_signature(canonical, bytes)?;
    Ok((kind, canonical))
}

fn validate_signature(media_type: &str, bytes: &[u8]) -> Result<(), EmbeddedMediaError> {
    let valid = match media_type {
        "video/mp4" | "audio/mp4" | "video/quicktime" => {
            bytes.len() >= 12 && bytes.get(4..8) == Some(b"ftyp")
        }
        "video/webm" | "audio/webm" => {
            bytes.starts_with(&[0x1a, 0x45, 0xdf, 0xa3])
                && bytes
                    .windows(4)
                    .take(64)
                    .any(|window| window.eq_ignore_ascii_case(b"webm"))
        }
        "video/ogg" | "audio/ogg" => bytes.starts_with(b"OggS"),
        "audio/mpeg" => {
            bytes.starts_with(b"ID3")
                || bytes.len() >= 2 && bytes[0] == 0xff && bytes[1] & 0xe0 == 0xe0
        }
        "audio/wav" => {
            bytes.len() >= 12 && bytes.starts_with(b"RIFF") && bytes.get(8..12) == Some(b"WAVE")
        }
        "audio/aac" => bytes.len() >= 2 && bytes[0] == 0xff && bytes[1] & 0xf6 == 0xf0,
        _ => false,
    };
    if valid {
        Ok(())
    } else {
        Err(EmbeddedMediaError::SignatureMismatch)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn identifies_supported_media_without_trusting_extensions() {
        let mp4 = b"\0\0\0\x18ftypisom\0\0\0\0";
        assert_eq!(
            embedded_media_type("media/video.mp4", mp4, None),
            Ok((MediaKind::Video, "video/mp4"))
        );
        assert_eq!(
            embedded_media_type("media/video.mp4", b"not mp4", None),
            Err(EmbeddedMediaError::SignatureMismatch)
        );
        #[cfg(any(feature = "native-formats", feature = "odf-formats"))]
        assert_eq!(
            embedded_media_type_from_mime("audio/wav", b"RIFF\0\0\0\0WAVE"),
            Ok((MediaKind::Audio, "audio/wav"))
        );
    }
}
