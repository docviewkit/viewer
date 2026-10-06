//! Bounded browser-measured font advances supplied by the JavaScript layout realm.

use std::collections::HashMap;

use crate::diagnostic::{Diagnostic, DiagnosticCode, Phase};
use crate::limits::Limits;

const MAGIC: &[u8; 4] = b"OVFM";
const VERSION: u16 = 3;
const LEGACY_VERSION: u16 = 1;
const HEADER_BYTES: usize = 20;
const FACE_RECORD_BYTES_V1: usize = 16;
const FACE_RECORD_BYTES_V2: usize = 28;
const FACE_RECORD_BYTES_V3: usize = 32;
const METRIC_RECORD_BYTES: usize = 8;
const NORMAL_STRETCH: u8 = 4;
const MAX_FAMILY_BYTES: usize = 1_024;
const MAX_ADVANCE_EM: f32 = 1_024.0;
const FALLBACK_FACE_FLAG: u16 = 1;
const UNKNOWN_LINE_GAP_FLAG: u16 = 2;

#[derive(Clone, Debug, Eq, Hash, PartialEq)]
struct FontFaceKey {
    family: String,
    style: u8,
    weight: u16,
    stretch: u8,
    size_bits: u32,
}

/// A document-scoped, immutable table. Missing records deliberately fall back
/// to the format adapter's bounded approximation.
#[derive(Clone, Debug, Default)]
pub struct FontMetricTable {
    faces: HashMap<FontFaceKey, FontFaceMetrics>,
    metric_count: usize,
}

#[derive(Clone, Debug, Default)]
struct FontFaceMetrics {
    advances: HashMap<u32, f32>,
    vertical_metrics: Option<[f32; 3]>,
    fallback: bool,
    line_gap_unknown: bool,
}

impl FontMetricTable {
    pub fn decode(bytes: &[u8], limits: Limits) -> Result<Self, Diagnostic> {
        if bytes.is_empty() {
            return Ok(Self::default());
        }
        if bytes.len() > limits.max_font_bytes {
            return Err(invalid(
                "font metric table exceeds the configured font byte limit",
            ));
        }
        if bytes.len() < HEADER_BYTES || bytes.get(0..4) != Some(MAGIC.as_slice()) {
            return Err(invalid("font metric table header is invalid"));
        }
        let version = read_u16(bytes, 4)?;
        let header_bytes = read_u16(bytes, 6)? as usize;
        if !matches!(version, LEGACY_VERSION | 2 | VERSION) || header_bytes != HEADER_BYTES {
            return Err(invalid("font metric table version is unsupported"));
        }
        let face_record_bytes = if version == LEGACY_VERSION {
            FACE_RECORD_BYTES_V1
        } else if version == 2 {
            FACE_RECORD_BYTES_V2
        } else {
            FACE_RECORD_BYTES_V3
        };
        let face_count = read_u32(bytes, 8)? as usize;
        let metric_count = read_u32(bytes, 12)? as usize;
        let metrics_offset = read_u32(bytes, 16)? as usize;
        if face_count > limits.max_document_objects || metric_count > limits.max_xml_nodes {
            return Err(invalid(
                "font metric table exceeds the configured record limit",
            ));
        }
        let metric_bytes = metric_count
            .checked_mul(METRIC_RECORD_BYTES)
            .and_then(|length| metrics_offset.checked_add(length))
            .ok_or_else(|| invalid("font metric table length overflow"))?;
        if metrics_offset < HEADER_BYTES || metric_bytes != bytes.len() {
            return Err(invalid("font metric table length is inconsistent"));
        }

        let mut faces = HashMap::new();
        faces
            .try_reserve(face_count)
            .map_err(|_| allocation_failed())?;
        let mut face_offset = HEADER_BYTES;
        let mut expected_metric_start = 0_usize;
        for _ in 0..face_count {
            let record_end = face_offset
                .checked_add(face_record_bytes)
                .ok_or_else(|| invalid("font metric face offset overflow"))?;
            if record_end > metrics_offset {
                return Err(invalid("font metric face record is truncated"));
            }
            let family_length = read_u16(bytes, face_offset)? as usize;
            let style = bytes[face_offset + 2];
            let stretch = bytes[face_offset + 3];
            let weight = read_u16(bytes, face_offset + 4)?;
            let reserved = read_u16(bytes, face_offset + 6)?;
            let metric_start = read_u32(bytes, face_offset + 8)? as usize;
            let face_metric_count = read_u32(bytes, face_offset + 12)? as usize;
            let font_size = if version >= 3 {
                read_f32(bytes, face_offset + 28)?
            } else {
                0.0
            };
            if !font_size.is_finite() || font_size < 0.0 {
                return Err(invalid("font metric size is invalid"));
            }
            let vertical_metrics = if version >= 2 {
                let ascent_em = read_f32(bytes, face_offset + 16)?;
                let descent_em = read_f32(bytes, face_offset + 20)?;
                let line_gap_em = read_f32(bytes, face_offset + 24)?;
                let total = ascent_em + descent_em + line_gap_em;
                if !ascent_em.is_finite()
                    || ascent_em < 0.0
                    || !descent_em.is_finite()
                    || descent_em < 0.0
                    || !line_gap_em.is_finite()
                    || line_gap_em < 0.0
                    || !total.is_finite()
                    || total > MAX_ADVANCE_EM
                {
                    return Err(invalid("font vertical metrics are invalid"));
                }
                (total > 0.0).then_some([ascent_em, descent_em, line_gap_em])
            } else {
                None
            };
            if family_length == 0
                || family_length > MAX_FAMILY_BYTES
                || style > 2
                || stretch > 8
                || !(1..=1_000).contains(&weight)
                || if version == LEGACY_VERSION {
                    reserved != 0
                } else {
                    reserved & !(FALLBACK_FACE_FLAG | UNKNOWN_LINE_GAP_FLAG) != 0
                }
                || metric_start != expected_metric_start
            {
                return Err(invalid("font metric face record is invalid"));
            }
            let family_end = record_end
                .checked_add(family_length)
                .ok_or_else(|| invalid("font metric family length overflow"))?;
            if family_end > metrics_offset {
                return Err(invalid("font metric family is truncated"));
            }
            let family = core::str::from_utf8(&bytes[record_end..family_end])
                .map_err(|_| invalid("font metric family is not valid UTF-8"))?;
            let family =
                normalize_family(family).ok_or_else(|| invalid("font metric family is invalid"))?;
            let metric_end = metric_start
                .checked_add(face_metric_count)
                .ok_or_else(|| invalid("font metric range overflow"))?;
            if metric_end > metric_count {
                return Err(invalid("font metric face range is invalid"));
            }
            let mut advances = HashMap::new();
            advances
                .try_reserve(face_metric_count)
                .map_err(|_| allocation_failed())?;
            let mut previous_code_point = None;
            for metric_index in metric_start..metric_end {
                let offset = metrics_offset + metric_index * METRIC_RECORD_BYTES;
                let code_point = read_u32(bytes, offset)?;
                let advance_em = read_f32(bytes, offset + 4)?;
                if char::from_u32(code_point).is_none()
                    || previous_code_point.is_some_and(|previous| code_point <= previous)
                    || !advance_em.is_finite()
                    || !(0.0..=MAX_ADVANCE_EM).contains(&advance_em)
                {
                    return Err(invalid("font advance record is invalid"));
                }
                advances.insert(code_point, advance_em);
                previous_code_point = Some(code_point);
            }
            let key = FontFaceKey {
                family,
                style,
                weight,
                stretch,
                size_bits: if font_size == 0.0 {
                    0
                } else {
                    font_size.to_bits()
                },
            };
            if faces
                .insert(
                    key,
                    FontFaceMetrics {
                        advances,
                        vertical_metrics,
                        fallback: version >= 2 && reserved & FALLBACK_FACE_FLAG != 0,
                        line_gap_unknown: version >= 2 && reserved & UNKNOWN_LINE_GAP_FLAG != 0,
                    },
                )
                .is_some()
            {
                return Err(invalid("font metric table contains a duplicate face"));
            }
            expected_metric_start = metric_end;
            face_offset = family_end;
        }
        if face_offset != metrics_offset || expected_metric_start != metric_count {
            return Err(invalid("font metric table contains unreferenced data"));
        }
        Ok(Self {
            faces,
            metric_count,
        })
    }

    pub fn advance_em_at_size(
        &self,
        family: &str,
        italic: bool,
        bold: bool,
        character: char,
        font_size: f32,
    ) -> Option<f32> {
        let family = normalize_family(family)?;
        let mut key = FontFaceKey {
            family,
            style: u8::from(italic),
            weight: if bold { 700 } else { 400 },
            stretch: NORMAL_STRETCH,
            size_bits: if font_size > 0.0 && font_size.is_finite() {
                font_size.to_bits()
            } else {
                0
            },
        };
        let sized = self
            .faces
            .get(&key)
            .and_then(|metrics| metrics.advances.get(&(character as u32)))
            .copied();
        sized.or_else(|| {
            key.size_bits = 0;
            self.faces
                .get(&key)
                .and_then(|metrics| metrics.advances.get(&(character as u32)))
                .copied()
        })
    }

    pub fn line_height_em(
        &self,
        family: &str,
        italic: bool,
        bold: bool,
        minimum_if_incomplete: f32,
    ) -> Option<f32> {
        let family = normalize_family(family)?;
        let key = FontFaceKey {
            family,
            style: u8::from(italic),
            weight: if bold { 700 } else { 400 },
            stretch: NORMAL_STRETCH,
            size_bits: 0,
        };
        let metrics = self.faces.get(&key);
        // STZhongsong's hhea ascent/descent (1007/318) plus OS/2 line gap
        // (144), in 1000 units/em. Substitute fonts lose this extra leading.
        // Preserve authored line geometry without shipping the font itself;
        // an exact measured face still takes precedence.
        if !italic
            && !bold
            && matches!(key.family.as_str(), "华文中宋" | "stzhongsong")
            && metrics.is_none_or(|metrics| {
                metrics.fallback || metrics.line_gap_unknown || metrics.vertical_metrics.is_none()
            })
        {
            return Some(1469.0 / 1000.0);
        }
        metrics.and_then(|metrics| {
            metrics.vertical_metrics.map(|values| {
                let height: f32 = values.into_iter().sum();
                // ponytail: Canvas omits leading; retain the adapter's estimate
                // until font bytes supply complete metrics, without clipping tall ink.
                if metrics.line_gap_unknown {
                    // Arial and Times New Roman both use a 2355/2048 hhea line
                    // box across their four regular/bold/italic faces. The generic
                    // 1.31em estimate inflates their single spacing. Never apply
                    // this to substitutes or replace complete font metrics / taller ink.
                    let minimum = if matches!(key.family.as_str(), "arial" | "times new roman")
                        && !metrics.fallback
                    {
                        2355.0 / 2048.0
                    } else if !metrics.fallback
                        && (key.family == "superclarendon"
                            || key.family.starts_with("superclarendon-"))
                    {
                        // All Superclarendon faces: hhea 980 + 225 + 56 / 1000.
                        // Canvas exposes ascent/descent, but omits the 56-unit leading.
                        1261.0 / 1000.0
                    } else if !metrics.fallback
                        && !italic
                        && (key.family == "helveticaneue"
                            || key.family.starts_with("helveticaneue-")
                            || key.family == "helvetica neue"
                            || key.family.starts_with("helvetica neue "))
                    {
                        // Helvetica Neue hhea leading is 0.028em (regular) or
                        // 0.029em (medium/bold); Canvas supplies the other two terms.
                        height
                            + if bold
                                || key.family.contains("medium")
                                || key.family.ends_with("-bold")
                            {
                                0.029
                            } else {
                                0.028
                            }
                    } else {
                        minimum_if_incomplete
                    };
                    height.max(minimum)
                } else {
                    height
                }
            })
        })
    }

    pub fn point_rounded_line_height(
        &self,
        family: &str,
        italic: bool,
        bold: bool,
        size: f32,
    ) -> Option<f32> {
        let key = FontFaceKey {
            family: normalize_family(family)?,
            style: u8::from(italic),
            weight: if bold { 700 } else { 400 },
            stretch: NORMAL_STRETCH,
            size_bits: 0,
        };
        let metrics = self.faces.get(&key)?;
        let [ascent, descent, _] = metrics.vertical_metrics?;
        let height = self.line_height_em(family, italic, bold, 1.0)?;
        let leading = (height - ascent - descent).max(0.0);
        // NSFont rounds ascender, descender and leading separately in points.
        Some(
            ((size * ascent).round() + (size * descent).round() + (size * leading).round())
                .max(1.0),
        )
    }

    pub fn is_fallback_face(&self, family: &str, italic: bool, bold: bool) -> bool {
        let Some(family) = normalize_family(family) else {
            return false;
        };
        let key = FontFaceKey {
            family,
            style: u8::from(italic),
            weight: if bold { 700 } else { 400 },
            stretch: NORMAL_STRETCH,
            size_bits: 0,
        };
        self.faces.get(&key).is_some_and(|metrics| metrics.fallback)
    }

    pub fn is_empty(&self) -> bool {
        self.metric_count == 0
    }
}

fn normalize_family(value: &str) -> Option<String> {
    let value = value.trim();
    let value = if value.len() >= 2
        && matches!(value.as_bytes().first(), Some(b'\'') | Some(b'"'))
        && value.as_bytes().first() == value.as_bytes().last()
    {
        &value[1..value.len() - 1]
    } else {
        value
    };
    if value.is_empty() || value.chars().any(char::is_control) {
        None
    } else {
        Some(value.to_lowercase())
    }
}

fn read_u16(bytes: &[u8], offset: usize) -> Result<u16, Diagnostic> {
    let value = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| invalid("font metric table is truncated"))?;
    Ok(u16::from_le_bytes([value[0], value[1]]))
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32, Diagnostic> {
    let value = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| invalid("font metric table is truncated"))?;
    Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

fn read_f32(bytes: &[u8], offset: usize) -> Result<f32, Diagnostic> {
    Ok(f32::from_bits(read_u32(bytes, offset)?))
}

fn invalid(message: &str) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Layout, None, message)
}

fn allocation_failed() -> Diagnostic {
    Diagnostic::fatal(
        DiagnosticCode::AllocationFailed,
        Phase::Layout,
        None,
        "unable to allocate the bounded font metric table",
    )
}

#[cfg(test)]
mod tests {
    use super::FontMetricTable;
    use crate::limits::Limits;

    fn table(family: &str, code_point: u32, advance: f32) -> Vec<u8> {
        let family = family.as_bytes();
        let metrics_offset = 20 + 16 + family.len();
        let mut bytes = Vec::new();
        bytes.extend_from_slice(b"OVFM");
        bytes.extend_from_slice(&1_u16.to_le_bytes());
        bytes.extend_from_slice(&20_u16.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(&(metrics_offset as u32).to_le_bytes());
        bytes.extend_from_slice(&(family.len() as u16).to_le_bytes());
        bytes.push(0);
        bytes.push(4);
        bytes.extend_from_slice(&400_u16.to_le_bytes());
        bytes.extend_from_slice(&0_u16.to_le_bytes());
        bytes.extend_from_slice(&0_u32.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(family);
        bytes.extend_from_slice(&code_point.to_le_bytes());
        bytes.extend_from_slice(&advance.to_bits().to_le_bytes());
        bytes
    }

    fn table_v2(
        family: &str,
        code_point: u32,
        advance: f32,
        ascent: f32,
        descent: f32,
        line_gap: f32,
    ) -> Vec<u8> {
        let family = family.as_bytes();
        let metrics_offset = 20 + 28 + family.len();
        let mut bytes = Vec::new();
        bytes.extend_from_slice(b"OVFM");
        bytes.extend_from_slice(&2_u16.to_le_bytes());
        bytes.extend_from_slice(&20_u16.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(&(metrics_offset as u32).to_le_bytes());
        bytes.extend_from_slice(&(family.len() as u16).to_le_bytes());
        bytes.push(0);
        bytes.push(4);
        bytes.extend_from_slice(&400_u16.to_le_bytes());
        bytes.extend_from_slice(&0_u16.to_le_bytes());
        bytes.extend_from_slice(&0_u32.to_le_bytes());
        bytes.extend_from_slice(&1_u32.to_le_bytes());
        bytes.extend_from_slice(&ascent.to_bits().to_le_bytes());
        bytes.extend_from_slice(&descent.to_bits().to_le_bytes());
        bytes.extend_from_slice(&line_gap.to_bits().to_le_bytes());
        bytes.extend_from_slice(family);
        bytes.extend_from_slice(&code_point.to_le_bytes());
        bytes.extend_from_slice(&advance.to_bits().to_le_bytes());
        bytes
    }

    #[test]
    fn decodes_and_matches_a_complete_face() {
        let metrics =
            FontMetricTable::decode(&table("P0 Sans", 'W' as u32, 0.9), Limits::default()).unwrap();
        assert_eq!(
            metrics.advance_em_at_size("p0 sans", false, false, 'W', 12.0),
            Some(0.9)
        );
        assert_eq!(
            metrics.advance_em_at_size("P0 Sans", false, true, 'W', 12.0),
            None
        );
    }

    #[test]
    fn size_overrides_are_sparse_and_preserve_scalable_fallbacks() {
        let family = "P0 Sans";
        let mut bytes = table_v2(family, 'W' as u32, 0.9, 0.8, 0.2, 0.0);
        // Two v3 records: scalable W/i, and a 12px W correction.
        let mut base = bytes[20..48].to_vec();
        base[12..16].copy_from_slice(&2_u32.to_le_bytes());
        base.extend_from_slice(&0_f32.to_le_bytes());
        base.extend_from_slice(family.as_bytes());
        let mut sized = base.clone();
        sized[8..12].copy_from_slice(&2_u32.to_le_bytes());
        sized[12..16].copy_from_slice(&1_u32.to_le_bytes());
        sized[16..28].fill(0);
        sized[28..32].copy_from_slice(&12_f32.to_le_bytes());
        bytes.truncate(20);
        bytes[4..6].copy_from_slice(&3_u16.to_le_bytes());
        bytes[8..12].copy_from_slice(&2_u32.to_le_bytes());
        bytes[12..16].copy_from_slice(&3_u32.to_le_bytes());
        bytes[16..20].copy_from_slice(&((20 + base.len() + sized.len()) as u32).to_le_bytes());
        bytes.extend_from_slice(&base);
        bytes.extend_from_slice(&sized);
        for (character, advance) in [('W', 0.9_f32), ('i', 0.2), ('W', 0.91)] {
            bytes.extend_from_slice(&(character as u32).to_le_bytes());
            bytes.extend_from_slice(&advance.to_le_bytes());
        }
        let metrics = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
        assert_eq!(
            metrics.advance_em_at_size(family, false, false, 'W', 12.0),
            Some(0.91)
        );
        assert_eq!(
            metrics.advance_em_at_size(family, false, false, 'i', 12.0),
            Some(0.2)
        );
        for size in [0.0, 24.0, f32::NAN] {
            assert_eq!(
                metrics.advance_em_at_size(family, false, false, 'W', size),
                Some(0.9)
            );
        }
        assert_eq!(
            metrics.advance_em_at_size(family, false, true, 'W', 12.0),
            None
        );
        assert_eq!(
            metrics.advance_em_at_size(family, false, false, 'X', 12.0),
            None
        );
        assert_eq!(metrics.line_height_em(family, false, false, 0.0), Some(1.0));
        for invalid in [-1.0_f32, f32::NAN, f32::INFINITY] {
            let mut corrupt = bytes.clone();
            corrupt[48..52].copy_from_slice(&invalid.to_le_bytes());
            assert!(FontMetricTable::decode(&corrupt, Limits::default()).is_err());
        }
        // A size of zero would collide with the scalable face.
        bytes[20 + base.len() + 28..20 + base.len() + 32].fill(0);
        assert!(FontMetricTable::decode(&bytes, Limits::default()).is_err());
    }

    #[test]
    fn rejects_non_finite_advances() {
        let error =
            FontMetricTable::decode(&table("P0 Sans", 'W' as u32, f32::NAN), Limits::default())
                .unwrap_err();
        assert!(error.message.contains("advance"));
    }

    #[test]
    fn decodes_vertical_metrics_without_rejecting_v1_tables() {
        let metrics = FontMetricTable::decode(
            &table_v2("P0 Sans", 'W' as u32, 0.9, 0.92, 0.26, 0.04),
            Limits::default(),
        )
        .unwrap();
        assert_eq!(
            metrics.line_height_em("P0 Sans", false, false, 0.0),
            Some(1.22)
        );
        let legacy =
            FontMetricTable::decode(&table("P0 Sans", 'W' as u32, 0.9), Limits::default()).unwrap();
        assert_eq!(legacy.line_height_em("P0 Sans", false, false, 0.0), None);
        let mut incomplete = table_v2("P0 Sans", 'W' as u32, 0.9, 0.86, 0.14, 0.0);
        let complete = FontMetricTable::decode(&incomplete, Limits::default()).unwrap();
        assert_eq!(
            complete.line_height_em("P0 Sans", false, false, 1.31),
            Some(1.0)
        );
        incomplete[26..28].copy_from_slice(&super::UNKNOWN_LINE_GAP_FLAG.to_le_bytes());
        let incomplete = FontMetricTable::decode(&incomplete, Limits::default()).unwrap();
        for estimate in [1.2, 1.25, 1.31] {
            assert_eq!(
                incomplete.line_height_em("P0 Sans", false, false, estimate),
                Some(estimate)
            );
        }
    }

    #[test]
    fn rounds_native_font_line_components_separately() {
        // NSLayoutManager.defaultLineHeight(for:) on the supplied newsletter's
        // Helvetica Neue face, including the half-point leading boundary.
        let metrics = FontMetricTable::decode(
            &table_v2("HelveticaNeue", 'a' as u32, 0.5, 0.952, 0.213, 0.028),
            Limits::default(),
        )
        .unwrap();
        for (size, height) in [(10.0, 12.0), (11.0, 12.0), (12.0, 14.0), (18.0, 22.0)] {
            assert_eq!(
                metrics.point_rounded_line_height("HelveticaNeue", false, false, size),
                Some(height)
            );
        }
        let mut bytes = table_v2("HelveticaNeue", 'a' as u32, 0.5, 0.952, 0.213, 0.0);
        bytes[26..28].copy_from_slice(&super::UNKNOWN_LINE_GAP_FLAG.to_le_bytes());
        let measured = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
        assert_eq!(
            measured.point_rounded_line_height("HelveticaNeue", false, false, 18.0),
            Some(22.0)
        );
    }

    #[test]
    fn helvetica_neue_preserves_native_leading_without_inflating_the_line_box() {
        for (family, ascent, descent, leading) in [
            ("Helvetica Neue", 0.952, 0.213, 0.028),
            ("Helvetica Neue Medium", 0.975, 0.217, 0.029),
        ] {
            let mut bytes = table_v2(family, 'H' as u32, 0.7, ascent, descent, 0.0);
            bytes[26..28].copy_from_slice(&super::UNKNOWN_LINE_GAP_FLAG.to_le_bytes());
            let metrics = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
            assert!(
                (metrics.line_height_em(family, false, false, 1.31).unwrap()
                    - (ascent + descent + leading))
                    .abs()
                    < 0.0001
            );
            bytes[26..28].copy_from_slice(
                &(super::UNKNOWN_LINE_GAP_FLAG | super::FALLBACK_FACE_FLAG).to_le_bytes(),
            );
            let fallback = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
            assert_eq!(
                fallback.line_height_em(family, false, false, 1.31),
                Some(1.31)
            );
        }
    }

    #[test]
    fn superclarendon_preserves_known_leading_only_for_incomplete_exact_faces() {
        let mut bytes = table_v2("Superclarendon", 'a' as u32, 0.5, 0.98, 0.225, 0.0);
        let complete = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
        assert!(
            (complete
                .line_height_em("Superclarendon", false, false, 1.0)
                .unwrap()
                - 1.205)
                .abs()
                < 0.0001
        );
        bytes[26..28].copy_from_slice(&super::UNKNOWN_LINE_GAP_FLAG.to_le_bytes());
        let incomplete = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
        assert_eq!(
            (18.0
                * incomplete
                    .line_height_em("Superclarendon", false, false, 1.0)
                    .unwrap())
            .round(),
            23.0
        );
        bytes[26..28].copy_from_slice(
            &(super::UNKNOWN_LINE_GAP_FLAG | super::FALLBACK_FACE_FLAG).to_le_bytes(),
        );
        let fallback = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
        assert!(
            (fallback
                .line_height_em("Superclarendon", false, false, 1.0)
                .unwrap()
                - 1.205)
                .abs()
                < 0.0001
        );
    }

    #[test]
    fn times_new_roman_uses_known_leading_only_for_incomplete_exact_faces() {
        let mut bytes = table_v2(
            "Times New Roman",
            'a' as u32,
            0.444,
            1825.0 / 2048.0,
            443.0 / 2048.0,
            0.0,
        );
        let complete = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
        assert_eq!(
            complete.line_height_em("Times New Roman", false, false, 1.31),
            Some(2268.0 / 2048.0)
        );
        bytes[26..28].copy_from_slice(&super::UNKNOWN_LINE_GAP_FLAG.to_le_bytes());
        for (italic, bold) in [(false, false), (false, true), (true, false), (true, true)] {
            bytes[22] = u8::from(italic);
            bytes[24..26].copy_from_slice(&(if bold { 700_u16 } else { 400_u16 }).to_le_bytes());
            let incomplete = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
            assert_eq!(
                incomplete.line_height_em("Times New Roman", italic, bold, 1.31),
                Some(2355.0 / 2048.0)
            );
        }
        bytes[22] = 0;
        bytes[24..26].copy_from_slice(&400_u16.to_le_bytes());
        bytes[26..28].copy_from_slice(
            &(super::UNKNOWN_LINE_GAP_FLAG | super::FALLBACK_FACE_FLAG).to_le_bytes(),
        );
        let fallback = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
        assert_eq!(
            fallback.line_height_em("Times New Roman", false, false, 1.31),
            Some(1.31)
        );
        bytes[26..28].copy_from_slice(&super::UNKNOWN_LINE_GAP_FLAG.to_le_bytes());
        bytes[36..40].copy_from_slice(&1.5_f32.to_le_bytes());
        let tall = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
        assert!(
            tall.line_height_em("Times New Roman", false, false, 1.31)
                .unwrap()
                > 1.5
        );
    }

    #[test]
    fn arial_uses_its_known_hhea_leading_when_canvas_omits_it() {
        let mut bytes = table_v2(
            "Arial",
            'a' as u32,
            0.556,
            1854.0 / 2048.0,
            434.0 / 2048.0,
            0.0,
        );
        bytes[26..28].copy_from_slice(&super::UNKNOWN_LINE_GAP_FLAG.to_le_bytes());
        let metrics = FontMetricTable::decode(&bytes, Limits::default()).unwrap();

        assert_eq!(
            metrics.line_height_em("Arial", false, false, 1.31),
            Some(2355.0 / 2048.0)
        );
    }

    #[test]
    fn preserves_authored_zhongsong_leading_with_missing_or_substituted_fonts() {
        let missing = FontMetricTable::default();
        for family in ["华文中宋", "STZhongsong", " 'STZhongsong' "] {
            assert_eq!(
                missing.line_height_em(family, false, false, 0.0),
                Some(1.469)
            );
        }
        assert_eq!(missing.line_height_em("Unknown", false, false, 0.0), None);
        assert_eq!(
            missing.line_height_em("STZhongsong", true, false, 0.0),
            None
        );
        let mut bytes = table_v2("华文中宋", '三' as u32, 1.0, 0.9, 0.3, 0.0);
        let exact = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
        assert_eq!(
            exact.line_height_em("华文中宋", false, false, 0.0),
            Some(1.2)
        );
        bytes[26..28].copy_from_slice(&super::UNKNOWN_LINE_GAP_FLAG.to_le_bytes());
        let incomplete = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
        assert_eq!(
            incomplete.line_height_em("华文中宋", false, false, 1.31),
            Some(1.469)
        );
        bytes[26..28].copy_from_slice(&1_u16.to_le_bytes());
        let fallback = FontMetricTable::decode(&bytes, Limits::default()).unwrap();
        assert_eq!(
            fallback.line_height_em("华文中宋", false, false, 0.0),
            Some(1.469)
        );
        assert_eq!(
            fallback.advance_em_at_size("华文中宋", false, false, '三', 12.0),
            Some(1.0)
        );
    }
}
