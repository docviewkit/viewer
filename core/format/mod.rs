//! Content-based dispatch for native Office formats.

// Decimal digit grouping is identical for worksheet cells and ChartML labels.
#[cfg(any(
    feature = "native-formats",
    feature = "legacy-office-formats",
    feature = "odf-formats"
))]
fn group_decimal_digits(rendered: &str) -> String {
    let (sign, unsigned) = rendered
        .strip_prefix('-')
        .map_or(("", rendered), |v| ("-", v));
    let (integer, fraction) = unsigned.split_once('.').unwrap_or((unsigned, ""));
    let mut grouped = String::with_capacity(rendered.len() + integer.len() / 3);
    grouped.push_str(sign);
    for (index, character) in integer.chars().enumerate() {
        if index != 0 && (integer.len() - index).is_multiple_of(3) {
            grouped.push(',');
        }
        grouped.push(character);
    }
    if !fraction.is_empty() {
        grouped.push('.');
        grouped.push_str(fraction);
    }
    grouped
}

/// Remaining leading space when adjacent paragraph margins overlap.
#[cfg(any(
    feature = "native-formats",
    feature = "iwork-formats",
    feature = "odf-formats"
))]
fn collapsed_paragraph_space_before(before: f32, previous_after: f32) -> f32 {
    (before - previous_after).max(0.0)
}
#[cfg(any(
    feature = "native-formats",
    feature = "legacy-office-formats",
    feature = "iwork-formats",
    feature = "odf-formats"
))]
use crate::model::{FillRule, Geometry, PathCommand};
#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
use crate::model::{PathFillMode, PathLayer, Rect};
#[cfg(feature = "native-formats")]
pub mod docx;
#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
// Legacy charts use the shared spreadsheet renderer, not the XML adapters.
#[cfg_attr(not(feature = "native-formats"), allow(dead_code))]
mod drawingml;
#[cfg(any(feature = "native-formats", feature = "odf-formats"))]
mod embedded_font;
#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "iwork-formats",
    feature = "legacy-office-formats"
))]
mod embedded_media;
#[cfg(feature = "native-formats")]
pub mod flat;
#[cfg(feature = "iwork-formats")]
pub mod iwork;
#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub mod legacy;
#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
mod odf_chart;
#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
mod odf_math;
#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
pub mod odp;
#[cfg(feature = "odf-formats")]
pub mod ods;
#[cfg(feature = "odf-formats")]
pub mod odt;
#[cfg(feature = "native-formats")]
mod omml;
#[cfg(any(feature = "iwork-formats", feature = "pdf-formats"))]
pub mod pdf;
#[cfg(any(feature = "iwork-formats", feature = "pdf-formats"))]
mod pdf_crypto;
#[cfg(feature = "native-formats")]
pub mod pptx;
#[cfg(feature = "native-formats")]
mod vml;

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
fn optional_xml_attribute(
    attributes: &[crate::xml::XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<Option<String>, crate::diagnostic::Diagnostic> {
    attributes
        .iter()
        .find(|attribute| local_name(attribute.name) == name)
        .map(|attribute| {
            crate::xml::decode_xml_text(attribute.value)
                .map(|value| value.into_owned())
                .map_err(|error| with_part(error, part))
        })
        .transpose()
}
#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "ofd-formats",
    feature = "iwork-formats",
    feature = "legacy-office-formats"
))]
pub(super) mod presentation_image;
#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub mod security;
#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "iwork-formats",
    feature = "legacy-office-formats"
))]
mod symbol_font_mappings;
#[cfg(feature = "native-formats")]
pub mod xlsx;
#[cfg(feature = "calculation-service")]
mod xlsx_formula;

#[cfg(feature = "calculation-service")]
pub(crate) fn calculate_spreadsheet(
    bytes: &[u8],
    limits: crate::limits::Limits,
) -> Result<crate::calculation::CalculationResults, String> {
    xlsx_formula::calculate(bytes, limits)
}
#[cfg(feature = "ofd-formats")]
pub mod ofd;
#[cfg(feature = "xps-formats")]
pub mod xps;

#[cfg(any(feature = "native-formats", feature = "odf-formats"))]
pub(super) const DEFAULT_SHEET_COLUMN_WIDTH: f32 = 96.0;
#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn default_excel_column_width(maximum_digit_width: f32) -> f32 {
    (8.0 * maximum_digit_width + 5.0).floor()
}

/// Fit authored ODF table-column size hints into an available table width.
/// Missing or non-positive hints use the mean of positive hints (or unit weight),
/// then every column is scaled so the total equals `available`.
#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
pub(super) fn normalized_odf_table_sizes(widths: &[f32], count: usize, available: f32) -> Vec<f32> {
    if count == 0 {
        return Vec::new();
    }
    let positive = widths
        .iter()
        .take(count)
        .copied()
        .filter(|width| width.is_finite() && *width > 0.0)
        .collect::<Vec<_>>();
    let fallback = if positive.is_empty() {
        1.0
    } else {
        positive.iter().sum::<f32>() / positive.len() as f32
    };
    let raw = (0..count)
        .map(|index| {
            widths
                .get(index)
                .copied()
                .filter(|width| width.is_finite() && *width > 0.0)
                .unwrap_or(fallback)
        })
        .collect::<Vec<_>>();
    let total = raw.iter().sum::<f32>();
    if !total.is_finite() || total <= f32::EPSILON {
        return vec![available / count as f32; count];
    }
    raw.into_iter()
        .map(|width| width / total * available)
        .collect()
}

#[cfg(all(test, any(feature = "odf-formats", feature = "legacy-office-formats")))]
#[test]
fn normalized_odf_table_sizes_preserves_authored_ratios() {
    let sizes = normalized_odf_table_sizes(&[1.0, 3.0], 2, 400.0);
    assert_eq!(sizes, vec![100.0, 300.0]);
    let equal = normalized_odf_table_sizes(&[], 2, 400.0);
    assert_eq!(equal, vec![200.0, 200.0]);
    let mixed = normalized_odf_table_sizes(&[2.0, 0.0], 2, 300.0);
    assert_eq!(mixed, vec![150.0, 150.0]);
}

/// DOC and DOCX select the same variants after adapter-specific inheritance.
/// A missing selected variant stays absent; it must not fall back to default.
#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn select_word_story<T>(
    default: T,
    first: T,
    even: T,
    use_first: bool,
    use_even: bool,
) -> T {
    if use_first {
        first
    } else if use_even {
        even
    } else {
        default
    }
}

#[cfg(all(
    test,
    any(feature = "native-formats", feature = "legacy-office-formats")
))]
#[test]
fn word_story_selection_prioritizes_first_without_default_fallback() {
    assert_eq!(
        select_word_story(Some(0), Some(1), Some(2), false, false),
        Some(0)
    );
    assert_eq!(
        select_word_story(Some(0), Some(1), Some(2), false, true),
        Some(2)
    );
    assert_eq!(
        select_word_story(Some(0), Some(1), Some(2), true, true),
        Some(1)
    );
    assert_eq!(select_word_story(Some(0), None, Some(2), true, false), None);
    assert_eq!(select_word_story(Some(0), Some(1), None, false, true), None);
}

/// DOC and DOCX frame a separate source paragraph; Pages caps are inline glyphs.
#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn word_drop_cap_bounds(
    x: f32,
    y: f32,
    width: f32,
    lines: u32,
    line_height: f32,
    space: f32,
    in_margin: bool,
) -> (crate::model::Rect, Option<crate::model::Rect>) {
    let bounds = crate::model::Rect {
        x: if in_margin { x - width - space } else { x },
        y,
        width,
        height: lines as f32 * line_height,
    };
    let exclusion = (!in_margin).then_some(crate::model::Rect {
        width: width + space,
        ..bounds
    });
    (bounds, exclusion)
}

#[cfg(all(
    test,
    any(feature = "native-formats", feature = "legacy-office-formats")
))]
#[test]
fn word_drop_cap_frame_and_margin_share_line_geometry() {
    let (cap, wrap) = word_drop_cap_bounds(96.0, 120.0, 30.0, 3, 20.0, 4.0, false);
    assert_eq!(
        cap,
        crate::model::Rect {
            x: 96.0,
            y: 120.0,
            width: 30.0,
            height: 60.0
        }
    );
    assert_eq!(wrap.unwrap(), crate::model::Rect { width: 34.0, ..cap });
    let (margin, wrap) = word_drop_cap_bounds(96.0, 120.0, 30.0, 3, 20.0, 4.0, true);
    assert_eq!(margin, crate::model::Rect { x: 62.0, ..cap });
    assert_eq!(wrap, None);
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(super) const DEFAULT_SHEET_ROW_HEIGHT: f32 = 24.0;

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn excel_indexed_color(index: u32) -> Option<u32> {
    const COLORS: [&str; 64] = [
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000",
        "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000",
        "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC",
        "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080", "FF00FF", "FFFF00", "00FFFF",
        "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF",
        "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600",
        "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399",
        "333333",
    ];
    COLORS.get(index as usize).map(|color| {
        (u32::from_str_radix(color, 16).expect("constant Excel palette color") << 8) | 0xff
    })
}

#[cfg(all(
    test,
    any(feature = "native-formats", feature = "legacy-office-formats")
))]
#[test]
fn excel_indexed_palette_preserves_chart_colors_and_reserved_boundary() {
    assert_eq!(excel_indexed_color(29), Some(0xff8080ff));
    assert_eq!(excel_indexed_color(60), Some(0x993300ff));
    assert_eq!(excel_indexed_color(63), Some(0x333333ff));
    assert_eq!(excel_indexed_color(64), None);
    assert_eq!(excel_indexed_color(u32::MAX), None);
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(super) fn office_chart_palette_color(index: usize) -> u32 {
    const COLORS: [u32; 6] = [
        0x4472_c4ff,
        0xed7d_31ff,
        0xa5a5_a5ff,
        0xffc0_00ff,
        0x5b9b_d5ff,
        0x70ad_47ff,
    ];
    COLORS[index % COLORS.len()]
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(super) fn alphabetic_number(mut value: u32, uppercase: bool) -> String {
    if value == 0 {
        return "0".to_owned();
    }
    let base = if uppercase { b'A' } else { b'a' };
    let mut characters = Vec::new();
    while value != 0 {
        value -= 1;
        characters.push((base + (value % 26) as u8) as char);
        value /= 26;
    }
    characters.into_iter().rev().collect()
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(super) fn roman_number(mut value: u32, uppercase: bool) -> String {
    if value == 0 || value > 3_999 {
        return value.to_string();
    }
    let mut output = String::new();
    for (number, numeral) in [
        (1_000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ] {
        while value >= number {
            output.push_str(numeral);
            value -= number;
        }
    }
    if uppercase {
        output
    } else {
        output.to_ascii_lowercase()
    }
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(super) fn circled_number(value: u32) -> Option<String> {
    let character = match value {
        1..=20 => char::from_u32(0x2460 + value - 1),
        21..=35 => char::from_u32(0x3251 + value - 21),
        36..=50 => char::from_u32(0x32b1 + value - 36),
        _ => None,
    }?;
    Some(character.to_string())
}

#[cfg(feature = "native-formats")]
pub(super) fn enclosed_fullstop_number(value: u32) -> Option<String> {
    (1..=20)
        .contains(&value)
        .then_some(0x2487 + value)
        .and_then(char::from_u32)
        .map(|character| character.to_string())
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(super) fn chinese_number(value: u32) -> String {
    const DIGITS: [&str; 10] = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
    match value {
        0..=9 => DIGITS[value as usize].to_owned(),
        10 => "十".to_owned(),
        11..=19 => format!("十{}", DIGITS[(value % 10) as usize]),
        20..=99 => {
            let tens = DIGITS[(value / 10) as usize];
            let ones = value % 10;
            if ones == 0 {
                format!("{tens}十")
            } else {
                format!("{tens}十{}", DIGITS[ones as usize])
            }
        }
        _ => value.to_string(),
    }
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
/// ODF paragraph alignment; spreadsheet cells retain their value-dependent defaults.
pub(super) fn odf_text_align(value: &str) -> crate::model::TextAlign {
    use crate::model::TextAlign;
    match value {
        "center" => TextAlign::Center,
        "right" | "end" => TextAlign::End,
        "justify" => TextAlign::Justify,
        _ => TextAlign::Start,
    }
}

/// ODF font-family is a list; commas inside a quoted family are literal.
#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(super) fn odf_primary_font_family(value: &str) -> &str {
    let value = value.trim();
    if let Some(quote @ ('\'' | '"')) = value.chars().next() {
        let family = &value[1..];
        family.split(quote).next().unwrap_or(family).trim()
    } else {
        value.split(',').next().unwrap_or(value).trim()
    }
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
pub(super) fn zero_padded_number(value: u32, format: &str) -> Option<String> {
    let width = if format == "decimalZero" {
        2
    } else {
        let example = format.split(',').next()?.trim();
        if example.len() < 2
            || !example.ends_with('1')
            || !example[..example.len() - 1]
                .bytes()
                .all(|byte| byte == b'0')
        {
            return None;
        }
        example.len()
    };
    Some(format!("{value:0width$}"))
}

#[cfg(any(feature = "odf-formats", feature = "legacy-office-formats"))]
pub(super) fn odf_border_stroke_style(pattern: &str, width: f32) -> crate::model::StrokeStyle {
    use crate::model::{LineCap, LineCompound, StrokeStyle};

    let unit = width.max(1.0);
    match pattern {
        "dotted" => StrokeStyle {
            cap: LineCap::Round,
            dash: vec![unit, unit * 2.0],
            ..StrokeStyle::default()
        },
        "dashed" => StrokeStyle {
            dash: vec![unit * 8.0, unit * 4.0],
            ..StrokeStyle::default()
        },
        "fine-dashed" => StrokeStyle {
            dash: vec![unit * 3.0, unit * 3.0],
            ..StrokeStyle::default()
        },
        "dash-dot" => StrokeStyle {
            dash: vec![unit * 8.0, unit * 4.0, unit * 2.0, unit * 4.0],
            ..StrokeStyle::default()
        },
        "dash-dot-dot" => StrokeStyle {
            dash: vec![
                unit * 8.0,
                unit * 4.0,
                unit * 2.0,
                unit * 4.0,
                unit * 2.0,
                unit * 4.0,
            ],
            ..StrokeStyle::default()
        },
        "double" | "double-thin" => StrokeStyle {
            compound: LineCompound::Double,
            ..StrokeStyle::default()
        },
        _ => StrokeStyle::default(),
    }
}

#[cfg(any(
    feature = "native-formats",
    feature = "legacy-office-formats",
    feature = "iwork-formats",
    feature = "odf-formats"
))]
pub(super) fn polygon_geometry(points: &[(f32, f32)]) -> Geometry {
    Geometry::Path {
        fill_rule: FillRule::NonZero,
        commands: points
            .iter()
            .enumerate()
            .map(|(index, &(x, y))| {
                if index == 0 {
                    PathCommand::MoveTo { x, y }
                } else {
                    PathCommand::LineTo { x, y }
                }
            })
            .chain(std::iter::once(PathCommand::ClosePath))
            .collect(),
    }
}

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn diamond_geometry(width: f32, height: f32) -> Geometry {
    polygon_geometry(&[
        (width / 2.0, 0.0),
        (width, height / 2.0),
        (width / 2.0, height),
        (0.0, height / 2.0),
    ])
}

// Adapters resolve their own adjustment units into the four inner edges.
#[cfg(any(
    feature = "native-formats",
    feature = "legacy-office-formats",
    feature = "iwork-formats"
))]
pub(super) fn cross_geometry(
    width: f32,
    height: f32,
    left: f32,
    right: f32,
    top: f32,
    bottom: f32,
) -> Geometry {
    polygon_geometry(&[
        (left, 0.0),
        (right, 0.0),
        (right, top),
        (width, top),
        (width, bottom),
        (right, bottom),
        (right, height),
        (left, height),
        (left, bottom),
        (0.0, bottom),
        (0.0, top),
        (left, top),
    ])
}

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn regular_polygon_geometry(
    width: f32,
    height: f32,
    sides: usize,
    rotation: f32,
) -> Geometry {
    let center_x = width / 2.0;
    let center_y = height / 2.0;
    let points = (0..sides)
        .map(|index| {
            let angle = rotation + index as f32 * std::f32::consts::TAU / sides as f32;
            (
                center_x + center_x * angle.cos(),
                center_y + center_y * angle.sin(),
            )
        })
        .collect::<Vec<_>>();
    polygon_geometry(&points)
}

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn star_geometry(width: f32, height: f32, points: usize, inner_radius: f32) -> Geometry {
    let center_x = width / 2.0;
    let center_y = height / 2.0;
    let vertices = (0..points * 2)
        .map(|index| {
            let angle =
                -std::f32::consts::FRAC_PI_2 + index as f32 * std::f32::consts::PI / points as f32;
            let radius = if index % 2 == 0 { 1.0 } else { inner_radius };
            (
                center_x + center_x * radius * angle.cos(),
                center_y + center_y * radius * angle.sin(),
            )
        })
        .collect::<Vec<_>>();
    polygon_geometry(&vertices)
}

#[cfg(all(
    test,
    any(
        feature = "native-formats",
        feature = "odf-formats",
        feature = "legacy-office-formats"
    )
))]
mod numbering_tests {
    #[cfg(feature = "native-formats")]
    use super::enclosed_fullstop_number;
    #[cfg(feature = "odf-formats")]
    use super::odf_border_stroke_style;
    #[cfg(any(feature = "native-formats", feature = "odf-formats"))]
    use super::zero_padded_number;
    use super::{alphabetic_number, chinese_number, circled_number, roman_number};

    #[test]
    fn formats_shared_office_numbering_boundaries() {
        assert_eq!(alphabetic_number(27, false), "aa");
        assert_eq!(roman_number(3_999, true), "MMMCMXCIX");
        assert_eq!(roman_number(4_000, false), "4000");
        assert_eq!(circled_number(50).as_deref(), Some("㊿"));
        assert_eq!(circled_number(51), None);
        #[cfg(feature = "native-formats")]
        {
            assert_eq!(enclosed_fullstop_number(1).as_deref(), Some("⒈"));
            assert_eq!(enclosed_fullstop_number(20).as_deref(), Some("⒛"));
            assert_eq!(enclosed_fullstop_number(21), None);
        }
        assert_eq!(chinese_number(99), "九十九");
        #[cfg(any(feature = "native-formats", feature = "odf-formats"))]
        {
            assert_eq!(zero_padded_number(1, "decimalZero").as_deref(), Some("01"));
            assert_eq!(
                zero_padded_number(1, "001, 002, 003, ...").as_deref(),
                Some("001")
            );
            assert_eq!(
                zero_padded_number(1, "0001, 0002, 0003, ...").as_deref(),
                Some("0001")
            );
            assert_eq!(
                zero_padded_number(1, "00001, 00002, 00003, ...").as_deref(),
                Some("00001")
            );
        }
        #[cfg(feature = "odf-formats")]
        {
            assert_eq!(odf_border_stroke_style("dotted", 1.0).dash, [1.0, 2.0]);
            assert_eq!(
                odf_border_stroke_style("dash-dot-dot", 1.0).dash,
                [8.0, 4.0, 2.0, 4.0, 2.0, 4.0]
            );
            assert_eq!(
                odf_border_stroke_style("double-thin", 1.0).compound,
                crate::model::LineCompound::Double
            );
        }
    }
}

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn smiley_geometry(width: f32, height: f32) -> Geometry {
    let center_x = width / 2.0;
    let center_y = height / 2.0;
    let kappa = 0.552_284_8;
    let outline = vec![
        PathCommand::MoveTo {
            x: center_x,
            y: 0.0,
        },
        PathCommand::BezierCurveTo {
            cp1x: center_x + center_x * kappa,
            cp1y: 0.0,
            cp2x: width,
            cp2y: center_y - center_y * kappa,
            x: width,
            y: center_y,
        },
        PathCommand::BezierCurveTo {
            cp1x: width,
            cp1y: center_y + center_y * kappa,
            cp2x: center_x + center_x * kappa,
            cp2y: height,
            x: center_x,
            y: height,
        },
        PathCommand::BezierCurveTo {
            cp1x: center_x - center_x * kappa,
            cp1y: height,
            cp2x: 0.0,
            cp2y: center_y + center_y * kappa,
            x: 0.0,
            y: center_y,
        },
        PathCommand::BezierCurveTo {
            cp1x: 0.0,
            cp1y: center_y - center_y * kappa,
            cp2x: center_x - center_x * kappa,
            cp2y: 0.0,
            x: center_x,
            y: 0.0,
        },
        PathCommand::ClosePath,
    ];
    let mut details = Vec::new();
    for center_x in [width * 0.33, width * 0.67] {
        let radius_x = width * 0.05;
        let radius_y = height * 0.055;
        details.extend([
            PathCommand::MoveTo {
                x: center_x + radius_x,
                y: height * 0.34,
            },
            PathCommand::BezierCurveTo {
                cp1x: center_x + radius_x,
                cp1y: height * 0.34 + radius_y * kappa,
                cp2x: center_x + radius_x * kappa,
                cp2y: height * 0.34 + radius_y,
                x: center_x,
                y: height * 0.34 + radius_y,
            },
            PathCommand::BezierCurveTo {
                cp1x: center_x - radius_x * kappa,
                cp1y: height * 0.34 + radius_y,
                cp2x: center_x - radius_x,
                cp2y: height * 0.34 + radius_y * kappa,
                x: center_x - radius_x,
                y: height * 0.34,
            },
            PathCommand::BezierCurveTo {
                cp1x: center_x - radius_x,
                cp1y: height * 0.34 - radius_y * kappa,
                cp2x: center_x - radius_x * kappa,
                cp2y: height * 0.34 - radius_y,
                x: center_x,
                y: height * 0.34 - radius_y,
            },
            PathCommand::BezierCurveTo {
                cp1x: center_x + radius_x * kappa,
                cp1y: height * 0.34 - radius_y,
                cp2x: center_x + radius_x,
                cp2y: height * 0.34 - radius_y * kappa,
                x: center_x + radius_x,
                y: height * 0.34,
            },
            PathCommand::ClosePath,
        ]);
    }
    details.extend([
        PathCommand::MoveTo {
            x: width * 0.25,
            y: height * 0.60,
        },
        PathCommand::QuadraticCurveTo {
            cpx: center_x,
            cpy: height * 0.88,
            x: width * 0.75,
            y: height * 0.60,
        },
    ]);
    Geometry::LayeredPath {
        layers: vec![
            PathLayer {
                fill_rule: FillRule::NonZero,
                fill: PathFillMode::Normal,
                stroke: true,
                commands: outline,
            },
            PathLayer {
                fill_rule: FillRule::NonZero,
                fill: PathFillMode::None,
                stroke: true,
                commands: details,
            },
        ],
    }
}

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn horizontal_action_button_geometry(bounds: Rect, forward: bool) -> Geometry {
    let inset = bounds.width.min(bounds.height) * 3.0 / 8.0;
    let center_x = bounds.width / 2.0;
    let center_y = bounds.height / 2.0;
    let direction = if forward { 1.0 } else { -1.0 };
    let triangle = vec![
        PathCommand::MoveTo {
            x: center_x + inset * direction,
            y: center_y,
        },
        PathCommand::LineTo {
            x: center_x - inset * direction,
            y: center_y - inset,
        },
        PathCommand::LineTo {
            x: center_x - inset * direction,
            y: center_y + inset,
        },
        PathCommand::ClosePath,
    ];
    let rectangle = vec![
        PathCommand::MoveTo { x: 0.0, y: 0.0 },
        PathCommand::LineTo {
            x: bounds.width,
            y: 0.0,
        },
        PathCommand::LineTo {
            x: bounds.width,
            y: bounds.height,
        },
        PathCommand::LineTo {
            x: 0.0,
            y: bounds.height,
        },
        PathCommand::ClosePath,
    ];
    let mut base = rectangle.clone();
    base.extend(triangle.iter().cloned());
    let layer = |fill, stroke, commands| PathLayer {
        fill_rule: FillRule::NonZero,
        fill,
        stroke,
        commands,
    };
    Geometry::LayeredPath {
        layers: vec![
            layer(PathFillMode::Normal, false, base),
            layer(PathFillMode::Darken, false, triangle.clone()),
            layer(PathFillMode::None, true, triangle),
            layer(PathFillMode::None, true, rectangle),
        ],
    }
}

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn lightning_bolt_geometry(bounds: Rect) -> Geometry {
    let point = |x: f32, y: f32| (bounds.width * x / 21_600.0, bounds.height * y / 21_600.0);
    polygon_geometry(&[
        point(8_472.0, 0.0),
        point(12_860.0, 6_080.0),
        point(11_050.0, 6_797.0),
        point(16_577.0, 12_007.0),
        point(14_767.0, 12_877.0),
        point(21_600.0, 21_600.0),
        point(10_012.0, 14_915.0),
        point(12_222.0, 13_987.0),
        point(5_022.0, 9_705.0),
        point(7_602.0, 8_382.0),
        point(0.0, 3_890.0),
    ])
}

/// Converts characters authored through legacy symbol fonts into stable Unicode.
///
/// Office, OpenDocument, and iWork formats carry the same visual bullets
/// through different structures. Keeping the conversion here prevents their
/// adapters from drifting apart and avoids depending on a symbol font being
/// installed in the browser.
#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "iwork-formats",
    feature = "legacy-office-formats"
))]
pub(super) fn normalize_symbol_font_character(value: &str, font_family: Option<&str>) -> String {
    let Some(font_family) = font_family else {
        return value.to_owned();
    };
    let trimmed_family = font_family.trim().trim_matches(['\'', '"']);
    let semantic_family = if trimmed_family.eq_ignore_ascii_case("Wingdings")
        || trimmed_family.eq_ignore_ascii_case("Wingdings-Regular")
    {
        "Wingdings"
    } else if trimmed_family.eq_ignore_ascii_case("Symbol")
        || trimmed_family.eq_ignore_ascii_case("SymbolMT")
    {
        "Symbol"
    } else {
        trimmed_family
    };
    if semantic_family.eq_ignore_ascii_case("Webdings") {
        let mut normalized = String::new();
        for character in value.chars() {
            let code_point = character as u32;
            let source = if code_point <= 0xff {
                Some(code_point)
            } else if (0xf000..=0xf0ff).contains(&code_point) {
                Some(code_point & 0xff)
            } else {
                None
            };
            if source == Some(0x4e) {
                normalized.push('👁');
            } else {
                normalized.push(character);
            }
        }
        return normalized;
    }
    value
        .chars()
        .map(|character| {
            let code_point = character as u32;
            let source = if code_point <= 0xff {
                Some(code_point)
            } else if (0xf000..=0xf0ff).contains(&code_point) {
                Some(code_point & 0xff)
            } else {
                None
            };
            if semantic_family.eq_ignore_ascii_case("Wingdings") {
                match source {
                    Some(0x6c) => return '●',
                    Some(0x70) => return '□',
                    Some(0xa8) => return '□',
                    Some(0xb2) => return '✧',
                    Some(0xd8) => return '➢',
                    _ => {}
                }
            }
            if semantic_family.eq_ignore_ascii_case("Wingdings 2") && source == Some(0xa4) {
                return '□';
            }
            symbol_font_mappings::semantic_symbol_font_character(character, semantic_family)
                .unwrap_or(character)
        })
        .collect()
}

#[cfg(all(
    test,
    any(
        feature = "native-formats",
        feature = "odf-formats",
        feature = "iwork-formats",
        feature = "legacy-office-formats"
    )
))]
mod symbol_font_tests {
    use super::normalize_symbol_font_character;

    #[test]
    fn normalizes_legacy_symbol_fonts_from_source_bytes_and_private_use_aliases() {
        assert_eq!(normalize_symbol_font_character("l", Some("Wingdings")), "●");
        assert_eq!(normalize_symbol_font_character("p", Some("Wingdings")), "□");
        assert_eq!(
            normalize_symbol_font_character("p", Some("Wingdings-Regular")),
            "□"
        );
        assert_eq!(normalize_symbol_font_character("Ø", Some("Wingdings")), "➢");
        assert_eq!(
            normalize_symbol_font_character("\u{f0b2}", Some("Wingdings")),
            "✧"
        );
        assert_eq!(normalize_symbol_font_character("q", Some("Wingdings")), "❑");
        assert_eq!(normalize_symbol_font_character("u", Some("wingdings")), "◆");
        assert_eq!(
            normalize_symbol_font_character("\u{f0a8}", Some("Wingdings")),
            "□"
        );
        assert_eq!(
            normalize_symbol_font_character("\u{f071}", Some("Wingdings")),
            "❑"
        );
        assert_eq!(
            normalize_symbol_font_character("\u{f075}", Some("Wingdings")),
            "◆"
        );
        assert_eq!(normalize_symbol_font_character("§", Some("Wingdings")), "▪");
        assert_eq!(
            normalize_symbol_font_character("\u{f0a7}", Some("Wingdings")),
            "▪"
        );
        assert_eq!(
            normalize_symbol_font_character("!;Jü", Some("Wingdings")),
            "🖉🖴☺✓"
        );
        assert_eq!(
            normalize_symbol_font_character("ü", Some("'Wingdings'")),
            "✓"
        );
        assert_eq!(
            normalize_symbol_font_character("\u{f0b7}", Some("Symbol")),
            "•"
        );
        assert_eq!(
            normalize_symbol_font_character("\u{f0d2}\u{f0d3}\u{f0d4}", Some("Symbol")),
            "®©™"
        );
        assert_eq!(
            normalize_symbol_font_character("\u{f0e2}\u{f0e3}\u{f0e4}", Some("Symbol")),
            "®©™"
        );
        assert_eq!(
            normalize_symbol_font_character("\u{f0a3}x\u{f0b3}", Some("Symbol")),
            "≤ξ≥"
        );
        assert_eq!(
            normalize_symbol_font_character("ABG", Some("Symbol")),
            "ΑΒΓ"
        );
        assert_eq!(
            normalize_symbol_font_character("@\\", Some("\"Symbol\"")),
            "≅∴"
        );
        assert_eq!(normalize_symbol_font_character("q", Some("Arial")), "q");
        assert_eq!(
            normalize_symbol_font_character("\u{f0a4}", Some("Wingdings 2")),
            "□"
        );
        assert_eq!(normalize_symbol_font_character("q", None), "q");
        assert_eq!(
            normalize_symbol_font_character("\u{f04e}", Some("Webdings")),
            "👁"
        );
    }
}

#[cfg(any(feature = "native-formats", feature = "xps-formats"))]
use std::collections::HashMap;

#[cfg(feature = "native-formats")]
use crate::RetainedInput;
use crate::diagnostic::Diagnostic;
#[cfg(feature = "native-formats")]
use crate::diagnostic::Fidelity;
#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "xps-formats"
))]
use crate::diagnostic::{DiagnosticCode, Phase};
use crate::font_metrics::FontMetricTable;
use crate::limits::Limits;
use crate::model::Document;
#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "xps-formats"
))]
use crate::package::Package;
#[cfg(feature = "native-formats")]
use crate::package::PackageCache;
#[cfg(any(feature = "native-formats", feature = "xps-formats"))]
use crate::xml::{XmlAttribute, XmlEvent, decode_xml_text, parse_xml};

#[cfg(feature = "native-formats")]
const PPTX_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
#[cfg(feature = "native-formats")]
const PPTM_MAIN: &str = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml";
#[cfg(feature = "native-formats")]
const PPSX_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml";
#[cfg(feature = "native-formats")]
const PPSM_MAIN: &str = "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml";
#[cfg(feature = "native-formats")]
const POTX_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml";
#[cfg(feature = "native-formats")]
const POTM_MAIN: &str = "application/vnd.ms-powerpoint.template.macroEnabled.main+xml";
#[cfg(feature = "native-formats")]
const XLSX_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
#[cfg(feature = "native-formats")]
const XLSM_MAIN: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
#[cfg(feature = "native-formats")]
const XLTX_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
#[cfg(feature = "native-formats")]
const XLTM_MAIN: &str = "application/vnd.ms-excel.template.macroEnabled.main+xml";
#[cfg(feature = "native-formats")]
const DOCX_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
#[cfg(feature = "native-formats")]
const DOCM_MAIN: &str = "application/vnd.ms-word.document.macroEnabled.main+xml";
#[cfg(feature = "native-formats")]
const DOTX_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
#[cfg(feature = "native-formats")]
const DOTM_MAIN: &str = "application/vnd.ms-word.template.macroEnabledTemplate.main+xml";
#[cfg(feature = "odf-formats")]
const ODP_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.presentation";
#[cfg(feature = "odf-formats")]
const OTP_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.presentation-template";
#[cfg(feature = "odf-formats")]
const ODG_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.graphics";
#[cfg(feature = "odf-formats")]
const OTG_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.graphics-template";
#[cfg(feature = "odf-formats")]
const ODS_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.spreadsheet";
#[cfg(feature = "odf-formats")]
const OTS_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.spreadsheet-template";
#[cfg(feature = "odf-formats")]
const ODT_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.text";
#[cfg(feature = "odf-formats")]
const OTT_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.text-template";

#[cfg(feature = "odf-formats")]
fn is_open_document_presentation_mimetype(mimetype: &[u8]) -> bool {
    matches!(
        mimetype,
        ODP_MIMETYPE | OTP_MIMETYPE | ODG_MIMETYPE | OTG_MIMETYPE
    )
}

#[cfg(feature = "odf-formats")]
fn is_open_document_spreadsheet_mimetype(mimetype: &[u8]) -> bool {
    mimetype == ODS_MIMETYPE || mimetype == OTS_MIMETYPE
}

#[cfg(feature = "odf-formats")]
fn is_open_document_text_mimetype(mimetype: &[u8]) -> bool {
    mimetype == ODT_MIMETYPE || mimetype == OTT_MIMETYPE
}

#[cfg(feature = "odf-formats")]
fn flat_open_document_mimetype(bytes: &[u8], limits: Limits) -> Result<Option<String>, Diagnostic> {
    let xml = bytes
        .strip_prefix(&[0xef, 0xbb, 0xbf])
        .unwrap_or(bytes)
        .iter()
        .copied()
        .skip_while(u8::is_ascii_whitespace)
        .next()
        == Some(b'<');
    if !xml {
        return Ok(None);
    }

    let mut root_checked = false;
    let mut mimetype = None;
    crate::xml::parse_xml(bytes, limits, |event| {
        if root_checked {
            return Ok(());
        }
        let crate::xml::XmlEvent::StartElement {
            name, attributes, ..
        } = event
        else {
            return Ok(());
        };
        root_checked = true;
        if local_name(name) == "document" {
            mimetype = attributes
                .iter()
                .find(|attribute| local_name(attribute.name) == "mimetype")
                .map(|attribute| crate::xml::decode_xml_text(attribute.value))
                .transpose()?
                .map(|value| value.into_owned());
        }
        Ok(())
    })?;
    Ok(mimetype)
}

#[cfg(feature = "native-formats")]
fn is_presentation_content_type(content_type: &str) -> bool {
    matches!(
        content_type,
        PPTX_MAIN | PPTM_MAIN | PPSX_MAIN | PPSM_MAIN | POTX_MAIN | POTM_MAIN
    )
}

#[cfg(feature = "native-formats")]
fn is_spreadsheet_content_type(content_type: &str) -> bool {
    matches!(content_type, XLSX_MAIN | XLSM_MAIN | XLTX_MAIN | XLTM_MAIN)
}

#[cfg(feature = "native-formats")]
fn is_text_content_type(content_type: &str) -> bool {
    matches!(content_type, DOCX_MAIN | DOCM_MAIN | DOTX_MAIN | DOTM_MAIN)
}

#[cfg(feature = "native-formats")]
fn is_macro_enabled_content_type(content_type: &str) -> bool {
    matches!(
        content_type,
        PPTM_MAIN | PPSM_MAIN | POTM_MAIN | XLSM_MAIN | XLTM_MAIN | DOCM_MAIN | DOTM_MAIN
    )
}

#[cfg(any(feature = "native-formats", feature = "xps-formats"))]
#[derive(Clone, Default)]
pub(super) struct ContentTypes {
    defaults: HashMap<String, String>,
    overrides: HashMap<String, String>,
}

#[cfg(any(feature = "native-formats", feature = "xps-formats"))]
impl ContentTypes {
    fn override_for_part(&self, part: &str) -> Option<&str> {
        self.overrides.get(part).map(String::as_str)
    }

    pub(super) fn for_part(&self, part: &str) -> Option<&str> {
        self.override_for_part(part).or_else(|| {
            let (_, extension) = part.rsplit_once('.')?;
            self.defaults
                .get(&extension.to_ascii_lowercase())
                .map(String::as_str)
        })
    }
}

#[derive(Clone, Copy, Debug)]
#[cfg_attr(not(feature = "native-formats"), allow(dead_code))]
pub(crate) struct FieldDateTime {
    pub year: u32,
    pub month: u32,
    pub day: u32,
    pub hour: u32,
    pub minute: u32,
}

/// Detects and parses a supported package without consulting a filename.
pub fn detect_and_parse(bytes: &[u8], limits: Limits) -> Result<Option<Document>, Diagnostic> {
    detect_and_parse_with_font_metrics(bytes, limits, &FontMetricTable::default())
}

/// Detects and parses a supported package with optional browser-measured font
/// advances. Only layout-aware adapters consume the table.
pub fn detect_and_parse_with_font_metrics(
    bytes: &[u8],
    limits: Limits,
    font_metrics: &FontMetricTable,
) -> Result<Option<Document>, Diagnostic> {
    detect_and_parse_with_font_metrics_at(bytes, limits, font_metrics, None)
}

pub(crate) fn detect_and_parse_with_font_metrics_at(
    bytes: &[u8],
    limits: Limits,
    _font_metrics: &FontMetricTable,
    _field_date_time: Option<FieldDateTime>,
) -> Result<Option<Document>, Diagnostic> {
    #[cfg(not(any(
        feature = "native-formats",
        feature = "odf-formats",
        feature = "iwork-formats",
        feature = "legacy-office-formats",
        feature = "pdf-formats",
        feature = "xps-formats",
        feature = "ofd-formats"
    )))]
    let _ = (bytes, limits);
    #[cfg(feature = "iwork-formats")]
    if let Some(document) = iwork::detect_and_parse_with_font_metrics(bytes, limits, _font_metrics)?
    {
        return Ok(Some(document));
    }
    #[cfg(feature = "legacy-office-formats")]
    if let Some(document) =
        legacy::detect_and_parse_with_font_metrics(bytes, limits, _font_metrics)?
    {
        return Ok(Some(document));
    }
    #[cfg(feature = "odf-formats")]
    if let Some(document) = detect_and_parse_odf(bytes, limits, _font_metrics)? {
        return Ok(Some(document));
    }
    #[cfg(feature = "pdf-formats")]
    if let Some(document) = pdf::detect_and_parse(bytes, limits)? {
        return Ok(Some(document));
    }
    #[cfg(feature = "xps-formats")]
    if let Some(document) = xps::detect_and_parse(bytes, limits)? {
        return Ok(Some(document));
    }
    #[cfg(feature = "ofd-formats")]
    if let Some(document) = ofd::detect_and_parse(bytes, limits)? {
        return Ok(Some(document));
    }
    #[cfg(feature = "native-formats")]
    if let Some(document) = detect_and_parse_native(bytes, limits, _font_metrics, _field_date_time)?
    {
        return Ok(Some(document));
    }
    Ok(None)
}

#[cfg(feature = "odf-formats")]
fn detect_and_parse_odf(
    bytes: &[u8],
    limits: Limits,
    font_metrics: &FontMetricTable,
) -> Result<Option<Document>, Diagnostic> {
    if !bytes.starts_with(b"PK") {
        let mimetype = flat_open_document_mimetype(bytes, limits)
            .map_err(|error| error.in_part("content.xml"))?;
        return match mimetype.as_deref() {
            Some(
                "application/vnd.oasis.opendocument.presentation"
                | "application/vnd.oasis.opendocument.graphics",
            ) => odp::parse_flat(bytes, limits, font_metrics).map(Some),
            Some("application/vnd.oasis.opendocument.spreadsheet") => {
                ods::parse_flat(bytes, limits).map(Some)
            }
            Some("application/vnd.oasis.opendocument.text") => {
                odt::parse_flat_with_font_metrics(bytes, limits, font_metrics).map(Some)
            }
            _ => Ok(None),
        };
    }
    let package = Package::open(bytes, limits)?;
    let security_diagnostics = security::inspect(&package)?;
    let mut document = match package.part("mimetype")?.as_deref() {
        Some(mimetype) if is_open_document_presentation_mimetype(mimetype) => {
            odp::parse_with_font_metrics(&package, font_metrics)?
        }
        Some(mimetype) if is_open_document_spreadsheet_mimetype(mimetype) => ods::parse(&package)?,
        Some(mimetype) if is_open_document_text_mimetype(mimetype) => {
            odt::parse_with_font_metrics(&package, font_metrics)?
        }
        _ => return Ok(None),
    };
    document = with_diagnostics(document, security_diagnostics);
    document.embedded_fonts = embedded_font::extract_odf(&package, &mut document.diagnostics);
    Ok(Some(document))
}

#[cfg(feature = "native-formats")]
fn detect_and_parse_native(
    bytes: &[u8],
    limits: Limits,
    font_metrics: &FontMetricTable,
    field_date_time: Option<FieldDateTime>,
) -> Result<Option<Document>, Diagnostic> {
    if let Some(document) = flat::detect_and_parse(bytes, limits, font_metrics)? {
        return Ok(Some(document));
    }
    if !bytes.starts_with(b"PK") {
        return Ok(None);
    }
    let package = Package::open(bytes, limits)?;
    let security_diagnostics = security::inspect(&package)?;
    if !package.has_part("[Content_Types].xml") {
        return Ok(None);
    }
    if is_xps_package(&package)? {
        return Ok(None);
    }
    let content_types = parse_content_types(&package)?;
    let relationships = package.relationships(None)?;
    let office_relationships: Vec<_> = relationships
        .iter()
        .filter(|relationship| relationship.type_uri.ends_with("/officeDocument"))
        .collect();
    if office_relationships.len() != 1 {
        return Err(format_error(
            "_rels/.rels",
            "OPC package must contain exactly one officeDocument relationship",
        ));
    }
    let office_relationship = office_relationships[0];
    if office_relationship.external {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ExternalResourceBlocked,
            Phase::Security,
            None,
            "external officeDocument relationship is forbidden",
        )
        .in_part("_rels/.rels"));
    }
    let Some(content_type) = content_types.for_part(&office_relationship.target) else {
        return Err(format_error(
            "[Content_Types].xml",
            "officeDocument part has no matching content type",
        ));
    };
    let mut diagnostics = security_diagnostics;
    if is_macro_enabled_content_type(content_type)
        && !diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message.contains("macro"))
    {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ActiveContentBlocked,
                Phase::Security,
                Fidelity::Blocked,
                "macro-enabled document detected; macros will not execute",
            )
            .in_part(&office_relationship.target),
        );
    }
    if is_spreadsheet_content_type(content_type) {
        return xlsx::parse(&package, &office_relationship.target, &content_types, None)
            .map(|document| with_diagnostics(document, diagnostics))
            .map(Some);
    }
    if is_text_content_type(content_type) {
        return docx::parse_with_font_metrics_at(
            &package,
            &office_relationship.target,
            diagnostics,
            font_metrics,
            field_date_time,
        )
        .map(Some);
    }
    if is_presentation_content_type(content_type) {
        return pptx::parse(
            &package,
            &office_relationship.target,
            diagnostics,
            &content_types,
            font_metrics,
        )
        .map(Some);
    }
    Ok(None)
}

#[cfg(feature = "native-formats")]
pub(crate) struct PreparedXlsx {
    bytes: std::sync::Arc<RetainedInput>,
    limits: Limits,
    cache: PackageCache,
    workbook_part: String,
    content_types: ContentTypes,
    diagnostics: Vec<Diagnostic>,
    requires_calculation: bool,
    calculation: Option<std::sync::Arc<crate::calculation::CalculationResults>>,
}

#[cfg(feature = "native-formats")]
pub(crate) struct PreparedPptx {
    bytes: std::sync::Arc<RetainedInput>,
    limits: Limits,
    cache: PackageCache,
    presentation_part: String,
    content_types: ContentTypes,
    diagnostics: Vec<Diagnostic>,
    initial_diagnostics: Vec<Diagnostic>,
    font_metrics: FontMetricTable,
}

#[cfg(feature = "native-formats")]
impl PreparedPptx {
    pub(crate) const fn limits(&self) -> Limits {
        self.limits
    }

    pub(crate) fn materialize(&self) -> Result<Document, Diagnostic> {
        let package = Package::open_with_cache(&self.bytes, self.limits, self.cache.clone())?;
        pptx::parse(
            &package,
            &self.presentation_part,
            self.diagnostics.clone(),
            &self.content_types,
            &self.font_metrics,
        )
    }

    pub(crate) fn materialize_unit_with_budget(
        &self,
        unit_index: u32,
        materialized_image_bytes: usize,
    ) -> Result<(Document, usize), Diagnostic> {
        let package = Package::open_with_cache(&self.bytes, self.limits, self.cache.clone())?;
        let (mut document, materialized_image_bytes) = pptx::parse_unit(
            &package,
            &self.presentation_part,
            &self.content_types,
            &self.font_metrics,
            unit_index,
            materialized_image_bytes,
        )?;
        document
            .diagnostics
            .retain(|diagnostic| !self.initial_diagnostics.contains(diagnostic));
        Ok((document, materialized_image_bytes))
    }
}

#[cfg(feature = "native-formats")]
impl PreparedXlsx {
    pub(crate) const fn limits(&self) -> Limits {
        self.limits
    }

    pub(crate) const fn requires_calculation(&self) -> bool {
        self.requires_calculation && self.calculation.is_none()
    }

    pub(crate) fn apply_calculation(
        &mut self,
        calculation: crate::calculation::CalculationResults,
    ) {
        self.calculation = Some(std::sync::Arc::new(calculation));
    }

    pub(crate) fn materialize(&self) -> Result<Document, Diagnostic> {
        let package = Package::open_with_cache(&self.bytes, self.limits, self.cache.clone())?;
        xlsx::parse(
            &package,
            &self.workbook_part,
            &self.content_types,
            self.calculation.as_deref(),
        )
        .map(|document| with_diagnostics(document, self.diagnostics.clone()))
    }

    pub(crate) fn materialize_unit(&self, unit_index: u32) -> Result<Document, Diagnostic> {
        let package = Package::open_with_cache(&self.bytes, self.limits, self.cache.clone())?;
        xlsx::parse_unit(
            &package,
            &self.workbook_part,
            &self.content_types,
            unit_index,
            self.calculation.as_deref(),
        )
    }

    pub(crate) fn materialize_unit_region(
        &self,
        unit_index: u32,
        max_row: i32,
        max_column: i32,
    ) -> Result<Document, Diagnostic> {
        let package = Package::open_with_cache(&self.bytes, self.limits, self.cache.clone())?;
        xlsx::parse_unit_region(
            &package,
            &self.workbook_part,
            &self.content_types,
            unit_index,
            max_row,
            max_column,
            self.calculation.as_deref(),
        )
    }
}

#[cfg(feature = "native-formats")]
pub(crate) enum PreparedNative {
    Pptx {
        document: Document,
        loaded_units: std::collections::BTreeSet<u32>,
        materialized_image_bytes: usize,
        prepared: PreparedPptx,
    },
    Xlsx {
        document: Document,
        prepared: PreparedXlsx,
    },
}

#[cfg(feature = "native-formats")]
pub(crate) fn prepare_native_retained(
    bytes: std::sync::Arc<RetainedInput>,
    limits: Limits,
    font_metrics: &FontMetricTable,
) -> Result<Option<PreparedNative>, Diagnostic> {
    let Some(package) = prepare_ooxml_package(&bytes, limits)? else {
        return Ok(None);
    };
    if is_presentation_content_type(&package.content_type) {
        let initial = pptx::parse_initial(
            &package.package,
            &package.main_part,
            package.diagnostics.clone(),
            &package.content_types,
            font_metrics,
        )?;
        let initial_diagnostics = initial.document.diagnostics.clone();
        return Ok(Some(PreparedNative::Pptx {
            document: initial.document,
            loaded_units: initial.loaded_units,
            materialized_image_bytes: initial.materialized_image_bytes,
            prepared: PreparedPptx {
                bytes: std::sync::Arc::clone(&bytes),
                limits,
                cache: package.package.cache(),
                presentation_part: package.main_part,
                content_types: package.content_types,
                diagnostics: package.diagnostics,
                initial_diagnostics,
                font_metrics: font_metrics.clone(),
            },
        }));
    }
    if is_spreadsheet_content_type(&package.content_type) {
        let document = with_diagnostics(
            xlsx::parse_metadata(&package.package, &package.main_part, &package.content_types)?,
            package.diagnostics.clone(),
        );
        let requires_calculation = document
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message == crate::calculation::REQUIRED_MESSAGE);
        return Ok(Some(PreparedNative::Xlsx {
            document,
            prepared: PreparedXlsx {
                bytes: std::sync::Arc::clone(&bytes),
                limits,
                cache: package.package.cache(),
                workbook_part: package.main_part,
                content_types: package.content_types,
                diagnostics: package.diagnostics,
                requires_calculation,
                calculation: None,
            },
        }));
    }
    Ok(None)
}

#[cfg(feature = "native-formats")]
struct PreparedOoxmlPackage<'a> {
    package: Package<'a>,
    content_types: ContentTypes,
    main_part: String,
    content_type: String,
    diagnostics: Vec<Diagnostic>,
}

#[cfg(feature = "native-formats")]
fn prepare_ooxml_package(
    bytes: &[u8],
    limits: Limits,
) -> Result<Option<PreparedOoxmlPackage<'_>>, Diagnostic> {
    if !bytes.starts_with(b"PK") {
        return Ok(None);
    }
    let package = Package::open(bytes, limits)?;
    if !package.has_part("[Content_Types].xml") {
        return Ok(None);
    }
    if is_xps_package(&package)? {
        return Ok(None);
    }
    let content_types = parse_content_types(&package)?;
    let relationships = package.relationships(None)?;
    let office_relationships = relationships
        .iter()
        .filter(|relationship| relationship.type_uri.ends_with("/officeDocument"))
        .collect::<Vec<_>>();
    if office_relationships.len() != 1 || office_relationships[0].external {
        return Ok(None);
    }
    let main_part = office_relationships[0].target.clone();
    let Some(content_type) = content_types.for_part(&main_part).map(str::to_owned) else {
        return Ok(None);
    };
    let mut diagnostics = security::inspect(&package)?;
    if is_macro_enabled_content_type(&content_type)
        && !diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message.contains("macro"))
    {
        diagnostics.push(
            Diagnostic::warning(
                DiagnosticCode::ActiveContentBlocked,
                Phase::Security,
                Fidelity::Blocked,
                "macro-enabled document detected; macros will not execute",
            )
            .in_part(&main_part),
        );
    }
    Ok(Some(PreparedOoxmlPackage {
        package,
        content_types,
        main_part,
        content_type,
        diagnostics,
    }))
}

#[cfg(any(feature = "native-formats", feature = "odf-formats"))]
fn with_diagnostics(mut document: Document, mut diagnostics: Vec<Diagnostic>) -> Document {
    diagnostics.append(&mut document.diagnostics);
    document.diagnostics = diagnostics;
    document
}

#[cfg(feature = "native-formats")]
fn is_xps_package(package: &Package<'_>) -> Result<bool, Diagnostic> {
    // Identify the owning parser before interpreting optional OPC metadata.
    if !package
        .relationships_filtered(None, Some("/officedocument"))?
        .is_empty()
    {
        return Ok(false);
    }
    if !package
        .relationships_filtered(None, Some("/fixedrepresentation"))?
        .is_empty()
    {
        return Ok(true);
    }
    Ok(find_xps_sequence(package, None)?.is_some())
}

#[cfg(any(feature = "native-formats", feature = "xps-formats"))]
fn find_xps_sequence(
    package: &Package<'_>,
    diagnostics: Option<&mut Vec<Diagnostic>>,
) -> Result<Option<String>, Diagnostic> {
    if !package.has_part("[Content_Types].xml") {
        return Ok(None);
    }
    let types = parse_content_types_with_diagnostics(package, diagnostics)?;
    let mut parts = package.entry_names().filter(|part| {
        matches!(
            types.for_part(part),
            Some(
                "application/vnd.ms-package.xps-fixeddocumentsequence+xml"
                    | "application/oxps-fixeddocumentsequence+xml"
            )
        )
    });
    let result = parts.next().map(str::to_owned);
    if parts.next().is_some() {
        return Err(format_error(
            "[Content_Types].xml",
            "ambiguous XPS document sequence without a fixedrepresentation relationship",
        ));
    }
    Ok(result)
}

#[cfg(feature = "native-formats")]
fn parse_content_types(package: &Package<'_>) -> Result<ContentTypes, Diagnostic> {
    parse_content_types_with_diagnostics(package, None)
}

#[cfg(any(feature = "native-formats", feature = "xps-formats"))]
fn parse_content_types_with_diagnostics(
    package: &Package<'_>,
    mut diagnostics: Option<&mut Vec<Diagnostic>>,
) -> Result<ContentTypes, Diagnostic> {
    let part = "[Content_Types].xml";
    let bytes = package.required_part(part)?;
    let mut content_types = ContentTypes::default();
    parse_xml(&bytes, package.limits(), |event| {
        let XmlEvent::StartElement {
            name, attributes, ..
        } = event
        else {
            return Ok(());
        };
        let result = (|| {
            match local_name(name) {
                "Override" => {
                    let part_name = required_attribute(&attributes, "PartName", part)?;
                    let content_type = required_attribute(&attributes, "ContentType", part)?;
                    let part_name = part_name.trim_start_matches('/').to_owned();
                    if part_name.is_empty()
                        || content_type.trim().is_empty()
                        || content_types.overrides.contains_key(&part_name)
                    {
                        return Err(format_error(
                            part,
                            format!("duplicate or empty content-type override: {part_name}"),
                        ));
                    }
                    content_types.overrides.insert(part_name, content_type);
                }
                "Default" => {
                    let extension = required_attribute(&attributes, "Extension", part)?;
                    let content_type = required_attribute(&attributes, "ContentType", part)?;
                    let extension = extension
                        .trim()
                        .trim_start_matches('.')
                        .to_ascii_lowercase();
                    if extension.is_empty()
                        || content_type.trim().is_empty()
                        || content_types.defaults.contains_key(&extension)
                    {
                        return Err(format_error(
                            part,
                            format!("duplicate or empty content-type default: {extension}"),
                        ));
                    }
                    content_types.defaults.insert(extension, content_type);
                }
                _ => {}
            }
            Ok(())
        })();
        if let (Err(error), Some(diagnostics)) = (&result, diagnostics.as_deref_mut()) {
            if error.code == DiagnosticCode::FormatInvalid {
                let mut warning = error.clone();
                warning.severity = crate::diagnostic::Severity::Warning;
                warning.fidelity = crate::diagnostic::Fidelity::Omitted;
                diagnostics.push(warning);
                return Ok(());
            }
        }
        result
    })
    .map_err(|error| with_part(error, part))?;
    Ok(content_types)
}

#[cfg(any(feature = "native-formats", feature = "xps-formats"))]
fn required_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<String, Diagnostic> {
    let attribute = attributes
        .iter()
        .find(|attribute| local_name(attribute.name) == name)
        .ok_or_else(|| format_error(part, format!("element is missing {name}")))?;
    decode_xml_text(attribute.value)
        .map(|value| value.into_owned())
        .map_err(|error| with_part(error, part))
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "iwork-formats",
    feature = "pdf-formats",
    feature = "xps-formats",
    feature = "calculation-service",
    feature = "legacy-office-formats"
))]
fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

#[cfg(any(feature = "native-formats", feature = "xps-formats"))]
fn format_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message).in_part(part)
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats",
    feature = "xps-formats"
))]
fn with_part(error: Diagnostic, part: &str) -> Diagnostic {
    if error.location.part.is_some() {
        error
    } else {
        error.in_part(part)
    }
}

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
fn rgb_channel(color: u32, shift: u32) -> f32 {
    ((color >> shift) & 0xff) as f32 / 255.0
}

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
fn hsl_component(p: f32, q: f32, mut hue: f32) -> f32 {
    if hue < 0.0 {
        hue += 1.0;
    }
    if hue > 1.0 {
        hue -= 1.0;
    }
    if hue < 1.0 / 6.0 {
        p + (q - p) * 6.0 * hue
    } else if hue < 0.5 {
        q
    } else if hue < 2.0 / 3.0 {
        p + (q - p) * (2.0 / 3.0 - hue) * 6.0
    } else {
        p
    }
}

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn color_to_hsl(color: u32) -> (f32, f32, f32) {
    let r = rgb_channel(color, 24);
    let g = rgb_channel(color, 16);
    let b = rgb_channel(color, 8);
    let maximum = r.max(g).max(b);
    let minimum = r.min(g).min(b);
    let mut hue = 0.0_f32;
    let mut saturation = 0.0_f32;
    let luminance = (maximum + minimum) / 2.0;
    if maximum != minimum {
        let delta = maximum - minimum;
        saturation = if luminance > 0.5 {
            delta / (2.0 - maximum - minimum)
        } else {
            delta / (maximum + minimum)
        };
        hue = if maximum == r {
            (g - b) / delta + if g < b { 6.0 } else { 0.0 }
        } else if maximum == g {
            (b - r) / delta + 2.0
        } else {
            (r - g) / delta + 4.0
        } / 6.0;
    }
    (hue, saturation, luminance)
}

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn color_from_hsl(color: u32, hue: f32, saturation: f32, luminance: f32) -> u32 {
    let (r, g, b) = if saturation == 0.0 {
        (luminance, luminance, luminance)
    } else {
        let q = if luminance < 0.5 {
            luminance * (1.0 + saturation)
        } else {
            luminance + saturation - luminance * saturation
        };
        let p = 2.0 * luminance - q;
        (
            hsl_component(p, q, hue + 1.0 / 3.0),
            hsl_component(p, q, hue),
            hsl_component(p, q, hue - 1.0 / 3.0),
        )
    };
    let channel = |value: f32| (value * 255.0).round() as u32;
    (channel(r) << 24) | (channel(g) << 16) | (channel(b) << 8) | (color & 0xff)
}

#[cfg(any(feature = "native-formats", feature = "legacy-office-formats"))]
pub(super) fn transform_luminance(color: u32, multiplier: f32, offset: f32) -> u32 {
    let (hue, saturation, luminance) = color_to_hsl(color);
    color_from_hsl(
        color,
        hue,
        saturation,
        (luminance * multiplier + offset).clamp(0.0, 1.0),
    )
}

#[cfg(feature = "native-formats")]
pub(super) fn excel_serial_date(serial_day: i64, date_1904: bool) -> (i64, u32, u32) {
    if !date_1904 && serial_day == 60 {
        return (1900, 2, 29);
    }
    let base = if date_1904 {
        days_from_civil(1904, 1, 1)
    } else {
        days_from_civil(1899, 12, 31)
    };
    let leap_adjustment = i64::from(!date_1904 && serial_day > 60);
    civil_from_days(base + serial_day - leap_adjustment)
}

#[cfg(feature = "native-formats")]
pub(super) fn excel_serial_from_date(year: i64, month: u32, day: u32, date_1904: bool) -> i64 {
    let base = if date_1904 {
        days_from_civil(1904, 1, 1)
    } else {
        days_from_civil(1899, 12, 31)
    };
    let serial = days_from_civil(year, month, day) - base;
    serial + i64::from(!date_1904 && serial >= 60)
}

#[cfg(feature = "native-formats")]
fn days_from_civil(year: i64, month: u32, day: u32) -> i64 {
    let year = year - i64::from(month <= 2);
    let era = if year >= 0 { year } else { year - 399 } / 400;
    let year_of_era = year - era * 400;
    let month = i64::from(month);
    let day_of_year = (153 * (month + if month > 2 { -3 } else { 9 }) + 2) / 5 + i64::from(day) - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era * 146_097 + day_of_era - 719_468
}

#[cfg(feature = "native-formats")]
fn civil_from_days(days: i64) -> (i64, u32, u32) {
    let days = days + 719_468;
    let era = if days >= 0 { days } else { days - 146_096 } / 146_097;
    let day_of_era = days - era * 146_097;
    let year_of_era =
        (day_of_era - day_of_era / 1_460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let mut year = year_of_era + era * 400;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_prime = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_prime + 2) / 5 + 1;
    let month = month_prime + if month_prime < 10 { 3 } else { -9 };
    year += i64::from(month <= 2);
    (year, month as u32, day as u32)
}

#[cfg(any(feature = "native-formats", feature = "odf-formats"))]
pub(super) fn reserve_materialized_text_bytes(
    total: &mut usize,
    bytes_per_copy: usize,
    copies: usize,
    limit: usize,
    part: &str,
) -> Result<(), Diagnostic> {
    let additional = bytes_per_copy.checked_mul(copies).ok_or_else(|| {
        Diagnostic::fatal(
            DiagnosticCode::ZipTotalSizeLimit,
            Phase::Parse,
            None,
            "materialized document text size overflow",
        )
        .in_part(part)
    })?;
    let next = total.checked_add(additional).ok_or_else(|| {
        Diagnostic::fatal(
            DiagnosticCode::ZipTotalSizeLimit,
            Phase::Parse,
            None,
            "materialized document text budget overflow",
        )
        .in_part(part)
    })?;
    if next > limit {
        return Err(Diagnostic::fatal(
            DiagnosticCode::ZipTotalSizeLimit,
            Phase::Parse,
            None,
            "materialized document text exceeds the configured uncompressed-byte budget",
        )
        .in_part(part));
    }
    *total = next;
    Ok(())
}

#[cfg(any(feature = "native-formats", feature = "odf-formats"))]
pub(super) fn clone_materialized_text(value: &str, part: &str) -> Result<String, Diagnostic> {
    let mut clone = String::new();
    clone.try_reserve_exact(value.len()).map_err(|_| {
        Diagnostic::fatal(
            DiagnosticCode::AllocationFailed,
            Phase::Parse,
            None,
            "unable to allocate materialized document text",
        )
        .in_part(part)
    })?;
    clone.push_str(value);
    Ok(clone)
}

#[cfg(all(test, any(feature = "native-formats", feature = "odf-formats")))]
mod tests {
    #[test]
    fn odf_font_family_list_preserves_quoted_names() {
        for (value, expected) in [
            ("'Liberation Sans', Arial", "Liberation Sans"),
            ("\"Family, With Comma\", serif", "Family, With Comma"),
            (" Arial, sans-serif ", "Arial"),
            ("Times New Roman", "Times New Roman"),
            ("'Unclosed Family", "Unclosed Family"),
            ("+mn-lt", "+mn-lt"),
            ("", ""),
        ] {
            assert_eq!(super::odf_primary_font_family(value), expected);
        }
    }

    #[test]
    fn odf_paragraph_alignment_preserves_justification_and_fallback() {
        use crate::model::TextAlign;
        for (value, expected) in [
            ("justify", TextAlign::Justify),
            ("center", TextAlign::Center),
            ("right", TextAlign::End),
            ("end", TextAlign::End),
            ("start", TextAlign::Start),
            ("unknown", TextAlign::Start),
        ] {
            assert_eq!(super::odf_text_align(value), expected);
        }
    }

    #[test]
    fn shared_decimal_grouping_preserves_sign_fraction_and_padding() {
        for (input, expected) in [
            ("1234.50", "1,234.50"),
            ("-1234567.89 ", "-1,234,567.89 "),
            ("0007.8", "0,007.8"),
            ("0", "0"),
        ] {
            assert_eq!(super::group_decimal_digits(input), expected);
        }
    }

    #[cfg(any(
        feature = "native-formats",
        feature = "odf-formats",
        feature = "legacy-office-formats"
    ))]
    #[test]
    fn shared_xml_attributes_preserve_first_match_decoding_and_error_location() {
        use crate::xml::XmlAttribute;
        let attributes = [
            XmlAttribute {
                name: "a:value",
                value: "A&amp;B",
            },
            XmlAttribute {
                name: "b:value",
                value: "ignored",
            },
        ];
        assert_eq!(
            super::optional_xml_attribute(&attributes, "value", "part.xml")
                .unwrap()
                .as_deref(),
            Some("A&B")
        );
        assert_eq!(
            super::optional_xml_attribute(&attributes, "missing", "part.xml").unwrap(),
            None
        );
        let attributes = [XmlAttribute {
            name: "value",
            value: "&unknown;",
        }];
        let expected = super::with_part(
            crate::xml::decode_xml_text("&unknown;").unwrap_err(),
            "part.xml",
        );
        assert_eq!(
            super::optional_xml_attribute(&attributes, "value", "part.xml").unwrap_err(),
            expected
        );
    }

    #[cfg(any(
        feature = "native-formats",
        feature = "iwork-formats",
        feature = "odf-formats"
    ))]
    #[test]
    fn adjacent_paragraph_spacing_overlaps_without_becoming_negative() {
        for (before, after, remaining) in [
            (24.0, 12.0, 12.0),
            (12.0, 24.0, 0.0),
            (0.0, 0.0, 0.0),
            (12.0, 12.0, 0.0),
            (12.0, 0.0, 12.0),
        ] {
            assert_eq!(
                super::collapsed_paragraph_space_before(before, after),
                remaining
            );
        }
    }
    use super::detect_and_parse;
    #[cfg(feature = "native-formats")]
    use super::{
        POTM_MAIN, POTX_MAIN, PPSM_MAIN, PPSX_MAIN, PPTX_MAIN, is_macro_enabled_content_type,
        is_presentation_content_type, is_spreadsheet_content_type, is_text_content_type,
    };
    #[cfg(feature = "odf-formats")]
    use super::{
        is_open_document_presentation_mimetype, is_open_document_spreadsheet_mimetype,
        is_open_document_text_mimetype,
    };
    use crate::limits::Limits;
    #[cfg(feature = "native-formats")]
    use crate::model::{
        DocumentFormat, DocumentKind, Geometry, MappingQuality, ObjectKind, SourceLocator,
        TextAlign, UnitKind, Visual,
    };

    #[cfg(feature = "native-formats")]
    #[test]
    fn shared_cross_preserves_independent_inner_edges() {
        use crate::model::PathCommand::{ClosePath, LineTo, MoveTo};
        let Geometry::Path { commands, .. } =
            super::cross_geometry(120.0, 80.0, 15.0, 105.0, 20.0, 60.0)
        else {
            panic!("cross path");
        };
        assert_eq!(
            commands,
            vec![
                MoveTo { x: 15.0, y: 0.0 },
                LineTo { x: 105.0, y: 0.0 },
                LineTo { x: 105.0, y: 20.0 },
                LineTo { x: 120.0, y: 20.0 },
                LineTo { x: 120.0, y: 60.0 },
                LineTo { x: 105.0, y: 60.0 },
                LineTo { x: 105.0, y: 80.0 },
                LineTo { x: 15.0, y: 80.0 },
                LineTo { x: 15.0, y: 60.0 },
                LineTo { x: 0.0, y: 60.0 },
                LineTo { x: 0.0, y: 20.0 },
                LineTo { x: 15.0, y: 20.0 },
                ClosePath,
            ]
        );
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn shared_diamond_preserves_office_vertices() {
        use crate::model::PathCommand::{ClosePath, LineTo, MoveTo};
        let Geometry::Path { commands, .. } = super::diamond_geometry(96.0, 64.0) else {
            panic!("diamond path");
        };
        assert_eq!(
            commands,
            vec![
                MoveTo { x: 48.0, y: 0.0 },
                LineTo { x: 96.0, y: 32.0 },
                LineTo { x: 48.0, y: 64.0 },
                LineTo { x: 0.0, y: 32.0 },
                ClosePath,
            ]
        );
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn recognizes_ooxml_template_and_slideshow_main_parts() {
        for content_type in [
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
            "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
            "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
            "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
            "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
        ] {
            assert!(is_presentation_content_type(content_type), "{content_type}");
        }
        for content_type in [
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            "application/vnd.ms-excel.template.macroEnabled.main+xml",
        ] {
            assert!(is_spreadsheet_content_type(content_type), "{content_type}");
        }
        for content_type in [
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            "application/vnd.ms-word.document.macroEnabled.main+xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
            "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
        ] {
            assert!(is_text_content_type(content_type), "{content_type}");
        }
        for content_type in [
            "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
            "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
            "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "application/vnd.ms-excel.template.macroEnabled.main+xml",
            "application/vnd.ms-word.document.macroEnabled.main+xml",
            "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
        ] {
            assert!(
                is_macro_enabled_content_type(content_type),
                "{content_type}"
            );
        }
    }

    #[cfg(feature = "odf-formats")]
    #[test]
    fn recognizes_open_document_template_mimetypes() {
        assert!(is_open_document_presentation_mimetype(
            b"application/vnd.oasis.opendocument.presentation-template"
        ));
        assert!(is_open_document_presentation_mimetype(
            b"application/vnd.oasis.opendocument.graphics-template"
        ));
        assert!(is_open_document_spreadsheet_mimetype(
            b"application/vnd.oasis.opendocument.spreadsheet-template"
        ));
        assert!(is_open_document_text_mimetype(
            b"application/vnd.oasis.opendocument.text-template"
        ));
    }

    #[cfg(feature = "odf-formats")]
    #[test]
    fn parses_flat_open_document_text_through_the_public_dispatch() {
        let bytes = br#"<?xml version="1.0"?>
          <office:document xmlns:office="office" xmlns:style="style" xmlns:fo="fo" xmlns:text="text"
            office:mimetype="application/vnd.oasis.opendocument.text">
            <office:automatic-styles>
              <style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in" fo:margin="1in"/></style:page-layout>
            </office:automatic-styles>
            <office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles>
            <office:body><office:text><text:p>Flat text</text:p></office:text></office:body>
          </office:document>"#;

        let document = detect_and_parse(bytes, Limits::default())
            .expect("valid flat ODT")
            .expect("flat ODT is recognized by content");
        assert_eq!(document.format, Some(crate::model::DocumentFormat::Odt));
        assert!(
            document
                .objects
                .iter()
                .any(|object| object.text.as_deref() == Some("Flat text"))
        );
    }

    #[cfg(feature = "odf-formats")]
    #[test]
    fn parses_flat_open_document_graphics_through_the_public_dispatch() {
        let bytes = br#"<?xml version="1.0"?>
          <office:document xmlns:office="office" xmlns:style="style" xmlns:fo="fo" xmlns:draw="draw"
            office:mimetype="application/vnd.oasis.opendocument.graphics">
            <office:automatic-styles>
              <style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in"/></style:page-layout>
            </office:automatic-styles>
            <office:master-styles><style:master-page style:name="Default" style:page-layout-name="pm1"/></office:master-styles>
            <office:body><office:drawing><draw:page draw:name="Page 1" draw:master-page-name="Default"/></office:drawing></office:body>
          </office:document>"#;

        let document = detect_and_parse(bytes, Limits::default())
            .expect("valid flat ODG")
            .expect("flat ODG is recognized by content");
        assert_eq!(document.format, Some(crate::model::DocumentFormat::Odp));
        assert_eq!(document.units.len(), 1);
    }

    #[cfg(feature = "odf-formats")]
    #[test]
    fn keeps_zero_width_text_frames_in_flat_open_document_text() {
        let bytes = br#"<office:document xmlns:office="office" xmlns:text="text" xmlns:draw="draw" xmlns:svg="svg"
          office:mimetype="application/vnd.oasis.opendocument.text">
          <office:body><office:text><text:p><draw:frame svg:width="0cm" svg:height="2cm">
            <draw:text-box><text:p>.</text:p></draw:text-box>
          </draw:frame></text:p></office:text></office:body>
        </office:document>"#;

        let document = detect_and_parse(bytes, Limits::default())
            .expect("valid flat ODT")
            .expect("flat ODT is recognized by content");
        let frame = document
            .objects
            .iter()
            .find(|object| object.text.as_deref() == Some("."))
            .expect("zero-width text frame is retained");
        assert_eq!(frame.bounds.width, 1.0);
        assert!(document.diagnostics.iter().any(|diagnostic| {
            diagnostic.message.contains("zero dimension")
                && diagnostic.fidelity == crate::diagnostic::Fidelity::Approximate
        }));
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn parses_presentation_templates_and_slideshows_through_the_public_dispatch() {
        for (content_type, macro_enabled) in [
            (PPSX_MAIN, false),
            (POTX_MAIN, false),
            (PPSM_MAIN, true),
            (POTM_MAIN, true),
        ] {
            let bytes = pptx_fixture_with_main_content_type(content_type);
            let document = detect_and_parse(&bytes, Limits::default())
                .expect("valid OOXML presentation variant")
                .expect("presentation variant is recognized by package content");

            assert_eq!(
                document.format,
                Some(DocumentFormat::Pptx),
                "{content_type}"
            );
            assert_eq!(
                document.kind,
                Some(DocumentKind::Presentation),
                "{content_type}"
            );
            assert_eq!(
                document.diagnostics.iter().any(|diagnostic| diagnostic.code
                    == crate::diagnostic::DiagnosticCode::ActiveContentBlocked),
                macro_enabled,
                "{content_type}"
            );
        }
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn parses_a_pptx_text_shape_with_source_mapping() {
        let bytes = pptx_fixture(&[]);
        let document = detect_and_parse(&bytes, Limits::default())
            .expect("valid package")
            .expect("PPTX is recognized by package content");

        assert!(!document.fatal);
        assert_eq!(document.format, Some(DocumentFormat::Pptx));
        assert_eq!(document.kind, Some(DocumentKind::Presentation));
        assert_eq!(document.units.len(), 1);
        let unit = &document.units[0];
        assert_eq!(unit.kind, UnitKind::Slide);
        assert_eq!(unit.id, "unit:0");
        assert_eq!((unit.width, unit.height), (960.0, 720.0));

        assert_eq!(document.objects.len(), 1);
        let object = &document.objects[0];
        assert_eq!(object.kind, ObjectKind::TextBox);
        assert_eq!(object.stable_id, "object:0");
        assert_eq!(object.text.as_deref(), Some("Hello local document"));
        assert_eq!(
            (
                object.bounds.x,
                object.bounds.y,
                object.bounds.width,
                object.bounds.height
            ),
            (96.0, 96.0, 384.0, 96.0)
        );
        assert_eq!(object.source.part, "ppt/slides/slide1.xml");
        assert_eq!(object.source.mapping, MappingQuality::Exact);
        assert_eq!(
            object.source.locator,
            SourceLocator::PptxShape {
                shape_id: 2,
                row: None,
                column: None,
                text_range: Some((0, 20)),
                metadata: crate::model::PptxObjectMetadata {
                    name: Some("Greeting".to_owned()),
                    ..crate::model::PptxObjectMetadata::default()
                },
            }
        );
        assert!(matches!(
            &object.visual,
            Visual::TextLayout { visual, .. }
                if matches!(visual.as_ref(), Visual::RichText {
                    geometry: Geometry::Rectangle,
                    align: TextAlign::Center,
                    runs,
                    ..
                } if runs.len() == 1
                    && runs[0].font_size == 32.0
                    && runs[0].bold
                    && !runs[0].italic)
        ));
        assert!(document.diagnostics.is_empty());
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn pptx_images_use_default_and_override_content_types_for_custom_extensions() {
        let cases = [
            (
                "default",
                r#"<Default Extension="asset" ContentType="image/svg+xml"/>"#,
                "image.asset",
            ),
            (
                "override",
                r#"<Default Extension="bin" ContentType="image/png"/><Override PartName="/ppt/media/image.bin" ContentType="image/svg+xml"/>"#,
                "image.bin",
            ),
        ];

        for (label, image_content_types, image_name) in cases {
            let bytes = pptx_image_fixture(image_content_types, image_name, b"<svg/>");
            let document = detect_and_parse(&bytes, Limits::default())
                .expect("valid package")
                .expect("PPTX is recognized");

            assert_eq!(document.objects.len(), 1, "{label}");
            assert!(
                matches!(
                    &document.objects[0].visual,
                    Visual::Image { media_type, bytes, .. }
                        if media_type == "image/svg+xml" && bytes == b"<svg/>"
                ),
                "{label}"
            );
            assert!(document.diagnostics.is_empty(), "{label}");
        }
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn pptx_images_recover_supported_content_type_and_signature_mismatches() {
        let bytes = pptx_image_fixture(
            r#"<Default Extension="svg" ContentType="image/png"/>"#,
            "image.svg",
            b"\xff\xd8\xffjpeg",
        );
        let document = detect_and_parse(&bytes, Limits::default())
            .expect("supported image mismatch is recoverable")
            .expect("PPTX is recognized");

        assert!(matches!(
            &document.objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/jpeg" && bytes.starts_with(b"\xff\xd8\xff")
        ));
        assert!(document.diagnostics.is_empty());
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn pptx_images_still_reject_unrecognized_mismatched_bytes() {
        let bytes = pptx_image_fixture(
            r#"<Default Extension="png" ContentType="image/png"/>"#,
            "image.png",
            b"not an image",
        );
        let document = detect_and_parse(&bytes, Limits::default())
            .expect("invalid image is non-fatal")
            .expect("PPTX is recognized");

        assert!(document.objects.is_empty());
        assert_eq!(document.diagnostics.len(), 1);
        assert_eq!(
            document.diagnostics[0].fidelity,
            crate::diagnostic::Fidelity::Omitted
        );
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn xlsx_images_use_content_types_for_custom_extensions() {
        let bytes = stored_zip(&[
            (
                "[Content_Types].xml",
                r#"<Types><Default Extension="tmp" ContentType="image/svg+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<workbook><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet><dimension ref="A1:B2"/><drawing r:id="rIdDrawing"/><sheetData/></worksheet>"#,
            ),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships><Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#,
            ),
            (
                "xl/drawings/drawing1.xml",
                r#"<xdr:wsDr xmlns:xdr="xdr" xmlns:a="a" xmlns:r="r"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor></xdr:wsDr>"#,
            ),
            (
                "xl/drawings/_rels/drawing1.xml.rels",
                r#"<Relationships><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image.tmp"/></Relationships>"#,
            ),
            (
                "xl/media/image.tmp",
                r#"<svg xmlns="http://www.w3.org/2000/svg"/>"#,
            ),
        ]);

        let document = detect_and_parse(&bytes, Limits::default())
            .expect("valid XLSX image package")
            .expect("XLSX is recognized");

        assert!(matches!(
            &document.objects[0].visual,
            Visual::Image { media_type, bytes, .. }
                if media_type == "image/svg+xml" && bytes.starts_with(b"<svg")
        ));
        assert!(document.diagnostics.is_empty());
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn rejects_xlsx_shared_string_materialization_above_the_document_budget() {
        let shared_text = "x".repeat(1_024);
        let cells = (1..=9)
            .map(|row| format!(r#"<c r="A{row}" t="s"><v>0</v></c>"#))
            .collect::<String>();
        let shared_strings = format!("<sst><si><t>{shared_text}</t></si></sst>");
        let worksheet = format!(
            r#"<worksheet><dimension ref="A1:A9"/><sheetData>{cells}</sheetData></worksheet>"#
        );
        let entries = [
            (
                "[Content_Types].xml",
                r#"<Types><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<workbook xmlns:r="r"><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>"#,
            ),
            ("xl/sharedStrings.xml", shared_strings.as_str()),
            ("xl/worksheets/sheet1.xml", worksheet.as_str()),
        ];
        let bytes = stored_zip(&entries);
        let limits = Limits {
            max_entry_uncompressed_bytes: 8 * 1_024,
            max_total_uncompressed_bytes: 8 * 1_024,
            max_xml_bytes: 8 * 1_024,
            ..Limits::default()
        };

        let error = detect_and_parse(&bytes, limits)
            .expect_err("shared-string copies must be rejected before they exhaust memory");

        assert_eq!(
            error.code,
            crate::diagnostic::DiagnosticCode::ZipTotalSizeLimit
        );
    }

    #[cfg(feature = "odf-formats")]
    #[test]
    fn rejects_ods_repeated_text_materialization_above_the_document_budget() {
        let cell_text = "x".repeat(1_024);
        let content = format!(
            r#"<office:document-content><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:number-columns-repeated="9" office:value-type="string" office:string-value="{cell_text}"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
        );
        let entries = [
            ("mimetype", "application/vnd.oasis.opendocument.spreadsheet"),
            ("content.xml", content.as_str()),
        ];
        let bytes = stored_zip(&entries);
        let limits = Limits {
            max_entry_uncompressed_bytes: 8 * 1_024,
            max_total_uncompressed_bytes: 8 * 1_024,
            max_xml_bytes: 8 * 1_024,
            ..Limits::default()
        };

        let error = detect_and_parse(&bytes, limits)
            .expect_err("repeated ODS text must be rejected before object expansion");

        assert_eq!(
            error.code,
            crate::diagnostic::DiagnosticCode::ZipTotalSizeLimit
        );
    }

    #[cfg(feature = "odf-formats")]
    #[test]
    fn allows_large_sparse_ods_sheets_for_viewport_rendering() {
        let content = r#"<office:document-content><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row table:number-rows-repeated="100000"/><table:table-row><table:table-cell office:value-type="string" office:string-value="x"/></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#;
        let entries = [
            ("mimetype", "application/vnd.oasis.opendocument.spreadsheet"),
            ("content.xml", content),
        ];
        let bytes = stored_zip(&entries);

        let document = detect_and_parse(&bytes, Limits::default())
            .expect("large sparse sheets are rendered through bounded viewports")
            .expect("ODS is recognized");

        assert_eq!(document.units[0].rows, 100_001);
        assert_eq!(document.objects.len(), 1);
    }

    #[cfg(feature = "odf-formats")]
    #[test]
    fn rejects_odt_repeated_text_materialization_above_the_document_budget() {
        let cell_text = "x".repeat(1_024);
        let content = format!(
            r#"<office:document-content><office:body><office:text><table:table><table:table-row><table:table-cell table:number-columns-repeated="9"><text:p>{cell_text}</text:p></table:table-cell></table:table-row></table:table></office:text></office:body></office:document-content>"#
        );
        let styles = r#"<office:document-styles><office:styles><style:page-layout style:name="page"><style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in" fo:margin="1in"/></style:page-layout></office:styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="page"/></office:master-styles></office:document-styles>"#;
        let entries = [
            ("mimetype", "application/vnd.oasis.opendocument.text"),
            ("styles.xml", styles),
            ("content.xml", content.as_str()),
        ];
        let bytes = stored_zip(&entries);
        let limits = Limits {
            max_entry_uncompressed_bytes: 8 * 1_024,
            max_total_uncompressed_bytes: 8 * 1_024,
            max_xml_bytes: 8 * 1_024,
            ..Limits::default()
        };

        let error = detect_and_parse(&bytes, limits)
            .expect_err("repeated ODT text must be rejected before table expansion");

        assert_eq!(
            error.code,
            crate::diagnostic::DiagnosticCode::ZipTotalSizeLimit
        );
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn pptx_initial_parse_keeps_first_and_table_slides_then_loads_other_units() {
        let presentation = r#"<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/><p:sldId id="258" r:id="rId3"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#;
        let relationships = r#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/></Relationships>"#;
        let table_slide = r#"<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></p:xfrm><a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="952500"/></a:tblGrid><a:tr h="952500"><a:tc><a:txBody><a:p/></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#;
        let deferred_slide = r#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="3"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm></p:spPr></p:pic></p:spTree></p:cSld></p:sld>"#;
        let bytes = pptx_fixture_with_presentation(
            presentation,
            relationships,
            &[
                ("ppt/slides/slide2.xml", table_slide),
                ("ppt/slides/slide3.xml", deferred_slide),
                (
                    "ppt/slides/_rels/slide3.xml.rels",
                    r#"<Relationships><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image.svg"/></Relationships>"#,
                ),
                ("ppt/media/image.svg", "<svg/>"),
            ],
        );
        let package = crate::package::Package::open(&bytes, Limits::default()).unwrap();
        let content_types = super::parse_content_types(&package).unwrap();
        let initial = super::pptx::parse_initial(
            &package,
            "ppt/presentation.xml",
            Vec::new(),
            &content_types,
            &crate::font_metrics::FontMetricTable::default(),
        )
        .unwrap();

        assert_eq!(
            initial.loaded_units,
            std::collections::BTreeSet::from([0, 1])
        );
        assert!(
            initial
                .document
                .objects
                .iter()
                .all(|object| object.unit_index != 2)
        );

        let (deferred, materialized_image_bytes) = super::pptx::parse_unit(
            &package,
            "ppt/presentation.xml",
            &content_types,
            &crate::font_metrics::FontMetricTable::default(),
            2,
            initial.materialized_image_bytes,
        )
        .unwrap();
        assert!(
            !deferred.objects.is_empty()
                && deferred.objects.iter().all(|object| object.unit_index == 2)
        );
        assert!(materialized_image_bytes > initial.materialized_image_bytes);
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn rejects_pptx_slides_that_reference_the_same_part() {
        let presentation = r#"<p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#;
        let relationships = r#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#;
        let bytes = pptx_fixture_with_presentation(presentation, relationships, &[]);

        let error = detect_and_parse(&bytes, Limits::default())
            .expect_err("a slide part must not be parsed more than once");

        assert_eq!(error.code, crate::diagnostic::DiagnosticCode::FormatInvalid);
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn rejects_xlsx_worksheets_that_reference_the_same_part() {
        let entries = [
            (
                "[Content_Types].xml",
                r#"<Types><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<workbook><sheets><sheet name="Sheet1" r:id="rId1"/><sheet name="Sheet2" r:id="rId2"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet><dimension ref="A1"/><sheetData/></worksheet>"#,
            ),
        ];
        let bytes = stored_zip(&entries);

        let error = detect_and_parse(&bytes, Limits::default())
            .expect_err("a worksheet part must not be parsed more than once");

        assert_eq!(error.code, crate::diagnostic::DiagnosticCode::FormatInvalid);
    }

    #[cfg(feature = "native-formats")]
    fn pptx_fixture(extra_entries: &[(&str, &str)]) -> Vec<u8> {
        pptx_fixture_with_presentation(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#,
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#,
            extra_entries,
        )
    }

    #[cfg(feature = "native-formats")]
    fn pptx_fixture_with_main_content_type(content_type: &str) -> Vec<u8> {
        pptx_fixture_with_presentation_and_content_type(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"#,
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#,
            content_type,
            &[],
        )
    }

    #[cfg(feature = "native-formats")]
    fn pptx_image_fixture(
        image_content_types: &str,
        image_name: &str,
        image_bytes: &[u8],
    ) -> Vec<u8> {
        let content_types = format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  {image_content_types}
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>"#
        );
        let slide = r#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>
  <p:pic><p:nvPicPr><p:cNvPr id="7"/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9525" cy="9525"/></a:xfrm></p:spPr></p:pic>
</p:spTree></p:cSld></p:sld>"#;
        let slide_relationships = format!(
            r#"<Relationships><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{image_name}"/></Relationships>"#
        );
        let image_part = format!("ppt/media/{image_name}");
        stored_zip_bytes(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            (
                "_rels/.rels",
                br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
            ),
            (
                "ppt/presentation.xml",
                br#"<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#,
            ),
            (
                "ppt/_rels/presentation.xml.rels",
                br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
            ),
            ("ppt/slides/slide1.xml", slide.as_bytes()),
            (
                "ppt/slides/_rels/slide1.xml.rels",
                slide_relationships.as_bytes(),
            ),
            (image_part.as_str(), image_bytes),
        ])
    }

    #[cfg(feature = "native-formats")]
    fn pptx_fixture_with_presentation(
        presentation: &str,
        presentation_relationships: &str,
        extra_entries: &[(&str, &str)],
    ) -> Vec<u8> {
        pptx_fixture_with_presentation_and_content_type(
            presentation,
            presentation_relationships,
            PPTX_MAIN,
            extra_entries,
        )
    }

    #[cfg(feature = "native-formats")]
    #[test]
    fn opc_recovery_preserves_strict_callers_and_resource_limits() {
        let bytes = stored_zip(&[
            (
                "[Content_Types].xml",
                "<Types><Default Extension='xml' ContentType='first'/><Default Extension='xml' ContentType='second'/><Override PartName='/ignored'/></Types>",
            ),
            (
                "_rels/.rels",
                "<Relationships><Relationship Type='metadata' Target='meta.xml'/><Relationship Id='r1' Type='x/fixedrepresentation' Target='main.fdseq'/></Relationships>",
            ),
        ]);
        let package = crate::package::Package::open(&bytes, Limits::default()).unwrap();
        assert!(super::parse_content_types(&package).is_err());
        let mut diagnostics = Vec::new();
        let types =
            super::parse_content_types_with_diagnostics(&package, Some(&mut diagnostics)).unwrap();
        assert_eq!(types.for_part("a.xml"), Some("first"));
        assert_eq!(diagnostics.len(), 2);
        assert!(package.relationships(None).is_err());
        let edges = package
            .relationships_filtered(None, Some("/fixedrepresentation"))
            .unwrap();
        assert_eq!(edges.len(), 1);
        assert_eq!(edges[0].target, "main.fdseq");
        let package = crate::package::Package::open(
            &bytes,
            Limits {
                max_relationship_edges: 1,
                ..Limits::default()
            },
        )
        .unwrap();
        assert_eq!(
            package
                .relationships_filtered(None, Some("/fixedrepresentation"))
                .unwrap_err()
                .code,
            crate::diagnostic::DiagnosticCode::RelationshipLimit
        );
    }

    #[cfg(feature = "native-formats")]
    fn pptx_fixture_with_presentation_and_content_type(
        presentation: &str,
        presentation_relationships: &str,
        main_content_type: &str,
        extra_entries: &[(&str, &str)],
    ) -> Vec<u8> {
        let content_types = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/ppt/presentation.xml" ContentType="{main_content_type}"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>"#
        );
        let mut entries = vec![
            ("[Content_Types].xml", content_types.as_str()),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#,
            ),
            ("ppt/presentation.xml", presentation),
            (
                "ppt/_rels/presentation.xml.rels",
                presentation_relationships,
            ),
            (
                "ppt/slides/slide1.xml",
                r#"<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree><p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Greeting"/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
      <a:prstGeom prst="rect"/>
    </p:spPr>
    <p:txBody><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2400" b="1" i="0"/><a:t>Hello local document</a:t></a:r></a:p></p:txBody>
  </p:sp></p:spTree></p:cSld>
</p:sld>"#,
            ),
        ];
        entries.extend_from_slice(extra_entries);
        stored_zip(&entries)
    }

    fn stored_zip(entries: &[(&str, &str)]) -> Vec<u8> {
        let entries = entries
            .iter()
            .map(|(name, content)| (*name, content.as_bytes()))
            .collect::<Vec<_>>();
        stored_zip_bytes(&entries)
    }

    fn stored_zip_bytes(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut central = Vec::new();
        for (name, content) in entries {
            let name = name.as_bytes();
            let crc = crate::zip::crc32(content);
            let offset = bytes.len() as u32;
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

    fn push_u16(bytes: &mut Vec<u8>, value: u16) {
        bytes.extend_from_slice(&value.to_le_bytes());
    }

    fn push_u32(bytes: &mut Vec<u8>, value: u32) {
        bytes.extend_from_slice(&value.to_le_bytes());
    }
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
fn ellipse_arc_bezier_points(
    center: (f32, f32),
    radii: (f32, f32),
    from: f32,
    to: f32,
) -> [(f32, f32); 3] {
    let alpha = 4.0 / 3.0 * ((to - from) / 4.0).tan();
    let from_point = (
        center.0 + radii.0 * from.cos(),
        center.1 + radii.1 * from.sin(),
    );
    let to_point = (center.0 + radii.0 * to.cos(), center.1 + radii.1 * to.sin());
    [
        (
            from_point.0 - alpha * radii.0 * from.sin(),
            from_point.1 + alpha * radii.1 * from.cos(),
        ),
        (
            to_point.0 + alpha * radii.0 * to.sin(),
            to_point.1 - alpha * radii.1 * to.cos(),
        ),
        to_point,
    ]
}

#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats"
))]
fn append_elliptical_arc(
    commands: &mut Vec<crate::model::PathCommand>,
    bounding_box: [f32; 4],
    radial_points: [f32; 4],
    clockwise: bool,
    move_to_start: bool,
) {
    use crate::model::PathCommand;
    let [x1, y1, x2, y2] = bounding_box;
    let [start_x, start_y, end_x, end_y] = radial_points;
    let center_x = (x1 + x2) / 2.0;
    let center_y = (y1 + y2) / 2.0;
    let radius_x = (x2 - x1).abs() / 2.0;
    let radius_y = (y2 - y1).abs() / 2.0;
    if radius_x <= f32::EPSILON || radius_y <= f32::EPSILON {
        return;
    }

    let angle = |x: f32, y: f32| ((y - center_y) / radius_y).atan2((x - center_x) / radius_x);
    let start_angle = angle(start_x, start_y);
    let end_angle = angle(end_x, end_y);
    let mut sweep = end_angle - start_angle;
    if clockwise {
        while sweep <= 0.0 {
            sweep += std::f32::consts::TAU;
        }
    } else {
        while sweep >= 0.0 {
            sweep -= std::f32::consts::TAU;
        }
    }
    let point = |angle: f32| {
        (
            center_x + radius_x * angle.cos(),
            center_y + radius_y * angle.sin(),
        )
    };
    let (arc_start_x, arc_start_y) = point(start_angle);
    commands.push(if move_to_start {
        PathCommand::MoveTo {
            x: arc_start_x,
            y: arc_start_y,
        }
    } else {
        PathCommand::LineTo {
            x: arc_start_x,
            y: arc_start_y,
        }
    });

    let segment_count = (sweep.abs() / std::f32::consts::FRAC_PI_2).ceil().max(1.0) as usize;
    let segment_sweep = sweep / segment_count as f32;
    for segment in 0..segment_count {
        let from = start_angle + segment_sweep * segment as f32;
        let to = from + segment_sweep;
        let [control_1, control_2, (to_x, to_y)] =
            ellipse_arc_bezier_points((center_x, center_y), (radius_x, radius_y), from, to);
        commands.push(PathCommand::BezierCurveTo {
            cp1x: control_1.0,
            cp1y: control_1.1,
            cp2x: control_2.0,
            cp2y: control_2.1,
            x: to_x,
            y: to_y,
        });
    }
}

// Ordinary least squares shared by OOXML and ODF; chart adapters own domains and clipping.
#[cfg(any(feature = "native-formats", feature = "odf-formats"))]
pub(super) fn linear_regression(x_values: &[f32], y_values: &[f32]) -> Option<(f32, f32, f32)> {
    let points = x_values
        .iter()
        .copied()
        .zip(y_values.iter().copied())
        .filter(|(x, y)| x.is_finite() && y.is_finite())
        .collect::<Vec<_>>();
    if points.len() < 2 {
        return None;
    }
    let count = points.len() as f32;
    let mean_x = points.iter().map(|(x, _)| x).sum::<f32>() / count;
    let mean_y = points.iter().map(|(_, y)| y).sum::<f32>() / count;
    let variance_x = points
        .iter()
        .map(|(x, _)| (x - mean_x).powi(2))
        .sum::<f32>();
    if variance_x <= f32::EPSILON {
        return None;
    }
    let slope = points
        .iter()
        .map(|(x, y)| (x - mean_x) * (y - mean_y))
        .sum::<f32>()
        / variance_x;
    let intercept = mean_y - slope * mean_x;
    let residual = points
        .iter()
        .map(|(x, y)| (y - (slope * x + intercept)).powi(2))
        .sum::<f32>();
    let total = points
        .iter()
        .map(|(_, y)| (y - mean_y).powi(2))
        .sum::<f32>();
    Some((
        slope,
        intercept,
        if total <= f32::EPSILON {
            1.0
        } else {
            1.0 - residual / total
        },
    ))
}

#[cfg(all(test, any(feature = "native-formats", feature = "odf-formats")))]
#[test]
fn chart_linear_regression_preserves_finite_pairs_and_degenerate_cases() {
    let (slope, intercept, r_squared) =
        linear_regression(&[0.0, 1.0, 2.0, f32::NAN], &[1.0, 3.0, 5.0, 9.0]).unwrap();
    assert_eq!((slope, intercept, r_squared), (2.0, 1.0, 1.0));
    assert!(linear_regression(&[1.0, 1.0], &[2.0, 3.0]).is_none());
    assert!(linear_regression(&[1.0], &[2.0]).is_none());
}

pub(super) fn nice_chart_step(raw_step: f32) -> f32 {
    let raw_step = raw_step.max(f32::MIN_POSITIVE);
    let magnitude = 10.0_f32.powf(raw_step.log10().floor());
    let normalized = raw_step / magnitude;
    if normalized < 1.5 {
        magnitude
    } else if normalized < 3.0 {
        2.0 * magnitude
    } else if normalized < 7.0 {
        5.0 * magnitude
    } else {
        10.0 * magnitude
    }
}

pub(super) fn format_general_number(number: f64, width: usize) -> String {
    if number == 0.0 {
        return "0".to_owned();
    }
    let width = width.clamp(1, 11);
    let exponent = number.abs().log10().floor() as i32;
    let significant_decimals = 10 - exponent;
    let rounded = if significant_decimals < 0 {
        let factor = 10_f64.powi(-significant_decimals);
        (number / factor).round() * factor
    } else {
        number
    };
    let integer_digits = exponent.max(0) as usize + 1;
    let width_decimals =
        width.saturating_sub(usize::from(number.is_sign_negative()) + integer_digits + 1);
    let decimals = significant_decimals.max(0) as usize;
    let mut rendered = format!(
        "{rounded:.decimals$}",
        decimals = decimals.min(width_decimals)
    );
    if rendered.contains('.') {
        rendered.truncate(rendered.trim_end_matches('0').trim_end_matches('.').len());
    }
    if rendered.len() <= width {
        return rendered;
    }
    for decimals in (0..10).rev() {
        let scientific = format!("{number:.decimals$e}");
        let Some((mantissa, exponent)) = scientific.split_once('e') else {
            continue;
        };
        let exponent = exponent.parse::<i32>().unwrap_or_default();
        let mantissa = mantissa.trim_end_matches('0').trim_end_matches('.');
        let rendered = format!("{mantissa}E{exponent:+03}");
        if rendered.len() <= width {
            return rendered;
        }
    }
    "#".repeat(width)
}
