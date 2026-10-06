//! Shared VML coordinate and paint semantics. Host anchoring stays in each adapter.
use super::optional_xml_attribute as optional_attribute;
use crate::diagnostic::{Diagnostic, DiagnosticCode, Phase};
use crate::model::{FillRule, Geometry, ImageAdjustment, PathCommand};
use crate::xml::XmlAttribute;
const POINTS_TO_CSS_PIXELS: f32 = 96.0 / 72.0;
fn format_error(part: &str, message: impl Into<String>) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Parse, None, message).in_part(part)
}

pub(super) fn parse_vml_length(value: &str, part: &str) -> Result<f32, Diagnostic> {
    const MAX_VML_LENGTH: f32 = 1_000_000.0;
    let lower = value.trim().to_ascii_lowercase();
    let (number, scale) = [
        ("pt", POINTS_TO_CSS_PIXELS),
        ("px", 1.0),
        ("in", 96.0),
        ("cm", 96.0 / 2.54),
        ("mm", 96.0 / 25.4),
        ("pc", 16.0),
    ]
    .into_iter()
    .find_map(|(unit, scale)| lower.strip_suffix(unit).map(|number| (number, scale)))
    .unwrap_or((lower.as_str(), 1.0));
    let length = number
        .trim()
        .parse::<f32>()
        .map_err(|_| format_error(part, format!("VML length {value:?} is invalid")))?
        * scale;
    if !length.is_finite() || length.abs() > MAX_VML_LENGTH {
        return Err(format_error(part, "VML length exceeds the supported range"));
    }
    Ok(length)
}

pub(super) fn parse_vml_path_geometry(
    path: &str,
    coordsize: Option<&str>,
    coordorigin: Option<&str>,
    width: f32,
    height: f32,
) -> Option<Geometry> {
    let (coord_width, coord_height) = coordsize
        .and_then(|value| value.split_once(','))
        .and_then(|(width, height)| Some((width.trim().parse().ok()?, height.trim().parse().ok()?)))
        .unwrap_or((21_600.0_f32, 21_600.0_f32));
    if !coord_width.is_finite()
        || !coord_height.is_finite()
        || coord_width <= 0.0
        || coord_height <= 0.0
    {
        return None;
    }
    let (origin_x, origin_y) = match coordorigin {
        Some(value) => {
            let (x, y) = value.split_once(',')?;
            (x.trim().parse::<f32>().ok()?, y.trim().parse::<f32>().ok()?)
        }
        None => (0.0, 0.0),
    };
    if !origin_x.is_finite() || !origin_y.is_finite() {
        return None;
    }
    let scale = |x: f32, y: f32| {
        (
            (x - origin_x) / coord_width * width,
            (y - origin_y) / coord_height * height,
        )
    };
    let mut commands = Vec::new();
    let mut current = (0.0_f32, 0.0_f32);
    let mut offset = 0;
    let bytes = path.as_bytes();
    while offset < bytes.len() {
        while offset < bytes.len() && bytes[offset].is_ascii_whitespace() {
            offset += 1;
        }
        let command = *bytes.get(offset)?;
        if !command.is_ascii_alphabetic() {
            return None;
        }
        offset += 1;
        let values_start = offset;
        while offset < bytes.len() && !bytes[offset].is_ascii_alphabetic() {
            offset += 1;
        }
        let values = &path[values_start..offset];
        match command.to_ascii_lowercase() {
            b'm' | b'l' | b'r' | b'c' | b'v' => {
                let values = values
                    .split(',')
                    .map(|value| {
                        let value = value.trim();
                        if value.is_empty() {
                            Some(0.0)
                        } else {
                            value.parse::<f32>().ok().filter(|value| value.is_finite())
                        }
                    })
                    .collect::<Option<Vec<_>>>()?;
                let cubic = matches!(command.to_ascii_lowercase(), b'c' | b'v');
                let stride = if cubic { 6 } else { 2 };
                if values.len() < stride || values.len() % stride != 0 {
                    return None;
                }
                if cubic {
                    for curve in values.chunks_exact(6) {
                        let base = if command.eq_ignore_ascii_case(&b'v') {
                            current
                        } else {
                            (0.0, 0.0)
                        };
                        let (cp1x, cp1y) = scale(base.0 + curve[0], base.1 + curve[1]);
                        let (cp2x, cp2y) = scale(base.0 + curve[2], base.1 + curve[3]);
                        current = (base.0 + curve[4], base.1 + curve[5]);
                        let (x, y) = scale(current.0, current.1);
                        if [cp1x, cp1y, cp2x, cp2y, x, y]
                            .iter()
                            .any(|v| !v.is_finite())
                        {
                            return None;
                        }
                        commands.push(PathCommand::BezierCurveTo {
                            cp1x,
                            cp1y,
                            cp2x,
                            cp2y,
                            x,
                            y,
                        });
                    }
                    continue;
                }
                for (index, pair) in values.chunks_exact(2).enumerate() {
                    let point = if command.eq_ignore_ascii_case(&b'r') {
                        (current.0 + pair[0], current.1 + pair[1])
                    } else {
                        (pair[0], pair[1])
                    };
                    current = point;
                    let (x, y) = scale(point.0, point.1);
                    commands.push(if command.eq_ignore_ascii_case(&b'm') && index == 0 {
                        PathCommand::MoveTo { x, y }
                    } else {
                        PathCommand::LineTo { x, y }
                    });
                }
            }
            b'x' => commands.push(PathCommand::ClosePath),
            b'e' => {}
            _ => return None,
        }
    }
    commands
        .iter()
        .any(|command| matches!(command, PathCommand::MoveTo { .. }))
        .then_some(Geometry::Path {
            fill_rule: FillRule::NonZero,
            commands,
        })
}

pub(super) fn parse_vml_color(value: &str, part: &str) -> Result<u32, Diagnostic> {
    let value = value.trim();
    let value = value
        .rfind('[')
        .and_then(|open| {
            let index = value[open + 1..].strip_suffix(']')?;
            (!index.is_empty() && index.bytes().all(|byte| byte.is_ascii_digit()))
                .then(|| value[..open].trim_end())
        })
        .unwrap_or(value);
    let lower = value.to_ascii_lowercase();
    let rgb = match lower.as_str() {
        "aqua" => Some(0x00ffff),
        "black" => Some(0x000000),
        "blue" => Some(0x0000ff),
        "fuchsia" => Some(0xff00ff),
        "gray" => Some(0x808080),
        "green" => Some(0x008000),
        "lime" => Some(0x00ff00),
        "maroon" => Some(0x800000),
        "navy" => Some(0x000080),
        "olive" => Some(0x808000),
        "purple" => Some(0x800080),
        "red" => Some(0xff0000),
        "silver" => Some(0xc0c0c0),
        "teal" => Some(0x008080),
        "white" | "window" => Some(0xffffff),
        "yellow" => Some(0xffff00),
        _ => None,
    };
    if let Some(rgb) = rgb {
        return Ok((rgb << 8) | 0xff);
    }
    if let Some(components) = lower
        .strip_prefix("rgb(")
        .and_then(|value| value.strip_suffix(')'))
    {
        let mut components = components.split(',').map(str::trim);
        let red = components.next().and_then(|value| value.parse::<u8>().ok());
        let green = components.next().and_then(|value| value.parse::<u8>().ok());
        let blue = components.next().and_then(|value| value.parse::<u8>().ok());
        if let (Some(red), Some(green), Some(blue), None) = (red, green, blue, components.next()) {
            return Ok((u32::from(red) << 24)
                | (u32::from(green) << 16)
                | (u32::from(blue) << 8)
                | 0xff);
        }
        return Err(format_error(part, "VML RGB triplet is invalid"));
    }
    let hex = value.strip_prefix('#').unwrap_or(value);
    let expanded = if hex.len() == 3 {
        hex.chars().flat_map(|value| [value, value]).collect()
    } else {
        hex.to_owned()
    };
    if expanded.len() != 6 {
        return Err(format_error(part, "VML color must contain six digits"));
    }
    u32::from_str_radix(&expanded, 16)
        .map(|rgb| (rgb << 8) | 0xff)
        .map_err(|_| format_error(part, "VML color is not hexadecimal RGB"))
}

pub(super) fn parse_vml_scalar(value: &str, part: &str) -> Result<f32, Diagnostic> {
    let value = value.trim();
    let scalar = if let Some(fixed) = value.strip_suffix('f') {
        fixed.parse::<f32>().map(|value| value / 65_536.0)
    } else {
        value.parse::<f32>()
    }
    .map_err(|_| format_error(part, format!("VML scalar {value:?} is invalid")))?;
    if !scalar.is_finite() || scalar.abs() > 32_766.0 {
        return Err(format_error(part, "VML scalar exceeds the supported range"));
    }
    Ok(scalar)
}

// Office VML stores brightness at half the shared normalized value and
// contrast as a gain. Keep the VML encoding conversion outside the renderer.
pub(super) fn parse_vml_image_adjustment(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<ImageAdjustment, Diagnostic> {
    let gain = optional_attribute(attributes, "gain", part)?
        .map(|value| parse_vml_scalar(&value, part))
        .transpose()?
        .unwrap_or(1.0);
    let black = optional_attribute(attributes, "blacklevel", part)?
        .map(|value| parse_vml_scalar(&value, part))
        .transpose()?
        .unwrap_or(0.0);
    if gain < 0.0 || !(-1.0..=1.0).contains(&black) {
        return Err(format_error(part, "unsupported VML gain or blacklevel"));
    }
    Ok(ImageAdjustment {
        contrast: if gain <= 1.0 {
            gain - 1.0
        } else {
            1.0 - 1.0 / gain
        },
        brightness: (2.0 * black).clamp(-1.0, 1.0),
        ..ImageAdjustment::default()
    })
}

pub(super) fn parse_vml_opacity(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<f32, Diagnostic> {
    optional_attribute(attributes, "opacity", part)?
        .map(|value| {
            let value = value.trim();
            if let Some(percent) = value.strip_suffix('%') {
                parse_vml_scalar(percent, part).map(|value| value / 100.0)
            } else {
                parse_vml_scalar(value, part)
            }
        })
        .transpose()
        .map(|value| value.unwrap_or(1.0).clamp(0.0, 1.0))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn image_adjustment_decodes_fixed_point_and_preserves_identity() {
        assert_eq!(
            parse_vml_image_adjustment(&[], "vml").unwrap(),
            ImageAdjustment::default()
        );
        for (gain, black, contrast, brightness) in [
            ("19661f", "22938f", -0.7, 0.7),
            ("0.3", "0.35", -0.7, 0.7),
            ("2", "-0.25", 0.5, -0.5),
            ("0", "1", -1.0, 1.0),
        ] {
            let adjustment = parse_vml_image_adjustment(
                &[
                    XmlAttribute {
                        name: "gain",
                        value: gain,
                    },
                    XmlAttribute {
                        name: "blacklevel",
                        value: black,
                    },
                ],
                "vml",
            )
            .unwrap();
            assert!((adjustment.contrast - contrast).abs() < 0.0001);
            assert!((adjustment.brightness - brightness).abs() < 0.0001);
            assert!(adjustment.is_valid());
        }
        for gain in ["NaN", "-1", "bad", "9999999999f"] {
            let result = parse_vml_image_adjustment(
                &[XmlAttribute {
                    name: "gain",
                    value: gain,
                }],
                "vml",
            );
            assert!(result.is_err());
        }
    }

    #[test]
    fn ink_curves_keep_relative_controls_omitted_zeros_and_all_strokes() {
        let Geometry::Path { commands, .. } = parse_vml_path_geometry(
            "m10,20v,2,3,,5,6em20,30c20,31,22,33,24,35e",
            Some("128,128"),
            Some("10,20"),
            256.0,
            128.0,
        )
        .unwrap() else {
            unreachable!()
        };
        assert_eq!(
            commands,
            vec![
                PathCommand::MoveTo { x: 0.0, y: 0.0 },
                PathCommand::BezierCurveTo {
                    cp1x: 0.0,
                    cp1y: 2.0,
                    cp2x: 6.0,
                    cp2y: 0.0,
                    x: 10.0,
                    y: 6.0
                },
                PathCommand::MoveTo { x: 20.0, y: 10.0 },
                PathCommand::BezierCurveTo {
                    cp1x: 20.0,
                    cp1y: 11.0,
                    cp2x: 24.0,
                    cp2y: 13.0,
                    x: 28.0,
                    y: 15.0
                },
            ]
        );
        for path in ["m0,0v1,2", "m0,0c1,2,3,4,5", "m0,0q1,2"] {
            assert!(parse_vml_path_geometry(path, None, None, 100.0, 100.0).is_none());
        }
        assert!(parse_vml_path_geometry("m0,0", Some("NaN,1"), None, 100.0, 100.0).is_none());
    }
}
