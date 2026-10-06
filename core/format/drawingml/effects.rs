pub(super) fn drawingml_shadow_alignment(value: Option<&str>) -> u8 {
    match value {
        Some("tl") => 0,
        Some("t") => 1,
        Some("tr") => 2,
        Some("l") => 3,
        Some("ctr") => 4,
        Some("r") => 5,
        Some("bl") => 6,
        Some("br") => 8,
        _ => 7,
    }
}

pub(super) fn drawingml_outer_shadow(
    shadow: Shadow,
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<OuterShadow, Diagnostic> {
    Ok(OuterShadow {
        shadow,
        scale_x: signed_numeric_attribute(attributes, "sx", part)?
            .map_or(1.0, |value| value as f32 / 100_000.0),
        scale_y: signed_numeric_attribute(attributes, "sy", part)?
            .map_or(1.0, |value| value as f32 / 100_000.0),
        skew_x: signed_numeric_attribute(attributes, "kx", part)?
            .map_or(0.0, |value| value as f32 / 60_000.0),
        skew_y: signed_numeric_attribute(attributes, "ky", part)?
            .map_or(0.0, |value| value as f32 / 60_000.0),
        alignment: drawingml_shadow_alignment(
            string_attribute(attributes, "algn", part)?.as_deref(),
        ),
    })
}

pub(super) fn drawingml_outer_shadow_is_identity(effect: &OuterShadow) -> bool {
    effect.scale_x == 1.0 && effect.scale_y == 1.0 && effect.skew_x == 0.0 && effect.skew_y == 0.0
}

pub(super) fn drawingml_fill_reference_has_paint(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<bool, Diagnostic> {
    Ok(!matches!(
        numeric_attribute(attributes, "idx", part)?,
        Some(0 | 1000)
    ))
}

#[derive(Clone, Debug, Default)]
pub(super) struct DrawingMlPictureEffects {
    pub(super) background_fill: Option<Paint>,
    pub(super) outer_shadow: Option<OuterShadow>,
    pub(super) inner_shadow: Option<Shadow>,
    pub(super) glow: Option<Glow>,
    pub(super) reflection: Option<Reflection>,
    pub(super) soft_edge: Option<f32>,
    pub(super) three_d: Option<ThreeDStyle>,
}

impl DrawingMlPictureEffects {
    pub(super) fn is_empty(&self) -> bool {
        self.outer_shadow.is_none()
            && self.inner_shadow.is_none()
            && self.glow.is_none()
            && self.reflection.is_none()
            && self.soft_edge.is_none()
            && self.three_d.is_none()
    }

    pub(super) fn wrap(self, visual: Visual, clip: Option<Geometry>) -> Visual {
        let visual = if clip.is_none() {
            visual
        } else {
            Visual::Effect {
                shadow: None,
                clip,
                visual: Box::new(visual),
            }
        };
        if self.is_empty() {
            visual
        } else {
            Visual::AdvancedEffect {
                outer_shadow: self.outer_shadow,
                inner_shadow: self.inner_shadow,
                glow: self.glow,
                reflection: self.reflection,
                soft_edge: self.soft_edge,
                three_d: self.three_d,
                visual: Box::new(visual),
            }
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum DrawingMlPictureEffectTarget {
    Shadow,
    InnerShadow,
    Glow,
}

#[derive(Debug, Default)]
pub(super) struct DrawingMlPictureEffectsCapture {
    effects: DrawingMlPictureEffects,
    effect: Option<(usize, DrawingMlPictureEffectTarget)>,
    color_depth: Option<usize>,
    fill_depth: Option<usize>,
    fill_color_depth: Option<usize>,
    scene_3d_depth: Option<usize>,
    camera_3d_depth: Option<usize>,
    light_3d_depth: Option<usize>,
    shape_3d_depth: Option<usize>,
}

impl DrawingMlPictureEffectsCapture {
    #[allow(clippy::too_many_arguments)]
    #[inline(never)]
    pub(super) fn start(
        &mut self,
        local: &str,
        attributes: &[XmlAttribute<'_>],
        empty: bool,
        depth: usize,
        part: &str,
        color: impl Fn(&str, &str) -> Result<u32, Diagnostic>,
    ) -> Result<bool, Diagnostic> {
        let target = match local {
            "outerShdw" | "prstShdw" => Some(DrawingMlPictureEffectTarget::Shadow),
            "innerShdw" => Some(DrawingMlPictureEffectTarget::InnerShadow),
            "glow" => Some(DrawingMlPictureEffectTarget::Glow),
            _ => None,
        };
        if let Some(target) = target {
            let blur = numeric_attribute(
                attributes,
                if matches!(target, DrawingMlPictureEffectTarget::Glow) {
                    "rad"
                } else {
                    "blurRad"
                },
                part,
            )?
            .unwrap_or(0) as f32
                / EMU_PER_CSS_PIXEL;
            if matches!(target, DrawingMlPictureEffectTarget::Glow) {
                self.effects.glow = Some(Glow {
                    color: 0x0000_0080,
                    radius: blur,
                });
            } else {
                let distance = numeric_attribute(attributes, "dist", part)?.unwrap_or(0) as f32
                    / EMU_PER_CSS_PIXEL;
                let direction = signed_numeric_attribute(attributes, "dir", part)?.unwrap_or(0)
                    as f32
                    / 60_000.0;
                let radians = direction.to_radians();
                let shadow = Shadow {
                    color: 0x0000_0080,
                    blur,
                    offset_x: distance * radians.cos(),
                    offset_y: distance * radians.sin(),
                };
                match target {
                    DrawingMlPictureEffectTarget::Shadow => {
                        self.effects.outer_shadow =
                            Some(drawingml_outer_shadow(shadow, attributes, part)?);
                    }
                    DrawingMlPictureEffectTarget::InnerShadow => {
                        self.effects.inner_shadow = Some(shadow);
                    }
                    DrawingMlPictureEffectTarget::Glow => unreachable!(),
                }
            }
            self.effect = (!empty).then_some((depth, target));
            return Ok(true);
        }
        match local {
            "scene3d" => {
                self.effects
                    .three_d
                    .get_or_insert_with(ThreeDStyle::default);
                self.scene_3d_depth = (!empty).then_some(depth);
                return Ok(true);
            }
            "camera" if self.scene_3d_depth.is_some() => {
                parse_three_d_camera(
                    self.effects
                        .three_d
                        .get_or_insert_with(ThreeDStyle::default),
                    attributes,
                    part,
                )?;
                self.camera_3d_depth = (!empty).then_some(depth);
                return Ok(true);
            }
            "lightRig" if self.scene_3d_depth.is_some() => {
                parse_three_d_light(
                    self.effects
                        .three_d
                        .get_or_insert_with(ThreeDStyle::default),
                    attributes,
                    part,
                )?;
                self.light_3d_depth = (!empty).then_some(depth);
                return Ok(true);
            }
            "rot" if self.camera_3d_depth.is_some() => {
                parse_three_d_rotation(
                    self.effects
                        .three_d
                        .get_or_insert_with(ThreeDStyle::default),
                    attributes,
                    part,
                    true,
                )?;
                return Ok(true);
            }
            "rot" if self.light_3d_depth.is_some() => {
                parse_three_d_rotation(
                    self.effects
                        .three_d
                        .get_or_insert_with(ThreeDStyle::default),
                    attributes,
                    part,
                    false,
                )?;
                return Ok(true);
            }
            "sp3d" => {
                parse_three_d_shape(
                    self.effects
                        .three_d
                        .get_or_insert_with(ThreeDStyle::default),
                    attributes,
                    part,
                    false,
                )?;
                self.shape_3d_depth = (!empty).then_some(depth);
                return Ok(true);
            }
            "bevelT" if self.shape_3d_depth.is_some() => {
                self.effects
                    .three_d
                    .get_or_insert_with(ThreeDStyle::default)
                    .bevel_top = Some(parse_three_d_bevel(attributes, part)?);
                return Ok(true);
            }
            "bevelB" if self.shape_3d_depth.is_some() => {
                self.effects
                    .three_d
                    .get_or_insert_with(ThreeDStyle::default)
                    .bevel_bottom = Some(parse_three_d_bevel(attributes, part)?);
                return Ok(true);
            }
            _ => {}
        }
        if self.effect.is_none() && self.effects.background_fill.is_none() && local == "solidFill" {
            self.effects.background_fill = Some(Paint::Solid(0x0000_00ff));
            self.fill_depth = (!empty).then_some(depth);
            return Ok(true);
        }
        if self.effect.is_none() && self.effects.background_fill.is_none() && local == "noFill" {
            self.effects.background_fill = Some(Paint::None);
            return Ok(true);
        }
        if self.fill_depth.is_some()
            && matches!(local, "srgbClr" | "schemeClr" | "prstClr" | "sysClr")
        {
            let value = string_attribute(
                attributes,
                if local == "sysClr" { "lastClr" } else { "val" },
                part,
            )?
            .ok_or_else(|| format_error(part, format!("{local} is missing its color value")))?;
            self.effects.background_fill = Some(Paint::Solid(color(local, &value)?));
            self.fill_color_depth = (!empty).then_some(depth);
            return Ok(true);
        }
        if self.fill_color_depth.is_some()
            && matches!(
                local,
                "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff"
            )
        {
            let ratio = numeric_attribute(attributes, "val", part)?.unwrap_or(0) as f32 / 100_000.0;
            if let Some(Paint::Solid(color)) = self.effects.background_fill.as_mut() {
                apply_color_transform(color, local, ratio);
            }
            return Ok(true);
        }
        match local {
            "reflection" => {
                self.effects.reflection = Some(parse_drawingml_reflection(attributes, part)?);
                return Ok(true);
            }
            "softEdge" => {
                self.effects.soft_edge = Some(
                    numeric_attribute(attributes, "rad", part)?.unwrap_or(0) as f32
                        / EMU_PER_CSS_PIXEL,
                );
                return Ok(true);
            }
            _ => {}
        }
        if let Some((_, target)) = self.effect
            && matches!(local, "srgbClr" | "schemeClr" | "prstClr" | "sysClr")
        {
            let value = string_attribute(
                attributes,
                if local == "sysClr" { "lastClr" } else { "val" },
                part,
            )?
            .ok_or_else(|| format_error(part, format!("{local} is missing its color value")))?;
            self.set_color(target, color(local, &value)?);
            self.color_depth = (!empty).then_some(depth);
            return Ok(true);
        }
        if let Some((_, target)) = self.effect
            && self.color_depth.is_some()
            && matches!(
                local,
                "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff"
            )
        {
            let ratio = numeric_attribute(attributes, "val", part)?.unwrap_or(0) as f32 / 100_000.0;
            self.update_color(target, |value| apply_color_transform(value, local, ratio));
            return Ok(true);
        }
        Ok(false)
    }

    pub(super) fn end(&mut self, depth: usize) {
        if self.shape_3d_depth == Some(depth) {
            self.shape_3d_depth = None;
        }
        if self.camera_3d_depth == Some(depth) {
            self.camera_3d_depth = None;
        }
        if self.light_3d_depth == Some(depth) {
            self.light_3d_depth = None;
        }
        if self.scene_3d_depth == Some(depth) {
            self.scene_3d_depth = None;
            self.camera_3d_depth = None;
            self.light_3d_depth = None;
        }
        if self.fill_color_depth == Some(depth) {
            self.fill_color_depth = None;
        }
        if self.fill_depth == Some(depth) {
            self.fill_depth = None;
            self.fill_color_depth = None;
        }
        if self.color_depth == Some(depth) {
            self.color_depth = None;
        }
        if self
            .effect
            .is_some_and(|(effect_depth, _)| effect_depth == depth)
        {
            self.effect = None;
            self.color_depth = None;
        }
    }

    pub(super) fn finish(self) -> DrawingMlPictureEffects {
        self.effects
    }

    fn set_color(&mut self, target: DrawingMlPictureEffectTarget, color: u32) {
        self.update_color(target, |value| *value = color);
    }

    fn update_color(
        &mut self,
        target: DrawingMlPictureEffectTarget,
        update: impl FnOnce(&mut u32),
    ) {
        match target {
            DrawingMlPictureEffectTarget::Shadow => {
                if let Some(effect) = self.effects.outer_shadow.as_mut() {
                    update(&mut effect.shadow.color);
                }
            }
            DrawingMlPictureEffectTarget::InnerShadow => {
                if let Some(shadow) = self.effects.inner_shadow.as_mut() {
                    update(&mut shadow.color);
                }
            }
            DrawingMlPictureEffectTarget::Glow => {
                if let Some(glow) = self.effects.glow.as_mut() {
                    update(&mut glow.color);
                }
            }
        }
    }
}

pub(super) fn parse_drawingml_reflection(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<Reflection, Diagnostic> {
    Ok(Reflection {
        start_opacity: (numeric_attribute(attributes, "stA", part)?.unwrap_or(100_000) as f32
            / 100_000.0)
            .clamp(0.0, 1.0),
        end_opacity: (numeric_attribute(attributes, "endA", part)?.unwrap_or(0) as f32 / 100_000.0)
            .clamp(0.0, 1.0),
        start_position: (numeric_attribute(attributes, "stPos", part)?.unwrap_or(0) as f32
            / 100_000.0)
            .clamp(0.0, 1.0),
        end_position: (numeric_attribute(attributes, "endPos", part)?.unwrap_or(100_000) as f32
            / 100_000.0)
            .clamp(0.0, 1.0),
        direction_degrees: signed_numeric_attribute(attributes, "dir", part)?.unwrap_or(5_400_000)
            as f32
            / 60_000.0,
        blur: numeric_attribute(attributes, "blurRad", part)?.unwrap_or(0) as f32
            / EMU_PER_CSS_PIXEL,
        distance: numeric_attribute(attributes, "dist", part)?.unwrap_or(0) as f32
            / EMU_PER_CSS_PIXEL,
        scale_x: signed_numeric_attribute(attributes, "sx", part)?
            .unwrap_or(100_000)
            .unsigned_abs() as f32
            / 100_000.0,
        scale_y: signed_numeric_attribute(attributes, "sy", part)?
            .unwrap_or(-100_000)
            .unsigned_abs() as f32
            / 100_000.0,
    })
}

fn drawingml_angle_attribute(
    attributes: &[XmlAttribute<'_>],
    name: &str,
    part: &str,
) -> Result<f32, Diagnostic> {
    Ok(signed_numeric_attribute(attributes, name, part)?.unwrap_or(0) as f32 / 60_000.0)
}

// Preset angles in Office's Y -> X -> Z camera order. Explicit a:rot replaces them.
// Values corroborated by LibreOffice oox/source/drawingml/scene3dhelper.cxx.
fn drawingml_camera_angles(preset: &str) -> (f32, f32, f32) {
    match preset {
        "isometricBottomDown" => (35.4, 314.7, 299.8),
        "isometricBottomUp" => (35.4, 45.3, 60.2),
        "isometricLeftDown" => (35.0, 45.0, 0.0),
        "isometricLeftUp" => (325.0, 45.0, 0.0),
        "isometricOffAxis1Left" => (18.0, 64.0, 0.0),
        "isometricOffAxis1Right" => (18.0, 334.0, 0.0),
        "isometricOffAxis1Top" => (301.3, 306.5, 57.6),
        "isometricOffAxis2Left" => (18.0, 26.0, 0.0),
        "isometricOffAxis2Right" => (18.0, 296.0, 0.0),
        "isometricOffAxis2Top" => (301.3, 53.5, 302.4),
        "isometricOffAxis3Bottom" => (58.7, 306.5, 302.4),
        "isometricOffAxis3Left" => (342.0, 64.0, 0.0),
        "isometricOffAxis3Right" => (342.0, 334.0, 0.0),
        "isometricOffAxis4Bottom" => (58.7, 53.5, 57.6),
        "isometricOffAxis4Left" => (342.0, 26.0, 0.0),
        "isometricOffAxis4Right" => (342.0, 296.0, 0.0),
        "isometricRightDown" => (325.0, 315.0, 0.0),
        "isometricRightUp" => (35.0, 315.0, 0.0),
        "isometricTopDown" => (324.6, 45.3, 299.8),
        "isometricTopUp" => (324.6, 314.7, 60.2),
        "perspectiveAbove" => (340.0, 0.0, 0.0),
        "perspectiveAboveLeftFacing" => (39.3, 14.3, 341.1),
        "perspectiveAboveRightFacing" => (39.3, 345.7, 18.9),
        "perspectiveBelow" => (20.0, 0.0, 0.0),
        "perspectiveContrastingLeftFacing" => (10.4, 43.9, 356.4),
        "perspectiveContrastingRightFacing" => (10.4, 316.1, 3.6),
        "perspectiveHeroicExtremeLeftFacing" => (8.1, 34.5, 357.1),
        "perspectiveHeroicExtremeRightFacing" => (8.1, 325.5, 2.9),
        "perspectiveHeroicLeftFacing" => (349.0, 14.3, 2.6),
        "perspectiveHeroicRightFacing" => (349.0, 345.7, 357.4),
        "perspectiveLeft" => (0.0, 20.0, 0.0),
        "perspectiveRelaxed" => (309.6, 0.0, 0.0),
        "perspectiveRelaxedModerately" => (324.8, 0.0, 0.0),
        "perspectiveRight" => (0.0, 340.0, 0.0),
        _ => (0.0, 0.0, 0.0),
    }
}

pub(super) fn parse_three_d_camera(
    style: &mut ThreeDStyle,
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<(), Diagnostic> {
    style.camera_preset = string_attribute(attributes, "prst", part)?
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| "orthographicFront".to_owned());
    (style.camera_latitude, style.camera_longitude, style.camera_revolution) =
        drawingml_camera_angles(&style.camera_preset);
    let default_fov = if style.camera_preset.starts_with("legacyPerspective") { 65.0 }
        else if style.camera_preset.contains("Extreme") { 80.0 }
        else if style.camera_preset.starts_with("perspective") { 45.0 }
        else { 0.0 };
    style.camera_fov = signed_numeric_attribute(attributes, "fov", part)?
        .map_or(default_fov, |value| (value as f32 / 60_000.0).clamp(0.0, 179.5));
    style.camera_zoom = numeric_attribute(attributes, "zoom", part)?
        .map_or(1.0, |value| value as f32 / 100_000.0)
        .max(0.01);
    Ok(())
}

pub(super) fn parse_three_d_light(
    style: &mut ThreeDStyle,
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<(), Diagnostic> {
    style.light_rig = string_attribute(attributes, "rig", part)?
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| "threePt".to_owned());
    style.light_direction = string_attribute(attributes, "dir", part)?
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| "t".to_owned());
    Ok(())
}

pub(super) fn parse_three_d_rotation(
    style: &mut ThreeDStyle,
    attributes: &[XmlAttribute<'_>],
    part: &str,
    camera: bool,
) -> Result<(), Diagnostic> {
    let latitude = drawingml_angle_attribute(attributes, "lat", part)?;
    let longitude = drawingml_angle_attribute(attributes, "lon", part)?;
    let revolution = drawingml_angle_attribute(attributes, "rev", part)?;
    if camera {
        style.camera_latitude = latitude;
        style.camera_longitude = longitude;
        style.camera_revolution = revolution;
    } else {
        style.light_latitude = latitude;
        style.light_longitude = longitude;
        style.light_revolution = revolution;
    }
    Ok(())
}

pub(super) fn parse_three_d_shape(
    style: &mut ThreeDStyle,
    attributes: &[XmlAttribute<'_>],
    part: &str,
    applies_to_text: bool,
) -> Result<(), Diagnostic> {
    style.applies_to_text = applies_to_text;
    style.z =
        signed_numeric_attribute(attributes, "z", part)?.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL;
    style.extrusion_height =
        numeric_attribute(attributes, "extrusionH", part)?.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL;
    style.contour_width =
        numeric_attribute(attributes, "contourW", part)?.unwrap_or(0) as f32 / EMU_PER_CSS_PIXEL;
    style.material = string_attribute(attributes, "prstMaterial", part)?
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| "warmMatte".to_owned());
    Ok(())
}

pub(super) fn drawingml_camera_point(
    style: &ThreeDStyle,
    bounds: Rect,
    point: (f32, f32),
) -> Option<(f32, f32)> {
    let latitude = style.camera_latitude.to_radians();
    let longitude = style.camera_longitude.to_radians();
    let revolution = style.camera_revolution.to_radians();
    let zoom = style.camera_zoom;
    let (sin_x, cos_x) = latitude.sin_cos();
    let (sin_y, cos_y) = longitude.sin_cos();
    let (sin_z, cos_z) = revolution.sin_cos();
    let a = (cos_z * cos_y + sin_z * sin_x * sin_y) * zoom;
    let b = (-sin_z * cos_y + cos_z * sin_x * sin_y) * zoom;
    let c = sin_z * cos_x * zoom;
    let d = cos_z * cos_x * zoom;
    let center_x = bounds.x + bounds.width / 2.0;
    let center_y = bounds.y + bounds.height / 2.0;
    let x = point.0 - center_x;
    let y = point.1 - center_y;
    let inverse_distance = if style.camera_fov > 0.0 {
        (style.camera_fov.to_radians() / 2.0).tan() / (15976.0 * 96.0 / 2540.0)
    } else { 0.0 };
    let w = 1.0 + (-cos_x * sin_y * x + sin_x * y) * inverse_distance;
    (w > 0.01).then_some((center_x + (a * x + c * y) / w, center_y + (b * x + d * y) / w))
}

pub(super) fn parse_three_d_backdrop_point(
    backdrop: &mut Backdrop3D,
    kind: &str,
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<(), Diagnostic> {
    let x = signed_numeric_attribute(attributes, if kind == "anchor" { "x" } else { "dx" }, part)?
        .unwrap_or(0) as f32;
    let y = signed_numeric_attribute(attributes, if kind == "anchor" { "y" } else { "dy" }, part)?
        .unwrap_or(0) as f32;
    let z = signed_numeric_attribute(attributes, if kind == "anchor" { "z" } else { "dz" }, part)?
        .unwrap_or(0) as f32;
    match kind {
        "anchor" => {
            backdrop.anchor_x = x / EMU_PER_CSS_PIXEL;
            backdrop.anchor_y = y / EMU_PER_CSS_PIXEL;
            backdrop.anchor_z = z / EMU_PER_CSS_PIXEL;
        }
        "norm" => {
            backdrop.normal_x = x;
            backdrop.normal_y = y;
            backdrop.normal_z = z;
        }
        "up" => {
            backdrop.up_x = x;
            backdrop.up_y = y;
            backdrop.up_z = z;
        }
        _ => {}
    }
    Ok(())
}

pub(super) fn parse_three_d_bevel(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<Bevel3D, Diagnostic> {
    Ok(Bevel3D {
        width: numeric_attribute(attributes, "w", part)?.unwrap_or(76_200) as f32
            / EMU_PER_CSS_PIXEL,
        height: numeric_attribute(attributes, "h", part)?.unwrap_or(76_200) as f32
            / EMU_PER_CSS_PIXEL,
        preset: string_attribute(attributes, "prst", part)?
            .filter(|value| !value.is_empty())
            .unwrap_or_else(|| "circle".to_owned()),
    })
}

#[cfg(feature = "native-formats")]
pub(super) fn visit_drawingml_theme_style(
    package: &Package<'_>,
    part: &str,
    list: &str,
    index: u64,
    mut visit: impl FnMut(XmlEvent<'_>, usize) -> Result<(), Diagnostic>,
) -> Result<(), Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut depth = 0_usize;
    let mut list_depth = None;
    let mut item = 0;
    let mut selected = None;
    parse_xml(&bytes, package.limits(), |event| {
        match &event {
            XmlEvent::StartElement { name, empty, .. } => {
                if local_name(name) == list {
                    list_depth = Some(depth);
                } else if list_depth.is_some_and(|start| depth == start + 1) {
                    if item == index {
                        selected = Some(depth);
                    }
                    item += 1;
                }
                let empty = *empty;
                if selected.is_some() {
                    visit(event, depth)?;
                }
                if empty {
                    if selected == Some(depth) {
                        selected = None;
                    }
                } else {
                    depth += 1;
                }
            }
            XmlEvent::EndElement { .. } => {
                depth = depth.saturating_sub(1);
                if selected.is_some() {
                    visit(event, depth)?;
                }
                if selected == Some(depth) {
                    selected = None;
                }
                if list_depth == Some(depth) {
                    list_depth = None;
                }
            }
            _ => {
                if selected.is_some() {
                    visit(event, depth)?;
                }
            }
        }
        Ok(())
    })
    .map(|_| ())
}

#[cfg(feature = "native-formats")]
pub(super) fn drawingml_theme_effects(
    package: &Package<'_>,
    part: &str,
    index: u64,
    color: impl Fn(&str, &str) -> Result<u32, Diagnostic>,
) -> Result<Option<DrawingMlPictureEffects>, Diagnostic> {
    let mut capture = DrawingMlPictureEffectsCapture::default();
    let mut found = false;
    visit_drawingml_theme_style(package, part, "effectStyleLst", index, |event, depth| {
        found = true;
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                capture.start(
                    local_name(name),
                    &attributes,
                    empty,
                    depth,
                    part,
                    &color,
                )?;
            }
            XmlEvent::EndElement { .. } => capture.end(depth),
            _ => {}
        }
        Ok(())
    })?;
    Ok(found.then(|| capture.finish()))
}

pub(super) fn drawingml_shape_transform(
    bounds: Rect,
    rotation_degrees: f32,
    flip_horizontal: bool,
    flip_vertical: bool,
) -> AffineTransform {
    if rotation_degrees == 0.0 && !flip_horizontal && !flip_vertical {
        return AffineTransform::IDENTITY;
    }
    let radians = rotation_degrees.to_radians();
    let cosine = radians.cos();
    let sine = radians.sin();
    let scale_x = if flip_horizontal { -1.0 } else { 1.0 };
    let scale_y = if flip_vertical { -1.0 } else { 1.0 };
    let a = cosine * scale_x;
    let b = sine * scale_x;
    let c = -sine * scale_y;
    let d = cosine * scale_y;
    let center_x = bounds.x + bounds.width / 2.0;
    let center_y = bounds.y + bounds.height / 2.0;
    AffineTransform {
        a,
        b,
        c,
        d,
        e: center_x - a * center_x - c * center_y,
        f: center_y - b * center_x - d * center_y,
    }
}

// Pictures use the same geometry language in all three OOXML hosts.
#[cfg(feature = "native-formats")]
#[derive(Debug, Default)]
pub(super) struct DrawingMlPictureGeometry {
    pub(super) preset: Option<String>,
    pub(super) adjustments: HashMap<String, f32>,
    custom: Option<super::pptx::CustomGeometryState>,
}

#[cfg(feature = "native-formats")]
impl DrawingMlPictureGeometry {
    #[allow(clippy::too_many_arguments)]
    pub(super) fn start(
        &mut self,
        local: &str,
        attributes: &[XmlAttribute<'_>],
        empty: bool,
        depth: usize,
        part: &str,
        width: f32,
        height: f32,
    ) -> Result<(), Diagnostic> {
        if local == "custGeom" {
            self.preset = None;
            self.adjustments.clear();
            self.custom =
                (!empty).then(|| super::pptx::CustomGeometryState::new(depth, width, height));
        } else if let Some(custom) = self.custom.as_mut() {
            custom.start(local, attributes, part, depth, empty)?;
        } else if local == "prstGeom" {
            self.preset = string_attribute(attributes, "prst", part)?;
            self.adjustments.clear();
        } else if local == "gd" && self.preset.is_some() {
            if let (Some(name), Some(formula)) = (
                string_attribute(attributes, "name", part)?,
                string_attribute(attributes, "fmla", part)?,
            ) && let Some(value) = formula
                .strip_prefix("val ")
                .and_then(|s| s.parse::<f32>().ok())
                .filter(|v| v.is_finite())
            {
                self.adjustments.insert(name, value);
            }
        }
        Ok(())
    }

    pub(super) fn end(&mut self, local: &str, depth: usize) {
        if let Some(custom) = self.custom.as_mut() {
            custom.end(local, depth);
        }
    }

    pub(super) fn custom_geometry(&mut self, bounds: Rect) -> Option<Geometry> {
        self.custom
            .take()
            .and_then(|custom| custom.into_geometry(bounds))
    }

    pub(super) fn geometry(&mut self, bounds: Rect) -> Option<Geometry> {
        self.custom_geometry(bounds).or_else(|| {
            self.preset.as_deref().and_then(|preset| {
                super::pptx::resolve_preset_geometry(
                    preset,
                    bounds,
                    &self.adjustments,
                    0.0,
                    None,
                    None,
                )
                .map(|value| value.0)
            })
        })
    }
}

pub(super) fn drawingml_picture_crop(
    attributes: &[XmlAttribute<'_>],
    part: &str,
) -> Result<ImageCrop, Diagnostic> {
    let value = |name| percentage_attribute(attributes, name, part).map(|v| v.unwrap_or(0.0));
    let crop = ImageCrop {
        left: value("l")?,
        top: value("t")?,
        right: value("r")?,
        bottom: value("b")?,
    };
    if !crop.is_valid() {
        return Err(format_error(part, "picture source crop is invalid"));
    }
    Ok(crop)
}
