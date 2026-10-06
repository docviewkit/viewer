#[derive(Default)]
struct DiagramColorList {
    colors: Vec<u32>,
    method: String,
    counterclockwise: bool,
}

impl DiagramColorList {
    fn sample(&self, index: usize, count: usize) -> Option<u32> {
        let first = *self.colors.first()?;
        if self.method == "repeat" {
            return Some(self.colors[index % self.colors.len()]);
        }
        if self.colors.len() == 1 || count <= 1 {
            return Some(first);
        }
        let mut progress = index.min(count - 1) as f32 / (count - 1) as f32;
        if self.method == "cycle" {
            progress = 1.0 - (2.0 * progress - 1.0).abs();
        }
        let position = progress * (self.colors.len() - 1) as f32;
        let left = (position.floor() as usize).min(self.colors.len() - 1);
        let right = (left + 1).min(self.colors.len() - 1);
        let ratio = position - left as f32;
        let (a, b) = (self.colors[left], self.colors[right]);
        let (h1, s1, l1) = color_to_hsl(a);
        let (h2, s2, l2) = color_to_hsl(b);
        let hue_delta = if (h2 - h1).abs() < f32::EPSILON {
            0.0
        } else if self.counterclockwise {
            -((h1 - h2).rem_euclid(1.0))
        } else {
            (h2 - h1).rem_euclid(1.0)
        };
        let alpha = ((a & 255) as f32 * (1.0 - ratio) + (b & 255) as f32 * ratio).round() as u32;
        Some(color_from_hsl(
            (a & 0xffff_ff00) | alpha,
            (h1 + hue_delta * ratio).rem_euclid(1.0),
            s1 + (s2 - s1) * ratio,
            l1 + (l2 - l1) * ratio,
        ))
    }
}

fn parse_diagram_color_list(
    package: &Package<'_>,
    part: &str,
    scheme_color: &impl Fn(&str) -> Option<u32>,
    list_name: &str,
) -> Result<HashMap<String, DiagramColorList>, Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut depth = 0_usize;
    let mut label = None::<(usize, String)>;
    let mut text_fill_depth = None;
    let mut color = None::<(usize, u32)>;
    let mut colors = HashMap::<String, DiagramColorList>::new();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "styleLbl" {
                    label = string_attribute(&attributes, "name", part)?.map(|name| (depth, name));
                } else if local == list_name && label.is_some() {
                    if let Some((_, label)) = &label {
                        let list = colors.entry(label.clone()).or_default();
                        list.method = string_attribute(&attributes, "meth", part)?
                            .unwrap_or_else(|| "span".to_owned());
                        list.counterclockwise = string_attribute(&attributes, "hueDir", part)?
                            .as_deref()
                            == Some("ccw");
                    }
                    text_fill_depth = (!empty).then_some(depth);
                } else if text_fill_depth.is_some()
                    && matches!(local, "srgbClr" | "schemeClr" | "prstClr" | "sysClr")
                {
                    let value = string_attribute(&attributes, "val", part)?;
                    let resolved = match local {
                        "srgbClr" => value.as_deref().and_then(parse_rgb_color),
                        "schemeClr" => value.as_deref().and_then(scheme_color),
                        "prstClr" if value.as_deref() == Some("black") => Some(0x0000_00ff),
                        "prstClr" if value.as_deref() == Some("white") => Some(0xffff_ffff),
                        "sysClr" => string_attribute(&attributes, "lastClr", part)?
                            .as_deref()
                            .and_then(parse_rgb_color),
                        _ => None,
                    };
                    if let Some(resolved) = resolved {
                        if empty {
                            if let Some((_, label)) = &label {
                                colors
                                    .entry(label.clone())
                                    .or_default()
                                    .colors
                                    .push(resolved);
                            }
                        } else {
                            color = Some((depth, resolved));
                        }
                    }
                } else if color.is_some()
                    && matches!(local, "alpha" | "tint" | "shade" | "lumMod" | "lumOff")
                    && let Some(value) = numeric_attribute(&attributes, "val", part)?
                    && let Some((_, color)) = color.as_mut()
                {
                    apply_color_transform(color, local, value as f32 / 100_000.0);
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if color.as_ref().is_some_and(|(start, _)| *start == depth)
                    && let Some((_, resolved)) = color.take()
                    && let Some((_, label)) = &label
                {
                    colors
                        .entry(label.clone())
                        .or_default()
                        .colors
                        .push(resolved);
                }
                if local == list_name && text_fill_depth == Some(depth) {
                    text_fill_depth = None;
                }
                if local == "styleLbl" && label.as_ref().is_some_and(|(start, _)| *start == depth) {
                    label = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(colors)
}

// Quick-style labels refer to the host theme's format matrix. Keep that
// indirection here so every diagram host uses the same paints and effects.
#[derive(Default)]
struct DiagramQuickStyle {
    effects: DrawingMlPictureEffects,
    fill_index: Option<u64>,
    line_index: Option<u64>,
    effect_index: Option<u64>,
}

fn diagram_quick_styles(
    package: &Package<'_>,
    part: &str,
    expected_id: &str,
    color: &dyn Fn(&str) -> Option<u32>,
) -> Result<HashMap<String, DiagramQuickStyle>, Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut depth = 0;
    let mut matched = false;
    let mut label = None;
    let mut style = DiagramQuickStyle::default();
    let mut capture = DrawingMlPictureEffectsCapture::default();
    let mut styles = HashMap::new();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if depth == 0 {
                    matched = string_attribute(&attributes, "uniqueId", part)?.as_deref()
                        == Some(expected_id);
                }
                if matched {
                    if local == "styleLbl" {
                        label = string_attribute(&attributes, "name", part)?;
                        style = DiagramQuickStyle::default();
                        capture = DrawingMlPictureEffectsCapture::default();
                    } else if label.is_some() {
                        match local {
                            "fillRef" => {
                                style.fill_index = numeric_attribute(&attributes, "idx", part)?
                            }
                            "lnRef" => {
                                style.line_index = numeric_attribute(&attributes, "idx", part)?
                            }
                            "effectRef" => {
                                style.effect_index = numeric_attribute(&attributes, "idx", part)?
                            }
                            _ => {
                                capture.start(
                                    local,
                                    &attributes,
                                    empty,
                                    depth,
                                    part,
                                    |kind, value| {
                                        Ok(if kind == "schemeClr" {
                                            color(value)
                                        } else {
                                            parse_rgb_color(value)
                                        }
                                        .unwrap_or(0x0000_00ff))
                                    },
                                )?;
                            }
                        }
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                capture.end(depth);
                if local_name(name) == "styleLbl"
                    && let Some(label) = label.take()
                {
                    style.effects = std::mem::take(&mut capture).finish();
                    styles.insert(label, std::mem::take(&mut style));
                }
            }
            _ => {}
        }
        Ok(())
    })
    .map_err(|e| with_part(e, part))?;
    Ok(styles)
}

fn apply_diagram_quick_style(
    package: &Package<'_>,
    theme_part: Option<&str>,
    quick: &DiagramQuickStyle,
    style: &mut DiagramStyle,
    explicit_fill: bool,
    color: &dyn Fn(&str) -> Option<u32>,
) -> Result<(), Diagnostic> {
    style.effects = quick.effects.clone();
    if quick.line_index == Some(0) {
        style.line_width = Some(0.0);
    }
    let Some(part) = theme_part else {
        return Ok(());
    };
    let placeholder = style
        .fill
        .as_ref()
        .and_then(|fill| match fill {
            ChartFill::Solid(color) => Some(*color),
            _ => None,
        })
        .unwrap_or(0x0000_00ff);
    let scheme = |name: &str| {
        if name == "phClr" {
            Some(placeholder)
        } else {
            color(name)
        }
    };
    if !explicit_fill && let Some(index) = quick.fill_index {
        if index == 0 || index == 1000 {
            style.fill = Some(ChartFill::Image(Paint::None));
        } else {
            let (list, index) = if index >= 1001 {
                ("bgFillStyleLst", index - 1001)
            } else {
                ("fillStyleLst", index - 1)
            };
            let mut capture = None::<ChartFillCapture>;
            visit_drawingml_theme_style(package, part, list, index, |event, depth| {
                match event {
                    XmlEvent::StartElement {
                        name, attributes, ..
                    } => {
                        let local = local_name(name);
                        if local == "noFill" {
                            style.fill = Some(ChartFill::Image(Paint::None));
                        }
                        if capture.is_none() {
                            capture = ChartFillCapture::new(local, depth);
                        }
                        if let Some(capture) = capture.as_mut() {
                            capture.start(local, &attributes, part, &scheme)?;
                        }
                    }
                    XmlEvent::EndElement { name } => {
                        if let Some(capture) = capture.as_mut() {
                            capture.end(local_name(name));
                        }
                        if capture.as_ref().is_some_and(|c| c.closes_at(depth)) {
                            if let Some(fill) = capture.take().unwrap().finish(package, part)? {
                                style.fill = Some(fill);
                            }
                        }
                    }
                    _ => {}
                }
                Ok(())
            })?;
        }
    }
    if let Some(index) = quick.line_index.and_then(|i| i.checked_sub(1)) {
        visit_drawingml_theme_style(package, part, "lnStyleLst", index, |event, _| {
            if let XmlEvent::StartElement {
                name, attributes, ..
            } = event
                && local_name(name) == "ln"
            {
                style.line_width = Some(
                    numeric_attribute(&attributes, "w", part)?.unwrap_or(12700) as f32
                        / EMU_PER_CSS_PIXEL,
                );
            }
            Ok(())
        })?;
    }
    if let Some(index) = quick.effect_index.and_then(|i| i.checked_sub(1))
        && let Some(mut effects) = drawingml_theme_effects(package, part, index, |kind, value| {
            Ok(if kind == "schemeClr" {
                scheme(value)
            } else {
                parse_rgb_color(value)
            }
            .unwrap_or(0x0000_00ff))
        })?
    {
        effects.three_d = style.effects.three_d.take().or(effects.three_d);
        style.effects = effects;
    }
    Ok(())
}

pub(super) fn parse_diagram(
    package: &Package<'_>,
    part: &str,
    owner_relationships: &HashMap<&str, &Relationship>,
    colors_part: Option<&str>,
    theme_part: Option<&str>,
    scheme_color: impl Fn(&str) -> Option<u32>,
) -> Result<Diagram, Diagnostic> {
    parse_diagram_shared(package, part, owner_relationships, colors_part, theme_part, &scheme_color)
}

// DOCX/PPTX and XLSX use the same parser with different theme callbacks.
#[inline(never)]
fn parse_diagram_shared(
    package: &Package<'_>,
    part: &str,
    owner_relationships: &HashMap<&str, &Relationship>,
    colors_part: Option<&str>,
    theme_part: Option<&str>,
    scheme_color: &dyn Fn(&str) -> Option<u32>,
) -> Result<Diagram, Diagnostic> {
    struct PointState {
        depth: usize,
        model_id: String,
        point_type: Option<String>,
        placeholder: bool,
        presentation_assoc_id: Option<String>,
        presentation_name: Option<String>,
        hierarchy_branch: Option<String>,
        presentation_style_label: Option<String>,
        presentation_style_index: usize,
        presentation_style_count: usize,
        preset_geometry: Option<String>,
        text_depth: Option<usize>,
        bold: bool,
        font_size: Option<f32>,
        text: String,
        custom_geometry: bool,
        shape_properties_depth: Option<usize>,
        fill_capture: Option<ChartFillCapture>,
        fill: Option<ChartFill>,
        no_fill: bool,
        shadow_depth: Option<usize>,
        shadow: Option<Shadow>,
        three_d: Option<ThreeDStyle>,
        fill_scheme: Option<String>,
        fill_transforms: Vec<(String, f32)>,
    }

    let data_relationships = package.relationships(Some(part))?;
    let legacy_drawing_part = data_relationships
        .iter()
        .find(|relationship| {
            !relationship.external
                && relationship
                    .type_uri
                    .to_ascii_lowercase()
                    .contains("diagramdrawing")
        })
        .map(|relationship| relationship.target.clone());
    let bytes = package.required_part(part)?;
    let mut depth = 0_usize;
    let mut point: Option<PointState> = None;
    let mut nodes = Vec::new();
    let mut presentation_geometries = HashMap::new();
    let mut presentation_fills = HashMap::new();
    let mut presentation_text_styles = HashMap::new();
    let mut presentation_roles = Vec::new();
    let mut parents = HashMap::new();
    let mut sibling_orders = HashMap::new();
    let mut hierarchy_branches = HashMap::new();
    let mut parent_transitions = HashMap::new();
    let mut drawing_relationship_id = None;
    let mut layout_type = None;
    let mut quick_style_id = None;
    let mut right_to_left = false;
    let mut background_depth = None;
    let mut background_fill_capture = None::<ChartFillCapture>;
    let mut background_fill = None;
    let mut background_shadow_depth = None;
    let mut background_shadow = None;
    let mut scene_three_d = false;
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "dataModelExt" && drawing_relationship_id.is_none() {
                    drawing_relationship_id = string_attribute(&attributes, "relId", part)?;
                } else if local == "prSet"
                    && layout_type.is_none()
                    && point
                        .as_ref()
                        .is_some_and(|point| point.point_type.as_deref() == Some("doc"))
                {
                    layout_type = string_attribute(&attributes, "loTypeId", part)?;
                    quick_style_id = string_attribute(&attributes, "qsTypeId", part)?;
                } else if local == "dir" {
                    right_to_left |=
                        string_attribute(&attributes, "val", part)?.as_deref() == Some("rev");
                } else if local == "bg" && point.is_none() {
                    background_depth = (!empty).then_some(depth);
                } else if matches!(local, "solidFill" | "gradFill" | "pattFill" | "blipFill")
                    && background_depth.is_some()
                    && background_fill_capture.is_none()
                {
                    background_fill_capture = ChartFillCapture::new(local, depth);
                } else if local == "outerShdw" && background_depth.is_some() && point.is_none() {
                    let distance =
                        numeric_attribute(&attributes, "dist", part)?.unwrap_or(0) as f32 / 9_525.0;
                    let direction = signed_numeric_attribute(&attributes, "dir", part)?.unwrap_or(0)
                        as f32
                        / 60_000.0;
                    let radians = direction.to_radians();
                    background_shadow_depth = (!empty).then_some(depth);
                    background_shadow = Some(Shadow {
                        color: 0x0000_0080,
                        blur: numeric_attribute(&attributes, "blurRad", part)?.unwrap_or(0) as f32
                            / 9_525.0,
                        offset_x: distance * radians.cos(),
                        offset_y: distance * radians.sin(),
                    });
                } else if matches!(local, "srgbClr" | "schemeClr" | "prstClr")
                    && background_shadow_depth.is_some()
                    && let Some(value) = string_attribute(&attributes, "val", part)?
                {
                    let color = match local {
                        "srgbClr" => parse_rgb_color(&value),
                        "schemeClr" => scheme_color(&value),
                        "prstClr" if value == "black" => Some(0x0000_00ff),
                        "prstClr" if value == "white" => Some(0xffff_ffff),
                        _ => None,
                    };
                    if let (Some(shadow), Some(color)) = (background_shadow.as_mut(), color) {
                        shadow.color = color;
                    }
                } else if matches!(local, "alpha" | "tint" | "shade" | "lumMod" | "lumOff")
                    && background_shadow_depth.is_some()
                    && let Some(value) = numeric_attribute(&attributes, "val", part)?
                    && let Some(shadow) = background_shadow.as_mut()
                {
                    apply_color_transform(&mut shadow.color, local, value as f32 / 100_000.0);
                } else if local == "pt" && point.is_none() {
                    let model_id = string_attribute(&attributes, "modelId", part)?
                        .unwrap_or_else(|| format!("node-{}", nodes.len() + 1));
                    let point_type = string_attribute(&attributes, "type", part)?;
                    if empty {
                        nodes.push(DiagramNode {
                            model_id,
                            parent_id: None,
                            sibling_order: 0,
                            assistant: point_type.as_deref() == Some("asst"),
                            hierarchy_branch: None,
                            parent_transition_id: None,
                            bold: false,
                            font_size: None,
                            text: String::new(),
                            placeholder: false,
                            preset_geometry: None,
                            custom_geometry: false,
                            fill: None,
                            no_fill: false,
                            shadow: None,
                            three_d: None,
                            fill_scheme: None,
                            fill_transforms: Vec::new(),
                        });
                    } else {
                        point = Some(PointState {
                            depth,
                            model_id,
                            point_type,
                            placeholder: false,
                            presentation_assoc_id: None,
                            presentation_name: None,
                            hierarchy_branch: None,
                            presentation_style_label: None,
                            presentation_style_index: 0,
                            presentation_style_count: 1,
                            preset_geometry: None,
                            text_depth: None,
                            bold: false,
                            font_size: None,
                            text: String::new(),
                            custom_geometry: false,
                            shape_properties_depth: None,
                            fill_capture: None,
                            fill: None,
                            no_fill: false,
                            shadow_depth: None,
                            shadow: None,
                            three_d: None,
                            fill_scheme: None,
                            fill_transforms: Vec::new(),
                        });
                    }
                } else if local == "prSet"
                    && let Some(point) = point.as_mut()
                {
                    point.placeholder = string_attribute(&attributes, "phldr", part)?
                        .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"));
                    point.presentation_assoc_id =
                        string_attribute(&attributes, "presAssocID", part)?;
                    point.presentation_name = string_attribute(&attributes, "presName", part)?;
                    point.presentation_style_label =
                        string_attribute(&attributes, "presStyleLbl", part)?;
                    point.presentation_style_index =
                        numeric_attribute(&attributes, "presStyleIdx", part)?
                            .and_then(|value| usize::try_from(value).ok())
                            .unwrap_or(0);
                    point.presentation_style_count =
                        numeric_attribute(&attributes, "presStyleCnt", part)?
                            .and_then(|value| usize::try_from(value).ok())
                            .unwrap_or(1)
                            .max(1);
                } else if local == "hierBranch"
                    && let Some(point) = point.as_mut()
                {
                    point.hierarchy_branch = string_attribute(&attributes, "val", part)?;
                } else if local == "rPr"
                    && let Some(point) = point.as_mut()
                {
                    if point.font_size.is_none() {
                        point.font_size = numeric_attribute(&attributes, "sz", part)?
                            .map(|size| size as f32 / 75.0)
                            .filter(|size| *size > 0.0);
                    }
                    point.bold |= string_attribute(&attributes, "b", part)?
                        .is_some_and(|value| value == "1" || value == "true");
                } else if local == "prstGeom"
                    && let Some(point) = point.as_mut()
                {
                    point.preset_geometry = string_attribute(&attributes, "prst", part)?;
                } else if local == "custGeom"
                    && let Some(point) = point.as_mut()
                {
                    point.custom_geometry = true;
                } else if local == "spPr"
                    && let Some(point) = point.as_mut()
                {
                    point.shape_properties_depth = (!empty).then_some(depth);
                } else if local == "scene3d"
                    && point.as_ref().and_then(|point| point.point_type.as_deref()) == Some("doc")
                {
                    scene_three_d = true;
                } else if local == "outerShdw"
                    && let Some(point) = point.as_mut()
                    && point.shape_properties_depth.is_some()
                {
                    let distance =
                        numeric_attribute(&attributes, "dist", part)?.unwrap_or(0) as f32 / 9_525.0;
                    let direction = signed_numeric_attribute(&attributes, "dir", part)?.unwrap_or(0)
                        as f32
                        / 60_000.0;
                    let radians = direction.to_radians();
                    point.shadow_depth = (!empty).then_some(depth);
                    point.shadow = Some(Shadow {
                        color: 0x0000_0080,
                        blur: numeric_attribute(&attributes, "blurRad", part)?.unwrap_or(0) as f32
                            / 9_525.0,
                        offset_x: distance * radians.cos(),
                        offset_y: distance * radians.sin(),
                    });
                } else if local == "bevelT"
                    && let Some(point) = point.as_mut()
                    && point.shape_properties_depth.is_some()
                {
                    point
                        .three_d
                        .get_or_insert_with(ThreeDStyle::default)
                        .bevel_top = Some(parse_three_d_bevel(&attributes, part)?);
                } else if local == "bevelB"
                    && let Some(point) = point.as_mut()
                    && point.shape_properties_depth.is_some()
                {
                    point
                        .three_d
                        .get_or_insert_with(ThreeDStyle::default)
                        .bevel_bottom = Some(parse_three_d_bevel(&attributes, part)?);
                } else if matches!(local, "srgbClr" | "schemeClr" | "prstClr")
                    && let Some(point) = point.as_mut().filter(|point| point.shadow_depth.is_some())
                    && let Some(value) = string_attribute(&attributes, "val", part)?
                {
                    let color = match local {
                        "srgbClr" => parse_rgb_color(&value),
                        "schemeClr" => scheme_color(&value),
                        "prstClr" if value == "black" => Some(0x0000_00ff),
                        "prstClr" if value == "white" => Some(0xffff_ffff),
                        _ => None,
                    };
                    if let (Some(shadow), Some(color)) = (point.shadow.as_mut(), color) {
                        shadow.color = color;
                    }
                } else if matches!(local, "solidFill" | "gradFill" | "pattFill" | "blipFill")
                    && let Some(point) = point.as_mut()
                    && point
                        .shape_properties_depth
                        .is_some_and(|shape_depth| depth == shape_depth + 1)
                {
                    point.fill_capture = ChartFillCapture::new(local, depth);
                } else if local == "noFill"
                    && let Some(point) = point.as_mut()
                    && point
                        .shape_properties_depth
                        .is_some_and(|shape_depth| depth == shape_depth + 1)
                {
                    point.no_fill = true;
                } else if local == "schemeClr"
                    && let Some(point) = point.as_mut().filter(|point| point.custom_geometry)
                    && point.fill_scheme.is_none()
                {
                    point.fill_scheme = string_attribute(&attributes, "val", part)?;
                } else if matches!(
                    local,
                    "alpha"
                        | "tint"
                        | "shade"
                        | "lumMod"
                        | "lumOff"
                        | "satMod"
                        | "satOff"
                        | "hueMod"
                        | "hueOff"
                ) && let Some(point) =
                    point.as_mut().filter(|point| point.fill_scheme.is_some())
                    && let Some(value) = numeric_attribute(&attributes, "val", part)?
                {
                    point
                        .fill_transforms
                        .push((local.to_owned(), value as f32 / 100_000.0));
                } else if matches!(local, "alpha" | "tint" | "shade" | "lumMod" | "lumOff")
                    && let Some(point) = point.as_mut().filter(|point| point.shadow_depth.is_some())
                    && let Some(value) = numeric_attribute(&attributes, "val", part)?
                    && let Some(shadow) = point.shadow.as_mut()
                {
                    apply_color_transform(&mut shadow.color, local, value as f32 / 100_000.0);
                } else if name != "dgm:t"
                    && local == "t"
                    && let Some(point) = point.as_mut()
                {
                    point.text_depth = (!empty).then_some(depth);
                } else if local == "cxn"
                    && string_attribute(&attributes, "type", part)?
                        .is_none_or(|kind| kind == "parOf")
                    && let (Some(source), Some(destination)) = (
                        string_attribute(&attributes, "srcId", part)?,
                        string_attribute(&attributes, "destId", part)?,
                    )
                {
                    sibling_orders.insert(
                        destination.clone(),
                        numeric_attribute(&attributes, "srcOrd", part)?.unwrap_or(0),
                    );
                    if let Some(transition) = string_attribute(&attributes, "parTransId", part)? {
                        parent_transitions.insert(destination.clone(), transition);
                    }
                    parents.insert(destination, source);
                }
                if let Some(capture) = background_fill_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = point.as_mut().and_then(|point| point.fill_capture.as_mut())
                {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if let Some(capture) = background_fill_capture.as_mut() {
                    capture.end(local);
                }
                if let Some(capture) = point.as_mut().and_then(|point| point.fill_capture.as_mut())
                {
                    capture.end(local);
                }
                if background_fill_capture
                    .as_ref()
                    .is_some_and(|capture| capture.depth == depth)
                {
                    background_fill = background_fill_capture
                        .take()
                        .ok_or_else(|| {
                            format_error(part, "diagram background fill state is missing")
                        })?
                        .finish(package, part)?;
                }
                if local == "bg" && background_depth == Some(depth) {
                    background_depth = None;
                }
                if local == "outerShdw" && background_shadow_depth == Some(depth) {
                    background_shadow_depth = None;
                }
                if point
                    .as_ref()
                    .and_then(|point| point.fill_capture.as_ref())
                    .is_some_and(|capture| capture.depth == depth)
                {
                    let current = point.as_mut().ok_or_else(|| {
                        format_error(part, "diagram point fill state has no owning point")
                    })?;
                    current.fill = current
                        .fill_capture
                        .take()
                        .ok_or_else(|| format_error(part, "diagram point fill state is missing"))?
                        .finish(package, part)?;
                }
                if local == "spPr"
                    && let Some(point) = point.as_mut()
                    && point.shape_properties_depth == Some(depth)
                {
                    point.shape_properties_depth = None;
                }
                if local == "outerShdw"
                    && let Some(point) = point.as_mut()
                    && point.shadow_depth == Some(depth)
                {
                    point.shadow_depth = None;
                }
                if local == "t"
                    && let Some(point) = point.as_mut()
                    && point.text_depth == Some(depth)
                {
                    point.text_depth = None;
                }
                if local == "pt" && point.as_ref().is_some_and(|point| point.depth == depth) {
                    let point = point.take().ok_or_else(|| {
                        format_error(part, "diagram point parser state ended unexpectedly")
                    })?;
                    if let Some(branch) = &point.hierarchy_branch {
                        hierarchy_branches.insert(
                            point
                                .presentation_assoc_id
                                .as_ref()
                                .unwrap_or(&point.model_id)
                                .clone(),
                            branch.clone(),
                        );
                    }
                    if point.point_type.as_deref() == Some("pres") {
                        if let Some(label) = point.presentation_style_label.clone() {
                            presentation_text_styles.insert(
                                point.model_id.clone(),
                                (
                                    label,
                                    point.presentation_style_index,
                                    point.presentation_style_count,
                                ),
                            );
                        }
                        if let (Some(associated_id), Some(role), Some(label)) = (
                            point.presentation_assoc_id.clone(),
                            point.presentation_name.clone(),
                            point.presentation_style_label,
                        ) {
                            presentation_roles.push((
                                associated_id,
                                role,
                                label,
                                point.presentation_style_index,
                                point.presentation_style_count,
                                point.fill.clone(),
                                point.shadow,
                            ));
                        }
                        if let Some(associated_id) = point.presentation_assoc_id {
                            if point.presentation_name.as_deref() == Some("pictRect")
                                && let Some(preset) = point.preset_geometry
                            {
                                presentation_geometries.insert(associated_id.clone(), preset);
                            }
                            if let Some(fill) = point.fill {
                                presentation_fills.insert(associated_id, fill);
                            }
                        }
                    } else if !matches!(
                        point.point_type.as_deref(),
                        Some("parTrans" | "sibTrans" | "doc")
                    ) {
                        nodes.push(DiagramNode {
                            model_id: point.model_id,
                            parent_id: None,
                            sibling_order: 0,
                            assistant: point.point_type.as_deref() == Some("asst"),
                            hierarchy_branch: None,
                            parent_transition_id: None,
                            bold: point.bold,
                            font_size: point.font_size,
                            text: point.text,
                            placeholder: point.placeholder,
                            preset_geometry: point.preset_geometry,
                            custom_geometry: point.custom_geometry,
                            fill: point.fill,
                            no_fill: point.no_fill,
                            shadow: point.shadow,
                            three_d: point.three_d,
                            fill_scheme: point.fill_scheme,
                            fill_transforms: point.fill_transforms,
                        });
                    }
                }
            }
            XmlEvent::Text(text) => {
                if let Some(point) = point.as_mut().filter(|point| point.text_depth.is_some()) {
                    point
                        .text
                        .push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(point) = point.as_mut().filter(|point| point.text_depth.is_some()) {
                    point.text.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    for node in &mut nodes {
        node.parent_id = parents.get(&node.model_id).cloned();
        node.hierarchy_branch = hierarchy_branches.remove(&node.model_id);
        node.parent_transition_id = parent_transitions.remove(&node.model_id);
        node.sibling_order = sibling_orders.get(&node.model_id).copied().unwrap_or(0);
        if node.preset_geometry.is_none() {
            if let Some(preset) = presentation_geometries.remove(&node.model_id) {
                node.preset_geometry = Some(preset);
                node.no_fill = false;
            }
        }
        if node.fill.is_none()
            && let Some(fill) = presentation_fills.remove(&node.model_id)
        {
            node.fill = Some(fill);
        }
    }
    nodes.retain(|node| {
        !node.text.trim().is_empty()
            || node.placeholder
            || node.preset_geometry.is_some()
            || node.custom_geometry
            || node.fill.is_some()
    });
    let owner_drawing_part = drawing_relationship_id
        .as_deref()
        .and_then(|relationship_id| owner_relationships.get(relationship_id))
        .filter(|relationship| {
            !relationship.external
                && relationship
                    .type_uri
                    .to_ascii_lowercase()
                    .contains("diagramdrawing")
        })
        .map(|relationship| relationship.target.clone());
    let mut drawing_parts = Vec::new();
    if let Some(drawing_part) = owner_drawing_part {
        drawing_parts.push((drawing_part, true));
    }
    let owner_diagram_count = owner_relationships
        .values()
        .filter(|relationship| {
            !relationship.external
                && relationship
                    .type_uri
                    .to_ascii_lowercase()
                    .contains("diagramdata")
        })
        .count();
    if owner_diagram_count == 1 {
        let mut alternatives = owner_relationships
            .values()
            .filter(|relationship| {
                !relationship.external
                    && relationship
                        .type_uri
                        .to_ascii_lowercase()
                        .contains("diagramdrawing")
            })
            .collect::<Vec<_>>();
        alternatives.sort_by(|left, right| left.target.cmp(&right.target));
        for relationship in alternatives {
            if !drawing_parts
                .iter()
                .any(|(part, _)| part == &relationship.target)
            {
                drawing_parts.push((relationship.target.clone(), true));
            }
        }
    }
    if let Some(drawing_part) = legacy_drawing_part
        && !drawing_parts.iter().any(|(part, _)| part == &drawing_part)
    {
        drawing_parts.push((drawing_part, false));
    }
    let diagram_text_colors = colors_part
        .map(|colors_part| {
            parse_diagram_color_list(package, colors_part, &scheme_color, "txFillClrLst")
        })
        .transpose()?
        .unwrap_or_default();
    let fill_colors = colors_part
        .map(|part| parse_diagram_color_list(package, part, &scheme_color, "fillClrLst"))
        .transpose()?
        .unwrap_or_default();
    let line_colors = colors_part
        .map(|part| parse_diagram_color_list(package, part, &scheme_color, "linClrLst"))
        .transpose()?
        .unwrap_or_default();
    let mut diagnostics = Vec::new();
    let mut quick_styles = HashMap::new();
    if let Some(id) = quick_style_id.as_deref() {
        for relationship in owner_relationships
            .values()
            .filter(|r| !r.external && r.type_uri.ends_with("/diagramQuickStyle"))
        {
            match diagram_quick_styles(package, &relationship.target, id, scheme_color) {
                Ok(styles) => quick_styles.extend(styles),
                Err(mut error) => {
                    error.severity = crate::diagnostic::Severity::Warning;
                    error.fidelity = crate::diagnostic::Fidelity::Approximate;
                    diagnostics.push(error);
                }
            }
        }
    }
    let fallback_theme = if theme_part.is_none() && !quick_styles.is_empty() {
        package
            .relationships(None)?
            .into_iter()
            .find(|r| !r.external && r.type_uri.ends_with("/officeDocument"))
            .map(|r| package.relationships(Some(&r.target)))
            .transpose()?
            .unwrap_or_default()
            .into_iter()
            .find(|r| !r.external && r.type_uri.ends_with("/theme"))
            .map(|r| r.target)
    } else {
        None
    };
    let theme_part = theme_part.or(fallback_theme.as_deref());
    let role_line_colors = presentation_roles
        .iter()
        .filter_map(|(id, role, label, index, count, _, _)| {
            line_colors
                .get(label)?
                .sample(*index, *count)
                .map(|color| ((id.clone(), role.clone()), color))
        })
        .collect();
    let role_fills = presentation_roles
        .into_iter()
        .map(|(id, role, label, index, count, fill, shadow)| {
            let explicit_fill = fill.is_some();
            let fill = fill.or_else(|| {
                fill_colors
                    .get(&label)
                    .and_then(|colors| colors.sample(index, count))
                    .map(ChartFill::Solid)
            });
            let mut style = DiagramStyle {
                fill,
                shadow,
                ..DiagramStyle::default()
            };
            style.text_color = diagram_text_colors
                .get(&label)
                .and_then(|colors| colors.sample(index, count));
            if let Some(quick) = quick_styles.get(&label) {
                if let Err(error) = apply_diagram_quick_style(
                    package,
                    theme_part,
                    quick,
                    &mut style,
                    explicit_fill,
                    scheme_color,
                ) {
                    let mut error = error;
                    error.severity = crate::diagnostic::Severity::Warning;
                    error.fidelity = crate::diagnostic::Fidelity::Approximate;
                    diagnostics.push(error);
                }
            }
            ((id, role), style)
        })
        .collect();
    let drawing_text_colors = presentation_text_styles
        .into_iter()
        .filter_map(|(model_id, (label, index, count))| {
            let colors = diagram_text_colors.get(&label)?;
            colors.sample(index, count).map(|color| (model_id, color))
        })
        .collect();
    Ok(Diagram {
        data_part: part.to_owned(),
        diagnostics,
        drawing_parts,
        drawing_text_colors,
        role_fills,
        role_line_colors,
        layout_type,
        right_to_left,
        scene_three_d,
        background_fill,
        background_shadow,
        nodes,
    })
}
