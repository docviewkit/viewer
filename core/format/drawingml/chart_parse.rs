#[derive(Clone, Copy, Debug)]
enum ChartValueTarget {
    Name,
    Categories,
    XValues,
    Values,
    BubbleSizes,
    ErrorPlus,
    ErrorMinus,
}

// A theme override is a partial palette; absent slots inherit the host theme.
// Unlike the strict DOCX/XLSX theme readers, invalid RGB slots are skipped here,
// preserving the existing DrawingML/PPTX best-effort color behavior.
pub(super) fn drawingml_theme_colors(
    bytes: &[u8],
    limits: crate::limits::Limits,
    part: &str,
) -> Result<HashMap<String, u32>, Diagnostic> {
    let mut colors = HashMap::new();
    let mut depth = 0_usize;
    let mut container = None;
    let mut scheme = None;
    let mut slot = None::<(usize, String)>;
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "themeElements" || (depth == 0 && local == "themeOverride") {
                    container = (!empty).then_some(depth);
                } else if local == "clrScheme" && container.is_some_and(|start| depth == start + 1)
                {
                    scheme = (!empty).then_some(depth);
                } else if scheme.is_some()
                    && matches!(local, "dk1" | "lt1" | "dk2" | "lt2" | "accent1" | "accent2" | "accent3" | "accent4" | "accent5" | "accent6" | "hlink" | "folHlink")
                {
                    slot = (!empty).then(|| (depth, local.to_owned()));
                } else if let Some((_, key)) = &slot {
                    let attribute = match local {
                        "srgbClr" => Some("val"),
                        "sysClr" => Some("lastClr"),
                        _ => None,
                    };
                    if let Some(attribute) = attribute
                        && let Some(color) = string_attribute(&attributes, attribute, part)?
                            .as_deref()
                            .and_then(parse_rgb_color)
                    {
                        colors.insert(key.clone(), color);
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            XmlEvent::EndElement { .. } => {
                depth = depth.saturating_sub(1);
                if slot.as_ref().is_some_and(|(start, _)| *start == depth) {
                    slot = None;
                }
                if scheme == Some(depth) {
                    scheme = None;
                }
                if container == Some(depth) {
                    container = None;
                }
            }
            _ => {}
        }
        Ok(())
    })?;
    Ok(colors)
}

pub(super) fn parse_chart(
    package: &Package<'_>,
    part: &str,
    scheme_color: impl Fn(&str) -> Option<u32>,
) -> Result<Option<Chart>, Diagnostic> {
    parse_chart_impl(package, part, None, &scheme_color)
}

pub(super) fn parse_chart_with_theme(
    package: &Package<'_>,
    part: &str,
    theme_part: Option<&str>,
    scheme_color: impl Fn(&str) -> Option<u32>,
) -> Result<Option<Chart>, Diagnostic> {
    parse_chart_impl(package, part, theme_part, &scheme_color)
}

// Keep the large parser shared across host-specific theme callbacks.
#[inline(never)]
fn parse_chart_impl(
    package: &Package<'_>,
    part: &str,
    theme_part: Option<&str>,
    scheme_color: &dyn Fn(&str) -> Option<u32>,
) -> Result<Option<Chart>, Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut theme_colors = HashMap::new();
    for relationship in package
        .relationships(Some(part))?
        .into_iter()
        .filter(|r| !r.external && r.type_uri.ends_with("/themeOverride"))
    {
        theme_colors.extend(drawingml_theme_colors(
            &package.required_part(&relationship.target)?,
            package.limits(),
            &relationship.target,
        )?);
    }
    // Chart-local mappings take precedence over the host slide/document mapping.
    let mut color_map = HashMap::new();
    let mut mapping_depth = 0_usize;
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                if mapping_depth == 1 && local_name(name) == "clrMapOvr" {
                    for key in [
                        "bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4",
                        "accent5", "accent6", "hlink", "folHlink",
                    ] {
                        if let Some(value) = string_attribute(&attributes, key, part)? {
                            color_map.insert(key, value);
                        }
                    }
                }
                if !empty {
                    mapping_depth += 1;
                }
            }
            XmlEvent::EndElement { .. } => mapping_depth = mapping_depth.saturating_sub(1),
            _ => {}
        }
        Ok(())
    })?;
    let scheme_color = |value: &str| {
        let slot = color_map.get(value).map_or(value, String::as_str);
        theme_colors
            .get(slot)
            .copied()
            .or_else(|| scheme_color(slot))
    };
    let mut depth = 0_usize;
    let mut kind = None;
    let mut chart_depth = None;
    let mut chart_three_d = false;
    let mut chart_grouping = ChartGrouping::Standard;
    let mut chart_bar_horizontal = false;
    let mut chart_bar_depth = false;
    let mut chart_bar_cone = false;
    let mut chart_bar_cylinder = false;
    let mut chart_vary_colors = false;
    let mut chart_first_slice_angle = 0.0_f32;
    let mut scatter_has_lines = false;
    let mut scatter_has_markers = false;
    let mut scatter_smooth = false;
    let mut line_has_markers = false;
    let mut explicit_marker_series = HashSet::<usize>::new();
    let mut series_marker_is_explicit = false;
    let mut series_lines = false;
    let mut bar_gap_width_percent = 150.0_f32;
    let mut chart_bar_gap_depth_percent = 150.0_f32;
    let mut chart_pie_of_pie = false;
    let mut radar_filled = false;
    let mut filled_radars = Vec::new();
    let mut chart_bar_of_pie_split = 0_u16;
    let mut chart_style = None;
    let mut default_series_effects = None;
    let mut surface_band_depth = None;
    let mut surface_band_index = 0;
    let mut surface_band_properties = None;
    let mut surface_band_capture = None::<ChartFillCapture>;
    let mut surface_band_fills = Vec::new();
    let mut chart_style_is_extended = false;
    let mut chart_series_start = 0_usize;
    let mut chart_axis_ids = Vec::new();
    let mut series_depth = None;
    let mut source_series_index = 0_usize;
    let mut source_series_indices = Vec::new();
    let mut series = None::<ChartSeries>;
    let mut value_target: Option<(usize, ChartValueTarget)> = None;
    let mut collecting_value: Option<(usize, String)> = None;
    let mut cache_point_index = None::<usize>;
    let mut category_level_depth = None;
    let mut shape_properties_depth = None;
    let mut shape_fill_depth = None;
    let mut shape_fill_capture = None::<ChartFillCapture>;
    let mut shape_line_depth = None;
    let mut line_fill_depth = None;
    let mut completed = Vec::new();
    let mut show_title = false;
    let mut title_deleted = false;
    let mut title_depth = None;
    let mut title_layout_depth = None;
    let mut title_coordinates = [None; 2];
    let mut title_edge = [false; 2];
    let mut title_position = None;
    let mut title_text_depth = None;
    let mut title = String::new();
    let mut title_all_caps = false;
    let mut title_properties_depth = None;
    let mut title_fill_capture = None::<ChartFillCapture>;
    let mut title_text_fill_capture = None::<ChartFillCapture>;
    let mut title_text_color = None;
    let mut title_fill = None;
    let mut title_font_size = None;
    let mut title_font_bold = None;
    let mut plot_bounds = None;
    let mut plot_area_depth = None;
    let mut plot_area_properties_depth = None;
    let mut plot_area_fill_capture = None::<ChartFillCapture>;
    let mut plot_area_fill = None;
    let mut plot_layout_depth = None;
    let mut manual_plot_layout_depth = None;
    let mut manual_plot_layout = [None::<f32>; 4];
    let mut show_legend = false;
    let mut legend_position = ChartLegendPosition::default();
    let mut chart_space_depth = None;
    let mut chart_area_properties_depth = None;
    let mut chart_area_no_fill = false;
    let mut chart_area_fill = None;
    let mut chart_area_border = None;
    let mut chart_area_line_depth = None;
    let mut chart_area_line_width = 1.0;
    let mut chart_area_fill_capture = None::<ChartFillCapture>;
    let mut chart_text_properties_depth = None;
    let mut chart_font_family = None;
    let mut chart_font_size = None;
    let mut chart_font_bold = false;
    let mut legend_depth = None;
    let mut legend_layout_depth = None;
    let mut manual_legend_layout_depth = None;
    let mut manual_legend_layout = [None::<f32>; 4];
    let mut manual_legend_size_edge = [false; 2];
    let mut legend_bounds = None;
    let mut legend_entry_depth = None;
    let mut legend_entry_index = None::<usize>;
    let mut deleted_legend_entries = Vec::new();
    let mut legend_properties_depth = None;
    let mut legend_line_depth = None;
    let mut legend_fill = None;
    let mut legend_fill_capture = None::<ChartFillCapture>;
    let mut legend_stroke = None;
    let mut legend_stroke_width = 0.0;
    let mut legend_stroke_capture = None::<ChartFillCapture>;
    let mut value_axis_depth = None;
    let mut value_axis_scaling_depth = None;
    let mut value_axis_options = ChartValueAxis::default();
    let mut secondary_value_axis_options = None;
    let mut horizontal_axis_options = ChartValueAxis::default();
    let mut series_axis_options = None;
    let mut current_value_axis_options = None::<ChartValueAxis>;
    let mut display_unit_label_depth = None;
    let mut display_unit_text_depth = None;
    let mut axis_title_depth = None;
    let mut axis_title_text_depth = None;
    let mut axis_title = String::new();
    let mut axis_title_properties_depth = None;
    let mut axis_title_line_depth = None::<usize>;
    let mut axis_title_body_depth = None::<usize>;
    let mut axis_title_fill_capture = None::<ChartFillCapture>;
    let mut axis_title_stroke_capture = None::<ChartFillCapture>;
    let mut axis_title_effects_capture = DrawingMlPictureEffectsCapture::default();
    let mut axis_text_properties_depth = None;
    let mut axis_properties_depth = None;
    let mut axis_line_depth = None;
    let mut gridlines_depth = None;
    let mut marker_depth = None;
    let mut trendline_depth = None;
    let mut trendline_label_depth = None;
    let mut trendline_label_layout_depth = None;
    let mut manual_trendline_label_layout_depth = None;
    let mut trendline_label_offset = [None::<f32>; 2];
    let mut trendline_label_text_properties_depth = None;
    let mut error_bars_depth = None;
    let mut error_bars_are_vertical = true;
    let mut error_bars = ChartErrorBars::default();
    let mut error_bar_line_depth = None;
    let mut error_bar_stroke_capture = None::<ChartFillCapture>;
    let mut data_table_depth = None;
    let mut data_table = None::<ChartDataTable>;
    let mut up_down_bars_depth = None;
    let mut up_down_bars = None::<ChartUpDownBars>;
    let mut up_down_bar_depth = None;
    let mut up_down_bar_is_up = None::<bool>;
    let mut up_down_bar_properties_depth = None;
    let mut up_down_bar_line_depth = None;
    let mut up_down_bar_fill_capture = None::<ChartFillCapture>;
    let mut up_down_bar_stroke_capture = None::<ChartFillCapture>;
    let mut date_1904 = false;
    let mut category_format_code = None::<String>;
    let mut collecting_format_code = None::<(usize, String)>;
    let mut view_3d = None::<ChartView3D>;
    let mut right_angle_axes_explicit = false;
    let mut data_label_depth = None::<usize>;
    let mut data_labels_depth = None::<usize>;
    let mut group_data_labels_start = None::<usize>;
    let mut explicit_label_flags = HashSet::<(usize, String)>::new();
    let mut series_label_border_width = 1.0;
    let mut deleted_series_data_labels = HashSet::<usize>::new();
    let mut data_labels_text_properties_depth = None::<usize>;
    let mut data_label_index = None::<usize>;
    let mut data_label = None::<ChartDataLabelCapture>;
    let mut data_label_text_depth = None::<usize>;
    let mut data_label_field = None::<String>;
    let mut data_label_fields = Vec::<(usize, usize, std::ops::Range<usize>, String)>::new();
    let mut collecting_data_label_text = None::<(usize, String)>;
    let mut data_label_shape_properties_depth = None::<usize>;
    let mut data_label_line_depth = None::<usize>;
    let mut data_label_line_fill_depth = None::<usize>;
    let mut data_label_manual_layout_depth = None::<usize>;
    let mut data_point_depth = None::<usize>;
    let mut data_point_index = None::<usize>;
    let mut data_point_color = None::<u32>;
    let mut data_point_fill = None::<ChartFill>;
    let mut data_point_explosion = None::<f32>;
    let mut series_border_color = None;
    let mut series_effects = DrawingMlPictureEffectsCapture::default();
    let mut series_effects_explicit = false;
    let mut data_point_border_color = None::<u32>;
    let mut data_point_border_width = None::<f32>;
    let mut series_explosion = None::<f32>;
    let mut data_point_color_overrides = Vec::<(usize, u32)>::new();
    let mut data_point_fill_overrides = Vec::<(usize, ChartFill)>::new();
    let mut data_point_explosion_overrides = Vec::<(usize, f32)>::new();
    let mut data_point_border_overrides = Vec::<(usize, Option<u32>, f32)>::new();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                match local {
                    "chartSpace" if chart_space_depth.is_none() => {
                        chart_space_depth = Some(depth);
                    }
                    "spPr"
                        if series.is_none()
                            && chart_space_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        chart_area_properties_depth = (!empty).then_some(depth);
                    }
                    "txPr"
                        if series.is_none()
                            && chart_space_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        chart_text_properties_depth = (!empty).then_some(depth);
                    }
                    "defRPr" | "rPr" if chart_text_properties_depth.is_some() => {
                        if let Some(size) = numeric_attribute(&attributes, "sz", part)? {
                            chart_font_size = Some(size as f32 / 100.0 * 96.0 / 72.0);
                        }
                        if let Some(bold) = string_attribute(&attributes, "b", part)? {
                            chart_font_bold = bold == "1" || bold.eq_ignore_ascii_case("true");
                        }
                    }
                    "latin" if chart_text_properties_depth.is_some() => {
                        chart_font_family = string_attribute(&attributes, "typeface", part)?;
                    }
                    "p" if axis_title_depth.is_none()
                        && series.is_none()
                        && title_depth.is_some_and(|start| depth == start + 3) =>
                    {
                        if !title.is_empty() && !title.ends_with('\n') {
                            title.push('\n');
                        }
                    }
                    "defRPr" | "rPr"
                        if title_depth.is_some()
                            && axis_title_depth.is_none()
                            && series.is_none() =>
                    {
                        if let Some(size) = numeric_attribute(&attributes, "sz", part)? {
                            title_font_size = Some(size as f32 / 100.0 * 96.0 / 72.0);
                        }
                        if let Some(bold) = string_attribute(&attributes, "b", part)? {
                            title_font_bold =
                                Some(bold == "1" || bold.eq_ignore_ascii_case("true"));
                        }
                        title_all_caps |=
                            string_attribute(&attributes, "cap", part)?.as_deref() == Some("all");
                    }
                    "solidFill" if title_depth.is_some() && axis_title_depth.is_none()
                        && title_properties_depth.is_none() && series.is_none() => {
                        title_text_fill_capture = ChartFillCapture::new(local, depth);
                    }
                    "noFill"
                        if chart_area_properties_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        chart_area_no_fill = true;
                    }
                    "ln" if chart_area_properties_depth.is_some_and(|start| depth == start + 1) => {
                        chart_area_line_depth = (!empty).then_some(depth);
                        chart_area_line_width = numeric_attribute(&attributes, "w", part)?
                            .map_or(1.0, |width| width as f32 / EMU_PER_CSS_PIXEL);
                    }
                    "noFill" if chart_area_line_depth.is_some_and(|start| depth == start + 1) => {
                        chart_area_border = Some((ChartFill::Solid(0), 0.0));
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if chart_area_properties_depth.is_some_and(|start| depth == start + 1)
                            || chart_area_line_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        chart_area_fill_capture = ChartFillCapture::new(local, depth);
                    }
                    "view3D" if series.is_none() => {
                        view_3d.get_or_insert_with(ChartView3D::default);
                    }
                    "rotX" if series.is_none() => {
                        if let Some(value) = string_attribute(&attributes, "val", part)? {
                            view_3d.get_or_insert_with(ChartView3D::default).rot_x = value
                                .parse::<i16>()
                                .map_err(|_| format_error(part, "chart rotX is not an integer"))?;
                        }
                    }
                    "rotY" if series.is_none() => {
                        if let Some(value) = numeric_attribute(&attributes, "val", part)? {
                            view_3d.get_or_insert_with(ChartView3D::default).rot_y =
                                u16::try_from(value).map_err(|_| {
                                    format_error(part, "chart rotY exceeds the supported range")
                                })?;
                        }
                    }
                    "rAngAx" if series.is_none() => {
                        right_angle_axes_explicit = true;
                        view_3d.get_or_insert_with(ChartView3D::default).right_angle_axes =
                            string_attribute(&attributes, "val", part)?.is_none_or(|value|
                                value == "1" || value.eq_ignore_ascii_case("true"));
                    }
                    "perspective" if series.is_none() => {
                        if let Some(value) = numeric_attribute(&attributes, "val", part)? {
                            view_3d.get_or_insert_with(ChartView3D::default).perspective =
                                u16::try_from(value).map_err(|_| {
                                    format_error(
                                        part,
                                        "chart perspective exceeds the supported range",
                                    )
                                })?;
                        }
                    }
                    "hPercent" if series.is_none() => {
                        if let Some(value) = numeric_attribute(&attributes, "val", part)? {
                            view_3d
                                .get_or_insert_with(ChartView3D::default)
                                .height_percent = Some(u16::try_from(value).map_err(|_| {
                                format_error(part, "chart hPercent exceeds the supported range")
                            })?);
                        }
                    }
                    "depthPercent" if series.is_none() => {
                        if let Some(value) = numeric_attribute(&attributes, "val", part)? {
                            view_3d
                                .get_or_insert_with(ChartView3D::default)
                                .depth_percent = Some(u16::try_from(value).map_err(|_| {
                                format_error(part, "chart depthPercent exceeds the supported range")
                            })?);
                        }
                    }
                    "style" if chart_depth.is_none() && series.is_none() => {
                        if name.starts_with("c14:") {
                            chart_style =
                                numeric_attribute(&attributes, "val", part)?.map(|style| {
                                    if (101..=148).contains(&style) {
                                        style - 100
                                    } else {
                                        style
                                    }
                                });
                            chart_style_is_extended = true;
                        } else if !chart_style_is_extended {
                            chart_style = numeric_attribute(&attributes, "val", part)?;
                        }
                    }
                    "bandFmt" => {
                        surface_band_depth = (!empty).then_some(depth);
                        surface_band_index = 0;
                    }
                    "idx" if surface_band_depth.is_some() => {
                        surface_band_index =
                            numeric_attribute(&attributes, "val", part)?.unwrap_or(0) as usize;
                        if surface_band_index >= package.limits().max_document_objects {
                            return Err(format_error(
                                part,
                                "surface band index exceeds the object limit",
                            ));
                        }
                    }
                    "spPr" if surface_band_depth.is_some_and(|d| depth == d + 1) => {
                        surface_band_properties = (!empty).then_some(depth);
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if surface_band_properties.is_some_and(|d| depth == d + 1) =>
                    {
                        surface_band_capture = ChartFillCapture::new(local, depth);
                    }
                    "date1904" => {
                        date_1904 =
                            string_attribute(&attributes, "val", part)?.is_some_and(|value| {
                                value == "1" || value.eq_ignore_ascii_case("true")
                            });
                    }
                    _ if chart_kind(local).is_some() => {
                        chart_three_d = local.contains("3D");
                        chart_grouping = ChartGrouping::Standard;
                        chart_bar_horizontal = false;
                        chart_bar_depth = false;
                        chart_bar_cone = false;
                        chart_bar_cylinder = false;
                        chart_vary_colors = false;
                        chart_first_slice_angle = 0.0;
                        scatter_has_lines = false;
                        scatter_has_markers = local == "scatterChart";
                        scatter_smooth = false;
                        line_has_markers = local == "lineChart" && chart_style_is_extended;
                        explicit_marker_series.clear();
                        chart_bar_of_pie_split = 0;
                        chart_pie_of_pie = false;
                        radar_filled = false;
                        begin_chart(
                            chart_kind(local).expect("matched chart element"),
                            empty,
                            depth,
                            completed.len(),
                            &mut kind,
                            &mut chart_depth,
                            &mut chart_series_start,
                            &mut chart_axis_ids,
                        );
                    }
                    "varyColors" if chart_depth.is_some() && series.is_none() => {
                        chart_vary_colors = string_attribute(&attributes, "val", part)?
                            .is_some_and(|value| {
                                value == "1" || value.eq_ignore_ascii_case("true")
                            });
                    }
                    "radarStyle" if chart_depth.is_some() && series.is_none() => {
                        radar_filled = string_attribute(&attributes, "val", part)?.as_deref() == Some("filled");
                    }
                    "scatterStyle" if chart_depth.is_some() && series.is_none() => {
                        if let Some(style) = string_attribute(&attributes, "val", part)? {
                            let style = style.to_ascii_lowercase();
                            scatter_smooth = style.contains("smooth");
                            scatter_has_lines = style.contains("line") || scatter_smooth;
                            scatter_has_markers = style.contains("marker");
                        }
                    }
                    "firstSliceAng" if chart_depth.is_some() && series.is_none() => {
                        let value = numeric_attribute(&attributes, "val", part)?.unwrap_or(0);
                        if value > 360 {
                            return Err(format_error(
                                part,
                                "chart firstSliceAng exceeds 360 degrees",
                            ));
                        }
                        chart_first_slice_angle = value as f32;
                    }
                    "ofPieType" if chart_depth.is_some() && series.is_none() => {
                        chart_pie_of_pie =
                            string_attribute(&attributes, "val", part)?.as_deref() == Some("pie");
                    }
                    "wireframe" if chart_depth.is_some() && series.is_none() => {
                        if kind == Some(ChartKind::Surface)
                            && string_attribute(&attributes, "val", part)?.as_deref() != Some("0")
                        {
                            kind = Some(ChartKind::SurfaceWireframe);
                        }
                    }
                    "splitPos" if chart_depth.is_some() && series.is_none() => {
                        chart_bar_of_pie_split = numeric_attribute(&attributes, "val", part)?
                            .and_then(|value| u16::try_from(value).ok())
                            .unwrap_or(0);
                    }
                    "grouping" if chart_depth.is_some() && series.is_none() => {
                        chart_bar_depth = chart_three_d && kind == Some(ChartKind::Bar)
                            && string_attribute(&attributes, "val", part)?.as_deref() == Some("standard");
                        chart_grouping =
                            match string_attribute(&attributes, "val", part)?.as_deref() {
                                Some("stacked") => ChartGrouping::Stacked,
                                Some("percentStacked") => ChartGrouping::PercentStacked,
                                _ => ChartGrouping::Standard,
                            };
                    }
                    "barDir" if chart_depth.is_some() && series.is_none() => {
                        chart_bar_horizontal = string_attribute(&attributes, "val", part)?
                            .is_some_and(|value| value == "bar");
                    }
                    "shape" if chart_depth.is_some() && series.is_none() => {
                        if let Some(value) = string_attribute(&attributes, "val", part)? {
                            chart_bar_cone = value.starts_with("cone");
                            chart_bar_cylinder = value == "cylinder";
                        }
                    }
                    "shape" if chart_depth.is_some() => {
                        if let (Some(series), Some(value)) =
                            (series.as_mut(), string_attribute(&attributes, "val", part)?)
                        {
                            series.bar_cone = value.starts_with("cone");
                            series.bar_cone_to_max = value.ends_with("ToMax");
                            series.bar_cylinder = value == "cylinder";
                        }
                    }
                    "title"
                        if value_axis_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        axis_title = if empty { "Axis Title".to_owned() } else { String::new() };
                        axis_title_effects_capture = DrawingMlPictureEffectsCapture::default();
                        axis_title_depth = (!empty).then_some(depth);
                    }
                    "spPr" if axis_title_depth.is_some_and(|start| depth == start + 1) => {
                        axis_title_properties_depth = (!empty).then_some(depth);
                    }
                    "ln" if axis_title_properties_depth.is_some_and(|start| depth == start + 1) => {
                        axis_title_line_depth = (!empty).then_some(depth);
                        if let Some(options) = current_value_axis_options.as_mut() {
                            (options.title_stroke_width, options.title_stroke_style) =
                                drawingml_stroke_style(&attributes, part)?;
                        }
                    }
                    "prstDash" if axis_title_line_depth.is_some() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.title_stroke_style.dash = DrawingMlDashPattern::from_attribute(
                                string_attribute(&attributes, "val", part)?.as_deref(),
                            )
                            .lengths()
                            .iter()
                            .map(|length| length * options.title_stroke_width.max(1.0))
                            .collect();
                        }
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if axis_title_line_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        axis_title_stroke_capture = ChartFillCapture::new(local, depth);
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if axis_title_properties_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        axis_title_fill_capture = ChartFillCapture::new(local, depth);
                    }
                    "rich" if axis_title_depth.is_some() => {
                        axis_title_body_depth = (!empty).then_some(depth);
                    }
                    "bodyPr" if axis_title_body_depth.is_some() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.title_rotation_degrees =
                                signed_numeric_attribute(&attributes, "rot", part)?
                                    .map(|value| value as f32 / 60_000.0)
                                    .filter(|rotation| rotation.abs() <= 90.0);
                            options.title_orientation =
                                string_attribute(&attributes, "vert", part)?
                                    .as_deref()
                                    .map(drawingml_text_orientation);
                        }
                    }
                    "defRPr" | "rPr" if axis_title_depth.is_some() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            if let Some(size) = numeric_attribute(&attributes, "sz", part)? {
                                options.title_font_size = Some(size as f32 / 100.0 * 96.0 / 72.0);
                            }
                            if let Some(bold) = string_attribute(&attributes, "b", part)? {
                                options.title_bold =
                                    Some(bold == "1" || bold.eq_ignore_ascii_case("true"));
                            }
                        }
                    }
                    "srgbClr" | "schemeClr" if axis_text_properties_depth.is_some() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.label_color = string_attribute(&attributes, "val", part)?.as_deref().and_then(|value| if local == "srgbClr" { parse_rgb_color(value) } else { scheme_color(value) });
                        }
                    }
                    "defRPr" | "rPr" if axis_text_properties_depth.is_some() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            if let Some(size) = numeric_attribute(&attributes, "sz", part)? {
                                options.label_font_size = Some(size as f32 / 100.0 * 96.0 / 72.0);
                            }
                            if let Some(bold) = string_attribute(&attributes, "b", part)? {
                                options.label_bold =
                                    Some(bold == "1" || bold.eq_ignore_ascii_case("true"));
                            }
                        }
                    }
                    "title" if series.is_none() && value_axis_depth.is_none() => {
                        show_title = true;
                        title_depth = (!empty).then_some(depth);
                    }
                    "autoTitleDeleted" if series.is_none() && value_axis_depth.is_none() => {
                        title_deleted = string_attribute(&attributes, "val", part)?
                            .is_none_or(|value| value == "1" || value.eq_ignore_ascii_case("true"));
                    }
                    "manualLayout" if title_depth.is_some() && axis_title_depth.is_none() => {
                        title_layout_depth = (!empty).then_some(depth);
                    }
                    "xMode" | "yMode" if title_layout_depth.is_some() => {
                        title_edge[usize::from(local == "yMode")] = string_attribute(&attributes, "val", part)?.as_deref() == Some("edge");
                    }
                    "x" | "y" if title_layout_depth.is_some() => {
                        title_coordinates[usize::from(local == "y")] = float_attribute(&attributes, "val", part)?;
                    }
                    "plotArea" if series.is_none() => {
                        plot_area_depth = (!empty).then_some(depth);
                    }
                    "spPr" if plot_area_depth.is_some_and(|start| depth == start + 1) => {
                        plot_area_properties_depth = (!empty).then_some(depth);
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if plot_area_properties_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        plot_area_fill_capture = ChartFillCapture::new(local, depth);
                    }
                    "layout"
                        if plot_area_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        plot_layout_depth = (!empty).then_some(depth);
                    }
                    "manualLayout"
                        if plot_layout_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        manual_plot_layout_depth = (!empty).then_some(depth);
                        manual_plot_layout = [None; 4];
                    }
                    "x" | "y" | "w" | "h" if manual_plot_layout_depth.is_some() => {
                        let slot = match local {
                            "x" => 0,
                            "y" => 1,
                            "w" => 2,
                            _ => 3,
                        };
                        manual_plot_layout[slot] = float_attribute(&attributes, "val", part)?;
                    }
                    "valAx" | "catAx" | "dateAx" | "serAx" if series.is_none() => {
                        value_axis_depth = (!empty).then_some(depth);
                        current_value_axis_options = Some(ChartValueAxis {
                            is_date: local == "dateAx",
                            ..ChartValueAxis::default()
                        });
                    }
                    "txPr"
                        if value_axis_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        axis_text_properties_depth = (!empty).then_some(depth);
                    }
                    "spPr"
                        if value_axis_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        axis_properties_depth = (!empty).then_some(depth);
                    }
                    "ln" if gridlines_depth.is_some_and(|(start, _)| depth == start + 2) => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            let minor = gridlines_depth.unwrap().1;
                            let (enabled, _, width) = options.gridline_properties_mut(minor);
                            *width = numeric_attribute(&attributes, "w", part)?.map(|width| width as f32 / EMU_PER_CSS_PIXEL);
                            if *width == Some(0.0) {
                                if minor { *width = None; } else { *enabled = false; }
                            }
                        }
                    }
                    "srgbClr" | "sysClr" | "schemeClr"
                        if gridlines_depth.is_some_and(|(start, _)| depth == start + 4) =>
                    {
                        let color = if local == "schemeClr" {
                            string_attribute(&attributes, "val", part)?
                                .as_deref()
                                .and_then(&scheme_color)
                        } else {
                            string_attribute(
                                &attributes,
                                if local == "sysClr" { "lastClr" } else { "val" },
                                part,
                            )?
                            .as_deref()
                            .and_then(parse_rgb_color)
                        };
                        if let Some(options) = current_value_axis_options.as_mut() {
                            *options.gridline_properties_mut(gridlines_depth.unwrap().1).1 = color;
                        }
                    }
                    "noFill" if gridlines_depth.is_some_and(|(start, _)| depth == start + 3) => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            *options.gridline_properties_mut(gridlines_depth.unwrap().1).0 = false;
                        }
                    }
                    "ln" if axis_properties_depth.is_some_and(|start| depth == start + 1) => {
                        axis_line_depth = (!empty).then_some(depth);
                    }
                    "noFill" if axis_line_depth.is_some_and(|start| depth == start + 1) => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.axis_line_hidden = true;
                        }
                    }
                    "bodyPr" if axis_text_properties_depth.is_some() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.label_rotation_degrees =
                                signed_numeric_attribute(&attributes, "rot", part)?
                                    .map(|value| value as f32 / 60_000.0)
                                    .filter(|rotation| rotation.abs() <= 90.0);
                        }
                    }
                    "bodyPr" if data_labels_text_properties_depth.is_some() => {
                        if let Some(series) = series.as_mut() {
                            series.data_label_rotation_degrees =
                                signed_numeric_attribute(&attributes, "rot", part)?
                                    .map(|value| value as f32 / 60_000.0)
                                    .filter(|rotation| rotation.abs() <= 90.0);
                        }
                    }
                    "scaling"
                        if value_axis_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        value_axis_scaling_depth = (!empty).then_some(depth);
                    }
                    "min" | "max" if value_axis_scaling_depth.is_some() => {
                        let value = float_attribute(&attributes, "val", part)?;
                        if local == "min" {
                            if let Some(options) = current_value_axis_options.as_mut() {
                                options.minimum = value;
                            }
                        } else {
                            if let Some(options) = current_value_axis_options.as_mut() {
                                options.maximum = value;
                            }
                        }
                    }
                    "orientation" if value_axis_scaling_depth.is_some() => {
                        if let Some(axis) = current_value_axis_options.as_mut() {
                            axis.reversed = string_attribute(&attributes, "val", part)?.as_deref() == Some("maxMin");
                        }
                    }
                    "logBase" if value_axis_scaling_depth.is_some() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.log_base = float_attribute(&attributes, "val", part)?
                                .filter(|value| (2.0..=1_000.0).contains(value));
                        }
                    }
                    "noMultiLvlLbl" if value_axis_depth.is_some() && series.is_none() => {
                        if let Some(axis) = current_value_axis_options.as_mut() {
                            axis.no_multi_level_labels = string_attribute(&attributes, "val", part)?.as_deref().is_none_or(|v| matches!(v, "1" | "true"));
                        }
                    }
                    "dispUnitsLbl" if value_axis_depth.is_some() && series.is_none() => {
                        if let Some(axis) = current_value_axis_options.as_mut() {
                            axis.display_unit_label = Some(String::new());
                        }
                        display_unit_label_depth = (!empty).then_some(depth);
                    }
                    "t" if display_unit_label_depth.is_some() => {
                        display_unit_text_depth = (!empty).then_some(depth);
                    }
                    "builtInUnit" | "custUnit" if value_axis_depth.is_some() && series.is_none() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.display_unit = if local == "custUnit" {
                                float_attribute(&attributes, "val", part)?.filter(|value| *value > 0.0)
                            } else {
                                match string_attribute(&attributes, "val", part)?.as_deref() {
                                    Some("hundreds") => Some(100.0), Some("thousands") => Some(1e3),
                                    Some("tenThousands") => Some(1e4), Some("hundredThousands") => Some(1e5),
                                    Some("millions") => Some(1e6), Some("tenMillions") => Some(1e7),
                                    Some("hundredMillions") => Some(1e8), Some("billions") => Some(1e9),
                                    Some("trillions") => Some(1e12), _ => None,
                                }
                            };
                        }
                    }
                    "majorUnit" | "minorUnit" if value_axis_depth.is_some() && series.is_none() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            let value = float_attribute(&attributes, "val", part)?.filter(|value| *value > 0.0);
                            if local == "minorUnit" { options.minor_unit = value; } else { options.major_unit = value; }
                        }
                    }
                    "baseTimeUnit" | "majorTimeUnit" | "minorTimeUnit"
                        if value_axis_depth.is_some() && series.is_none() =>
                    {
                        if let Some(axis) = current_value_axis_options.as_mut().filter(|axis| axis.is_date) {
                            let unit = string_attribute(&attributes, "val", part)?
                                .as_deref().and_then(ChartTimeUnit::from_ooxml);
                            match local {
                                "baseTimeUnit" => axis.base_time_unit = unit,
                                "majorTimeUnit" => axis.major_time_unit = unit,
                                _ => axis.minor_time_unit = unit,
                            }
                        }
                    }
                    "tickLblSkip" | "tickMarkSkip"
                        if value_axis_depth.is_some() && series.is_none() =>
                    {
                        let skip = numeric_attribute(&attributes, "val", part)?
                            .and_then(|value| usize::try_from(value).ok())
                            .filter(|value| *value > 0);
                        if let Some(options) = current_value_axis_options.as_mut() {
                            if local == "tickLblSkip" {
                                options.label_skip = skip;
                            } else {
                                options.tick_skip = skip;
                            }
                        }
                    }
                    "majorTickMark" | "minorTickMark" if value_axis_depth.is_some() && series.is_none() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            let mark = ChartAxisTickMark::from_ooxml(string_attribute(&attributes, "val", part)?.as_deref());
                            if local == "minorTickMark" { options.minor_tick_mark = mark; } else { options.major_tick_mark = mark; }
                        }
                    }
                    "majorGridlines" | "minorGridlines"
                        if value_axis_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            let minor = local == "minorGridlines";
                            *options.gridline_properties_mut(minor).0 = true;
                            gridlines_depth = (!empty).then_some((depth, minor));
                        }
                    }
                    "delete" if value_axis_depth.is_some() && series.is_none() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.deleted = string_attribute(&attributes, "val", part)?
                                .is_some_and(|value| {
                                    value == "1" || value.eq_ignore_ascii_case("true")
                                });
                        }
                    }
                    "axId"
                        if value_axis_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.id = string_attribute(&attributes, "val", part)?;
                        }
                    }
                    "crossBetween" if value_axis_depth.is_some() && series.is_none() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.cross_between_categories =
                                string_attribute(&attributes, "val", part)?.as_deref()
                                    == Some("between");
                        }
                    }
                    "numFmt" if value_axis_depth.is_some() && series.is_none() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.number_format =
                                string_attribute(&attributes, "formatCode", part)?;
                        }
                    }
                    "axPos" if value_axis_depth.is_some() && series.is_none() => {
                        if let Some(options) = current_value_axis_options.as_mut() {
                            options.position = string_attribute(&attributes, "val", part)?;
                        }
                    }
                    "dTable" if series.is_none() => {
                        data_table_depth = (!empty).then_some(depth);
                        data_table = Some(ChartDataTable::default());
                    }
                    "serLines" if chart_depth.is_some() && series.is_none() => {
                        series_lines = true;
                    }
                    "upDownBars" if chart_depth.is_some() && series.is_none() => {
                        up_down_bars_depth = (!empty).then_some(depth);
                        up_down_bars = Some(ChartUpDownBars {
                            series_start: chart_series_start,
                            ..ChartUpDownBars::default()
                        });
                    }
                    "gapWidth" if up_down_bars_depth.is_some() => {
                        if let Some(value) = numeric_attribute(&attributes, "val", part)? {
                            if value > 500 {
                                return Err(format_error(
                                    part,
                                    "chart up/down bar gapWidth exceeds 500 percent",
                                ));
                            }
                            if let Some(options) = up_down_bars.as_mut() {
                                options.gap_width = value as u16;
                            }
                        }
                    }
                    "gapWidth" if chart_depth.is_some() && series.is_none() => {
                        if let Some(value) = float_attribute(&attributes, "val", part)?
                            .filter(|value| (0.0..=500.0).contains(value))
                        {
                            bar_gap_width_percent = value;
                        }
                    }
                    "gapDepth" if chart_depth.is_some() && series.is_none() => {
                        if let Some(value) = float_attribute(&attributes, "val", part)?
                            .filter(|value| (0.0..=500.0).contains(value))
                        {
                            chart_bar_gap_depth_percent = value;
                        }
                    }
                    "upBars" | "downBars"
                        if up_down_bars_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        up_down_bar_depth = (!empty).then_some(depth);
                        up_down_bar_is_up = (!empty).then_some(local == "upBars");
                    }
                    "spPr" if up_down_bar_depth.is_some_and(|start| depth == start + 1) => {
                        up_down_bar_properties_depth = (!empty).then_some(depth);
                    }
                    "ln" if up_down_bar_properties_depth
                        .is_some_and(|start| depth == start + 1) =>
                    {
                        up_down_bar_line_depth = (!empty).then_some(depth);
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if up_down_bar_line_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        up_down_bar_stroke_capture = ChartFillCapture::new(local, depth);
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if up_down_bar_properties_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        up_down_bar_fill_capture = ChartFillCapture::new(local, depth);
                    }
                    "showHorzBorder" | "showVertBorder" | "showOutline" | "showKeys"
                        if data_table_depth.is_some() =>
                    {
                        let shown =
                            string_attribute(&attributes, "val", part)?.is_some_and(|value| {
                                value == "1" || value.eq_ignore_ascii_case("true")
                            });
                        if let Some(table) = data_table.as_mut() {
                            match local {
                                "showHorzBorder" => table.show_horizontal_borders = shown,
                                "showVertBorder" => table.show_vertical_borders = shown,
                                "showOutline" => table.show_outline = shown,
                                _ => table.show_keys = shown,
                            }
                        }
                    }
                    "spPr"
                        if title_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        title_properties_depth = (!empty).then_some(depth);
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if title_properties_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        title_fill_capture = ChartFillCapture::new(local, depth);
                    }
                    "spPr"
                        if legend_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        legend_properties_depth = (!empty).then_some(depth);
                    }
                    "ln" if legend_properties_depth.is_some_and(|start| depth == start + 1) => {
                        legend_line_depth = (!empty).then_some(depth);
                        legend_stroke_width = numeric_attribute(&attributes, "w", part)?
                            .map_or(0.0, |width| width as f32 / EMU_PER_CSS_PIXEL);
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if legend_line_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        legend_stroke_capture = ChartFillCapture::new(local, depth);
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if legend_properties_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        legend_fill_capture = ChartFillCapture::new(local, depth);
                    }
                    "t" if axis_title_depth.is_some() && series.is_none() => {
                        axis_title_text_depth = (!empty).then_some(depth);
                    }
                    "t" if title_depth.is_some() && series.is_none() => {
                        title_text_depth = (!empty).then_some(depth);
                    }
                    "legend" if series.is_none() => {
                        show_legend = true;
                        legend_depth = (!empty).then_some(depth);
                    }
                    "layout"
                        if legend_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        legend_layout_depth = (!empty).then_some(depth);
                    }
                    "manualLayout"
                        if legend_layout_depth.is_some_and(|start| depth == start + 1)
                            && series.is_none() =>
                    {
                        manual_legend_layout_depth = (!empty).then_some(depth);
                        manual_legend_layout = [None; 4];
                        manual_legend_size_edge = [false; 2];
                    }
                    "wMode" | "hMode" if manual_legend_layout_depth.is_some() => {
                        let slot = if local == "wMode" { 0 } else { 1 };
                        manual_legend_size_edge[slot] =
                            string_attribute(&attributes, "val", part)?.as_deref() == Some("edge");
                    }
                    "x" | "y" | "w" | "h" if manual_legend_layout_depth.is_some() => {
                        let slot = match local {
                            "x" => 0,
                            "y" => 1,
                            "w" => 2,
                            _ => 3,
                        };
                        manual_legend_layout[slot] = float_attribute(&attributes, "val", part)?;
                    }
                    "legendEntry" if legend_depth.is_some() && series.is_none() => {
                        legend_entry_depth = (!empty).then_some(depth);
                        legend_entry_index = None;
                    }
                    "idx" if legend_entry_depth.is_some() => {
                        legend_entry_index = numeric_attribute(&attributes, "val", part)?
                            .and_then(|value| usize::try_from(value).ok());
                    }
                    "delete" if legend_entry_depth.is_some() => {
                        if string_attribute(&attributes, "val", part)?
                            .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"))
                            && let Some(index) = legend_entry_index
                        {
                            deleted_legend_entries.push(index);
                        }
                    }
                    "legendPos" if series.is_none() => {
                        legend_position = ChartLegendPosition::from_ooxml(
                            string_attribute(&attributes, "val", part)?.as_deref(),
                        );
                    }
                    "idx" if series_depth.is_some_and(|start| depth == start + 1) => {
                        source_series_index = numeric_attribute(&attributes, "val", part)?
                            .and_then(|value| usize::try_from(value).ok())
                            .unwrap_or(completed.len());
                    }
                    "ser" if kind.is_some() && series.is_none() => {
                        series_depth = (!empty).then_some(depth);
                        source_series_index = completed.len();
                        series_marker_is_explicit = false;
                        series_border_color = None;
                        series_effects = DrawingMlPictureEffectsCapture::default();
                        series_effects_explicit = false;
                        data_point_color_overrides.clear();
                        data_point_fill_overrides.clear();
                        data_point_explosion_overrides.clear();
                        data_point_border_overrides.clear();
                        series_explosion = None;
                        series = Some(ChartSeries {
                            kind: kind.unwrap_or(ChartKind::Line),
                            first_slice_angle: 0.0,
                            grouping: chart_grouping,
                            bar_horizontal: chart_bar_horizontal,
                            bar_depth: false,
                            bar_cone: chart_bar_cone,
                            bar_cone_to_max: false,
                            bar_cylinder: chart_bar_cylinder,
                            bar_gap_depth_percent: 150.0,
                            three_d: chart_three_d,
                            axis_id: None,
                            name: String::new(),
                            category_levels: Vec::new(),
                            categories: Vec::new(),
                            x_values: Vec::new(),
                            values: Vec::new(),
                            raw_values: Vec::new(),
                            bubble_sizes: Vec::new(),
                            color: None,
                            fill: None,
                            point_colors: Vec::new(),
                            point_fills: Vec::new(),
                            point_explosions: Vec::new(),
                            effects: Default::default(),
                            point_border_colors: Vec::new(),
                            point_border_widths: Vec::new(),
                            stroke_width: None,
                            line_visible: true,
                            smooth: scatter_smooth,
                            marker_symbol: None,
                            marker_size: None,
                            subtotals: Vec::new(),
                            show_values: false,
                            show_category_name: false,
                            show_percent: false,
                            number_format: None,
                            data_label_position: None,
                            data_label_text_color: None,
                            data_label_font_size: None,
                            data_label_font_bold: None,
                            data_label_rotation_degrees: None,
                            hidden_labels: Vec::new(),
                            data_labels: Vec::new(),
                            data_label_border: None,
                            linear_trendline: false,
                            show_trendline_equation: false,
                            show_trendline_r_squared: false,
                            trendline_label_offset: None,
                            trendline_label_font_size: None,
                            x_error_bars: None,
                            y_error_bars: None,
                        });
                    }
                    "marker"
                        if series_depth.is_some_and(|start| depth == start + 1)
                            && marker_depth.is_none() =>
                    {
                        marker_depth = (!empty).then_some(depth);
                    }
                    "marker"
                        if kind == Some(ChartKind::Line)
                            && chart_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        line_has_markers = string_attribute(&attributes, "val", part)?
                            .map_or(true, |value| {
                                value == "1" || value.eq_ignore_ascii_case("true")
                            });
                    }
                    "smooth" if series.as_ref().is_some_and(|s| s.kind == ChartKind::Scatter) => {
                        series.as_mut().unwrap().smooth = string_attribute(&attributes, "val", part)?
                            .is_none_or(|v| v == "1" || v.eq_ignore_ascii_case("true"));
                    }
                    "symbol" if marker_depth.is_some() => {
                        series_marker_is_explicit = true;
                        if let Some(series) = series.as_mut() {
                            series.marker_symbol = string_attribute(&attributes, "val", part)?
                                .filter(|symbol| symbol != "none");
                        }
                    }
                    "size" if marker_depth.is_some() => {
                        if let Some(series) = series.as_mut() {
                            series.marker_size = numeric_attribute(&attributes, "val", part)?
                                .map(|size| size as f32 * 96.0 / 72.0);
                        }
                    }
                    "trendline" if series.is_some() => {
                        trendline_depth = (!empty).then_some(depth);
                    }
                    "trendlineLbl" if trendline_depth.is_some() => {
                        trendline_label_depth = (!empty).then_some(depth);
                        trendline_label_offset = [None; 2];
                    }
                    "layout" if trendline_label_depth.is_some_and(|start| depth == start + 1) => {
                        trendline_label_layout_depth = (!empty).then_some(depth);
                    }
                    "manualLayout"
                        if trendline_label_layout_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        manual_trendline_label_layout_depth = (!empty).then_some(depth);
                    }
                    "x" | "y" if manual_trendline_label_layout_depth.is_some() => {
                        let slot = if local == "x" { 0 } else { 1 };
                        trendline_label_offset[slot] = float_attribute(&attributes, "val", part)?;
                    }
                    "txPr" if trendline_label_depth.is_some_and(|start| depth == start + 1) => {
                        trendline_label_text_properties_depth = (!empty).then_some(depth);
                    }
                    "defRPr" | "rPr" if trendline_label_text_properties_depth.is_some() => {
                        if let Some(series) = series.as_mut() {
                            series.trendline_label_font_size =
                                numeric_attribute(&attributes, "sz", part)?
                                    .map(|size| size as f32 / 100.0 * 96.0 / 72.0);
                        }
                    }
                    "trendlineType" if trendline_depth.is_some() => {
                        if let Some(series) = series.as_mut() {
                            series.linear_trendline = string_attribute(&attributes, "val", part)?
                                .is_some_and(|value| value == "linear");
                        }
                    }
                    "dispEq" | "dispRSqr" if trendline_depth.is_some() => {
                        let shown =
                            string_attribute(&attributes, "val", part)?.is_some_and(|value| {
                                value == "1" || value.eq_ignore_ascii_case("true")
                            });
                        if let Some(series) = series.as_mut() {
                            if local == "dispEq" {
                                series.show_trendline_equation = shown;
                            } else {
                                series.show_trendline_r_squared = shown;
                            }
                        }
                    }
                    "errBars" if series.is_some() => {
                        error_bars_depth = (!empty).then_some(depth);
                        error_bars_are_vertical = true;
                        error_bars = ChartErrorBars::default();
                    }
                    "errDir" if error_bars_depth.is_some() => {
                        error_bars_are_vertical = string_attribute(&attributes, "val", part)?
                            .is_none_or(|value| value != "x");
                    }
                    "errValType" if error_bars_depth.is_some() => {
                        error_bars.standard_error = string_attribute(&attributes, "val", part)?
                            .as_deref()
                            == Some("stdErr");
                    }
                    "errBarType" if error_bars_depth.is_some() => {
                        let value = string_attribute(&attributes, "val", part)?;
                        error_bars.show_plus = value.as_deref() != Some("minus");
                        error_bars.show_minus = value.as_deref() != Some("plus");
                    }
                    "noEndCap" if error_bars_depth.is_some() => {
                        error_bars.end_caps = !string_attribute(&attributes, "val", part)?
                            .is_none_or(|value| value == "1" || value.eq_ignore_ascii_case("true"));
                    }
                    "plus" if error_bars_depth.is_some() => {
                        value_target = (!empty).then_some((depth, ChartValueTarget::ErrorPlus));
                    }
                    "minus" if error_bars_depth.is_some() => {
                        value_target = (!empty).then_some((depth, ChartValueTarget::ErrorMinus));
                    }
                    "ln" if error_bars_depth.is_some_and(|start| depth == start + 2) => {
                        error_bar_line_depth = (!empty).then_some(depth);
                        error_bars.stroke_width = numeric_attribute(&attributes, "w", part)?
                            .map_or(0.75, |width| width as f32 / EMU_PER_CSS_PIXEL);
                    }
                    "noFill" if error_bar_line_depth.is_some() => {
                        error_bars.stroke_width = 0.0;
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if error_bar_line_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        error_bar_stroke_capture = ChartFillCapture::new(local, depth);
                    }
                    "dPt" if series.is_some() && data_point_depth.is_none() => {
                        data_point_depth = (!empty).then_some(depth);
                        data_point_index = None;
                        data_point_color = None;
                        data_point_fill = None;
                        data_point_explosion = None;
                        data_point_border_color = None;
                        data_point_border_width = None;
                    }
                    "idx" if data_point_depth.is_some() => {
                        data_point_index = numeric_attribute(&attributes, "val", part)?
                            .and_then(|value| usize::try_from(value).ok());
                    }
                    "explosion" if series.is_some() => {
                        let value = numeric_attribute(&attributes, "val", part)?.unwrap_or(0);
                        if value > 400 {
                            return Err(format_error(
                                part,
                                "chart data-point explosion exceeds 400 percent",
                            ));
                        }
                        if data_point_depth.is_some() {
                            data_point_explosion = Some(value as f32 / 100.0);
                        } else {
                            series_explosion = Some(value as f32 / 100.0);
                        }
                    }
                    "dLbls" if series.is_some() => {
                        data_labels_depth = (!empty).then_some(depth);
                        series_label_border_width = 1.0;
                    }
                    "dLbls" if kind.is_some() && completed.len() > chart_series_start => {
                        data_labels_depth = (!empty).then_some(depth);
                        group_data_labels_start = Some(chart_series_start);
                        series = completed.last().cloned().map(|mut defaults: ChartSeries| {
                            defaults.show_values = false;
                            defaults.show_category_name = false;
                            defaults.show_percent = false;
                            defaults.number_format = None;
                            defaults.data_label_position = None;
                            defaults.data_label_text_color = None;
                            defaults.data_label_font_size = None;
                            defaults.data_label_font_bold = None;
                            defaults.data_label_rotation_degrees = None;
                            defaults.hidden_labels.clear();
                            defaults.data_labels.clear();
                            defaults.data_label_border = None;
                            defaults
                        });
                    }
                    "dLbl" if data_labels_depth.is_some() => {
                        data_label_depth = (!empty).then_some(depth);
                        data_label_index = None;
                        data_label = (!empty).then(ChartDataLabelCapture::default);
                    }
                    "txPr" if data_labels_depth.is_some() && data_label_depth.is_none() => {
                        data_labels_text_properties_depth = (!empty).then_some(depth);
                    }
                    "manualLayout" if data_label_depth.is_some() => {
                        data_label_manual_layout_depth = (!empty).then_some(depth);
                    }
                    "xMode" | "yMode" if data_label_manual_layout_depth.is_some() => {
                        let slot = if local == "xMode" { 0 } else { 1 };
                        if let Some(label) = data_label.as_mut() {
                            label.manual_offset_edge[slot] =
                                string_attribute(&attributes, "val", part)?.as_deref()
                                    == Some("edge");
                        }
                    }
                    "x" | "y" if data_label_manual_layout_depth.is_some() => {
                        let slot = if local == "x" { 0 } else { 1 };
                        if let Some(label) = data_label.as_mut() {
                            label.manual_offset[slot] = float_attribute(&attributes, "val", part)?;
                        }
                    }
                    "w" | "h" if data_label_manual_layout_depth.is_some() => {
                        let slot = if local == "w" { 0 } else { 1 };
                        if let Some(label) = data_label.as_mut() {
                            label.manual_size[slot] = float_attribute(&attributes, "val", part)?;
                        }
                    }
                    "idx" if data_label_depth.is_some() => {
                        data_label_index = numeric_attribute(&attributes, "val", part)?
                            .and_then(|value| usize::try_from(value).ok());
                    }
                    "delete" if data_label_depth.is_some() => {
                        let deleted =
                            string_attribute(&attributes, "val", part)?.is_some_and(|value| {
                                value == "1" || value.eq_ignore_ascii_case("true")
                            });
                        if deleted
                            && let (Some(series), Some(index)) = (series.as_mut(), data_label_index)
                        {
                            series.hidden_labels.push(index);
                        }
                    }
                    "delete" if data_labels_depth.is_some() => {
                        if group_data_labels_start.is_none()
                            && string_attribute(&attributes, "val", part)?.is_some_and(|value| {
                                value == "1" || value.eq_ignore_ascii_case("true")
                            })
                        {
                            deleted_series_data_labels.insert(completed.len());
                        }
                    }
                    "tx" if data_label_depth
                        .is_some_and(|start| depth == start.saturating_add(1)) =>
                    {
                        data_label_text_depth = (!empty).then_some(depth);
                    }
                    "fld" if data_label_text_depth.is_some() => {
                        data_label_field = string_attribute(&attributes, "type", part)?;
                    }
                    "p" | "br" if data_label_text_depth.is_some() => {
                        if let Some(label) = data_label.as_mut() {
                            if !label.text.is_empty() && !label.text.ends_with('\n') {
                                label.text.push('\n');
                            }
                        }
                    }
                    "t" if data_label_text_depth.is_some() => {
                        collecting_data_label_text = (!empty).then_some((depth, String::new()));
                    }
                    "spPr"
                        if data_label_depth
                            .or(data_labels_depth)
                            .is_some_and(|start| depth == start.saturating_add(1)) =>
                    {
                        data_label_shape_properties_depth = (!empty).then_some(depth);
                    }
                    "ln" if data_label_shape_properties_depth
                        .is_some_and(|start| depth == start.saturating_add(1)) =>
                    {
                        data_label_line_depth = (!empty).then_some(depth);
                        let width = numeric_attribute(&attributes, "w", part)?.unwrap_or(9525)
                            as f32
                            / EMU_PER_CSS_PIXEL;
                        if let Some(label) = data_label.as_mut() {
                            label.border_width = width;
                        } else {
                            series_label_border_width = width;
                        }
                    }
                    "solidFill"
                        if data_label_line_depth
                            .is_some_and(|start| depth == start.saturating_add(1)) =>
                    {
                        data_label_line_fill_depth = (!empty).then_some(depth);
                    }
                    "srgbClr"
                        if data_label_depth.is_some()
                            || data_label_line_fill_depth.is_some()
                            || data_labels_text_properties_depth.is_some() =>
                    {
                        if let Some(color) = string_attribute(&attributes, "val", part)?
                            .as_deref()
                            .and_then(parse_rgb_color)
                        {
                            if let Some(label) = data_label.as_mut() {
                                if data_label_line_fill_depth.is_some() {
                                    label.border_color = Some(color);
                                } else if data_label_shape_properties_depth.is_none() {
                                    label.text_color = Some(color);
                                }
                            } else if data_label_line_fill_depth.is_some()
                                && let Some(series) = series.as_mut()
                            {
                                series.data_label_border = Some((color, series_label_border_width));
                            } else if data_labels_text_properties_depth.is_some()
                                && let Some(series) = series.as_mut()
                            {
                                series.data_label_text_color = Some(color);
                            }
                        }
                    }
                    "schemeClr"
                        if data_label_depth.is_some()
                            || data_label_line_fill_depth.is_some()
                            || data_labels_text_properties_depth.is_some() =>
                    {
                        if let Some(color) = string_attribute(&attributes, "val", part)?
                            .as_deref()
                            .and_then(&scheme_color)
                        {
                            if let Some(label) = data_label.as_mut() {
                                if data_label_line_fill_depth.is_some() {
                                    label.border_color = Some(color);
                                } else if data_label_shape_properties_depth.is_none() {
                                    label.text_color = Some(color);
                                }
                            } else if data_label_line_fill_depth.is_some()
                                && let Some(series) = series.as_mut()
                            {
                                series.data_label_border = Some((color, series_label_border_width));
                            } else if data_labels_text_properties_depth.is_some()
                                && let Some(series) = series.as_mut()
                            {
                                series.data_label_text_color = Some(color);
                            }
                        }
                    }
                    "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff"
                        if (data_label_depth.is_some()
                            && data_label_shape_properties_depth.is_none())
                            || data_labels_text_properties_depth.is_some() =>
                    {
                        let color = data_label
                            .as_mut()
                            .and_then(|label| label.text_color.as_mut())
                            .or_else(|| {
                                series
                                    .as_mut()
                                    .and_then(|series| series.data_label_text_color.as_mut())
                            });
                        if let Some(color) = color {
                            apply_color_transform(
                                color,
                                local,
                                percentage_attribute(&attributes, "val", part)?.unwrap_or(1.0),
                            );
                        }
                    }
                    "defRPr" | "rPr"
                        if (data_label_depth.is_some()
                            && data_label_shape_properties_depth.is_none())
                            || data_labels_text_properties_depth.is_some() =>
                    {
                        if let Some(size) = numeric_attribute(&attributes, "sz", part)? {
                            let size = Some(size as f32 / 100.0 * 96.0 / 72.0);
                            if let Some(label) = data_label.as_mut() {
                                label.font_size = size;
                            } else if let Some(series) = series.as_mut() {
                                series.data_label_font_size = size;
                            }
                        }
                        if let Some(bold) = string_attribute(&attributes, "b", part)? {
                            let bold = Some(bold == "1" || bold.eq_ignore_ascii_case("true"));
                            if let Some(label) = data_label.as_mut() {
                                label.font_bold = bold;
                            } else if let Some(series) = series.as_mut() {
                                series.data_label_font_bold = bold;
                            }
                        }
                    }
                    "showVal" if series.is_some() => {
                        let shown =
                            string_attribute(&attributes, "val", part)?.is_some_and(|value| {
                                value == "1" || value.eq_ignore_ascii_case("true")
                            });
                        if let Some(label) = data_label.as_mut() {
                            label.show_values = Some(shown);
                        } else if let Some(series) = series.as_mut() {
                            series.show_values = shown;
                            if group_data_labels_start.is_none() {
                                explicit_label_flags
                                    .insert((completed.len(), "showVal".to_owned()));
                            }
                        }
                    }
                    "showCatName" if series.is_some() => {
                        let shown =
                            string_attribute(&attributes, "val", part)?.is_some_and(|value| {
                                value == "1" || value.eq_ignore_ascii_case("true")
                            });
                        if let Some(label) = data_label.as_mut() {
                            label.show_category_name = Some(shown);
                        } else if let Some(series) = series.as_mut() {
                            series.show_category_name = shown;
                            if group_data_labels_start.is_none() {
                                explicit_label_flags
                                    .insert((completed.len(), "showCatName".to_owned()));
                            }
                        }
                    }
                    "showPercent" if series.is_some() => {
                        let shown =
                            string_attribute(&attributes, "val", part)?.is_some_and(|value| {
                                value == "1" || value.eq_ignore_ascii_case("true")
                            });
                        if let Some(label) = data_label.as_mut() {
                            label.show_percent = Some(shown);
                        } else if let Some(series) = series.as_mut() {
                            series.show_percent = shown;
                            if group_data_labels_start.is_none() {
                                explicit_label_flags
                                    .insert((completed.len(), "showPercent".to_owned()));
                            }
                        }
                    }
                    "numFmt" if series.is_some() => {
                        let format = string_attribute(&attributes, "formatCode", part)?;
                        if let Some(label) = data_label.as_mut() {
                            label.number_format = format;
                        } else if let Some(series) = series.as_mut() {
                            series.number_format = format;
                        }
                    }
                    "dLblPos" if series.is_some() => {
                        let position = string_attribute(&attributes, "val", part)?;
                        if let Some(label) = data_label.as_mut() {
                            label.position = position;
                        } else if let Some(series) = series.as_mut() {
                            series.data_label_position = position;
                        }
                    }
                    "tx" if series.is_some() => {
                        value_target = (!empty).then_some((depth, ChartValueTarget::Name));
                    }
                    "cat" if series.is_some() => {
                        category_format_code = None;
                        value_target = (!empty).then_some((depth, ChartValueTarget::Categories));
                    }
                    "lvl" if value_target.is_some_and(|(_, target)| matches!(target, ChartValueTarget::Categories)) => {
                        if let Some(series) = series.as_mut() {
                            series.category_levels.push(Vec::new());
                            category_level_depth = (!empty).then_some(depth);
                        }
                    }
                    "xVal" if series.is_some() => {
                        value_target = (!empty).then_some((depth, ChartValueTarget::XValues));
                    }
                    "val" | "yVal" if series.is_some() => {
                        value_target = (!empty).then_some((depth, ChartValueTarget::Values));
                    }
                    "bubbleSize" if series.is_some() => {
                        value_target = (!empty).then_some((depth, ChartValueTarget::BubbleSizes));
                    }
                    "bubble3D" if series.as_ref().is_some_and(|series| series.kind == ChartKind::Bubble) && data_point_depth.is_none() => {
                        if let Some(series) = series.as_mut() {
                            series.three_d = string_attribute(&attributes, "val", part)?
                                .is_none_or(|value| {
                                    value == "1" || value.eq_ignore_ascii_case("true")
                                });
                        }
                    }
                    "v" if series.is_some() && value_target.is_some() => {
                        collecting_value = (!empty).then_some((depth, String::new()));
                    }
                    "pt" if series.is_some() && value_target.is_some() => {
                        cache_point_index = numeric_attribute(&attributes, "idx", part)?
                            .and_then(|value| usize::try_from(value).ok());
                    }
                    "formatCode"
                        if value_target.is_some_and(|(_, target)| {
                            matches!(
                                target,
                                ChartValueTarget::Categories | ChartValueTarget::Values
                            )
                        }) =>
                    {
                        collecting_format_code = (!empty).then_some((depth, String::new()));
                    }
                    "spPr"
                        if data_point_depth.is_some_and(|start| depth == start + 1)
                            || series_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        shape_properties_depth = (!empty).then_some(depth);
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if shape_properties_depth.is_some_and(|start| depth == start + 1) =>
                    {
                        shape_fill_depth = (!empty).then_some(depth);
                        shape_fill_capture = ChartFillCapture::new(local, depth);
                    }
                    "ln" if shape_properties_depth.is_some_and(|start| depth == start + 1) => {
                        shape_line_depth = (!empty).then_some(depth);
                        if let Some(width) = numeric_attribute(&attributes, "w", part)? {
                            let width = width as f32 / EMU_PER_CSS_PIXEL;
                            if data_point_depth.is_some() {
                                data_point_border_width = Some(width);
                            } else if let Some(series) = series
                                .as_mut()
                                .filter(|series| series.stroke_width.is_none())
                            {
                                series.stroke_width = Some(width);
                            }
                        }
                    }
                    "solidFill" if shape_line_depth.is_some_and(|start| depth == start + 1) => {
                        line_fill_depth = (!empty).then_some(depth);
                    }
                    "noFill" if shape_line_depth.is_some_and(|start| depth == start + 1) => {
                        if data_point_depth.is_none()
                            && let Some(series) = series.as_mut()
                        {
                            series.line_visible = false;
                            series.stroke_width = Some(0.0);
                            series_border_color = None;
                        } else if data_point_depth.is_some() {
                            data_point_border_width = Some(0.0);
                            data_point_border_color = None;
                        }
                    }
                    "axId" if chart_depth.is_some() && series.is_none() => {
                        if let Some(axis_id) = string_attribute(&attributes, "val", part)? {
                            chart_axis_ids.push(axis_id);
                        }
                    }
                    "srgbClr" if shape_fill_depth.is_some() || line_fill_depth.is_some() => {
                        if let Some(color) = string_attribute(&attributes, "val", part)?
                            .as_deref()
                            .and_then(parse_rgb_color)
                        {
                            if line_fill_depth.is_some() {
                                if data_point_depth.is_some() {
                                    data_point_border_color = Some(color);
                                } else {
                                    series_border_color = Some(color);
                                    if let Some(series) = series.as_mut().filter(|series| series.color.is_none()) { series.color = Some(color); }
                                }
                            } else if data_point_depth.is_some() {
                                data_point_color.get_or_insert(color);
                            } else if let Some(series) =
                                series.as_mut().filter(|series| series.color.is_none())
                            {
                                series.color = Some(color);
                            }
                        }
                    }
                    "schemeClr" if shape_fill_depth.is_some() || line_fill_depth.is_some() => {
                        if let Some(color) = string_attribute(&attributes, "val", part)?
                            .as_deref()
                            .and_then(&scheme_color)
                        {
                            if line_fill_depth.is_some() {
                                if data_point_depth.is_some() {
                                    data_point_border_color = Some(color);
                                } else {
                                    series_border_color = Some(color);
                                    if let Some(series) = series.as_mut().filter(|series| series.color.is_none()) { series.color = Some(color); }
                                }
                            } else if data_point_depth.is_some() {
                                data_point_color.get_or_insert(color);
                            } else if let Some(series) =
                                series.as_mut().filter(|series| series.color.is_none())
                            {
                                series.color = Some(color);
                            }
                        }
                    }
                    "alpha" | "tint" | "shade" | "lumMod" | "lumOff" | "satMod" | "satOff"
                        if shape_fill_depth.is_some() || line_fill_depth.is_some() =>
                    {
                        let ratio = percentage_attribute(&attributes, "val", part)?.unwrap_or(1.0);
                        if line_fill_depth.is_some() {
                            if let Some(color) = data_point_border_color.as_mut() {
                                apply_color_transform(color, local, ratio);
                            } else if let Some(color) =
                                series.as_mut().and_then(|series| series.color.as_mut())
                            {
                                apply_color_transform(color, local, ratio);
                            }
                        } else if let Some(color) = data_point_color.as_mut() {
                            apply_color_transform(color, local, ratio);
                        } else if let Some(color) =
                            series.as_mut().and_then(|series| series.color.as_mut())
                        {
                            apply_color_transform(color, local, ratio);
                        }
                    }
                    _ => {}
                }
                if let Some(capture) = error_bar_stroke_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = surface_band_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = title_text_fill_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = title_fill_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = chart_area_fill_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = plot_area_fill_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = legend_fill_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = legend_stroke_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = axis_title_fill_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = axis_title_stroke_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if shape_properties_depth.is_some() && data_point_depth.is_none() {
                    if matches!(local, "effectLst" | "effectDag" | "scene3d" | "sp3d") { series_effects_explicit = true; }
                    series_effects.start(local, &attributes, empty, depth, part, |kind, value| Ok(if kind == "schemeClr" { scheme_color(value).unwrap_or(0x000000ff) } else { parse_rgb_color(value).unwrap_or(0x000000ff) }))?;
                }
                if axis_title_properties_depth.is_some() {
                    axis_title_effects_capture.start(
                        local,
                        &attributes,
                        empty,
                        depth,
                        part,
                        |kind, value| {
                            Ok(match kind {
                                "srgbClr" | "sysClr" => {
                                    parse_rgb_color(value).unwrap_or(0x0000_00ff)
                                }
                                "schemeClr" => scheme_color(value).unwrap_or(0x0000_00ff),
                                "prstClr" if value == "white" => 0xffff_ffff,
                                _ => 0x0000_00ff,
                            })
                        },
                    )?;
                }
                if let Some(capture) = up_down_bar_fill_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = up_down_bar_stroke_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if let Some(capture) = shape_fill_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if let Some(capture) = error_bar_stroke_capture.as_mut() {
                    capture.end(local);
                }
                if error_bar_stroke_capture
                    .as_ref()
                    .is_some_and(|capture| capture.closes_at(depth))
                {
                    error_bars.stroke = error_bar_stroke_capture
                        .take()
                        .unwrap()
                        .finish(package, part)?;
                }
                if local == "ln" && error_bar_line_depth == Some(depth) {
                    error_bar_line_depth = None;
                }
                if let Some(capture) = surface_band_capture.as_mut() {
                    capture.end(local);
                }
                if surface_band_capture
                    .as_ref()
                    .is_some_and(|capture| capture.closes_at(depth))
                {
                    let fill = surface_band_capture.take().unwrap().finish(package, part)?;
                    surface_band_fills
                        .resize_with(surface_band_fills.len().max(surface_band_index + 1), || {
                            None
                        });
                    surface_band_fills[surface_band_index] = fill;
                }
                if local == "spPr" && surface_band_properties == Some(depth) {
                    surface_band_properties = None;
                }
                if local == "bandFmt" && surface_band_depth == Some(depth) {
                    surface_band_depth = None;
                }

                if let Some(capture) = title_text_fill_capture.as_mut() {
                    capture.end(local);
                }
                if let Some(capture) = title_fill_capture.as_mut() {
                    capture.end(local);
                }
                if let Some(capture) = chart_area_fill_capture.as_mut() {
                    capture.end(local);
                }
                if let Some(capture) = plot_area_fill_capture.as_mut() {
                    capture.end(local);
                }
                if let Some(capture) = legend_fill_capture.as_mut() {
                    capture.end(local);
                }
                if let Some(capture) = legend_stroke_capture.as_mut() {
                    capture.end(local);
                }
                if let Some(capture) = axis_title_fill_capture.as_mut() {
                    capture.end(local);
                }
                if let Some(capture) = axis_title_stroke_capture.as_mut() {
                    capture.end(local);
                }
                axis_title_effects_capture.end(depth);
                series_effects.end(depth);
                if let Some(capture) = up_down_bar_fill_capture.as_mut() {
                    capture.end(local);
                }
                if let Some(capture) = up_down_bar_stroke_capture.as_mut() {
                    capture.end(local);
                }
                if let Some(capture) = shape_fill_capture.as_mut() {
                    capture.end(local);
                }
                if shape_fill_capture
                    .as_ref()
                    .is_some_and(|capture| capture.depth == depth)
                {
                    let fill = shape_fill_capture
                        .take()
                        .ok_or_else(|| format_error(part, "chart series fill state is missing"))?
                        .finish(package, part)?;
                    if data_point_depth.is_some() {
                        data_point_fill = fill;
                    } else if let Some(series) = series.as_mut() {
                        series.fill = fill;
                    }
                }
                if title_text_fill_capture.as_ref().is_some_and(|capture| capture.depth == depth) {
                    if let Some(ChartFill::Solid(color)) = title_text_fill_capture.take().unwrap().finish(package, part)? {
                        title_text_color = Some(color);
                    }
                }
                if title_fill_capture
                    .as_ref()
                    .is_some_and(|capture| capture.depth == depth)
                {
                    title_fill = title_fill_capture
                        .take()
                        .ok_or_else(|| format_error(part, "chart title fill state is missing"))?
                        .finish(package, part)?;
                }
                if chart_area_fill_capture
                    .as_ref()
                    .is_some_and(|capture| capture.depth == depth)
                {
                    let fill = chart_area_fill_capture
                        .take()
                        .ok_or_else(|| format_error(part, "chart-area fill state is missing"))?
                        .finish(package, part)?;
                    if chart_area_line_depth.is_some() {
                        chart_area_border = fill.map(|fill| (fill, chart_area_line_width));
                    } else {
                        chart_area_fill = fill;
                    }
                }
                if plot_area_fill_capture
                    .as_ref()
                    .is_some_and(|capture| capture.depth == depth)
                {
                    plot_area_fill = plot_area_fill_capture
                        .take()
                        .ok_or_else(|| format_error(part, "plot-area fill state is missing"))?
                        .finish(package, part)?;
                }
                if legend_fill_capture
                    .as_ref()
                    .is_some_and(|capture| capture.depth == depth)
                {
                    legend_fill = legend_fill_capture
                        .take()
                        .ok_or_else(|| format_error(part, "chart legend fill state is missing"))?
                        .finish(package, part)?;
                }
                if legend_stroke_capture
                    .as_ref()
                    .is_some_and(|capture| capture.depth == depth)
                {
                    legend_stroke = legend_stroke_capture
                        .take()
                        .ok_or_else(|| format_error(part, "chart legend stroke state is missing"))?
                        .finish(package, part)?;
                }
                if axis_title_fill_capture
                    .as_ref()
                    .is_some_and(|capture| capture.depth == depth)
                {
                    let fill = axis_title_fill_capture
                        .take()
                        .ok_or_else(|| format_error(part, "axis title fill state is missing"))?
                        .finish(package, part)?;
                    if let Some(options) = current_value_axis_options.as_mut() {
                        options.title_fill = fill;
                    }
                }
                if axis_title_stroke_capture
                    .as_ref()
                    .is_some_and(|capture| capture.depth == depth)
                {
                    let stroke = axis_title_stroke_capture
                        .take()
                        .ok_or_else(|| format_error(part, "axis title stroke state is missing"))?
                        .finish(package, part)?;
                    if let Some(options) = current_value_axis_options.as_mut() {
                        options.title_stroke = stroke;
                    }
                }
                if up_down_bar_fill_capture
                    .as_ref()
                    .is_some_and(|capture| capture.depth == depth)
                {
                    let fill = up_down_bar_fill_capture
                        .take()
                        .ok_or_else(|| {
                            format_error(part, "chart up/down bar fill state is missing")
                        })?
                        .finish(package, part)?;
                    if let (Some(options), Some(is_up)) = (up_down_bars.as_mut(), up_down_bar_is_up)
                    {
                        if is_up {
                            options.up_fill = fill;
                        } else {
                            options.down_fill = fill;
                        }
                    }
                }
                if up_down_bar_stroke_capture
                    .as_ref()
                    .is_some_and(|capture| capture.depth == depth)
                {
                    let stroke = up_down_bar_stroke_capture
                        .take()
                        .ok_or_else(|| {
                            format_error(part, "chart up/down bar stroke state is missing")
                        })?
                        .finish(package, part)?;
                    if let (Some(options), Some(is_up)) = (up_down_bars.as_mut(), up_down_bar_is_up)
                    {
                        if is_up {
                            options.up_stroke = stroke;
                        } else {
                            options.down_stroke = stroke;
                        }
                    }
                }
                if local == "manualLayout" && title_layout_depth == Some(depth) {
                    if let [Some(x), Some(y)] = title_coordinates { title_position = Some([(x, title_edge[0]), (y, title_edge[1])]); }
                    title_layout_depth = None;
                } else if local == "manualLayout" && manual_plot_layout_depth == Some(depth) {
                    if let [Some(x), Some(y), Some(width), Some(height)] = manual_plot_layout {
                        let bounds = Rect {
                            x,
                            y,
                            width,
                            height,
                        };
                        if bounds.is_valid() {
                            plot_bounds = Some(bounds);
                        }
                    }
                    manual_plot_layout_depth = None;
                } else if local == "layout" && plot_layout_depth == Some(depth) {
                    plot_layout_depth = None;
                } else if local == "manualLayout" && manual_legend_layout_depth == Some(depth) {
                    if let [Some(x), Some(y), Some(mut width), Some(mut height)] =
                        manual_legend_layout
                    {
                        if manual_legend_size_edge[0] {
                            width -= x;
                        }
                        if manual_legend_size_edge[1] {
                            height -= y;
                        }
                        let bounds = Rect {
                            x,
                            y,
                            width,
                            height,
                        };
                        if bounds.is_valid() {
                            legend_bounds = Some(bounds);
                        }
                    }
                    manual_legend_layout_depth = None;
                } else if local == "layout" && legend_layout_depth == Some(depth) {
                    legend_layout_depth = None;
                } else if local == "spPr" && plot_area_properties_depth == Some(depth) {
                    plot_area_properties_depth = None;
                } else if local == "plotArea" && plot_area_depth == Some(depth) {
                    plot_area_depth = None;
                }
                if local == "txPr" && chart_text_properties_depth == Some(depth) {
                    chart_text_properties_depth = None;
                }
                if local == "txPr" && data_labels_text_properties_depth == Some(depth) {
                    data_labels_text_properties_depth = None;
                }
                if local == "txPr" && axis_text_properties_depth == Some(depth) {
                    axis_text_properties_depth = None;
                }
                if local == "ln" && chart_area_line_depth == Some(depth) {
                    chart_area_line_depth = None;
                }
                if matches!(local, "majorGridlines" | "minorGridlines") && gridlines_depth.is_some_and(|(start, _)| start == depth) {
                    gridlines_depth = None;
                }
                if local == "ln" && axis_line_depth == Some(depth) {
                    axis_line_depth = None;
                }
                if local == "spPr" && axis_properties_depth == Some(depth) {
                    axis_properties_depth = None;
                }
                if local == "ln" && axis_title_line_depth == Some(depth) {
                    axis_title_line_depth = None;
                } else if local == "spPr" && axis_title_properties_depth == Some(depth) {
                    if let Some(options) = current_value_axis_options.as_mut() {
                        options.title_effects =
                            std::mem::take(&mut axis_title_effects_capture).finish();
                    }
                    axis_title_properties_depth = None;
                } else if local == "rich" && axis_title_body_depth == Some(depth) {
                    axis_title_body_depth = None;
                }
                if local == "scaling" && value_axis_scaling_depth == Some(depth) {
                    value_axis_scaling_depth = None;
                } else if matches!(local, "valAx" | "catAx" | "dateAx" | "serAx")
                    && value_axis_depth == Some(depth)
                {
                    let mut options = current_value_axis_options.take().unwrap_or_default();
                    options.title = std::mem::take(&mut axis_title);
                    if local == "serAx" {
                        series_axis_options = Some(options);
                    } else {
                        // Horizontal bars rotate category/value roles even when an
                        // exporter leaves column-chart axPos values in the source.
                        let horizontal = if completed.iter().any(|series: &ChartSeries| {
                            series.kind == ChartKind::Bar && series.bar_horizontal
                        }) {
                            local == "valAx"
                        } else {
                            matches!(options.position.as_deref(), Some("b" | "t"))
                        };
                        let retained = if horizontal {
                            &mut horizontal_axis_options
                        } else {
                            &mut value_axis_options
                        };
                        if options.title.trim().is_empty() && !retained.title.trim().is_empty() {
                            options.title = std::mem::take(&mut retained.title);
                            options.title_font_size = retained.title_font_size;
                            options.title_bold = retained.title_bold;
                        }
                        if retained.id.is_some() && retained.id != options.id {
                            // A second category axis must not overwrite the price/value
                            // axis of a combination chart.
                            if local == "valAx" {
                                secondary_value_axis_options = Some(options);
                            }
                        } else {
                            *retained = options;
                        }
                    }
                    value_axis_depth = None;
                } else if local == "dTable" && data_table_depth == Some(depth) {
                    data_table_depth = None;
                } else if local == "ln" && up_down_bar_line_depth == Some(depth) {
                    up_down_bar_line_depth = None;
                } else if local == "spPr" && up_down_bar_properties_depth == Some(depth) {
                    up_down_bar_properties_depth = None;
                } else if matches!(local, "upBars" | "downBars") && up_down_bar_depth == Some(depth)
                {
                    up_down_bar_depth = None;
                    up_down_bar_is_up = None;
                } else if local == "upDownBars" && up_down_bars_depth == Some(depth) {
                    up_down_bars_depth = None;
                }
                if local == "dispUnitsLbl" && display_unit_label_depth == Some(depth) {
                    display_unit_label_depth = None;
                }
                if local == "t" && display_unit_text_depth == Some(depth) {
                    display_unit_text_depth = None;
                }
                if local == "t" && axis_title_text_depth == Some(depth) {
                    axis_title_text_depth = None;
                } else if local == "t" && title_text_depth == Some(depth) {
                    title_text_depth = None;
                } else if local == "t"
                    && collecting_data_label_text
                        .as_ref()
                        .is_some_and(|(start, _)| *start == depth)
                {
                    let (_, text) = collecting_data_label_text.take().ok_or_else(|| {
                        format_error(
                            part,
                            "chart data-label text parser state ended unexpectedly",
                        )
                    })?;
                    if let Some(label) = data_label.as_mut() {
                        let start = label.text.len();
                        label.text.push_str(&text);
                        if let (Some(index), Some(field)) =
                            (data_label_index, data_label_field.as_ref())
                        {
                            data_label_fields.push((
                                completed.len(),
                                index,
                                start..label.text.len(),
                                field.clone(),
                            ));
                        }
                    }
                } else if local == "fld" && data_label_text_depth.is_some() {
                    data_label_field = None;
                } else if local == "tx" && data_label_text_depth == Some(depth) {
                    data_label_text_depth = None;
                } else if local == "title" && axis_title_depth == Some(depth) {
                    if axis_title.trim().is_empty() { axis_title = "Axis Title".to_owned(); }
                    axis_title_depth = None;
                } else if local == "title" && title_depth == Some(depth) {
                    title_depth = None;
                } else if local == "spPr" && title_properties_depth == Some(depth) {
                    title_properties_depth = None;
                } else if local == "spPr" && chart_area_properties_depth == Some(depth) {
                    chart_area_properties_depth = None;
                } else if local == "ln" && legend_line_depth == Some(depth) {
                    legend_line_depth = None;
                } else if local == "spPr" && legend_properties_depth == Some(depth) {
                    legend_properties_depth = None;
                } else if local == "legend" && legend_depth == Some(depth) {
                    legend_depth = None;
                } else if local == "legendEntry" && legend_entry_depth == Some(depth) {
                    legend_entry_depth = None;
                    legend_entry_index = None;
                }
                if local == "v"
                    && collecting_value
                        .as_ref()
                        .is_some_and(|(start, _)| *start == depth)
                {
                    let (_, value) = collecting_value.take().ok_or_else(|| {
                        format_error(part, "chart value parser state ended unexpectedly")
                    })?;
                    if let (Some((_, target)), Some(series)) = (value_target, series.as_mut()) {
                        match target {
                            ChartValueTarget::Name => {
                                if !series.name.is_empty() && !value.is_empty() {
                                    series.name.push(' ');
                                }
                                series.name.push_str(&value);
                            }
                            ChartValueTarget::Categories => {
                                if category_level_depth.is_some() {
                                    if let Some(level) = series.category_levels.last_mut() {
                                        set_string_at(level, cache_point_index, value);
                                    }
                                } else {
                                    if let Ok(number) = value.parse::<f32>() && number.is_finite() {
                                        set_number_at(&mut series.x_values, cache_point_index, number);
                                    }
                                    set_string_at(&mut series.categories, cache_point_index,
                                        format_chart_category(&value, category_format_code.as_deref(), date_1904).unwrap_or(value));
                                }
                            }
                            ChartValueTarget::XValues => {
                                if let Ok(number) = value.parse::<f32>()
                                    && number.is_finite()
                                {
                                    set_number_at(&mut series.x_values, cache_point_index, number);
                                } else {
                                    set_string_at(&mut series.categories, cache_point_index, value);
                                }
                            }
                            ChartValueTarget::Values => {
                                set_parsed_number_at(&mut series.values, cache_point_index, &value);
                                set_string_at(&mut series.raw_values, cache_point_index, value);
                            }
                            ChartValueTarget::BubbleSizes => set_parsed_number_at(
                                &mut series.bubble_sizes,
                                cache_point_index,
                                &value,
                            ),
                            ChartValueTarget::ErrorPlus => set_parsed_number_at(
                                &mut error_bars.plus,
                                cache_point_index,
                                &value,
                            ),
                            ChartValueTarget::ErrorMinus => set_parsed_number_at(
                                &mut error_bars.minus,
                                cache_point_index,
                                &value,
                            ),
                        }
                    }
                } else if local == "formatCode"
                    && collecting_format_code
                        .as_ref()
                        .is_some_and(|(start, _)| *start == depth)
                {
                    let format_code = collecting_format_code.take().map(|(_, value)| value);
                    if value_target
                        .is_some_and(|(_, target)| matches!(target, ChartValueTarget::Categories))
                    {
                        category_format_code = format_code;
                    } else if let Some(series) = series.as_mut() {
                        series.number_format = format_code;
                    }
                } else if matches!(
                    local,
                    "tx" | "cat" | "xVal" | "val" | "yVal" | "plus" | "minus"
                ) && value_target
                    .as_ref()
                    .is_some_and(|(start, _)| *start == depth)
                {
                    value_target = None;
                } else if local == "lvl" && category_level_depth == Some(depth) {
                    category_level_depth = None;
                    if let Some(series) = series.as_mut() {
                        if series.category_levels.len() == 1 {
                            series.categories.clone_from(&series.category_levels[0]);
                        }
                    }
                } else if local == "pt" && value_target.is_some() {
                    cache_point_index = None;
                } else if matches!(local, "solidFill" | "gradFill" | "pattFill" | "blipFill")
                    && shape_fill_depth == Some(depth)
                {
                    shape_fill_depth = None;
                } else if local == "solidFill" && line_fill_depth == Some(depth) {
                    line_fill_depth = None;
                } else if local == "ln" && shape_line_depth == Some(depth) {
                    shape_line_depth = None;
                } else if local == "solidFill" && data_label_line_fill_depth == Some(depth) {
                    data_label_line_fill_depth = None;
                } else if local == "ln" && data_label_line_depth == Some(depth) {
                    data_label_line_depth = None;
                } else if local == "spPr" && data_label_shape_properties_depth == Some(depth) {
                    data_label_shape_properties_depth = None;
                } else if local == "spPr" && shape_properties_depth == Some(depth) {
                    shape_properties_depth = None;
                } else if local == "dPt" && data_point_depth == Some(depth) {
                    if let (Some(index), Some(color)) = (data_point_index, data_point_color) {
                        data_point_color_overrides.push((index, color));
                    }
                    if let (Some(index), Some(fill)) = (data_point_index, data_point_fill.take()) {
                        data_point_fill_overrides.push((index, fill));
                    }
                    if let (Some(index), Some(explosion)) = (data_point_index, data_point_explosion)
                    {
                        data_point_explosion_overrides.push((index, explosion));
                    }
                    if let (Some(index), Some(width)) = (data_point_index, data_point_border_width)
                    {
                        data_point_border_overrides.push((index, data_point_border_color.or(series_border_color), width));
                    }
                    data_point_depth = None;
                    data_point_index = None;
                    data_point_color = None;
                    data_point_fill = None;
                    data_point_explosion = None;
                    data_point_border_color = None;
                    data_point_border_width = None;
                } else if local == "manualLayout"
                    && manual_trendline_label_layout_depth == Some(depth)
                {
                    if let [Some(x), Some(y)] = trendline_label_offset
                        && let Some(series) = series.as_mut()
                    {
                        series.trendline_label_offset = Some((x, y));
                    }
                    manual_trendline_label_layout_depth = None;
                } else if local == "layout" && trendline_label_layout_depth == Some(depth) {
                    trendline_label_layout_depth = None;
                } else if local == "txPr" && trendline_label_text_properties_depth == Some(depth) {
                    trendline_label_text_properties_depth = None;
                } else if local == "trendlineLbl" && trendline_label_depth == Some(depth) {
                    trendline_label_depth = None;
                } else if local == "errBars" && error_bars_depth == Some(depth) {
                    if let Some(series) = series.as_mut() {
                        let target = if error_bars_are_vertical {
                            &mut series.y_error_bars
                        } else {
                            &mut series.x_error_bars
                        };
                        *target = Some(std::mem::take(&mut error_bars));
                    }
                    error_bars_depth = None;
                } else if local == "ser" && series_depth == Some(depth) {
                    let mut current = series.take().ok_or_else(|| {
                        format_error(part, "chart series parser state ended unexpectedly")
                    })?;
                    if current.kind == ChartKind::Scatter
                        && current.x_values.is_empty()
                        && !current.categories.is_empty()
                    {
                        current.x_values = (1..=current.values.len())
                            .map(|value| value as f32)
                            .collect();
                    }
                    if let Some(errors) = current.x_error_bars.as_mut() {
                        errors.resolve(&current.x_values);
                    }
                    if let Some(errors) = current.y_error_bars.as_mut() {
                        errors.resolve(&current.values);
                    }
                    if current.categories.is_empty()
                        && !matches!(current.kind, ChartKind::Scatter | ChartKind::Bubble)
                    {
                        current.categories = (1..=current.values.len())
                            .map(|index| index.to_string())
                            .collect();
                    }
                    if current.kind == ChartKind::Scatter
                        && scatter_has_markers
                        && current.marker_symbol.is_none()
                        && !series_marker_is_explicit
                    {
                        current.marker_symbol = Some(
                            ["diamond", "triangle", "square", "circle"][completed.len() % 4]
                                .to_owned(),
                        );
                    }
                    if !current.values.is_empty() {
                        if chart_vary_colors {
                            current.point_colors = current.color.map_or_else(
                                || {
                                    (0..current.values.len() + usize::from(matches!(current.kind, ChartKind::BarOfPie(_))))
                                        .map(|index| {
                                            scheme_color(CHART_ACCENTS[index % CHART_ACCENTS.len()])
                                                .unwrap_or_else(|| {
                                                    super::office_chart_palette_color(index)
                                                })
                                        })
                                        .collect()
                                },
                                |color| vec![color; current.values.len()],
                            );
                        } else if matches!(
                            current.kind,
                            ChartKind::Pie | ChartKind::BarOfPie(_) | ChartKind::Doughnut
                        ) && let Some(color) = current.color
                        {
                            current.point_colors = vec![color; current.values.len()];
                        }
                        if !data_point_color_overrides.is_empty() {
                            if current.point_colors.len() != current.values.len() {
                                let series_index = source_series_index;
                                let color = current.color.unwrap_or_else(|| {
                                    scheme_color(CHART_ACCENTS[series_index % CHART_ACCENTS.len()])
                                        .unwrap_or_else(|| {
                                            super::office_chart_palette_color(series_index)
                                        })
                                });
                                current.point_colors = vec![color; current.values.len()];
                            }
                            for (index, color) in data_point_color_overrides.drain(..) {
                                if let Some(point_color) = current.point_colors.get_mut(index) {
                                    *point_color = color;
                                }
                            }
                        }
                        current.point_fills = vec![current.fill.clone(); current.values.len()];
                        for (index, fill) in data_point_fill_overrides.drain(..) {
                            if let Some(point_fill) = current.point_fills.get_mut(index) {
                                *point_fill = Some(fill);
                            }
                        }
                        current.point_explosions =
                            vec![series_explosion.unwrap_or(0.0); current.values.len()];
                        for (index, explosion) in data_point_explosion_overrides.drain(..) {
                            if let Some(point_explosion) = current.point_explosions.get_mut(index) {
                                *point_explosion = explosion;
                            }
                        }
                        current.effects = std::mem::take(&mut series_effects).finish();
                        let beveled_fill = chart_style.is_some_and(|style| (25..=32).contains(&style))
                            && (matches!(current.kind, ChartKind::Bar | ChartKind::Area | ChartKind::Pie | ChartKind::Doughnut | ChartKind::Bubble)
                                || current.kind == ChartKind::Radar && radar_filled);
                        if !series_effects_explicit && !current.three_d && (chart_style == Some(13) || beveled_fill) {
                            if default_series_effects.is_none() {
                                default_series_effects = Some(chart_theme_effects(package, part, theme_part, if beveled_fill { 2 } else { 0 }, &scheme_color)?.unwrap_or_default());
                            }
                            current.effects = default_series_effects.as_ref().unwrap().clone();
                        }
                        current.point_border_colors = vec![series_border_color; current.values.len()];
                        current.point_border_widths = vec![if series_border_color.is_some() { current.stroke_width.unwrap_or(1.0) } else { 0.0 }; current.values.len()];
                        for (index, color, width) in data_point_border_overrides.drain(..) {
                            if let Some(border_color) = current.point_border_colors.get_mut(index) {
                                *border_color = color;
                            }
                            if let Some(border_width) = current.point_border_widths.get_mut(index) {
                                *border_width = width;
                            }
                        }
                        if series_marker_is_explicit {
                            explicit_marker_series.insert(completed.len());
                        }
                        source_series_indices.push(source_series_index);
                        completed.push(current);
                    }
                    if completed.len() > package.limits().max_document_objects {
                        return Err(Diagnostic::fatal(
                            DiagnosticCode::ObjectLimit,
                            Phase::Parse,
                            None,
                            "chart series exceeds the configured object limit",
                        )
                        .in_part(part));
                    }
                    series_depth = None;
                } else if local == "trendline" && trendline_depth == Some(depth) {
                    trendline_depth = None;
                } else if local == "marker" && marker_depth == Some(depth) {
                    marker_depth = None;
                } else if local == "dLbl" && data_label_depth == Some(depth) {
                    if let (Some(index), Some(label)) = (data_label_index, data_label.take())
                        && let Some(series) = series.as_mut()
                    {
                        series.data_labels.push(ChartDataLabel {
                            index,
                            text: label.text,
                            text_color: label.text_color,
                            font_size: label.font_size,
                            font_bold: label.font_bold,
                            manual_offset: match (label.manual_offset, label.manual_offset_edge) {
                                ([Some(x), Some(y)], [false, false]) => Some((x, y)),
                                _ => None,
                            },
                            manual_size: match label.manual_size {
                                [Some(width), Some(height)] => Some((width, height)),
                                _ => None,
                            },
                            border_color: label.border_color,
                            border_width: label.border_width,
                            show_values: label.show_values,
                            show_category_name: label.show_category_name,
                            show_percent: label.show_percent,
                            number_format: label.number_format,
                            position: label.position,
                        });
                    }
                    data_label_depth = None;
                    data_label_index = None;
                    data_label_text_depth = None;
                    collecting_data_label_text = None;
                    data_label_shape_properties_depth = None;
                    data_label_line_depth = None;
                    data_label_line_fill_depth = None;
                    data_label_manual_layout_depth = None;
                } else if local == "dLbls" && data_labels_depth == Some(depth) {
                    if let Some(start) = group_data_labels_start.take()
                        && let Some(defaults) = series.take()
                    {
                        for (index, target) in completed.iter_mut().enumerate().skip(start) {
                            if deleted_series_data_labels.contains(&index) {
                                continue;
                            }
                            if !explicit_label_flags.contains(&(index, "showVal".to_owned())) {
                                target.show_values = defaults.show_values;
                            }
                            if !explicit_label_flags.contains(&(index, "showCatName".to_owned())) {
                                target.show_category_name = defaults.show_category_name;
                            }
                            if !explicit_label_flags.contains(&(index, "showPercent".to_owned())) {
                                target.show_percent = defaults.show_percent;
                            }
                            if target.number_format.is_none() {
                                target.number_format.clone_from(&defaults.number_format);
                            }
                            if target.data_label_position.is_none() {
                                target
                                    .data_label_position
                                    .clone_from(&defaults.data_label_position);
                            }
                            if target.data_label_text_color.is_none() {
                                target
                                    .data_label_text_color
                                    .clone_from(&defaults.data_label_text_color);
                            }
                            target.data_label_font_size = target
                                .data_label_font_size
                                .or(defaults.data_label_font_size);
                            target.data_label_font_bold = target
                                .data_label_font_bold
                                .or(defaults.data_label_font_bold);
                            target.data_label_rotation_degrees = target
                                .data_label_rotation_degrees
                                .or(defaults.data_label_rotation_degrees);
                            for hidden in &defaults.hidden_labels {
                                if !target.hidden_labels.contains(hidden) {
                                    target.hidden_labels.push(*hidden);
                                }
                            }
                            for label in &defaults.data_labels {
                                if !target
                                    .data_labels
                                    .iter()
                                    .any(|own| own.index == label.index)
                                {
                                    target.data_labels.push(label.clone());
                                }
                            }
                            target.data_label_border =
                                target.data_label_border.or(defaults.data_label_border);
                        }
                    }
                    data_labels_depth = None;
                    data_labels_text_properties_depth = None;
                } else if is_chart_element(local) && chart_depth == Some(depth) {
                    if let Some(options) = up_down_bars.as_mut() {
                        options.series_end = completed.len();
                    }
                    let value_axis_id = chart_axis_ids
                        .get(1)
                        .cloned()
                        .or_else(|| Some(format!("chart-group:{chart_series_start}")));
                    for (series_index, series) in
                        completed.iter_mut().enumerate().skip(chart_series_start)
                    {
                        series.axis_id.clone_from(&value_axis_id);
                        series.first_slice_angle = chart_first_slice_angle;
                        if line_has_markers
                            && !explicit_marker_series.contains(&series_index)
                            && series.marker_symbol.is_none()
                        {
                            series.marker_symbol = Some(
                                ["diamond", "square", "triangle", "x", "circle"]
                                    [source_series_indices[series_index] % 5]
                                    .to_owned(),
                            );
                            // Automatic markers need space beyond a thick series stroke;
                            // an explicitly authored marker size still takes precedence.
                            series
                                .marker_size
                                .get_or_insert((series.stroke_width.unwrap_or(2.0) * 2.0).max(8.0));
                        }
                        if matches!(series.kind, ChartKind::BarOfPie(_)) {
                            let split = if chart_bar_of_pie_split == 0 {
                                series.values.len() / 3 + 1
                            } else {
                                series
                                    .values
                                    .len()
                                    .saturating_sub(usize::from(chart_bar_of_pie_split))
                            };
                            series.kind = if chart_pie_of_pie {
                                ChartKind::PieOfPie(u16::try_from(split).unwrap_or(u16::MAX))
                            } else {
                                ChartKind::BarOfPie(u16::try_from(split).unwrap_or(u16::MAX))
                            };
                        }
                        if series.kind == ChartKind::Radar && radar_filled {
                            filled_radars.push(series_index);
                        }
                        series.grouping = chart_grouping;
                        series.bar_horizontal = chart_bar_horizontal;
                        series.bar_depth = chart_bar_depth;
                        series.bar_cone |= chart_bar_cone;
                        series.bar_cylinder = chart_bar_cylinder;
                        series.bar_gap_depth_percent = chart_bar_gap_depth_percent;
                    }
                    chart_depth = None;
                    chart_three_d = false;
                    chart_grouping = ChartGrouping::Standard;
                    chart_bar_horizontal = false;
                    chart_bar_depth = false;
                    chart_bar_cone = false;
                    chart_bar_cylinder = false;
                    chart_bar_gap_depth_percent = 150.0;
                    chart_vary_colors = false;
                    chart_first_slice_angle = 0.0;
                    chart_bar_of_pie_split = 0;
                    kind = None;
                    chart_axis_ids.clear();
                }
            }
            XmlEvent::Text(text) => {
                if display_unit_text_depth.is_some() {
                    if let Some(label) = current_value_axis_options.as_mut().and_then(|axis| axis.display_unit_label.as_mut()) {
                        label.push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                    }
                } else if axis_title_text_depth.is_some() {
                    axis_title
                        .push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                } else if title_text_depth.is_some() {
                    title.push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                } else if let Some((_, value)) = collecting_data_label_text.as_mut() {
                    value.push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                } else if let Some((_, value)) = collecting_format_code.as_mut() {
                    value.push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                } else if let Some((_, value)) = collecting_value.as_mut() {
                    value.push_str(&decode_xml_text(text).map_err(|error| with_part(error, part))?);
                }
            }
            XmlEvent::Cdata(text) => {
                if display_unit_text_depth.is_some() {
                    if let Some(label) = current_value_axis_options.as_mut().and_then(|axis| axis.display_unit_label.as_mut()) {
                        label.push_str(text);
                    }
                } else if axis_title_text_depth.is_some() {
                    axis_title.push_str(text);
                } else if title_text_depth.is_some() {
                    title.push_str(text);
                } else if let Some((_, value)) = collecting_data_label_text.as_mut() {
                    value.push_str(text);
                } else if let Some((_, value)) = collecting_format_code.as_mut() {
                    value.push_str(text);
                } else if let Some((_, value)) = collecting_value.as_mut() {
                    value.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    // Resolve fields only after the series caches are complete; rich labels often precede c:val.
    // Reverse ranges keep earlier offsets stable when replacement lengths differ.
    for (series_index, point_index, range, field) in data_label_fields.into_iter().rev() {
        let Some(series) = completed.get_mut(series_index) else {
            continue;
        };
        let value = series.values.get(point_index).copied();
        let replacement = match field.as_str() {
            "VALUE" => value.map(|v| format_chart_value(v, series.number_format.as_deref())),
            "CATEGORYNAME" => series.categories.get(point_index).cloned(),
            "SERIESNAME" => Some(series.name.clone()),
            "PERCENTAGE" => series
                .pie_percentages()
                .get(point_index)
                .map(|v| format!("{v}%")),
            _ => None,
        };
        if let (Some(text), Some(label)) = (
            replacement,
            series
                .data_labels
                .iter_mut()
                .find(|l| l.index == point_index),
        ) {
            if label.text.get(range.clone()).is_some() {
                label.text.replace_range(range, &text);
            }
        }
    }
    if title_all_caps {
        title = title.to_uppercase();
    }
    // A present title can use Office's automatic text without authored text runs.
    show_title &= !title_deleted;
    let color_count = source_series_indices.iter().copied().max().unwrap_or(0) as f32 + 1.0;
    for (parsed_index, series) in completed.iter_mut().enumerate() {
        let index = source_series_indices[parsed_index];
        if series.color.is_none()
            && let Some(accent) = chart_style.and_then(|style| {
                let index = style.checked_sub(2)? % 8;
                CHART_ACCENTS.get(index.checked_sub(1)? as usize)
            })
        {
            series.color = scheme_color(accent).map(|mut color| {
                if chart_style.is_some_and(|style| (27..=32).contains(&style)) {
                    let tint = (index as f32 + 1.0) / (color_count + 1.0) * 1.4 - 0.7;
                    color = transform_luminance(color, 1.0 - tint.abs(), tint.max(0.0));
                }
                if chart_style.is_some_and(|style| (11..=16).contains(&style)) && index == 1 {
                    color = transform_luminance(color, 0.6, 0.4);
                }
                color
            });
        }
        if chart_style == Some(1) && !series.point_colors.is_empty() {
            let base = scheme_color("dk1").unwrap_or(0x0000_00ff);
            series.point_colors = [0.60, 0.82, 0.68]
                .into_iter()
                .cycle()
                .take(series.values.len())
                .map(|tint| {
                    let mut color = base;
                    apply_color_transform(&mut color, "tint", tint);
                    color
                })
                .collect();
        }
        if series.color.is_none() {
            series.color = scheme_color(CHART_ACCENTS[index % CHART_ACCENTS.len()]).or_else(|| {
                (index != parsed_index).then(|| super::office_chart_palette_color(index))
            });
        }
        if filled_radars.contains(&parsed_index) && series.fill.is_none() {
            series.fill = Some(ChartFill::Solid(series.color.unwrap_or_else(|| super::office_chart_palette_color(index))));
        }
    }
    let plot_area_color = chart_style
        .filter(|style| !chart_style_is_extended && *style > 32)
        .map(|_| CLASSIC_DARK_CHART_PLOT_COLOR);
    // With no chart style, the classic Office defaults apply. Keep authored
    // text and line properties, including explicit false/noFill, authoritative.
    if chart_style.is_none() {
        chart_area_border.get_or_insert((ChartFill::Solid(0x8080_80ff), 0.75));
        for axis in [&mut value_axis_options, &mut horizontal_axis_options] {
            if axis.major_gridlines { axis.major_gridline_color.get_or_insert(0x8686_86ff); }
        }
    }
    // 3-D columns use a perspective camera unless explicitly overridden.
    // Keep the existing defaults for other chart families.
    if !right_angle_axes_explicit && completed.iter().any(ChartSeries::projected_column) {
        if let Some(view) = &mut view_3d { view.right_angle_axes = false; }
    }
    Ok((!completed.is_empty()).then(|| Chart {
        source_part: part.to_owned(),
        date_1904,
        series: completed,
        show_title,
        title,
        title_fill,
        title_position,
        title_font_size,
        title_font_bold,
        title_text_color,
        show_legend,
        legend_position,
        deleted_legend_entries,
        chart_area_no_fill,
        chart_area_fill,
        chart_area_border,
        plot_area_fill,
        legend_fill,
        legend_stroke,
        legend_stroke_width,
        plot_area_color,
        plot_area_border_color: None,
        plot_area_border_width: 0.0,
        data_label_position: None,
        style: None,
        classic_defaults: !chart_style_is_extended,
        font_family: chart_font_family,
        font_size: chart_font_size,
        font_bold: chart_font_bold,
        font_scale_basis: None,
        native_size: None,
        value_axis_options,
        secondary_value_axis_options,
        horizontal_axis_options,
        series_axis_options,
        scatter_has_lines,
        series_lines,
        data_table,
        up_down_bars,
        plot_bounds,
        legend_bounds,
        view_3d,
        surface_band_fills,
        bar_gap_width_percent,
        category_gap_width: None,
    }))
}

fn format_chart_category(
    value: &str,
    format_code: Option<&str>,
    date_1904: bool,
) -> Option<String> {
    let format_code = format_code?.to_ascii_lowercase();
    if !format_code.contains('y') || !format_code.contains('m') {
        return None;
    }
    let serial = value.parse::<f64>().ok()?;
    if !serial.is_finite() || !(-2_000_000.0..=2_000_000.0).contains(&serial) {
        return None;
    }
    let serial_day = serial.floor() as i64 - i64::from(date_1904);
    let (year, month, day) = excel_serial_date(serial_day, date_1904);
    if !(1..=9999).contains(&year) {
        return None;
    }
    if format_code.contains("mmm") {
        const MONTHS: [&str; 12] = [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        ];
        return MONTHS.get(month.saturating_sub(1) as usize).map(|month| {
            if format_code.contains('d') {
                format!("{day}-{month}-{:02}", year % 100)
            } else {
                format!("{month} {year:04}")
            }
        });
    }
    format_code
        .contains('d')
        .then(|| format!("{month}/{day}/{year:04}"))
}

pub(super) fn parse_basic_chart(
    package: &Package<'_>,
    part: &str,
) -> Result<Option<Chart>, Diagnostic> {
    parse_chart(package, part, default_scheme_color)
}

pub(super) fn parse_chart_ex(
    package: &Package<'_>,
    part: &str,
    scheme_color: impl Fn(&str) -> Option<u32>,
) -> Result<Option<Chart>, Diagnostic> {
    parse_chart_ex_with_data(package, part, scheme_color, |_| None)
}

pub(super) fn parse_chart_ex_with_data(
    package: &Package<'_>,
    part: &str,
    scheme_color: impl Fn(&str) -> Option<u32>,
    resolve_data: impl Fn(&str) -> Option<ChartSourceData>,
) -> Result<Option<Chart>, Diagnostic> {
    parse_chart_ex_impl(package, part, &scheme_color, &resolve_data)
}

#[inline(never)]
fn parse_chart_ex_impl(
    package: &Package<'_>,
    part: &str,
    scheme_color: &dyn Fn(&str) -> Option<u32>,
    resolve_data: &dyn Fn(&str) -> Option<ChartSourceData>,
) -> Result<Option<Chart>, Diagnostic> {
    struct SeriesState {
        depth: usize,
        kind: Option<ChartKind>,
        name_depth: Option<usize>,
        name: String,
        data_id: Option<String>,
        subtotals_depth: Option<usize>,
        subtotals: Vec<usize>,
        data_label_position: Option<String>,
        binning: ChartBinning,
        show_values: bool,
        fill: Option<ChartFill>,
        point_fills: HashMap<usize, ChartFill>,
        stroke_color: Option<u32>,
        stroke_width: f32,
        point_borders: HashMap<usize, (Option<u32>, f32)>,
    }

    let bytes = package.required_part(part)?;
    let mut depth = 0_usize;
    let mut data_depth = None;
    let mut data_id = None;
    let mut numeric_dimension_depth = None;
    let mut numeric_formula_depth = None;
    let mut numeric_formula = String::new();
    let mut string_formula = String::new();
    let mut string_formula_depth = None;
    let mut category_levels = HashMap::<String, Vec<Vec<String>>>::new();
    let mut string_dimension_depth = None;
    let mut string_level_depth = None;
    let mut point = None::<(usize, usize, String)>;
    let mut string_point = None::<(usize, usize, String)>;
    let mut points = Vec::new();
    let mut string_points = Vec::new();
    let mut data = HashMap::<String, Vec<f32>>::new();
    let mut categories = HashMap::<String, Vec<String>>::new();
    let mut series = None::<SeriesState>;
    let mut completed = Vec::new();
    let mut pareto_owners = Vec::new();
    let mut show_title = false;
    let mut title_depth = None;
    let mut title_text_depth = None;
    let mut title = String::new();
    let mut title_fallback = String::new();
    let mut title_properties_depth = None;
    let mut show_legend = false;
    let mut legend_position = ChartLegendPosition::default();
    let mut data_label_position = None;
    let mut category_gap_width = None;
    let mut shape_properties = None;
    let mut line_depth = None;
    let mut line_width = 1.0;
    let mut point_style = None;
    let mut fill_capture = None::<ChartFillCapture>;
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                match local {
                    "data" if data_depth.is_none() => {
                        data_id = string_attribute(&attributes, "id", part)?;
                        data_depth = (!empty).then_some(depth);
                        points.clear();
                        numeric_formula.clear();
                        string_formula.clear();
                    }
                    "numDim"
                        if data_depth.is_some()
                            && matches!(
                                string_attribute(&attributes, "type", part)?.as_deref(),
                                Some("val" | "size")
                            ) =>
                    {
                        numeric_dimension_depth = (!empty).then_some(depth);
                    }
                    "f" if numeric_dimension_depth.is_some() => {
                        numeric_formula_depth = (!empty).then_some(depth);
                    }
                    "f" if string_dimension_depth.is_some() => {
                        string_formula_depth = (!empty).then_some(depth);
                    }
                    "strDim"
                        if data_depth.is_some()
                            && string_attribute(&attributes, "type", part)?.as_deref()
                                == Some("cat") =>
                    {
                        string_dimension_depth = (!empty).then_some(depth);
                    }
                    "lvl" if string_dimension_depth.is_some() && string_level_depth.is_none() => {
                        string_level_depth = (!empty).then_some(depth);
                        string_points.clear();
                    }
                    "pt" if numeric_dimension_depth.is_some() && point.is_none() => {
                        let index = numeric_attribute(&attributes, "idx", part)?
                            .and_then(|value| usize::try_from(value).ok())
                            .unwrap_or(points.len());
                        if !empty {
                            point = Some((depth, index, String::new()));
                        }
                    }
                    "pt" if string_level_depth.is_some() && string_point.is_none() => {
                        let index = numeric_attribute(&attributes, "idx", part)?
                            .and_then(|value| usize::try_from(value).ok())
                            .unwrap_or(string_points.len());
                        if !empty {
                            string_point = Some((depth, index, String::new()));
                        }
                    }
                    "series" if series.is_none() => {
                        let layout_id = string_attribute(&attributes, "layoutId", part)?;
                        if layout_id.as_deref() == Some("paretoLine") {
                            pareto_owners.push(
                                numeric_attribute(&attributes, "ownerIdx", part)?.unwrap_or(0)
                                    as usize,
                            );
                        }
                        let kind = match layout_id.as_deref() {
                            Some("waterfall") => Some(ChartKind::Waterfall),
                            Some("treemap") => Some(ChartKind::Treemap),
                            Some("boxWhisker") => Some(ChartKind::BoxWhisker),
                            Some("funnel") => Some(ChartKind::Funnel),
                            Some("sunburst") => Some(ChartKind::Sunburst),
                            Some("clusteredColumn") => Some(ChartKind::Histogram),
                            _ => None,
                        };
                        let kind = if string_attribute(&attributes, "hidden", part)?.as_deref()
                            == Some("1")
                        {
                            None
                        } else {
                            kind
                        };
                        if !empty {
                            series = Some(SeriesState {
                                depth,
                                kind,
                                name_depth: None,
                                name: String::new(),
                                data_id: None,
                                subtotals_depth: None,
                                subtotals: Vec::new(),
                                data_label_position: None,
                                binning: ChartBinning::default(),
                                show_values: false,
                                fill: None,
                                point_fills: HashMap::new(),
                                stroke_color: None,
                                stroke_width: 1.0,
                                point_borders: HashMap::new(),
                            });
                        }
                    }
                    "binning" if series.is_some() => {
                        let binning = &mut series.as_mut().unwrap().binning;
                        binning.right_closed =
                            string_attribute(&attributes, "intervalClosed", part)?.as_deref()
                                != Some("l");
                        binning.underflow = float_attribute(&attributes, "underflow", part)?;
                        binning.overflow = float_attribute(&attributes, "overflow", part)?;
                    }
                    "binSize" if series.is_some() => {
                        series.as_mut().unwrap().binning.size =
                            float_attribute(&attributes, "val", part)?.filter(|v| *v > 0.0);
                    }
                    "binCount" if series.is_some() => {
                        series.as_mut().unwrap().binning.count =
                            numeric_attribute(&attributes, "val", part)?
                                .map(|v| v as usize)
                                .filter(|v| *v > 0);
                    }
                    "aggregation" if series.is_some() => {
                        series.as_mut().unwrap().binning.aggregate = true;
                    }
                    "dataPt" if series.is_some() => {
                        point_style = numeric_attribute(&attributes, "idx", part)?
                            .map(|v| (depth, v as usize));
                    }
                    "spPr"
                        if series.as_ref().is_some_and(|s| {
                            depth == point_style.map_or(s.depth, |(d, _)| d) + 1
                        }) =>
                    {
                        shape_properties = (!empty).then_some(depth);
                    }
                    "solidFill" | "gradFill" | "pattFill" | "blipFill"
                        if shape_properties.is_some_and(|d| depth == d + 1)
                            || line_depth.is_some_and(|d| depth == d + 1) =>
                    {
                        fill_capture = ChartFillCapture::new(local, depth);
                    }
                    "ln" if shape_properties.is_some_and(|d| depth == d + 1) => {
                        line_depth = (!empty).then_some(depth);
                        line_width = float_attribute(&attributes, "w", part)?
                            .map_or(1.0, |v| v / EMU_PER_CSS_PIXEL)
                            .max(0.0);
                    }
                    "noFill" if shape_properties.is_some_and(|d| depth == d + 1) => {
                        if let Some(series) = series.as_mut() {
                            if let Some((_, index)) = point_style {
                                series.point_fills.insert(index, ChartFill::Solid(0));
                            } else {
                                series.fill = Some(ChartFill::Solid(0));
                            }
                        }
                    }
                    "visibility" if series.is_some() => {
                        series.as_mut().unwrap().show_values =
                            string_attribute(&attributes, "value", part)?.as_deref() == Some("1");
                    }
                    "dataLabels" if series.is_some() => {
                        series
                            .as_mut()
                            .ok_or_else(|| format_error(part, "ChartEx series state is missing"))?
                            .data_label_position = string_attribute(&attributes, "pos", part)?;
                    }
                    "v" if series.is_some() => {
                        if !empty {
                            let series = series.as_mut().ok_or_else(|| {
                                format_error(part, "ChartEx series state is missing")
                            })?;
                            series.name_depth = Some(depth);
                            series.name.clear();
                        }
                    }
                    "dataId" if series.is_some() => {
                        let value = string_attribute(&attributes, "val", part)?;
                        series
                            .as_mut()
                            .ok_or_else(|| format_error(part, "ChartEx series state is missing"))?
                            .data_id = value;
                    }
                    "subtotals" if series.is_some() => {
                        if !empty {
                            series
                                .as_mut()
                                .ok_or_else(|| {
                                    format_error(part, "ChartEx series state is missing")
                                })?
                                .subtotals_depth = Some(depth);
                        }
                    }
                    "idx"
                        if series
                            .as_ref()
                            .is_some_and(|series| series.subtotals_depth.is_some()) =>
                    {
                        if let Some(index) = numeric_attribute(&attributes, "val", part)?
                            .and_then(|value| usize::try_from(value).ok())
                        {
                            series
                                .as_mut()
                                .ok_or_else(|| {
                                    format_error(part, "ChartEx series state is missing")
                                })?
                                .subtotals
                                .push(index);
                        }
                    }
                    "title" if series.is_none() => {
                        show_title = true;
                        title_depth = (!empty).then_some(depth);
                    }
                    "txPr" if title_depth.is_some() && series.is_none() => {
                        title_properties_depth = (!empty).then_some(depth);
                    }
                    "t" | "v" if title_depth.is_some() && series.is_none() => {
                        title_text_depth = (!empty).then_some(depth);
                    }
                    "catScaling" if series.is_none() => {
                        category_gap_width = float_attribute(&attributes, "gapWidth", part)?
                            .filter(|value| *value >= 0.0);
                    }
                    "legend" if series.is_none() => {
                        show_legend = true;
                        legend_position = ChartLegendPosition::from_ooxml(
                            string_attribute(&attributes, "pos", part)?.as_deref(),
                        );
                    }
                    _ => {}
                }
                if let Some(capture) = fill_capture.as_mut() {
                    capture.start(local, &attributes, part, &scheme_color)?;
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if let Some(capture) = fill_capture.as_mut() {
                    capture.end(local);
                }
                if fill_capture.as_ref().is_some_and(|c| c.closes_at(depth)) {
                    if let Some(fill) = fill_capture.take().unwrap().finish(package, part)? {
                        if let Some(series) = series.as_mut() {
                            if line_depth.is_some() {
                                let color = if let ChartFill::Solid(color) = fill {
                                    Some(color)
                                } else {
                                    None
                                };
                                if let Some((_, index)) = point_style {
                                    series.point_borders.insert(index, (color, line_width));
                                } else {
                                    series.stroke_color = color;
                                    series.stroke_width = line_width;
                                }
                            } else if let Some((_, index)) = point_style {
                                series.point_fills.insert(index, fill);
                            } else {
                                series.fill = Some(fill);
                            }
                        }
                    }
                }
                if local == "txPr" && title_properties_depth == Some(depth) {
                    title_properties_depth = None;
                }
                if local == "ln" && line_depth == Some(depth) {
                    line_depth = None;
                }
                if local == "spPr" && shape_properties == Some(depth) {
                    shape_properties = None;
                }
                if local == "dataPt" && point_style.is_some_and(|(d, _)| d == depth) {
                    point_style = None;
                }
                if local == "pt" && point.as_ref().is_some_and(|(start, _, _)| *start == depth) {
                    let (_, index, value) = point
                        .take()
                        .ok_or_else(|| format_error(part, "ChartEx point state is missing"))?;
                    if let Ok(value) = value.trim().parse::<f32>()
                        && value.is_finite()
                    {
                        points.push((index, value));
                    }
                } else if local == "pt"
                    && string_point
                        .as_ref()
                        .is_some_and(|(start, _, _)| *start == depth)
                {
                    let (_, index, value) = string_point.take().ok_or_else(|| {
                        format_error(part, "ChartEx string point state is missing")
                    })?;
                    string_points.push((index, value));
                } else if local == "f" && numeric_formula_depth == Some(depth) {
                    numeric_formula_depth = None;
                } else if local == "numDim" && numeric_dimension_depth == Some(depth) {
                    numeric_dimension_depth = None;
                } else if local == "lvl" && string_level_depth == Some(depth) {
                    string_points.sort_unstable_by_key(|(index, _)| *index);
                    if let Some(id) = data_id.as_ref() {
                        let level = string_points
                            .drain(..)
                            .map(|(_, value)| value)
                            .collect::<Vec<_>>();
                        categories
                            .entry(id.clone())
                            .or_insert_with(|| level.clone());
                        category_levels
                            .entry(id.clone())
                            .or_default()
                            .insert(0, level);
                    }
                    string_level_depth = None;
                } else if local == "f" && string_formula_depth == Some(depth) {
                    string_formula_depth = None;
                } else if local == "strDim" && string_dimension_depth == Some(depth) {
                    string_dimension_depth = None;
                } else if local == "data" && data_depth == Some(depth) {
                    points.sort_unstable_by_key(|(index, _)| *index);
                    if let Some(id) = data_id.take() {
                        let values = if points.is_empty() {
                            resolve_data(numeric_formula.trim())
                                .map(|data| data.values)
                                .unwrap_or_default()
                        } else {
                            points.drain(..).map(|(_, value)| value).collect()
                        };
                        if !category_levels.contains_key(&id) {
                            if let Some(source) = resolve_data(string_formula.trim()) {
                                categories.insert(
                                    id.clone(),
                                    source.category_levels.last().cloned().unwrap_or_default(),
                                );
                                category_levels.insert(id.clone(), source.category_levels);
                            }
                        }
                        data.insert(id, values);
                    }
                    data_depth = None;
                } else if matches!(local, "t" | "v") && title_text_depth == Some(depth) {
                    title_text_depth = None;
                } else if local == "title" && title_depth == Some(depth) {
                    title_depth = None;
                } else if local == "v"
                    && series
                        .as_ref()
                        .is_some_and(|series| series.name_depth == Some(depth))
                {
                    series
                        .as_mut()
                        .ok_or_else(|| format_error(part, "ChartEx series state is missing"))?
                        .name_depth = None;
                } else if local == "subtotals"
                    && series
                        .as_ref()
                        .is_some_and(|series| series.subtotals_depth == Some(depth))
                {
                    series
                        .as_mut()
                        .ok_or_else(|| format_error(part, "ChartEx series state is missing"))?
                        .subtotals_depth = None;
                } else if local == "series"
                    && series.as_ref().is_some_and(|series| series.depth == depth)
                {
                    let mut current = series
                        .take()
                        .ok_or_else(|| format_error(part, "ChartEx series state is missing"))?;
                    current.subtotals.sort_unstable();
                    current.subtotals.dedup();
                    if data_label_position.is_none() {
                        data_label_position.clone_from(&current.data_label_position);
                    }
                    if let (Some(kind), Some(values)) = (
                        current.kind,
                        current.data_id.as_deref().and_then(|id| data.get(id)),
                    ) && !values.is_empty()
                    {
                        let mut labels = current
                            .data_id
                            .as_deref()
                            .and_then(|id| categories.get(id))
                            .cloned()
                            .unwrap_or_default();
                        let values = if kind == ChartKind::Histogram {
                            let (values, bins) = chart_histogram(
                                values,
                                &labels,
                                &current.binning,
                                package.limits().max_document_objects,
                            )
                            .map_err(|error| with_part(error, part))?;
                            labels = bins;
                            values
                        } else {
                            values.clone()
                        };
                        let point_fills = (0..values.len())
                            .map(|i| current.point_fills.remove(&i))
                            .collect();
                        let borders = (0..values.len())
                            .map(|i| {
                                current
                                    .point_borders
                                    .get(&i)
                                    .copied()
                                    .unwrap_or((current.stroke_color, current.stroke_width))
                            })
                            .collect::<Vec<_>>();
                        completed.push(ChartSeries {
                            kind,
                            first_slice_angle: 0.0,
                            grouping: ChartGrouping::Standard,
                            bar_horizontal: false,
                            bar_depth: false,
                            bar_cone: false,
                            bar_cone_to_max: false,
                            bar_cylinder: false,
                            bar_gap_depth_percent: 150.0,
                            three_d: false,
                            axis_id: None,
                            name: current.name,
                            category_levels: current
                                .data_id
                                .as_ref()
                                .and_then(|id| category_levels.get(id))
                                .cloned()
                                .unwrap_or_default(),
                            categories: if kind == ChartKind::BoxWhisker {
                                Vec::new()
                            } else if labels.is_empty() {
                                (1..=values.len()).map(|i| i.to_string()).collect()
                            } else {
                                labels
                            },
                            x_values: Vec::new(),
                            values,
                            raw_values: Vec::new(),
                            bubble_sizes: Vec::new(),
                            color: (kind != ChartKind::Sunburst)
                                .then(|| scheme_color("accent1").unwrap_or(0x4472_c4ff)),
                            fill: current.fill,
                            point_colors: Vec::new(),
                            point_fills,
                            point_explosions: Vec::new(),
                            effects: Default::default(),
                            point_border_colors: borders.iter().map(|b| b.0).collect(),
                            point_border_widths: borders.iter().map(|b| b.1).collect(),
                            stroke_width: None,
                            line_visible: true,
                            smooth: false,
                            marker_symbol: None,
                            marker_size: None,
                            subtotals: current.subtotals,
                            show_values: current.show_values,
                            show_category_name: false,
                            show_percent: false,
                            number_format: None,
                            data_label_position: current.data_label_position,
                            data_label_text_color: None,
                            data_label_font_size: None,
                            data_label_font_bold: None,
                            data_label_rotation_degrees: None,
                            hidden_labels: Vec::new(),
                            data_labels: Vec::new(),
                            data_label_border: None,
                            linear_trendline: false,
                            show_trendline_equation: false,
                            show_trendline_r_squared: false,
                            trendline_label_offset: None,
                            trendline_label_font_size: None,
                            x_error_bars: None,
                            y_error_bars: None,
                        });
                    }
                }
            }
            XmlEvent::Text(text) => {
                let text = decode_xml_text(text).map_err(|error| with_part(error, part))?;
                if let Some((_, _, value)) = point.as_mut() {
                    value.push_str(&text);
                } else if string_formula_depth.is_some() {
                    string_formula.push_str(&text);
                } else if numeric_formula_depth.is_some() {
                    numeric_formula.push_str(&text);
                } else if let Some((_, _, value)) = string_point.as_mut() {
                    value.push_str(&text);
                } else if title_text_depth.is_some() {
                    if title_properties_depth.is_some() {
                        title_fallback.push_str(&text);
                    } else {
                        title.push_str(&text);
                    }
                } else if let Some(series) =
                    series.as_mut().filter(|series| series.name_depth.is_some())
                {
                    series.name.push_str(&text);
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some((_, _, value)) = point.as_mut() {
                    value.push_str(text);
                } else if string_formula_depth.is_some() {
                    string_formula.push_str(&text);
                } else if numeric_formula_depth.is_some() {
                    numeric_formula.push_str(text);
                } else if let Some((_, _, value)) = string_point.as_mut() {
                    value.push_str(text);
                } else if title_text_depth.is_some() {
                    if title_properties_depth.is_some() {
                        title_fallback.push_str(text);
                    } else {
                        title.push_str(text);
                    }
                } else if let Some(series) =
                    series.as_mut().filter(|series| series.name_depth.is_some())
                {
                    series.name.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    for index in pareto_owners {
        if let Some(series) = completed
            .get_mut(index)
            .filter(|s| s.kind == ChartKind::Histogram)
        {
            let mut order = (0..series.values.len()).collect::<Vec<_>>();
            order.sort_by(|&a, &b| series.values[b].total_cmp(&series.values[a]));
            series.values = order.iter().map(|&i| series.values[i]).collect();
            series.categories = order
                .iter()
                .map(|&i| series.categories[i].clone())
                .collect();
            series.kind = ChartKind::Pareto;
        }
    }
    let style = chart_ex_style(package, part, &scheme_color)?;
    if title.trim().is_empty() {
        title = title_fallback;
    }
    if show_title && title.trim().is_empty() {
        title = "Chart Title".to_owned();
    }
    Ok((!completed.is_empty()).then(|| Chart {
        source_part: part.to_owned(),
        date_1904: false,
        series: completed,
        show_title,
        title,
        title_fill: None,
        title_position: None,
        title_font_size: None,
        title_font_bold: None,
        title_text_color: None,
        show_legend,
        legend_position,
        deleted_legend_entries: Vec::new(),
        chart_area_no_fill: false,
        chart_area_fill: None,
        chart_area_border: None,
        plot_area_fill: None,
        legend_fill: None,
        legend_stroke: None,
        legend_stroke_width: 0.0,
        plot_area_color: None,
        plot_area_border_color: None,
        plot_area_border_width: 0.0,
        data_label_position,
        style,
        classic_defaults: false,
        font_family: None,
        font_size: None,
        font_bold: false,
        font_scale_basis: None,
        native_size: None,
        value_axis_options: ChartValueAxis::default(),
        secondary_value_axis_options: None,
        horizontal_axis_options: ChartValueAxis::default(),
        series_axis_options: None,
        scatter_has_lines: false,
        series_lines: false,
        data_table: None,
        up_down_bars: None,
        plot_bounds: None,
        legend_bounds: None,
        view_3d: None,
        surface_band_fills: Vec::new(),
        bar_gap_width_percent: 150.0,
        category_gap_width,
    }))
}

fn chart_ex_style(
    package: &Package<'_>,
    chart_part: &str,
    scheme_color: &impl Fn(&str) -> Option<u32>,
) -> Result<Option<ChartStyle>, Diagnostic> {
    let relationships = package.relationships(Some(chart_part))?;
    let style_part = relationships.iter().find(|relationship| {
        !relationship.external && relationship.type_uri.ends_with("/chartStyle")
    });
    let Some(style_part) = style_part else {
        return Ok(None);
    };
    let Some(style_id @ (372 | 395)) = chart_style_id(package, &style_part.target)? else {
        return Ok(None);
    };
    let (method, colors) = relationships
        .iter()
        .find(|relationship| {
            !relationship.external && relationship.type_uri.ends_with("/chartColorStyle")
        })
        .map(|relationship| chart_color_style(package, &relationship.target, scheme_color))
        .transpose()?
        .unwrap_or_default();
    let role_colors = if method.as_deref() == Some("withinLinear")
        && let Some(base) = colors.first().copied()
    {
        within_linear_waterfall_colors(base)
    } else {
        [
            colors
                .first()
                .copied()
                .or_else(|| scheme_color("accent1"))
                .unwrap_or(0x5b9b_d5ff),
            colors
                .get(1)
                .copied()
                .or_else(|| scheme_color("accent2"))
                .unwrap_or(0xed7d_31ff),
            colors
                .get(2)
                .copied()
                .or_else(|| scheme_color("accent3"))
                .unwrap_or(0xa5a5_a5ff),
        ]
    };
    Ok(Some(match style_id {
        372 => ChartStyle::office_372(role_colors, scheme_color),
        395 => ChartStyle::office_395(role_colors, scheme_color),
        _ => unreachable!("style id was constrained above"),
    }))
}

fn within_linear_waterfall_colors(base: u32) -> [u32; 3] {
    // ponytail: ChartEx waterfall has three semantic roles; generalize the ramp when another
    // chart type needs a different object count.
    let transform = |kind, ratio| {
        let mut color = base;
        apply_color_transform(&mut color, kind, ratio);
        color
    };
    [
        transform("shade", 0.58),
        transform("shade", 0.86),
        transform("tint", 0.86),
    ]
}

fn chart_style_id(package: &Package<'_>, part: &str) -> Result<Option<u64>, Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut id = None;
    parse_xml(&bytes, package.limits(), |event| {
        if let XmlEvent::StartElement {
            name, attributes, ..
        } = event
            && local_name(name) == "chartStyle"
        {
            id = numeric_attribute(&attributes, "id", part)?;
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok(id)
}

fn chart_color_style(
    package: &Package<'_>,
    part: &str,
    scheme_color: &impl Fn(&str) -> Option<u32>,
) -> Result<(Option<String>, Vec<u32>), Diagnostic> {
    let bytes = package.required_part(part)?;
    let mut depth = 0_usize;
    let mut root_depth = None;
    let mut method = None;
    let mut colors = Vec::new();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "colorStyle" {
                    method = string_attribute(&attributes, "meth", part)?;
                    root_depth = (!empty).then_some(depth);
                } else if local == "schemeClr"
                    && root_depth.is_some_and(|root| depth == root + 1)
                    && let Some(color) = string_attribute(&attributes, "val", part)?
                        .as_deref()
                        .and_then(scheme_color)
                {
                    colors.push(color);
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if local_name(name) == "colorStyle" && root_depth == Some(depth) {
                    root_depth = None;
                }
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| with_part(error, part))?;
    Ok((method, colors))
}

fn chart_theme_effects(
    package: &Package<'_>,
    chart_part: &str,
    theme_part: Option<&str>,
    index: u64,
    color: &dyn Fn(&str) -> Option<u32>,
) -> Result<Option<DrawingMlPictureEffects>, Diagnostic> {
    let mut theme = None;
    if let Some(main) = package
        .relationships(None)?
        .iter()
        .find(|r| !r.external && r.type_uri.ends_with("/officeDocument"))
    {
        theme = package
            .relationships(Some(&main.target))?
            .into_iter()
            .find(|r| !r.external && r.type_uri.ends_with("/theme"))
            .map(|r| r.target);
    }
    if let Some(part) = theme_part {
        theme = Some(part.to_owned());
    }
    let override_part = package
        .relationships(Some(chart_part))?
        .into_iter()
        .find(|r| !r.external && r.type_uri.ends_with("/themeOverride"))
        .map(|r| r.target);
    let mut result = None;
    for part in theme.iter().chain(override_part.iter()) {
        if let Some(effects) = drawingml_theme_effects(package, part, index, |kind, value| {
            Ok(if kind == "schemeClr" {
                color(value).unwrap_or(0x000000ff)
            } else {
                parse_rgb_color(value).unwrap_or(0x000000ff)
            })
        })? {
            result = Some(effects);
        }
    }
    Ok(result)
}
