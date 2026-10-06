#[cfg(test)]
mod tests {
    use super::{
        CLASSIC_DARK_CHART_PLOT_COLOR, Chart, ChartFill, ChartGrouping, ChartKind,
        ChartLegendPosition, ChartSeries, ChartSourceData, ChartValueAxis, Diagram, DiagramNode,
        DrawingMlPictureEffects, apply_color_transform, chart_box_whisker, chart_pie_label_layout,
        chart_pie_slices, chart_series_line_geometries, diagram_horizontal_list_elements,
        drawingml_fill_reference_has_paint, format_chart_category, parse_chart, parse_chart_ex,
        parse_chart_ex_with_data, parse_three_d_bevel, radar_geometry,
        within_linear_waterfall_colors,
    };
    use crate::format::presentation_image::stored_zip;
    use crate::limits::Limits;
    use crate::model::{Geometry, OuterShadow, Paint, PathCommand, Rect, Shadow, Visual};
    use crate::package::Package;
    use crate::xml::XmlAttribute;

    #[test]
    fn supplied_radar_style_preserves_bevel_shadow_and_series_shades() {
        let package = Package::open(include_bytes!("../../../tests/fixtures/chart-Radar.docx"), Limits::default()).unwrap();
        let chart = parse_chart(&package, "word/charts/chart1.xml", |name| (name == "accent2").then_some(0xc0504dff)).unwrap().unwrap();
        assert_ne!(chart.series[0].color, chart.series[1].color, "style 28 uses different shades of accent2");
        for series in &chart.series {
            assert!(series.effects.outer_shadow.is_some());
            assert!(series.effects.three_d.as_ref().unwrap().bevel_top.is_some());
        }
        let elements = super::chart_extended_elements(&chart, Rect { x: 0.0, y: 0.0, width: 290.0, height: 220.0 }, 10000).unwrap().unwrap();
        assert_eq!(elements.iter().filter(|e| matches!(&e.visual, Visual::AdvancedEffect { three_d: Some(_), outer_shadow: Some(_), .. })).count(), 4, "both polygons and legend keys retain the effects in every host");
        for (bytes, part) in [
            (include_bytes!("../../../tests/fixtures/chart-2d.pptx").as_slice(), "ppt/charts/chart1.xml"),
            (include_bytes!("../../../tests/fixtures/chart-empty-title-original.xlsx").as_slice(), "xl/charts/chart1.xml"),
        ] {
            let host = Package::open(bytes, Limits::default()).unwrap();
            let parts = host.entry_names().map(|name| (name, if name == part {
                package.required_part("word/charts/chart1.xml").unwrap().to_vec()
            } else if name.contains("/theme/theme") {
                package.required_part("word/theme/theme1.xml").unwrap().to_vec()
            } else { host.required_part(name).unwrap().to_vec() })).collect::<Vec<_>>();
            let zip = stored_zip(&parts.iter().map(|(name, data)| (*name, data.as_slice())).collect::<Vec<_>>());
            let document = crate::format::detect_and_parse(&zip, Limits::default()).unwrap().unwrap();
            let effects = document.objects.iter().filter(|o| o.source.part == part && matches!(&o.visual,
                Visual::AdvancedEffect { three_d: Some(_), outer_shadow: Some(_), .. })).count();
            assert_eq!(effects, 4, "{part}: both radar polygons and legend keys retain the theme effects");
        }
        let original = String::from_utf8(package.required_part("word/charts/chart1.xml").unwrap().to_vec()).unwrap();
        for effect in ["<a:effectLst/>", "<a:sp3d><a:bevelT w=\"19050\" h=\"9525\"/></a:sp3d>"] {
            let xml = original.replacen("<c:ser>", &format!("<c:ser><c:spPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill>{effect}</c:spPr>"), 1);
            let parts = package.entry_names().map(|name| (name, if name == "word/charts/chart1.xml" {
                xml.as_bytes().to_vec()
            } else { package.required_part(name).unwrap().to_vec() })).collect::<Vec<_>>();
            let zip = stored_zip(&parts.iter().map(|(name, data)| (*name, data.as_slice())).collect::<Vec<_>>());
            let modified = Package::open(&zip, Limits::default()).unwrap();
            let chart = parse_chart(&modified, "word/charts/chart1.xml", super::default_scheme_color).unwrap().unwrap();
            assert_eq!(chart.series[0].color, Some(0x00ff00ff), "authored fill wins over style palette");
            assert!(chart.series[0].effects.outer_shadow.is_none(), "explicit effects suppress theme defaults");
            if effect.contains("sp3d") {
                assert_eq!(chart.series[0].effects.three_d.as_ref().unwrap().bevel_top.as_ref().unwrap().width, 2.0);
            } else { assert!(chart.series[0].effects.three_d.is_none()); }
        }

    }

    #[test]
    fn supplied_chart3_preserves_manual_title_and_compact_axis_layout() {
        let package = Package::open(include_bytes!("../../../tests/fixtures/chart3.docx"), Limits::default()).unwrap();
        let parse = |index| parse_chart(&package, &format!("word/charts/chart{index}.xml"), |_| None).unwrap().unwrap();
        let line = parse(1);
        assert_eq!(line.value_axis_options.title, "Axis Title");
        let frame = Rect { x: 120.0, y: 180.0, width: 262.667, height: 183.697 };
        let title = line.positioned_title_bounds(frame, Rect { width: 100.0, height: 36.0, ..frame });
        assert!((title.x - 121.17).abs() < 0.1 && (title.y - 233.57).abs() < 0.1);
        assert_eq!(line.legend_entries()[0].1, "Series 2");
        let mut radar = parse(6);
        let bounds = Rect { x: 0.0, y: 0.0, width: 125.0, height: 124.0 };
        radar.resolve_value_axis_layout(bounds);
        assert_eq!(radar.title_text_style().font_size, 24.0);
        assert_eq!(radar.value_axis_ticks().iter().map(|(v, _)| *v).collect::<Vec<_>>(), [0.0, 50.0], "radar ticks occupy a radius, not its full diameter");
        assert_eq!(radar.constrained_legend(bounds).unwrap().2.len(), 1);
        let plot = radar.plot_area_bounds(bounds, Rect { x: 17.5, y: 21.08, width: 80.0, height: 86.8 });
        assert!(plot.height >= 30.0, "radial labels must not reserve Cartesian date-axis margins");
        let mut scatter = parse(5);
        scatter.resolve_value_axis_layout(Rect { width: 195.0, height: 101.0, ..bounds });
        assert!(scatter.scatter_axis_ticks(true).len() <= 3);
        let area = parse(3);
        let plot = Rect { width: 150.0, height: 70.0, ..bounds };
        assert!(super::chart_area_3d_axis_labels(&area, plot, 12.0).1.iter().all(|label| label.3 == -90.0));
    }

    #[test]
    fn supplied_chart4_first_chart_uses_projected_clustered_columns() {
        let package = Package::open(include_bytes!("../../../tests/fixtures/chart4.docx"), Limits::default()).unwrap();
        let mut chart = parse_chart(&package, "word/charts/chart1.xml", |_| None).unwrap().unwrap();
        let bounds = Rect { x: 0.0, y: 0.0, width: 518.0, height: 326.0 };
        chart.resolve_value_axis_layout(bounds);
        assert_eq!(chart.value_axis_ticks().iter().map(|(value, _)| chart.value_axis_label(*value)).collect::<Vec<_>>(), ["0", "1E+11"]);
        let plot = chart.plot_area_bounds(bounds, bounds);
        assert!(chart.depth_box_faces(0, 2, plot, plot).is_some(), "clustered columns must use the same projection as their axes");
        assert!(chart.title_text_style().bold, "classic chart title defaults to bold");
        assert_eq!(chart.title_text_style().font_size, 24.0, "native Word title is 18pt");
        let table = super::chart_data_table_layout(&chart, bounds, plot).unwrap();
        assert_eq!(table.lines.iter().filter(|l| l.series_index.is_none()).count(), 2, "only internal category separators");
        assert_eq!(table.lines.iter().filter(|l| matches!(l.geometry, Geometry::Rectangle)).count(), 3, "column keys are filled squares");
        assert_eq!(chart.series[0].value_label(2, chart.series[0].values[2], Some("General")), "56456456454");
        let document = crate::format::detect_and_parse(include_bytes!("../../../tests/fixtures/chart4.docx"), Limits::default()).unwrap().unwrap();
        let name = document.objects.iter().find(|o| o.source.part == "word/charts/chart1.xml" && o.text.as_deref() == Some("Series 1")).unwrap();
        assert!(matches!(&name.visual, Visual::TextLayout { layout, .. } if !layout.wrap && layout.inset_left == 0.0 && layout.inset_right == 0.0), "table names must fit on one line without shape text insets");
    }

    #[test]
    fn supplied_chart4_composite_and_radar_survive_pptx_and_xlsx_adapters() {
        let source = Package::open(include_bytes!("../../../tests/fixtures/chart4.docx"), Limits::default()).unwrap();
        for (bytes, part) in [
            (include_bytes!("../../../tests/fixtures/chart-2d.pptx").as_slice(), "ppt/charts/chart1.xml"),
            (include_bytes!("../../../tests/fixtures/chart-empty-title-original.xlsx").as_slice(), "xl/charts/chart1.xml"),
        ] {
            let host = Package::open(bytes, Limits::default()).unwrap();
            for chart_index in [1, 3, 5, 6] {
                let xml = source.required_part(&format!("word/charts/chart{chart_index}.xml")).unwrap();
                let parts = host.entry_names().map(|name| (name, if name == part { xml.to_vec() } else { host.required_part(name).unwrap().into_vec() })).collect::<Vec<_>>();
                let zip = stored_zip(&parts.iter().map(|(name, data)| (*name, data.as_slice())).collect::<Vec<_>>());
                let document = crate::format::detect_and_parse(&zip, Limits::default()).unwrap().unwrap();
                let objects = document.objects.iter().filter(|o| o.source.part == part).collect::<Vec<_>>();
                if chart_index == 1 {
                    assert!(objects.iter().any(|o| o.text.as_deref() == Some("1E+11")), "{part}: projected axis scale");
                    assert!(objects.iter().any(|o| o.text.as_deref() == Some("56456456454")), "{part}: full cached value precision");
                } else if chart_index == 3 {
                    for value in ["8", "20", "21", "22", "13", "9"] {
                        assert!(objects.iter().any(|o| o.text.as_deref() == Some(value)), "{part}: missing {value}");
                    }
                } else if chart_index == 5 {
                    assert!(objects.iter().any(|o| matches!(&o.visual, Visual::PaintedShape { geometry: Geometry::Path { commands, .. }, .. } if commands.iter().any(|c| matches!(c, PathCommand::BezierCurveTo { .. })))), "{part}: smooth scatter curve");
                    assert!(!objects.iter().any(|o| matches!(&o.visual, Visual::PaintedShape { geometry: Geometry::Ellipse, .. })), "{part}: explicit marker none");
                } else {
                    assert!(objects.iter().filter(|o| {
                        let visual = match &o.visual { Visual::AdvancedEffect { visual, .. } => visual.as_ref(), visual => visual };
                        matches!(visual, Visual::PaintedShape { geometry: Geometry::Path { .. }, fill: Paint::Solid(_), .. })
                    }).count() >= 2, "{part}: filled radar polygons");
                }
            }
        }
    }

    #[test]
    fn supplied_chart4_stacked_areas_share_depth_plane() {
        let package = Package::open(include_bytes!("../../../tests/fixtures/chart4.docx"), Limits::default()).unwrap();
        let chart = parse_chart(&package, "word/charts/chart4.xml", |_| None).unwrap().unwrap();
        let plot = Rect { x: 0.0, y: 0.0, width: 300.0, height: 150.0 };
        let first = super::chart_area_3d_faces(&chart, 0, plot, (0.0, 50.0), chart.view_3d.unwrap(), 0x0000_ffff);
        let second = super::chart_area_3d_faces(&chart, 1, plot, (0.0, 50.0), chart.view_3d.unwrap(), 0x00ff_00ff);
        assert_eq!(first[0].0, second[0].0, "stacked series must meet on the same depth plane");
    }

    #[test]
    fn supplied_chart4_preserves_chart_semantics() {
        let package = Package::open(include_bytes!("../../../tests/fixtures/chart4.docx"), Limits::default()).unwrap();
        let chart = |index| parse_chart(&package, &format!("word/charts/chart{index}.xml"), |_| None).unwrap().unwrap();
        assert_eq!(chart(1).title_text_style().color, 0xff00_00ff, "authored red title");
        assert_eq!(chart(3).series[0].kind, ChartKind::BarOfPie(2), "automatic split keeps three of five categories in the main pie");
        assert!(chart(6).series.iter().all(|s| s.fill.is_some()), "filled radar retains its fill");
        let bounds = Rect { x: 0.0, y: 0.0, width: 520.0, height: 326.0 };
        let first = chart(1);
        let plot = first.plot_area_bounds(bounds, bounds);
        let table = super::chart_data_table_layout(&first, bounds, plot).unwrap();
        assert!(table.texts.iter().all(|t| t.bounds.height <= 24.0), "manual plot must not stretch the table to the frame bottom");
    }

    #[test]
    fn real_olap_chart_automatic_labels_stay_inside_the_chart() {
        let package = Package::open(include_bytes!("../../../tests/fixtures/OlapPivotA3.xlsx"), Limits::default()).unwrap();
        let chart = parse_chart(&package, "xl/charts/chart1.xml", |_| None).unwrap().unwrap();
        assert_eq!(chart.series.len(), 207);
        let bounds = Rect { x: 24.0, y: 117.0, width: 480.0, height: 288.0 };
        let mut objects = Vec::new();
        super::push_spreadsheet_chart(chart, bounds, &crate::model::SourceRef {
            part: "xl/charts/chart1.xml".into(), mapping: crate::model::MappingQuality::Exact,
            locator: crate::model::SourceLocator::Flat { kind: "chart", index: None, row: None, column: None, text_range: None },
        }, 0, 100000, &mut objects, &mut Vec::new()).unwrap();
        let lines = objects.iter().filter_map(|object| {
            let Visual::AdvancedEffect { visual, .. } = &object.visual else { return None; };
            let Visual::PaintedShape { geometry: Geometry::Path { commands, .. }, fill: Paint::None, .. } = visual.as_ref() else { return None; };
            Some(commands.len())
        }).collect::<Vec<_>>();
        assert_eq!(lines.len(), 177, "one shadowed path per nonempty series; empty caches stay empty");
        assert_eq!(lines.iter().map(|count| count - 1).sum::<usize>(), 9552, "all original line segments survive batching");
        let labels = objects.iter().filter(|o| o.text.as_ref().is_some_and(|text| text.starts_with("Measure"))).collect::<Vec<_>>();
        assert_eq!(labels.len(), 6, "Excel shows the six legend entries that fit, without shrinking the font");
        for label in labels {
            assert!(label.bounds.x >= bounds.x && label.bounds.y >= bounds.y);
            assert!(label.bounds.x + label.bounds.width <= bounds.x + bounds.width);
            assert!(label.bounds.y + label.bounds.height <= bounds.y + bounds.height);
            assert!(label.text.as_ref().unwrap().contains('\n'));
        }
        let categories = objects.iter().filter(|o| o.text.as_ref().is_some_and(|text| text.starts_with("Level1Item"))).count();
        assert!((10..=20).contains(&categories), "automatic category labels must be thinned: {categories}");
    }

    #[test]
    fn camera_defaults_and_explicit_rotation_share_one_resolved_model() {
        let mut style = crate::model::ThreeDStyle::default();
        super::parse_three_d_camera(&mut style, &[XmlAttribute {
            name: "prst", value: "perspectiveHeroicExtremeLeftFacing",
        }], "test.xml").unwrap();
        assert_eq!((style.camera_latitude, style.camera_longitude, style.camera_revolution), (8.1, 34.5, 357.1));
        assert_eq!(style.camera_fov, 80.0);
        super::parse_three_d_rotation(&mut style, &[
            XmlAttribute { name: "lat", value: "0" },
            XmlAttribute { name: "lon", value: "0" },
            XmlAttribute { name: "rev", value: "0" },
        ], "test.xml", true).unwrap();
        assert_eq!((style.camera_latitude, style.camera_longitude, style.camera_revolution), (0.0, 0.0, 0.0));
    }

    #[test]
    fn real_3d_column_preserves_perspective() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/chart-3d-column.pptx"),
            Limits::default(),
        ).unwrap();
        let chart = parse_chart(&package, "ppt/charts/chart1.xml", |_| None).unwrap().unwrap();
        assert!(!chart.view_3d.unwrap().right_angle_axes,
            "omitted rAngAx must not disable the authored perspective");
        let part = package.part("ppt/charts/chart1.xml").unwrap().unwrap();
        let xml = std::str::from_utf8(&part).unwrap();
        for (element, expected) in [("<c:rAngAx/>", true), ("<c:rAngAx val=\"1\"/>", true), ("<c:rAngAx val=\"0\"/>", false)] {
            let changed = xml.replace("<c:perspective", &format!("{element}<c:perspective"));
            let bytes = stored_zip(&[("ppt/charts/chart1.xml", changed.as_bytes())]);
            let package = Package::open(&bytes, Limits::default()).unwrap();
            let explicit = parse_chart(&package, "ppt/charts/chart1.xml", |_| None).unwrap().unwrap();
            assert_eq!(explicit.view_3d.unwrap().right_angle_axes, expected);
        }
        let plot = Rect { x: 40.0, y: 50.0, width: 600.0, height: 360.0 };
        let axis = chart.value_axis();
        for series in 0..3 {
            for category in 0..4 {
                let bounds = super::chart_bar_segment_bounds(&chart, series, category, plot,
                    (axis.0, axis.1), chart.bar_gap_width_percent).unwrap();
                let faces = chart.depth_box_faces(series, category, plot, bounds).unwrap();
                for face in &faces {
                    let Geometry::Path { commands, .. } = face else { panic!("cuboid face must be projected"); };
                    assert_eq!(commands.len(), 5);
                    for command in commands {
                        if let PathCommand::MoveTo { x, y } | PathCommand::LineTo { x, y } = command {
                            assert!(*x >= -0.001 && *x <= bounds.width + 0.001);
                            assert!(*y >= -0.001 && *y <= bounds.height + 0.001);
                        }
                    }
                }
                let Geometry::Path { commands, .. } = &faces[0] else { unreachable!() };
                let (PathCommand::MoveTo { y: left, .. }, PathCommand::LineTo { y: right, .. }) = (&commands[0], &commands[1]) else { unreachable!() };
                assert!((left-right).abs() > 0.1, "front edge must follow the sloping floor");
            }
        }
        let near = super::chart_bar_segment_bounds(&chart, 0, 0, plot, (axis.0, axis.1), chart.bar_gap_width_percent).unwrap();
        let far = super::chart_bar_segment_bounds(&chart, 2, 0, plot, (axis.0, axis.1), chart.bar_gap_width_percent).unwrap();
        assert!(near.width > far.width, "distant columns must shrink");
        let (_, Geometry::Path { commands, .. }) = chart.bar_grid_line(plot, 0.0, true, false) else { panic!("projected grid"); };
        let PathCommand::MoveTo { x,y } = commands[0] else { unreachable!() };
        let floor_corner = super::chart_project_3d(plot, chart.view_3d.unwrap(), 0.0, 0.0, 0.0);
        assert!((plot.x+x-floor_corner.0).abs() < 0.001 && (plot.y+y-floor_corner.1).abs() < 0.001);
    }

    #[test]
    fn real_chart_2d_automatic_layout_and_spreadsheet_adapter() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/chart-2d.pptx"),
            Limits::default(),
        )
        .unwrap();
        let mut chart = parse_chart(&package, "ppt/charts/chart1.xml", |_| None)
            .unwrap()
            .unwrap();
        let bounds = Rect {
            x: 48.0,
            y: 168.0,
            width: 864.0,
            height: 475.1667,
        };
        let (plot, legend) = chart
            .automatic_cartesian_layout(bounds)
            .expect("automatic layout");
        assert!((plot.x - 91.0).abs() < 6.0);
        assert!((plot.width - 688.0).abs() < 12.0);
        assert!((plot.height - 401.0).abs() < 6.0);
        assert!(legend.unwrap().x > plot.x + plot.width);
        assert_eq!(chart.legend_text_style().font_size, 24.0);
        let mut objects = Vec::new();
        super::push_spreadsheet_chart(
            chart.clone(),
            bounds,
            &crate::model::SourceRef {
                part: "xl/charts/chart1.xml".into(),
                mapping: crate::model::MappingQuality::Exact,
                locator: crate::model::SourceLocator::Flat {
                    kind: "chart",
                    index: None,
                    row: None,
                    column: None,
                    text_range: None,
                },
            },
            0,
            10000,
            &mut objects,
            &mut Vec::new(),
        )
        .unwrap();
        let grid = objects
            .iter()
            .filter(|o| {
                matches!(
                    &o.visual,
                    Visual::PaintedShape {
                        geometry: Geometry::Line,
                        ..
                    }
                ) && o.bounds.width > 100.0
                    && o.bounds.height < 1.0
            })
            .collect::<Vec<_>>();
        assert!(grid.len() >= 7);
        assert!(
            grid.iter().filter(|o| (o.bounds.width - plot.width).abs() < 0.01).count() >= 7,
            "{:?}", grid.iter().map(|o| o.bounds).collect::<Vec<_>>()
        );
        let manual = Rect { x: 0.2, y: 0.2, width: 0.5, height: 0.5 };
        chart.plot_bounds = Some(manual);
        assert!(chart.automatic_cartesian_layout(bounds).is_none());
        assert_eq!(chart.plot_area_bounds(bounds, bounds).width, bounds.width * manual.width);
        chart.plot_bounds = None;
        chart.series[0].three_d = true;
        assert!(chart.automatic_cartesian_layout(bounds).is_none());
    }

    #[test]
    fn real_org_chart_3d_resolves_quick_style_through_shared_effects() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/smartart-orgchart-3d.pptx"),
            Limits::default(),
        )
        .unwrap();
        let relationships = package
            .relationships(Some("ppt/slides/slide1.xml"))
            .unwrap();
        let map = relationships.iter().map(|r| (r.id.as_str(), r)).collect();
        let diagram = super::parse_diagram(
            &package,
            "ppt/diagrams/data1.xml",
            &map,
            Some("ppt/diagrams/colors1.xml"),
            Some("ppt/theme/theme1.xml"),
            super::default_scheme_color,
        )
        .unwrap();
        assert!(diagram.diagnostics.is_empty());
        let elements = super::diagram_org_chart_elements(
            &diagram,
            Rect {
                x: 48.0,
                y: 168.0,
                width: 864.0,
                height: 475.16672,
            },
            0x4f81bdff,
        )
        .unwrap();
        let boxes = elements
            .iter()
            .filter(|e| e.text.is_some())
            .collect::<Vec<_>>();
        assert_eq!(boxes.len(), 9);
        let expected = [
            (365., 168.),
            (219., 269.),
            (254., 370.),
            (254., 471.),
            (511., 269.),
            (426., 370.),
            (426., 471.),
            (462., 572.),
            (598., 370.),
        ];
        for (element, (x, y)) in boxes.into_iter().zip(expected) {
            assert!(
                (element.bounds.x - x).abs() < 3.0 && (element.bounds.y - y).abs() < 3.0,
                "{:?}",
                element.bounds
            );
            let Visual::AdvancedEffect {
                three_d: Some(style),
                outer_shadow: Some(shadow),
                visual,
                ..
            } = &element.visual
            else {
                panic!("missing quick style: {:?}", element.visual);
            };
            assert_eq!(style.material, "plastic");
            assert_eq!(style.bevel_top.as_ref().unwrap().preset, "relaxedInset");
            assert_eq!(style.light_revolution, 125.0);
            assert_eq!(shadow.shadow.color, 0x00000059);
            let Visual::TextLayout { visual, .. } = visual.as_ref() else {
                panic!()
            };
            let Visual::RichText { fill, stroke, .. } = visual.as_ref() else {
                panic!()
            };
            assert!(matches!(fill, Paint::LinearGradient { .. }));
            assert!(matches!(stroke, Paint::None));
        }
    }

    #[test]
    fn real_org_chart_uses_assistants_styles_and_subtree_extents() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/smartart-orgchart.pptx"),
            Limits::default(),
        )
        .unwrap();
        let mut diagram = super::parse_diagram(
            &package,
            "ppt/diagrams/data1.xml",
            &Default::default(),
            Some("ppt/diagrams/colors1.xml"),
            None,
            |name| match name {
                "accent1" => Some(0x4f81_bdff),
                "lt1" => Some(0xffff_ffff),
                _ => None,
            },
        )
        .unwrap();
        assert_eq!(diagram.nodes.len(), 9);
        assert_eq!(
            diagram.nodes.iter().filter(|node| node.assistant).count(),
            3
        );
        let bounds = Rect {
            x: 48.0,
            y: 168.0,
            width: 864.0,
            height: 475.16672,
        };
        let render = |diagram: &Diagram| super::diagram_org_chart_elements(diagram, bounds, 0).unwrap();
        let elements = render(&diagram);
        let boxes = elements
            .iter()
            .filter(|e| e.text.is_some())
            .collect::<Vec<_>>();
        assert_eq!(boxes.len(), 9);
        assert!(boxes[3].bounds.y < boxes[4].bounds.y);
        assert!((boxes[3].bounds.x - boxes[4].bounds.x).abs() < 0.01);
        assert!(boxes[7].bounds.y < boxes[8].bounds.y);
        assert!(boxes[8].bounds.x < boxes[7].bounds.x);
        let expected = [
            (326.0, 168.0),
            (147.0, 269.0),
            (583.0, 269.0),
            (240.0, 370.0),
            (240.0, 471.0),
            (497.0, 370.0),
            (669.0, 370.0),
            (497.0, 471.0),
            (411.0, 572.0),
        ];
        for (element, (x, y)) in boxes.iter().zip(expected) {
            assert!(
                (element.bounds.x - x).abs() < 3.0 && (element.bounds.y - y).abs() < 3.0,
                "{:?}",
                element.bounds
            );
            let Visual::TextLayout { visual, .. } = &element.visual else {
                panic!("no fabricated shadow")
            };
            assert!(matches!(
                visual.as_ref(),
                Visual::RichText {
                    geometry: Geometry::Rectangle,
                    fill: Paint::Solid(0x4f81_bdff),
                    ..
                }
            ));
        }
        // XML point order must not override connection srcOrd; mirroring must
        // preserve the same topology and fit, including the deepest assistant.
        diagram.nodes.reverse();
        let reordered = render(&diagram);
        for (a, b) in boxes
            .iter()
            .zip(reordered.iter().filter(|e| e.text.is_some()).rev())
        {
            assert_eq!(a.bounds, b.bounds);
        }
        diagram.right_to_left = true;
        let mirrored = render(&diagram);
        for (a, b) in reordered
            .iter()
            .filter(|e| e.text.is_some())
            .zip(mirrored.iter().filter(|e| e.text.is_some()))
        {
            assert!(
                (a.bounds.x + b.bounds.x + a.bounds.width - (2.0 * bounds.x + bounds.width)).abs()
                    < 0.01
            );
            assert_eq!(a.bounds.y, b.bounds.y);
        }
        let id = diagram.nodes[8].model_id.clone();
        diagram.nodes[8].parent_id = Some(id);
        assert!(
            super::diagram_org_chart_elements(&diagram, bounds, 0).is_none(),
            "cycle must use diagnostic fallback"
        );
    }

    #[test]
    fn xl8galry_exploded_pie_labels_clear_faces_and_remain_compact() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/xl8galry.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let chart = parse_chart(
            &package,
            "xl/charts/chart11.xml",
            super::default_scheme_color,
        )
        .unwrap()
        .unwrap();
        let bounds = Rect {
            x: 0.0,
            y: 0.0,
            width: 800.0,
            height: 600.0,
        };
        let pie = chart.pie_bounds(bounds, bounds, true);
        let radius = pie.width * 0.52;
        let view = chart.view_3d.unwrap();
        let vertical_radius = radius * view.pie_vertical_ratio();
        let depth = radius * view.pie_depth_ratio();
        let (_, slices) = chart_pie_slices(
            &chart.series[0],
            (pie.x + pie.width / 2.0, pie.y + (pie.height - depth) / 2.0),
            (radius, vertical_radius),
            depth,
            view.pie_depth_perspective(),
            None,
        );
        let labels = chart_pie_label_layout(
            &chart,
            &chart.series[0],
            &slices,
            bounds,
            radius,
            vertical_radius,
        );
        assert_eq!(labels.len(), 12);
        // Excel places these single-line category names beside the slices, with
        // Central below the front face and no gratuitous lines through the pie.
        for label in &labels {
            assert!(
                label.bounds.height <= label.style.font_size * 1.5,
                "{} must not reserve multiple empty lines",
                label.text
            );
        }
        let central = labels.iter().find(|label| label.text == "Central").unwrap();
        let front = slices
            .iter()
            .find(|slice| slice.index == central.index)
            .unwrap()
            .side
            .as_ref()
            .unwrap()
            .0;
        assert!(
            central.bounds.y >= front.y + front.height,
            "Central overlaps the front face"
        );
        assert!(
            labels.iter().filter(|label| label.leader.is_some()).count() <= 4,
            "nearby category labels should not all gain leader lines"
        );
        for (index, left) in labels.iter().enumerate() {
            for right in &labels[index + 1..] {
                assert!(
                    left.bounds.x + left.bounds.width <= right.bounds.x
                        || right.bounds.x + right.bounds.width <= left.bounds.x
                        || left.bounds.y + left.bounds.height <= right.bounds.y
                        || right.bounds.y + right.bounds.height <= left.bounds.y,
                    "{} overlaps {}",
                    left.text,
                    right.text
                );
            }
        }
    }

    #[test]
    fn xl8galry_all_visible_legends_match_excel_order() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/xl8galry.xlsx"),
            Limits::default(),
        )
        .unwrap();
        // Excel reference: side legends on Outdoor Bars and B&W Area read top-to-bottom;
        // horizontal legends, including Tubes, read in the original series order.
        let visible = [
            (1, 10, true),
            (2, 4, false),
            (3, 2, false),
            (4, 2, false),
            (5, 2, false),
            (6, 2, false),
            (7, 2, false),
            (10, 10, false),
            (12, 10, false),
            (13, 10, false),
            (16, 3, false),
            (19, 10, true),
        ];
        for number in 1..=20 {
            let chart = parse_chart(
                &package,
                &format!("xl/charts/chart{number}.xml"),
                super::default_scheme_color,
            )
            .unwrap()
            .unwrap();
            if let Some((_, count, reversed)) = visible.iter().find(|(n, _, _)| *n == number) {
                assert!(chart.show_legend, "Sheet{number}");
                let mut expected = (0..*count).collect::<Vec<_>>();
                if *reversed {
                    expected.reverse();
                }
                assert_eq!(
                    chart
                        .legend_entries()
                        .iter()
                        .map(|e| e.0)
                        .collect::<Vec<_>>(),
                    expected,
                    "Sheet{number}"
                );
            } else {
                assert!(!chart.show_legend, "Sheet{number} has no legend in Excel");
            }
        }
    }

    #[test]
    fn xl8galry_inherits_series_borders_theme_shadow_and_axis_color() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/xl8galry.xlsx"),
            Limits::default(),
        )
        .unwrap();
        for n in [1, 3, 10, 12, 18, 19] {
            let chart = parse_chart(
                &package,
                &format!("xl/charts/chart{n}.xml"),
                super::default_scheme_color,
            )
            .unwrap()
            .unwrap();
            for series in &chart.series {
                assert!(
                    series.effects.outer_shadow.is_some(),
                    "chart{n}: missing theme shadow"
                );
            }
        }
        let chart = parse_chart(
            &package,
            "xl/charts/chart1.xml",
            super::default_scheme_color,
        )
        .unwrap()
        .unwrap();
        assert_eq!(chart.category_axis_options().label_color, Some(0x000000ff));
        for series in &chart.series {
            assert_eq!(series.point_border_colors, vec![Some(0x000000ff)]);
            assert!((series.point_stroke_width(0) - 12700.0 / 9525.0).abs() < 0.001);
            let shadow = series.effects.outer_shadow.unwrap().shadow;
            assert_eq!(shadow.color & 0xff, 163);
            assert!((shadow.blur - 50800.0 / 9525.0).abs() < 0.001);
        }
    }

    #[test]
    fn xl8galry_explicit_empty_effect_list_suppresses_theme_shadow() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/xl8galry.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let xml = String::from_utf8(
            package
                .required_part("xl/charts/chart1.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        let xml = xml.replacen("</c:spPr>", "<a:effectLst/></c:spPr>", 1);
        let parts = package
            .entry_names()
            .map(|name| {
                (
                    name.to_owned(),
                    if name == "xl/charts/chart1.xml" {
                        xml.as_bytes().to_vec()
                    } else {
                        package.required_part(name).unwrap().to_vec()
                    },
                )
            })
            .collect::<Vec<_>>();
        let entries = parts
            .iter()
            .map(|(name, bytes)| (name.as_str(), bytes.as_slice()))
            .collect::<Vec<_>>();
        let bytes = stored_zip(&entries);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(
            &package,
            "xl/charts/chart1.xml",
            super::default_scheme_color,
        )
        .unwrap()
        .unwrap();
        assert!(chart.series[0].effects.is_empty());
        assert!(chart.series[1].effects.outer_shadow.is_some());
    }

    #[test]
    fn xl8galry_column_faces_preserve_authored_depth_percent() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/xl8galry.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let chart = parse_chart(
            &package,
            "xl/charts/chart13.xml",
            super::default_scheme_color,
        )
        .unwrap()
        .unwrap();
        let view = chart.view_3d.unwrap();
        assert_eq!(view.depth_percent, Some(500));
        let bar = Rect {
            x: 0.0,
            y: 0.0,
            width: 15.0,
            height: 100.0,
        };
        let deep = super::chart_bar_3d_faces(bar, 0x9999ffff, view, false);
        let normal = super::chart_bar_3d_faces(
            bar,
            0x9999ffff,
            super::ChartView3D {
                depth_percent: Some(100),
                ..view
            },
            false,
        );
        assert!(
            deep[0].0.width > normal[0].0.width * 1.5,
            "500 percent depth must not render as 100 percent"
        );
    }

    #[test]
    fn xl8galry_general_labels_keep_cached_precision() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/xl8galry.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let chart = parse_chart(
            &package,
            "xl/charts/chart1.xml",
            super::default_scheme_color,
        )
        .unwrap()
        .unwrap();
        let label = super::chart_bar_data_label(
            &chart,
            &chart.series[6],
            0,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 30.0,
            },
            chart.series[6].values[0],
        )
        .unwrap();
        assert_eq!(label.1, "1189.679009");
        assert!(
            label.0.width > 60.0,
            "outside label box must fit the full cached value"
        );
        assert_eq!(
            label.0.x, 202.0,
            "outside labels begin after the bar, never over it"
        );
        let tubes = parse_chart(
            &package,
            "xl/charts/chart10.xml",
            super::default_scheme_color,
        )
        .unwrap()
        .unwrap();
        let label = super::chart_bar_data_label(
            &tubes,
            &tubes.series[0],
            0,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 40.0,
                height: 70.0,
            },
            tubes.series[0].values[0],
        )
        .unwrap();
        assert!(
            label.0.width <= 70.0,
            "rotated label fits the bar thickness"
        );
    }

    #[test]
    fn xl8galry_area_blocks_rotate_categories_behind_series() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/xl8galry.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let chart = parse_chart(
            &package,
            "xl/charts/chart9.xml",
            super::default_scheme_color,
        )
        .unwrap()
        .unwrap();
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 600.0,
            height: 400.0,
        };
        let (series, categories) = super::chart_area_3d_axis_labels(&chart, plot, 12.0);
        assert!(
            series.last().unwrap().1.x - series[0].1.x > plot.width * 0.4,
            "100 degree rotation puts series across the viewport"
        );
        assert!(
            categories.last().unwrap().1.x < categories[0].1.x,
            "categories recede to the right at rotY=100"
        );
    }

    #[test]
    fn xl8galry_cones_use_series_depth_and_area_blocks_follow_rotation() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/xl8galry.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let cones = parse_chart(
            &package,
            "xl/charts/chart8.xml",
            super::default_scheme_color,
        )
        .unwrap()
        .unwrap();
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 600.0,
            height: 400.0,
        };
        let bounds = |i| {
            super::chart_bar_segment_bounds(
                &cones,
                i,
                0,
                plot,
                (0.0, 4000.0),
                cones.bar_gap_width_percent,
            )
            .unwrap()
        };
        let (first, last) = (bounds(0), bounds(9));
        assert!(
            (first.x - last.x).abs() < 0.1,
            "rotY=0 puts series behind one another, not beside one another"
        );
        assert!((first.y + first.height) - (last.y + last.height) > 50.0);
        assert!(
            first.width > 20.0,
            "depth grouping must not divide column width by the series count"
        );
    }

    #[test]
    fn xl8galry_point_bubble_flag_does_not_flatten_other_chart_families() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/xl8galry.xlsx"),
            Limits::default(),
        )
        .unwrap();
        for i in [11, 17] {
            let chart = parse_chart(
                &package,
                &format!("xl/charts/chart{i}.xml"),
                super::default_scheme_color,
            )
            .unwrap()
            .unwrap();
            assert!(
                chart.series.iter().all(|s| s.three_d),
                "chart {i}: point bubble3D must not override the chart family"
            );
        }
    }

    #[test]
    fn xl8galry_bar_point_pattern_reaches_spreadsheet_objects() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/xl8galry.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let chart = parse_chart(
            &package,
            "xl/charts/chart17.xml",
            super::default_scheme_color,
        )
        .unwrap()
        .unwrap();
        assert!(matches!(
            chart.series[2].point_fills[0],
            Some(ChartFill::Pattern { .. })
        ));
        let mut objects = Vec::new();
        let mut diagnostics = Vec::new();
        super::push_spreadsheet_chart(
            chart,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 678.0,
                height: 461.0,
            },
            &crate::model::SourceRef {
                part: "xl/charts/chart17.xml".into(),
                mapping: crate::model::MappingQuality::Exact,
                locator: crate::model::SourceLocator::Flat {
                    kind: "chart",
                    index: None,
                    row: None,
                    column: None,
                    text_range: None,
                },
            },
            0,
            10000,
            &mut objects,
            &mut diagnostics,
        )
        .unwrap();
        assert!(objects.iter().any(|o| matches!(&o.visual, Visual::PaintedShape { fill: Paint::Pattern { preset, .. }, .. } if preset == "pct75")));
    }

    #[test]
    fn real_reversed_categories_share_area_and_group_label_order() {
        let package = Package::open(include_bytes!("../../../tests/fixtures/ooxml-reversed-categories.docx"), Limits::default()).unwrap();
        let mut chart = parse_chart(&package, "word/charts/chart1.xml", |_| None).unwrap().unwrap();
        let plot = Rect { x: 100.0, y: 100.0, width: 400.0, height: 300.0 };
        chart.series[0].category_levels = vec![chart.series[0].categories.clone(), vec!["First".into(), "Second".into()]];
        let labels = chart.supplemental_axis_labels(plot, 12.0);
        assert!(labels[0].0.y < labels[1].0.y);
        chart.series[0].kind = ChartKind::Area;
        chart.series[0].bar_horizontal = false;
        std::mem::swap(&mut chart.value_axis_options, &mut chart.horizontal_axis_options);
        let (_, points, _) = super::chart_area_geometry(&chart, 0, plot, (0.0, 1.0)).unwrap();
        assert!(points[0].0 > points.last().unwrap().0);
    }

    #[test]
    fn real_date_axis_uses_month_boundaries_and_weekly_minor_lines() {
        let package = Package::open(include_bytes!("../../../tests/fixtures/ooxml-calendar-axis.xlsx"), Limits::default()).unwrap();
        let chart = parse_chart(&package, "xl/charts/chart1.xml", |_| None).unwrap().unwrap();
        let plot = Rect { x: 100.0, y: 100.0, width: 400.0, height: 300.0 };
        let marks = chart.date_axis_marks(&chart.series[0], plot).unwrap();
        let months = marks.iter().filter(|(_, _, major)| *major).collect::<Vec<_>>();
        assert_eq!(months.iter().map(|(_, text, _)| text.as_str()).collect::<Vec<_>>(),
            ["1/1/2019", "2/1/2019", "3/1/2019", "4/1/2019", "5/1/2019"]);
        assert!((months[1].0 - (100.0 + 400.0 * 31.0 / 120.0)).abs() < 0.1);
        assert!((months[2].0 - (100.0 + 400.0 * 59.0 / 120.0)).abs() < 0.1);
        assert!(marks.iter().filter(|(_, _, major)| !*major).count() >= 17);
    }

    #[test]
    fn real_ooxml_chart_display_units() {
        let package = Package::open(include_bytes!("../../../tests/fixtures/ooxml-display-units.docx"), Limits::default()).unwrap();
        let chart = parse_chart(&package, "word/charts/chart1.xml", |_| None).unwrap().unwrap();
        assert_eq!(chart.value_axis_label(2_000_000_000.0), "2");
        assert_eq!(chart.value_axis_label(0.5), "5E-10");
        assert!(chart.supplemental_axis_labels(Rect { x: 100.0, y: 100.0, width: 400.0, height: 300.0 }, 12.0)
            .iter().any(|(_, text)| text == "Billions"));
    }

    #[test]
    fn real_ooxml_multilevel_categories_preserve_leaf_labels() {
        let package = Package::open(include_bytes!("../../../tests/fixtures/ooxml-multilevel-axis.docx"), Limits::default()).unwrap();
        let chart = parse_chart(&package, "word/charts/chart1.xml", |_| None).unwrap().unwrap();
        assert_eq!(chart.series[0].categories, ["Categoria 1", "Categoria 2", "Categoria 3", "Categoria 4"]);
        assert_eq!(chart.series[0].category_levels.len(), 2);
        assert_eq!(chart.category_x(&chart.series[0], 0, Rect { x: 100.0, y: 100.0, width: 400.0, height: 300.0 }, false), Some(150.0));
        let labels = chart.supplemental_axis_labels(Rect { x: 100.0, y: 100.0, width: 400.0, height: 300.0 }, 12.0);
        assert_eq!(labels.iter().map(|(_, text)| text.as_str()).collect::<Vec<_>>(), ["2011", "2012"]);
        assert!(labels[0].0.x < labels[1].0.x);
    }

    #[test]
    fn supplied_tdf105517_chart_styles_and_number_formats() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/corpus-tdf105517.pptx"),
            Limits::default(),
        )
        .unwrap();
        let chart = parse_chart(&package, "ppt/charts/chart1.xml", |_| Some(0xff0000ff))
            .unwrap()
            .unwrap();
        assert_eq!(
            chart.series.iter().map(|s| s.color).collect::<Vec<_>>(),
            vec![Some(0xf0ab00ff), Some(0x226ca9ff)]
        );
        assert_eq!(
            chart.series[0].marker_symbol, None,
            "point 11 must not override the explicit series none"
        );
        assert_eq!(
            chart.series[1].marker_symbol.as_deref(),
            Some("triangle"),
            "automatic marker uses source series idx 2"
        );
        assert_eq!(
            chart.series[1].marker_size,
            Some(16.0),
            "automatic triangles must extend beyond the 8px stroke"
        );
        assert_eq!(chart.value_axis_label(1_200_000.0), "1,200,000");
        assert_eq!(
            chart.value_axis_grid_style(0xd9d9d9ff, 0.75),
            (0x000000ff, 15875.0 / 9525.0)
        );
        let (bounds, text, style, _) = super::chart_bar_data_label(
            &chart,
            &chart.series[1],
            2,
            Rect {
                x: 400.0,
                y: 300.0,
                width: 0.0,
                height: 0.0,
            },
            220_000.0,
        )
        .unwrap();
        assert_eq!(text, "220,000");
        assert_eq!((style.font_size, style.bold), (32.0, true));
        assert!(bounds.width > 110.0 && bounds.y + bounds.height < 300.0);
        assert!(super::chart_bar_data_label(&chart, &chart.series[1], 1, bounds, 0.0).is_none());
        for (value, format, expected) in [
            (1200000.0, "#,##0", "1,200,000"),
            (-1234.5, "#,##0.00", "-1,234.50"),
            (0.0, "#,##0", "0"),
            (0.25, "0.0%", "25.0%"),
            (-1234.0, "$#,##0;($#,##0)", "($1,234)"),
        ] {
            assert_eq!(super::format_chart_value(value, Some(format)), expected);
        }
        // Chart-local mappings and partial palettes must still inherit unmapped host slots.
        let colors = super::drawingml_theme_colors(br#"<a:themeOverride xmlns:a="a"><a:clrScheme><a:accent1><a:srgbClr val="123456"/></a:accent1></a:clrScheme></a:themeOverride>"#, Limits::default(), "theme.xml").unwrap();
        assert_eq!(colors.len(), 1);
        assert_eq!(colors["accent1"], 0x123456ff);
    }

    #[test]
    fn next_twenty_chart_fields_and_automatic_marker() {
        for (name, part, expected) in [
            ("tdf125444.pptx", "ppt/charts/chart1.xml", "0-1 Vendor 38%"),
            (
                "CustomDataLabel_tdf115107.pptx",
                "ppt/charts/chart1.xml",
                "90.0 = 90",
            ),
        ] {
            let bytes = std::fs::read(format!("tests/fixtures/pptx-next20/{name}")).unwrap();
            let package = Package::open(&bytes, Limits::default()).unwrap();
            let chart = parse_chart(&package, part, |_| None).unwrap().unwrap();
            assert_eq!(chart.series[0].data_labels[0].text, expected, "{name}");
        }
        let bytes = std::fs::read("tests/fixtures/pptx-next20/tdf137691_dataTable.pptx").unwrap();
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, "ppt/charts/chart1.xml", |_| None)
            .unwrap()
            .unwrap();
        assert_eq!(chart.series[0].marker_symbol.as_deref(), Some("diamond"));
    }

    #[test]
    fn supplied_chart_wall_zero_width_gridlines_are_hidden() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/corpus-chart-wall.pptx"),
            Limits::default(),
        )
        .unwrap();
        let chart = parse_chart(&package, "ppt/charts/chart1.xml", |_| None)
            .unwrap()
            .unwrap();
        assert!(
            !chart.value_axis_grid_lines_visible(),
            "Office omits the authored zero-width gridlines"
        );
    }

    #[test]
    fn supplied_3d_columns_do_not_add_an_extra_axis_interval() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/corpus-bnc864396.pptx"),
            Limits::default(),
        )
        .unwrap();
        let chart = parse_chart(&package, "ppt/charts/chart1.xml", |_| None)
            .unwrap()
            .unwrap();
        assert_eq!(chart.value_axis(), (0.0, 9.0, 1.0), "cached Office axis");
    }

    #[test]
    fn supplied_horizontal_bar_uses_value_axis_despite_stale_axis_positions() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/corpus-tdf48041.pptx"),
            Limits::default(),
        )
        .unwrap();
        let chart = parse_chart(&package, "ppt/charts/chart3.xml", |_| None)
            .unwrap()
            .unwrap();
        assert_eq!(
            chart.numerical_axis_options().id.as_deref(),
            Some("61787032")
        );
        let (_, maximum, _) = chart.value_axis();
        assert!(
            (4.0..=6.0).contains(&maximum),
            "Office horizontal scale is 0–5, got {maximum}"
        );
    }

    #[test]
    fn supplied_mixed_percent_area_keeps_absolute_column_axis() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/corpus-stacked-mix.pptx"),
            Limits::default(),
        )
        .unwrap();
        for (part, percent) in [
            ("ppt/charts/chart1.xml", false),
            ("ppt/charts/chart2.xml", true),
            ("ppt/charts/chart3.xml", false),
        ] {
            let chart = parse_chart(&package, part, |name| match name {
                "accent1" => Some(0x4f81bdff),
                "accent2" => Some(0xc0504dff),
                "accent3" => Some(0x9bbb59ff),
                _ => None,
            })
            .unwrap()
            .unwrap();
            assert_eq!(chart.value_axis_is_percent(), percent, "{part}");
            if part.ends_with("chart3.xml") {
                assert_eq!(
                    chart.series[0].color,
                    Some(0x9bbb59ff),
                    "Series 3 keeps accent3 despite being parsed first"
                );
                let (_, maximum, _) = chart.value_axis();
                assert!(
                    (7.3..=10.0).contains(&maximum),
                    "Office has an absolute axis up to 8, got {maximum}"
                );
            }
        }
    }

    #[test]
    fn surface_source_data_supports_line_and_wireframe_boundaries() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/tdf128207.docx"),
            Limits::default(),
        )
        .unwrap();
        let xml = String::from_utf8(
            package
                .required_part("word/charts/chart2.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        for (xml, kind) in [
            (
                xml.replace("surface3DChart", "line3DChart"),
                ChartKind::Line3D,
            ),
            (
                xml.replace("surface3DChart", "surfaceChart"),
                ChartKind::Surface,
            ),
            (
                xml.replace(
                    "<c:surface3DChart>",
                    "<c:surface3DChart><c:wireframe val=\"1\"/>",
                ),
                ChartKind::SurfaceWireframe,
            ),
        ] {
            let bytes = stored_zip(&[("chart.xml", xml.as_bytes())]);
            let package = Package::open(&bytes, Limits::default()).unwrap();
            let chart = parse_chart(&package, "chart.xml", super::default_scheme_color)
                .unwrap()
                .unwrap();
            assert_eq!(chart.series[0].kind, kind);
            assert!(
                chart
                    .series
                    .iter()
                    .all(|s| s.three_d == (kind != ChartKind::Surface))
            );
            assert!(super::chart_extended_elements(&chart, Rect::default(), 10).is_err());
            let elements = super::chart_extended_elements(
                &chart,
                Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 600.0,
                    height: 400.0,
                },
                Limits::default().max_document_objects,
            )
            .unwrap()
            .unwrap();
            if kind == ChartKind::Surface {
                assert!(elements.iter().any(|e| e.text.as_deref() == Some("Row 1")));
            }
            assert!(elements.iter().any(|e| matches!(
                &e.visual,
                Visual::PaintedShape {
                    geometry: Geometry::Path { .. },
                    fill: Paint::None,
                    stroke: Paint::Solid(_),
                    ..
                }
            )));
        }
    }

    #[test]
    fn supplied_smartart_color_lists_cover_span_cycle_and_repeat() {
        let original = Package::open(
            include_bytes!("../../../tests/fixtures/fill-color-list.pptx"),
            Limits::default(),
        )
        .unwrap();
        let xml = String::from_utf8(
            original
                .required_part("ppt/diagrams/colors1.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        for method in ["span", "cycle", "repeat"] {
            let xml = xml.replace("meth=\"repeat\"", &format!("meth=\"{method}\""));
            let bytes = stored_zip(&[("colors.xml", xml.as_bytes())]);
            let package = Package::open(&bytes, Limits::default()).unwrap();
            let lists = super::parse_diagram_color_list(
                &package,
                "colors.xml",
                &super::default_scheme_color,
                "fillClrLst",
            )
            .unwrap();
            let list = lists.values().find(|list| list.colors.len() >= 3).unwrap();
            assert_eq!(list.sample(0, 5), list.colors.first().copied());
            match method {
                "span" => assert_eq!(list.sample(4, 5), list.colors.last().copied()),
                "cycle" => {
                    assert_eq!(list.sample(4, 5), list.colors.first().copied());
                    assert_eq!(list.sample(2, 5), list.colors.last().copied());
                }
                _ => assert_eq!(
                    list.sample(list.colors.len(), 9),
                    list.colors.first().copied()
                ),
            }
        }
    }

    #[test]
    fn supplied_smartart_without_drawings_preserves_shape_families_and_text() {
        for (bytes, count) in [
            (
                include_bytes!("../../../tests/fixtures/smartart-chevron.pptx").as_slice(),
                3,
            ),
            (
                include_bytes!("../../../tests/fixtures/smartart-cycle.pptx").as_slice(),
                5,
            ),
            (
                include_bytes!("../../../tests/fixtures/smartart-dir.pptx").as_slice(),
                2,
            ),
        ] {
            let package = Package::open(bytes, Limits::default()).unwrap();
            let diagram = super::parse_diagram(
                &package,
                "ppt/diagrams/data1.xml",
                &Default::default(),
                Some("ppt/diagrams/colors1.xml"),
                None,
                super::default_scheme_color,
            )
            .unwrap();
            let bounds = Rect {
                x: 160.0,
                y: 146.67,
                width: 640.0,
                height: 426.67,
            };
            let elements = diagram_horizontal_list_elements(&diagram, bounds, 0x4f81_bdff)
                .or_else(|| super::diagram_semantic_elements(&diagram, bounds, 0x4f81_bdff))
                .unwrap();
            let labels = elements
                .iter()
                .filter_map(|e| e.text.as_deref())
                .collect::<Vec<_>>();
            assert!(labels.len() >= count);
            for node in &diagram.nodes {
                assert!(
                    labels.iter().any(|text| text.contains(&node.text)),
                    "missing {}",
                    node.text
                );
            }
            if count == 5 {
                assert_eq!(elements.iter().filter(|e| e.text.is_some()).count(), 5);
                let labels = elements
                    .iter()
                    .filter(|e| e.text.is_some())
                    .collect::<Vec<_>>();
                assert!(labels[0].bounds.y < labels[1].bounds.y);
                assert!(labels[1].bounds.x > labels[0].bounds.x);
                assert!(labels[4].bounds.x < labels[0].bounds.x);
            }
            if diagram.right_to_left {
                let first = elements
                    .iter()
                    .find(|e| e.text.as_deref() == Some("first"))
                    .unwrap();
                let second = elements
                    .iter()
                    .find(|e| e.text.as_deref() == Some("second"))
                    .unwrap();
                assert!(second.bounds.x < first.bounds.x);
            }
        }
    }

    #[test]
    fn supplied_missing_chart_families_are_parsed() {
        for (bytes, part) in [
            (
                include_bytes!("../../../tests/fixtures/funnel-pp1.pptx").as_slice(),
                "ppt/charts/chartEx1.xml",
            ),
            (
                include_bytes!("../../../tests/fixtures/testStockChart.docx").as_slice(),
                "word/charts/chart1.xml",
            ),
            (
                include_bytes!("../../../tests/fixtures/tdf128207.docx").as_slice(),
                "word/charts/chart2.xml",
            ),
        ] {
            let package = Package::open(bytes, Limits::default()).unwrap();
            let chart = if part.contains("chartEx") {
                super::parse_chart_ex(&package, part, super::default_scheme_color)
            } else {
                super::parse_chart(&package, part, super::default_scheme_color)
            }
            .unwrap();
            let chart = chart.expect("real chart must not be omitted");
            let expected = if part.contains("chartEx") {
                ChartKind::Funnel
            } else if part.ends_with("chart2.xml") {
                ChartKind::Surface
            } else {
                ChartKind::Stock
            };
            assert!(
                chart.series.iter().any(|s| s.kind == expected),
                "real family misclassified: {part}"
            );
            if expected == ChartKind::Surface {
                assert!(chart.surface_band_fills.iter().flatten().count() >= 3);
                assert_eq!(chart.view_3d.unwrap().rot_y, 170);
                assert_eq!(
                    chart
                        .legend_entries()
                        .iter()
                        .map(|(_, label, _)| label.as_str())
                        .collect::<Vec<_>>(),
                    ["0–2", "2–4", "4–6"]
                );
            }
            if expected == ChartKind::Stock {
                assert_eq!(
                    chart
                        .series
                        .iter()
                        .filter(|s| s.kind == ChartKind::Stock)
                        .count(),
                    4
                );
                assert!(chart.series.iter().any(|s| s.kind == ChartKind::Bar));
                assert!(chart.secondary_value_axis_options.is_some());
                let stock = chart
                    .series
                    .iter()
                    .find(|s| s.kind == ChartKind::Stock)
                    .unwrap();
                assert_eq!(stock.axis_id.as_deref(), Some("80166912"));
                assert_eq!(chart.value_axis_for_series(stock).1, 70.0);
                assert!(chart.secondary_value_axis_ticks().is_some());
            }
            let elements = super::chart_extended_elements(
                &chart,
                Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 600.0,
                    height: 400.0,
                },
                Limits::default().max_document_objects,
            )
            .unwrap()
            .unwrap();
            assert!(
                elements
                    .iter()
                    .any(|e| matches!(e.visual, Visual::PaintedShape { .. }))
            );
        }
    }

    #[test]
    fn chart_color_mapping_overrides_host_and_preserves_unmapped_colors() {
        for mapping in ["", "<c:clrMapOvr bg1=\"lt1\"/>"] {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:pieChart><c:ser>
                <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></c:spPr></c:dPt>
                <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr></c:dPt>
                <c:val><c:numLit><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>95</c:v></c:pt></c:numLit></c:val>
                </c:ser></c:pieChart></c:plotArea></c:chart>{mapping}</c:chartSpace>"#
            );
            let bytes = stored_zip(&[("chart.xml", xml.as_bytes())]);
            let package = Package::open(&bytes, Limits::default()).unwrap();
            let chart = parse_chart(&package, "chart.xml", |name| match name {
                "bg1" => Some(0x0000_00ff),
                "lt1" => Some(0xffff_ffff),
                "accent1" => Some(0x4472_c4ff),
                _ => None,
            })
            .unwrap()
            .unwrap();
            assert_eq!(
                chart.series[0].point_colors,
                [
                    if mapping.is_empty() {
                        0x0000_00ff
                    } else {
                        0xffff_ffff
                    },
                    0x4472_c4ff
                ]
            );
        }
    }

    #[test]
    fn supplied_empty_chart_title_keeps_office_default_title() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/chart-empty-title-original.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let chart = parse_chart(&package, "xl/charts/chart1.xml", |_| None)
            .unwrap()
            .unwrap();
        assert!(chart.title.is_empty(), "preserve the authored empty text");
        assert!(chart.show_title, "the existing title must remain visible");
        assert_eq!(chart.title_text(), "Chart Title");
        assert_eq!(chart.title_font_size, Some(14.0 * 96.0 / 72.0));
    }

    #[test]
    fn empty_chart_title_uses_the_only_named_series() {
        const PART: &str = "word/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:title/><c:plotArea><c:pieChart><c:ser><c:tx><c:v>p_n</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(chart.title_text(), "p_n");
    }

    #[test]
    fn chart_title_visibility_distinguishes_empty_absent_and_deleted() {
        const PART: &str = "xl/charts/chart1.xml";
        for (title, visible, text) in [
            ("", false, "Chart Title"),
            ("<c:autoTitleDeleted val=\"0\"/>", false, "Chart Title"),
            ("<c:title/>", true, "Chart Title"),
            (
                "<c:title/><c:autoTitleDeleted val=\"false\"/>",
                true,
                "Chart Title",
            ),
            (
                "<c:title/><c:autoTitleDeleted val=\"1\"/>",
                false,
                "Chart Title",
            ),
            (
                "<c:autoTitleDeleted val=\"true\"/><c:title/>",
                false,
                "Chart Title",
            ),
            ("<c:title/><c:autoTitleDeleted/>", false, "Chart Title"),
            (
                "<c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>",
                true,
                "Sales",
            ),
        ] {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart>{title}<c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
            );
            let bytes = stored_zip(&[(PART, xml.as_bytes())]);
            let package = Package::open(&bytes, Limits::default()).unwrap();
            let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
            assert_eq!(chart.show_title, visible, "{title}");
            assert_eq!(chart.title_text(), text, "{title}");
        }
    }

    #[test]
    fn supplied_chart_titles_and_axis_style_inheritance_are_shared() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/chart-original.docx"),
            Limits::default(),
        )
        .unwrap();
        let mut chart = parse_chart(&package, "word/charts/chart1.xml", |_| None)
            .unwrap()
            .unwrap();
        assert_eq!(chart.title, "折线统计图");
        assert_eq!(chart.title_text(), "折线统计图");
        chart.title.clear();
        assert_eq!(chart.title_text(), "Chart Title");
        assert_eq!(chart.value_axis_options.title, "人数");
        assert_eq!(chart.horizontal_axis_options.title, "日期");
        assert_eq!(
            chart
                .axis_title_text_style(&chart.value_axis_options)
                .font_size,
            10.0 * 96.0 / 72.0
        );
        chart.font_size = Some(16.0);
        chart.font_bold = true;
        let inherited = chart.axis_title_text_style(&chart.value_axis_options);
        assert_eq!(inherited.font_size, 16.0);
        assert!(inherited.bold);
        chart.value_axis_options.title_font_size = Some(20.0);
        chart.value_axis_options.title_bold = Some(false);
        let explicit = chart.axis_title_text_style(&chart.value_axis_options);
        assert_eq!(explicit.font_size, 20.0);
        assert!(!explicit.bold);
        assert_eq!(explicit.color, 0x0000_00ff);
    }

    #[test]
    fn fill_style_zero_and_one_thousand_mean_no_paint() {
        for idx in ["0", "1000"] {
            assert!(
                !drawingml_fill_reference_has_paint(
                    &[XmlAttribute {
                        name: "idx",
                        value: idx
                    }],
                    "test.xml",
                )
                .unwrap()
            );
        }
        assert!(
            drawingml_fill_reference_has_paint(
                &[XmlAttribute {
                    name: "idx",
                    value: "1"
                }],
                "test.xml",
            )
            .unwrap()
        );
    }

    #[test]
    fn picture_outer_shadow_uses_clipping_safe_composition() {
        let visual = DrawingMlPictureEffects {
            outer_shadow: Some(OuterShadow {
                shadow: Shadow {
                    color: 0x0000_0059,
                    blur: 4.0,
                    offset_x: 2.0,
                    offset_y: 2.0,
                },
                scale_x: 1.0,
                scale_y: 1.0,
                skew_x: 0.0,
                skew_y: 0.0,
                alignment: 7,
            }),
            ..DrawingMlPictureEffects::default()
        }
        .wrap(Visual::None, None);

        assert!(matches!(
            visual,
            Visual::AdvancedEffect {
                outer_shadow: Some(_),
                ..
            }
        ));
    }

    fn iso_strict_logarithmic_chart() -> Chart {
        const PART: &str = "xl/charts/chart2.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/><c:ser><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>5.33776433546107</c:v></c:pt><c:pt idx="2"><c:v>20.193474219937901</c:v></c:pt><c:pt idx="3"><c:v>30.795380276430301</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart><c:catAx><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:tickLblPos val="nextTo"/></c:catAx><c:valAx><c:scaling><c:logBase val="10"/></c:scaling></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="edge"/><c:hMode val="edge"/><c:x val="0.9166667183334366"/><c:y val="0.45537076343717908"/><c:w val="0.99671058342116681"/><c:h val="0.52042382745635063"/></c:manualLayout></c:layout><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr></c:legend></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        parse_chart(&package, PART, |_| None).unwrap().unwrap()
    }

    #[test]
    fn chart_title_explicit_color_preserves_default_typography() {
        let mut chart = iso_strict_logarithmic_chart();
        let original = chart.title_text_style();
        chart.title_text_color = Some(0x008080ff);
        let explicit = chart.title_text_style();
        assert_eq!(explicit.color, 0x008080ff);
        assert_eq!(explicit.font_size, original.font_size);
        assert_eq!(explicit.bold, original.bold);
        chart.title_text_color = None;
        assert_eq!(chart.title_text_style().color, original.color);
    }

    #[test]
    fn category_axis_without_cached_labels_uses_office_point_numbers() {
        assert_eq!(
            iso_strict_logarithmic_chart().series[0].categories,
            ["1", "2", "3", "4"]
        );
    }

    #[test]
    fn legend_preserves_its_authored_outline() {
        let chart = iso_strict_logarithmic_chart();
        assert!(matches!(
            chart.legend_stroke,
            Some(ChartFill::Solid(0x0000_00ff))
        ));
        assert!((chart.legend_stroke_width - 3175.0 / super::EMU_PER_CSS_PIXEL).abs() < 0.001);
        let bounds = chart.legend_bounds.unwrap();
        assert!((bounds.x - 0.916_666_7).abs() < 0.001);
        assert!((bounds.y - 0.455_370_75).abs() < 0.001);
        assert!((bounds.width - 0.080_043_87).abs() < 0.001);
        assert!((bounds.height - 0.065_053_06).abs() < 0.001);
    }

    #[test]
    fn shared_color_transform_preserves_alpha_and_applies_transparency() {
        let mut color = 0x3366_99ff;
        apply_color_transform(&mut color, "alpha", 0.5);
        assert_eq!(color, 0x3366_9980);
    }

    #[test]
    fn drawingml_bevel_uses_schema_dimensions_when_omitted() {
        let bevel = parse_three_d_bevel(&[], "ppt/slides/slide2.xml").unwrap();

        assert!((bevel.width - 8.0).abs() < 0.001);
        assert!((bevel.height - 8.0).abs() < 0.001);
        assert_eq!(bevel.preset, "circle");
    }

    #[test]
    fn supplied_target_list_uses_authored_order_and_nested_details() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/complex2005_12rtm.docx"),
            Limits::default(),
        )
        .unwrap();
        let mut diagram = super::parse_diagram(
            &package,
            "word/diagrams/data1.xml",
            &Default::default(),
            None,
            None,
            |_| None,
        )
        .unwrap();
        let bounds = Rect {
            x: 20.0,
            y: 30.0,
            width: 600.0,
            height: 360.0,
        };
        let elements = super::diagram_target_list_elements(&diagram, bounds, 0x4f81_bdff).unwrap();
        let labels = elements
            .iter()
            .filter_map(|element| element.text.as_deref())
            .collect::<Vec<_>>();
        assert_eq!(labels[0], "SMART ART!");
        assert_eq!(labels[1], "Typical Doc");
        assert_eq!(labels[3], "Complex Doc");
        assert!(labels[2].starts_with("• 5 pages\n• 2 pictures"));
        assert!(labels[2].contains("• Total e20s\n  • With e20 text boxes: 4"));
        assert!(labels[4].contains("• 150 pages\n• 5 text boxes"));
        assert_eq!(elements.len(), 11);
        for (index, size) in [360.0, 234.0, 108.0].into_iter().enumerate() {
            assert_eq!(elements[index].bounds.width, size);
            assert_eq!(elements[index].bounds.y, 30.0 + index as f32 * 108.0);
            assert_eq!(elements[index].bounds.x + size / 2.0, 200.0);
        }
        // The adapter must leave unknown layouts and disconnected cycles to the
        // existing best-effort fallback, rather than render a misleading subset.
        diagram.layout_type =
            Some("urn:microsoft.com/office/officeart/2005/8/layout/target30".to_owned());
        assert!(super::diagram_target_list_elements(&diagram, bounds, 0).is_none());
        diagram.layout_type =
            Some("urn:microsoft.com/office/officeart/2005/8/layout/target3#1".to_owned());
        diagram.nodes[0].parent_id = Some(diagram.nodes[0].model_id.clone());
        assert!(super::diagram_target_list_elements(&diagram, bounds, 0).is_none());
    }

    #[test]
    fn supplied_fill_color_list_keeps_columns_and_indexed_colors() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/fill-color-list.pptx"),
            Limits::default(),
        )
        .unwrap();
        let mut diagram = super::parse_diagram(
            &package,
            "ppt/diagrams/data1.xml",
            &Default::default(),
            Some("ppt/diagrams/colors1.xml"),
            None,
            |name| match name {
                "accent2" => Some(0xc050_4dff),
                "accent3" => Some(0x9bbb_59ff),
                "accent4" => Some(0x8064_a2ff),
                _ => None,
            },
        )
        .unwrap();
        let bounds = Rect {
            x: 133.0,
            y: 66.0,
            width: 694.0,
            height: 462.0,
        };
        let elements = diagram_horizontal_list_elements(&diagram, bounds, 0x4f81_bdff)
            .expect("hList1 must not fall back to a vertical hierarchy");
        let headers = elements
            .iter()
            .filter(|element| element.text.is_some())
            .collect::<Vec<_>>();
        assert_eq!(headers.len(), 3);
        for (i, (text, color)) in [("A", 0xc050_4dff), ("B", 0x9bbb_59ff), ("C", 0x8064_a2ff)]
            .into_iter()
            .enumerate()
        {
            assert_eq!(headers[i].text.as_deref(), Some(text));
            assert!((headers[i].bounds.height / headers[i].bounds.width - 0.4).abs() < 0.001);
            assert!(headers[i].bounds.y > bounds.y);
            let mut visual = &headers[i].visual;
            if let Visual::Effect { visual: inner, .. } = visual {
                visual = inner;
            }
            let Visual::TextLayout { visual, .. } = visual else {
                panic!("text layout");
            };
            assert!(
                matches!(visual.as_ref(), Visual::RichText { fill: Paint::Solid(actual), .. } if *actual == color)
            );
            if i > 0 {
                assert!(
                    headers[i].bounds.x > headers[i - 1].bounds.x + headers[i - 1].bounds.width
                );
                assert_eq!(headers[i].bounds.y, headers[0].bounds.y);
            }
        }
        assert_eq!(elements.len(), 6, "three headers and three tinted bodies");
        for (i, body) in elements.iter().step_by(2).enumerate() {
            let Visual::Effect {
                shadow: Some(_),
                visual,
                ..
            } = &body.visual
            else {
                panic!("authored body shadow");
            };
            let Visual::PaintedShape {
                fill: Paint::Solid(color),
                ..
            } = visual.as_ref()
            else {
                panic!("tinted body");
            };
            assert_eq!(*color & 0xff, 230, "90 percent alpha");
            assert!(body.bounds.y > headers[i].bounds.y);
            assert_eq!(
                diagram.nodes[i].font_size,
                Some(24.0),
                "authored 18 pt text"
            );
        }
        diagram.nodes[0].fill = Some(ChartFill::Solid(0x1234_56ff));
        let overridden = diagram_horizontal_list_elements(&diagram, bounds, 0).unwrap();
        let Visual::Effect { visual, .. } = &overridden[1].visual else {
            panic!("shadow");
        };
        let Visual::TextLayout { visual, .. } = visual.as_ref() else {
            panic!("text");
        };
        assert!(matches!(
            visual.as_ref(),
            Visual::RichText {
                fill: Paint::Solid(0x1234_56ff),
                ..
            }
        ));
        diagram.nodes[1].parent_id = Some(diagram.nodes[0].model_id.clone());
        let nested = diagram_horizontal_list_elements(&diagram, bounds, 0).unwrap();
        assert!(
            nested.iter().any(|element| element
                .text
                .as_deref()
                .is_some_and(|text| text.contains(&diagram.nodes[1].text))),
            "nested content must not be dropped"
        );
        diagram.nodes[1].parent_id = None;
        diagram.layout_type =
            Some("urn:microsoft.com/office/officeart/2005/8/layout/hList10".to_owned());
        assert!(diagram_horizontal_list_elements(&diagram, bounds, 0).is_none());
    }

    #[test]
    fn shared_horizontal_list_materializes_all_smartart_roles() {
        let nodes = (0..3)
            .map(|index| DiagramNode {
                model_id: index.to_string(),
                parent_id: None,
                sibling_order: 0,
                assistant: false,
                hierarchy_branch: None,
                parent_transition_id: None,
                bold: false,
                font_size: None,
                text: String::new(),
                placeholder: true,
                preset_geometry: None,
                custom_geometry: false,
                fill: Some(ChartFill::Solid(0xff00_00ff)),
                no_fill: false,
                shadow: None,
                three_d: None,
                fill_scheme: None,
                fill_transforms: Vec::new(),
            })
            .collect();
        let diagram = Diagram {
            data_part: "diagram.xml".to_owned(),
            drawing_parts: Vec::new(),
            drawing_text_colors: Default::default(),
            right_to_left: false,
            role_fills: Default::default(),
            diagnostics: Vec::new(),
            role_line_colors: Default::default(),
            layout_type: Some(
                "urn:microsoft.com/office/officeart/2005/8/layout/hList7#1".to_owned(),
            ),
            scene_three_d: false,
            background_fill: None,
            background_shadow: None,
            nodes,
        };

        let elements = diagram_horizontal_list_elements(
            &diagram,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 900.0,
                height: 600.0,
            },
            0x4f81_bdff,
        )
        .expect("hList7 uses the shared semantic layout");

        assert_eq!(elements.len(), 10);
        assert_eq!(
            elements
                .iter()
                .filter(|element| element.text.is_some())
                .count(),
            3
        );
        assert_eq!(
            elements
                .iter()
                .filter(|element| matches!(
                    element.visual,
                    crate::model::Visual::PaintedShape {
                        geometry: Geometry::Ellipse,
                        ..
                    }
                ))
                .count(),
            3,
        );
        assert!(elements.iter().any(|element| {
            matches!(
                element.visual,
                crate::model::Visual::PaintedShape {
                    geometry: Geometry::Path { .. },
                    ..
                }
            )
        }));
    }

    #[test]
    fn waterfall_spans_accumulate_changes_and_reset_at_subtotals() {
        let series = ChartSeries {
            kind: ChartKind::Waterfall,
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
            name: String::new(),
            category_levels: Vec::new(),
            categories: Vec::new(),
            x_values: Vec::new(),
            values: vec![100.0, 20.0, -50.0, 70.0],
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
            smooth: false,
            marker_symbol: None,
            marker_size: None,
            subtotals: vec![0, 3],
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
        };

        assert_eq!(
            series.value_spans(),
            vec![(0.0, 100.0), (100.0, 120.0), (120.0, 70.0), (0.0, 70.0)]
        );
    }

    #[test]
    fn within_linear_waterfall_uses_office_dark_to_light_role_colors() {
        assert_eq!(
            within_linear_waterfall_colors(0x5b9b_d5ff),
            [0x4679_a7ff, 0x5591_c7ff, 0x84ae_dcff]
        );
    }

    #[test]
    fn radar_geometry_uses_the_shared_chart_axis_maximum() {
        let geometry = radar_geometry(
            &[10.0, 20.0, 30.0],
            40.0,
            Rect {
                x: 200.0,
                y: 300.0,
                width: 100.0,
                height: 100.0,
            },
        )
        .unwrap();
        let Geometry::Path { commands, .. } = geometry else {
            panic!("radar geometry must be a path");
        };

        assert!(matches!(
            commands.first(),
            Some(PathCommand::MoveTo { x, y })
                if (*x - 50.0).abs() < 0.001 && (*y - 39.5).abs() < 0.001
        ));
        assert!(matches!(commands.last(), Some(PathCommand::ClosePath)));
    }

    #[test]
    fn chart_date_categories_use_their_cached_number_format() {
        assert_eq!(
            format_chart_category("37261", Some("m/d/yyyy"), false).as_deref(),
            Some("1/5/2002")
        );
        assert_eq!(
            format_chart_category("37257", Some("m/d/yyyy"), true).as_deref(),
            Some("1/1/2006")
        );
        assert_eq!(
            format_chart_category("44362", Some(r"d\-mmm\-yy"), false).as_deref(),
            Some("15-Jun-21")
        );
        assert_eq!(format_chart_category("37261", Some("General"), false), None);
    }

    #[test]
    fn date_axis_uses_numeric_bounds_instead_of_all_category_slots() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:lineChart><c:ser>
              <c:cat><c:numLit><c:formatCode>m/d/yyyy</c:formatCode><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:numLit></c:cat>
              <c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt><c:pt idx="3"><c:v>40</c:v></c:pt><c:pt idx="4"><c:v>50</c:v></c:pt></c:numLit></c:val>
            </c:ser></c:lineChart><c:dateAx><c:scaling><c:min val="2"/><c:max val="4"/></c:scaling><c:axPos val="b"/><c:majorTickMark val="out"/></c:dateAx></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let series = &chart.series[0];
        let plot = Rect {
            x: 10.0,
            y: 0.0,
            width: 100.0,
            height: 50.0,
        };

        assert!(chart.horizontal_axis_options.is_date);
        assert_eq!(series.x_values, [1.0, 2.0, 3.0, 4.0, 5.0]);
        assert_eq!(chart.category_x(series, 0, plot, false), None);
        assert_eq!(chart.category_x(series, 1, plot, false), Some(10.0));
        assert_eq!(chart.category_x(series, 3, plot, false), Some(110.0));
        assert_eq!(chart.category_x(series, 4, plot, false), None);
        assert_eq!(
            chart.category_axis_tick_bounds(series, 1, plot),
            Some(Rect {
                x: 10.0,
                y: 50.0,
                width: 0.01,
                height: 4.0,
            })
        );

        let narrow_plot = Rect {
            width: 18.0,
            ..plot
        };
        assert!(chart.category_has_label(series, 1, narrow_plot, 18.0));
        assert!(!chart.category_has_label(series, 2, narrow_plot, 18.0));
        assert!(chart.category_has_label(series, 3, narrow_plot, 18.0));
    }

    #[test]
    fn category_axis_honors_authored_label_and_tick_skips() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:lineChart><c:ser>
              <c:cat><c:strLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt></c:numLit></c:val>
            </c:ser></c:lineChart><c:catAx><c:axPos val="b"/><c:majorTickMark val="out"/><c:tickLblSkip val="2"/><c:tickMarkSkip val="3"/></c:catAx></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let series = &chart.series[0];
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 50.0,
        };

        assert!(chart.category_has_label(series, 0, plot, 18.0));
        assert!(!chart.category_has_label(series, 1, plot, 18.0));
        assert!(chart.category_has_label(series, 2, plot, 18.0));
        assert!(chart.category_axis_tick_bounds(series, 0, plot).is_some());
        assert_eq!(chart.category_axis_tick_bounds(series, 1, plot), None);
        assert!(chart.category_axis_tick_bounds(series, 3, plot).is_some());
    }

    #[test]
    fn category_ticks_follow_cross_between_boundaries() {
        const PART: &str = "ppt/charts/chart4.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:ser>
              <c:cat><c:strLit><c:pt idx="0"><c:v>Anne</c:v></c:pt><c:pt idx="1"><c:v>Bob</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:pt idx="0"><c:v>9</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
            </c:ser></c:barChart><c:catAx><c:axPos val="b"/><c:majorTickMark val="out"/></c:catAx><c:valAx><c:axPos val="l"/><c:majorTickMark val="out"/><c:crossBetween val="between"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let series = &chart.series[0];
        let plot = Rect {
            x: 10.0,
            y: 20.0,
            width: 100.0,
            height: 50.0,
        };

        assert_eq!(
            chart.category_axis_tick_bounds(series, 0, plot).unwrap().x,
            10.0
        );
        assert_eq!(
            chart.category_axis_tick_bounds(series, 1, plot).unwrap().x,
            60.0
        );
        assert_eq!(
            chart.category_axis_tick_bounds(series, 2, plot).unwrap().x,
            110.0
        );
        assert_eq!(chart.category_axis_tick_bounds(series, 3, plot), None);
        assert_eq!(
            chart.value_axis_tick_bounds(plot, 0.5, false),
            Some(Rect {
                x: 6.0,
                y: 45.0,
                width: 4.0,
                height: 0.01,
            })
        );
    }

    #[test]
    fn line_chart_keeps_date_axis_title_and_month_labels() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Network Usage</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea>
              <c:lineChart><c:ser><c:marker><c:symbol val="square"/><c:size val="8"/></c:marker><c:cat><c:numRef><c:numCache><c:formatCode>mmm\ yyyy</c:formatCode><c:pt idx="0"><c:v>41521</c:v></c:pt><c:pt idx="1"><c:v>41551</c:v></c:pt></c:numCache></c:numRef></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2.37</c:v></c:pt><c:pt idx="1"><c:v>9.2</c:v></c:pt></c:numLit></c:val></c:ser><c:ser><c:marker><c:symbol val="diamond"/><c:size val="8"/></c:marker><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
              <c:dateAx><c:axPos val="b"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Months</a:t></a:r></a:p></c:rich></c:tx></c:title><c:numFmt formatCode="mmm\ yyyy"/><c:txPr><a:bodyPr rot="-5400000"/></c:txPr></c:dateAx>
              <c:valAx><c:axPos val="l"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Usage GB</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(chart.title, "Network Usage");
        assert_eq!(chart.horizontal_axis_options.title, "Months");
        assert_eq!(chart.value_axis_options.title, "Usage GB");
        assert_eq!(
            chart.horizontal_axis_options.label_rotation_degrees,
            Some(-90.0)
        );
        assert_eq!(chart.series[0].categories, ["Sep 2013", "Oct 2013"]);
        assert_eq!(chart.series[0].marker_symbol.as_deref(), Some("square"));
        assert_eq!(chart.series[1].marker_symbol.as_deref(), Some("diamond"));
        assert_eq!(chart.series[0].marker_size, Some(8.0 * 96.0 / 72.0));
        assert_eq!(chart.legend_marker(0), Some(("square", 8.0 * 96.0 / 72.0)));
    }

    #[test]
    fn line_chart_level_marker_assigns_series_shapes_without_overriding_none() {
        const PART: &str = "word/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:ser><c:marker><c:symbol val="none"/></c:marker><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser><c:marker val="1"/></c:lineChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(chart.series[0].marker_symbol.as_deref(), Some("diamond"));
        assert_eq!(chart.series[1].marker_symbol.as_deref(), Some("square"));
        assert_eq!(chart.series[2].marker_symbol, None);
    }

    #[test]
    fn logarithmic_axis_uses_power_ticks_and_positions() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>1000</c:v></c:pt><c:pt idx="2"><c:v>100000</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart><c:valAx><c:axPos val="l"/><c:scaling><c:logBase val="10"/></c:scaling></c:valAx></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(chart.value_axis(), (10.0, 100_000.0, 10.0));
        assert_eq!(
            chart.value_axis_ticks(),
            vec![
                (10.0, 0.0),
                (100.0, 0.25),
                (1000.0, 0.5),
                (10_000.0, 0.75),
                (100_000.0, 1.0)
            ]
        );
        assert_eq!(
            chart
                .line_points(
                    0,
                    Rect {
                        x: 0.0,
                        y: 0.0,
                        width: 100.0,
                        height: 100.0,
                    },
                    (10.0, 100_000.0),
                    false,
                )
                .iter()
                .map(|(_, y, _)| *y)
                .collect::<Vec<_>>(),
            [100.0, 50.0, 0.0]
        );
    }

    #[test]
    fn stacked_bar_series_lines_follow_cached_boundaries() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:grouping val="percentStacked"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:serLines/></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(
            chart_series_line_geometries(
                &chart,
                Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 100.0,
                    height: 100.0,
                },
                (0.0, 1.0),
            )
            .len(),
            2
        );
    }

    #[test]
    fn chart_axis_ignores_rotation_outside_office_range() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea>
              <c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
              <c:catAx><c:axPos val="b"/><c:txPr><a:bodyPr rot="-60000000"/></c:txPr></c:catAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(chart.horizontal_axis_options.label_rotation_degrees, None);
    }

    #[test]
    fn axis_line_no_fill_hides_only_that_axis() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart><c:catAx><c:axPos val="b"/><c:spPr><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr></c:catAx><c:valAx><c:axPos val="l"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:valAx></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert!(!chart.axis_line_visible(false));
        assert!(chart.axis_line_visible(true));
    }

    #[test]
    fn chartml_preserves_manual_layout_deleted_axes_secondary_scale_and_label_rotation() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea>
              <c:layout><c:manualLayout><c:x val="0.1"/><c:y val="0.2"/><c:w val="0.7"/><c:h val="0.6"/></c:manualLayout></c:layout>
              <c:bar3DChart><c:barDir val="col"/><c:ser><c:dLbls><c:txPr><a:bodyPr rot="5400000"/></c:txPr></c:dLbls><c:shape val="coneToMax"/><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="cat"/><c:axId val="left"/></c:bar3DChart>
              <c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>9</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="cat2"/><c:axId val="right"/></c:lineChart>
              <c:catAx><c:axId val="cat"/><c:delete val="1"/><c:axPos val="b"/></c:catAx>
              <c:valAx><c:axId val="left"/><c:axPos val="l"/><c:majorGridlines/><c:majorTickMark val="cross"/></c:valAx>
              <c:valAx><c:axId val="right"/><c:scaling><c:logBase val="10"/></c:scaling><c:axPos val="r"/><c:majorTickMark val="out"/></c:valAx>
            </c:plotArea><c:legend/></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert!(!chart.axis_line_visible(true));
        assert!(chart.value_axis_grid_lines_visible());
        assert!(chart.secondary_value_axis_ticks().is_some());
        assert_eq!(
            chart.value_axis_for_series(&chart.series[0]),
            (0.0, 2.5, 0.5)
        );
        assert_eq!(
            chart.value_axis_for_series(&chart.series[1]),
            (1.0, 10.0, 10.0)
        );
        assert!(chart.series[0].bar_cone_to_max);
        assert_eq!(chart.series[0].data_label_rotation_degrees, Some(90.0));
        assert_eq!(chart.legend_entries()[0].1, "Series 1");
        let layout = chart.plot_area_bounds(
            Rect {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 100.0,
            },
            Rect::default(),
        );
        assert!((layout.x - 20.0).abs() < 0.001);
        assert!((layout.y - 20.0).abs() < 0.001);
        assert!((layout.width - 140.0).abs() < 0.001);
        assert!((layout.height - 60.0).abs() < 0.001);
    }

    #[test]
    fn chart_defaults_match_office_title_legend_and_value_axis() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:title/><c:plotArea><c:barChart><c:ser>
              <c:val><c:numLit><c:pt idx="0"><c:v>4.3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt></c:numLit></c:val>
            </c:ser></c:barChart></c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart><c:spPr><a:noFill/></c:spPr>
            <c:txPr><a:p><a:pPr><a:defRPr sz="1800" b="1"><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert!(chart.show_title);
        assert!(chart.title.is_empty());
        assert_eq!(chart.title_text(), "Chart Title");
        assert_eq!(chart.legend_position, ChartLegendPosition::Bottom);
        assert!(chart.chart_area_no_fill);
        assert_eq!(chart.plot_area_color, None);
        assert_eq!(chart.value_axis(), (0.0, 6.0, 1.0));
        assert_eq!(chart.font_family.as_deref(), Some("Arial"));
        assert_eq!(chart.font_size, Some(24.0));
        assert!(chart.font_bold);
    }

    #[test]
    fn stacked_chart_axis_keeps_office_headroom_above_its_peak() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:areaChart><c:grouping val="stacked"/>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>32</c:v></c:pt><c:pt idx="1"><c:v>32</c:v></c:pt><c:pt idx="2"><c:v>28</c:v></c:pt><c:pt idx="3"><c:v>12</c:v></c:pt><c:pt idx="4"><c:v>15</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>12</c:v></c:pt><c:pt idx="2"><c:v>12</c:v></c:pt><c:pt idx="3"><c:v>21</c:v></c:pt><c:pt idx="4"><c:v>28</c:v></c:pt></c:numLit></c:val></c:ser>
            </c:areaChart><c:dateAx><c:axPos val="b"/><c:numFmt formatCode="m/d/yyyy"/></c:dateAx>
            <c:valAx><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:numFmt formatCode="General"/></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(chart.value_axis(), (0.0, 50.0, 5.0));
    }

    #[test]
    fn value_axis_adds_a_grid_interval_when_the_auto_ceiling_has_only_nine() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>9</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(chart.value_axis(), (0.0, 10.0, 1.0));
    }

    #[test]
    fn classic_dark_chart_styles_fill_the_plot_area() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:style val="33"/><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(chart.plot_area_color, Some(CLASSIC_DARK_CHART_PLOT_COLOR));
    }

    #[test]
    fn chart_title_gradient_fill_uses_shared_drawingml_paint() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Title</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="D9E2F3"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill></c:spPr></c:title><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let Some(ChartFill::LinearGradient {
            angle_degrees,
            angle_scaled,
            stops,
        }) = chart.title_fill
        else {
            panic!("chart title gradient fill was not parsed");
        };

        assert_eq!(angle_degrees, 90.0);
        assert!(angle_scaled);
        assert_eq!(stops.len(), 2);
        assert_eq!(stops[0].color, 0xffff_ffff);
        assert_eq!(stops[1].color, 0xd9e2_f3ff);
    }

    #[test]
    fn chart_area_and_legend_share_drawingml_gradient_fills() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea><c:legend><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="32CD32"/></a:gs><a:gs pos="50000"><a:srgbClr val="32CD32"><a:alpha val="0"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="32CD32"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></c:spPr></c:legend></c:chart><c:spPr><a:gradFill><a:gsLst><a:gs pos="20000"><a:srgbClr val="FFFF00"><a:alpha val="30000"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="4169E1"><a:alpha val="30000"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="3600000"/></a:gradFill></c:spPr></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        let Some(ChartFill::LinearGradient { stops, .. }) = chart.chart_area_fill else {
            panic!("chart-area gradient fill was not parsed");
        };
        assert_eq!(stops[0].color, 0xffff_004d);
        assert_eq!(stops[1].color, 0x4169_e14d);
        let Some(ChartFill::LinearGradient { stops, .. }) = chart.legend_fill else {
            panic!("legend gradient fill was not parsed");
        };
        assert_eq!(stops.len(), 3);
        assert_eq!(stops[1].color, 0x32cd_3200);
    }

    #[test]
    fn plot_area_gradient_fill_uses_the_shared_chart_fill_model() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="32CD32"><a:alpha val="0"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="32CD32"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></c:spPr><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        let Some(ChartFill::LinearGradient { stops, .. }) = chart.plot_area_fill else {
            panic!("plot-area gradient fill was not parsed");
        };
        assert_eq!(stops[0].color, 0x32cd_3200);
        assert_eq!(stops[1].color, 0x32cd_32ff);
    }

    #[test]
    fn supplied_scatter_standard_error_bars_keep_values_style_and_scale() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/testErrorBarProp.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let mut chart = parse_chart(&package, "xl/charts/chart1.xml", |name| match name {
            "accent1" => Some(0x4472_c4ff),
            "tx1" => Some(0x0000_00ff),
            _ => None,
        })
        .unwrap()
        .unwrap();
        let series = &chart.series[0];
        assert_eq!(series.name, "Aucun Coefficient");
        assert_eq!(series.color, Some(0x4472_c4ff));
        for (errors, expected, color, width) in [
            (
                series.x_error_bars.as_ref().unwrap(),
                0.9574271,
                0x5959_59ff,
                1.0,
            ),
            (
                series.y_error_bars.as_ref().unwrap(),
                0.2327611,
                0xff00_00ff,
                96.0 / 72.0,
            ),
        ] {
            assert_eq!(errors.plus.len(), 10);
            assert_eq!(errors.plus, errors.minus);
            assert!((errors.plus[0] - expected).abs() < 0.00001, "{errors:?}");
            assert!(matches!(errors.stroke, Some(ChartFill::Solid(actual)) if actual == color));
            assert!((errors.stroke_width - width).abs() < 0.001);
        }
        assert_eq!(chart.scatter_axis(true), (1918.0, 1932.0, 2.0));
        assert_eq!(chart.scatter_axis(false), (88.0, 91.5, 0.5));
        assert_eq!(chart.value_axis(), chart.scatter_axis(false));
        let plot = Rect {
            x: 40.0,
            y: 20.0,
            width: 700.0,
            height: 350.0,
        };
        let bars = chart.scatter_error_bars(series, plot);
        assert_eq!(bars.len(), 20);
        for (_, element) in bars {
            let Visual::PaintedShape {
                geometry: Geometry::Path { commands, .. },
                ..
            } = element.visual
            else {
                panic!("error bar must be a path");
            };
            assert_eq!(commands.len(), 6, "stem and two caps");
        }
        // A manual scale clips stems without inventing caps at the plot boundary.
        chart.horizontal_axis_options.minimum = Some(1920.0);
        chart.horizontal_axis_options.maximum = Some(1929.0);
        let bars = chart.scatter_error_bars(&chart.series[0], plot);
        let Visual::PaintedShape {
            geometry: Geometry::Path { commands, .. },
            ..
        } = &bars[0].1.visual
        else {
            panic!();
        };
        assert_eq!(commands.len(), 4);
        assert_eq!(bars[0].1.bounds.x, plot.x);
    }

    #[test]
    fn scatter_error_bars_honor_one_sided_caps_and_missing_observations() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/testErrorBarProp.xlsx"),
            Limits::default(),
        )
        .unwrap();
        let xml = String::from_utf8(
            package
                .required_part("xl/charts/chart1.xml")
                .unwrap()
                .to_vec(),
        )
        .unwrap();
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 700.0,
            height: 350.0,
        };
        for (direction, caps, command_count) in
            [("plus", "0", 4), ("minus", "0", 4), ("both", "1", 2)]
        {
            let xml = xml
                .replace(
                    "errBarType val=\"both\"",
                    &format!("errBarType val=\"{direction}\""),
                )
                .replace("noEndCap val=\"0\"", &format!("noEndCap val=\"{caps}\""));
            let bytes = stored_zip(&[("xl/charts/chart1.xml", xml.as_bytes())]);
            let package = Package::open(&bytes, Limits::default()).unwrap();
            let chart = parse_chart(&package, "xl/charts/chart1.xml", |_| None)
                .unwrap()
                .unwrap();
            let bars = chart.scatter_error_bars(&chart.series[0], plot);
            assert_eq!(bars.len(), 20);
            for (_, element) in bars {
                let Visual::PaintedShape {
                    geometry: Geometry::Path { commands, .. },
                    ..
                } = element.visual
                else {
                    panic!();
                };
                assert_eq!(commands.len(), command_count);
            }
        }
        for values in [vec![], vec![1.0], vec![f32::NAN, 1.0], vec![3.0, 3.0]] {
            let mut errors = super::ChartErrorBars {
                standard_error: true,
                ..Default::default()
            };
            errors.resolve(&values);
            assert_eq!(errors.amounts(0), (0.0, 0.0));
        }
        let mut errors = super::ChartErrorBars {
            standard_error: true,
            ..Default::default()
        };
        errors.resolve(&[1.0, f32::NAN, 3.0]);
        assert_eq!(errors.amounts(2), (1.0, 1.0));
    }

    #[test]
    fn scatter_chart_keeps_both_value_axes_and_linear_trendline() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:scatterChart><c:ser><c:xVal><c:numLit><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>100</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>0.6</c:v></c:pt><c:pt idx="1"><c:v>2.3</c:v></c:pt></c:numLit></c:yVal><c:trendline><c:trendlineType val="linear"/><c:dispRSqr val="1"/><c:dispEq val="1"/><c:trendlineLbl><c:layout><c:manualLayout><c:x val="0.137"/><c:y val="0.136"/></c:manualLayout></c:layout><c:txPr><a:p><a:pPr><a:defRPr sz="1400"/></a:pPr></a:p></c:txPr></c:trendlineLbl></c:trendline><c:errBars><c:errDir val="y"/><c:plus><c:numLit><c:pt idx="0"><c:v>0.1</c:v></c:pt><c:pt idx="1"><c:v>0.1</c:v></c:pt></c:numLit></c:plus><c:minus><c:numLit><c:pt idx="0"><c:v>0.1</c:v></c:pt><c:pt idx="1"><c:v>0.1</c:v></c:pt></c:numLit></c:minus></c:errBars></c:ser></c:scatterChart><c:valAx><c:axPos val="b"/><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="1600"/></a:pPr><a:r><a:rPr sz="1600"/><a:t>Dissolved Oxygen (%)</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx><c:valAx><c:axPos val="l"/><c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr sz="1600"/></a:pPr><a:r><a:rPr sz="1600"/><a:t>R1 (1/T1 sec-1)</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx><c:valAx><c:axPos val="r"/></c:valAx><c:valAx><c:axPos val="t"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert!(!chart.show_title);
        assert_eq!(chart.horizontal_axis_options.title, "Dissolved Oxygen (%)");
        assert_eq!(chart.value_axis_options.title, "R1 (1/T1 sec-1)");
        assert_eq!(
            chart.horizontal_axis_options.title_font_size,
            Some(16.0 * 96.0 / 72.0)
        );
        assert_eq!(chart.horizontal_axis_options.title_bold, None);
        assert_eq!(
            chart.value_axis_options.title_font_size,
            Some(16.0 * 96.0 / 72.0)
        );
        assert_eq!(chart.value_axis_options.title_bold, None);
        assert!(chart.series[0].linear_trendline);
        assert!(chart.series[0].show_trendline_equation);
        assert!(chart.series[0].show_trendline_r_squared);
        assert_eq!(chart.series[0].trendline_label_offset, Some((0.137, 0.136)));
        assert_eq!(
            chart.series[0].trendline_label_font_size,
            Some(14.0 * 96.0 / 72.0)
        );
        assert_eq!(
            chart.series[0].y_error_bars.as_ref().unwrap().plus,
            [0.1, 0.1]
        );
        assert_eq!(
            chart.series[0].y_error_bars.as_ref().unwrap().minus,
            [0.1, 0.1]
        );
        assert_eq!(chart.scatter_axis(true), (0.0, 120.0, 20.0));
        assert_eq!(chart.scatter_axis(false), (0.0, 3.0, 0.5));
    }

    #[test]
    fn chart_title_bitmap_fill_resolves_its_image_relationship() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[
            (
                PART,
                br#"<c:chartSpace xmlns:c="c" xmlns:a="a" xmlns:r="r"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Title</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:blipFill><a:blip r:embed="rId3"/><a:tile/></a:blipFill></c:spPr></c:title><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
            ),
            (
                "ppt/charts/_rels/chart1.xml.rels",
                br#"<Relationships xmlns="r"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/></Relationships>"#,
            ),
            ("ppt/media/image1.jpeg", b"\xff\xd8\xff"),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let Some(ChartFill::Image(Paint::Image {
            media_type,
            bytes,
            tile,
            ..
        })) = chart.title_fill
        else {
            panic!("chart title bitmap fill was not resolved");
        };

        assert_eq!(media_type, "image/jpeg");
        assert_eq!(bytes, b"\xff\xd8\xff");
        assert!(tile);
    }

    #[test]
    fn parses_every_ooxml_legend_position_into_shared_semantics() {
        for (value, expected) in [
            ("b", ChartLegendPosition::Bottom),
            ("l", ChartLegendPosition::Left),
            ("r", ChartLegendPosition::Right),
            ("t", ChartLegendPosition::Top),
            ("tr", ChartLegendPosition::TopRight),
        ] {
            let xml = format!(
                r#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea><c:legend><c:legendPos val="{value}"/></c:legend></c:chart></c:chartSpace>"#,
            );
            let bytes = stored_zip(&[("chart.xml", xml.as_bytes())]);
            let package = Package::open(&bytes, Limits::default()).unwrap();
            let chart = parse_chart(&package, "chart.xml", |_| None)
                .unwrap()
                .unwrap();
            assert_eq!(chart.legend_position, expected);
        }
    }

    #[test]
    fn stacked_area_legend_follows_visible_top_to_bottom_order() {
        const PART: &str = "xl/charts/chart19.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:areaChart><c:grouping val="stacked"/>
              <c:ser><c:tx><c:v>North</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:tx><c:v>South</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:tx><c:v>Bar</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
            </c:areaChart></c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(
            chart
                .legend_entries()
                .into_iter()
                .map(|(_, label, _)| label)
                .collect::<Vec<_>>(),
            ["Bar", "South", "North"]
        );
    }

    #[test]
    fn horizontal_bar_legend_follows_visible_top_to_bottom_order() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:barDir val="bar"/>
              <c:ser><c:tx><c:v>Bottom</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:tx><c:v>Top</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
            </c:barChart></c:plotArea><c:legend/></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(
            chart
                .legend_entries()
                .into_iter()
                .map(|(_, label, _)| label)
                .collect::<Vec<_>>(),
            ["Top", "Bottom"]
        );
    }

    #[test]
    fn preserves_horizontal_cone_bar_3d_semantics() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="1"/><c:perspective val="60"/></c:view3D><c:plotArea><c:bar3DChart><c:barDir val="bar"/><c:ser><c:tx><c:v>Series 1</c:v></c:tx>
              <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
              <c:dLbls><c:dLbl><c:idx val="0"/><c:delete val="1"/></c:dLbl><c:showVal val="1"/></c:dLbls>
            </c:ser><c:ser><c:tx><c:v>Series 2</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser><c:shape val="cone"/></c:bar3DChart></c:plotArea><c:legend/></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None)
            .expect("chart parses")
            .expect("bar3D chart is supported");

        assert_eq!(chart.series[0].kind, ChartKind::Bar);
        assert_eq!(chart.series[0].values, [1.0, 2.0]);
        assert!(chart.series[0].bar_horizontal);
        assert!(chart.series[0].bar_cone);
        assert!(chart.series[0].three_d);
        assert_eq!(
            chart.view_3d,
            Some(super::ChartView3D {
                perspective: 60,
                ..super::ChartView3D::default()
            })
        );
        assert_eq!(chart.series[0].hidden_labels, [0]);
        assert_eq!(
            chart
                .legend_entries()
                .into_iter()
                .map(|(_, label, _)| label)
                .collect::<Vec<_>>(),
            ["Series 2", "Series 1"]
        );
        let bounds = Rect {
            x: 0.0,
            y: 0.0,
            width: 40.0,
            height: 10.0,
        };
        assert!(matches!(
            super::chart_bar_geometry(bounds, true, true, false),
            Geometry::Path { .. }
        ));
        assert!(matches!(
            super::chart_bar_paint(bounds, 0x4472_c4ff, true, true, false),
            Paint::LinearGradient { .. }
        ));
        let cap = super::chart_bar_cone_cap(bounds, 0x4472_c4ff, true, true, 0.5).unwrap();
        assert!(matches!(cap.1, Geometry::Ellipse));
        assert_eq!(cap.0.height, bounds.height / 2.0);
    }

    #[test]
    fn horizontal_3d_bar_depth_is_bounded_by_bar_thickness() {
        let bounds = Rect {
            x: 10.0,
            y: 20.0,
            width: 300.0,
            height: 16.0,
        };
        let faces = super::chart_bar_3d_faces(
            bounds,
            0x4472_c4ff,
            super::ChartView3D {
                rot_x: 30,
                rot_y: 50,
                right_angle_axes: true,
                ..super::ChartView3D::default()
            },
            true,
        );

        assert!(faces[0].0.width - bounds.width <= bounds.height);
        assert!(faces[1].0.width <= bounds.height);
        let Paint::Solid(top_color) = faces[0].2 else {
            panic!("3D bar top face must have a solid color");
        };
        let brightness =
            |color: u32| (color >> 24) + ((color >> 16) & 0xff) + ((color >> 8) & 0xff);
        assert!(brightness(top_color) < brightness(0x4472_c4ff));
    }

    #[test]
    fn parses_bubble_coordinates_and_scales_area_from_size() {
        const PART: &str = "word/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:bubbleChart><c:ser><c:tx><c:v>Y-Values</c:v></c:tx><c:xVal><c:numLit><c:pt idx="0"><c:v>0.7</c:v></c:pt><c:pt idx="1"><c:v>1.8</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>2.7</c:v></c:pt><c:pt idx="1"><c:v>3.2</c:v></c:pt></c:numLit></c:yVal><c:bubbleSize><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:bubbleSize></c:ser></c:bubbleChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let bubbles = super::chart_bubble_bounds(
            &chart,
            0,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 200.0,
                height: 100.0,
            },
        );

        assert_eq!(chart.series[0].kind, ChartKind::Bubble);
        assert_eq!(chart.series[0].bubble_sizes, [10.0, 4.0]);
        assert_eq!(bubbles.len(), 2);
        assert!((bubbles[1].0.width / bubbles[0].0.width - (4.0_f32 / 10.0).sqrt()).abs() < 0.001);
    }

    #[test]
    fn preserves_cylinder_bar_3d_semantics() {
        const PART: &str = "word/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:bar3DChart><c:barDir val="col"/><c:ser>
              <c:val><c:numLit><c:pt idx="0"><c:v>4.3</c:v></c:pt></c:numLit></c:val>
            </c:ser><c:shape val="cylinder"/></c:bar3DChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert!(chart.series[0].bar_cylinder);
        let bounds = Rect {
            x: 0.0,
            y: 0.0,
            width: 20.0,
            height: 80.0,
        };
        assert!(matches!(
            super::chart_bar_geometry(bounds, false, false, true),
            Geometry::Path { .. }
        ));
        let Paint::LinearGradient { stops, .. } =
            super::chart_bar_paint(bounds, 0x4472_c4ff, false, false, true)
        else {
            panic!("cylinder must use a linear gradient");
        };
        assert!((stops[1].color >> 24) & 0xff < 128);
        let (_, _, Paint::Solid(cap)) =
            super::chart_bar_cylinder_cap(bounds, 0x4472_c4ff, false, true)
        else {
            panic!("cylinder cap must use a solid fill");
        };
        assert!((cap >> 24) & 0xff < 0x44);
    }

    #[test]
    fn standard_3d_bar_series_use_depth_slots_and_marker_depth() {
        const PART: &str = "word/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:bar3DChart><c:barDir val="col"/><c:grouping val="standard"/>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:gapWidth val="0"/>
            </c:bar3DChart><c:serAx><c:delete val="0"/></c:serAx></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 100.0,
        };
        let first = super::chart_bar_segment_bounds(
            &chart,
            0,
            0,
            plot,
            (0.0, 1.0),
            chart.bar_gap_width_percent,
        )
        .unwrap();
        let second = super::chart_bar_segment_bounds(
            &chart,
            1,
            0,
            plot,
            (0.0, 1.0),
            chart.bar_gap_width_percent,
        )
        .unwrap();

        assert!(first.width > plot.width / 2.0);
        assert!((first.width - second.width).abs() < 0.001);
        assert!(first.height < plot.height);
        assert_eq!(
            super::chart_3d_series_axis_labels(&chart, plot, 12.0).len(),
            2
        );
        assert!(second.x > first.x && second.x < first.x + first.width);
        assert!(second.y < first.y);
    }

    #[test]
    fn clustered_horizontal_3d_bars_end_at_the_office_axis_ceiling() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:bar3DChart><c:barDir val="bar"/><c:grouping val="clustered"/>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>6</c:v></c:pt></c:numLit></c:val></c:ser>
            </c:bar3DChart></c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(chart.value_axis(), (0.0, 6.0, 1.0));
        assert!(chart.uses_compact_horizontal_3d_layout());
    }

    #[test]
    fn horizontal_3d_bar_gridlines_join_the_back_wall_to_floor_ticks() {
        let plot = Rect {
            x: 10.0,
            y: 20.0,
            width: 100.0,
            height: 80.0,
        };
        let (bounds, Geometry::Path { commands, .. }) =
            super::chart_value_axis_grid_line(plot, 0.5, super::ChartView3D::default(), true, true)
        else {
            panic!("3D horizontal bar gridline must be projected");
        };

        assert_eq!(bounds.x, 60.0);
        assert_eq!(bounds.y, plot.y);
        assert!(bounds.width > 0.0);
        assert_eq!(bounds.height, plot.height);
        assert!(matches!(
            commands.as_slice(),
            [
                PathCommand::MoveTo { x, y: 0.0 },
                PathCommand::LineTo { x: x2, y },
                PathCommand::LineTo {
                    x: 0.0,
                    y: 80.0
                }
            ] if *x == bounds.width && *x2 == bounds.width && *y < plot.height
        ));

        let (_, _, fill) = super::chart_bar_3d_floor(plot, super::ChartView3D::default());
        assert!(matches!(fill, Paint::None));
        assert!(
            super::chart_bar_3d_walls(plot, super::ChartView3D::default())
                .iter()
                .all(|(_, _, fill)| matches!(fill, Paint::None))
        );

        let bar = Rect {
            x: plot.x,
            y: plot.y + 20.0,
            width: 50.0,
            height: 10.0,
        };
        let projected =
            super::chart_bar_3d_plane_bounds(bar, plot, super::ChartView3D::default(), true);
        assert!(projected.x > bar.x && projected.x < bar.x + bounds.width);
        assert!(projected.y < bar.y && projected.y > plot.y);
        assert_eq!(
            super::chart_bar_3d_plane_bounds(bar, plot, super::ChartView3D::default(), false,),
            bar
        );
    }

    #[test]
    fn perspective_3d_columns_follow_the_office_view_plane() {
        let plot = Rect {
            x: 10.0,
            y: 20.0,
            width: 120.0,
            height: 80.0,
        };
        let view = super::ChartView3D {
            rot_x: 15,
            rot_y: 20,
            right_angle_axes: false,
            perspective: 30,
            ..super::ChartView3D::default()
        };
        let left = Rect {
            x: 25.0,
            y: 50.0,
            width: 10.0,
            height: 50.0,
        };
        let right = Rect { x: 105.0, ..left };

        let left = super::chart_bar_3d_plane_bounds(left, plot, view, false);
        let right = super::chart_bar_3d_plane_bounds(right, plot, view, false);
        assert!(right.y > left.y + 10.0);

        let (_, Geometry::Path { commands, .. }) =
            super::chart_value_axis_grid_line(plot, 0.5, view, true, false)
        else {
            panic!("3D column gridline must be projected");
        };
        assert!(matches!(
            commands.as_slice(),
            [
                PathCommand::MoveTo { .. },
                PathCommand::LineTo { y: left_y, .. },
                PathCommand::LineTo { y: right_y, .. }
            ] if right_y > left_y
        ));

        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="0"/><c:perspective val="30"/></c:view3D><c:plotArea><c:bar3DChart><c:barDir val="col"/><c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser><c:shape val="cylinder"/><c:axId val="cat"/><c:axId val="series"/></c:bar3DChart><c:catAx><c:axId val="cat"/><c:axPos val="b"/><c:majorTickMark val="out"/></c:catAx><c:serAx><c:axId val="series"/><c:delete val="0"/></c:serAx></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        assert!(!chart.axis_line_visible(true));
        assert!(!chart.series_axis_line_visible());
        let tick = chart
            .category_axis_tick_bounds(&chart.series[0], 2, plot)
            .unwrap();
        assert!(tick.y > plot.y + plot.height + 5.0);
    }

    #[test]
    fn percent_stacked_bars_and_per_series_label_positions_share_geometry() {
        const PART: &str = "word/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="percentStacked"/>
              <c:ser><c:dLbls><c:showVal val="1"/></c:dLbls><c:val><c:numLit><c:pt idx="0"><c:v>4.3</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:dLbls><c:dLblPos val="inEnd"/><c:showVal val="1"/></c:dLbls><c:val><c:numLit><c:pt idx="0"><c:v>2.4</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:dLbls><c:dLblPos val="inBase"/><c:showVal val="1"/></c:dLbls><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
            </c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 40.0,
        };
        let bars = (0..3)
            .map(|index| {
                super::chart_bar_segment_bounds(
                    &chart,
                    index,
                    0,
                    plot,
                    (0.0, 1.0),
                    chart.bar_gap_width_percent,
                )
                .unwrap()
            })
            .collect::<Vec<_>>();

        assert_eq!(
            chart.series[1].data_label_position.as_deref(),
            Some("inEnd")
        );
        assert_eq!(
            chart.series[2].data_label_position.as_deref(),
            Some("inBase")
        );
        assert!((bars[0].width + bars[1].width + bars[2].width - plot.width).abs() < 0.001);
        assert!((bars[1].x - (bars[0].x + bars[0].width)).abs() < 0.001);
        assert!((bars[2].x - (bars[1].x + bars[1].width)).abs() < 0.001);
        let in_end = super::chart_bar_label_bounds(&chart.series[1], bars[1], 2.4);
        let in_base = super::chart_bar_label_bounds(&chart.series[2], bars[2], 2.0);
        assert!(in_end.x + in_end.width / 2.0 > bars[1].x + bars[1].width / 2.0);
        assert!(in_base.x + in_base.width / 2.0 < bars[2].x + bars[2].width / 2.0);
    }

    #[test]
    fn stacked_bar_keeps_authored_series_label_colors_and_value_widths() {
        const PART: &str = "xl/charts/chart10.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/>
              <c:ser><c:spPr><a:solidFill><a:srgbClr val="9999FF"/></a:solidFill></c:spPr><c:dLbls><c:txPr><a:bodyPr rot="5400000"/><a:p><a:pPr><a:defRPr sz="1025" b="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr><c:showVal val="1"/></c:dLbls><c:cat><c:strLit><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1000</c:v></c:pt><c:pt idx="1"><c:v>246.235253149996</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:spPr><a:solidFill><a:srgbClr val="993366"/></a:solidFill></c:spPr><c:dLbls><c:txPr><a:bodyPr rot="5400000"/><a:p><a:pPr><a:defRPr sz="1025" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr><c:showVal val="1"/></c:dLbls><c:cat><c:strLit><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>926.700928129913</c:v></c:pt><c:pt idx="1"><c:v>145.772846206949</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:dLbls><c:showVal val="1"/><c:showCatName val="0"/></c:dLbls>
            </c:barChart><c:valAx><c:axId val="value"/><c:scaling><c:max val="18000"/></c:scaling><c:axPos val="b"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 1_800.0,
            height: 200.0,
        };
        let first = super::chart_bar_segment_bounds(
            &chart,
            0,
            0,
            plot,
            (0.0, 18_000.0),
            chart.bar_gap_width_percent,
        )
        .unwrap();
        let second = super::chart_bar_segment_bounds(
            &chart,
            1,
            0,
            plot,
            (0.0, 18_000.0),
            chart.bar_gap_width_percent,
        )
        .unwrap();

        assert_eq!(chart.series[0].color, Some(0x9999_ffff));
        assert_eq!(chart.series[1].color, Some(0x9933_66ff));
        assert_eq!(chart.series[0].data_label_text_color, Some(0x0000_00ff));
        assert_eq!(chart.series[1].data_label_text_color, Some(0xffff_ffff));
        assert!((first.width - 100.0).abs() < 0.001);
        assert!((second.width - 92.670_09).abs() < 0.001);
    }

    #[test]
    fn clustered_bar_automatic_labels_use_office_outside_end() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>
              <c:ser><c:tx><c:v>North</c:v></c:tx><c:dLbls><c:numFmt formatCode="General" sourceLinked="0"/><c:showVal val="1"/></c:dLbls><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>1000</c:v></c:pt></c:numLit></c:val></c:ser>
            </c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let bar = Rect {
            x: 20.0,
            y: 10.0,
            width: 80.0,
            height: 20.0,
        };
        let label = super::chart_bar_label_bounds(&chart.series[0], bar, 1000.0);

        assert!(label.x >= bar.x + bar.width);
        let negative_label = super::chart_bar_label_bounds(&chart.series[0], bar, -1000.0);
        assert!(negative_label.x + negative_label.width <= bar.x);
    }

    #[test]
    fn percent_stacked_area_uses_shared_normalized_bands() {
        const PART: &str = "word/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:areaChart><c:grouping val="percentStacked"/>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>32</c:v></c:pt><c:pt idx="1"><c:v>28</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>12</c:v></c:pt></c:numLit></c:val></c:ser>
            </c:areaChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 100.0,
        };
        let (_, _, first_labels) = super::chart_area_geometry(&chart, 0, plot, (0.0, 1.0))
            .expect("first area band");
        let (_, _, second_labels) = super::chart_area_geometry(&chart, 1, plot, (0.0, 1.0))
            .expect("second area band");

        assert_eq!(chart.series[0].grouping, ChartGrouping::PercentStacked);
        assert!((first_labels[0].1 - 63.636_364).abs() < 0.001);
        assert!((second_labels[0].1 - 13.636_364).abs() < 0.001);
    }

    #[test]
    fn area_3d_keeps_categories_on_the_horizontal_axis() {
        const PART: &str = "word/charts/chart1.xml";
        let xml = br#"<c:chartSpace xmlns:c="c"><c:style val="11"/><c:chart><c:view3D><c:perspective val="30"/></c:view3D><c:plotArea><c:area3DChart><c:grouping val="stacked"/>
            <c:ser><c:tx><c:v>Series 1</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>1/5/2002</c:v></c:pt><c:pt idx="1"><c:v>1/6/2002</c:v></c:pt><c:pt idx="2"><c:v>1/7/2002</c:v></c:pt><c:pt idx="3"><c:v>1/8/2002</c:v></c:pt><c:pt idx="4"><c:v>1/9/2002</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>32</c:v></c:pt><c:pt idx="1"><c:v>32</c:v></c:pt><c:pt idx="2"><c:v>28</c:v></c:pt><c:pt idx="3"><c:v>12</c:v></c:pt><c:pt idx="4"><c:v>15</c:v></c:pt></c:numLit></c:val></c:ser>
            <c:ser><c:tx><c:v>Series 2</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>1/5/2002</c:v></c:pt><c:pt idx="1"><c:v>1/6/2002</c:v></c:pt><c:pt idx="2"><c:v>1/7/2002</c:v></c:pt><c:pt idx="3"><c:v>1/8/2002</c:v></c:pt><c:pt idx="4"><c:v>1/9/2002</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>12</c:v></c:pt><c:pt idx="2"><c:v>12</c:v></c:pt><c:pt idx="3"><c:v>21</c:v></c:pt><c:pt idx="4"><c:v>28</c:v></c:pt></c:numLit></c:val></c:ser>
            <c:axId val="cat"/><c:axId val="val"/></c:area3DChart><c:catAx><c:axId val="cat"/><c:axPos val="b"/></c:catAx><c:valAx><c:axId val="val"/><c:axPos val="l"/><c:crossBetween val="midCat"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#;
        let bytes = stored_zip(&[(PART, xml)]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |name| {
            (name == "accent1").then_some(0x4f81_bdff)
        })
        .unwrap()
        .unwrap();
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 600.0,
            height: 400.0,
        };
        let axis = chart.value_axis();
        let view = chart.view_3d.unwrap();
        let first =
            super::chart_area_3d_faces(&chart, 0, plot, (axis.0, axis.1), view, 0x4f81_bdff);
        let second =
            super::chart_area_3d_faces(&chart, 1, plot, (axis.0, axis.1), view, 0x95b3_d7ff);
        let (series_labels, category_labels) = super::chart_area_3d_axis_labels(&chart, plot, 12.0);

        assert_eq!(axis, (0.0, 50.0, 5.0));
        assert_eq!(chart.series[0].color, Some(0x4f81_bdff));
        assert_eq!(chart.series[1].color, Some(0x95b3_d7ff));
        assert!(first.iter().any(|(bounds, _, _)| bounds.width > 400.0));
        assert!(second.iter().any(|(bounds, _, _)| bounds.width > 400.0));
        assert_eq!(second[0].0, first[0].0, "stacked bands meet on one plane, as in the native chart4 reference");
        let mut standard = chart.clone();
        for series in &mut standard.series { series.grouping = ChartGrouping::Standard; }
        let first = super::chart_area_3d_faces(&standard, 0, plot, (axis.0, axis.1), view, 0x4f81_bdff);
        let second = super::chart_area_3d_faces(&standard, 1, plot, (axis.0, axis.1), view, 0x95b3_d7ff);
        assert!(second[0].0.x > first[0].0.x);
        assert!(second[0].0.y < first[0].0.y);
        assert!(series_labels.is_empty());
        assert_eq!(category_labels.len(), 5);
        assert_eq!(category_labels[0].2, "1/5/2002");
        assert_eq!(category_labels[0].3, -45.0);
        assert!(category_labels[0].1.x < category_labels[4].1.x);
        assert_eq!(category_labels[0].1.y, category_labels[4].1.y);
    }

    #[test]
    fn standard_area_geometry_uses_the_final_axis_ceiling() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:areaChart><c:grouping val="standard"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>32</c:v></c:pt><c:pt idx="1"><c:v>32</c:v></c:pt></c:numLit></c:val></c:ser><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>12</c:v></c:pt></c:numLit></c:val></c:ser></c:areaChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 100.0,
        };
        let (minimum, maximum, _) = chart.value_axis();
        let (_, points, _) = super::chart_area_geometry(&chart, 0, plot, (minimum, maximum))
            .expect("standard area");

        assert_eq!((minimum, maximum), (0.0, 35.0));
        assert!((points[0].1 - 8.571_428).abs() < 0.001);
    }

    #[test]
    fn mixed_stacked_area_and_columns_share_the_office_visible_axis() {
        const PART: &str = "xl/charts/chart3.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea>
              <c:areaChart><c:grouping val="stacked"/><c:ser><c:tx><c:v>North</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>Foo</c:v></c:pt><c:pt idx="1"><c:v>Bar</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="area-cat"/><c:axId val="area-val"/></c:areaChart>
              <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:tx><c:v>South</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>Foo</c:v></c:pt><c:pt idx="1"><c:v>Bar</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="bar-cat"/><c:axId val="bar-val"/></c:barChart>
              <c:catAx><c:axId val="area-cat"/><c:delete val="0"/><c:axPos val="b"/></c:catAx><c:valAx><c:axId val="area-val"/><c:delete val="0"/><c:axPos val="l"/><c:crossBetween val="between"/></c:valAx>
              <c:catAx><c:axId val="bar-cat"/><c:delete val="1"/><c:axPos val="b"/></c:catAx><c:valAx><c:axId val="bar-val"/><c:delete val="1"/><c:axPos val="l"/></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let axis = chart.value_axis();
        let area_axis = chart.value_axis_for_series(&chart.series[0]);
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 100.0,
            height: 100.0,
        };
        let (_, points, _) =
            super::chart_area_geometry(&chart, 0, plot, (area_axis.0, area_axis.1))
                .expect("mixed chart area");

        assert_eq!(axis, (0.0, 4.5, 0.5));
        assert_eq!(area_axis, axis);
        let area_plot = chart.area_plot_bounds(&chart.series[0], plot).unwrap();
        assert_eq!((area_plot.x, area_plot.width), (25.0, 50.0));
        assert!((points[0].1 - 77.777_78).abs() < 0.001);
        assert!((points[1].1 - 55.555_557).abs() < 0.001);
    }

    #[test]
    fn dual_axes_use_the_same_office_ranges_for_ticks_and_series() {
        const PART: &str = "xl/charts/chart4.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea>
              <c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="cat-left"/><c:axId val="left"/></c:lineChart>
              <c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="cat-right"/><c:axId val="right"/></c:lineChart>
              <c:catAx><c:axId val="cat-left"/><c:axPos val="b"/></c:catAx><c:valAx><c:axId val="left"/><c:axPos val="l"/></c:valAx>
              <c:catAx><c:axId val="cat-right"/><c:delete val="1"/><c:axPos val="b"/></c:catAx><c:valAx><c:axId val="right"/><c:axPos val="r"/></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let primary = chart.value_axis();
        let secondary = chart.secondary_value_axis_ticks().unwrap().1;

        assert_eq!(primary, (0.0, 2.5, 0.5));
        assert_eq!(chart.value_axis_for_series(&chart.series[0]), primary);
        assert_eq!(secondary.first().unwrap().0, 0.0);
        assert_eq!(secondary.last().unwrap().0, 4.5);
    }

    #[test]
    fn parses_chartex_treemap_cached_sizes_and_leaf_labels() {
        const PART: &str = "word/charts/chartEx1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<cx:chartSpace xmlns:cx="cx"><cx:chartData><cx:data id="0">
              <cx:strDim type="cat">
                <cx:lvl><cx:pt idx="0">Leaf 1</cx:pt><cx:pt idx="1">Leaf 2</cx:pt></cx:lvl>
                <cx:lvl><cx:pt idx="0">Branch 1</cx:pt><cx:pt idx="1">Branch 1</cx:pt></cx:lvl>
              </cx:strDim>
              <cx:numDim type="size"><cx:lvl><cx:pt idx="0">22</cx:pt><cx:pt idx="1">12</cx:pt></cx:lvl></cx:numDim>
            </cx:data></cx:chartData><cx:chart><cx:title/><cx:plotArea><cx:plotAreaRegion>
              <cx:series layoutId="treemap"><cx:tx><cx:txData><cx:v>Series 1</cx:v></cx:txData></cx:tx><cx:dataId val="0"/></cx:series>
            </cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart_ex(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(chart.series[0].kind, ChartKind::Treemap);
        assert_eq!(chart.series[0].values, [22.0, 12.0]);
        assert_eq!(chart.series[0].categories, ["Leaf 1", "Leaf 2"]);
        assert_eq!(chart.title_text(), "Chart Title");
    }

    #[test]
    fn parses_bar_of_pie_and_resolves_chartex_box_whisker_data() {
        const PIE: &str = "xl/charts/chart1.xml";
        const BOX: &str = "xl/charts/chartEx1.xml";
        let bytes = stored_zip(&[
            (
                PIE,
                br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:ofPieChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>9</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt><c:pt idx="2"><c:v>7</c:v></c:pt><c:pt idx="3"><c:v>6</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt><c:pt idx="5"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser><c:splitType val="pos"/><c:splitPos val="4"/></c:ofPieChart></c:plotArea></c:chart></c:chartSpace>"#,
            ),
            (
                BOX,
                br#"<cx:chartSpace xmlns:cx="cx" xmlns:a="a"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title><cx:txPr><a:p><a:r><a:t>Box</a:t></a:r></a:p></cx:txPr></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="boxWhisker"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion><cx:axis><cx:catScaling gapWidth="2.47"/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>"#,
            ),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let pie = parse_chart(&package, PIE, |_| None).unwrap().unwrap();
        let box_chart = parse_chart_ex_with_data(
            &package,
            BOX,
            |_| None,
            |name| {
                (name == "_xlchart.v1.0").then(|| ChartSourceData {
                    values: vec![1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 100.0],
                    category_levels: Vec::new(),
                })
            },
        )
        .unwrap()
        .unwrap();

        assert_eq!(pie.series[0].kind, ChartKind::BarOfPie(2));
        assert_eq!(box_chart.series[0].kind, ChartKind::BoxWhisker);
        assert_eq!(
            box_chart.series[0].values,
            [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 100.0]
        );
        assert_eq!(box_chart.title, "Box");
        assert_eq!(box_chart.category_gap_width, Some(2.47));
        let geometry = chart_box_whisker(
            &box_chart.series[0].values,
            0,
            1,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 100.0,
                height: 100.0,
            },
            (0.0, 20.0),
            2.47,
        )
        .unwrap();
        assert_eq!(geometry.outlier_y.len(), 1);
    }

    #[test]
    fn preserves_pie_3d_and_its_theme_point_colors() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:view3D><c:rotX val="30"/><c:rAngAx val="0"/><c:perspective val="30"/></c:view3D><c:plotArea><c:pie3DChart><c:varyColors val="1"/><c:ser><c:explosion val="25"/>
              <c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
            </c:ser></c:pie3DChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |name| match name {
            "accent1" => Some(0x4f81_bdff),
            "accent2" => Some(0xc050_4dff),
            _ => None,
        })
        .expect("chart parses")
        .expect("pie3D chart is supported");

        assert!(chart.series[0].three_d);
        assert_eq!(chart.series[0].point_colors, [0x4f81_bdff, 0xc050_4dff]);
        assert_eq!(chart.series[0].point_explosions, [0.25, 0.25]);
        let view = chart.view_3d.expect("3D view is preserved");
        let (_, slices) = chart_pie_slices(
            &chart.series[0],
            (50.0, 50.0),
            (40.0, 20.0),
            8.0,
            view.pie_depth_perspective(),
            None,
        );
        let (_, Geometry::Path { commands, .. }) = slices
            .iter()
            .find(|slice| slice.index == 0)
            .expect("first data point remains addressable after painter sorting")
            .cut_side
            .as_ref()
            .expect("exploded 3D pie has cut faces")
        else {
            panic!("3D pie cut face must be a path");
        };
        let [
            PathCommand::MoveTo { x: center_x, .. },
            PathCommand::LineTo { x: outer_x, .. },
            PathCommand::LineTo {
                x: projected_outer_x,
                ..
            },
            PathCommand::LineTo {
                x: projected_center_x,
                ..
            },
            ..,
        ] = commands.as_slice()
        else {
            panic!("3D pie cut face must contain a projected quad");
        };
        assert!((center_x - outer_x).abs() < 0.001);
        assert!((projected_outer_x - outer_x).abs() < 0.001);
        assert!(projected_center_x > center_x);
    }

    #[test]
    fn chartml_preserves_category_crossing_deleted_legend_entries_and_solid_pie_sides() {
        let line_bytes = stored_zip(&[(
            "line.xml",
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:lineChart><c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart><c:valAx><c:crossBetween val="between"/></c:valAx></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&line_bytes, Limits::default()).unwrap();
        let line = parse_chart(&package, "line.xml", |_| None)
            .unwrap()
            .unwrap();
        let points = line.line_points(
            0,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 100.0,
                height: 100.0,
            },
            (0.0, 2.0),
            false,
        );
        assert_eq!(points[0].0, 25.0);

        let pie_bytes = stored_zip(&[(
            "pie.xml",
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:pie3DChart><c:ser><c:explosion val="25"/><c:cat><c:strLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>8.2</c:v></c:pt><c:pt idx="1"><c:v>3.2</c:v></c:pt><c:pt idx="2"><c:v>1.4</c:v></c:pt><c:pt idx="3"><c:v>1.2</c:v></c:pt></c:numLit></c:val></c:ser></c:pie3DChart></c:plotArea><c:legend><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry></c:legend></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&pie_bytes, Limits::default()).unwrap();
        let pie = parse_chart(&package, "pie.xml", |_| None).unwrap().unwrap();
        assert_eq!(
            pie.legend_entries()
                .into_iter()
                .map(|(_, label, _)| label)
                .collect::<Vec<_>>(),
            ["1", "3", "4"]
        );
        let (_, slices) =
            chart_pie_slices(&pie.series[0], (50.0, 50.0), (40.0, 20.0), 8.0, 0.125, None);
        assert_eq!(
            slices.iter().map(|slice| slice.index).collect::<Vec<_>>(),
            [3, 2, 0, 1],
            "exploded 3D pie slices must be ordered back to front",
        );
        assert!(
            slices
                .iter()
                .filter_map(|slice| slice.side.as_ref())
                .all(|(bounds, _)| bounds.x.is_finite()
                    && bounds.y.is_finite()
                    && bounds.width.is_finite()
                    && bounds.height.is_finite()
                    && bounds.width >= 0.0
                    && bounds.height >= 0.0)
        );
        for slice in &slices {
            let (bounds, Geometry::Path { commands, .. }) = slice
                .cut_side
                .as_ref()
                .expect("exploded 3D pie has cut faces")
            else {
                panic!("3D pie cut face must be a path");
            };
            let radial_faces = commands
                .iter()
                .filter(|command| matches!(command, PathCommand::MoveTo { .. }))
                .count();
            assert!(
                (1..=2).contains(&radial_faces),
                "exploded 3D pie slice must close every visible radial face"
            );
            assert!(bounds.width > 0.0 && bounds.height > 0.0);
        }
    }

    #[test]
    fn explicit_pie_series_and_data_point_colors_override_vary_colors() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:pieChart><c:varyColors val="1"/><c:ser>
              <c:spPr><a:solidFill><a:schemeClr val="bg2"/></a:solidFill></c:spPr>
              <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:schemeClr val="tx2"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr></c:dPt>
              <c:val><c:numLit><c:pt idx="0"><c:v>66</c:v></c:pt><c:pt idx="1"><c:v>33</c:v></c:pt></c:numLit></c:val>
            </c:ser><c:firstSliceAng val="238"/></c:pieChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |name| match name {
            "bg2" => Some(0xe7e6_e6ff),
            "tx2" => Some(0x4454_6aff),
            "accent1" => Some(0x4472_c4ff),
            "accent2" => Some(0xed7d_31ff),
            _ => None,
        })
        .unwrap()
        .unwrap();

        assert_eq!(chart.series[0].point_colors, [0xe7e6_e6ff, 0x4454_6aff]);
        assert_eq!(
            chart.series[0].point_border_colors,
            [None, Some(0x0000_ffff)]
        );
        assert_eq!(chart.series[0].point_border_widths, [0.0, 2.0]);
        assert_eq!(chart.series[0].first_slice_angle, 238.0);
        let dark_slice_midpoint = chart.series[0].pie_start_angle_radians()
            + std::f32::consts::TAU * (66.0 / 99.0 + 33.0 / 99.0 / 2.0);
        assert!((dark_slice_midpoint.sin() - 0.999_391).abs() < 0.001);
    }

    #[test]
    fn pie_data_point_pattern_fill_is_preserved() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:pieChart><c:ser>
              <c:dPt><c:idx val="0"/><c:spPr><a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="000000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill></c:spPr></c:dPt>
              <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
            </c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert!(matches!(
            &chart.series[0].point_fills[0],
            Some(ChartFill::Pattern {
                preset,
                foreground: 0x0000_00ff,
                background: 0xffff_ffff,
            }) if preset == "pct50"
        ));
    }

    #[test]
    fn explicit_centered_pie_labels_do_not_gain_leader_lines() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:pieChart><c:ser>
              <c:dLbls><c:dLblPos val="ctr"/><c:showCatName val="1"/><c:showPercent val="1"/></c:dLbls>
              <c:cat><c:strLit><c:pt idx="0"><c:v>Small</c:v></c:pt><c:pt idx="1"><c:v>Large</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt></c:numLit></c:val>
            </c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let (_, slices) = chart_pie_slices(
            &chart.series[0],
            (300.0, 300.0),
            (200.0, 200.0),
            0.0,
            0.0,
            None,
        );
        let labels = chart_pie_label_layout(
            &chart,
            &chart.series[0],
            &slices,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 600.0,
                height: 600.0,
            },
            200.0,
            200.0,
        );

        assert!(labels.iter().all(|label| label.leader.is_none()));
    }

    #[test]
    fn unpositioned_pie_labels_do_not_gain_leader_lines() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:pieChart><c:ser>
              <c:dLbls><c:showCatName val="1"/><c:showPercent val="1"/><c:showLeaderLines val="1"/></c:dLbls>
              <c:cat><c:strLit><c:pt idx="0"><c:v>Revenue</c:v></c:pt><c:pt idx="1"><c:v>A</c:v></c:pt><c:pt idx="2"><c:v>H</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>87</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val>
            </c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let (_, slices) = chart_pie_slices(
            &chart.series[0],
            (300.0, 300.0),
            (200.0, 200.0),
            0.0,
            0.0,
            None,
        );
        let labels = chart_pie_label_layout(
            &chart,
            &chart.series[0],
            &slices,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 600.0,
                height: 600.0,
            },
            200.0,
            200.0,
        );

        assert!(labels.iter().all(|label| label.leader.is_none()));
        let revenue = labels.iter().find(|label| label.index == 0).unwrap();
        let center = (
            revenue.bounds.x + revenue.bounds.width / 2.0,
            revenue.bounds.y + revenue.bounds.height / 2.0,
        );
        let slice = slices.iter().find(|slice| slice.index == 0).unwrap();
        assert!((center.0 - slice.center.0).hypot(center.1 - slice.center.1) > 150.0);
    }

    #[test]
    fn chart_path_gradient_and_authored_3d_layout_are_preserved() {
        const PART: &str = "xl/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:view3D><c:rotX val="25"/><c:hPercent val="50"/><c:rotY val="0"/><c:perspective val="14"/></c:view3D><c:plotArea><c:pie3DChart><c:ser>
              <c:dLbls><c:dLblPos val="outEnd"/><c:showCatName val="1"/><c:showLeaderLines val="1"/></c:dLbls>
              <c:cat><c:strLit><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt><c:pt idx="2"><c:v>East</c:v></c:pt><c:pt idx="3"><c:v>West</c:v></c:pt><c:pt idx="4"><c:v>Central</c:v></c:pt><c:pt idx="5"><c:v>Mountain</c:v></c:pt><c:pt idx="6"><c:v>Pacific</c:v></c:pt><c:pt idx="7"><c:v>Northeast</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:pt idx="0"><c:v>38</c:v></c:pt><c:pt idx="1"><c:v>42</c:v></c:pt><c:pt idx="2"><c:v>36</c:v></c:pt><c:pt idx="3"><c:v>40</c:v></c:pt><c:pt idx="4"><c:v>8</c:v></c:pt><c:pt idx="5"><c:v>7</c:v></c:pt><c:pt idx="6"><c:v>6</c:v></c:pt><c:pt idx="7"><c:v>5</c:v></c:pt></c:numLit></c:val>
            </c:ser></c:pie3DChart></c:plotArea></c:chart><c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="CCE0F5"/></a:gs><a:gs pos="100000"><a:srgbClr val="0066CC"/></a:gs></a:gsLst><a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></c:spPr></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        assert!(matches!(
            chart.chart_area_fill.as_ref().unwrap().paint(Rect {
                x: 0.0,
                y: 0.0,
                width: 600.0,
                height: 400.0,
            }),
            Paint::RectGradient { .. }
        ));
        let view = chart.view_3d.unwrap();
        assert!(view.pie_vertical_ratio() > 0.38);
        let (_, slices) = chart_pie_slices(
            &chart.series[0],
            (300.0, 220.0),
            (180.0, 180.0 * view.pie_vertical_ratio()),
            180.0 * view.pie_depth_ratio(),
            view.pie_depth_perspective(),
            None,
        );
        let labels = chart_pie_label_layout(
            &chart,
            &chart.series[0],
            &slices,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 600.0,
                height: 440.0,
            },
            180.0,
            180.0 * view.pie_vertical_ratio(),
        );
        for (index, left) in labels.iter().enumerate() {
            for right in labels.iter().skip(index + 1) {
                assert!(
                    left.bounds.x + left.bounds.width <= right.bounds.x
                        || right.bounds.x + right.bounds.width <= left.bounds.x
                        || left.bounds.y + left.bounds.height <= right.bounds.y
                        || right.bounds.y + right.bounds.height <= left.bounds.y,
                    "outside pie labels overlap: {} and {}",
                    left.text,
                    right.text,
                );
            }
        }

        assert!(labels.iter().all(|label| {
            label
                .leader
                .as_ref()
                .is_none_or(|(bounds, _)| bounds.width < 90.0)
        }));

        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:depthPercent val="500"/><c:rAngAx val="1"/></c:view3D><c:plotArea><c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val></c:ser><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>9</c:v></c:pt></c:numLit></c:val></c:ser><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>8</c:v></c:pt></c:numLit></c:val></c:ser><c:gapWidth val="150"/><c:gapDepth val="0"/></c:bar3DChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let plot = Rect {
            x: 0.0,
            y: 0.0,
            width: 600.0,
            height: 400.0,
        };
        let first = super::chart_bar_segment_bounds(
            &chart,
            0,
            0,
            plot,
            (0.0, 10.0),
            chart.bar_gap_width_percent,
        )
        .unwrap();
        let second = super::chart_bar_segment_bounds(
            &chart,
            1,
            0,
            plot,
            (0.0, 10.0),
            chart.bar_gap_width_percent,
        )
        .unwrap();
        let deep_spacing = second.x - first.x;
        assert!(deep_spacing > 0.0, "clustered series retain their left-to-right order");
        let mut shallow = chart.clone();
        shallow.view_3d.as_mut().unwrap().depth_percent = Some(100);
        let shallow_first = super::chart_bar_segment_bounds(&shallow, 0, 0, plot, (0.0, 10.0), shallow.bar_gap_width_percent).unwrap();
        let shallow_second = super::chart_bar_segment_bounds(&shallow, 1, 0, plot, (0.0, 10.0), shallow.bar_gap_width_percent).unwrap();
        assert!(shallow_second.x - shallow_first.x > deep_spacing,
            "fitting an authored 500% depth must compress projected category spacing");
    }

    #[test]
    fn clustered_gap_width_is_measured_in_bar_widths() {
        const PART: &str = "word/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
              <c:gapWidth val="150"/>
            </c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let bar = super::chart_bar_segment_bounds(
            &chart,
            0,
            0,
            Rect {
                x: 0.0,
                y: 0.0,
                width: 100.0,
                height: 350.0,
            },
            (0.0, 1.0),
            chart.bar_gap_width_percent,
        )
        .unwrap();

        assert!((bar.height - 100.0).abs() < 0.001);
    }

    #[test]
    fn pie_manual_layout_and_point_explosion_share_geometry() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:layout><c:manualLayout><c:x val="0.1"/><c:y val="0.2"/><c:w val="0.6"/><c:h val="0.5"/></c:manualLayout></c:layout><c:pieChart><c:ser><c:dPt><c:idx val="1"/><c:explosion val="17"/></c:dPt><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let pie = chart.pie_bounds(
            Rect {
                x: 100.0,
                y: 200.0,
                width: 400.0,
                height: 300.0,
            },
            Rect::default(),
            false,
        );
        let (_, slices) = chart_pie_slices(
            &chart.series[0],
            (0.0, 0.0),
            (pie.width / 2.0, pie.height / 2.0),
            0.0,
            0.0,
            None,
        );
        let exploded = slices.iter().find(|slice| slice.index == 1).unwrap();

        assert_eq!(
            pie,
            Rect {
                x: 185.0,
                y: 260.0,
                width: 150.0,
                height: 150.0
            }
        );
        let pie_3d = chart.pie_bounds(
            Rect {
                x: 100.0,
                y: 200.0,
                width: 400.0,
                height: 300.0,
            },
            Rect::default(),
            true,
        );
        assert!((pie_3d.x + pie_3d.width / 2.0 - 260.0).abs() < 0.001);
        assert!((pie_3d.y - 260.0).abs() < 0.001);
        assert!((pie_3d.width * 1.17 - 240.0).abs() < 0.001);
        assert!((pie_3d.height - 150.0).abs() < 0.001);
        assert!((exploded.center.0 + 9.016).abs() < 0.001);
        assert!((exploded.center.1 + 9.016).abs() < 0.001);
    }

    #[test]
    fn chart_data_table_and_explicit_value_axis_share_semantics() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:lineChart><c:ser><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>&quot;$&quot;#,##0;[Red]\(&quot;$&quot;#,##0\)</c:formatCode><c:pt idx="0"><c:v>-479325</c:v></c:pt><c:pt idx="1"><c:v>685146</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart><c:valAx><c:scaling><c:max val="700000"/></c:scaling><c:numFmt formatCode="&quot;$&quot;#,##0;[Red]\(&quot;$&quot;#,##0\)"/></c:valAx><c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/></c:dTable></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();
        let table = chart.data_table.unwrap();

        assert!(table.show_horizontal_borders);
        assert!(table.show_vertical_borders);
        assert!(table.show_outline);
        assert!(table.show_keys);
        assert_eq!(chart.value_axis_options.maximum, Some(700_000.0));
        assert_eq!(chart.value_axis(), (-700_000.0, 700_000.0, 200_000.0));
        assert!(
            chart.series[0]
                .number_format
                .as_deref()
                .is_some_and(|format| format.contains('$'))
        );
    }

    #[test]
    fn line_series_and_axis_labels_use_their_explicit_properties() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:lineChart><c:ser><c:spPr><a:ln w="101600"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill></a:ln></c:spPr><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart><c:valAx><c:txPr><a:p><a:pPr><a:defRPr sz="1200" b="0"/></a:pPr></a:p></c:txPr></c:valAx></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |name| {
            (name == "tx1").then_some(0x0000_00ff)
        })
        .unwrap()
        .unwrap();

        assert_eq!(chart.series[0].color, Some(0x4040_40ff));
        assert_eq!(chart.series[0].stroke_width, Some(8.0 * 96.0 / 72.0));
        assert_eq!(chart.value_axis_options.label_font_size, Some(16.0));
        assert_eq!(chart.value_axis_options.label_bold, Some(false));
    }

    #[test]
    fn data_point_line_color_transforms_do_not_mutate_its_fill() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:pieChart><c:ser>
              <c:dPt><c:idx val="0"/><c:spPr>
                <a:solidFill><a:schemeClr val="accent3"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill>
                <a:ln><a:solidFill><a:schemeClr val="accent3"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></a:ln>
              </c:spPr></c:dPt>
              <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
            </c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let accent3 = 0x2453_ffff;
        let chart = parse_chart(&package, PART, |name| {
            (name == "accent3").then_some(accent3)
        })
        .unwrap()
        .unwrap();
        let mut expected = accent3;
        apply_color_transform(&mut expected, "lumMod", 0.2);
        apply_color_transform(&mut expected, "lumOff", 0.8);

        assert_eq!(chart.series[0].point_colors, [expected]);
    }

    #[test]
    fn legacy_chart_styles_select_grayscale_and_accent_palettes() {
        const PIE: &str = "ppt/charts/pie.xml";
        const BAR: &str = "ppt/charts/bar.xml";
        const REPEATED_BAR: &str = "ppt/charts/repeated-bar.xml";
        const EXTENDED_BAR: &str = "ppt/charts/extended-bar.xml";
        let bytes = stored_zip(&[
            (PIE, br#"<c:chartSpace xmlns:c="c"><c:style val="1"/><c:chart><c:plotArea><c:pieChart><c:varyColors val="1"/><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>"#),
            (BAR, br#"<c:chartSpace xmlns:c="c"><c:style val="4"/><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#),
            (REPEATED_BAR, br#"<c:chartSpace xmlns:c="c"><c:style val="13"/><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#),
            (EXTENDED_BAR, br#"<c:chartSpace xmlns:c="c" xmlns:c14="c14"><c14:style val="102"/><c:style val="2"/><c:chart><c:plotArea><c:bar3DChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser><c:shape val="cylinder"/></c:bar3DChart></c:plotArea></c:chart></c:chartSpace>"#),
        ]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let color = |name: &str| match name {
            "dk1" => Some(0x0000_00ff),
            "accent1" => Some(0x4472_c4ff),
            "accent2" => Some(0xc050_4dff),
            "accent3" => Some(0xde6c_36ff),
            "accent4" => Some(0x8064_a2ff),
            _ => None,
        };

        let pie = parse_chart(&package, PIE, color).unwrap().unwrap();
        let bar = parse_chart(&package, BAR, color).unwrap().unwrap();
        let repeated_bar = parse_chart(&package, REPEATED_BAR, color).unwrap().unwrap();
        let extended_bar = parse_chart(&package, EXTENDED_BAR, color).unwrap().unwrap();

        assert_eq!(pie.series[0].point_colors.len(), 2);
        assert_ne!(pie.series[0].point_colors[0], 0x4472_c4ff);
        assert_eq!(bar.series[0].color, Some(0xc050_4dff));
        assert_eq!(repeated_bar.series[0].color, Some(0xde6c_36ff));
        assert_eq!(extended_bar.series[0].color, Some(0x4472_c4ff));
    }

    #[test]
    fn supplied_word_chart_series_labels_override_group_defaults() {
        let package = Package::open(
            include_bytes!("../../../tests/fixtures/word-data-label-borders.docx"),
            Limits::default(),
        )
        .unwrap();
        let chart = parse_chart(&package, "word/charts/chart1.xml", |_| None)
            .unwrap()
            .unwrap();
        assert!(
            chart.series[0].show_values,
            "series showVal must override chart group false"
        );
        assert_eq!(
            chart.series[0].data_labels.len(),
            2,
            "point border overrides must survive group defaults"
        );
        assert_eq!(
            chart.series[0].data_labels[0].border_color,
            Some(0xffff00ff)
        );
        let chart2 = parse_chart(&package, "word/charts/chart2.xml", |_| None)
            .unwrap()
            .unwrap();
        let (color, width) = chart2.series[0].data_label_border.unwrap();
        assert_eq!(color, 0x00ff00ff);
        assert!((width - 44450.0 / 9525.0).abs() < 0.001);
        for (index, value) in chart.series[0].values.iter().copied().enumerate() {
            let label = super::chart_bar_data_label(
                &chart,
                &chart.series[0],
                index,
                Rect {
                    x: 0.0,
                    y: 0.0,
                    width: 200.0,
                    height: 40.0,
                },
                value,
            )
            .unwrap();
            assert_eq!(label.1, super::format_chart_value(value, None));
        }
    }

    #[test]
    fn preserves_chart_text_and_data_label_settings() {
        const PART: &str = "ppt/charts/chart1.xml";
        let bytes = stored_zip(&[(
            PART,
            br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart>
              <c:title><c:tx><c:rich><a:p><a:r><a:t>Approval</a:t></a:r></a:p></c:rich></c:tx></c:title>
              <c:plotArea><c:barChart><c:ser>
                <c:dLbls><c:dLbl><c:idx val="0"/><c:layout><c:manualLayout><c:x val="0.2"/><c:y val="-0.1"/><c:w val="0.4"/><c:h val="0.3"/></c:manualLayout></c:layout><c:tx><c:rich><a:p><a:r><a:rPr sz="2000" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Custom</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:ln w="22860"><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:ln></c:spPr><c:numFmt formatCode="0.0%"/><c:dLblPos val="ctr"/><c:showVal val="0"/><c:showPercent val="1"/></c:dLbl><c:dLbl><c:idx val="1"/></c:dLbl><c:txPr><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="404040"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr><c:numFmt formatCode="0%"/><c:showVal val="1"/><c:showCatName val="1"/></c:dLbls>
                <c:cat><c:strLit><c:pt idx="0"><c:v>Text only</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>0.29</c:v></c:pt></c:numLit></c:val>
              </c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        )]);
        let package = Package::open(&bytes, Limits::default()).unwrap();
        let chart = parse_chart(&package, PART, |_| None).unwrap().unwrap();

        assert_eq!(chart.title, "Approval");
        assert_eq!(chart.series[0].categories, ["Text only"]);
        assert!(chart.series[0].show_values);
        assert!(chart.series[0].show_category_name);
        assert_eq!(chart.series[0].number_format.as_deref(), Some("0%"));
        let label = &chart.series[0].data_labels[0];
        assert_eq!((label.index, label.text.as_str()), (0, "Custom"));
        assert_eq!(label.text_color, Some(0xff00_00ff));
        assert_eq!(label.font_bold, Some(true));
        assert_eq!(label.manual_offset, Some((0.2, -0.1)));
        assert_eq!(label.manual_size, Some((0.4, 0.3)));
        assert_eq!(label.border_color, Some(0xffff_00ff));
        assert_eq!(label.show_values, Some(false));
        assert_eq!(label.show_percent, Some(true));
        assert_eq!(label.number_format.as_deref(), Some("0.0%"));
        assert_eq!(label.position.as_deref(), Some("ctr"));
        assert!((label.font_size.unwrap() - 20.0 * 96.0 / 72.0).abs() < 0.001);
        assert!((label.border_width - 2.4).abs() < 0.001);
        let inherited =
            chart.data_label_text_style(&chart.series[0], chart.series[0].data_labels.get(1));
        assert_eq!(inherited.color, 0x4040_40ff);
        assert_eq!(inherited.font_size, 12.0);
        assert!(inherited.bold);
    }

    #[test]
    fn waterfall_axis_uses_office_nice_major_units() {
        let chart = Chart {
            source_part: String::new(),
            date_1904: false,
            title: String::new(),
            title_fill: None,
            title_position: None,
            title_font_size: None,
            title_font_bold: None,
            title_text_color: None,
            series: vec![ChartSeries {
                kind: ChartKind::Waterfall,
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
                name: String::new(),
                category_levels: Vec::new(),
                categories: Vec::new(),
                x_values: Vec::new(),
                values: vec![100.0, 20.0, 50.0, -40.0, 130.0, -60.0, 70.0, 140.0],
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
                smooth: false,
                marker_symbol: None,
                marker_size: None,
                subtotals: vec![0, 4, 7],
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
            }],
            show_title: true,
            show_legend: true,
            legend_position: ChartLegendPosition::Top,
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
            data_label_position: Some("inEnd".to_owned()),
            style: None,
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
            category_gap_width: None,
        };

        assert_eq!(chart.value_axis(), (0.0, 180.0, 20.0));
    }
}

#[test]
fn shared_dash_lengths_keep_thin_and_wide_stroke_semantics() {
    assert_eq!(
        drawingml_dash_lengths(&[4.0, 3.0, 1.0, 3.0], 0.5),
        [4.0, 3.0, 1.0, 3.0]
    );
    assert_eq!(
        drawingml_dash_lengths(&[4.0, 3.0, 1.0, 3.0], 3.0),
        [12.0, 9.0, 3.0, 9.0]
    );
    assert!(drawingml_dash_lengths(&[], 8.0).is_empty());
}
