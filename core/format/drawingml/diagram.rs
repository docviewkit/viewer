#[derive(Debug)]
pub(super) struct Diagram {
    pub(super) data_part: String,
    pub(super) diagnostics: Vec<Diagnostic>,
    pub(super) drawing_parts: Vec<(String, bool)>,
    pub(super) drawing_text_colors: HashMap<String, u32>,
    pub(super) role_line_colors: HashMap<(String, String), u32>,
    pub(super) role_fills: HashMap<(String, String), DiagramStyle>,
    pub(super) layout_type: Option<String>,
    pub(super) right_to_left: bool,
    pub(super) scene_three_d: bool,
    pub(super) background_fill: Option<ChartFill>,
    pub(super) background_shadow: Option<Shadow>,
    pub(super) nodes: Vec<DiagramNode>,
}

#[derive(Debug, Default)]
pub(super) struct DiagramStyle {
    pub(super) fill: Option<ChartFill>,
    pub(super) shadow: Option<Shadow>,
    pub(super) effects: DrawingMlPictureEffects,
    pub(super) line_width: Option<f32>,
    pub(super) text_color: Option<u32>,
}

#[derive(Debug)]
pub(super) struct DiagramNode {
    pub(super) model_id: String,
    pub(super) parent_id: Option<String>,
    pub(super) sibling_order: u64,
    pub(super) assistant: bool,
    pub(super) hierarchy_branch: Option<String>,
    pub(super) parent_transition_id: Option<String>,
    pub(super) bold: bool,
    pub(super) font_size: Option<f32>,
    pub(super) text: String,
    pub(super) placeholder: bool,
    pub(super) preset_geometry: Option<String>,
    pub(super) custom_geometry: bool,
    pub(super) fill: Option<ChartFill>,
    pub(super) no_fill: bool,
    pub(super) shadow: Option<Shadow>,
    pub(super) three_d: Option<ThreeDStyle>,
    pub(super) fill_scheme: Option<String>,
    pub(super) fill_transforms: Vec<(String, f32)>,
}

pub(super) fn diagram_fallback_elements(
    diagram: &Diagram,
    bounds: Rect,
    accent: u32,
) -> Option<Vec<DrawingMlElement>> {
    diagram_org_chart_elements(diagram, bounds, accent)
        .or_else(|| diagram_target_list_elements(diagram, bounds, accent))
        .or_else(|| diagram_horizontal_list_elements(diagram, bounds, accent))
        .or_else(|| diagram_vertical_list_elements(diagram, bounds, accent))
        .or_else(|| diagram_semantic_elements(diagram, bounds, accent))
}

// orgChart1's built-in hierarchy constraints use boxes of height 0.5w and 0.21w
// spacing. Assistants occupy paired branches; shallow ordinary branches hang
// vertically. Work in node-width units, then fit the entire tree to the frame.
pub(super) fn diagram_org_chart_elements(
    diagram: &Diagram,
    bounds: Rect,
    accent: u32,
) -> Option<Vec<DrawingMlElement>> {
    if diagram.layout_type.as_deref()?.split('#').next()?
        != "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1"
        || diagram.nodes.is_empty()
    {
        return None;
    }
    let nodes = &diagram.nodes;
    let ids = nodes
        .iter()
        .enumerate()
        .map(|(i, node)| (node.model_id.as_str(), i))
        .collect::<HashMap<_, _>>();
    let mut children = vec![Vec::new(); nodes.len()];
    let mut roots = Vec::new();
    for (i, node) in nodes.iter().enumerate() {
        if let Some(parent) = node.parent_id.as_deref().and_then(|id| ids.get(id)) {
            children[*parent].push(i);
        } else {
            roots.push(i);
        }
    }
    for siblings in children.iter_mut().chain(std::iter::once(&mut roots)) {
        siblings.sort_by_key(|i| nodes[*i].sibling_order);
    }
    let root_ids = roots.iter().copied().collect::<HashSet<_>>();
    // Iterative postorder avoids a call-stack limit on deeply nested documents.
    let mut order = Vec::new();
    let mut pending = roots.clone();
    let mut seen = vec![false; nodes.len()];
    while let Some(i) = pending.pop() {
        if std::mem::replace(&mut seen[i], true) {
            return None;
        }
        order.push(i);
        pending.extend(&children[i]);
    }
    if order.len() != nodes.len() {
        return None;
    }
    let mut sizes = vec![(1.0_f32, 0.5_f32); nodes.len()];
    let mut root_x = vec![0.0_f32; nodes.len()];
    let mut offsets = vec![(0.0_f32, 0.0_f32); nodes.len()];
    let mut depth = vec![0_usize; nodes.len()];
    let mut hanging = vec![false; nodes.len()];
    for &i in order.iter().rev() {
        depth[i] = children[i].iter().map(|j| depth[*j] + 1).max().unwrap_or(0);
        let assistants = children[i]
            .iter()
            .copied()
            .filter(|j| nodes[*j].assistant)
            .collect::<Vec<_>>();
        let ordinary = children[i]
            .iter()
            .copied()
            .filter(|j| !nodes[*j].assistant)
            .collect::<Vec<_>>();
        let mut y = 0.71_f32;
        for pair in assistants.chunks(2) {
            for (side, &j) in pair.iter().enumerate() {
                offsets[j] = (
                    if side == 0 {
                        0.5 - 0.105 - sizes[j].0
                    } else {
                        0.5 + 0.105
                    },
                    y,
                );
            }
            y += pair.iter().map(|j| sizes[*j].1).fold(0.0_f32, f32::max) + 0.21;
        }
        let branch = nodes[i].hierarchy_branch.as_deref().unwrap_or("init");
        // Unsupported authored branch modes retain the diagnostic generic path.
        if !matches!(branch, "init" | "std") {
            return None;
        }
        hanging[i] = branch == "init" && !root_ids.contains(&i) && depth[i] <= 1;
        if hanging[i] {
            let x = if nodes[i].assistant || !assistants.is_empty() {
                0.65
            } else {
                0.25
            };
            for &j in &ordinary {
                offsets[j] = (x, y);
                y += sizes[j].1 + 0.21;
            }
        } else if !ordinary.is_empty() {
            let mut x = 0.0;
            let mut previous: Option<usize> = None;
            for &j in &ordinary {
                // A leaf occupies only the root row: a preceding subtree's
                // deeper hanging branch cannot collide with it.
                if children[j].is_empty()
                    && let Some(k) = previous
                {
                    x = offsets[k].0 + root_x[k] + 1.21;
                }
                offsets[j] = (x, y);
                x += sizes[j].0 + 0.21;
                previous = Some(j);
            }
            let first = ordinary[0];
            let last = *ordinary.last()?;
            let centre = (offsets[first].0 + root_x[first] + offsets[last].0 + root_x[last]) / 2.0;
            for &j in &ordinary {
                offsets[j].0 -= centre;
            }
        }
        let left = children[i]
            .iter()
            .map(|j| offsets[*j].0)
            .fold(0.0_f32, f32::min);
        let right = children[i]
            .iter()
            .map(|j| offsets[*j].0 + sizes[*j].0)
            .fold(1.0_f32, f32::max);
        let bottom = children[i]
            .iter()
            .map(|j| offsets[*j].1 + sizes[*j].1)
            .fold(0.5_f32, f32::max);
        root_x[i] = -left;
        sizes[i] = (right - left, bottom);
        for &j in &children[i] {
            offsets[j].0 -= left;
        }
    }
    let width = roots.iter().map(|i| sizes[*i].0).sum::<f32>()
        + 0.21 * roots.len().saturating_sub(1) as f32;
    let height = roots.iter().map(|i| sizes[*i].1).fold(0.5_f32, f32::max);
    let scale = (bounds.width / width).min(bounds.height / height);
    let left = bounds.x + (bounds.width - width * scale) / 2.0;
    let top = bounds.y + (bounds.height - height * scale) / 2.0;
    let mut x = 0.0;
    for &i in &roots {
        offsets[i] = (x, 0.0);
        x += sizes[i].0 + 0.21;
    }
    let mut rects = vec![bounds; nodes.len()];
    for &i in &order {
        let (x, y) = offsets[i];
        let node_x = x + root_x[i];
        rects[i] = Rect {
            x: left
                + if diagram.right_to_left {
                    width - node_x - 1.0
                } else {
                    node_x
                } * scale,
            y: top + y * scale,
            width: scale,
            height: 0.5 * scale,
        };
        for &j in &children[i] {
            offsets[j].0 += x;
            offsets[j].1 += y;
        }
    }
    let mut connector_colors = HashMap::<&str, (&str, u32)>::new();
    for ((model, role), &color) in &diagram.role_line_colors {
        let entry = connector_colors
            .entry(model.as_str())
            .or_insert((role.as_str(), color));
        if role.as_str() < entry.0 {
            *entry = (role.as_str(), color);
        }
    }
    let mut elements = Vec::new();
    for (i, node) in nodes.iter().enumerate() {
        let Some(&parent) = node.parent_id.as_deref().and_then(|id| ids.get(id)) else {
            continue;
        };
        let p = rects[parent];
        let r = rects[i];
        let start = (
            p.x + p.width
                * if hanging[parent] && !node.assistant && !nodes[parent].assistant
                && !children[parent].iter().any(|child| nodes[*child].assistant) {
                    if diagram.right_to_left { 0.9 } else { 0.1 }
                } else {
                    0.5
                },
            p.y + p.height,
        );
        let points = if node.assistant || hanging[parent] {
            let end_x = if r.x + r.width / 2.0 < start.0 {
                r.x + r.width
            } else {
                r.x
            };
            vec![
                start,
                (start.0, r.y + r.height / 2.0),
                (end_x, r.y + r.height / 2.0),
            ]
        } else {
            let end = (r.x + r.width / 2.0, r.y);
            let bend = r.y - scale * 0.105;
            vec![start, (start.0, bend), (end.0, bend), end]
        };
        let x = points.iter().map(|p| p.0).fold(f32::INFINITY, f32::min);
        let y = points.iter().map(|p| p.1).fold(f32::INFINITY, f32::min);
        let connector_bounds = Rect {
            x,
            y,
            width: (points.iter().map(|p| p.0).fold(x, f32::max) - x).max(1.0),
            height: (points.iter().map(|p| p.1).fold(y, f32::max) - y).max(1.0),
        };
        let color = node
            .parent_transition_id
            .as_ref()
            .and_then(|id| connector_colors.get(id.as_str()).map(|(_, color)| *color))
            .unwrap_or(accent);
        elements.push(DrawingMlElement {
            bounds: connector_bounds,
            text: None,
            visual: Visual::PaintedShape {
                geometry: Geometry::Path {
                    fill_rule: FillRule::NonZero,
                    commands: points
                        .iter()
                        .enumerate()
                        .map(|(index, point)| {
                            if index == 0 {
                                PathCommand::MoveTo {
                                    x: point.0 - x,
                                    y: point.1 - y,
                                }
                            } else {
                                PathCommand::LineTo {
                                    x: point.0 - x,
                                    y: point.1 - y,
                                }
                            }
                        })
                        .collect(),
                },
                fill: Paint::None,
                stroke: Paint::Solid(color),
                stroke_width: 2.0,
            },
        });
    }
    for (i, node) in nodes.iter().enumerate() {
        let role = if root_ids.contains(&i) {
            "rootText1"
        } else if node.assistant {
            "rootText3"
        } else {
            "rootText"
        };
        let key = (node.model_id.clone(), role.to_owned());
        let style = diagram.role_fills.get(&key);
        let fill = node
            .fill
            .as_ref()
            .or_else(|| style.and_then(|style| style.fill.as_ref()));
        let rect = rects[i];
        let font_size = node
            .font_size
            .unwrap_or((rect.height * 0.28).min(100.0 * 4.0 / 3.0));
        let visual = Visual::TextLayout {
            layout: TextLayout {
                vertical_align: TextVerticalAlign::Center,
                auto_fit: crate::model::TextAutoFit::Shrink,
                ..TextLayout::default()
            },
            visual: Box::new(Visual::RichText {
                geometry: Geometry::Rectangle,
                fill: if node.no_fill {
                    Paint::None
                } else {
                    fill.map_or(Paint::Solid(accent), |fill| fill.paint(rect))
                },
                stroke: if style.and_then(|s| s.line_width) == Some(0.0) {
                    Paint::None
                } else {
                    Paint::Solid(
                        diagram
                            .role_line_colors
                            .get(&key)
                            .copied()
                            .unwrap_or(0xffff_ffff),
                    )
                },
                stroke_width: style.and_then(|s| s.line_width).unwrap_or(2.0 * 4.0 / 3.0),
                align: TextAlign::Center,
                line_height: font_size * 1.2,
                runs: vec![TextRun {
                    text: node.text.clone(),
                    font_family: "Calibri".to_owned(),
                    font_size,
                    color: style.and_then(|s| s.text_color).unwrap_or(0xffff_ffff),
                    bold: node.bold,
                    italic: false,
                    underline: false,
                    strikethrough: false,
                    highlight: 0,
                    baseline_shift: 0.0,
                    letter_spacing: 0.0,
                    horizontal_scale: 1.0,
                    east_asian_line_breaks: true,
                    paint: None,
                }],
            }),
        };
        let visual = if let Some(style) = style {
            style.effects.clone().wrap(visual, None)
        } else {
            visual
        };
        let shadow = node.shadow.or_else(|| style.and_then(|style| style.shadow));
        elements.push(DrawingMlElement {
            bounds: rect,
            text: Some(node.text.clone()),
            visual: if shadow.is_some() {
                Visual::Effect {
                    shadow,
                    clip: None,
                    visual: Box::new(visual),
                }
            } else {
                visual
            },
        });
    }
    Some(elements)
}

pub(super) fn diagram_target_list_elements(
    diagram: &Diagram,
    bounds: Rect,
    accent: u32,
) -> Option<Vec<DrawingMlElement>> {
    let layout = diagram.layout_type.as_deref()?.split('#').next()?;
    if layout != "urn:microsoft.com/office/officeart/2005/8/layout/target3" {
        return None;
    }
    let branches = diagram_branches(diagram)?;
    if branches.is_empty() || branches.len() > 7 {
        return None;
    }
    let roots = branches.iter().map(|(node, _)| *node).collect::<Vec<_>>();
    let groups = branches
        .iter()
        .map(|(_, descendants)| diagram_bullet_text(descendants))
        .collect::<Vec<_>>();

    // Built-in target3 constraints: circle1 diameter = 0.6w, vertical spacer
    // = 0.05 diameter; each successive circle loses 1/n of diameter+spacer.
    // Rectangles share the circles' tops/bottoms and begin at their centre X.
    let count = roots.len();
    let diameter = (bounds.width * 0.6).min(bounds.height).max(1.0);
    let gap = if count > 1 { diameter * 0.05 } else { 0.0 };
    let top = bounds.y + (bounds.height - diameter) / 2.0;
    let panel_x = bounds.x + diameter / 2.0;
    let panel_width = bounds.x + bounds.width - panel_x;
    let row_height = (diameter - gap * (count - 1) as f32) / count as f32;
    let primary_size = (row_height / 3.0).min(65.0 * 4.0 / 3.0);
    let secondary_size = primary_size / 5.0;
    let split_text = groups.iter().any(|text| !text.is_empty());
    let mut elements = Vec::new();
    for (index, node) in roots.iter().enumerate() {
        let size = diameter - (diameter + gap) * index as f32 / count as f32;
        let circle = Rect {
            x: panel_x - size / 2.0,
            y: top + index as f32 * row_height,
            width: size,
            height: size,
        };
        elements.push(DrawingMlElement {
            bounds: circle,
            text: None,
            visual: Visual::PaintedShape {
                geometry: Geometry::Ellipse,
                fill: node
                    .fill
                    .as_ref()
                    .map_or(Paint::Solid(accent), |fill| fill.paint(circle)),
                stroke: Paint::Solid(0xffff_ffff),
                stroke_width: 2.0,
            },
        });
    }
    // Opaque panels cover the right halves, leaving the authored semicircles.
    for index in 0..count {
        let size = diameter - (diameter + gap) * index as f32 / count as f32;
        elements.push(DrawingMlElement {
            bounds: Rect {
                x: panel_x,
                y: top + index as f32 * row_height,
                width: panel_width,
                height: size,
            },
            text: None,
            visual: Visual::PaintedShape {
                geometry: Geometry::Rectangle,
                fill: Paint::Solid(0xffff_ffff),
                stroke: Paint::Solid(accent),
                stroke_width: 2.0,
            },
        });
    }
    for (index, node) in roots.iter().enumerate() {
        let width = if split_text {
            panel_width / 2.0
        } else {
            panel_width
        };
        for (text, x, width, font_size, bold, align) in [
            (
                node.text.clone(),
                panel_x,
                width,
                primary_size,
                node.bold,
                TextAlign::Center,
            ),
            (
                groups[index].clone(),
                panel_x + width,
                panel_width - width,
                secondary_size,
                false,
                TextAlign::Start,
            ),
        ] {
            if text.is_empty() {
                continue;
            }
            let inset = font_size * 0.3;
            elements.push(DrawingMlElement {
                bounds: Rect {
                    x: x + inset,
                    y: top + index as f32 * row_height + inset,
                    width: (width - 2.0 * inset).max(1.0),
                    height: (row_height - 2.0 * inset).max(1.0),
                },
                text: Some(text.clone()),
                visual: Visual::TextLayout {
                    layout: TextLayout {
                        vertical_align: TextVerticalAlign::Center,
                        auto_fit: crate::model::TextAutoFit::Shrink,
                        ..TextLayout::default()
                    },
                    visual: Box::new(Visual::RichText {
                        geometry: Geometry::Rectangle,
                        fill: Paint::None,
                        stroke: Paint::None,
                        stroke_width: 0.0,
                        align,
                        line_height: font_size * 1.2,
                        runs: vec![TextRun {
                            paint: None,
                            text,
                            font_family: "Calibri".to_owned(),
                            font_size,
                            color: 0x0000_00ff,
                            bold,
                            italic: false,
                            underline: false,
                            strikethrough: false,
                            highlight: 0,
                            baseline_shift: 0.0,
                            letter_spacing: 0.0,
                            horizontal_scale: 1.0,
                            east_asian_line_breaks: true,
                        }],
                    }),
                },
            });
        }
    }
    Some(elements)
}

pub(super) fn diagram_vertical_list_elements(
    diagram: &Diagram,
    bounds: Rect,
    accent: u32,
) -> Option<Vec<DrawingMlElement>> {
    diagram
        .layout_type
        .as_deref()
        .is_some_and(|value| value.contains("/vList6"))
        .then(|| {
            let ids = diagram
                .nodes
                .iter()
                .map(|node| node.model_id.as_str())
                .collect::<HashSet<_>>();
            let nodes = diagram
                .nodes
                .iter()
                .filter(|node| {
                    node.parent_id
                        .as_deref()
                        .is_none_or(|parent| !ids.contains(parent))
                })
                .collect::<Vec<_>>();
            let count = nodes.len().max(1);
            let gap = (bounds.height / 24.0).clamp(6.0, 18.0);
            let row_height = ((bounds.height - gap * (count + 1) as f32) / count as f32).max(1.0);
            let mut elements = Vec::with_capacity(count.saturating_mul(2));
            for (index, node) in nodes.into_iter().enumerate() {
                let row_y = bounds.y + gap + index as f32 * (row_height + gap);
                let arrow = Rect {
                    x: bounds.x + bounds.width * 0.34,
                    y: row_y + row_height * 0.16,
                    width: bounds.width * 0.62,
                    height: row_height * 0.68,
                };
                let head = arrow.height;
                elements.push(DrawingMlElement {
                    bounds: arrow,
                    text: None,
                    visual: Visual::PaintedShape {
                        geometry: Geometry::Path {
                            fill_rule: FillRule::NonZero,
                            commands: vec![
                                PathCommand::MoveTo {
                                    x: 0.0,
                                    y: arrow.height * 0.22,
                                },
                                PathCommand::LineTo {
                                    x: arrow.width - head,
                                    y: arrow.height * 0.22,
                                },
                                PathCommand::LineTo {
                                    x: arrow.width - head,
                                    y: 0.0,
                                },
                                PathCommand::LineTo {
                                    x: arrow.width,
                                    y: arrow.height / 2.0,
                                },
                                PathCommand::LineTo {
                                    x: arrow.width - head,
                                    y: arrow.height,
                                },
                                PathCommand::LineTo {
                                    x: arrow.width - head,
                                    y: arrow.height * 0.78,
                                },
                                PathCommand::LineTo {
                                    x: 0.0,
                                    y: arrow.height * 0.78,
                                },
                                PathCommand::ClosePath,
                            ],
                        },
                        fill: Paint::Solid(0xb4c7_e7ff),
                        stroke: Paint::None,
                        stroke_width: 0.0,
                    },
                });
                let panel = Rect {
                    x: bounds.x + bounds.width * 0.04,
                    y: row_y,
                    width: bounds.width * 0.42,
                    height: row_height,
                };
                elements.push(DrawingMlElement {
                    bounds: panel,
                    text: None,
                    visual: Visual::PaintedShape {
                        geometry: Geometry::RoundedRectangle {
                            radius_x: panel.height * 0.12,
                            radius_y: panel.height * 0.12,
                        },
                        fill: Paint::Solid(accent),
                        stroke: Paint::None,
                        stroke_width: 0.0,
                    },
                });
                if !node.text.trim().is_empty() {
                    let text = node.text.clone();
                    elements.push(DrawingMlElement {
                        bounds: Rect {
                            x: panel.x + panel.width * 0.08,
                            y: panel.y + panel.height * 0.12,
                            width: panel.width * 0.84,
                            height: panel.height * 0.76,
                        },
                        text: Some(text.clone()),
                        visual: Visual::TextLayout {
                            layout: TextLayout {
                                vertical_align: TextVerticalAlign::Center,
                                wrap: false,
                                ..TextLayout::default()
                            },
                            visual: Box::new(Visual::RichText {
                                geometry: Geometry::Rectangle,
                                fill: Paint::None,
                                stroke: Paint::None,
                                stroke_width: 0.0,
                                align: TextAlign::Center,
                                line_height: 0.0,
                                runs: vec![TextRun {
                                    paint: None,
                                    east_asian_line_breaks: true,
                                    text,
                                    font_family: "Arial".to_owned(),
                                    font_size: (panel.height * 0.16).clamp(12.0, 72.0),
                                    color: 0xffff_ffff,
                                    bold: false,
                                    italic: false,
                                    underline: false,
                                    strikethrough: false,
                                    highlight: 0,
                                    baseline_shift: 0.0,
                                    letter_spacing: 0.0,
                                    horizontal_scale: 1.0,
                                }],
                            }),
                        },
                    });
                }
            }
            elements
        })
}

pub(super) fn diagram_horizontal_list_elements(
    diagram: &Diagram,
    bounds: Rect,
    accent: u32,
) -> Option<Vec<DrawingMlElement>> {
    if diagram.layout_type.as_deref()?.split('#').next()?
        == "urn:microsoft.com/office/officeart/2005/8/layout/hList1"
    {
        return diagram_column_list_elements(diagram, bounds, accent);
    }
    diagram
        .layout_type
        .as_deref()
        .is_some_and(|value| value.contains("/hList7"))
        .then(|| {
            let count = diagram.nodes.len().max(1);
            let gap = (bounds.width / 80.0).clamp(4.0, 12.0);
            let panel_width = ((bounds.width - gap * (count + 1) as f32) / count as f32).max(1.0);
            let mut elements = Vec::with_capacity(count.saturating_mul(3).saturating_add(1));
            for (index, node) in diagram.nodes.iter().enumerate() {
                let panel = Rect {
                    x: bounds.x + gap + index as f32 * (panel_width + gap),
                    y: bounds.y,
                    width: panel_width,
                    height: bounds.height,
                };
                let diameter = panel.width.min(panel.height) * 0.60;
                let picture = Rect {
                    x: panel.x + (panel.width - diameter) / 2.0,
                    y: panel.y + panel.height * 0.06,
                    width: diameter,
                    height: diameter,
                };
                let text = if node.text.trim().is_empty() && node.placeholder {
                    "[Text]".to_owned()
                } else {
                    node.text.clone()
                };
                elements.push(DrawingMlElement {
                    bounds: panel,
                    text: None,
                    visual: Visual::PaintedShape {
                        geometry: Geometry::RoundedRectangle {
                            radius_x: panel.width.min(panel.height) * 0.10,
                            radius_y: panel.width.min(panel.height) * 0.10,
                        },
                        fill: Paint::Solid(accent),
                        stroke: Paint::Solid(0xffff_ffff),
                        stroke_width: 3.0,
                    },
                });
                elements.push(DrawingMlElement {
                    bounds: picture,
                    text: None,
                    visual: Visual::PaintedShape {
                        geometry: Geometry::Ellipse,
                        fill: node
                            .fill
                            .as_ref()
                            .map_or(Paint::Solid(0xffff_ffff), |fill| fill.paint(picture)),
                        stroke: Paint::Solid(0xffff_ffff),
                        stroke_width: 4.0,
                    },
                });
                let text_bounds = Rect {
                    x: panel.x + panel.width * 0.04,
                    y: panel.y + panel.height * 0.46,
                    width: panel.width * 0.92,
                    height: panel.height * 0.28,
                };
                elements.push(DrawingMlElement {
                    bounds: text_bounds,
                    text: Some(text.clone()),
                    visual: Visual::TextLayout {
                        layout: TextLayout {
                            vertical_align: TextVerticalAlign::Center,
                            wrap: false,
                            ..TextLayout::default()
                        },
                        visual: Box::new(Visual::RichText {
                            geometry: Geometry::Rectangle,
                            fill: Paint::None,
                            stroke: Paint::None,
                            stroke_width: 0.0,
                            align: TextAlign::Center,
                            line_height: 0.0,
                            runs: vec![TextRun {
                                paint: None,
                                east_asian_line_breaks: true,
                                text,
                                font_family: "Arial".to_owned(),
                                font_size: (panel.height * 0.12).clamp(12.0, 72.0),
                                color: 0xffff_ffff,
                                bold: false,
                                italic: false,
                                underline: false,
                                strikethrough: false,
                                highlight: 0,
                                baseline_shift: 0.0,
                                letter_spacing: 0.0,
                                horizontal_scale: 1.0,
                            }],
                        }),
                    },
                });
            }
            let arrow = Rect {
                x: bounds.x + bounds.width * 0.04,
                y: bounds.y + bounds.height * 0.81,
                width: bounds.width * 0.92,
                height: bounds.height * 0.12,
            };
            let head = arrow.height;
            let shaft_top = arrow.height * 0.35;
            let shaft_bottom = arrow.height * 0.65;
            elements.push(DrawingMlElement {
                bounds: arrow,
                text: None,
                visual: Visual::PaintedShape {
                    geometry: Geometry::Path {
                        fill_rule: FillRule::NonZero,
                        commands: vec![
                            PathCommand::MoveTo {
                                x: 0.0,
                                y: arrow.height / 2.0,
                            },
                            PathCommand::LineTo { x: head, y: 0.0 },
                            PathCommand::LineTo {
                                x: head,
                                y: shaft_top,
                            },
                            PathCommand::LineTo {
                                x: arrow.width - head,
                                y: shaft_top,
                            },
                            PathCommand::LineTo {
                                x: arrow.width - head,
                                y: 0.0,
                            },
                            PathCommand::LineTo {
                                x: arrow.width,
                                y: arrow.height / 2.0,
                            },
                            PathCommand::LineTo {
                                x: arrow.width - head,
                                y: arrow.height,
                            },
                            PathCommand::LineTo {
                                x: arrow.width - head,
                                y: shaft_bottom,
                            },
                            PathCommand::LineTo {
                                x: head,
                                y: shaft_bottom,
                            },
                            PathCommand::LineTo {
                                x: head,
                                y: arrow.height,
                            },
                            PathCommand::ClosePath,
                        ],
                    },
                    fill: Paint::Solid(0xb4c7_e7ff),
                    stroke: Paint::Solid(0xffff_ffff),
                    stroke_width: 3.0,
                },
            });
            elements
        })
}

// hList1 has separate parent/descendant rectangles; hList7 has picture circles
// and an arrow, so only their existing host dispatch is shared.
fn diagram_column_list_elements(
    diagram: &Diagram,
    bounds: Rect,
    accent: u32,
) -> Option<Vec<DrawingMlElement>> {
    let mut branches = diagram_branches(diagram)?;
    if branches.is_empty() {
        return None;
    }
    if diagram.right_to_left {
        branches.reverse();
    }
    let nodes = branches.iter().map(|(node, _)| *node).collect::<Vec<_>>();
    let nested = branches.iter().any(|(_, children)| !children.is_empty());
    let width = bounds.width / (nodes.len() as f32 + 0.14 * (nodes.len() - 1) as f32);
    let header_height = (width * 0.4).min(if nested {
        bounds.height * 0.8 / 2.02
    } else {
        bounds.height
    });
    // Built-in empty-text constraints: parTx top/bottom margins are 0.32 em;
    // desTx height is 1.22 em. Clamp the header, then centre the resulting pair.
    let body_height = (bounds.height * 1.22 / (if nested { 0.8 } else { 0.64 } + 1.22))
        .min(bounds.height - header_height);
    let top = bounds.y + (bounds.height - header_height - body_height) / 2.0;
    let mut elements = Vec::new();
    for (index, node) in nodes.into_iter().enumerate() {
        for (role, y, height) in [
            ("desTx", top + header_height, body_height),
            ("parTx", top, header_height),
        ] {
            let rect = Rect {
                x: bounds.x + index as f32 * width * 1.14,
                y,
                width,
                height,
            };
            let style = diagram
                .role_fills
                .get(&(node.model_id.clone(), role.to_owned()));
            let fill = if role == "parTx" {
                node.fill.as_ref()
            } else {
                None
            }
            .or_else(|| style.and_then(|style| style.fill.as_ref()));
            let paint = if node.no_fill {
                Paint::None
            } else {
                fill.map_or(Paint::Solid(accent), |fill| fill.paint(rect))
            };
            let text = if role == "parTx" {
                Some(node.text.clone())
            } else {
                let text = diagram_bullet_text(&branches[index].1);
                (!text.is_empty()).then_some(text)
            };
            let visual = if let Some(text) = &text {
                Visual::TextLayout {
                    layout: TextLayout {
                        vertical_align: if role == "parTx" {
                            TextVerticalAlign::Center
                        } else {
                            TextVerticalAlign::Top
                        },
                        inset_left: if role == "parTx" {
                            0.0
                        } else {
                            header_height * 0.16
                        },
                        inset_top: if role == "parTx" {
                            0.0
                        } else {
                            header_height * 0.12
                        },
                        auto_fit: crate::model::TextAutoFit::Shrink,
                        ..TextLayout::default()
                    },
                    visual: Box::new(Visual::RichText {
                        geometry: Geometry::Rectangle,
                        fill: paint,
                        stroke: Paint::None,
                        stroke_width: 0.0,
                        align: if role == "parTx" {
                            TextAlign::Center
                        } else {
                            TextAlign::Start
                        },
                        line_height: 0.0,
                        runs: vec![TextRun {
                            text: text.clone(),
                            font_family: "Calibri".to_owned(),
                            font_size: node.font_size.unwrap_or(if nested {
                                header_height / 2.4
                            } else {
                                24.0
                            }),
                            color: if role == "parTx" {
                                0xffff_ffff
                            } else {
                                0x0000_00ff
                            },
                            bold: node.bold,
                            italic: false,
                            underline: false,
                            strikethrough: false,
                            highlight: 0,
                            baseline_shift: 0.0,
                            letter_spacing: 0.0,
                            horizontal_scale: 1.0,
                            east_asian_line_breaks: true,
                            paint: None,
                        }],
                    }),
                }
            } else {
                Visual::PaintedShape {
                    geometry: Geometry::Rectangle,
                    fill: paint,
                    stroke: Paint::None,
                    stroke_width: 0.0,
                }
            };
            let shadow = style
                .and_then(|style| style.shadow)
                .or_else(|| (role == "parTx").then_some(node.shadow).flatten());
            let visual = if shadow.is_some() {
                Visual::Effect {
                    shadow,
                    clip: None,
                    visual: Box::new(visual),
                }
            } else {
                visual
            };
            elements.push(DrawingMlElement {
                bounds: rect,
                text,
                visual,
            });
        }
    }
    Some(elements)
}

fn diagram_branches(diagram: &Diagram) -> Option<Vec<(&DiagramNode, Vec<(&DiagramNode, usize)>)>> {
    let ids = diagram
        .nodes
        .iter()
        .map(|node| node.model_id.as_str())
        .collect::<HashSet<_>>();
    let mut children = HashMap::<&str, Vec<&DiagramNode>>::new();
    let mut roots = Vec::new();
    for node in &diagram.nodes {
        match node
            .parent_id
            .as_deref()
            .filter(|parent| ids.contains(parent))
        {
            Some(parent) => children.entry(parent).or_default().push(node),
            None => roots.push(node),
        }
    }
    roots.sort_by_key(|node| node.sibling_order);
    for nodes in children.values_mut() {
        nodes.sort_by_key(|node| node.sibling_order);
    }
    let mut visited = HashSet::new();
    let mut result = Vec::new();
    for root in roots {
        if !visited.insert(root.model_id.as_str()) {
            return None;
        }
        let mut descendants = Vec::new();
        let mut stack = children
            .get(root.model_id.as_str())
            .into_iter()
            .flatten()
            .rev()
            .map(|node| (*node, 0_usize))
            .collect::<Vec<_>>();
        while let Some((node, depth)) = stack.pop() {
            if !visited.insert(node.model_id.as_str()) {
                return None;
            }
            descendants.push((node, depth));
            if let Some(next) = children.get(node.model_id.as_str()) {
                stack.extend(next.iter().rev().map(|node| (*node, depth + 1)));
            }
        }
        result.push((root, descendants));
    }
    (visited.len() == diagram.nodes.len()).then_some(result)
}

fn diagram_bullet_text(nodes: &[(&DiagramNode, usize)]) -> String {
    nodes
        .iter()
        .map(|(node, depth)| format!("{}• {}", "  ".repeat((*depth).min(8)), node.text))
        .collect::<Vec<_>>()
        .join("\n")
}

// ponytail: without generated DrawingML, these family layouts preserve topology,
// text and paints; an Office layout-constraint interpreter is needed for exact
// geometry of every authored layout variation. Host adapters retain diagnostics.
pub(super) fn diagram_semantic_elements(
    diagram: &Diagram,
    bounds: Rect,
    accent: u32,
) -> Option<Vec<DrawingMlElement>> {
    let layout = diagram
        .layout_type
        .as_deref()?
        .rsplit('/')
        .next()?
        .split('#')
        .next()?;
    let circular = matches!(
        layout,
        "cycle2" | "cycle3" | "cycle7" | "radial3" | "relationship"
    );
    let pyramid = layout == "pyramid1";
    let venn = matches!(layout, "venn1" | "venn2" | "target2");
    let gear = layout == "gear1";
    let matrix = matches!(layout, "matrix1" | "cycle4" | "HexagonCluster");
    let chevron = matches!(layout, "chevron1" | "chevron2");
    let vertical = matches!(
        layout,
        "list1"
            | "vList3"
            | "vList4"
            | "vList5"
            | "vProcess5"
            | "VerticalCurvedList"
            | "BracketList+Icon"
    );
    let horizontal = matches!(
        layout,
        "hList3"
            | "hProcess3"
            | "hProcess6"
            | "hProcess9"
            | "lProcess2"
            | "bProcess2"
            | "process3"
            | "process4"
            | "arrow5"
            | "pList2"
            | "PictureStrips"
            | "bList2"
            | "funnel1"
            | "chart3"
            | "hierarchy2"
            | "hierarchy4"
    );
    if !(circular || pyramid || venn || gear || matrix || chevron || vertical || horizontal) {
        return None;
    }
    let mut branches = diagram_branches(diagram)?;
    if branches.is_empty() {
        return None;
    }
    if diagram.right_to_left {
        branches.reverse();
    }
    let count = branches.len();
    let mut elements = Vec::new();
    let mut panels = Vec::new();
    if circular {
        // The cycle algorithm's 0.5-node sibling spacing determines its radius.
        // Fit the union of node circles, not an assumed symmetric circle box.
        let radius = if count > 1 {
            0.75 / (std::f32::consts::PI / count as f32).sin()
        } else {
            0.0
        };
        let centers = (0..count)
            .map(|i| {
                let angle =
                    -std::f32::consts::FRAC_PI_2 + std::f32::consts::TAU * i as f32 / count as f32;
                (radius * angle.cos(), radius * angle.sin())
            })
            .collect::<Vec<_>>();
        let left = centers
            .iter()
            .map(|p| p.0 - 0.5)
            .fold(f32::INFINITY, f32::min);
        let right = centers
            .iter()
            .map(|p| p.0 + 0.5)
            .fold(f32::NEG_INFINITY, f32::max);
        let top = centers
            .iter()
            .map(|p| p.1 - 0.5)
            .fold(f32::INFINITY, f32::min);
        let bottom = centers
            .iter()
            .map(|p| p.1 + 0.5)
            .fold(f32::NEG_INFINITY, f32::max);
        let scale = (bounds.width / (right - left)).min(bounds.height / (bottom - top));
        let x = bounds.x + (bounds.width - (right - left) * scale) / 2.0;
        let y = bounds.y + (bounds.height - (bottom - top) * scale) / 2.0;
        for center in centers {
            panels.push(Rect {
                x: x + (center.0 - 0.5 - left) * scale,
                y: y + (center.1 - 0.5 - top) * scale,
                width: scale,
                height: scale,
            });
        }
        if count > 1 {
            for i in 0..count {
                let (a, b) = (panels[i], panels[(i + 1) % count]);
                let (dx, dy) = (b.x - a.x, b.y - a.y);
                let distance = dx.hypot(dy);
                if distance <= 0.0 {
                    continue;
                }
                let rect = Rect {
                    x: (a.x + b.x + a.width) / 2.0 - scale * 0.125,
                    y: (a.y + b.y + a.height) / 2.0 - scale * 0.125,
                    width: scale * 0.25,
                    height: scale * 0.25,
                };
                let geometry =
                    super::pptx::drawingml_preset_geometry("rightArrow", rect, &HashMap::new())?.0;
                let mut element =
                    extended_chart_shape(rect, geometry, Paint::Solid(0xb0c3_e1ff), Paint::None);
                let angle = dy.atan2(dx);
                let (cx, cy) = (rect.x + rect.width / 2.0, rect.y + rect.height / 2.0);
                element.visual = Visual::Layer {
                    transform: AffineTransform {
                        a: angle.cos(),
                        b: angle.sin(),
                        c: -angle.sin(),
                        d: angle.cos(),
                        e: cx - cx * angle.cos() + cy * angle.sin(),
                        f: cy - cx * angle.sin() - cy * angle.cos(),
                    },
                    opacity: 1.0,
                    blend_mode: crate::model::BlendMode::Normal,
                    visual: Box::new(element.visual),
                };
                elements.push(element);
            }
        }
    } else if venn || gear {
        let size = bounds.width.min(bounds.height) * if count <= 2 { 0.72 } else { 0.58 };
        let radius = size * if gear { 0.56 } else { 0.27 };
        for i in 0..count {
            let angle =
                -std::f32::consts::FRAC_PI_2 + std::f32::consts::TAU * i as f32 / count as f32;
            panels.push(Rect {
                x: bounds.x + bounds.width / 2.0 + radius * angle.cos() - size / 2.0,
                y: bounds.y + bounds.height / 2.0 + radius * angle.sin() - size / 2.0,
                width: size,
                height: size,
            });
        }
    } else if matrix {
        let columns = (count as f32).sqrt().ceil() as usize;
        let rows = count.div_ceil(columns);
        let gap = bounds.width.min(bounds.height) * 0.02;
        let width = (bounds.width - gap * (columns - 1) as f32) / columns as f32;
        let height = (bounds.height - gap * (rows - 1) as f32) / rows as f32;
        for i in 0..count {
            panels.push(Rect {
                x: bounds.x + (i % columns) as f32 * (width + gap),
                y: bounds.y + (i / columns) as f32 * (height + gap),
                width,
                height,
            });
        }
    } else if pyramid {
        for i in 0..count {
            panels.push(Rect {
                x: bounds.x,
                y: bounds.y + i as f32 * bounds.height / count as f32,
                width: bounds.width,
                height: bounds.height / count as f32,
            });
        }
    } else {
        let gap = if chevron {
            -8.0
        } else {
            (bounds.width.min(bounds.height) * 0.04).min(
                if vertical {
                    bounds.height
                } else {
                    bounds.width
                } / count.max(1) as f32
                    * 0.25,
            )
        };
        let width = if vertical {
            bounds.width
        } else {
            (bounds.width - gap * (count - 1) as f32) / count as f32
        };
        let height = if vertical {
            (bounds.height - gap * (count - 1) as f32) / count as f32
        } else if chevron {
            (width * 0.4).min(bounds.height)
        } else {
            bounds.height
        };
        for i in 0..count {
            panels.push(Rect {
                x: bounds.x
                    + if vertical {
                        0.0
                    } else {
                        i as f32 * (width + gap)
                    },
                y: bounds.y
                    + if vertical {
                        i as f32 * (height + gap)
                    } else {
                        (bounds.height - height) / 2.0
                    },
                width,
                height,
            });
        }
    }
    for (i, ((node, children), rect)) in branches.iter().zip(&panels).enumerate() {
        let preset = if circular || venn {
            "ellipse"
        } else if gear {
            "gear6"
        } else if chevron {
            "chevron"
        } else if layout == "HexagonCluster" {
            "hexagon"
        } else if matches!(layout, "arrow5" | "hProcess6") {
            "rightArrow"
        } else {
            "rect"
        };
        let geometry = if pyramid {
            let a = i as f32 / count as f32 * rect.width / 2.0;
            let b = (i + 1) as f32 / count as f32 * rect.width / 2.0;
            extended_chart_path(
                &[
                    (rect.width / 2.0 - a, 0.0),
                    (rect.width / 2.0 + a, 0.0),
                    (rect.width / 2.0 + b, rect.height),
                    (rect.width / 2.0 - b, rect.height),
                ],
                true,
            )
        } else {
            super::pptx::drawingml_preset_geometry(
                node.preset_geometry.as_deref().unwrap_or(preset),
                *rect,
                &HashMap::new(),
            )?
            .0
        };
        let role_fill = ["node", "parTxOnly", "parTx", "node1", "tx"]
            .iter()
            .find_map(|role| {
                diagram
                    .role_fills
                    .get(&(node.model_id.clone(), (*role).to_owned()))
                    .and_then(|style| style.fill.as_ref())
            });
        let fill = node
            .fill
            .as_ref()
            .or(role_fill)
            .map_or(Paint::Solid(accent), |fill| fill.paint(*rect));
        let mut shape = extended_chart_shape(
            *rect,
            geometry,
            if node.no_fill { Paint::None } else { fill },
            Paint::Solid(0xffff_ffff),
        );
        if node.shadow.is_some() {
            shape.visual = Visual::Effect {
                shadow: node.shadow,
                clip: None,
                visual: Box::new(shape.visual),
            };
        }
        elements.push(shape);
        let mut text = node.text.clone();
        if !children.is_empty() {
            text.push('\n');
            text.push_str(&diagram_bullet_text(children));
        }
        if text.is_empty() && node.placeholder {
            text = "[Text]".to_owned();
        }
        if !text.is_empty() {
            let font = node.font_size.unwrap_or_else(|| {
                (rect.height
                    * if circular {
                        0.48
                    } else if chevron {
                        0.7
                    } else {
                        0.18
                    })
                .min(86.6667)
            });
            let inset = rect.width.min(rect.height) * 0.12;
            elements.push(extended_chart_text(
                Rect {
                    x: rect.x + inset,
                    y: rect.y + inset,
                    width: rect.width - 2.0 * inset,
                    height: rect.height - 2.0 * inset,
                },
                text,
                font,
                0xffff_ffff,
            ));
        }
    }
    Some(elements)
}
