//! Minimal, bounded XFA layout for pure-XFA PDFs.
//!
//! XFA is an XML form language layered on top of PDF.  This adapter keeps the
//! XML as the source of truth and synthesizes a vector page for the normal
//! renderer; it never uses the PDF fallback page or an embedded preview.

use std::collections::BTreeMap;

use crate::diagnostic::{Diagnostic, DiagnosticCode, Fidelity, Phase};
use crate::limits::Limits;
use crate::model::{
    Document, ImageCrop, MappingQuality, Object, ObjectKind, Rect, SheetAxis, SourceLocator,
    SourceRef, Unit, UnitKind, Visual,
};
use crate::xml::{XmlEvent, decode_xml_text, parse_xml};

use super::super::local_name;

const MM_TO_CSS: f32 = 96.0 / 25.4;
const PT_TO_CSS: f32 = 96.0 / 72.0;

#[derive(Clone, Debug, Default)]
struct Node {
    name: String,
    attributes: BTreeMap<String, String>,
    children: Vec<usize>,
    text: String,
}

#[derive(Clone, Debug)]
pub(super) struct XfaPage {
    pub(super) width: f32,
    pub(super) height: f32,
    pub(super) svg: Vec<u8>,
    pub(super) text: String,
}

#[derive(Clone, Copy)]
struct PageGeometry {
    width: f32,
    height: f32,
    content_x: f32,
    content_y: f32,
    content_height: f32,
}

struct Layout<'a> {
    nodes: &'a [Node],
    page_templates: Vec<usize>,
    geometry: PageGeometry,
    pages: Vec<String>,
    page_text: Vec<String>,
    values: BTreeMap<String, Vec<String>>,
    value_indices: BTreeMap<String, usize>,
    page: usize,
    cursor_y: f32,
    max_pages: usize,
}

pub(super) fn render(
    template: &[u8],
    datasets: &[u8],
    form: &[u8],
    limits: Limits,
) -> Result<Vec<XfaPage>, Diagnostic> {
    let nodes = parse_nodes(template, limits)?;
    let Some(root) = nodes.first().map(|_| 0) else {
        return Ok(Vec::new());
    };
    let Some(form_root) =
        descendants(&nodes, root).find(|index| local_name(&nodes[*index].name) == "subform")
    else {
        return Ok(Vec::new());
    };
    let page_set = nodes[form_root]
        .children
        .iter()
        .copied()
        .find(|index| local_name(&nodes[*index].name) == "pageSet");
    let page_templates = page_set
        .map(|index| {
            nodes[index]
                .children
                .iter()
                .copied()
                .filter(|child| local_name(&nodes[*child].name) == "pageArea")
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let Some(first_page) = page_templates.first().copied() else {
        return Ok(Vec::new());
    };
    let geometry = page_geometry(&nodes, first_page);
    let max_pages = saved_page_count(form, limits).max(1);
    let values = dataset_values(datasets, limits)?;
    let mut layout = Layout {
        nodes: &nodes,
        page_templates,
        geometry,
        pages: Vec::new(),
        page_text: Vec::new(),
        values,
        value_indices: BTreeMap::new(),
        page: 0,
        cursor_y: geometry.content_y,
        max_pages,
    };
    layout.ensure_page(0);
    let dynamic_children = nodes[form_root].children.clone();
    for child in dynamic_children {
        if Some(child) == page_set || !visible(&nodes[child]) {
            continue;
        }
        layout.flow_node(child, geometry.content_x, 0.0)?;
    }
    while layout.pages.len() < max_pages {
        layout.ensure_page(layout.pages.len());
    }
    Ok(layout
        .pages
        .into_iter()
        .zip(layout.page_text)
        .map(|(body, text)| XfaPage {
            width: geometry.width,
            height: geometry.height,
            svg: format!(
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{}\" height=\"{}\" viewBox=\"0 0 {} {}\"><rect width=\"100%\" height=\"100%\" fill=\"white\"/>{body}</svg>",
                geometry.width, geometry.height, geometry.width, geometry.height
            )
            .into_bytes(),
            text,
        })
        .collect())
}

pub(super) fn materialize(
    pages: &[XfaPage],
    selected_unit: Option<usize>,
    document: &mut Document,
) -> Result<(), Diagnostic> {
    for (index, page) in pages.iter().enumerate() {
        let unit_index = u32::try_from(index).map_err(|_| object_limit())?;
        document.units.push(Unit {
            kind: UnitKind::Page,
            index: unit_index,
            id: format!("pdf:xfa-page:{}", index + 1),
            name: format!("Page {}", index + 1),
            width: page.width,
            height: page.height,
            rows: 0,
            columns: 0,
            frozen_rows: 0,
            frozen_columns: 0,
            frozen_width: 0.0,
            frozen_height: 0.0,
            row_axis: SheetAxis::default(),
            column_axis: SheetAxis::default(),
            show_grid_lines: false,
            tab_color: None,
            sheet: None,
            slide: None,
        });
        if selected_unit.is_some_and(|selected| selected != index) {
            continue;
        }
        document.objects.push(Object {
            numeric_id: 0,
            parent_numeric_id: None,
            stable_id: format!("pdf:{unit_index}:0"),
            parent_stable_id: None,
            kind: ObjectKind::Image,
            unit_index,
            bounds: Rect {
                x: 0.0,
                y: 0.0,
                width: page.width,
                height: page.height,
            },
            z: 0,
            text: (!page.text.is_empty()).then(|| page.text.clone()),
            source: SourceRef {
                part: "document.pdf".to_owned(),
                mapping: MappingQuality::Derived,
                locator: SourceLocator::Pdf {
                    kind: "form",
                    object_number: None,
                    byte_offset: None,
                    action: None,
                },
            },
            visual: Visual::Image {
                media_type: "image/svg+xml".to_owned(),
                bytes: page.svg.clone(),
                crop: ImageCrop::default(),
            },
        });
    }
    document.diagnostics.push(
        Diagnostic::warning(
            DiagnosticCode::ApproximateLayout,
            Phase::Layout,
            Fidelity::Approximate,
            "PDF_XFA_LAYOUT: the pure-XFA form was rendered from its XML template and datasets",
        )
        .in_part("document.pdf"),
    );
    Ok(())
}

impl Layout<'_> {
    fn ensure_page(&mut self, page: usize) {
        while self.pages.len() <= page && self.pages.len() < self.max_pages {
            let index = self.pages.len();
            self.pages.push(String::new());
            self.page_text.push(String::new());
            let template = if index == 0 {
                self.page_templates.first().copied()
            } else {
                self.page_templates
                    .get(1)
                    .or_else(|| self.page_templates.first())
                    .copied()
            };
            if let Some(template) = template {
                self.render_positioned(template, 0.0, 0.0, index);
            }
        }
    }

    fn next_page(&mut self) {
        if self.page + 1 < self.max_pages {
            self.page += 1;
            self.ensure_page(self.page);
            self.cursor_y = self.geometry.content_y;
        }
    }

    fn flow_node(&mut self, index: usize, x: f32, y_offset: f32) -> Result<(), Diagnostic> {
        let node = &self.nodes[index];
        if !visible(node) {
            return Ok(());
        }
        if direct_child(node, self.nodes, "breakBefore").is_some() {
            self.next_page();
        }
        let layout = node
            .attributes
            .get("layout")
            .map(String::as_str)
            .unwrap_or("position");
        let height = node_height(self.nodes, index);
        if layout == "tb" {
            let start_y = self.cursor_y;
            let start_x = x + length_attr(node, "x").unwrap_or(0.0);
            for child in node.children.clone() {
                if local_name(&self.nodes[child].name) == "pageSet" || !visible(&self.nodes[child])
                {
                    continue;
                }
                self.flow_node(child, start_x, y_offset)?;
            }
            if let Some(height) = length_attr(node, "h") {
                self.cursor_y = self.cursor_y.max(start_y + height);
            }
            return Ok(());
        }
        if matches!(layout, "lr-tb" | "rl-tb") {
            let available_width = length_attr(node, "w")
                .unwrap_or_else(|| node_width(self.nodes, index))
                .max(0.1);
            let start_x = x + length_attr(node, "x").unwrap_or(0.0);
            let right_to_left = layout == "rl-tb";
            let children = node
                .children
                .iter()
                .copied()
                .filter(|child| {
                    renderable_child(&self.nodes[*child]) && visible(&self.nodes[*child])
                })
                .collect::<Vec<_>>();
            let mut row = Vec::<(usize, f32)>::new();
            let mut row_width = 0.0;
            let mut row_height = 0.0_f32;
            for child in children {
                let child_width = node_width(self.nodes, child).min(available_width).max(0.1);
                if !row.is_empty() && row_width + child_width > available_width + 0.01 {
                    self.render_flow_row(
                        &row,
                        start_x,
                        available_width,
                        row_height,
                        y_offset,
                        right_to_left,
                    );
                    row.clear();
                    row_width = 0.0;
                    row_height = 0.0;
                }
                row.push((child, row_width));
                row_width += child_width;
                row_height = row_height.max(node_height(self.nodes, child));
            }
            if !row.is_empty() {
                self.render_flow_row(
                    &row,
                    start_x,
                    available_width,
                    row_height,
                    y_offset,
                    right_to_left,
                );
            }
            return Ok(());
        }
        if self.cursor_y + y_offset + height
            > self.geometry.content_y + self.geometry.content_height
            && self.cursor_y > self.geometry.content_y
        {
            self.next_page();
        }
        // In a top-to-bottom flowed container the layout engine owns the
        // child's origin. Authored x/y coordinates belong to positioned
        // containers and must not be added again here.
        let authored_x = length_attr(node, "x").unwrap_or(0.0);
        let authored_y = length_attr(node, "y").unwrap_or(0.0);
        self.render_positioned(
            index,
            x - authored_x,
            self.cursor_y + y_offset - authored_y,
            self.page,
        );
        self.cursor_y += height.max(0.0);
        Ok(())
    }

    fn render_flow_row(
        &mut self,
        row: &[(usize, f32)],
        start_x: f32,
        available_width: f32,
        row_height: f32,
        y_offset: f32,
        right_to_left: bool,
    ) {
        if self.cursor_y + y_offset + row_height
            > self.geometry.content_y + self.geometry.content_height
            && self.cursor_y > self.geometry.content_y
        {
            self.next_page();
        }
        for &(child, offset_x) in row {
            let child_width = node_width(self.nodes, child).min(available_width).max(0.1);
            let authored_x = length_attr(&self.nodes[child], "x").unwrap_or(0.0);
            let authored_y = length_attr(&self.nodes[child], "y").unwrap_or(0.0);
            let child_x = if right_to_left {
                available_width - offset_x - child_width
            } else {
                offset_x
            };
            self.render_positioned(
                child,
                start_x + child_x - authored_x,
                self.cursor_y + y_offset - authored_y,
                self.page,
            );
        }
        self.cursor_y += row_height.max(0.0);
    }

    fn render_positioned(&mut self, index: usize, parent_x: f32, parent_y: f32, page: usize) {
        if page >= self.pages.len() || !visible(&self.nodes[index]) {
            return;
        }
        let node = &self.nodes[index];
        let x = parent_x + length_attr(node, "x").unwrap_or(0.0);
        let y = parent_y + length_attr(node, "y").unwrap_or(0.0);
        match local_name(&node.name) {
            "draw" | "field" | "exclGroup" => self.render_primitive(index, x, y, page),
            "line" => self.render_line(index, x, y, page),
            "rectangle" | "arc" => self.render_box(index, x, y, page),
            _ => {}
        }
        let children = node
            .children
            .iter()
            .copied()
            .filter(|child| renderable_child(&self.nodes[*child]) && visible(&self.nodes[*child]))
            .collect::<Vec<_>>();
        match node.attributes.get("layout").map(String::as_str) {
            Some("tb") => {
                let mut offset_y = 0.0;
                for child in children {
                    let authored_x = length_attr(&self.nodes[child], "x").unwrap_or(0.0);
                    let authored_y = length_attr(&self.nodes[child], "y").unwrap_or(0.0);
                    self.render_positioned(child, x - authored_x, y + offset_y - authored_y, page);
                    offset_y += node_height(self.nodes, child);
                }
            }
            Some("lr-tb" | "rl-tb") => {
                let right_to_left =
                    node.attributes.get("layout").map(String::as_str) == Some("rl-tb");
                let available_width = length_attr(node, "w")
                    .unwrap_or_else(|| node_width(self.nodes, index))
                    .max(0.1);
                let mut row_x = 0.0;
                let mut row_y = 0.0;
                let mut row_height = 0.0_f32;
                for child in children {
                    let child_width = node_width(self.nodes, child).min(available_width).max(0.1);
                    let child_height = node_height(self.nodes, child);
                    if row_x > 0.0 && row_x + child_width > available_width + 0.01 {
                        row_y += row_height;
                        row_x = 0.0;
                        row_height = 0.0;
                    }
                    let authored_x = length_attr(&self.nodes[child], "x").unwrap_or(0.0);
                    let authored_y = length_attr(&self.nodes[child], "y").unwrap_or(0.0);
                    let child_x = if right_to_left {
                        available_width - row_x - child_width
                    } else {
                        row_x
                    };
                    self.render_positioned(
                        child,
                        x + child_x - authored_x,
                        y + row_y - authored_y,
                        page,
                    );
                    row_x += child_width;
                    row_height = row_height.max(child_height);
                }
            }
            _ => {
                for child in children {
                    self.render_positioned(child, x, y, page);
                }
            }
        }
    }

    fn render_primitive(&mut self, index: usize, x: f32, y: f32, page: usize) {
        let node = &self.nodes[index];
        let width = length_attr(node, "w").unwrap_or(80.0).max(0.1);
        let height = length_attr(node, "h").unwrap_or(16.0).max(0.1);
        let is_field = matches!(local_name(&node.name), "field" | "exclGroup");
        let caption = descendant(node, self.nodes, "caption");
        let caption_text = caption
            .map(|value| textual_value(value, self.nodes))
            .unwrap_or_default();
        let caption_reserve = caption
            .and_then(|caption| length_attr(caption, "reserve"))
            .unwrap_or(0.0)
            .max(0.0);
        let caption_placement = caption
            .and_then(|caption| caption.attributes.get("placement"))
            .map(String::as_str)
            .unwrap_or("left");
        let (mut field_x, mut field_y, mut field_width, mut field_height) =
            if is_field && !caption_text.is_empty() && caption_reserve > 0.0 {
                match caption_placement {
                    "top" => (
                        x,
                        y + caption_reserve,
                        width,
                        (height - caption_reserve).max(0.1),
                    ),
                    "right" => (x, y, (width - caption_reserve).max(0.1), height),
                    "bottom" => (x, y, width, (height - caption_reserve).max(0.1)),
                    _ => (
                        x + caption_reserve,
                        y,
                        (width - caption_reserve).max(0.1),
                        height,
                    ),
                }
            } else {
                (x, y, width, height)
            };
        let margin = direct_child(node, self.nodes, "margin");
        let inset_left = margin
            .and_then(|margin| length_attr(margin, "leftInset"))
            .unwrap_or(0.0);
        let inset_right = margin
            .and_then(|margin| length_attr(margin, "rightInset"))
            .unwrap_or(0.0);
        let inset_top = margin
            .and_then(|margin| length_attr(margin, "topInset"))
            .unwrap_or(0.0);
        let inset_bottom = margin
            .and_then(|margin| length_attr(margin, "bottomInset"))
            .unwrap_or(0.0);
        if is_field {
            field_x += inset_left;
            field_y += inset_top;
            field_width = (field_width - inset_left - inset_right).max(0.1);
            field_height = (field_height - inset_top - inset_bottom).max(0.1);
        }
        let fill = descendant(node, self.nodes, "color")
            .and_then(|color| color.attributes.get("value"))
            .and_then(|value| rgb(value))
            .unwrap_or_else(|| {
                if is_field {
                    "white".to_owned()
                } else {
                    "none".to_owned()
                }
            });
        let bordered = is_field || descendant(node, self.nodes, "border").is_some();
        if bordered || fill != "none" {
            self.pages[page].push_str(&format!(
                "<rect x=\"{field_x}\" y=\"{field_y}\" width=\"{field_width}\" height=\"{field_height}\" fill=\"{fill}\" stroke=\"{}\" stroke-width=\"{}\"/>",
                if bordered { "black" } else { "none" },
                if bordered { 0.7 } else { 0.0 }
            ));
        }
        if let Some(image) = descendant(node, self.nodes, "image") {
            let media_type = image
                .attributes
                .get("contentType")
                .map(String::as_str)
                .unwrap_or("image/png");
            let data = image.text.split_whitespace().collect::<String>();
            if !data.is_empty() {
                self.pages[page].push_str(&format!(
                    "<image x=\"{field_x}\" y=\"{field_y}\" width=\"{field_width}\" height=\"{field_height}\" preserveAspectRatio=\"xMidYMid meet\" href=\"data:{};base64,{}\"/>",
                    escape(media_type), data
                ));
            }
        }
        let mut text = if is_field {
            self.field_value(index)
        } else {
            descendant(node, self.nodes, "value")
                .map(|value| textual_value(value, self.nodes))
                .unwrap_or_default()
        };
        if text.is_empty() && !is_field && !caption_text.is_empty() {
            text = caption_text.clone();
        }
        if !caption_text.is_empty() && caption_text != text {
            let (mut caption_x, mut caption_y, mut caption_width, mut caption_height) =
                match caption_placement {
                    "top" => (x, y, width, caption_reserve.max(height)),
                    "right" => (x + field_width, y, caption_reserve, height),
                    "bottom" => (x, y + field_height, width, caption_reserve),
                    _ => (x, y, caption_reserve.max(width), height),
                };
            caption_x += inset_left;
            caption_y += inset_top;
            caption_width = (caption_width - inset_left - inset_right).max(0.1);
            caption_height = (caption_height - inset_top - inset_bottom).max(0.1);
            self.render_text(
                &caption_text,
                index,
                caption_x,
                caption_y,
                caption_width,
                caption_height,
                page,
                true,
            );
        }
        let choice_list = descendant(node, self.nodes, "choiceList");
        let open = choice_list
            .and_then(|choice| choice.attributes.get("open"))
            .map(String::as_str);
        if is_field && matches!(open, Some("always" | "multiSelect")) {
            let items = choice_items(node, self.nodes);
            let selected = self
                .values
                .get(
                    node.attributes
                        .get("name")
                        .map(String::as_str)
                        .unwrap_or_default(),
                )
                .cloned()
                .unwrap_or_default();
            let font = descendant(node, self.nodes, "font");
            let size = font
                .and_then(|font| length_attr(font, "size"))
                .unwrap_or(10.0 * PT_TO_CSS);
            let row_height = (size * 1.15).max(1.0);
            for (row, item) in items.iter().enumerate() {
                let row_y = field_y + row as f32 * row_height;
                if row_y + row_height > field_y + field_height + 0.01 {
                    break;
                }
                if selected.iter().any(|value| value == item) {
                    self.pages[page].push_str(&format!(
                        "<rect x=\"{field_x}\" y=\"{row_y}\" width=\"{field_width}\" height=\"{row_height}\" fill=\"rgb(158,199,220)\"/>",
                    ));
                }
                self.render_text(
                    item,
                    index,
                    field_x + 1.0,
                    row_y,
                    field_width - 2.0,
                    row_height,
                    page,
                    false,
                );
            }
            text.clear();
        }
        if is_field
            && let Some(cells) = descendant(node, self.nodes, "comb")
                .and_then(|comb| comb.attributes.get("numberOfCells"))
                .and_then(|value| value.parse::<usize>().ok())
                .filter(|cells| *cells > 0 && *cells <= 256)
        {
            let cell_width = field_width / cells as f32;
            for cell in 1..cells {
                let line_x = field_x + cell as f32 * cell_width;
                self.pages[page].push_str(&format!(
                    "<line x1=\"{line_x}\" y1=\"{field_y}\" x2=\"{line_x}\" y2=\"{}\" stroke=\"black\" stroke-width=\"0.7\"/>",
                    field_y + field_height
                ));
            }
            let font = descendant(node, self.nodes, "font");
            let size = font
                .and_then(|font| length_attr(font, "size"))
                .unwrap_or(10.0 * PT_TO_CSS);
            let family = font
                .and_then(|font| font.attributes.get("typeface"))
                .map(String::as_str)
                .unwrap_or("Arial");
            for (cell, character) in text.chars().take(cells).enumerate() {
                let text_x = field_x + (cell as f32 + 0.5) * cell_width;
                let text_y = field_y + (field_height - size).max(0.0) / 2.0 + size;
                self.pages[page].push_str(&format!(
                    "<text x=\"{text_x}\" y=\"{text_y}\" text-anchor=\"middle\" font-family=\"{}\" font-size=\"{size}\" fill=\"black\">{}</text>",
                    escape(family), escape(&character.to_string())
                ));
            }
            text.clear();
        }
        if !text.is_empty() {
            let inset = if is_field { 4.0 } else { 1.0 };
            self.render_text(
                &text,
                index,
                field_x + inset,
                field_y + inset,
                field_width - inset * 2.0,
                field_height,
                page,
                false,
            );
            if !self.page_text[page].is_empty() {
                self.page_text[page].push('\n');
            }
            self.page_text[page].push_str(&text);
        }
    }

    fn field_value(&mut self, index: usize) -> String {
        let node = &self.nodes[index];
        let name = node.attributes.get("name").cloned().unwrap_or_default();
        let cursor = self.value_indices.entry(name.clone()).or_default();
        let value = self
            .values
            .get(&name)
            .and_then(|values| values.get(*cursor))
            .cloned();
        *cursor += usize::from(value.is_some());
        let value = value.unwrap_or_else(|| {
            direct_child(node, self.nodes, "value")
                .map(|value| textual_value(value, self.nodes))
                .unwrap_or_default()
        });
        visible_item(node, self.nodes, &value).unwrap_or(value)
    }

    fn render_text(
        &mut self,
        text: &str,
        index: usize,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        page: usize,
        caption: bool,
    ) {
        let font = descendant(&self.nodes[index], self.nodes, "font");
        let size = font
            .and_then(|font| length_attr(font, "size"))
            .unwrap_or(10.0 * PT_TO_CSS);
        let family = font
            .and_then(|font| font.attributes.get("typeface"))
            .map(String::as_str)
            .unwrap_or("Arial");
        let weight = font
            .and_then(|font| font.attributes.get("weight"))
            .map(String::as_str)
            .unwrap_or("normal");
        let line_height = (size * 1.15).max(1.0);
        let max_chars = ((width / (size * 0.52)).floor() as usize).max(1);
        let lines = wrap(
            text,
            max_chars,
            ((height / line_height).floor() as usize).max(1),
        );
        let node = &self.nodes[index];
        let para = if caption {
            descendant(node, self.nodes, "caption")
                .and_then(|caption| descendant(caption, self.nodes, "para"))
        } else {
            direct_child(node, self.nodes, "para")
        };
        let horizontal_align = para
            .and_then(|para| para.attributes.get("hAlign"))
            .map(String::as_str)
            .unwrap_or("left");
        let (text_x, anchor) = match horizontal_align {
            "center" => (x + width / 2.0, "middle"),
            "right" => (x + width, "end"),
            _ => (x, "start"),
        };
        let start_y = y + if caption { size } else { size.min(height) };
        self.pages[page].push_str(&format!(
            "<text x=\"{text_x}\" y=\"{start_y}\" text-anchor=\"{anchor}\" font-family=\"{}\" font-size=\"{size}\" font-weight=\"{}\" fill=\"black\">",
            escape(family), escape(weight)
        ));
        for (line_index, line) in lines.iter().enumerate() {
            self.pages[page].push_str(&format!(
                "<tspan x=\"{text_x}\" dy=\"{}\">{}</tspan>",
                if line_index == 0 { 0.0 } else { line_height },
                escape(line)
            ));
        }
        self.pages[page].push_str("</text>");
    }

    fn render_line(&mut self, index: usize, x: f32, y: f32, page: usize) {
        let node = &self.nodes[index];
        let width = length_attr(node, "w").unwrap_or(0.0);
        let height = length_attr(node, "h").unwrap_or(0.0);
        self.pages[page].push_str(&format!(
            "<line x1=\"{x}\" y1=\"{y}\" x2=\"{}\" y2=\"{}\" stroke=\"black\" stroke-width=\"0.7\"/>",
            x + width,
            y + height
        ));
    }

    fn render_box(&mut self, index: usize, x: f32, y: f32, page: usize) {
        let node = &self.nodes[index];
        let width = length_attr(node, "w").unwrap_or(0.0);
        let height = length_attr(node, "h").unwrap_or(0.0);
        let tag = if local_name(&node.name) == "arc" {
            "ellipse"
        } else {
            "rect"
        };
        if tag == "ellipse" {
            self.pages[page].push_str(&format!(
                "<ellipse cx=\"{}\" cy=\"{}\" rx=\"{}\" ry=\"{}\" fill=\"none\" stroke=\"black\" stroke-width=\"0.7\"/>",
                x + width / 2.0, y + height / 2.0, width / 2.0, height / 2.0
            ));
        } else {
            self.pages[page].push_str(&format!(
                "<rect x=\"{x}\" y=\"{y}\" width=\"{width}\" height=\"{height}\" fill=\"none\" stroke=\"black\" stroke-width=\"0.7\"/>"
            ));
        }
    }
}

fn parse_nodes(bytes: &[u8], limits: Limits) -> Result<Vec<Node>, Diagnostic> {
    match parse_nodes_once(bytes, limits) {
        Ok(nodes) => Ok(nodes),
        Err(error) if bytes.contains(&b'&') => {
            let repaired = escape_bare_ampersands(bytes);
            parse_nodes_once(&repaired, limits).map_err(|_| error)
        }
        Err(error) => Err(error),
    }
}

fn parse_nodes_once(bytes: &[u8], limits: Limits) -> Result<Vec<Node>, Diagnostic> {
    let mut nodes = Vec::<Node>::new();
    let mut stack = Vec::<usize>::new();
    parse_xml(bytes, limits, |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let index = nodes.len();
                nodes.push(Node {
                    name: name.to_owned(),
                    attributes: attributes
                        .into_iter()
                        .map(|attribute| {
                            Ok((
                                local_name(attribute.name).to_owned(),
                                decode_xml_text(attribute.value)?.into_owned(),
                            ))
                        })
                        .collect::<Result<_, Diagnostic>>()?,
                    ..Node::default()
                });
                if let Some(parent) = stack.last().copied() {
                    nodes[parent].children.push(index);
                }
                if !empty {
                    stack.push(index);
                }
            }
            XmlEvent::EndElement { .. } => {
                stack.pop();
            }
            XmlEvent::Text(text) | XmlEvent::Cdata(text) => {
                if let Some(index) = stack.last().copied() {
                    nodes[index].text.push_str(&decode_xml_text(text)?);
                }
            }
        }
        Ok(())
    })?;
    Ok(nodes)
}

fn escape_bare_ampersands(bytes: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] != b'&' {
            output.push(bytes[index]);
            index += 1;
            continue;
        }
        let tail = &bytes[index + 1..];
        let valid = [b"amp;".as_slice(), b"lt;", b"gt;", b"quot;", b"apos;"]
            .into_iter()
            .any(|entity| tail.starts_with(entity))
            || tail.first() == Some(&b'#') && tail.iter().take(12).any(|byte| *byte == b';');
        if valid {
            output.push(b'&');
        } else {
            output.extend_from_slice(b"&amp;");
        }
        index += 1;
    }
    output
}

fn dataset_values(
    bytes: &[u8],
    limits: Limits,
) -> Result<BTreeMap<String, Vec<String>>, Diagnostic> {
    let nodes = parse_nodes(bytes, limits)?;
    let mut values = BTreeMap::<String, Vec<String>>::new();
    for node in &nodes {
        let child_values = node
            .children
            .iter()
            .filter_map(|index| {
                let child = &nodes[*index];
                child
                    .children
                    .is_empty()
                    .then(|| child.text.trim())
                    .filter(|value| !value.is_empty())
            })
            .map(str::to_owned)
            .collect::<Vec<_>>();
        if !child_values.is_empty() {
            values
                .entry(local_name(&node.name).to_owned())
                .or_default()
                .extend(child_values);
        }
        if node.children.is_empty() {
            let value = node.text.trim();
            if !value.is_empty() {
                values
                    .entry(local_name(&node.name).to_owned())
                    .or_default()
                    .push(value.to_owned());
            }
        }
    }
    Ok(values)
}

fn choice_items(node: &Node, nodes: &[Node]) -> Vec<String> {
    node.children
        .iter()
        .find_map(|index| {
            let child = &nodes[*index];
            (local_name(&child.name) == "items"
                && child.attributes.get("save").map(String::as_str) == Some("1"))
            .then_some(child)
        })
        .map(|items| {
            items
                .children
                .iter()
                .map(|index| textual_value(&nodes[*index], nodes))
                .filter(|value| !value.is_empty())
                .collect()
        })
        .unwrap_or_default()
}

fn saved_page_count(bytes: &[u8], limits: Limits) -> usize {
    let Ok(nodes) = parse_nodes(bytes, limits) else {
        return 1;
    };
    nodes
        .iter()
        .filter(|node| local_name(&node.name) == "pageArea")
        .count()
        .max(1)
}

fn page_geometry(nodes: &[Node], page: usize) -> PageGeometry {
    let medium = descendant(&nodes[page], nodes, "medium");
    let width = medium
        .and_then(|node| length_attr(node, "short"))
        .unwrap_or(816.0);
    let height = medium
        .and_then(|node| length_attr(node, "long"))
        .unwrap_or(1056.0);
    let content = descendant(&nodes[page], nodes, "contentArea");
    PageGeometry {
        width,
        height,
        content_x: content
            .and_then(|node| length_attr(node, "x"))
            .unwrap_or(0.0),
        content_y: content
            .and_then(|node| length_attr(node, "y"))
            .unwrap_or(0.0),
        content_height: content
            .and_then(|node| length_attr(node, "h"))
            .unwrap_or(height),
    }
}

fn node_height(nodes: &[Node], index: usize) -> f32 {
    let node = &nodes[index];
    if let Some(height) = length_attr(node, "h") {
        return height;
    }
    if node.attributes.get("layout").map(String::as_str) == Some("tb") {
        return node
            .children
            .iter()
            .filter(|child| visible(&nodes[**child]))
            .map(|child| node_height(nodes, *child))
            .sum();
    }
    if matches!(
        node.attributes.get("layout").map(String::as_str),
        Some("lr-tb" | "rl-tb")
    ) {
        let available_width = length_attr(node, "w")
            .unwrap_or_else(|| node_width(nodes, index))
            .max(0.1);
        let mut row_width = 0.0;
        let mut row_height = 0.0_f32;
        let mut height = 0.0;
        for &child in node
            .children
            .iter()
            .filter(|child| renderable_child(&nodes[**child]) && visible(&nodes[**child]))
        {
            let width = node_width(nodes, child).min(available_width).max(0.1);
            if row_width > 0.0 && row_width + width > available_width + 0.01 {
                height += row_height;
                row_width = 0.0;
                row_height = 0.0;
            }
            row_width += width;
            row_height = row_height.max(node_height(nodes, child));
        }
        return height + row_height;
    }
    node.children
        .iter()
        .filter(|child| visible(&nodes[**child]))
        .map(|child| length_attr(&nodes[*child], "y").unwrap_or(0.0) + node_height(nodes, *child))
        .fold(0.0, f32::max)
}

fn node_width(nodes: &[Node], index: usize) -> f32 {
    let node = &nodes[index];
    if let Some(width) = length_attr(node, "w") {
        return width;
    }
    node.children
        .iter()
        .filter(|child| renderable_child(&nodes[**child]) && visible(&nodes[**child]))
        .map(|child| length_attr(&nodes[*child], "x").unwrap_or(0.0) + node_width(nodes, *child))
        .fold(0.0, f32::max)
}

fn renderable_child(node: &Node) -> bool {
    !matches!(
        local_name(&node.name),
        "value"
            | "caption"
            | "font"
            | "para"
            | "border"
            | "ui"
            | "assist"
            | "bind"
            | "event"
            | "script"
            | "items"
            | "margin"
            | "breakBefore"
            | "breakAfter"
            | "contentArea"
            | "medium"
    )
}

fn visible(node: &Node) -> bool {
    !matches!(
        node.attributes.get("presence").map(String::as_str),
        Some("hidden" | "invisible")
    )
}

fn descendant<'a>(node: &'a Node, nodes: &'a [Node], name: &str) -> Option<&'a Node> {
    node.children.iter().find_map(|index| {
        let child = &nodes[*index];
        (local_name(&child.name) == name)
            .then_some(child)
            .or_else(|| descendant(child, nodes, name))
    })
}

fn direct_child<'a>(node: &'a Node, nodes: &'a [Node], name: &str) -> Option<&'a Node> {
    node.children
        .iter()
        .map(|index| &nodes[*index])
        .find(|child| local_name(&child.name) == name)
}

fn descendants<'a>(nodes: &'a [Node], root: usize) -> impl Iterator<Item = usize> + 'a {
    let mut pending = vec![root];
    std::iter::from_fn(move || {
        let index = pending.pop()?;
        pending.extend(nodes[index].children.iter().rev().copied());
        Some(index)
    })
}

fn textual_value(node: &Node, nodes: &[Node]) -> String {
    fn visit(node: &Node, nodes: &[Node], output: &mut String) {
        if matches!(local_name(&node.name), "script" | "items" | "image") {
            return;
        }
        let text = node.text.trim();
        if !text.is_empty() {
            if !output.is_empty() {
                output.push(' ');
            }
            output.push_str(text);
        }
        for child in &node.children {
            visit(&nodes[*child], nodes, output);
        }
    }
    let mut output = String::new();
    visit(node, nodes, &mut output);
    output
}

fn visible_item(node: &Node, nodes: &[Node], saved: &str) -> Option<String> {
    let items = node
        .children
        .iter()
        .filter_map(|index| (local_name(&nodes[*index].name) == "items").then_some(&nodes[*index]))
        .collect::<Vec<_>>();
    let saved_items = items
        .iter()
        .find(|items| items.attributes.get("save").map(String::as_str) == Some("1"))?;
    let visible_items = items
        .iter()
        .find(|items| items.attributes.get("save").map(String::as_str) == Some("0"))?;
    let saved_values = saved_items
        .children
        .iter()
        .map(|index| textual_value(&nodes[*index], nodes))
        .collect::<Vec<_>>();
    let visible_values = visible_items
        .children
        .iter()
        .map(|index| textual_value(&nodes[*index], nodes))
        .collect::<Vec<_>>();
    saved_values
        .iter()
        .position(|value| value == saved)
        .and_then(|index| visible_values.get(index))
        .cloned()
}

fn length_attr(node: &Node, name: &str) -> Option<f32> {
    parse_length(node.attributes.get(name)?)
}

fn parse_length(value: &str) -> Option<f32> {
    let value = value.trim();
    let (number, scale) = if let Some(number) = value.strip_suffix("mm") {
        (number, MM_TO_CSS)
    } else if let Some(number) = value.strip_suffix("pt") {
        (number, PT_TO_CSS)
    } else if let Some(number) = value.strip_suffix("in") {
        (number, 96.0)
    } else if let Some(number) = value.strip_suffix("cm") {
        (number, MM_TO_CSS * 10.0)
    } else if let Some(number) = value.strip_suffix("px") {
        (number, 1.0)
    } else {
        (value, 1.0)
    };
    number
        .trim()
        .parse::<f32>()
        .ok()
        .filter(|value| value.is_finite())
        .map(|value| value * scale)
}

fn rgb(value: &str) -> Option<String> {
    let components = value
        .split(',')
        .map(str::trim)
        .map(str::parse::<u8>)
        .collect::<Result<Vec<_>, _>>()
        .ok()?;
    (components.len() >= 3)
        .then(|| format!("rgb({},{},{})", components[0], components[1], components[2]))
}

fn wrap(text: &str, max_chars: usize, max_lines: usize) -> Vec<String> {
    let mut lines = Vec::new();
    for paragraph in text
        .split(['\r', '\n'])
        .filter(|paragraph| !paragraph.is_empty())
    {
        let mut current = String::new();
        for word in paragraph.split_whitespace() {
            if !current.is_empty() && current.chars().count() + 1 + word.chars().count() > max_chars
            {
                lines.push(std::mem::take(&mut current));
                if lines.len() >= max_lines {
                    return lines;
                }
            }
            if !current.is_empty() {
                current.push(' ');
            }
            current.push_str(word);
        }
        if !current.is_empty() {
            lines.push(current);
            if lines.len() >= max_lines {
                return lines;
            }
        }
    }
    lines
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn object_limit() -> Diagnostic {
    Diagnostic::fatal(
        DiagnosticCode::ObjectLimit,
        Phase::Layout,
        None,
        "PDF XFA output exceeds the configured object budget",
    )
    .in_part("document.pdf")
}

#[cfg(test)]
mod tests {
    use super::{parse_length, wrap};

    #[test]
    fn converts_xfa_lengths_to_css_pixels() {
        assert!((parse_length("25.4mm").unwrap() - 96.0).abs() < 0.001);
        assert!((parse_length("72pt").unwrap() - 96.0).abs() < 0.001);
    }

    #[test]
    fn wraps_xfa_text_within_the_requested_line_budget() {
        assert_eq!(wrap("one two three four", 7, 2), ["one two", "three"]);
        assert_eq!(
            wrap("one two\rthree four", 20, 2),
            ["one two", "three four"]
        );
    }
}
