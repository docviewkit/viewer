//! Format-independent rendering primitives and the separate source-object table.

use std::cmp::Reverse;
use std::collections::BinaryHeap;

use crate::diagnostic::Diagnostic;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DocumentFormat {
    Pptx,
    Odp,
    Xlsx,
    Ods,
    Docx,
    Odt,
    Csv,
    Rtf,
    Ppt,
    Xls,
    Doc,
    Pages,
    Numbers,
    Keynote,
    Pdf,
    Xps,
    Ofd,
}

impl DocumentFormat {
    pub const fn code(self) -> u8 {
        match self {
            Self::Pptx => 1,
            Self::Odp => 2,
            Self::Xlsx => 3,
            Self::Ods => 4,
            Self::Docx => 5,
            Self::Odt => 6,
            Self::Csv => 8,
            Self::Rtf => 10,
            Self::Ppt => 11,
            Self::Xls => 12,
            Self::Doc => 13,
            Self::Pages => 14,
            Self::Numbers => 15,
            Self::Keynote => 16,
            Self::Pdf => 17,
            Self::Xps => 18,
            Self::Ofd => 19,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DocumentKind {
    Presentation,
    Spreadsheet,
    Text,
}

impl DocumentKind {
    pub const fn code(self) -> u8 {
        match self {
            Self::Presentation => 1,
            Self::Spreadsheet => 2,
            Self::Text => 3,
        }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Rect {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
}

impl Rect {
    pub fn is_valid(self) -> bool {
        self.x.is_finite()
            && self.y.is_finite()
            && self.width.is_finite()
            && self.height.is_finite()
            && self.width >= 0.0
            && self.height >= 0.0
    }

    pub fn contains(self, x: f32, y: f32) -> bool {
        self.is_valid()
            && x >= self.x
            && y >= self.y
            && x <= self.x + self.width
            && y <= self.y + self.height
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum UnitKind {
    Slide,
    Sheet,
    Page,
}

impl UnitKind {
    pub const fn code(self) -> u8 {
        match self {
            Self::Slide => 1,
            Self::Sheet => 2,
            Self::Page => 3,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct SlideMetadata {
    pub source_id: Option<String>,
    pub source_part: Option<String>,
    pub speaker_notes: Option<String>,
    pub speaker_notes_part: Option<String>,
    pub speaker_note_paragraphs: Vec<SpeakerNoteParagraph>,
    pub number: u32,
    pub hidden: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub struct SpeakerNoteParagraph {
    pub layout: TextParagraphLayout,
    pub runs: Vec<TextRun>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct SheetAxisSpan {
    pub start: u32,
    pub end: u32,
    pub size: f32,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct SheetAxis {
    pub default_size: f32,
    pub spans: Vec<SheetAxisSpan>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SheetOrientation {
    Portrait,
    Landscape,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SheetViewMode {
    PageLayout,
    PageBreakPreview,
}

impl SheetOrientation {
    pub const fn code(self) -> u8 {
        match self {
            Self::Portrait => 1,
            Self::Landscape => 2,
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct SheetMargins {
    pub left: Option<f32>,
    pub right: Option<f32>,
    pub top: Option<f32>,
    pub bottom: Option<f32>,
    pub header: Option<f32>,
    pub footer: Option<f32>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct SheetPrintSettings {
    pub view_mode: Option<SheetViewMode>,
    pub print_area: Option<String>,
    pub print_titles: Option<String>,
    // Boundary id and inclusive perpendicular span, in worksheet coordinates.
    pub row_breaks: Vec<[u32; 3]>,
    pub column_breaks: Vec<[u32; 3]>,
    pub paper_size: Option<u32>,
    pub orientation: Option<SheetOrientation>,
    pub scale: Option<u32>,
    pub fit_to_width: Option<u32>,
    pub fit_to_height: Option<u32>,
    pub fit_to_page: bool,
    pub margins: SheetMargins,
    pub different_odd_even: bool,
    pub different_first: bool,
    pub odd_header: Option<String>,
    pub odd_footer: Option<String>,
    pub even_header: Option<String>,
    pub even_footer: Option<String>,
    pub first_header: Option<String>,
    pub first_footer: Option<String>,
}

impl SheetAxis {
    pub fn offset(&self, index: u32) -> f32 {
        self.spans
            .iter()
            .fold(index as f32 * self.default_size, |offset, span| {
                let covered = index
                    .min(span.end.saturating_add(1))
                    .saturating_sub(span.start);
                offset + covered as f32 * (span.size - self.default_size)
            })
    }

    pub fn uniform(count: u32, total_size: f32) -> Self {
        Self {
            default_size: if count == 0 {
                0.0
            } else {
                total_size / count as f32
            },
            spans: Vec::new(),
        }
    }

    pub fn from_sizes(sizes: &[f32], fallback: f32) -> Self {
        let default_size = sizes.first().copied().unwrap_or(fallback);
        let mut spans: Vec<SheetAxisSpan> = Vec::new();
        for (index, &size) in sizes
            .iter()
            .enumerate()
            .filter(|(_, size)| size.to_bits() != default_size.to_bits())
        {
            let index = u32::try_from(index).unwrap_or(u32::MAX);
            if let Some(last) = spans.last_mut()
                && last.end.saturating_add(1) == index
                && last.size.to_bits() == size.to_bits()
            {
                last.end = index;
            } else {
                spans.push(SheetAxisSpan {
                    start: index,
                    end: index,
                    size,
                });
            }
        }
        Self {
            default_size,
            spans,
        }
    }
}

/// Smallest axis count covering an extent, bounded by the format's row/column limit.
pub(crate) fn sheet_axis_count_for_extent(
    extent: f32,
    limit: u32,
    offset: impl Fn(u32) -> f32,
) -> u32 {
    if extent <= 0.0 {
        return 0;
    }
    let mut low = 0_u32;
    let mut high = limit;
    while low < high {
        let middle = low + (high - low) / 2;
        if offset(middle) >= extent {
            high = middle;
        } else {
            low = middle.saturating_add(1);
        }
    }
    low.max(1)
}

#[derive(Clone, Debug, PartialEq)]
pub struct Unit {
    pub kind: UnitKind,
    pub index: u32,
    pub id: String,
    pub name: String,
    pub width: f32,
    pub height: f32,
    pub rows: u32,
    pub columns: u32,
    pub frozen_rows: u32,
    pub frozen_columns: u32,
    pub frozen_width: f32,
    pub frozen_height: f32,
    pub row_axis: SheetAxis,
    pub column_axis: SheetAxis,
    pub show_grid_lines: bool,
    /// Authored worksheet tab color, packed as 0xRRGGBBAA.
    pub tab_color: Option<u32>,
    pub sheet: Option<SheetPrintSettings>,
    pub slide: Option<SlideMetadata>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MappingQuality {
    Exact,
    Derived,
    Approximate,
}

impl MappingQuality {
    pub const fn code(self) -> u8 {
        match self {
            Self::Exact => 0,
            Self::Derived => 1,
            Self::Approximate => 2,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct PptxAction {
    pub kind: String,
    pub action: Option<String>,
    pub target: Option<String>,
    pub tooltip: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct PptxObjectMetadata {
    pub name: Option<String>,
    pub title: Option<String>,
    pub description: Option<String>,
    pub hidden: bool,
    pub click_action: Option<PptxAction>,
    pub hover_action: Option<PptxAction>,
}

#[derive(Clone, Debug, PartialEq)]
pub enum SourceLocator {
    PptxShape {
        shape_id: u32,
        row: Option<u32>,
        column: Option<u32>,
        text_range: Option<(u32, u32)>,
        metadata: PptxObjectMetadata,
    },
    OdpElement {
        element_id: Option<String>,
        path: String,
        row: Option<u32>,
        column: Option<u32>,
    },
    Xlsx {
        kind: &'static str,
        sheet_name: String,
        address: Option<String>,
        formula: Option<String>,
        drawing_id: Option<u32>,
    },
    Ods {
        kind: &'static str,
        table_name: String,
        row: Option<u32>,
        column: Option<u32>,
        element_id: Option<String>,
        path: String,
    },
    Docx {
        kind: &'static str,
        paragraph_id: Option<String>,
        paragraph_index: Option<u32>,
        drawing_id: Option<u32>,
        row: Option<u32>,
        column: Option<u32>,
        text_range: Option<(u32, u32)>,
        action: Option<PptxAction>,
    },
    Odt {
        kind: &'static str,
        element_id: Option<String>,
        path: String,
        row: Option<u32>,
        column: Option<u32>,
        text_range: Option<(u32, u32)>,
    },
    Flat {
        kind: &'static str,
        index: Option<u32>,
        row: Option<u32>,
        column: Option<u32>,
        text_range: Option<(u32, u32)>,
    },
    Legacy {
        kind: &'static str,
        stream: String,
        record_offset: Option<u32>,
        row: Option<u32>,
        column: Option<u32>,
        text_range: Option<(u32, u32)>,
    },
    Iwork {
        kind: &'static str,
        component: String,
    },
    Pdf {
        kind: &'static str,
        object_number: Option<u32>,
        byte_offset: Option<u32>,
        action: Option<PptxAction>,
    },
    Xps {
        kind: &'static str,
        path: String,
    },
    Ofd {
        kind: &'static str,
        path: String,
    },
}

#[derive(Clone, Debug, PartialEq)]
pub struct SourceRef {
    pub part: String,
    pub mapping: MappingQuality,
    pub locator: SourceLocator,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ObjectKind {
    Group,
    TextBox,
    Paragraph,
    Image,
    Shape,
    Table,
    Cell,
    Unknown,
}

impl ObjectKind {
    pub const fn code(self) -> u8 {
        match self {
            Self::Group => 1,
            Self::TextBox => 2,
            Self::Paragraph => 3,
            Self::Image => 4,
            Self::Shape => 5,
            Self::Table => 6,
            Self::Cell => 7,
            Self::Unknown => 8,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct AffineTransform {
    pub a: f32,
    pub b: f32,
    pub c: f32,
    pub d: f32,
    pub e: f32,
    pub f: f32,
}

impl AffineTransform {
    pub const IDENTITY: Self = Self {
        a: 1.0,
        b: 0.0,
        c: 0.0,
        d: 1.0,
        e: 0.0,
        f: 0.0,
    };

    pub fn is_valid(self) -> bool {
        [self.a, self.b, self.c, self.d, self.e, self.f]
            .into_iter()
            .all(f32::is_finite)
    }

    /// Concatenates `next` after this transform, matching Canvas 2D matrix
    /// multiplication and parent-before-child scene traversal.
    pub fn concat(self, next: Self) -> Self {
        Self {
            a: self.a * next.a + self.c * next.b,
            b: self.b * next.a + self.d * next.b,
            c: self.a * next.c + self.c * next.d,
            d: self.b * next.c + self.d * next.d,
            e: self.a * next.e + self.c * next.f + self.e,
            f: self.b * next.e + self.d * next.f + self.f,
        }
    }

    /// Rotation by `degrees` about `(center_x, center_y)`, expressed as
    /// `translate(center) · rotate · translate(-center)` so it matches the
    /// Canvas 2D convention used by `TextLayout::rotation_degrees`.
    ///
    /// Returns `IDENTITY` for non-finite inputs or a zero angle, so callers can
    /// apply it unconditionally.
    pub fn rotation_about(center_x: f32, center_y: f32, degrees: f32) -> Self {
        if !degrees.is_finite() || degrees == 0.0 {
            return Self::IDENTITY;
        }
        let radians = degrees.to_radians();
        let (sin, cos) = (radians.sin(), radians.cos());
        if !sin.is_finite() || !cos.is_finite() {
            return Self::IDENTITY;
        }
        Self {
            a: cos,
            b: sin,
            c: -sin,
            d: cos,
            e: center_x - center_x * cos + center_y * sin,
            f: center_y - center_x * sin - center_y * cos,
        }
    }

    pub fn inverse_transform_point(self, x: f32, y: f32) -> Option<(f32, f32)> {
        let determinant = self.a * self.d - self.b * self.c;
        if !determinant.is_finite() || determinant.abs() <= f32::EPSILON {
            return None;
        }
        let translated_x = x - self.e;
        let translated_y = y - self.f;
        Some((
            (self.d * translated_x - self.c * translated_y) / determinant,
            (-self.b * translated_x + self.a * translated_y) / determinant,
        ))
    }
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum BlendMode {
    #[default]
    Normal,
    Multiply,
    Screen,
    Overlay,
    Darken,
    Lighten,
    ColorDodge,
    ColorBurn,
    HardLight,
    SoftLight,
    Difference,
    Exclusion,
    Hue,
    Saturation,
    Color,
    Luminosity,
    DestinationOut,
}

impl BlendMode {
    pub const fn code(self) -> u8 {
        match self {
            Self::Normal => 0,
            Self::Multiply => 1,
            Self::Screen => 2,
            Self::Overlay => 3,
            Self::Darken => 4,
            Self::Lighten => 5,
            Self::ColorDodge => 6,
            Self::ColorBurn => 7,
            Self::HardLight => 8,
            Self::SoftLight => 9,
            Self::Difference => 10,
            Self::Exclusion => 11,
            Self::Hue => 12,
            Self::Saturation => 13,
            Self::Color => 14,
            Self::Luminosity => 15,
            Self::DestinationOut => 16,
        }
    }
}

impl Default for AffineTransform {
    fn default() -> Self {
        Self::IDENTITY
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum FillRule {
    NonZero,
    EvenOdd,
}

impl FillRule {
    pub const fn code(self) -> u8 {
        match self {
            Self::NonZero => 0,
            Self::EvenOdd => 1,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum PathCommand {
    MoveTo {
        x: f32,
        y: f32,
    },
    LineTo {
        x: f32,
        y: f32,
    },
    QuadraticCurveTo {
        cpx: f32,
        cpy: f32,
        x: f32,
        y: f32,
    },
    BezierCurveTo {
        cp1x: f32,
        cp1y: f32,
        cp2x: f32,
        cp2y: f32,
        x: f32,
        y: f32,
    },
    ClosePath,
}

impl PathCommand {
    /// Apply an affine mapping to endpoints and control points alike.
    pub fn transform(&mut self, transform: AffineTransform) {
        let point = |x: &mut f32, y: &mut f32| {
            (*x, *y) = (
                transform.a * *x + transform.c * *y + transform.e,
                transform.b * *x + transform.d * *y + transform.f,
            );
        };
        match self {
            Self::MoveTo { x, y } | Self::LineTo { x, y } => point(x, y),
            Self::QuadraticCurveTo { cpx, cpy, x, y } => {
                point(cpx, cpy);
                point(x, y);
            }
            Self::BezierCurveTo {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                point(cp1x, cp1y);
                point(cp2x, cp2y);
                point(x, y);
            }
            Self::ClosePath => {}
        }
    }

    pub const fn code(&self) -> u8 {
        match self {
            Self::MoveTo { .. } => 0,
            Self::LineTo { .. } => 1,
            Self::QuadraticCurveTo { .. } => 2,
            Self::BezierCurveTo { .. } => 3,
            Self::ClosePath => 4,
        }
    }
}

#[cfg(any(
    test,
    feature = "odf-formats",
    feature = "legacy-office-formats",
    feature = "xps-formats"
))]
pub(crate) fn append_dashed_polyline(
    commands: &mut Vec<PathCommand>,
    points: &[(f32, f32)],
    dash: &[f32],
    offset: f32,
) {
    let mut dash_index = 0_usize;
    let mut dash_remaining = dash[0];
    let period: f32 = dash.iter().sum();
    let mut phase = offset.rem_euclid(period);
    while phase > 0.0 {
        if phase < dash_remaining {
            dash_remaining -= phase;
            break;
        }
        phase -= dash_remaining;
        dash_index = (dash_index + 1) % dash.len();
        dash_remaining = dash[dash_index];
    }
    let mut drawing = dash_index.is_multiple_of(2);
    let mut dash_open = false;
    for segment in points.windows(2) {
        let mut current = segment[0];
        let end = segment[1];
        let delta_x = end.0 - current.0;
        let delta_y = end.1 - current.1;
        let length = delta_x.hypot(delta_y);
        if length <= f32::EPSILON {
            continue;
        }
        let direction_x = delta_x / length;
        let direction_y = delta_y / length;
        let mut segment_remaining = length;
        // A period ending at the path endpoint has no following dash. Ignore
        // only f32 subtraction residue, which round/square caps would amplify.
        let tolerance = length * f32::EPSILON * 8.0;
        while segment_remaining > tolerance {
            let step = dash_remaining.min(segment_remaining);
            let next = (
                current.0 + direction_x * step,
                current.1 + direction_y * step,
            );
            if drawing {
                if !dash_open {
                    commands.push(PathCommand::MoveTo {
                        x: current.0,
                        y: current.1,
                    });
                    dash_open = true;
                }
                commands.push(PathCommand::LineTo {
                    x: next.0,
                    y: next.1,
                });
            }
            current = next;
            segment_remaining -= step;
            dash_remaining -= step;
            if dash_remaining <= f32::EPSILON {
                dash_index = (dash_index + 1) % dash.len();
                dash_remaining = dash[dash_index];
                drawing = dash_index.is_multiple_of(2);
                if !drawing {
                    dash_open = false;
                }
            }
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PathFillMode {
    Normal,
    None,
    Darken,
    DarkenLess,
    Lighten,
    LightenLess,
}

impl PathFillMode {
    pub const fn code(self) -> u8 {
        match self {
            Self::Normal => 0,
            Self::None => 1,
            Self::Darken => 2,
            Self::DarkenLess => 3,
            Self::Lighten => 4,
            Self::LightenLess => 5,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct PathLayer {
    pub fill_rule: FillRule,
    pub fill: PathFillMode,
    pub stroke: bool,
    pub commands: Vec<PathCommand>,
}

#[derive(Clone, Debug, PartialEq)]
pub enum Geometry {
    Rectangle,
    Ellipse,
    Line,
    RoundedRectangle {
        radius_x: f32,
        radius_y: f32,
    },
    /// Commands use object-local document-space coordinates, relative to the
    /// top-left of the object's bounds.
    Path {
        fill_rule: FillRule,
        commands: Vec<PathCommand>,
    },
    /// DrawingML preset/custom geometries may assign paint semantics to each
    /// path independently. Layers retain that authored order and styling.
    LayeredPath {
        layers: Vec<PathLayer>,
    },
}

impl Geometry {
    pub const fn code(&self) -> u8 {
        match self {
            Self::Rectangle => 0,
            Self::Ellipse => 1,
            Self::Line => 2,
            Self::RoundedRectangle { .. } => 3,
            Self::Path { .. } => 4,
            Self::LayeredPath { .. } => 5,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct GradientStop {
    pub offset: f32,
    pub color: u32,
}

#[derive(Clone, Debug, PartialEq)]
pub enum XpsColor {
    Rgba(u32),
    Context {
        alpha: f32,
        profile: Vec<u8>,
        channels: Vec<f32>,
    },
}

#[derive(Clone, Debug, PartialEq)]
pub struct XpsGradientStop {
    pub offset: f32,
    pub color: XpsColor,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TileMode {
    None,
    Tile,
    FlipX,
    FlipY,
    FlipXY,
}

impl TileMode {
    pub const fn code(self) -> u8 {
        match self {
            Self::None => 0,
            Self::Tile => 1,
            Self::FlipX => 2,
            Self::FlipY => 3,
            Self::FlipXY => 4,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum StretchMode {
    None,
    Fill,
    Uniform,
    UniformToFill,
}

impl StretchMode {
    pub const fn code(self) -> u8 {
        match self {
            Self::None => 0,
            Self::Fill => 1,
            Self::Uniform => 2,
            Self::UniformToFill => 3,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct VisualBrushChild {
    pub bounds: Rect,
    pub visual: Visual,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum GradientSpread {
    Pad,
    Reflect,
    Repeat,
}

impl GradientSpread {
    pub const fn code(self) -> u8 {
        match self {
            Self::Pad => 0,
            Self::Reflect => 1,
            Self::Repeat => 2,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum Paint {
    None,
    Solid(u32),
    /// Coordinates use the same object-local document space as paths.
    LinearGradient {
        x0: f32,
        y0: f32,
        x1: f32,
        y1: f32,
        stops: Vec<GradientStop>,
    },
    /// Two-circle radial gradient in object-local document space.
    RadialGradient {
        x0: f32,
        y0: f32,
        r0: f32,
        x1: f32,
        y1: f32,
        r1: f32,
        stops: Vec<GradientStop>,
    },
    /// DrawingML rectangular path gradient in object-local document space.
    RectGradient {
        center_x: f32,
        center_y: f32,
        stops: Vec<GradientStop>,
    },
    /// DrawingML circular path gradient between two object-local circles.
    CircleGradient {
        x0: f32,
        y0: f32,
        r0: f32,
        x1: f32,
        y1: f32,
        r1: f32,
        stops: Vec<GradientStop>,
    },
    /// Gradient from an authored focus rectangle to the actual painted contour.
    ShapeGradient {
        focus: ImageCrop,
        stops: Vec<GradientStop>,
    },
    MappedGradient {
        paint: Box<Paint>,
        tile: ImageCrop,
        flip: TileMode,
        rotate_with_shape: bool,
    },
    Pattern {
        preset: String,
        foreground: u32,
        background: u32,
    },
    Image {
        mapping: Option<Box<ImageFillMapping>>,
        media_type: String,
        bytes: Vec<u8>,
        crop: ImageCrop,
        tile: bool,
        tile_width: Option<f32>,
        tile_height: Option<f32>,
    },
    Visual {
        viewbox: Rect,
        viewport: Rect,
        viewbox_relative: bool,
        viewport_relative: bool,
        tile_mode: TileMode,
        stretch: StretchMode,
        alignment_x: f32,
        alignment_y: f32,
        transform: AffineTransform,
        relative_transform: AffineTransform,
        opacity: f32,
        children: Vec<VisualBrushChild>,
    },
    XpsGradient {
        radial: bool,
        start_x: f32,
        start_y: f32,
        end_x: f32,
        end_y: f32,
        radius_x: f32,
        radius_y: f32,
        relative: bool,
        spread: GradientSpread,
        linear_rgb: bool,
        transform: AffineTransform,
        relative_transform: AffineTransform,
        stops: Vec<XpsGradientStop>,
    },
}

impl Paint {
    /// Rebase absolute gradient coordinates into a new object's local origin.
    /// Relative gradients and non-gradient paints keep their authored mapping.
    pub fn localize_gradient(mut self, x: f32, y: f32) -> Self {
        match &mut self {
            Self::LinearGradient { x0, y0, x1, y1, .. }
            | Self::RadialGradient { x0, y0, x1, y1, .. } => {
                *x0 -= x;
                *y0 -= y;
                *x1 -= x;
                *y1 -= y;
            }
            Self::XpsGradient {
                start_x,
                start_y,
                end_x,
                end_y,
                relative: false,
                transform,
                ..
            } => {
                *start_x -= x;
                *start_y -= y;
                *end_x -= x;
                *end_y -= y;
                transform.e += transform.a * x + transform.c * y - x;
                transform.f += transform.b * x + transform.d * y - y;
            }
            _ => {}
        }
        self
    }

    pub const fn code(&self) -> u8 {
        match self {
            Self::None => 0,
            Self::Solid(_) => 1,
            Self::LinearGradient { .. } => 2,
            Self::RadialGradient { .. } => 3,
            Self::Pattern { .. } => 4,
            Self::Image { .. } => 5,
            Self::RectGradient { .. } => 6,
            Self::CircleGradient { .. } => 7,
            Self::Visual { .. } => 8,
            Self::XpsGradient { .. } => 9,
            Self::ShapeGradient { .. } => 10,
            Self::MappedGradient { .. } => 11,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextAlign {
    Start,
    Center,
    End,
    Justify,
    Distribute,
    MediumKashida,
    HighKashida,
    LowKashida,
    ThaiDistribute,
}

#[cfg(any(feature = "native-formats", feature = "iwork-formats"))]
pub(crate) fn spreadsheet_text_overflow_is_clipped(
    align: TextAlign,
    previous_occupied: bool,
    next_occupied: bool,
) -> bool {
    match align {
        TextAlign::Start => next_occupied,
        TextAlign::Center => previous_occupied || next_occupied,
        TextAlign::End => previous_occupied,
        _ => true,
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct ImageFillMapping {
    pub scale_x: f32,
    pub scale_y: f32,
    pub offset_x: f32,
    pub offset_y: f32,
    pub alignment_x: f32,
    pub alignment_y: f32,
    pub flip: TileMode,
    pub rotate_with_shape: bool,
    pub fill_rectangle: ImageCrop,
    pub dpi: f32,
}

impl Default for ImageFillMapping {
    fn default() -> Self {
        Self {
            scale_x: 1.0,
            scale_y: 1.0,
            offset_x: 0.0,
            offset_y: 0.0,
            alignment_x: 0.0,
            alignment_y: 0.0,
            flip: TileMode::None,
            rotate_with_shape: true,
            fill_rectangle: ImageCrop::default(),
            dpi: 0.0,
        }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct ImageCrop {
    pub left: f32,
    pub top: f32,
    pub right: f32,
    pub bottom: f32,
}

impl ImageCrop {
    pub fn is_valid(self) -> bool {
        [self.left, self.top, self.right, self.bottom]
            .into_iter()
            .all(|value| value.is_finite() && (-1.0..=1.0).contains(&value))
            && self.left + self.right <= 1.0
            && self.top + self.bottom <= 1.0
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct ImageAdjustment {
    pub grayscale: bool,
    pub bilevel_threshold: Option<f32>,
    pub brightness: f32,
    pub contrast: f32,
    pub duotone: Option<[u32; 2]>,
}

impl ImageAdjustment {
    pub fn is_valid(self) -> bool {
        self.bilevel_threshold
            .is_none_or(|value| value.is_finite() && (0.0..=1.0).contains(&value))
            && self.brightness.is_finite()
            && (-1.0..=1.0).contains(&self.brightness)
            && self.contrast.is_finite()
            && (-1.0..=1.0).contains(&self.contrast)
    }
}

impl TextAlign {
    pub const fn code(self) -> u8 {
        match self {
            Self::Start => 0,
            Self::Center => 1,
            Self::End => 2,
            Self::Justify => 3,
            Self::Distribute => 4,
            Self::MediumKashida => 5,
            Self::HighKashida => 6,
            Self::LowKashida => 7,
            Self::ThaiDistribute => 8,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextTabLeader {
    None,
    Dot,
    Hyphen,
    Underscore,
    MiddleDot,
}

impl TextTabLeader {
    pub const fn code(self) -> u8 {
        match self {
            Self::None => 0,
            Self::Dot => 1,
            Self::Hyphen => 2,
            Self::Underscore => 3,
            Self::MiddleDot => 4,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct TextTabStop {
    pub position: f32,
    pub align: TextAlign,
    pub leader: TextTabLeader,
}

impl From<f32> for TextTabStop {
    fn from(position: f32) -> Self {
        Self {
            position,
            align: TextAlign::Start,
            leader: TextTabLeader::None,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct TextRun {
    /// Optional authored glyph paint; absent uses the run color.
    pub paint: Option<Box<Paint>>,
    /// Apply East Asian punctuation line-start/end restrictions to this run.
    pub east_asian_line_breaks: bool,
    pub text: String,
    pub font_family: String,
    pub font_size: f32,
    pub color: u32,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    /// RGBA highlight color; an alpha byte of zero means no highlight.
    pub highlight: u32,
    /// Positive values raise the run, negative values lower it.
    pub baseline_shift: f32,
    pub letter_spacing: f32,
    /// Authored glyph width multiplier; 1.0 preserves the font's natural width.
    pub horizontal_scale: f32,
}

impl TextRun {
    #[cfg(any(
        feature = "native-formats",
        feature = "iwork-formats",
        feature = "legacy-office-formats"
    ))]
    pub(crate) fn same_style(&self, other: &Self) -> bool {
        self.paint == other.paint
            && self.east_asian_line_breaks == other.east_asian_line_breaks
            && self.font_family == other.font_family
            && self.font_size == other.font_size
            && self.color == other.color
            && self.bold == other.bold
            && self.italic == other.italic
            && self.underline == other.underline
            && self.strikethrough == other.strikethrough
            && self.highlight == other.highlight
            && self.baseline_shift == other.baseline_shift
            && self.letter_spacing == other.letter_spacing
            && self.horizontal_scale == other.horizontal_scale
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Shadow {
    pub color: u32,
    pub blur: f32,
    pub offset_x: f32,
    pub offset_y: f32,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct OuterShadow {
    pub shadow: Shadow,
    pub scale_x: f32,
    pub scale_y: f32,
    pub skew_x: f32,
    pub skew_y: f32,
    pub alignment: u8,
}

#[derive(Clone, Debug, PartialEq)]
pub struct TextEffect {
    pub stroke: Option<Box<Paint>>,
    pub stroke_width: f32,
    pub glow: Option<Glow>,
    pub fill_to_text: bool,
    pub shadow: Option<Shadow>,
    pub inner_shadow: Option<Shadow>,
    pub reflection: Option<Reflection>,
    pub wavy_underline: bool,
    pub dotted_underline: bool,
    pub heavy_underline: bool,
    pub double_underline: bool,
    pub dot_dash_underline: bool,
    pub double_strikethrough: bool,
    pub shadow_scale_x: f32,
    pub shadow_scale_y: f32,
    pub shadow_skew_x: f32,
    pub shadow_skew_y: f32,
    pub shadow_alignment: u8,
}

impl Default for TextEffect {
    fn default() -> Self {
        Self {
            stroke: None,
            stroke_width: 0.0,
            glow: None,
            fill_to_text: false,
            shadow: None,
            inner_shadow: None,
            reflection: None,
            wavy_underline: false,
            dotted_underline: false,
            heavy_underline: false,
            double_underline: false,
            dot_dash_underline: false,
            double_strikethrough: false,
            shadow_scale_x: 1.0,
            shadow_scale_y: 1.0,
            shadow_skew_x: 0.0,
            shadow_skew_y: 0.0,
            shadow_alignment: 7,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Glow {
    pub color: u32,
    pub radius: f32,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Reflection {
    pub start_opacity: f32,
    pub end_opacity: f32,
    pub start_position: f32,
    pub end_position: f32,
    pub direction_degrees: f32,
    pub blur: f32,
    pub distance: f32,
    pub scale_x: f32,
    pub scale_y: f32,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Bevel3D {
    pub width: f32,
    pub height: f32,
    pub preset: String,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Backdrop3D {
    pub anchor_x: f32,
    pub anchor_y: f32,
    pub anchor_z: f32,
    pub normal_x: f32,
    pub normal_y: f32,
    pub normal_z: f32,
    pub up_x: f32,
    pub up_y: f32,
    pub up_z: f32,
}

impl Default for Backdrop3D {
    fn default() -> Self {
        Self {
            anchor_x: 0.0,
            anchor_y: 0.0,
            anchor_z: 0.0,
            normal_x: 0.0,
            normal_y: 0.0,
            normal_z: 1.0,
            up_x: 0.0,
            up_y: -1.0,
            up_z: 0.0,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct ThreeDStyle {
    pub camera_preset: String,
    pub camera_fov: f32,
    pub camera_zoom: f32,
    /// Resolved OOXML angles: explicit a:rot overrides preset defaults.
    pub camera_latitude: f32,
    pub camera_longitude: f32,
    pub camera_revolution: f32,
    pub light_rig: String,
    pub light_direction: String,
    pub light_latitude: f32,
    pub light_longitude: f32,
    pub light_revolution: f32,
    pub z: f32,
    pub extrusion_height: f32,
    pub contour_width: f32,
    pub material: String,
    pub bevel_top: Option<Bevel3D>,
    pub bevel_bottom: Option<Bevel3D>,
    pub extrusion_color: Option<u32>,
    pub contour_color: Option<u32>,
    pub backdrop: Option<Backdrop3D>,
    /// `a:flatTx` excludes shape text from the 3D scene at the authored Z depth.
    pub flat_text_z: Option<f32>,
    /// The 3D scene was authored on `a:bodyPr` and therefore applies to text.
    pub applies_to_text: bool,
}

impl Default for ThreeDStyle {
    fn default() -> Self {
        Self {
            camera_preset: "orthographicFront".to_owned(),
            camera_fov: 0.0,
            camera_zoom: 1.0,
            camera_latitude: 0.0,
            camera_longitude: 0.0,
            camera_revolution: 0.0,
            light_rig: "threePt".to_owned(),
            light_direction: "t".to_owned(),
            light_latitude: 0.0,
            light_longitude: 0.0,
            light_revolution: 0.0,
            z: 0.0,
            extrusion_height: 0.0,
            contour_width: 0.0,
            material: "warmMatte".to_owned(),
            bevel_top: None,
            bevel_bottom: None,
            extrusion_color: None,
            contour_color: None,
            backdrop: None,
            flat_text_z: None,
            applies_to_text: false,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum LineCap {
    Flat,
    Round,
    Square,
}

impl LineCap {
    pub const fn code(self) -> u8 {
        match self {
            Self::Flat => 0,
            Self::Round => 1,
            Self::Square => 2,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum LineJoin {
    Miter,
    Round,
    Bevel,
}

impl LineJoin {
    pub const fn code(self) -> u8 {
        match self {
            Self::Miter => 0,
            Self::Round => 1,
            Self::Bevel => 2,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum LineCompound {
    Single,
    Double,
    ThickThin,
    ThinThick,
    Triple,
}

impl LineCompound {
    pub const fn code(self) -> u8 {
        match self {
            Self::Single => 0,
            Self::Double => 1,
            Self::ThickThin => 2,
            Self::ThinThick => 3,
            Self::Triple => 4,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum LineAlignment {
    Center,
    Inset,
}

impl LineAlignment {
    pub const fn code(self) -> u8 {
        match self {
            Self::Center => 0,
            Self::Inset => 1,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct StrokeStyle {
    pub cap: LineCap,
    pub join: LineJoin,
    pub compound: LineCompound,
    pub alignment: LineAlignment,
    pub miter_limit: f32,
    pub dash: Vec<f32>,
    pub dash_offset: f32,
}

impl Default for StrokeStyle {
    fn default() -> Self {
        Self {
            cap: LineCap::Flat,
            join: LineJoin::Miter,
            compound: LineCompound::Single,
            alignment: LineAlignment::Center,
            miter_limit: 10.0,
            dash: Vec::new(),
            dash_offset: 0.0,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextDirection {
    Auto,
    Ltr,
    Rtl,
}

impl TextDirection {
    pub const fn code(self) -> u8 {
        match self {
            Self::Auto => 0,
            Self::Ltr => 1,
            Self::Rtl => 2,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextOrientation {
    Horizontal,
    VerticalRl,
    VerticalLr,
    Rotated90,
    Rotated270,
    StackedRl,
    StackedLr,
}

impl TextOrientation {
    pub const fn code(self) -> u8 {
        match self {
            Self::Horizontal => 0,
            Self::VerticalRl => 1,
            Self::VerticalLr => 2,
            Self::Rotated90 => 3,
            Self::Rotated270 => 4,
            Self::StackedRl => 5,
            Self::StackedLr => 6,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextAutoFit {
    None,
    Shrink,
    FitFrame,
}

impl TextAutoFit {
    pub const fn code(self) -> u8 {
        match self {
            Self::None => 0,
            Self::Shrink => 1,
            Self::FitFrame => 2,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextVerticalAlign {
    Top,
    Center,
    Bottom,
}

impl TextVerticalAlign {
    pub const fn code(self) -> u8 {
        match self {
            Self::Top => 0,
            Self::Center => 1,
            Self::Bottom => 2,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextHorizontalOverflow {
    Overflow,
    Clip,
}

impl TextHorizontalOverflow {
    pub const fn code(self) -> u8 {
        match self {
            Self::Overflow => 0,
            Self::Clip => 1,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum TextVerticalOverflow {
    Overflow,
    Clip,
    Ellipsis,
}

impl TextVerticalOverflow {
    pub const fn code(self) -> u8 {
        match self {
            Self::Overflow => 0,
            Self::Clip => 1,
            Self::Ellipsis => 2,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct TextParagraphRule {
    pub color: u32,
    pub stroke_width: f32,
    pub offset_x: f32,
    pub offset_y: f32,
    pub width: f32,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct TextDropCap {
    pub characters: u32,
    pub lines: u32,
    pub raised_lines: u32,
    pub padding: f32,
    pub outdent: f32,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct TextParagraphLayout {
    pub align: TextAlign,
    pub margin_left: f32,
    pub margin_right: f32,
    pub first_line_indent: f32,
    pub default_tab_stop: f32,
    /// Absolute line height for this paragraph; zero inherits the text box line height.
    pub line_height: f32,
    pub space_before: f32,
    pub space_after: f32,
    /// Whether an oversized Latin word may be split at character boundaries.
    pub latin_line_break: bool,
    /// Whether terminal punctuation may extend beyond the paragraph's right edge.
    pub hanging_punctuation: bool,
    /// Authored horizontal rule above this paragraph.
    pub rule_above: Option<TextParagraphRule>,
    pub rule_below: Option<TextParagraphRule>,
    /// The paragraph's leading characters carry their authored drop-cap style.
    pub drop_cap: Option<TextDropCap>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct TextLayout {
    /// Repeat a glyph into spare line width at a UTF-16 source-text offset.
    pub fill_character: Option<(u32, char)>,
    /// Rectangular exclusions relative to the inset text origin, in layout units.
    pub wrap_regions: Vec<Rect>,
    /// Center the font box within fixed-height lines; bottom-align if it is taller.
    pub fixed_line_height: bool,
    /// Permit removal of full-width punctuation whitespace while fitting a line.
    pub compress_punctuation: bool,
    /// The final paragraph continues in another frame; its last visible line
    /// is not a paragraph-ending line for justification.
    pub continues_after: bool,
    pub direction: TextDirection,
    pub orientation: TextOrientation,
    pub auto_fit: TextAutoFit,
    pub vertical_align: TextVerticalAlign,
    /// Literal bullet or numbering prefix.
    pub prefix: Option<String>,
    /// Absolute positions from the text box's leading edge.
    pub tab_stops: Vec<TextTabStop>,
    pub default_tab_stop: f32,
    pub hanging_indent: f32,
    /// Additional vertical distance after an explicit paragraph break.
    pub paragraph_spacing: f32,
    pub inset_left: f32,
    pub inset_right: f32,
    pub inset_top: f32,
    pub inset_bottom: f32,
    pub margin_left: f32,
    pub margin_right: f32,
    pub first_line_indent: f32,
    pub column_count: u32,
    pub column_spacing: f32,
    pub rotation_degrees: f32,
    pub font_scale: f32,
    pub line_spacing_reduction: f32,
    pub horizontal_overflow: TextHorizontalOverflow,
    pub vertical_overflow: TextVerticalOverflow,
    pub wrap: bool,
    pub warp: Option<String>,
    /// Whether glyph interiors are painted. PDF text rendering modes can disable fill.
    pub text_fill: bool,
    /// Optional authored paint for glyph interiors. `None` preserves per-run colors.
    pub text_paint: Paint,
    /// Horizontally scale a single authored line to the source advance width.
    pub text_scale_to_fit: bool,
    /// The source text matrix itself requires horizontal scaling, even when
    /// browser advances came from an exact embedded PDF font.
    pub text_matrix_scale_to_fit: bool,
    /// Requests bounded supersampling when a browser rasterizes this text below 1x.
    pub low_resolution_supersample: bool,
    /// RGBA glyph outline; an alpha byte of zero disables the outline.
    pub text_stroke_color: u32,
    /// Optional authored paint for glyph outlines. `None` preserves the solid color.
    pub text_stroke_paint: Paint,
    pub text_stroke_width: f32,
    /// Absolute glyph baseline from the text box top; zero uses browser metrics.
    pub text_baseline: f32,
    /// Paragraph-local horizontal layout, in source paragraph order.
    pub paragraphs: Vec<TextParagraphLayout>,
    pub min_scale: f32,
}

impl Default for TextLayout {
    fn default() -> Self {
        Self {
            fill_character: None,
            wrap_regions: Vec::new(),
            compress_punctuation: false,
            fixed_line_height: false,
            continues_after: false,
            direction: TextDirection::Auto,
            orientation: TextOrientation::Horizontal,
            auto_fit: TextAutoFit::None,
            vertical_align: TextVerticalAlign::Top,
            prefix: None,
            tab_stops: Vec::new(),
            default_tab_stop: 36.0,
            hanging_indent: 0.0,
            paragraph_spacing: 0.0,
            inset_left: 4.0,
            inset_right: 4.0,
            inset_top: 4.0,
            inset_bottom: 4.0,
            margin_left: 0.0,
            margin_right: 0.0,
            first_line_indent: 0.0,
            column_count: 1,
            column_spacing: 0.0,
            rotation_degrees: 0.0,
            font_scale: 1.0,
            line_spacing_reduction: 0.0,
            horizontal_overflow: TextHorizontalOverflow::Overflow,
            vertical_overflow: TextVerticalOverflow::Overflow,
            wrap: true,
            warp: None,
            text_fill: true,
            text_paint: Paint::None,
            text_scale_to_fit: false,
            text_matrix_scale_to_fit: false,
            low_resolution_supersample: false,
            text_stroke_color: 0,
            text_stroke_paint: Paint::None,
            text_stroke_width: 0.0,
            text_baseline: 0.0,
            paragraphs: Vec::new(),
            min_scale: 0.1,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum Visual {
    None,
    Shape {
        geometry: Geometry,
        fill: u32,
        stroke: u32,
        stroke_width: f32,
    },
    Text {
        geometry: Geometry,
        fill: u32,
        stroke: u32,
        stroke_width: f32,
        font_family: String,
        font_size: f32,
        color: u32,
        bold: bool,
        italic: bool,
        align: TextAlign,
    },
    Image {
        media_type: String,
        bytes: Vec<u8>,
        crop: ImageCrop,
    },
    ImageWithFallback {
        media_type: String,
        bytes: Vec<u8>,
        fallback_media_type: String,
        fallback_bytes: Vec<u8>,
        crop: ImageCrop,
    },
    /// An encoded image whose decoded pixels are modulated by an authored
    /// 8-bit alpha plane of the same dimensions.
    MaskedImage {
        media_type: String,
        bytes: Vec<u8>,
        resource_id: u32,
        mask_resource_id: u32,
        mask_width: u32,
        mask_height: u32,
        mask_media_type: String,
        alpha_mask: Vec<u8>,
        crop: ImageCrop,
    },
    Group {
        children: Vec<VisualBrushChild>,
    },
    OpacityMask {
        mask: Paint,
        visual: Box<Visual>,
    },
    ColorManagedImage {
        source_profile: Vec<u8>,
        destination_profile: Option<Vec<u8>>,
        visual: Box<Visual>,
    },
    /// A locally embedded audio or video asset attached to its authored poster.
    /// The poster remains the static Canvas representation; browser hosts can
    /// expose the bounded media bytes as an interactive playback control.
    Media {
        kind: MediaKind,
        media_type: String,
        bytes: Vec<u8>,
        poster: Box<Visual>,
    },
    /// DrawingML `a:clrChange`: replace exact source pixels with the authored
    /// target color. `use_alpha` controls whether source alpha participates in
    /// matching; the target alpha modulates the matched source pixel's alpha.
    ImageColorChange {
        from: u32,
        to: u32,
        use_alpha: bool,
        visual: Box<Visual>,
    },
    ImageAdjustment {
        adjustment: ImageAdjustment,
        visual: Box<Visual>,
    },
    /// A compositing layer. Leading layers on a group object are inherited by
    /// its descendants; leading layers on any object affect its own visual.
    Layer {
        transform: AffineTransform,
        opacity: f32,
        blend_mode: BlendMode,
        visual: Box<Visual>,
    },
    PaintedShape {
        geometry: Geometry,
        fill: Paint,
        stroke: Paint,
        stroke_width: f32,
    },
    RichText {
        geometry: Geometry,
        fill: Paint,
        stroke: Paint,
        stroke_width: f32,
        align: TextAlign,
        /// Absolute document-space line height. Non-positive values request
        /// the renderer's font-derived default.
        line_height: f32,
        runs: Vec<TextRun>,
    },
    /// Canvas-compatible effects around any visual. Effects may be nested to
    /// preserve shadow/clip ordering.
    Effect {
        shadow: Option<Shadow>,
        clip: Option<Geometry>,
        visual: Box<Visual>,
    },
    /// Paragraph layout controls around a text visual. Keeping this separate
    /// preserves the v3/v4 rich-text defaults for existing format adapters.
    TextLayout {
        layout: TextLayout,
        visual: Box<Visual>,
    },
    TextEffects {
        effects: Vec<TextEffect>,
        visual: Box<Visual>,
    },
    StrokeStyle {
        style: StrokeStyle,
        visual: Box<Visual>,
    },
    AdvancedEffect {
        outer_shadow: Option<OuterShadow>,
        inner_shadow: Option<Shadow>,
        glow: Option<Glow>,
        reflection: Option<Reflection>,
        soft_edge: Option<f32>,
        three_d: Option<ThreeDStyle>,
        visual: Box<Visual>,
    },
}

impl Visual {
    pub const fn code(&self) -> u8 {
        match self {
            Self::None => 0,
            Self::Shape { .. } => 1,
            Self::Text { .. } => 2,
            Self::Image { .. } => 3,
            Self::ImageWithFallback { .. } => 4,
            Self::Layer { .. } => 5,
            Self::PaintedShape { .. } => 6,
            Self::RichText { .. } => 7,
            Self::Effect { .. } => 8,
            Self::TextLayout { .. } => 9,
            Self::StrokeStyle { .. } => 10,
            Self::AdvancedEffect { .. } => 11,
            Self::ImageColorChange { .. } => 12,
            Self::Media { .. } => 13,
            Self::MaskedImage { .. } => 14,
            Self::Group { .. } => 15,
            Self::OpacityMask { .. } => 16,
            Self::ColorManagedImage { .. } => 17,
            Self::ImageAdjustment { .. } => 18,
            Self::TextEffects { .. } => 19,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MediaKind {
    Audio,
    Video,
}

impl MediaKind {
    pub const fn code(self) -> u8 {
        match self {
            Self::Audio => 0,
            Self::Video => 1,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Object {
    pub numeric_id: u32,
    pub parent_numeric_id: Option<u32>,
    pub stable_id: String,
    pub parent_stable_id: Option<String>,
    pub kind: ObjectKind,
    pub unit_index: u32,
    pub bounds: Rect,
    pub z: i32,
    pub text: Option<String>,
    pub source: SourceRef,
    pub visual: Visual,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum FontStyle {
    Normal,
    Italic,
    Oblique,
}

impl FontStyle {
    pub const fn code(self) -> u8 {
        match self {
            Self::Normal => 0,
            Self::Italic => 1,
            Self::Oblique => 2,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct EmbeddedFont {
    pub family: String,
    pub bytes: Vec<u8>,
    pub style: FontStyle,
    pub weight: u16,
}

#[derive(Clone, Debug, PartialEq)]
pub struct OutlineItem {
    pub title: String,
    pub unit_index: u32,
    pub level: u32,
}

#[derive(Clone, Debug)]
pub struct Document {
    pub fatal: bool,
    pub format: Option<DocumentFormat>,
    pub kind: Option<DocumentKind>,
    pub units: Vec<Unit>,
    pub outline: Vec<OutlineItem>,
    pub objects: Vec<Object>,
    pub embedded_fonts: Vec<EmbeddedFont>,
    /// Ordered lookup candidates; authored names remain intact. Unlike legacy
    /// FFN canonical-name normalization, availability is resolved by the host.
    pub font_alternate_names: Vec<(String, Vec<String>)>,
    pub diagnostics: Vec<Diagnostic>,
}

impl Document {
    pub fn fatal(diagnostic: Diagnostic) -> Self {
        Self {
            fatal: true,
            format: None,
            kind: None,
            units: Vec::new(),
            outline: Vec::new(),
            objects: Vec::new(),
            embedded_fonts: Vec::new(),
            font_alternate_names: Vec::new(),
            diagnostics: vec![diagnostic],
        }
    }

    pub fn hit_test(&self, unit_index: u32, x: f32, y: f32, limit: usize) -> Vec<u32> {
        let limit = limit.min(256);
        if limit == 0 {
            return Vec::new();
        }
        let dense_numeric_ids = self
            .objects
            .iter()
            .enumerate()
            .all(|(index, object)| usize::try_from(object.numeric_id) == Ok(index));
        let mut hits = BinaryHeap::with_capacity(limit);
        for object in self.objects.iter().filter(|object| {
            object.unit_index == unit_index
                && object_is_visible(self, object, dense_numeric_ids)
                && transformed_contains(self, object, dense_numeric_ids, x, y)
        }) {
            let candidate = (
                object.z,
                depth(self, object, dense_numeric_ids),
                object.numeric_id,
            );
            if hits.len() < limit {
                hits.push(Reverse(candidate));
            } else if hits.peek().is_some_and(|worst| candidate > worst.0) {
                hits.pop();
                hits.push(Reverse(candidate));
            }
        }

        let mut hits: Vec<_> = hits.into_iter().map(|Reverse(hit)| hit).collect();
        hits.sort_unstable_by(|left, right| right.cmp(left));
        hits.into_iter()
            .map(|(_, _, numeric_id)| numeric_id)
            .collect()
    }
}

fn object_is_visible(document: &Document, object: &Object, dense_numeric_ids: bool) -> bool {
    let mut current = Some(object);
    let mut depth = 0_usize;
    while let Some(object) = current {
        if matches!(
            &object.source.locator,
            SourceLocator::PptxShape { metadata, .. } if metadata.hidden
        ) {
            return false;
        }
        if depth >= 128 {
            return false;
        }
        depth += 1;
        current = object
            .parent_numeric_id
            .and_then(|numeric_id| object_by_numeric_id(document, numeric_id, dense_numeric_ids));
    }
    true
}

fn object_by_numeric_id(
    document: &Document,
    numeric_id: u32,
    dense_numeric_ids: bool,
) -> Option<&Object> {
    if dense_numeric_ids {
        usize::try_from(numeric_id)
            .ok()
            .and_then(|index| document.objects.get(index))
    } else {
        document
            .objects
            .iter()
            .find(|candidate| candidate.numeric_id == numeric_id)
    }
}

pub(crate) fn leading_transform(visual: &Visual) -> AffineTransform {
    let mut transform = AffineTransform::IDENTITY;
    let mut visual = visual;
    for _ in 0..=64 {
        match visual {
            Visual::Layer {
                transform: layer,
                visual: nested,
                ..
            } => {
                transform = transform.concat(*layer);
                visual = nested;
            }
            Visual::Effect { visual: nested, .. }
            | Visual::TextLayout { visual: nested, .. }
            | Visual::TextEffects { visual: nested, .. }
            | Visual::StrokeStyle { visual: nested, .. }
            | Visual::AdvancedEffect { visual: nested, .. }
            | Visual::ImageColorChange { visual: nested, .. }
            | Visual::ImageAdjustment { visual: nested, .. } => visual = nested,
            Visual::Media { poster, .. } => visual = poster,
            _ => break,
        }
    }
    transform
}

fn transformed_contains(
    document: &Document,
    object: &Object,
    dense_numeric_ids: bool,
    x: f32,
    y: f32,
) -> bool {
    let mut ancestors: [Option<&Object>; 128] = [None; 128];
    let mut ancestor_count = 0;
    let mut parent = object.parent_numeric_id;
    while let Some(parent_id) = parent {
        if ancestor_count == ancestors.len() {
            return false;
        }
        let Some(parent_object) = object_by_numeric_id(document, parent_id, dense_numeric_ids)
        else {
            return false;
        };
        ancestors[ancestor_count] = Some(parent_object);
        ancestor_count += 1;
        parent = parent_object.parent_numeric_id;
    }
    let mut transform = AffineTransform::IDENTITY;
    for ancestor in ancestors[..ancestor_count].iter().rev().flatten() {
        if ancestor.kind == ObjectKind::Group {
            transform = transform.concat(leading_transform(&ancestor.visual));
        }
    }
    transform = transform.concat(leading_transform(&object.visual));
    transform
        .inverse_transform_point(x, y)
        .is_some_and(|(local_x, local_y)| object.bounds.contains(local_x, local_y))
}

fn depth(document: &Document, object: &Object, dense_numeric_ids: bool) -> usize {
    let mut depth = 0;
    let mut parent = object.parent_numeric_id;
    while let Some(parent_id) = parent {
        if depth >= 128 {
            break;
        }
        parent = object_by_numeric_id(document, parent_id, dense_numeric_ids)
            .and_then(|candidate| candidate.parent_numeric_id);
        depth += 1;
    }
    depth
}

#[cfg(test)]
mod tests {
    #[test]
    fn sheet_extents_cover_hidden_and_nonuniform_axes_without_extra_boundary_cells() {
        let axis = super::SheetAxis::from_sizes(&[10.0, 0.0, 0.0, 20.0, 10.0], 10.0);
        for (extent, expected) in [
            (0.0, 0),
            (10.0, 1),
            (10.1, 4),
            (30.0, 4),
            (40.0, 5),
            (41.0, 6),
            (1000.0, 8),
        ] {
            assert_eq!(
                super::sheet_axis_count_for_extent(extent, 8, |index| axis.offset(index)),
                expected
            );
        }
    }

    #[test]
    fn sheet_axis_offsets_preserve_hidden_spans_and_boundaries() {
        let axis = super::SheetAxis::from_sizes(&[10.0, 0.0, 0.0, 20.0, 10.0], 10.0);
        for (index, expected) in [0.0, 10.0, 10.0, 10.0, 30.0, 40.0].into_iter().enumerate() {
            assert_eq!(axis.offset(index as u32), expected);
        }
        let axis = super::SheetAxis {
            default_size: 0.0,
            spans: vec![super::SheetAxisSpan {
                start: u32::MAX - 1,
                end: u32::MAX,
                size: 2.0,
            }],
        };
        assert_eq!(axis.offset(u32::MAX), 2.0);
    }

    use super::{
        AffineTransform, BlendMode, Document, DocumentFormat, MappingQuality, Object, ObjectKind,
        PptxObjectMetadata, Rect, SourceLocator, SourceRef, TextLayout, Visual,
    };

    #[test]
    fn dashed_paths_do_not_start_a_dash_at_the_terminal_boundary() {
        use super::{PathCommand, append_dashed_polyline};
        for length in [50.0, 50.0 - 0.00001, 50.0 + 0.00001] {
            let mut commands = Vec::new();
            append_dashed_polyline(
                &mut commands,
                &[(0.0, 0.0), (length, 0.0)],
                &[3.0, 7.0],
                0.0,
            );
            assert_eq!(commands.len(), 10, "no terminal dot at {length}");
            assert_eq!(
                commands.last(),
                Some(&PathCommand::LineTo { x: 43.0, y: 0.0 })
            );
        }
        let mut commands = Vec::new();
        append_dashed_polyline(&mut commands, &[(0.0, 0.0), (50.01, 0.0)], &[3.0, 7.0], 0.0);
        assert_eq!(commands.len(), 12, "a real partial dash must remain");
        let mut commands = Vec::new();
        append_dashed_polyline(
            &mut commands,
            &[(0.0, 0.0), (5.0, 0.0), (10.0, 0.0)],
            &[3.0, 7.0],
            1.0,
        );
        assert_eq!(
            commands.last(),
            Some(&PathCommand::LineTo { x: 10.0, y: 0.0 })
        );
    }

    #[test]
    fn shared_affine_path_and_dash_phase_preserve_geometry() {
        use super::{PathCommand, append_dashed_polyline};
        let mut command = PathCommand::BezierCurveTo {
            cp1x: 1.0,
            cp1y: 2.0,
            cp2x: 3.0,
            cp2y: 4.0,
            x: 5.0,
            y: 6.0,
        };
        command.transform(AffineTransform {
            a: 2.0,
            b: 1.0,
            c: -1.0,
            d: 3.0,
            e: 7.0,
            f: -2.0,
        });
        assert_eq!(
            command,
            PathCommand::BezierCurveTo {
                cp1x: 7.0,
                cp1y: 5.0,
                cp2x: 9.0,
                cp2y: 13.0,
                x: 11.0,
                y: 21.0,
            }
        );
        let mut commands = Vec::new();
        append_dashed_polyline(&mut commands, &[(0.0, 0.0), (10.0, 0.0)], &[3.0, 2.0], 1.0);
        assert_eq!(
            &commands[..4],
            &[
                PathCommand::MoveTo { x: 0.0, y: 0.0 },
                PathCommand::LineTo { x: 2.0, y: 0.0 },
                PathCommand::MoveTo { x: 4.0, y: 0.0 },
                PathCommand::LineTo { x: 7.0, y: 0.0 },
            ]
        );
    }

    #[test]
    fn localizes_absolute_gradients_without_changing_relative_paints() {
        use super::{GradientSpread, Paint};
        let linear = Paint::LinearGradient {
            x0: 5.0,
            y0: 7.0,
            x1: 15.0,
            y1: 27.0,
            stops: vec![],
        };
        assert_eq!(
            linear.localize_gradient(3.0, 4.0),
            Paint::LinearGradient {
                x0: 2.0,
                y0: 3.0,
                x1: 12.0,
                y1: 23.0,
                stops: vec![]
            }
        );
        let radial = Paint::RadialGradient {
            x0: 5.0,
            y0: 7.0,
            r0: 2.0,
            x1: 15.0,
            y1: 27.0,
            r1: 9.0,
            stops: vec![],
        };
        assert_eq!(
            radial.localize_gradient(3.0, 4.0),
            Paint::RadialGradient {
                x0: 2.0,
                y0: 3.0,
                r0: 2.0,
                x1: 12.0,
                y1: 23.0,
                r1: 9.0,
                stops: vec![]
            }
        );
        for relative in [false, true] {
            let paint = Paint::XpsGradient {
                radial: true,
                start_x: 5.0,
                start_y: 7.0,
                end_x: 15.0,
                end_y: 27.0,
                radius_x: 2.0,
                radius_y: 9.0,
                relative,
                spread: GradientSpread::Pad,
                linear_rgb: false,
                transform: AffineTransform::IDENTITY,
                relative_transform: AffineTransform::IDENTITY,
                stops: vec![],
            };
            let mut expected = paint.clone();
            if let Paint::XpsGradient {
                start_x,
                start_y,
                end_x,
                end_y,
                ..
            } = &mut expected
            {
                if !relative {
                    *start_x = 2.0;
                    *start_y = 3.0;
                    *end_x = 12.0;
                    *end_y = 23.0;
                }
            }
            assert_eq!(paint.localize_gradient(3.0, 4.0), expected);
        }
        assert_eq!(
            Paint::Solid(0xff).localize_gradient(3.0, 4.0),
            Paint::Solid(0xff)
        );
    }

    #[test]
    fn remaining_flat_formats_keep_their_existing_wire_codes() {
        assert_eq!(DocumentFormat::Csv.code(), 8);
        assert_eq!(DocumentFormat::Rtf.code(), 10);
    }

    #[test]
    fn optional_native_formats_have_stable_non_retired_wire_codes() {
        assert_eq!(DocumentFormat::Ppt.code(), 11);
        assert_eq!(DocumentFormat::Xls.code(), 12);
        assert_eq!(DocumentFormat::Doc.code(), 13);
        assert_eq!(DocumentFormat::Pages.code(), 14);
        assert_eq!(DocumentFormat::Numbers.code(), 15);
        assert_eq!(DocumentFormat::Keynote.code(), 16);
        assert_eq!(DocumentFormat::Pdf.code(), 17);
    }

    #[test]
    fn hit_test_caps_results_at_the_public_protocol_limit() {
        let document = document(
            (0..300)
                .map(|numeric_id| object(numeric_id, None, numeric_id as i32))
                .collect(),
        );

        let hits = document.hit_test(0, 5.0, 5.0, 300);

        assert_eq!(hits.len(), 256);
        assert_eq!(hits.first(), Some(&299));
        assert_eq!(hits.last(), Some(&44));
    }

    #[test]
    fn hit_test_orders_by_z_then_depth_then_numeric_id() {
        let document = document(vec![
            object(0, None, 5),
            object(1, Some(0), 5),
            object(2, Some(1), 5),
            object(3, None, 6),
            object(4, Some(1), 5),
            object(5, None, 4),
        ]);

        assert_eq!(
            document.hit_test(0, 5.0, 5.0, usize::MAX),
            vec![3, 4, 2, 1, 0, 5]
        );
    }

    #[test]
    fn hit_test_preserves_depth_order_for_sparse_numeric_ids() {
        let document = document(vec![
            object(100, None, 5),
            object(200, Some(100), 5),
            object(300, None, 5),
        ]);

        assert_eq!(document.hit_test(0, 5.0, 5.0, 3), vec![200, 300, 100]);
    }

    #[test]
    fn hit_test_applies_inherited_group_affine_transforms() {
        let mut group = object(0, None, 0);
        group.kind = ObjectKind::Group;
        group.bounds = Rect::default();
        group.visual = Visual::Layer {
            transform: AffineTransform {
                e: 100.0,
                f: 20.0,
                ..AffineTransform::IDENTITY
            },
            opacity: 0.5,
            blend_mode: BlendMode::Normal,
            visual: Box::new(Visual::None),
        };
        let child = object(1, Some(0), 1);
        let document = document(vec![group, child]);

        assert!(document.hit_test(0, 5.0, 5.0, 10).is_empty());
        assert_eq!(document.hit_test(0, 105.0, 25.0, 10), vec![1]);
    }

    #[test]
    fn hit_test_traverses_text_wrappers_before_affine_transforms() {
        let mut wrapped = object(0, None, 1);
        wrapped.visual = Visual::TextLayout {
            layout: TextLayout::default(),
            visual: Box::new(Visual::TextEffects {
                effects: Vec::new(),
                visual: Box::new(Visual::Layer {
                    transform: AffineTransform {
                        e: 100.0,
                        f: 20.0,
                        ..AffineTransform::IDENTITY
                    },
                    opacity: 1.0,
                    blend_mode: BlendMode::Normal,
                    visual: Box::new(Visual::None),
                }),
            }),
        };
        let document = document(vec![wrapped]);

        assert!(document.hit_test(0, 5.0, 5.0, 10).is_empty());
        assert_eq!(document.hit_test(0, 105.0, 25.0, 10), vec![0]);
    }

    #[test]
    fn image_adjustment_and_text_effects_have_distinct_protocol_codes() {
        let image_adjustment = Visual::ImageAdjustment {
            adjustment: super::ImageAdjustment::default(),
            visual: Box::new(Visual::None),
        };
        let text_effects = Visual::TextEffects {
            effects: vec![super::TextEffect::default()],
            visual: Box::new(Visual::None),
        };

        assert_ne!(image_adjustment.code(), text_effects.code());
    }

    fn document(objects: Vec<Object>) -> Document {
        Document {
            fatal: false,
            format: None,
            kind: None,
            units: Vec::new(),
            outline: Vec::new(),
            objects,
            embedded_fonts: Vec::new(),
            font_alternate_names: Vec::new(),
            diagnostics: Vec::new(),
        }
    }

    fn object(numeric_id: u32, parent_numeric_id: Option<u32>, z: i32) -> Object {
        Object {
            numeric_id,
            parent_numeric_id,
            stable_id: format!("object:{numeric_id}"),
            parent_stable_id: parent_numeric_id.map(|parent| format!("object:{parent}")),
            kind: ObjectKind::Shape,
            unit_index: 0,
            bounds: Rect {
                x: 0.0,
                y: 0.0,
                width: 10.0,
                height: 10.0,
            },
            z,
            text: None,
            source: SourceRef {
                part: "test.xml".to_owned(),
                mapping: MappingQuality::Exact,
                locator: SourceLocator::PptxShape {
                    shape_id: numeric_id,
                    row: None,
                    column: None,
                    text_range: None,
                    metadata: PptxObjectMetadata::default(),
                },
            },
            visual: Visual::None,
        }
    }
}
