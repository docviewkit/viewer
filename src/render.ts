import { byteFingerprint, bytesEqual } from "./bytes.js";
import {
  multiplyTransform,
  type SceneAffineTransform,
  type SceneGeometry,
  type SceneObject,
  type ScenePaint,
  type ScenePathCommand,
  type SceneReflection,
  type SceneStrokeStyle,
  type SceneTextAlign,
  type SceneTextLayout,
  type SceneTextEffect,
  type SceneTextParagraphLayout,
  type SceneTextRun,
  type SceneTextTabStop,
  type SceneThreeDStyle,
  type SceneVisual,
  type SceneXpsColor,
} from "./scene.js";
import type {
  Diagnostic,
  DocumentFontRun,
  Rect,
  RenderRequest,
  RenderResult,
  RenderedTextFragment,
  ResourceLimits,
  SheetAxis,
  SheetSizeOverride,
  UnitDescriptor,
  Viewport,
} from "./types.js";
import { OfficeEngineError } from "./types.js";
import { identifyImage, reducedJpeg2000Dimensions } from "./image.js";
import { decodeOfficeImage, type DecodedOfficeImage } from "./image-codec-client.js";
import type { OfficeImageCodecFont } from "./image-codecs.js";
import {
  approximateFontFamily,
  compatibleFallbackFamily,
  fallbackBaselineShift,
  fallbackFontSizeFactor,
  fallbackSpaceAdvanceEm,
  fontFallbackTextSegments,
  fontShorthand,
  probeBrowserFontFace,
  semanticSymbolFontText,
  symbolFontGlyphText,
  type FontResolution,
  type FontProbeContext,
  type RuntimeFontResolver,
  type RuntimeFontSet,
} from "./font.js";

const IMAGE_DECODE_TIMEOUT_MS = 10_000;
const IMAGE_EFFECT_CHUNK_PIXELS = 1_048_576;
const IMAGE_PREFETCH_CONCURRENCY = 8;
const IMAGE_PREFETCH_PIXEL_BUDGET = 16_000_000;
const MAX_FONT_RESOLUTION_CODE_POINTS = 65_536;
const MAX_VISUAL_DEPTH = 64;
const LCMS_FLAGS_COPY_ALPHA = 0x0400_0000;
const LCMS_TYPE_RGBA_8 = (4 << 16) | (1 << 7) | (3 << 3) | 1;
const RERASTERIZED_VECTOR_FORMATS = new Set(["svg", "emf", "wmf", "emz", "wmz"]);
const RERASTERIZED_VECTOR_MEDIA_TYPES = new Set([
  "image/svg+xml",
  "image/x-emf",
  "image/x-wmf",
  "image/x-emz",
  "image/x-wmz",
]);

function fallbackGlyphText(text: string, fontFamily: string): string {
  const semantic = semanticSymbolFontText(text, fontFamily);
  return fontFamily.trim().replace(/^(['"])(.*)\1$/u, "$2").toLowerCase() === "algerian"
    ? semantic.replace(/[a-z]/gu, (character) => character.toUpperCase())
    : semantic;
}

class SheetAxisMapper {
  readonly #axis: SheetAxis;
  readonly #count: number;
  readonly #overrides: readonly SheetSizeOverride[];
  readonly #overrideSizes: ReadonlyMap<number, number>;
  readonly #overrideDeltas: readonly number[];
  readonly #spanDeltas: readonly number[];

  constructor(axis: SheetAxis, count: number, overrides: readonly SheetSizeOverride[] = []) {
    const sorted = [...overrides].sort((left, right) => left.index - right.index);
    for (let index = 0; index < sorted.length; index += 1) {
      const override = sorted[index]!;
      if (!Number.isInteger(override.index) || override.index < 0 || override.index >= count
        || !Number.isFinite(override.size) || override.size <= 0
        || sorted[index - 1]?.index === override.index) {
        throw new OfficeEngineError("INVALID_ARGUMENT", "Sheet size overrides must have unique valid indexes and positive sizes");
      }
    }
    this.#axis = axis;
    this.#count = count;
    this.#overrides = sorted;
    this.#overrideSizes = new Map(sorted.map((override) => [override.index, override.size]));
    let overrideDelta = 0;
    this.#overrideDeltas = [0, ...sorted.map((override) => (
      overrideDelta += override.size - this.baseSize(override.index)
    ))];
    let spanDelta = 0;
    this.#spanDeltas = [0, ...axis.spans.map((span) => (
      spanDelta += (span.end - span.start + 1) * (span.size - axis.defaultSize)
    ))];
  }

  size(index: number): number {
    return this.#overrideSizes.get(index) ?? this.baseSize(index);
  }

  baseSize(index: number): number {
    let low = 0;
    let high = this.#axis.spans.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      if (this.#axis.spans[middle]!.start <= index) low = middle + 1;
      else high = middle;
    }
    const span = this.#axis.spans[low - 1];
    return span !== undefined && index <= span.end ? span.size : this.#axis.defaultSize;
  }

  baseOffset(index: number): number {
    let low = 0;
    let high = this.#axis.spans.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      if (this.#axis.spans[middle]!.start < index) low = middle + 1;
      else high = middle;
    }
    let delta = this.#spanDeltas[low] ?? 0;
    const span = this.#axis.spans[low - 1];
    if (span !== undefined && index <= span.end) {
      delta -= (span.end - index + 1) * (span.size - this.#axis.defaultSize);
    }
    return index * this.#axis.defaultSize + delta;
  }

  offset(index: number): number {
    let low = 0;
    let high = this.#overrides.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      if (this.#overrides[middle]!.index < index) low = middle + 1;
      else high = middle;
    }
    return this.baseOffset(index) + (this.#overrideDeltas[low] ?? 0);
  }

  map(value: number): number {
    if (value <= 0) return value;
    const baseTotal = this.baseOffset(this.#count);
    if (value >= baseTotal) return this.offset(this.#count) + value - baseTotal;
    let low = 0;
    let high = this.#count;
    while (low < high) {
      const middle = Math.floor((low + high + 1) / 2);
      if (this.baseOffset(middle) <= value) low = middle;
      else high = middle - 1;
    }
    const start = this.baseOffset(low);
    const baseSize = this.baseSize(low);
    return this.offset(low) + (baseSize <= 0 ? 0 : (value - start) * this.size(low) / baseSize);
  }

  indexAt(value: number): number {
    if (this.#count === 0) return -1;
    let low = 0;
    let high = this.#count;
    while (low < high) {
      const middle = Math.floor((low + high + 1) / 2);
      if (this.offset(middle) <= value) low = middle;
      else high = middle - 1;
    }
    return Math.min(this.#count - 1, Math.max(0, low));
  }

  get total(): number { return this.offset(this.#count); }
}

const SHEET_CSS_PIXELS_PER_INCH = 96;
const SHEET_DEFAULT_MARGINS = Object.freeze({ left: .7, right: .7, top: .75, bottom: .75 });
const SHEET_PAPER_INCHES: Readonly<Record<number, readonly [number, number]>> = Object.freeze({
  1: [8.5, 11],
  3: [11, 17],
  5: [8.5, 14],
  8: [11.6929, 16.5354],
  9: [8.2677, 11.6929],
  11: [5.8268, 8.2677],
  13: [7.1654, 10.1181],
});

export interface SheetPrintPage {
  readonly viewport: Rect;
  readonly width: number;
  readonly height: number;
  readonly fragments: readonly Readonly<{ viewport: Rect; x: number; y: number }>[];
  readonly paper: Readonly<{ width: number; height: number }>;
  readonly margins: Readonly<{ left: number; right: number; top: number; bottom: number }>;
  readonly scale: number;
}

function sheetRepeatedBand(
  repeat: readonly [number, number] | undefined,
  start: number,
  end: number,
): readonly [number, number] | undefined {
  if (repeat === undefined) return undefined;
  if (repeat[0] >= end) return repeat;
  const last = Math.min(start, repeat[1]);
  return last > repeat[0] ? [repeat[0], last] : undefined;
}

function sheetPageSpans(
  mapper: SheetAxisMapper,
  start: number,
  end: number,
  capacity: number,
  breaks: ReadonlySet<number>,
  repeat?: readonly [number, number],
): readonly (readonly [number, number])[] {
  const spans: (readonly [number, number])[] = [];
  let pageStart = mapper.offset(start);
  let position = pageStart;
  const endPosition = mapper.offset(end);
  const reserved = (): number => {
    const band = sheetRepeatedBand(repeat, pageStart, endPosition);
    return band === undefined ? 0 : band[1] - band[0];
  };
  let titleSize = reserved();
  for (let index = start; index < end; index += 1) {
    const size = mapper.size(index);
    if (position > pageStart && (breaks.has(index) || position + size - pageStart > Math.max(1, capacity - titleSize) + .01)) {
      spans.push([pageStart, position]);
      pageStart = position;
      titleSize = reserved();
    }
    position += size;
  }
  if (position > pageStart || spans.length === 0) spans.push([pageStart, position]);
  return spans;
}

function sheetColumnIndex(label: string): number {
  let value = 0;
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64;
  return value - 1;
}

function sheetPrintRange(unit: Extract<UnitDescriptor, { type: "sheet" }>): Readonly<{
  startColumn: number;
  endColumn: number;
  startRow: number;
  endRow: number;
}> {
  const formula = unit.printSettings?.printArea;
  const area = formula?.includes(",") === true ? undefined : formula?.split("!").pop();
  const match = area?.match(/^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/iu);
  if (match === undefined || match === null) {
    return { startColumn: 0, endColumn: unit.columns, startRow: 0, endRow: unit.rows };
  }
  const startColumn = Math.max(0, Math.min(unit.columns, sheetColumnIndex(match[1] ?? "A")));
  const endColumn = Math.max(startColumn + 1, Math.min(unit.columns, sheetColumnIndex(match[3] ?? "A") + 1));
  const startRow = Math.max(0, Math.min(unit.rows, Number(match[2]) - 1));
  const endRow = Math.max(startRow + 1, Math.min(unit.rows, Number(match[4])));
  return { startColumn, endColumn, startRow, endRow };
}

/** Derives Excel print pages without changing the worksheet coordinate system. */
export function sheetPrintPages(
  unit: Extract<UnitDescriptor, { type: "sheet" }>,
  sizes: NonNullable<RenderRequest["sheetSizes"]> = {},
): readonly SheetPrintPage[] {
  const settings = unit.printSettings;
  let [paperWidth, paperHeight] = SHEET_PAPER_INCHES[settings?.paperSize ?? 1] ?? SHEET_PAPER_INCHES[1]!;
  if (settings?.orientation === "landscape") [paperWidth, paperHeight] = [paperHeight, paperWidth];
  const paper = Object.freeze({
    width: paperWidth * SHEET_CSS_PIXELS_PER_INCH,
    height: paperHeight * SHEET_CSS_PIXELS_PER_INCH,
  });
  const margins = Object.freeze({
    left: (settings?.margins.left ?? SHEET_DEFAULT_MARGINS.left) * SHEET_CSS_PIXELS_PER_INCH,
    right: (settings?.margins.right ?? SHEET_DEFAULT_MARGINS.right) * SHEET_CSS_PIXELS_PER_INCH,
    top: (settings?.margins.top ?? SHEET_DEFAULT_MARGINS.top) * SHEET_CSS_PIXELS_PER_INCH,
    bottom: (settings?.margins.bottom ?? SHEET_DEFAULT_MARGINS.bottom) * SHEET_CSS_PIXELS_PER_INCH,
  });
  const printableWidth = Math.max(1, paper.width - margins.left - margins.right);
  const printableHeight = Math.max(1, paper.height - margins.top - margins.bottom);
  const rows = new SheetAxisMapper(unit.rowAxis, unit.rows, sizes.rows);
  const columns = new SheetAxisMapper(unit.columnAxis, unit.columns, sizes.columns);
  const range = sheetPrintRange(unit);
  const areaWidth = columns.offset(range.endColumn) - columns.offset(range.startColumn);
  const areaHeight = rows.offset(range.endRow) - rows.offset(range.startRow);
  let scale = Math.min(4, Math.max(.1, (settings?.scale ?? 100) / 100));
  if (settings?.fitToPage === true) {
    const fitWidth = settings.fitToWidth ?? 0;
    const fitHeight = settings.fitToHeight ?? 0;
    const fitted = [
      fitWidth > 0 && areaWidth > 0 ? printableWidth * fitWidth / areaWidth : Infinity,
      fitHeight > 0 && areaHeight > 0 ? printableHeight * fitHeight / areaHeight : Infinity,
    ];
    const target = Math.min(...fitted.filter(Number.isFinite));
    if (Number.isFinite(target)) scale = Math.min(4, Math.max(.1, target));
  }
  const titles = settings?.printTitles ?? "";
  const rowTitle = titles.match(/!\$?(\d+):\$?(\d+)(?=,|$)/u);
  const columnTitle = titles.match(/!\$?([A-Z]+):\$?([A-Z]+)(?=,|$)/iu);
  const repeatRange = (mapper: SheetAxisMapper, start: number, end: number, count: number): readonly [number, number] | undefined =>
    start >= 0 && end > start && end <= count ? [mapper.offset(start), mapper.offset(end)] : undefined;
  const repeatRows = rowTitle === null ? undefined : repeatRange(rows, Number(rowTitle[1]) - 1, Number(rowTitle[2]), unit.rows);
  const repeatColumns = columnTitle === null ? undefined : repeatRange(columns, sheetColumnIndex(columnTitle[1]!), sheetColumnIndex(columnTitle[2]!) + 1, unit.columns);
  const boundaries = (breaks: NonNullable<typeof settings>["rowBreaks"], start: number, end: number): ReadonlySet<number> =>
    new Set(settings?.fitToPage === true ? [] : (breaks ?? []).filter(([, min, max]) => min < end && max >= start).map(([id]) => id));
  const columnPages = sheetPageSpans(columns, range.startColumn, range.endColumn, printableWidth / scale,
    boundaries(settings?.columnBreaks, range.startRow, range.endRow), repeatColumns);
  return columnPages.flatMap(([left, right]) => {
    const rowPages = sheetPageSpans(rows, range.startRow, range.endRow, printableHeight / scale,
      boundaries(settings?.rowBreaks, columns.indexAt(left), columns.indexAt(Math.max(left, right - .01)) + 1), repeatRows);
    return rowPages.map(([top, bottom]) => {
      const rowTitle = sheetRepeatedBand(repeatRows, top, rows.offset(range.endRow));
      const columnTitle = sheetRepeatedBand(repeatColumns, left, columns.offset(range.endColumn));
      const rowBands = rowTitle === undefined ? [[top, bottom] as const] : [rowTitle, [top, bottom] as const];
      const columnBands = columnTitle === undefined ? [[left, right] as const] : [columnTitle, [left, right] as const];
      const fragments: { viewport: Rect; x: number; y: number }[] = [];
      let y = 0, width = 0;
      for (const [y1, y2] of rowBands) {
        let x = 0;
        for (const [x1, x2] of columnBands) {
          fragments.push({ viewport: { x: x1, y: y1, width: x2 - x1, height: y2 - y1 }, x, y });
          x += x2 - x1;
        }
        width = x;
        y += y2 - y1;
      }
      return Object.freeze({
        viewport: Object.freeze({ x: left, y: top, width: right - left, height: bottom - top }),
        width, height: y, fragments, paper, margins, scale,
      });
    });
  });
}

export function sheetFormulaViewportBounds(
  unit: UnitDescriptor,
  request: Pick<RenderRequest, "viewport" | "sheetSizes">,
): { readonly maxRow: number; readonly maxColumn: number } | undefined {
  if (unit.type !== "sheet" || request.viewport === undefined) return undefined;
  const rows = new SheetAxisMapper(unit.rowAxis, unit.rows, request.sheetSizes?.rows);
  const columns = new SheetAxisMapper(unit.columnAxis, unit.columns, request.sheetSizes?.columns);
  const bottom = request.viewport.y + request.viewport.height;
  const right = request.viewport.x + request.viewport.width;
  // Two cells of look-ahead preserve adjacent overflow while keeping the
  // dependency roots proportional to the visible rectangle.
  return {
    maxRow: Math.min(unit.rows, rows.indexAt(bottom) + 3),
    maxColumn: Math.min(unit.columns, columns.indexAt(right) + 3),
  };
}

function sheetMappers(unit: UnitDescriptor, request: RenderRequest): {
  readonly rows: SheetAxisMapper;
  readonly columns: SheetAxisMapper;
} | undefined {
  if (unit.type !== "sheet" || request.sheetSizes === undefined) return undefined;
  return {
    rows: new SheetAxisMapper(unit.rowAxis, unit.rows, request.sheetSizes.rows),
    columns: new SheetAxisMapper(unit.columnAxis, unit.columns, request.sheetSizes.columns),
  };
}

function drawSheetGridLines(
  context: OffscreenCanvasRenderingContext2D,
  unit: Extract<UnitDescriptor, { type: "sheet" }>,
  viewport: Viewport,
  rows: SheetAxisMapper,
  columns: SheetAxisMapper,
): void {
  if (!unit.showGridLines) return;
  const left = Math.max(0, viewport.x);
  const top = Math.max(0, viewport.y);
  const right = Math.min(columns.total, viewport.x + viewport.width);
  const bottom = Math.min(rows.total, viewport.y + viewport.height);
  if (left >= right || top >= bottom) return;

  const transform = context.getTransform?.() ?? { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const pixelWidth = transform.a > 0 ? 1 / transform.a : 1;
  const pixelHeight = transform.d > 0 ? 1 / transform.d : 1;
  const snapX = (value: number): number => transform.a > 0
    ? (Math.round(value * transform.a + transform.e) - transform.e) / transform.a
    : value;
  const snapY = (value: number): number => transform.d > 0
    ? (Math.round(value * transform.d + transform.f) - transform.f) / transform.d
    : value;
  context.fillStyle = color(0xd0d0_d0ff);

  let firstColumn = columns.indexAt(left);
  if (columns.offset(firstColumn) < left) firstColumn += 1;
  const lastColumn = Math.min(unit.columns, columns.indexAt(Math.max(left, right - Number.EPSILON)) + 1);
  let previousX = Number.NaN;
  for (let column = firstColumn; column <= lastColumn; column += 1) {
    const boundary = columns.offset(column);
    const x = Math.max(left, snapX(boundary) - (boundary > left ? pixelWidth : 0));
    if (x >= right || x === previousX) continue;
    context.fillRect(x, top, Math.min(pixelWidth, right - x), bottom - top);
    previousX = x;
  }

  let firstRow = rows.indexAt(top);
  if (rows.offset(firstRow) < top) firstRow += 1;
  const lastRow = Math.min(unit.rows, rows.indexAt(Math.max(top, bottom - Number.EPSILON)) + 1);
  let previousY = Number.NaN;
  for (let row = firstRow; row <= lastRow; row += 1) {
    const boundary = rows.offset(row);
    const y = Math.max(top, snapY(boundary) - (boundary > top ? pixelHeight : 0));
    if (y >= bottom || y === previousY) continue;
    context.fillRect(left, y, right - left, Math.min(pixelHeight, bottom - y));
    previousY = y;
  }
}

function mapSheetBounds(
  bounds: SceneObject["bounds"],
  rows: SheetAxisMapper,
  columns: SheetAxisMapper,
): SceneObject["bounds"] {
  const right = columns.map(bounds.x + bounds.width);
  const bottom = rows.map(bounds.y + bounds.height);
  const x = columns.map(bounds.x);
  const y = rows.map(bounds.y);
  return { x, y, width: Math.max(0, right - x), height: Math.max(0, bottom - y) };
}

function sameSheetCoordinate(left: number, right: number): boolean {
  return Math.abs(left - right) < 0.001;
}

function mapSheetObject(object: SceneObject, unitIndex: number, rows: SheetAxisMapper, columns: SheetAxisMapper): SceneObject {
  if (object.unitIndex !== unitIndex) return object;
  const bounds = mapSheetBounds(object.bounds, rows, columns);
  if (object.visual.kind !== "group") return { ...object, bounds };
  return {
    ...object,
    bounds,
    visual: {
      ...object.visual,
      children: object.visual.children.map((child) => {
        const mapped = mapSheetBounds(child.bounds, rows, columns);
        if (sameSheetCoordinate(child.bounds.height, object.bounds.height)
          && child.bounds.width < object.bounds.width
          && (sameSheetCoordinate(child.bounds.x, object.bounds.x)
            || sameSheetCoordinate(
              child.bounds.x + child.bounds.width,
              object.bounds.x + object.bounds.width,
            ))) {
          const width = Math.min(child.bounds.width, bounds.width);
          return { ...child, bounds: {
            x: sameSheetCoordinate(child.bounds.x, object.bounds.x)
              ? bounds.x
              : bounds.x + bounds.width - width,
            y: bounds.y,
            width,
            height: bounds.height,
          } };
        }
        if (sameSheetCoordinate(child.bounds.width, object.bounds.width)
          && child.bounds.height < object.bounds.height
          && (sameSheetCoordinate(child.bounds.y, object.bounds.y)
            || sameSheetCoordinate(
              child.bounds.y + child.bounds.height,
              object.bounds.y + object.bounds.height,
            ))) {
          const height = Math.min(child.bounds.height, bounds.height);
          return { ...child, bounds: {
            x: bounds.x,
            y: sameSheetCoordinate(child.bounds.y, object.bounds.y)
              ? bounds.y
              : bounds.y + bounds.height - height,
            width: bounds.width,
            height,
          } };
        }
        return { ...child, bounds: mapped };
      }),
    },
  };
}

function xlsxPaintLayer(object: SceneObject): number {
  if (object.source.format !== "xlsx") return 0;
  // Worksheet picture backgrounds precede both the grid and all cells.
  if (object.z < 0) return -1;
  if (object.source.kind !== "cell") return 2;
  return object.type === "shape" && object.visual.kind === "group" ? 1 : 0;
}

function collapseXlsxCellBorder(object: SceneObject, sheetWidth: number, sheetHeight: number): SceneObject {
  if (xlsxPaintLayer(object) !== 1 || object.visual.kind !== "group") return object;
  const right = object.bounds.x + object.bounds.width;
  const bottom = object.bounds.y + object.bounds.height;
  return {
    ...object,
    visual: {
      ...object.visual,
      children: object.visual.children.map((child) => {
        if (child.visual.kind !== "painted-shape" || child.visual.geometry !== "rectangle") return child;
        if (sameSheetCoordinate(child.bounds.height, object.bounds.height)
          && sameSheetCoordinate(child.bounds.x + child.bounds.width, right)
          && right < sheetWidth) {
          return { ...child, bounds: { ...child.bounds, x: right } };
        }
        if (sameSheetCoordinate(child.bounds.width, object.bounds.width)
          && sameSheetCoordinate(child.bounds.y + child.bounds.height, bottom)
          && bottom < sheetHeight) {
          return { ...child, bounds: { ...child.bounds, y: bottom } };
        }
        return child;
      }),
    },
  };
}

function comparePaintOrder(left: SceneObject, right: SceneObject): number {
  return xlsxPaintLayer(left) - xlsxPaintLayer(right)
    || left.z - right.z
    || left.numericId - right.numericId;
}

function alignXlsxCellBorderToDevicePixels(
  context: OffscreenCanvasRenderingContext2D,
  bounds: SceneObject["bounds"],
): SceneObject["bounds"] {
  const { a, b, c, d, e, f } = context.getTransform();
  if (a <= 0 || d <= 0 || Math.abs(b) > Number.EPSILON || Math.abs(c) > Number.EPSILON) return bounds;
  const left = Math.round(bounds.x * a + e);
  const top = Math.round(bounds.y * d + f);
  const right = Math.round((bounds.x + bounds.width) * a + e);
  const bottom = Math.round((bounds.y + bounds.height) * d + f);
  return {
    x: (left - e) / a,
    y: (top - f) / d,
    width: Math.max(0, right - left) / a,
    height: Math.max(0, bottom - top) / d,
  };
}

interface DecodedSceneImage extends DecodedOfficeImage {
  readonly mediaType: string;
  readonly usedFallback: boolean;
}

interface CachedSceneImage {
  promise: Promise<DecodedSceneImage>;
  image?: DecodedSceneImage;
  pixels?: number;
}

interface CachedRenderedFrame {
  readonly unitIndex: number;
  readonly bitmap: ImageBitmap;
  readonly pixels: number;
  readonly result: Omit<RenderResult, "bitmap">;
}

interface ImageRasterTarget {
  readonly width: number;
  readonly height: number;
}

interface ImagePrefetch {
  readonly take: (state: VisualState) => Promise<DecodedSceneImage> | undefined;
  readonly consume: (state: VisualState) => void;
  readonly release: () => void;
}

function approximateImageMessage(mediaType: string): string {
  if (mediaType === "image/svg+xml") {
    return "The SVG contains text rasterized with document or host font assets; text metrics may differ from Microsoft Office";
  }
  if (["image/x-emf", "image/x-wmf", "image/x-emz", "image/x-wmz"].includes(mediaType)) {
    return "The Office metafile was rendered through a Canvas compatibility layer; complex GDI effects may differ";
  }
  return "The image was rendered through a compatibility decoder; some visual details may differ from Microsoft Office";
}

function vectorRasterTargetFor(
  mediaType: string,
  bounds: Viewport,
  deviceScale: number,
  maxPixels: number,
  crop?: {
    readonly cropLeft: number;
    readonly cropTop: number;
    readonly cropRight: number;
    readonly cropBottom: number;
  },
): ImageRasterTarget | undefined {
  if (!RERASTERIZED_VECTOR_MEDIA_TYPES.has(mediaType)) {
    return undefined;
  }
  const metafileScale = mediaType === "image/svg+xml" ? 1 : 2;
  const preserveMetafileEdge = metafileScale === 1 ? 0 : 1;
  const visibleWidth = Math.max(Number.EPSILON, 1 - (crop?.cropLeft ?? 0) - (crop?.cropRight ?? 0));
  const visibleHeight = Math.max(Number.EPSILON, 1 - (crop?.cropTop ?? 0) - (crop?.cropBottom ?? 0));
  const requestedWidth = Math.max(
    1,
    Math.ceil(bounds.width * deviceScale * metafileScale / visibleWidth) + preserveMetafileEdge,
  );
  const requestedHeight = Math.max(
    1,
    Math.ceil(bounds.height * deviceScale * metafileScale / visibleHeight) + preserveMetafileEdge,
  );
  const requestedPixels = requestedWidth * requestedHeight;
  const scale = requestedPixels > maxPixels
    ? Math.sqrt(maxPixels / requestedPixels)
    : 1;
  return {
    width: Math.max(1, Math.floor(requestedWidth * scale)),
    height: Math.max(1, Math.floor(requestedHeight * scale)),
  };
}

function vectorRasterTarget(
  object: SceneObject,
  state: VisualState,
  deviceScale: number,
  maxPixels: number,
): ImageRasterTarget | undefined {
  return state.visual.kind === "image"
    ? vectorRasterTargetFor(
        state.visual.mediaType,
        transformedBounds(object, state.transform),
        deviceScale,
        maxPixels,
        state.visual,
      )
    : undefined;
}

function nestedVectorRasterTarget(
  context: OffscreenCanvasRenderingContext2D,
  bounds: Viewport,
  mediaType: string,
  maxPixels: number,
  crop?: {
    readonly cropLeft: number;
    readonly cropTop: number;
    readonly cropRight: number;
    readonly cropBottom: number;
  },
): ImageRasterTarget | undefined {
  const transform = canvasTransform(context) ?? { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  return vectorRasterTargetFor(mediaType, {
    x: 0,
    y: 0,
    width: Math.hypot(transform.a, transform.b) * bounds.width,
    height: Math.hypot(transform.c, transform.d) * bounds.height,
  }, 1, maxPixels, crop);
}

function color(value: number): string {
  const red = (value >>> 24) & 0xff;
  const green = (value >>> 16) & 0xff;
  const blue = (value >>> 8) & 0xff;
  const alpha = (value & 0xff) / 255;
  return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
}

function visible(value: number): boolean {
  return (value & 0xff) !== 0;
}

function transparentPaint(paint: ScenePaint): boolean {
  if (paint.kind === "mapped-gradient") return transparentPaint(paint.paint);
  if (paint.kind === "none") return true;
  if (paint.kind === "solid") return !visible(paint.color);
  if (paint.kind === "pattern") return !visible(paint.foreground) && !visible(paint.background);
  if (paint.kind === "image") return false;
  if (paint.kind === "visual") return paint.opacity === 0;
  if (paint.kind === "xps-gradient") {
    return paint.stops.length > 0 && paint.stops.every((stop) => (
      stop.color.kind === "rgba" ? !visible(stop.color.color) : stop.color.alpha === 0
    ));
  }
  return paint.stops.length > 0 && paint.stops.every((stop) => !visible(stop.color));
}

function opaquePaint(paint: ScenePaint): boolean {
  if (paint.kind === "mapped-gradient") return opaquePaint(paint.paint);
  if (paint.kind === "solid") return (paint.color & 255) === 255;
  return (paint.kind === "linear-gradient" || paint.kind === "radial-gradient" || paint.kind === "shape-gradient")
    && paint.stops.length > 0 && paint.stops.every(stop => (stop.color & 255) === 255);
}

function transparentVisual(object: SceneObject, state: VisualState): boolean {
  if (state.media !== undefined) return false;
  if (state.operations.some(({ visual }) => visual.kind === "layer" && visual.opacity === 0)) {
    return true;
  }
  const { visual } = state;
  if (visual.kind === "none") return true;
  if (visual.kind === "shape") {
    return (visual.geometry === "line" || !visible(visual.fill))
      && (!visible(visual.stroke)
        || (visual.strokeWidth <= 0 && object.source.format !== "pdf"));
  }
  if (visual.kind === "painted-shape") {
    return (visual.geometry === "line" || transparentPaint(visual.fill))
      && (transparentPaint(visual.stroke)
        || (visual.strokeWidth <= 0 && object.source.format !== "pdf"));
  }
  return false;
}

function containsRect(outer: Viewport, inner: Viewport): boolean {
  return inner.x >= outer.x && inner.y >= outer.y
    && inner.x + inner.width <= outer.x + outer.width
    && inner.y + inner.height <= outer.y + outer.height;
}

function opaqueRectangle({ object, state }: PreparedSceneObject): Viewport | undefined {
  if (state.media !== undefined || state.operations.length !== 0) return undefined;
  const { visual } = state;
  if (visual.kind === "shape"
    && visual.geometry === "rectangle"
    && (visual.fill & 0xff) === 0xff) return object.bounds;
  if (visual.kind === "painted-shape"
    && visual.geometry === "rectangle"
    && visual.fill.kind === "solid"
    && (visual.fill.color & 0xff) === 0xff) return object.bounds;
  return undefined;
}

function strictlyBoundedVisual({ object, state }: PreparedSceneObject): boolean {
  if (state.media !== undefined || state.operations.length !== 0) return false;
  const { visual } = state;
  if (visual.kind === "image") return true;
  if (visual.kind === "shape" && visual.geometry === "rectangle") {
    return !visible(visual.stroke) || (visual.strokeWidth <= 0 && object.source.format !== "pdf");
  }
  return visual.kind === "painted-shape"
    && visual.geometry === "rectangle"
    && (transparentPaint(visual.stroke)
      || (visual.strokeWidth <= 0 && object.source.format !== "pdf"));
}

function cullInvisibleObjects(objects: readonly PreparedSceneObject[]): PreparedSceneObject[] {
  const result: PreparedSceneObject[] = [];
  const covers: Viewport[] = [];
  for (let index = objects.length - 1; index >= 0; index -= 1) {
    const object = objects[index]!;
    if (transparentVisual(object.object, object.state)) continue;
    if (strictlyBoundedVisual(object) && covers.some((cover) => containsRect(cover, object.bounds))) {
      continue;
    }
    result.push(object);
    const cover = opaqueRectangle(object);
    if (cover !== undefined) {
      // ponytail: bound cover checks; use a spatial index if real documents exceed this useful set.
      if (covers.length < 32) covers.push(cover);
      else {
        const smallest = covers.reduce(
          (candidate, value, candidateIndex) => value.width * value.height
            < covers[candidate]!.width * covers[candidate]!.height ? candidateIndex : candidate,
          0,
        );
        if (cover.width * cover.height > covers[smallest]!.width * covers[smallest]!.height) {
          covers[smallest] = cover;
        }
      }
    }
  }
  return result.reverse();
}

function runCodePoints(text: string): readonly number[] {
  const codePoints = new Set<number>();
  for (const character of text) {
    codePoints.add(character.codePointAt(0)!);
    if (codePoints.size === MAX_FONT_RESOLUTION_CODE_POINTS) break;
  }
  return [...codePoints].sort((left, right) => left - right);
}

const IDENTITY_TRANSFORM: SceneAffineTransform = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
type SceneOperationVisual = Extract<
  SceneVisual,
  { kind: "layer" | "effect" | "advanced-effect" | "image-color-change" | "image-adjustment" }
>;
type SceneAdvancedEffect = Extract<SceneVisual, { kind: "advanced-effect" }>;
type SceneImageColorChange = Extract<SceneVisual, { kind: "image-color-change" }>;
type SceneImageAdjustment = Extract<SceneVisual, { kind: "image-adjustment" }>;
type SceneMedia = Extract<SceneVisual, { kind: "media" }>;

interface VisualState {
  readonly visual: SceneVisual;
  readonly operations: readonly {
    readonly owner: SceneObject;
    readonly visual: SceneOperationVisual;
  }[];
  readonly textLayout?: SceneTextLayout;
  readonly textEffects?: readonly SceneTextEffect[];
  readonly strokeStyle?: SceneStrokeStyle;
  readonly media?: SceneMedia;
  readonly transform: SceneAffineTransform;
  /** Uniform-equivalent scale contributed by ancestor groups. */
  readonly groupedTextScale: number;
  readonly groupedTextScaleX: number;
  readonly groupedTextScaleY: number;
}

interface PreparedSceneObject {
  readonly object: SceneObject;
  readonly bounds: Viewport;
  readonly state: VisualState;
}

interface PreparedSceneUnit {
  readonly key: string;
  readonly objects: readonly PreparedSceneObject[];
  readonly maximumBottoms: Float64Array;
}

function visualText(visual: SceneVisual): string | undefined {
  if (visual.kind === "rich-text") return visual.runs.map(({ text }) => text).join("");
  if (visual.kind === "layer"
    || visual.kind === "effect"
    || visual.kind === "text-layout"
    || visual.kind === "text-effects"
    || visual.kind === "stroke-style"
    || visual.kind === "advanced-effect") return visualText(visual.visual);
  return undefined;
}

function expandPositionedTextBatch(object: SceneObject): readonly SceneObject[] {
  if (object.type !== "text-box"
    || (object.source.format !== "pdf" && object.source.format !== "ofd")
    || object.source.kind !== "text"
    || object.visual.kind !== "group"
    || object.visual.children.length < 2) return [object];
  const denominator = object.visual.children.length + 1;
  const childText = object.visual.children.map((child) => visualText(child.visual));
  const semantic = object.text !== undefined
    && childText.reduce((length, text) => length + (text?.length ?? 0), 0) === object.text.length
    ? object.text
    : undefined;
  let textOffset = 0;
  return object.visual.children.map((child, index) => {
    const { text: _text, ...original } = object;
    const visual = childText[index];
    const text = semantic === undefined || visual === undefined
      ? visual
      : semantic.slice(textOffset, textOffset + visual.length);
    textOffset += visual?.length ?? 0;
    return {
      ...original,
      numericId: object.numericId + (index + 1) / denominator,
      bounds: child.bounds,
      ...(text === undefined ? {} : { text }),
      z: object.z + index,
      visual: child.visual,
    };
  });
}

type BatchablePathPaint = Extract<ScenePaint, {
  readonly kind: "solid" | "linear-gradient";
}>;
type BatchablePdfPath = Extract<SceneVisual, { readonly kind: "painted-shape" }> & {
  readonly geometry: Extract<SceneGeometry, { readonly kind: "path" }>;
  readonly fill: BatchablePathPaint;
};

function batchablePdfPath(entry: {
  readonly object: SceneObject;
  readonly state: VisualState;
}): BatchablePdfPath | undefined {
  const { object, state } = entry;
  const visual = state.visual;
  if (object.source.format !== "pdf"
    || state.operations.length !== 0
    || visual.kind !== "painted-shape"
    || typeof visual.geometry !== "object"
    || visual.geometry.kind !== "path"
    || visual.stroke.kind !== "none"
    || (visual.fill.kind !== "solid"
      && visual.fill.kind !== "linear-gradient")) return undefined;
  const paint = visual.fill as BatchablePathPaint;
  const opaque = paint.kind === "solid"
    ? (paint.color & 0xff) === 0xff
    : paint.stops.every(({ color: value }) => (value & 0xff) === 0xff);
  return opaque ? visual as BatchablePdfPath : undefined;
}

function sameBatchPaint(
  leftObject: SceneObject,
  left: BatchablePathPaint,
  rightObject: SceneObject,
  right: BatchablePathPaint,
): boolean {
  if (left.kind !== right.kind) return false;
  if (left.kind === "solid" && right.kind === "solid") return left.color === right.color;
  if (left.kind === "linear-gradient" && right.kind === "linear-gradient") {
    if (leftObject.bounds.x + left.start.x !== rightObject.bounds.x + right.start.x
      || leftObject.bounds.y + left.start.y !== rightObject.bounds.y + right.start.y
      || leftObject.bounds.x + left.end.x !== rightObject.bounds.x + right.end.x
      || leftObject.bounds.y + left.end.y !== rightObject.bounds.y + right.end.y) return false;
  } else return false;
  return left.stops.length === right.stops.length && left.stops.every((stop, index) => (
    stop.offset === right.stops[index]?.offset && stop.color === right.stops[index]?.color
  ));
}

function relativeBrushTransform(
  bounds: SceneObject["bounds"],
  transform: SceneAffineTransform,
): SceneAffineTransform {
  if (bounds.width === 0 || bounds.height === 0) return transform;
  const toBounds = { a: bounds.width, b: 0, c: 0, d: bounds.height, e: bounds.x, f: bounds.y };
  const fromBounds = {
    a: 1 / bounds.width,
    b: 0,
    c: 0,
    d: 1 / bounds.height,
    e: -bounds.x / bounds.width,
    f: -bounds.y / bounds.height,
  };
  return multiplyTransform(toBounds, multiplyTransform(transform, fromBounds));
}

function visualTransform(visual: SceneVisual): SceneAffineTransform {
  let transform = IDENTITY_TRANSFORM;
  let current = visual;
  for (let depth = 0; depth <= MAX_VISUAL_DEPTH; depth += 1) {
    if (current.kind === "layer") transform = multiplyTransform(transform, current.transform);
    if (current.kind === "layer"
      || current.kind === "effect"
      || current.kind === "text-layout"
      || current.kind === "text-effects"
      || current.kind === "stroke-style"
      || current.kind === "advanced-effect"
      || current.kind === "image-color-change"
      || current.kind === "image-adjustment"
      || current.kind === "media") {
      current = current.visual;
    } else {
      break;
    }
  }
  return transform;
}

function unwrapVisual(
  owner: SceneObject,
): {
  readonly visual: SceneVisual;
  readonly operations: readonly {
    readonly owner: SceneObject;
    readonly visual: SceneOperationVisual;
  }[];
  readonly textLayout?: SceneTextLayout;
  readonly textEffects?: readonly SceneTextEffect[];
  readonly strokeStyle?: SceneStrokeStyle;
  readonly media?: SceneMedia;
} {
  const operations: {
    readonly owner: SceneObject;
    readonly visual: SceneOperationVisual;
  }[] = [];
  let textLayout: SceneTextLayout | undefined;
  let textEffects: readonly SceneTextEffect[] | undefined;
  let strokeStyle: SceneStrokeStyle | undefined;
  let media: SceneMedia | undefined;
  let current = owner.visual;
  let depth = 0;
  while (current.kind === "layer"
    || current.kind === "effect"
    || current.kind === "text-layout"
    || current.kind === "text-effects"
    || current.kind === "stroke-style"
    || current.kind === "advanced-effect"
    || current.kind === "image-color-change"
    || current.kind === "image-adjustment"
    || current.kind === "media") {
    if (depth >= MAX_VISUAL_DEPTH) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Visual nesting exceeds the renderer limit");
    }
    depth += 1;
    if (current.kind === "text-layout") textLayout = current.layout;
    else if (current.kind === "text-effects") textEffects = current.effects;
    else if (current.kind === "stroke-style") strokeStyle = current.style;
    else if (current.kind === "media") {
      if (media !== undefined) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "A visual cannot contain nested media assets");
      }
      media = current;
    }
    else operations.push({ owner, visual: current });
    current = current.visual;
  }
  return {
    visual: current,
    operations,
    ...(textLayout === undefined ? {} : { textLayout }),
    ...(textEffects === undefined ? {} : { textEffects }),
    ...(strokeStyle === undefined ? {} : { strokeStyle }),
    ...(media === undefined ? {} : { media }),
  };
}

function resolveVisualState(
  object: SceneObject,
  objectsById: ReadonlyMap<number, SceneObject>,
): VisualState {
  const ancestors: SceneObject[] = [];
  const visited = new Set<number>();
  let parentId = object.parentNumericId;
  while (parentId !== undefined && ancestors.length < 128 && !visited.has(parentId)) {
    visited.add(parentId);
    const parent = objectsById.get(parentId);
    if (parent === undefined) break;
    ancestors.unshift(parent);
    parentId = parent.parentNumericId;
  }
  const operations: {
    readonly owner: SceneObject;
    readonly visual: SceneOperationVisual;
  }[] = [];
  let groupedTransform = IDENTITY_TRANSFORM;
  for (const ancestor of ancestors) {
    if (ancestor.type !== "group") continue;
    const ancestorOperations = unwrapVisual(ancestor).operations;
    operations.push(...ancestorOperations);
    for (const operation of ancestorOperations) {
      if (operation.visual.kind !== "layer") continue;
      groupedTransform = multiplyTransform(groupedTransform, operation.visual.transform);
    }
  }
  const own = unwrapVisual(object);
  operations.push(...own.operations);
  let transform = IDENTITY_TRANSFORM;
  for (const operation of operations) {
    if (operation.visual.kind === "layer") {
      transform = multiplyTransform(transform, operation.visual.transform);
    }
  }
  const groupedTextScaleX = Math.hypot(groupedTransform.a, groupedTransform.b);
  const groupedTextScaleY = Math.hypot(groupedTransform.c, groupedTransform.d);
  const groupedTextScale = Math.sqrt(Math.abs(
    groupedTransform.a * groupedTransform.d - groupedTransform.b * groupedTransform.c,
  ));
  return {
    visual: own.visual,
    operations,
    ...(own.textLayout === undefined ? {} : { textLayout: own.textLayout }),
    ...(own.textEffects === undefined ? {} : { textEffects: own.textEffects }),
    ...(own.strokeStyle === undefined ? {} : { strokeStyle: own.strokeStyle }),
    ...(own.media === undefined ? {} : { media: own.media }),
    transform,
    groupedTextScale: Number.isFinite(groupedTextScale) && groupedTextScale > 0 ? groupedTextScale : 1,
    groupedTextScaleX: Number.isFinite(groupedTextScaleX) && groupedTextScaleX > 0 ? groupedTextScaleX : 1,
    groupedTextScaleY: Number.isFinite(groupedTextScaleY) && groupedTextScaleY > 0 ? groupedTextScaleY : 1,
  };
}

function isSceneObjectVisible(
  object: SceneObject,
  objectsById: ReadonlyMap<number, SceneObject>,
): boolean {
  const visited = new Set<number>();
  let current: SceneObject | undefined = object;
  while (current !== undefined && !visited.has(current.numericId)) {
    if (current.hidden) return false;
    visited.add(current.numericId);
    current = current.parentNumericId === undefined
      ? undefined
      : objectsById.get(current.parentNumericId);
  }
  return current === undefined;
}

function nestedGeometry(visual: SceneVisual): SceneGeometry | undefined {
  let current = visual;
  let depth = 0;
  while (depth < MAX_VISUAL_DEPTH) {
    if (current.kind === "shape"
      || current.kind === "text"
      || current.kind === "painted-shape"
      || current.kind === "rich-text") {
      return current.geometry;
    }
    if (current.kind === "effect" && current.clip !== undefined) return current.clip;
    if (current.kind === "layer"
      || current.kind === "effect"
      || current.kind === "text-layout"
      || current.kind === "text-effects"
      || current.kind === "stroke-style"
      || current.kind === "advanced-effect"
      || current.kind === "image-color-change"
      || current.kind === "image-adjustment"
      || current.kind === "media") {
      current = current.visual;
      depth += 1;
      continue;
    }
    return undefined;
  }
  return undefined;
}

function visualNeedsLowResolutionSupersampling(visual: SceneVisual): boolean {
  for (let depth = 0; depth <= 64; depth += 1) {
    if (visual.kind === "text-layout" && visual.layout.lowResolutionSupersample === true) return true;
    if (visual.kind === "layer"
      || visual.kind === "effect"
      || visual.kind === "text-layout"
      || visual.kind === "text-effects"
      || visual.kind === "stroke-style"
      || visual.kind === "advanced-effect"
      || visual.kind === "image-color-change"
      || visual.kind === "image-adjustment"
      || visual.kind === "media") {
      visual = visual.visual;
      continue;
    }
    return false;
  }
  return false;
}

function advancedEffectsFor(state: VisualState): readonly SceneAdvancedEffect[] {
  return state.operations
    .map(({ visual }) => visual)
    .filter((visual): visual is SceneAdvancedEffect => visual.kind === "advanced-effect");
}

function imageColorChangesFor(state: VisualState): readonly SceneImageColorChange[] {
  return state.operations
    .map(({ visual }) => visual)
    .filter((visual): visual is SceneImageColorChange => visual.kind === "image-color-change");
}

function imageAdjustmentsFor(state: VisualState): readonly SceneImageAdjustment[] {
  return state.operations
    .map(({ visual }) => visual)
    .filter((visual): visual is SceneImageAdjustment => visual.kind === "image-adjustment");
}

function imageAdjustmentKey(adjustments: readonly SceneImageAdjustment[]): string {
  return adjustments
    .map((value) => `${value.grayscale ? 1 : 0}:${value.bilevelThreshold ?? ""}:${value.brightness}:${value.contrast}:${value.duotone?.join(":") ?? ""}`)
    .join(",");
}

function adjustedChannel(value: number, brightness: number, contrast: number): number {
  const scale = contrast >= 0
    ? 1 / Math.max(1 - contrast, Number.EPSILON)
    : 1 + contrast;
  const black = 0.5 - 0.5 * scale;
  let channel = black + value / 255 * scale;
  channel += brightness >= 0
    ? (1 - black) * brightness
    : (black + scale) * brightness;
  return Math.round(Math.max(0, Math.min(1, channel)) * 255);
}

function applyImageAdjustments(
  bitmap: ImageBitmap,
  adjustments: readonly SceneImageAdjustment[],
  maxPixels: number,
  retainedPixels: number,
  maxTotalPixels: number,
): ImageBitmap {
  if (adjustments.length === 0) return bitmap;
  const pixels = bitmap.width * bitmap.height;
  if (!Number.isSafeInteger(pixels) || pixels <= 0 || pixels > maxPixels) {
    throw new OfficeEngineError("IMAGE_DIMENSION_LIMIT", `Decoded image adjustment requires ${pixels} pixels; limit is ${maxPixels}`);
  }
  if (retainedPixels + pixels > maxTotalPixels) {
    throw new OfficeEngineError("IMAGE_TOTAL_PIXEL_LIMIT", "Image-adjustment canvas exceeds the total image pixel limit");
  }
  const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
  const context = canvas.getContext("2d", { alpha: true, willReadFrequently: true });
  if (context === null) throw new OfficeEngineError("UNSUPPORTED_ENVIRONMENT", "A Canvas 2D context is required for image adjustments");
  context.drawImage(bitmap, 0, 0);
  const rowsPerChunk = Math.max(1, Math.floor(IMAGE_EFFECT_CHUNK_PIXELS / bitmap.width));
  for (let y = 0; y < bitmap.height; y += rowsPerChunk) {
    const height = Math.min(rowsPerChunk, bitmap.height - y);
    const image = context.getImageData(0, y, bitmap.width, height);
    for (let offset = 0; offset < image.data.length; offset += 4) {
      let red = image.data[offset]!;
      let green = image.data[offset + 1]!;
      let blue = image.data[offset + 2]!;
      for (const adjustment of adjustments) {
        const luminance = 0.2126 * red + 0.7152 * green + 0.0722 * blue;
        if (adjustment.grayscale) red = green = blue = luminance;
        if (adjustment.duotone !== undefined) {
          const ratio = luminance / 255;
          const [dark, light] = adjustment.duotone;
          red = ((dark >>> 24) & 0xff) * (1 - ratio) + ((light >>> 24) & 0xff) * ratio;
          green = ((dark >>> 16) & 0xff) * (1 - ratio) + ((light >>> 16) & 0xff) * ratio;
          blue = ((dark >>> 8) & 0xff) * (1 - ratio) + ((light >>> 8) & 0xff) * ratio;
        }
        if (adjustment.bilevelThreshold !== undefined) {
          red = green = blue = luminance >= adjustment.bilevelThreshold * 255 ? 255 : 0;
        }
        red = adjustedChannel(red, adjustment.brightness, adjustment.contrast);
        green = adjustedChannel(green, adjustment.brightness, adjustment.contrast);
        blue = adjustedChannel(blue, adjustment.brightness, adjustment.contrast);
      }
      image.data[offset] = red;
      image.data[offset + 1] = green;
      image.data[offset + 2] = blue;
    }
    context.putImageData(image, 0, y);
  }
  const result = canvas.transferToImageBitmap();
  bitmap.close();
  return result;
}

function applyImageColorChanges(
  bitmap: ImageBitmap,
  changes: readonly SceneImageColorChange[],
  tolerance: number,
  maxPixels: number,
  retainedPixels: number,
  maxTotalPixels: number,
): ImageBitmap {
  if (changes.length === 0) return bitmap;
  const pixels = bitmap.width * bitmap.height;
  if (!Number.isSafeInteger(pixels) || pixels <= 0 || pixels > maxPixels) {
    throw new OfficeEngineError(
      "IMAGE_DIMENSION_LIMIT",
      `Decoded image color change requires ${pixels} pixels; limit is ${maxPixels}`,
    );
  }
  const canvasPeakPixels = retainedPixels + pixels;
  if (!Number.isSafeInteger(canvasPeakPixels) || canvasPeakPixels > maxTotalPixels) {
    throw new OfficeEngineError(
      "IMAGE_TOTAL_PIXEL_LIMIT",
      `Image color-change canvas requires ${canvasPeakPixels} simultaneous image pixels; limit is ${maxTotalPixels}`,
    );
  }
  const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
  const context = canvas.getContext("2d", { alpha: true, willReadFrequently: true });
  if (context === null) {
    throw new OfficeEngineError(
      "UNSUPPORTED_ENVIRONMENT",
      "A Canvas 2D context is required for image color-change effects",
    );
  }
  context.drawImage(bitmap, 0, 0);
  let changed = false;
  const rowsPerChunk = Math.max(1, Math.floor(IMAGE_EFFECT_CHUNK_PIXELS / bitmap.width));
  for (let y = 0; y < bitmap.height; y += rowsPerChunk) {
    const height = Math.min(rowsPerChunk, bitmap.height - y);
    const image = context.getImageData(0, y, bitmap.width, height);
    let chunkChanged = false;
    for (let offset = 0; offset < image.data.length; offset += 4) {
      for (const change of changes) {
        if (Math.abs(image.data[offset]! - ((change.from >>> 24) & 0xff)) > tolerance
          || Math.abs(image.data[offset + 1]! - ((change.from >>> 16) & 0xff)) > tolerance
          || Math.abs(image.data[offset + 2]! - ((change.from >>> 8) & 0xff)) > tolerance
          || (change.useAlpha && image.data[offset + 3] !== (change.from & 0xff))) {
          continue;
        }
        image.data[offset] = (change.to >>> 24) & 0xff;
        image.data[offset + 1] = (change.to >>> 16) & 0xff;
        image.data[offset + 2] = (change.to >>> 8) & 0xff;
        image.data[offset + 3] = Math.round(image.data[offset + 3]! * (change.to & 0xff) / 255);
        changed = true;
        chunkChanged = true;
      }
    }
    if (chunkChanged) context.putImageData(image, 0, y);
  }
  if (!changed) return bitmap;
  const bitmapPeakPixels = canvasPeakPixels + pixels;
  if (!Number.isSafeInteger(bitmapPeakPixels) || bitmapPeakPixels > maxTotalPixels) {
    throw new OfficeEngineError(
      "IMAGE_TOTAL_PIXEL_LIMIT",
      `Image color-change bitmap requires ${bitmapPeakPixels} simultaneous image pixels; limit is ${maxTotalPixels}`,
    );
  }
  const result = canvas.transferToImageBitmap();
  bitmap.close();
  return result;
}

function applyImageAlphaMask(
  bitmap: ImageBitmap,
  mask: ImageBitmap,
  invertMask: boolean,
  flipMaskY: boolean,
  maxPixels: number,
  retainedPixels: number,
  maxTotalPixels: number,
): ImageBitmap {
  // PDF image masks may deliberately have a higher resolution than their
  // color source. Composite at the mask resolution so its silhouette is not
  // reduced to the source raster first (issue4246.pdf).
  const width = mask.width;
  const height = mask.height;
  const pixels = width * height;
  if (pixels <= 0 || pixels > maxPixels) {
    throw new OfficeEngineError(
      "IMAGE_DIMENSION_LIMIT",
      "Decoded PDF image mask exceeds the image dimension limit",
    );
  }
  const canvasPeakPixels = retainedPixels + pixels;
  if (!Number.isSafeInteger(canvasPeakPixels) || canvasPeakPixels > maxTotalPixels) {
    throw new OfficeEngineError(
      "IMAGE_TOTAL_PIXEL_LIMIT",
      `Image soft-mask canvas requires ${canvasPeakPixels} simultaneous image pixels; limit is ${maxTotalPixels}`,
    );
  }
  const canvas = new OffscreenCanvas(width, height);
  const maskCanvas = new OffscreenCanvas(width, height);
  const context = canvas.getContext("2d", { alpha: true, willReadFrequently: true });
  const maskContext = maskCanvas.getContext("2d", { alpha: false, willReadFrequently: true });
  if (context === null || maskContext === null) {
    throw new OfficeEngineError(
      "UNSUPPORTED_ENVIRONMENT",
      "A Canvas 2D context is required for PDF image soft masks",
    );
  }
  context.drawImage(bitmap, 0, 0, width, height);
  maskContext.drawImage(mask, 0, 0);
  const rowsPerChunk = Math.max(1, Math.floor(IMAGE_EFFECT_CHUNK_PIXELS / width));
  for (let y = 0; y < height; y += rowsPerChunk) {
    const chunkHeight = Math.min(rowsPerChunk, height - y);
    const image = context.getImageData(0, y, width, chunkHeight);
    const maskImage = maskContext.getImageData(
      0,
      flipMaskY ? height - y - chunkHeight : y,
      width,
      chunkHeight,
    );
    for (let pixel = 0, offset = 3; pixel < width * chunkHeight; pixel += 1, offset += 4) {
      const maskPixel = flipMaskY
        ? (chunkHeight - 1 - Math.floor(pixel / width)) * width + pixel % width
        : pixel;
      const maskAlpha = invertMask
        ? 255 - maskImage.data[maskPixel * 4]!
        : maskImage.data[maskPixel * 4]!;
      image.data[offset] = Math.round(image.data[offset]! * maskAlpha / 255);
    }
    context.putImageData(image, 0, y);
  }
  const bitmapPeakPixels = canvasPeakPixels + pixels;
  if (!Number.isSafeInteger(bitmapPeakPixels) || bitmapPeakPixels > maxTotalPixels) {
    throw new OfficeEngineError(
      "IMAGE_TOTAL_PIXEL_LIMIT",
      `Image soft-mask bitmap requires ${bitmapPeakPixels} simultaneous image pixels; limit is ${maxTotalPixels}`,
    );
  }
  const result = canvas.transferToImageBitmap();
  bitmap.close();
  mask.close();
  return result;
}

function threeDProjection(style: SceneThreeDStyle): readonly [number, number] {
  const preset = style.cameraPreset.toLowerCase();
  let x = preset.includes("left") ? -1 : preset.includes("right") ? 1 : 0;
  let y = preset.includes("above") ? -1 : preset.includes("below") ? 1 : 0;
  if (preset.includes("facing") && y === 0) y = 0.45;
  if (preset.includes("heroic") && y === 0) y = -0.45;
  let normalize = x !== 0 || y !== 0;
  x += Math.sin(style.cameraLongitude * Math.PI / 180);
  y -= Math.sin(style.cameraLatitude * Math.PI / 180);
  if (Math.abs(x) + Math.abs(y) < 0.001) {
    const direction = style.lightDirection.toLowerCase();
    x = direction.includes("l") ? 0.5 : direction.includes("r") ? -0.5 : 0.35;
    y = direction.startsWith("t") ? 0.5 : direction.startsWith("b") ? -0.5 : 0.35;
    normalize = true;
  }
  const length = Math.hypot(x, y) || 1;
  const distance = style.extrusionHeight / Math.max(0.1, style.cameraZoom);
  const divisor = normalize ? length : Math.max(1, length);
  return [x / divisor * distance, y / divisor * distance];
}

function threeDMaterialFilter(style: SceneThreeDStyle, depthRatio: number): string {
  const material = style.material.toLowerCase();
  const contrast = material.includes("metal") ? 1.35
    : material.includes("plastic") ? 1.18
      : material.includes("edge") ? 1.25
        : 1.08;
  const saturation = material.includes("metal") ? 0.82
    : material.includes("plastic") ? 1.12
      : material.includes("powder") ? 0.9
        : 1;
  const rig = style.lightRig.toLowerCase();
  const rigLight = rig.includes("bright") || rig.includes("flood") ? 1.15
    : rig.includes("harsh") || rig.includes("chilly") ? 0.88
      : 1;
  const rotationLight = 0.12 * Math.sin(
    (style.lightLatitude + style.lightLongitude + style.lightRevolution) * Math.PI / 180,
  );
  const brightness = Math.max(0.45, rigLight + rotationLight - depthRatio * 0.35);
  return `brightness(${brightness}) contrast(${contrast}) saturate(${saturation})`;
}

function threeDCameraMatrix(style: SceneThreeDStyle): readonly [number, number, number, number, number, number] {
  const latitude = style.cameraLatitude * Math.PI / 180;
  const longitude = style.cameraLongitude * Math.PI / 180;
  const revolution = style.cameraRevolution * Math.PI / 180;
  const zoom = style.cameraZoom;
  const cosX = Math.cos(latitude);
  const sinX = Math.sin(latitude);
  const cosY = Math.cos(longitude);
  const sinY = Math.sin(longitude);
  const cosZ = Math.cos(revolution);
  const sinZ = Math.sin(revolution);
  const a = (cosZ * cosY + sinZ * sinX * sinY) * zoom;
  const b = (-sinZ * cosY + cosZ * sinX * sinY) * zoom;
  const c = sinZ * cosX * zoom;
  const d = cosZ * cosX * zoom;
  // Office camera distance in CSS pixels (96 dpi), with the authored field of view.
  const inverseDistance = style.cameraFov > 0 ? Math.tan(style.cameraFov * Math.PI / 360) / (15976 * 96 / 2540) : 0;
  return [a, b, c, d, -cosX * sinY * inverseDistance, sinX * inverseDistance];
}

function applyThreeDCamera(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  style: SceneThreeDStyle,
): void {
  const [a, b, c, d] = threeDCameraMatrix(style);
  if (Math.abs(a - 1) < 0.0001 && Math.abs(b) < 0.0001
    && Math.abs(c) < 0.0001 && Math.abs(d - 1) < 0.0001 && style.backdrop === undefined) return;
  const centerX = object.bounds.x + object.bounds.width / 2;
  const centerY = object.bounds.y + object.bounds.height / 2;
  context.translate(centerX, centerY);
  if (style.backdrop !== undefined) {
    const normalLength = Math.hypot(
      style.backdrop.normalX,
      style.backdrop.normalY,
      style.backdrop.normalZ,
    ) || 1;
    const upLength = Math.hypot(style.backdrop.upX, style.backdrop.upY, style.backdrop.upZ) || 1;
    const shearX = style.backdrop.normalY / normalLength * 0.2;
    const shearY = style.backdrop.upX / upLength * 0.2;
    context.transform(
      1,
      shearY,
      shearX,
      1,
      style.backdrop.anchorX + style.backdrop.normalX / normalLength * style.backdrop.anchorZ,
      style.backdrop.anchorY + style.backdrop.normalY / normalLength * style.backdrop.anchorZ,
    );
  }
  context.transform(a, b, c, d, 0, 0);
  context.translate(-centerX, -centerY);
}

// Inverse mapping avoids the alpha seams produced by overlapping Canvas triangles.
function drawPerspectiveSurface(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  style: SceneThreeDStyle,
  draw: (target: OffscreenCanvasRenderingContext2D) => void,
): void {
  const [a, b, c, d, g, h] = threeDCameraMatrix(style);
  const determinant = a * d - b * c;
  if (Math.abs(determinant) < 1e-8) return; // An edge-on plane has no visible area.
  const { x, y, width, height } = object.bounds;
  if (width <= 0 || height <= 0) return;
  const centerX = x + width / 2, centerY = y + height / 2;
  let padding = 2 + style.extrusionHeight + style.contourWidth;
  for (const { visual } of unwrapVisual(object).operations) {
    if (visual.kind === "advanced-effect") {
      padding = Math.max(padding, (visual.glow?.radius ?? 0) * 2, (visual.softEdge ?? 0) * 2);
    }
  }
  const left = -width / 2 - padding, top = -height / 2 - padding;
  const sourceWidth = width + padding * 2, sourceHeight = height + padding * 2;
  const corners = [[left, top], [left + sourceWidth, top],
    [left, top + sourceHeight], [left + sourceWidth, top + sourceHeight]] as const;
  if (corners.some(([px, py]) => 1 + g * px + h * py <= 0.01)) {
    throw new OfficeEngineError("RENDER_FAILED", "3D picture intersects the camera near plane");
  }
  const projected = corners.map(([px, py]) => {
    const w = 1 + g * px + h * py;
    return [(a * px + c * py) / w, (b * px + d * py) / w] as const;
  });
  const minX = Math.min(...projected.map(p => p[0])), minY = Math.min(...projected.map(p => p[1]));
  const outputWidth = Math.max(...projected.map(p => p[0])) - minX;
  const outputHeight = Math.max(...projected.map(p => p[1])) - minY;
  const transform = context.getTransform();
  const scaleX = Math.hypot(transform.a, transform.b), scaleY = Math.hypot(transform.c, transform.d);
  // ponytail: CPU resampling is bounded to the existing 8M-pixel raster budget; use a GPU pass if profiling warrants it.
  const sourceSize = gradientRasterSize(sourceWidth, sourceHeight,
    Math.min(scaleX, 8192 / sourceWidth), Math.min(scaleY, 8192 / sourceHeight));
  const outputSize = gradientRasterSize(outputWidth, outputHeight,
    Math.min(scaleX, 8192 / outputWidth), Math.min(scaleY, 8192 / outputHeight));
  const source = new OffscreenCanvas(sourceSize.width, sourceSize.height);
  const sourceContext = source.getContext("2d", { willReadFrequently: true });
  const output = new OffscreenCanvas(outputSize.width, outputSize.height);
  const outputContext = output.getContext("2d");
  if (sourceContext === null || outputContext === null) throw new OfficeEngineError("RENDER_FAILED", "Cannot allocate 3D surface");
  sourceContext.scale(sourceSize.scaleX, sourceSize.scaleY);
  sourceContext.translate(-centerX - left, -centerY - top);
  draw(sourceContext);
  const pixels = sourceContext.getImageData(0, 0, source.width, source.height).data;
  const result = outputContext.createImageData(output.width, output.height);
  for (let row = 0; row < output.height; row++) {
    const v = minY + (row + 0.5) / outputSize.scaleY;
    for (let column = 0; column < output.width; column++) {
      const u = minX + (column + 0.5) / outputSize.scaleX;
      const denominator = determinant + (b * h - d * g) * u + (c * g - a * h) * v;
      if (Math.abs(denominator) < 1e-8) continue;
      const sx = ((d * u - c * v) / denominator - left) * sourceSize.scaleX - 0.5;
      const sy = ((a * v - b * u) / denominator - top) * sourceSize.scaleY - 0.5;
      if (sx < 0 || sy < 0 || sx >= source.width - 1 || sy >= source.height - 1) continue;
      const ix = Math.floor(sx), iy = Math.floor(sy), fx = sx - ix, fy = sy - iy;
      const index = (row * output.width + column) * 4;
      let alpha = 0, red = 0, green = 0, blue = 0;
      for (let dy = 0; dy < 2; dy++) for (let dx = 0; dx < 2; dx++) {
        const sample = ((iy + dy) * source.width + ix + dx) * 4;
        const weight = (dx ? fx : 1 - fx) * (dy ? fy : 1 - fy) * pixels[sample + 3]!;
        alpha += weight;
        red += pixels[sample]! * weight;
        green += pixels[sample + 1]! * weight;
        blue += pixels[sample + 2]! * weight;
      }
      if (alpha > 0) {
        result.data[index] = red / alpha;
        result.data[index + 1] = green / alpha;
        result.data[index + 2] = blue / alpha;
        result.data[index + 3] = alpha;
      }
    }
  }
  outputContext.putImageData(result, 0, 0);
  context.drawImage(output, centerX + minX, centerY + minY, outputWidth, outputHeight);
}

function threeDLightDirection(style: SceneThreeDStyle): readonly [number, number] {
  const rotation = style.lightRevolution * Math.PI / 180;
  const direction = style.lightDirection.toLowerCase();
  const x = direction.includes("l") ? -1 : direction.includes("r") ? 1 : 0;
  const y = direction.startsWith("b") ? 1 : -1;
  return [x * Math.cos(rotation) - y * Math.sin(rotation),
    x * Math.sin(rotation) + y * Math.cos(rotation)];
}

function drawWithThreeD(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  style: SceneThreeDStyle,
  draw: (target: OffscreenCanvasRenderingContext2D) => void,
  surfaceOnly = false,
): void {
  const [offsetX, offsetY] = threeDProjection(style);
  if (surfaceOnly && style.flatTextZ !== undefined) {
    context.save();
    try {
      const depthScale = style.flatTextZ / Math.max(1, style.extrusionHeight);
      context.translate(offsetX * depthScale, offsetY * depthScale);
      draw(context);
    } finally {
      context.restore();
    }
    return;
  }
  const [, , , , g, h] = threeDCameraMatrix(style);
  if (g !== 0 || h !== 0) {
    drawPerspectiveSurface(context, object, style, target => drawWithThreeD(target, object, {
      ...style, cameraFov: 0, cameraZoom: 1, cameraLatitude: 0, cameraLongitude: 0, cameraRevolution: 0,
    }, draw, surfaceOnly));
    return;
  }
  // WebKit accepts a filter expando but does not implement Canvas filters.
  const supportsFilters = "filter" in Object.getPrototypeOf(context);
  let textBounds: Viewport | undefined;
  if (style.appliesToText) {
    // Resolve font metrics once. Repainting on differently transformed scratch
    // contexts can move WebKit glyph baselines between the face and bevel.
    const surface = new OffscreenCanvas(context.canvas.width, context.canvas.height);
    const target = surface.getContext("2d");
    const transform = context.getTransform();
    const determinant = transform.a * transform.d - transform.b * transform.c;
    if (target !== null && Math.abs(determinant) > 1e-12) {
      target.setTransform(transform);
      draw(target);
      if (style.extrusionHeight > 0 && style.extrusionColor !== undefined && supportsFilters) {
        // Measure the rendered alpha, not the text box: glyph overhangs, glow,
        // and nested transforms can all paint outside the authored bounds.
        const pixels = target.getImageData(0, 0, surface.width, surface.height).data;
        let left = surface.width, top = surface.height, right = 0, bottom = 0;
        for (let y = 0; y < surface.height; y++) for (let x = 0; x < surface.width; x++) {
          if (pixels[(y * surface.width + x) * 4 + 3] === 0) continue;
          left = Math.min(left, x); top = Math.min(top, y);
          right = Math.max(right, x + 1); bottom = Math.max(bottom, y + 1);
        }
        if (right <= left || bottom <= top) {
          if (context.globalCompositeOperation === "source-over") return;
        } else if (context.globalCompositeOperation === "source-over") {
          textBounds = transformedBounds({ ...object,
            bounds: { x: left - 2, y: top - 2, width: right - left + 4, height: bottom - top + 4 },
          }, transform.inverse());
        }
      }
      draw = layer => {
        layer.save();
        try {
          layer.transform(transform.d / determinant, -transform.b / determinant,
            -transform.c / determinant, transform.a / determinant,
            (transform.c * transform.f - transform.d * transform.e) / determinant,
            (transform.b * transform.e - transform.a * transform.f) / determinant);
          layer.drawImage(surface, 0, 0);
        } finally { layer.restore(); }
      };
    }
  }
  const steps = Math.min(32, Math.max(0, Math.ceil(style.extrusionHeight)));
  // Reuse the full-size raster to preserve subpixel sampling; limit only the
  // filtered composite to the ink bounds, with a transparent sampling margin.
  let extrusionLayer: OffscreenCanvas | undefined;
  for (let step = surfaceOnly ? 0 : steps; step >= 1; step -= 1) {
    context.save();
    try {
      const ratio = step / Math.max(1, steps);
      applyThreeDCamera(context, object, style);
      context.translate(offsetX * ratio, offsetY * ratio);
      context.filter = threeDMaterialFilter(style, ratio);
      if (style.extrusionColor !== undefined) {
        const projected = textBounds === undefined ? undefined
          : transformedBounds({ ...object, bounds: textBounds }, context.getTransform());
        const left = projected === undefined ? 0 : Math.max(0, Math.floor(projected.x));
        const top = projected === undefined ? 0 : Math.max(0, Math.floor(projected.y));
        const right = projected === undefined ? context.canvas.width
          : Math.min(context.canvas.width, Math.ceil(projected.x + projected.width));
        const bottom = projected === undefined ? context.canvas.height
          : Math.min(context.canvas.height, Math.ceil(projected.y + projected.height));
        if (right <= left || bottom <= top) continue;
        const layer = colorizedLayer(context, style.extrusionColor, draw,
          supportsFilters ? extrusionLayer : undefined);
        if (layer !== undefined) {
          extrusionLayer = layer;
          context.save();
          try {
            context.setTransform(1, 0, 0, 1, 0, 0);
            if (textBounds === undefined) context.drawImage(layer, 0, 0);
            else context.drawImage(layer, left, top, right - left, bottom - top,
              left, top, right - left, bottom - top);
          } finally {
            context.restore();
          }
          continue;
        }
      }
      draw(context);
    } finally {
      context.restore();
    }
  }
  if (!surfaceOnly && (style.contourWidth > 0 || style.bevelBottom !== undefined)) {
    context.save();
    try {
      applyThreeDCamera(context, object, style);
      if (style.contourColor !== undefined) context.shadowColor = color(style.contourColor);
      context.shadowBlur = style.contourWidth;
      context.filter = threeDMaterialFilter(style, 0.55);
      draw(context);
    } finally {
      context.restore();
    }
  }
  context.save();
  try {
    applyThreeDCamera(context, object, style);
    if (style.z !== 0) context.translate(offsetX * style.z / Math.max(1, style.extrusionHeight), offsetY * style.z / Math.max(1, style.extrusionHeight));
    // DrawingML's solid/gradient paint describes the visible front face. The
    // lighting/material treatment belongs to the extrusion, contour, and
    // bevel layers; applying it to the final face recolors the entire shape.
    // Reset explicitly because a nested soft-edge/effect may have installed a
    // filter before the 3D wrapper was entered.
    context.filter = "none";
    draw(context);
  } finally {
    context.restore();
  }
  if (!surfaceOnly
    && style.bevelTop !== undefined
    && (style.bevelTop.width > 0 || style.bevelTop.height > 0)) {
    const insetRectangle = style.bevelTop.preset.toLowerCase() === "relaxedinset"
      && nestedGeometry(object.visual) === "rectangle";
    const bevelCanvas = new OffscreenCanvas(context.canvas.width, context.canvas.height);
    const bevelContext = bevelCanvas.getContext("2d");
    if (bevelContext === null) return;
    const transform = canvasTransform(context) ?? { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
    bevelContext.transform(
      transform.a,
      transform.b,
      transform.c,
      transform.d,
      transform.e,
      transform.f,
    );
    bevelContext.imageSmoothingEnabled = context.imageSmoothingEnabled;
    bevelContext.imageSmoothingQuality = context.imageSmoothingQuality;
    applyThreeDCamera(bevelContext, object, style);
    draw(bevelContext);
    if (style.appliesToText) {
      // A text bevel follows glyph contours. Shrinking the entire text box
      // creates a displaced second copy, particularly with WebKit compositing.
      const source = bevelContext.getImageData(0, 0, bevelCanvas.width, bevelCanvas.height);
      const rim = bevelContext.createImageData(source.width, source.height);
      const [lightX, lightY] = threeDLightDirection(style);
      const radius = Math.min(style.bevelTop.width, style.bevelTop.height);
      const camera = bevelContext.getTransform();
      const shiftX = Math.round((camera.a * lightX + camera.c * lightY) * radius);
      const shiftY = Math.round((camera.b * lightX + camera.d * lightY) * radius);
      const alpha = (x: number, y: number): number => x < 0 || y < 0 || x >= source.width || y >= source.height
        ? 0 : source.data[(y * source.width + x) * 4 + 3]!;
      // ponytail: directional alpha relief; curved bevel presets need a distance field for exact normals.
      for (let y = 0; y < source.height; y++) for (let x = 0; x < source.width; x++) {
        const index = (y * source.width + x) * 4;
        const a = source.data[index + 3]!;
        if (a === 0) continue;
        const relief = alpha(x - shiftX, y - shiftY) - alpha(x + shiftX, y + shiftY);
        rim.data[index] = rim.data[index + 1] = rim.data[index + 2] = relief > 0 ? 255 : 0;
        rim.data[index + 3] = Math.abs(relief) * a / 255 * 0.15;
      }
      bevelContext.putImageData(rim, 0, 0);
      context.save();
      try { context.setTransform(1, 0, 0, 1, 0, 0); context.drawImage(bevelCanvas, 0, 0); }
      finally { context.restore(); }
      return;
    }
    const geometry = nestedGeometry(object.visual);
    const { x, y, width, height } = object.bounds;
    const polygon = typeof geometry === "object" && geometry.kind === "path"
      && geometry.commands[0]?.kind === "moveTo" && geometry.commands.at(-1)?.kind === "closePath"
      && geometry.commands.slice(1, -1).every(command => command.kind === "lineTo")
      ? geometry.commands.flatMap(command => command.kind === "moveTo" || command.kind === "lineTo"
        ? [[x + command.x, y + command.y]] : [])
      : geometry === "rectangle" && !insetRectangle
        ? [[x, y], [x + width, y], [x + width, y + height], [x, y + height]] : [];
    if (polygon.length >= 3 && polygon.length <= 64 && style.bevelTop.width > 0
      && style.bevelTop.preset.toLowerCase() === "circle") {
      const area = polygon.reduce((sum, a, i) => {
        const b = polygon[(i + 1) % polygon.length]!;
        return sum + a[0]! * b[1]! - b[0]! * a[1]!;
      }, 0);
      const orientation = Math.sign(area);
      const normals = polygon.map((a, i) => {
        const b = polygon[(i + 1) % polygon.length]!;
        const dx = b[0]! - a[0]!, dy = b[1]! - a[1]!;
        const length = Math.hypot(dx, dy);
        return [orientation * dy / length, -orientation * dx / length];
      });
      const center = [0, 1].map(axis => polygon.reduce((sum, p) => sum + p[axis]!, 0) / polygon.length);
      const distances = polygon.map((a, i) => (a[0]! - center[0]!) * normals[i]![0]!
        + (a[1]! - center[1]!) * normals[i]![1]!);
      // ponytail: at most 64 convex edges; complex contours retain the raster bevel below.
      const convex = polygon.every((a, i) => polygon.every(b =>
        (b[0]! - a[0]!) * normals[i]![0]! + (b[1]! - a[1]!) * normals[i]![1]! <= 1e-5));
      if (area !== 0 && convex && distances.every(d => Number.isFinite(d) && d > 0)) {
        const inset = Math.min(style.bevelTop.width, ...distances.map(d => d * 0.95));
        const inner = polygon.map((a, i) => {
          const previous = normals[(i + polygon.length - 1) % polygon.length]!, n = normals[i]!;
          const divisor = 1 + previous[0]! * n[0]! + previous[1]! * n[1]!;
          return [a[0]! - inset * (previous[0]! + n[0]!) / divisor,
            a[1]! - inset * (previous[1]! + n[1]!) / divisor];
        });
        const [lightX, lightY] = threeDLightDirection(style);
        const mask = new OffscreenCanvas(bevelCanvas.width, bevelCanvas.height);
        const maskContext = mask.getContext("2d");
        if (maskContext === null) return;
        maskContext.drawImage(bevelCanvas, 0, 0);
        bevelContext.save();
        bevelContext.setTransform(1, 0, 0, 1, 0, 0);
        bevelContext.clearRect(0, 0, bevelCanvas.width, bevelCanvas.height);
        bevelContext.restore();
        bevelContext.globalCompositeOperation = "source-over";
        for (let i = 0; i < polygon.length; i++) {
          const next = (i + 1) % polygon.length;
          const a = polygon[i]!, b = polygon[next]!, c = inner[next]!, d = inner[i]!;
          const illumination = normals[i]![0]! * lightX + normals[i]![1]! * lightY;
          const rim = bevelContext.createLinearGradient((a[0]! + b[0]!) / 2, (a[1]! + b[1]!) / 2,
            (c[0]! + d[0]!) / 2, (c[1]! + d[1]!) / 2);
          const tint = illumination >= 0 ? "255,255,255" : "0,0,0";
          rim.addColorStop(0, `rgba(${tint},${Math.abs(illumination) * 0.6})`);
          rim.addColorStop(1, `rgba(${tint},${Math.abs(illumination) * 0.15})`);
          bevelContext.fillStyle = rim;
          bevelContext.beginPath(); bevelContext.moveTo(a[0]!, a[1]!);
          for (const point of [b, c, d]) bevelContext.lineTo(point[0]!, point[1]!);
          bevelContext.closePath(); bevelContext.fill();
        }
        // Keep authored fill/stroke alpha; do not paint the front face twice.
        bevelContext.globalCompositeOperation = "destination-in";
        bevelContext.setTransform(1, 0, 0, 1, 0, 0);
        bevelContext.drawImage(mask, 0, 0);
        context.save();
        try { context.setTransform(1, 0, 0, 1, 0, 0); context.drawImage(bevelCanvas, 0, 0); }
        finally { context.restore(); }
        return;
      }
    }
    bevelContext.save();
    try {
      const centerX = object.bounds.x + object.bounds.width / 2;
      const centerY = object.bounds.y + object.bounds.height / 2;
      const scaleX = Math.max(0.01, 1 - style.bevelTop.width * (insetRectangle ? 2 : 1) / Math.max(1, object.bounds.width));
      const scaleY = Math.max(0.01, 1 - (insetRectangle ? style.bevelTop.width * 2 : style.bevelTop.height) / Math.max(1, object.bounds.height));
      bevelContext.globalCompositeOperation = "destination-out";
      bevelContext.translate(centerX, centerY);
      bevelContext.scale(scaleX, scaleY);
      bevelContext.translate(-centerX, -centerY);
      draw(bevelContext);
    } finally {
      bevelContext.restore();
    }
    bevelContext.save();
    try {
      if (insetRectangle) {
        // Bevel width is measured in the face plane; height is its Z depth.
        // relaxedInset has a ridge inside each face, not a flat tinted border.
        const { x, y, width, height } = object.bounds;
        const inset = Math.min(style.bevelTop.width, width * 0.495, height * 0.495);
        const outer = [[x, y], [x + width, y], [x + width, y + height], [x, y + height]] as const;
        const inner = [[x + inset, y + inset], [x + width - inset, y + inset],
          [x + width - inset, y + height - inset], [x + inset, y + height - inset]] as const;
        // ponytail: faceted lighting; per-pixel normals are needed for exact Office highlights.
        const [lightX, lightY] = threeDLightDirection(style);
        const normals = [[0, -1], [1, 0], [0, 1], [-1, 0]] as const;
        const depth = Math.min(1, style.bevelTop.height / Math.max(0.01, inset));
        const specular = style.material.toLowerCase().includes("plastic") ? 0.8 : 0.5;
        bevelContext.globalCompositeOperation = "source-atop";
        for (let side = 0; side < 4; side++) {
          const next = (side + 1) % 4;
          const a = outer[side]!, b = outer[next]!, c = inner[next]!, d = inner[side]!;
          const n = normals[side]!;
          const illumination = (n[0] * lightX + n[1] * lightY)
            * (style.lightRig === "threePt" && n[0] === 0 ? depth : 1);
          const rim = bevelContext.createLinearGradient((a[0] + b[0]) / 2, (a[1] + b[1]) / 2,
            (c[0] + d[0]) / 2, (c[1] + d[1]) / 2);
          rim.addColorStop(0, `rgba(0,0,0,${0.25 + depth * 0.3})`);
          rim.addColorStop(0.45, `rgba(255,255,255,${Math.max(0.08, illumination * specular)})`);
          // Three-point lighting has a secondary rim light opposite the key.
          const secondary = style.lightRig === "threePt" ? Math.abs(n[0]) * specular * 0.65 : 0;
          rim.addColorStop(0.55, `rgba(255,255,255,${Math.max(secondary, illumination * specular)})`);
          rim.addColorStop(1, "rgba(255,255,255,0)");
          bevelContext.fillStyle = rim;
          bevelContext.beginPath();
          bevelContext.moveTo(a[0], a[1]); bevelContext.lineTo(b[0], b[1]);
          bevelContext.lineTo(c[0], c[1]); bevelContext.lineTo(d[0], d[1]);
          bevelContext.closePath(); bevelContext.fill();
        }
      } else {
        const preset = style.bevelTop.preset.toLowerCase();
        const [light, shadow, transition] = preset === "softround" ? [0.32, 0.28, 0.24]
          : preset === "convex" ? [0.2, 0.4, 0.38]
            : preset === "slope" ? [0.12, 0.46, 0.18]
              : preset === "hardedge" ? [0.16, 0.52, 0.46]
                : preset === "artdeco" ? [0.28, 0.44, 0.42]
                  : [0.24, 0.36, 0.32];
        const direction = style.lightDirection.toLowerCase();
        const lightX = direction.includes("r") ? object.bounds.x + object.bounds.width
          : object.bounds.x;
        const lightY = direction.startsWith("b") ? object.bounds.y + object.bounds.height
          : object.bounds.y;
        const gradient = bevelContext.createLinearGradient(
          lightX,
          lightY,
          object.bounds.x + object.bounds.width - (lightX - object.bounds.x),
          object.bounds.y + object.bounds.height - (lightY - object.bounds.y),
        );
        gradient.addColorStop(0, `rgba(255, 255, 255, ${light})`);
        gradient.addColorStop(transition, "rgba(255, 255, 255, 0)");
        gradient.addColorStop(1 - transition, "rgba(0, 0, 0, 0)");
        gradient.addColorStop(1, `rgba(0, 0, 0, ${shadow})`);
        bevelContext.globalCompositeOperation = "source-in";
        bevelContext.fillStyle = gradient;
        bevelContext.fillRect(
          object.bounds.x,
          object.bounds.y,
          object.bounds.width,
          object.bounds.height,
        );
      }
    } finally {
      bevelContext.restore();
    }
    context.save();
    try {
      context.setTransform(1, 0, 0, 1, 0, 0);
      context.drawImage(bevelCanvas, 0, 0);
    } finally {
      context.restore();
    }
  }
}

function colorizedLayer(
  target: OffscreenCanvasRenderingContext2D,
  fill: number,
  draw: (layer: OffscreenCanvasRenderingContext2D) => void,
  scratch?: OffscreenCanvas,
): OffscreenCanvas | undefined {
  const canvas = scratch ?? new OffscreenCanvas(target.canvas.width, target.canvas.height);
  const context = canvas.getContext("2d");
  if (context === null) return undefined;
  if (scratch !== undefined) context.clearRect(0, 0, canvas.width, canvas.height);
  context.save();
  try {
    context.setTransform(target.getTransform());
    context.imageSmoothingEnabled = target.imageSmoothingEnabled;
    context.imageSmoothingQuality = target.imageSmoothingQuality;
    draw(context);
    context.setTransform(1, 0, 0, 1, 0, 0);
    context.globalAlpha = 1;
    context.globalCompositeOperation = "source-in";
    context.fillStyle = color(fill);
    context.fillRect(0, 0, canvas.width, canvas.height);
    return canvas;
  } finally { context.restore(); }
}

// Separable sliding maxima spread alpha in O(pixels), independent of glow radius.
function spreadGlowAlpha(image: ImageData, radius: number): void {
  const { width, height, data } = image;
  const alpha = new Uint8Array(width * height);
  const scratch = new Uint8Array(alpha.length);
  const queue = new Int32Array(Math.max(width, height));
  for (let i = 0; i < alpha.length; i++) alpha[i] = data[i * 4 + 3]!;
  for (const horizontal of [true, false]) {
    const length = horizontal ? width : height;
    const lines = horizontal ? height : width;
    const stride = horizontal ? 1 : width;
    const source = horizontal ? alpha : scratch;
    const output = horizontal ? scratch : alpha;
    for (let line = 0; line < lines; line++) {
      const base = horizontal ? line * width : line;
      let head = 0;
      let tail = 0;
      let next = 0;
      for (let position = 0; position < length; position++) {
        const end = Math.min(length - 1, position + radius);
        while (next <= end) {
          while (tail > head && source[base + queue[tail - 1]! * stride]! <= source[base + next * stride]!) tail--;
          queue[tail++] = next++;
        }
        while (head < tail && queue[head]! < position - radius) head++;
        output[base + position * stride] = source[base + queue[head]! * stride]!;
      }
    }
  }
  for (let i = 0; i < alpha.length; i++) data[i * 4 + 3] = alpha[i]!;
}

function drawWithGlow(
  target: OffscreenCanvasRenderingContext2D,
  glow: NonNullable<SceneAdvancedEffect["glow"]>,
  draw: (surface: OffscreenCanvasRenderingContext2D) => void,
  outline?: SceneObject,
): void {
  if (glow.radius <= 0 || !visible(glow.color)) { draw(target); return; }
  const source = new OffscreenCanvas(target.canvas.width, target.canvas.height);
  const surface = source.getContext("2d");
  const mask = new OffscreenCanvas(source.width, source.height);
  const maskContext = mask.getContext("2d");
  if (surface === null || maskContext === null) { draw(target); return; }
  const transform = target.getTransform();
  surface.setTransform(transform);
  surface.imageSmoothingEnabled = target.imageSmoothingEnabled;
  surface.imageSmoothingQuality = target.imageSmoothingQuality;
  draw(surface);
  const scale = Math.max(Math.hypot(transform.a, transform.b), Math.hypot(transform.c, transform.d));
  const radius = glow.radius * scale;
  maskContext.drawImage(source, 0, 0);
  const geometry = outline === undefined ? undefined : nestedGeometry(outline.visual);
  if (outline !== undefined && geometry !== undefined) {
    // Keep the 3D picture silhouette in the glow mask, not as a colored source stroke.
    maskContext.setTransform(transform);
    maskContext.strokeStyle = "black";
    maskContext.lineWidth = 1;
    traceGeometry(maskContext, outline, geometry);
    maskContext.stroke();
    maskContext.setTransform(1, 0, 0, 1, 0, 0);
  }
  const pixels = maskContext.getImageData(0, 0, source.width, source.height);
  // Spread before feathering: a one-pixel line must glow as visibly as a solid arrow.
  // ponytail: square dilation plus Gaussian feather; use a distance field for exact circular falloff.
  spreadGlowAlpha(pixels, Math.min(Math.max(source.width, source.height), Math.ceil(radius / 4)));
  maskContext.putImageData(pixels, 0, 0);
  target.save();
  try {
    target.setTransform(1, 0, 0, 1, 0, 0);
    // Canvas shadows work in WebKit as well as engines with Canvas filter support.
    target.shadowColor = color(glow.color);
    target.shadowBlur = radius / 2;
    target.shadowOffsetX = source.width;
    target.shadowOffsetY = 0;
    target.drawImage(mask, -source.width, 0);
    target.shadowColor = "transparent";
    target.shadowBlur = 0;
    target.shadowOffsetX = 0;
    target.drawImage(source, 0, 0);
  } finally {
    target.restore();
  }
}

const shadowSurfaces = new WeakMap<
  OffscreenCanvasRenderingContext2D, { source: OffscreenCanvas; blurred: OffscreenCanvas }
>();

function drawWithAdvancedEffects(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  effects: readonly SceneAdvancedEffect[],
  draw: (target: OffscreenCanvasRenderingContext2D) => void,
  threeDMode: "full" | "surface" = "full",
  prepareSurface?: (target: OffscreenCanvasRenderingContext2D) => void,
  strokeWidth?: number,
): void {
  const drawAt = (index: number, target: OffscreenCanvasRenderingContext2D): void => {
    const effect = effects[index];
    if (effect === undefined) {
      if (prepareSurface === undefined) {
        draw(target);
        return;
      }
      target.save();
      try {
        prepareSurface?.(target);
        draw(target);
      } finally {
        target.restore();
      }
      return;
    }
    drawOne(effect, target, (nested) => drawAt(index + 1, nested));
  };
  const drawOne = (
    effect: SceneAdvancedEffect,
    target: OffscreenCanvasRenderingContext2D,
    drawNested: (nested: OffscreenCanvasRenderingContext2D) => void,
  ): void => {
    if (effect.outerShadow !== undefined) {
      const shadow = effect.outerShadow;
      const transform = target.getTransform();
      const horizontal = shadow.alignment % 3;
      const vertical = Math.floor(shadow.alignment / 3);
      const anchorX = object.bounds.x + object.bounds.width * horizontal / 2;
      const anchorY = object.bounds.y + object.bounds.height * vertical / 2;
      let region: Rect | undefined;
      const geometry = effect.visual.kind === "painted-shape" ? effect.visual.geometry : undefined;
      // A stroked polyline has bounded ink, unlike text or nested effects.
      // Keep the full-surface path for those cases until their ink bounds are known.
      if (strokeWidth !== undefined && effects.length === 1 && prepareSurface === undefined
        && (target.filter === undefined || target.filter === "none")
        && typeof geometry === "object" && geometry.kind === "path"
        && geometry.commands.length >= 2) {
        const points = geometry.commands.flatMap(command =>
          command.kind === "moveTo" || command.kind === "lineTo" ? [command] : []);
        if (points.length === geometry.commands.length) {
          let leftX = Infinity, rightX = -Infinity, topY = Infinity, bottomY = -Infinity;
          for (const point of points) {
            leftX = Math.min(leftX, point.x); rightX = Math.max(rightX, point.x);
            topY = Math.min(topY, point.y); bottomY = Math.max(bottomY, point.y);
          }
          // Default Canvas miter limit is ten half-widths; straight segments have no joins.
          const ink = strokeWidth * (points.length === 2 ? 1 : 5);
          const skewX = Math.tan(shadow.skewX * Math.PI / 180);
          const skewY = Math.tan(shadow.skewY * Math.PI / 180);
          const shadowTransform = multiplyTransform(transform, {
            a: shadow.scaleX, b: skewY, c: skewX, d: shadow.scaleY,
            e: anchorX + shadow.offsetX - shadow.scaleX * anchorX - skewX * anchorY,
            f: anchorY + shadow.offsetY - skewY * anchorX - shadow.scaleY * anchorY,
          });
          const bounds = transformedBounds({ ...object, bounds: {
            x: object.bounds.x + leftX - ink,
            y: object.bounds.y + topY - ink,
            width: rightX - leftX + ink * 2,
            height: bottomY - topY + ink * 2,
          } }, shadowTransform);
          const padding = Math.ceil(shadow.blur * Math.max(
            Math.hypot(transform.a, transform.b), Math.hypot(transform.c, transform.d),
          ) * 2) + 2;
          const left = Math.max(0, Math.floor(bounds.x) - padding);
          const top = Math.max(0, Math.floor(bounds.y) - padding);
          const right = Math.min(target.canvas.width, Math.ceil(bounds.x + bounds.width) + padding);
          const bottom = Math.min(target.canvas.height, Math.ceil(bounds.y + bounds.height) + padding);
          region = { x: left, y: top, width: Math.max(0, right - left), height: Math.max(0, bottom - top) };
        }
      }
      let scratch: OffscreenCanvas | undefined;
      let blurred: OffscreenCanvas | undefined;
      if (region !== undefined && region.width > 0 && region.height > 0) {
        const surfaces = shadowSurfaces.get(target) ?? {
          source: new OffscreenCanvas(region.width, region.height),
          blurred: new OffscreenCanvas(region.width, region.height),
        };
        for (const surface of [surfaces.source, surfaces.blurred]) {
          if (surface.width !== region.width) surface.width = region.width;
          if (surface.height !== region.height) surface.height = region.height;
        }
        shadowSurfaces.set(target, surfaces);
        scratch = surfaces.source;
        blurred = surfaces.blurred;
      }
      const shadowCanvas = region?.width === 0 || region?.height === 0 ? undefined
        : colorizedLayer(target, shadow.color, (shadowContext) => {
          shadowContext.setTransform(transform.a, transform.b, transform.c, transform.d,
            transform.e - (region?.x ?? 0), transform.f - (region?.y ?? 0));
          shadowContext.imageSmoothingEnabled = target.imageSmoothingEnabled;
          shadowContext.imageSmoothingQuality = target.imageSmoothingQuality;
          shadowContext.translate(
            anchorX + shadow.offsetX,
            anchorY + shadow.offsetY,
          );
          shadowContext.transform(
            shadow.scaleX,
            Math.tan(shadow.skewY * Math.PI / 180),
            Math.tan(shadow.skewX * Math.PI / 180),
            shadow.scaleY,
            0,
            0,
          );
          shadowContext.translate(-anchorX, -anchorY);
          drawNested(shadowContext);
        }, scratch);
      if (shadowCanvas !== undefined) {
        target.save();
        try {
          target.setTransform(1, 0, 0, 1, 0, 0);
          if (shadow.blur > 0) {
            const scale = Math.max(
              Math.hypot(transform.a, transform.b),
              Math.hypot(transform.c, transform.d),
            );
            // DrawingML stores a blur radius; CSS blur() takes a standard deviation.
            const blur = shadow.blur * scale / 2;
            // Filtering the destination would still process the entire viewport per segment.
            const local = blurred?.getContext("2d");
            if (local !== undefined && local !== null) {
              local.clearRect(0, 0, local.canvas.width, local.canvas.height);
              local.filter = `blur(${blur}px)`;
              local.drawImage(shadowCanvas, 0, 0);
            } else {
              blurred = undefined;
              target.filter = `${target.filter === "none" ? "" : `${target.filter} `}blur(${blur}px)`;
            }
          } else blurred = undefined;
          target.drawImage(blurred ?? shadowCanvas, region?.x ?? 0, region?.y ?? 0);
        } finally {
          target.restore();
        }
      }
    }
    const drawSurface = (surface: OffscreenCanvasRenderingContext2D): void => {
      surface.save();
      try {
        if (effect.softEdge !== undefined && effect.softEdge > 0) {
          const layer = new OffscreenCanvas(surface.canvas.width, surface.canvas.height);
          const isolated = layer.getContext("2d");
          const mask = new OffscreenCanvas(layer.width, layer.height);
          const alpha = mask.getContext("2d");
          if (isolated === null || alpha === null) { drawNested(surface); return; }
          const transform = surface.getTransform();
          isolated.setTransform(transform);
          isolated.imageSmoothingEnabled = surface.imageSmoothingEnabled;
          isolated.imageSmoothingQuality = surface.imageSmoothingQuality;
          drawNested(isolated);
          isolated.setTransform(1, 0, 0, 1, 0, 0);
          const scale = Math.max(Math.hypot(transform.a, transform.b), Math.hypot(transform.c, transform.d));
          // Feather alpha only: blurring the source colors erases fine image detail.
          // ponytail: Gaussian alpha feather approximates Office's edge profile; use erosion if exact falloff is required.
          // Canvas shadows also work in WebKit, where the 2D filter property is unavailable.
          alpha.shadowColor = "black";
          alpha.shadowBlur = effect.softEdge * scale * 2;
          alpha.shadowOffsetX = layer.width;
          alpha.drawImage(layer, -layer.width, 0);
          isolated.globalCompositeOperation = "destination-in";
          isolated.drawImage(mask, 0, 0);
          surface.setTransform(1, 0, 0, 1, 0, 0);
          surface.drawImage(layer, 0, 0);
        } else {
          drawNested(surface);
        }
      } finally {
        surface.restore();
      }
    };
    const glow = effect.glow;
    const drawGlowingSurface = glow === undefined ? drawSurface
      : (surface: OffscreenCanvasRenderingContext2D): void => drawWithGlow(surface, glow, drawSurface, effect.threeD === undefined ? undefined : object);
    if (effect.threeD === undefined) drawGlowingSurface(target);
    else drawWithThreeD(target, object, effect.threeD, drawGlowingSurface, threeDMode === "surface");
    if (effect.innerShadow !== undefined) {
      const source = colorizedLayer(target, 0x000000ff, drawNested);
      if (source !== undefined) {
        const layer = new OffscreenCanvas(source.width, source.height);
        const mask = layer.getContext("2d");
        if (mask !== null) {
          const transform = target.getTransform();
          const shadow = effect.innerShadow;
          const scale = Math.max(Math.hypot(transform.a, transform.b), Math.hypot(transform.c, transform.d));
          // Inner shadow = source alpha times the inverse of its blurred, shifted alpha.
          // Draw only the shadow, so this also works without Canvas filter in WebKit.
          mask.shadowColor = "black";
          mask.shadowBlur = shadow.blur * scale;
          mask.shadowOffsetX = source.width + transform.a * shadow.offsetX + transform.c * shadow.offsetY;
          mask.shadowOffsetY = transform.b * shadow.offsetX + transform.d * shadow.offsetY;
          mask.drawImage(source, -source.width, 0);
          mask.shadowColor = "transparent";
          mask.globalCompositeOperation = "source-out";
          mask.fillStyle = color(shadow.color);
          mask.fillRect(0, 0, layer.width, layer.height);
          mask.globalCompositeOperation = "destination-in";
          mask.drawImage(source, 0, 0);
          target.save();
          try {
            target.setTransform(1, 0, 0, 1, 0, 0);
            target.drawImage(layer, 0, 0);
          } finally { target.restore(); }
        }
      }
    }
    const reflection: SceneReflection | undefined = effect.reflection;
    if (reflection === undefined
      || reflection.startOpacity === 0 && reflection.endOpacity === 0) return;
    const paintReflection = (
      reflectionContext: OffscreenCanvasRenderingContext2D,
      opacity: number,
    ): void => {
      reflectionContext.save();
      try {
        const centerX = object.bounds.x + object.bounds.width / 2;
        const baseline = object.bounds.y + object.bounds.height;
        reflectionContext.translate(centerX, baseline + reflection.distance);
        reflectionContext.scale(reflection.scaleX, -reflection.scaleY);
        reflectionContext.translate(-centerX, -baseline);
        reflectionContext.globalAlpha *= opacity;
        const filters: string[] = [];
        if (reflection.blur > 0) filters.push(`blur(${reflection.blur}px)`);
        if (filters.length > 0) reflectionContext.filter = filters.join(" ");
        if (glow !== undefined && effect.threeD === undefined) drawWithGlow(reflectionContext, glow, drawNested);
        else drawNested(reflectionContext);
      } finally {
        reflectionContext.restore();
      }
    };
    if (reflection.startOpacity === reflection.endOpacity) {
      paintReflection(target, reflection.startOpacity);
      return;
    }
    const reflectedBounds = {
      x: object.bounds.x + object.bounds.width * (1 - reflection.scaleX) / 2,
      y: object.bounds.y + object.bounds.height + reflection.distance,
      width: object.bounds.width * reflection.scaleX,
      height: object.bounds.height * reflection.scaleY,
    };
    // The destination-in fade clips to this rectangle. Allocate only its device
    // bounds, rather than a full slide for every reflected text fragment.
    const transform = target.getTransform();
    const deviceBounds = transformedBounds({ ...object, bounds: reflectedBounds }, transform);
    const left = Math.max(0, Math.floor(deviceBounds.x) - 1);
    const top = Math.max(0, Math.floor(deviceBounds.y) - 1);
    const right = Math.min(target.canvas.width, Math.ceil(deviceBounds.x + deviceBounds.width) + 1);
    const bottom = Math.min(target.canvas.height, Math.ceil(deviceBounds.y + deviceBounds.height) + 1);
    if (right <= left || bottom <= top) return;
    const reflectionCanvas = new OffscreenCanvas(right - left, bottom - top);
    const reflectionContext = reflectionCanvas.getContext("2d");
    if (reflectionContext === null) return;
    reflectionContext.setTransform(transform.a, transform.b, transform.c, transform.d,
      transform.e - left, transform.f - top);
    reflectionContext.imageSmoothingEnabled = target.imageSmoothingEnabled;
    reflectionContext.imageSmoothingQuality = target.imageSmoothingQuality;
    reflectionContext.shadowColor = target.shadowColor;
    reflectionContext.shadowBlur = target.shadowBlur;
    reflectionContext.shadowOffsetX = target.shadowOffsetX;
    reflectionContext.shadowOffsetY = target.shadowOffsetY;
    paintReflection(reflectionContext, 1);
    const radians = reflection.directionDegrees * Math.PI / 180;
    const axisX = Math.cos(radians);
    const axisY = Math.sin(radians);
    const axisLength = Math.abs(reflectedBounds.width * axisX)
      + Math.abs(reflectedBounds.height * axisY);
    const centerX = reflectedBounds.x + reflectedBounds.width / 2;
    const centerY = reflectedBounds.y + reflectedBounds.height / 2;
    const gradient = reflectionContext.createLinearGradient(
      centerX - axisX * axisLength / 2,
      centerY - axisY * axisLength / 2,
      centerX + axisX * axisLength / 2,
      centerY + axisY * axisLength / 2,
    );
    gradient.addColorStop(0, `rgba(0, 0, 0, ${reflection.startOpacity})`);
    gradient.addColorStop(
      reflection.startPosition,
      `rgba(0, 0, 0, ${reflection.startOpacity})`,
    );
    gradient.addColorStop(reflection.endPosition, `rgba(0, 0, 0, ${reflection.endOpacity})`);
    gradient.addColorStop(1, `rgba(0, 0, 0, ${reflection.endOpacity})`);
    reflectionContext.save();
    try {
      reflectionContext.globalCompositeOperation = "destination-in";
      reflectionContext.filter = "none";
      reflectionContext.globalAlpha = 1;
      reflectionContext.fillStyle = gradient;
      reflectionContext.fillRect(
        reflectedBounds.x,
        reflectedBounds.y,
        reflectedBounds.width,
        reflectedBounds.height,
      );
    } finally {
      reflectionContext.restore();
    }
    target.save();
    try {
      target.setTransform(1, 0, 0, 1, 0, 0);
      target.drawImage(reflectionCanvas, left, top);
    } finally {
      target.restore();
    }
  };
  drawAt(0, context);
}

function transformedBounds(object: SceneObject, transform: SceneAffineTransform): Viewport {
  const { x, y, width, height } = object.bounds;
  const points = [
    [x, y],
    [x + width, y],
    [x, y + height],
    [x + width, y + height],
  ].map(([pointX = 0, pointY = 0]) => ({
    x: transform.a * pointX + transform.c * pointY + transform.e,
    y: transform.b * pointX + transform.d * pointY + transform.f,
  }));
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}

function paintedBounds(object: SceneObject, state: VisualState): Viewport {
  let bounds = transformedBounds(object, state.transform);
  const scaleX = Math.hypot(state.transform.a, state.transform.b);
  const scaleY = Math.hypot(state.transform.c, state.transform.d);
  let outsetX = 0;
  let outsetY = 0;
  for (const operation of state.operations) {
    if (operation.visual.kind === "effect" && operation.visual.shadow !== undefined) {
      outsetX = Math.max(
        outsetX,
        (Math.abs(operation.visual.shadow.offsetX) + operation.visual.shadow.blur * 2) * Math.max(1, scaleX),
      );
      outsetY = Math.max(
        outsetY,
        (Math.abs(operation.visual.shadow.offsetY) + operation.visual.shadow.blur * 2) * Math.max(1, scaleY),
      );
    } else if (operation.visual.kind === "advanced-effect") {
      if (operation.visual.outerShadow !== undefined) {
        const shadow = operation.visual.outerShadow;
        outsetX = Math.max(
          outsetX,
          (Math.abs(shadow.offsetX) + shadow.blur * 2
            + bounds.width * Math.max(0, Math.abs(shadow.scaleX) - 1)
            + bounds.height * Math.abs(Math.tan(shadow.skewX * Math.PI / 180))) * Math.max(1, scaleX),
        );
        outsetY = Math.max(
          outsetY,
          (Math.abs(shadow.offsetY) + shadow.blur * 2
            + bounds.height * Math.max(0, Math.abs(shadow.scaleY) - 1)
            + bounds.width * Math.abs(Math.tan(shadow.skewY * Math.PI / 180))) * Math.max(1, scaleY),
        );
      }
      if (operation.visual.glow !== undefined) {
        outsetX = Math.max(outsetX, operation.visual.glow.radius * 2 * Math.max(1, scaleX));
        outsetY = Math.max(outsetY, operation.visual.glow.radius * 2 * Math.max(1, scaleY));
      }
      if (operation.visual.softEdge !== undefined) {
        outsetX = Math.max(outsetX, operation.visual.softEdge * 2 * Math.max(1, scaleX));
        outsetY = Math.max(outsetY, operation.visual.softEdge * 2 * Math.max(1, scaleY));
      }
      if (operation.visual.reflection !== undefined) {
        outsetY = Math.max(
          outsetY,
          (bounds.height * operation.visual.reflection.scaleY
            + operation.visual.reflection.distance
            + operation.visual.reflection.blur * 2) * Math.max(1, scaleY),
        );
      }
      if (operation.visual.threeD !== undefined) {
        const [a, b, c, d, g, h] = threeDCameraMatrix(operation.visual.threeD);
        const centerX = object.bounds.x + object.bounds.width / 2;
        const centerY = object.bounds.y + object.bounds.height / 2;
        for (const x of [-object.bounds.width / 2, object.bounds.width / 2]) {
          for (const y of [-object.bounds.height / 2, object.bounds.height / 2]) {
            const w = 1 + g * x + h * y;
            if (w <= 0.01) continue;
            const px = centerX + (a * x + c * y) / w, py = centerY + (b * x + d * y) / w;
            const tx = state.transform.a * px + state.transform.c * py + state.transform.e;
            const ty = state.transform.b * px + state.transform.d * py + state.transform.f;
            const minX = Math.min(bounds.x, tx), minY = Math.min(bounds.y, ty);
            bounds = { x: minX, y: minY, width: Math.max(bounds.x + bounds.width, tx) - minX,
              height: Math.max(bounds.y + bounds.height, ty) - minY };
          }
        }
        const [offsetX, offsetY] = threeDProjection(operation.visual.threeD);
        const bevel = Math.max(
          operation.visual.threeD.bevelTop?.width ?? 0,
          operation.visual.threeD.bevelTop?.height ?? 0,
          operation.visual.threeD.bevelBottom?.width ?? 0,
          operation.visual.threeD.bevelBottom?.height ?? 0,
          operation.visual.threeD.contourWidth,
        );
        outsetX = Math.max(outsetX, (Math.abs(offsetX) + bevel) * Math.max(1, scaleX));
        outsetY = Math.max(outsetY, (Math.abs(offsetY) + bevel) * Math.max(1, scaleY));
      }
    }
  }
  return {
    x: bounds.x - outsetX,
    y: bounds.y - outsetY,
    width: bounds.width + outsetX * 2,
    height: bounds.height + outsetY * 2,
  };
}

function intersects(a: Viewport, b: Viewport): boolean {
  return a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y;
}

function sheetSizeCacheKey(request: RenderRequest): string {
  const axisKey = (overrides: readonly SheetSizeOverride[] | undefined): string => (
    [...(overrides ?? [])]
      .sort((left, right) => left.index - right.index)
      .map(({ index, size }) => `${index}:${size}`)
      .join(",")
  );
  return `${axisKey(request.sheetSizes?.rows)}|${axisKey(request.sheetSizes?.columns)}`;
}

function firstObjectStartingAtOrAfter(objects: readonly PreparedSceneObject[], y: number): number {
  let low = 0;
  let high = objects.length;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    if (objects[middle]!.bounds.y < y) low = middle + 1;
    else high = middle;
  }
  return low;
}

function firstMaximumBottomAfter(maximumBottoms: Float64Array, y: number): number {
  let low = 0;
  let high = maximumBottoms.length;
  while (low < high) {
    const middle = Math.floor((low + high) / 2);
    if (maximumBottoms[middle]! <= y) low = middle + 1;
    else high = middle;
  }
  return low;
}

function validateNumber(value: number, name: string, allowZero = false): void {
  if (!Number.isFinite(value) || (allowZero ? value < 0 : value <= 0)) {
    throw new OfficeEngineError("INVALID_RENDER_REQUEST", `${name} must be a finite ${allowZero ? "non-negative" : "positive"} number`);
  }
}

function drawWatermark(
  context: OffscreenCanvasRenderingContext2D,
  width: number,
  height: number,
  text: string,
  pixelRatio: number,
): void {
  const spacingX = 240 * pixelRatio;
  const spacingY = 110 * pixelRatio;
  context.save();
  try {
    context.globalAlpha = 0.18;
    context.globalCompositeOperation = "source-over";
    context.fillStyle = "#5c687a";
    context.font = `italic 600 ${14 * pixelRatio}px sans-serif`;
    context.textAlign = "center";
    context.textBaseline = "middle";
    for (let y = Math.min(height / 2, spacingY / 2); y < height; y += spacingY) {
      for (let x = Math.min(width / 2, spacingX / 2); x < width; x += spacingX) {
        context.save();
        context.translate(x, y);
        context.rotate(-Math.PI / 7);
        context.fillText(text, 0, 0, spacingX * 0.8);
        context.restore();
      }
    }
  } finally {
    context.restore();
  }
}

function traceGeometry(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  geometry: SceneGeometry,
): CanvasFillRule {
  const { x, y, width, height } = object.bounds;
  context.beginPath();
  if (geometry === "ellipse") {
    context.ellipse(x + width / 2, y + height / 2, width / 2, height / 2, 0, 0, Math.PI * 2);
  } else if (geometry === "line") {
    context.moveTo(x, y);
    context.lineTo(x + width, y + height);
  } else if (geometry === "rectangle") {
    context.rect(x, y, width, height);
  } else if (geometry.kind === "rounded-rectangle") {
    const radiusX = Math.min(geometry.radiusX, width / 2);
    const radiusY = Math.min(geometry.radiusY, height / 2);
    const kappa = 0.5522847498307936;
    context.moveTo(x + radiusX, y);
    context.lineTo(x + width - radiusX, y);
    context.bezierCurveTo(x + width - radiusX + radiusX * kappa, y, x + width, y + radiusY - radiusY * kappa, x + width, y + radiusY);
    context.lineTo(x + width, y + height - radiusY);
    context.bezierCurveTo(x + width, y + height - radiusY + radiusY * kappa, x + width - radiusX + radiusX * kappa, y + height, x + width - radiusX, y + height);
    context.lineTo(x + radiusX, y + height);
    context.bezierCurveTo(x + radiusX - radiusX * kappa, y + height, x, y + height - radiusY + radiusY * kappa, x, y + height - radiusY);
    context.lineTo(x, y + radiusY);
    context.bezierCurveTo(x, y + radiusY - radiusY * kappa, x + radiusX - radiusX * kappa, y, x + radiusX, y);
    context.closePath();
  } else if (geometry.kind === "layered-path") {
    for (const layer of geometry.layers) tracePathCommands(context, object, layer.commands);
  } else {
    tracePathCommands(context, object, geometry.commands);
  }
  return typeof geometry === "object" && geometry.kind === "path" ? geometry.fillRule : "nonzero";
}

function path2d(object: SceneObject, geometry: Extract<SceneGeometry, { readonly kind: "path" }>): Path2D {
  const path = new Path2D();
  tracePathCommands(path, object, geometry.commands);
  return path;
}

function tracePathCommands(
  context: CanvasPath,
  object: SceneObject,
  commands: readonly ScenePathCommand[],
): void {
  const { x, y } = object.bounds;
  for (const command of commands) {
    if (command.kind === "moveTo") context.moveTo(x + command.x, y + command.y);
    else if (command.kind === "lineTo") context.lineTo(x + command.x, y + command.y);
    else if (command.kind === "quadraticCurveTo") {
      context.quadraticCurveTo(x + command.cpx, y + command.cpy, x + command.x, y + command.y);
    } else if (command.kind === "bezierCurveTo") {
      context.bezierCurveTo(
        x + command.cp1x,
        y + command.cp1y,
        x + command.cp2x,
        y + command.cp2y,
        x + command.x,
        y + command.y,
      );
    } else context.closePath();
  }
}

function geometrySubpaths(commands: readonly ScenePathCommand[]): ScenePathCommand[][] {
  const subpaths: ScenePathCommand[][] = [];
  let current: ScenePathCommand[] = [];
  for (const command of commands) {
    if (command.kind === "moveTo" && current.length !== 0) {
      subpaths.push(current);
      current = [];
    }
    current.push(command);
    if (command.kind === "closePath") {
      subpaths.push(current);
      current = [];
    }
  }
  if (current.length !== 0) subpaths.push(current);
  return subpaths;
}

function isClosedLineEndSubpath(commands: readonly ScenePathCommand[]): boolean {
  const lineCount = commands.filter(({ kind }) => kind === "lineTo").length;
  const curveCount = commands.filter(({ kind }) => kind === "bezierCurveTo").length;
  return commands[0]?.kind === "moveTo"
    && commands.at(-1)?.kind === "closePath"
    && ((lineCount >= 2
      && lineCount <= 4
      && commands.every(({ kind }) => kind === "moveTo" || kind === "lineTo" || kind === "closePath"))
      || (curveCount === 4
        && commands.every(({ kind }) => kind === "moveTo" || kind === "bezierCurveTo" || kind === "closePath")));
}

function attachedClosedLineEnds(
  subpaths: readonly (readonly ScenePathCommand[])[],
): readonly (readonly ScenePathCommand[])[] {
  const firstOpenSubpath = subpaths.findIndex((commands) => commands.at(-1)?.kind !== "closePath");
  if (firstOpenSubpath < 0) return [];
  const endpoints = subpaths
    .filter((commands) => commands.at(-1)?.kind !== "closePath")
    .flatMap((commands) => {
      const first = commands[0];
      const last = commands.at(-1);
      return [first, last].filter(
        (command): command is Exclude<ScenePathCommand, { readonly kind: "closePath" }> =>
          command !== undefined && command.kind !== "closePath",
      );
    });
  return subpaths.filter((commands, index) => {
    // Connector line ends are appended after their open shaft path. Preset
    // outlines such as bevel also contain closed polygons whose vertices touch
    // later open detail lines; those polygons are not arrowheads and must not
    // be filled with the stroke color.
    if (index <= firstOpenSubpath) return false;
    if (!isClosedLineEndSubpath(commands)) return false;
    const tip = commands[0];
    return tip?.kind === "moveTo"
      && endpoints.some(({ x, y }) => Math.abs(x - tip.x) < 0.01 && Math.abs(y - tip.y) < 0.01);
  });
}

function closedLineEndAttachment(
  commands: readonly ScenePathCommand[],
): { readonly tip: { readonly x: number; readonly y: number }; readonly x: number; readonly y: number } | undefined {
  const tip = commands[0];
  const firstSide = commands[1];
  const centerOrSecondSide = commands[2];
  if (tip?.kind !== "moveTo"
    || (firstSide?.kind !== "lineTo" && firstSide?.kind !== "bezierCurveTo")
    || (centerOrSecondSide?.kind !== "lineTo" && centerOrSecondSide?.kind !== "bezierCurveTo")) return undefined;
  if (commands.length === 6 && firstSide.kind === "bezierCurveTo"
    && centerOrSecondSide.kind === "bezierCurveTo") {
    return {
      tip,
      x: (tip.x + centerOrSecondSide.x) / 2,
      y: (tip.y + centerOrSecondSide.y) / 2,
    };
  }
  if (commands.length === 4) {
    return {
      tip,
      x: (firstSide.x + centerOrSecondSide.x) / 2,
      y: (firstSide.y + centerOrSecondSide.y) / 2,
    };
  }
  return { tip, x: centerOrSecondSide.x, y: centerOrSecondSide.y };
}

function fillStrokeColoredLineEnds(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  geometry: SceneGeometry,
  strokeStyle: string | CanvasGradient | CanvasPattern,
  groupedScale = 1,
): void {
  if (typeof geometry !== "object" || geometry.kind !== "path") return;
  const subpaths = geometrySubpaths(geometry.commands);
  const closedLineEnds = attachedClosedLineEnds(subpaths);
  if (closedLineEnds.length === 0
    || !subpaths.some((commands) => !isClosedLineEndSubpath(commands))) return;

  const { x, y } = object.bounds;
  context.beginPath();
  for (const commands of closedLineEnds) {
    const attachment = closedLineEndAttachment(commands);
    const scale = Number.isFinite(groupedScale) && groupedScale > 0 ? groupedScale : 1;
    const point = (pointX: number, pointY: number): readonly [number, number] => [
      x + (attachment?.tip.x ?? 0) + (pointX - (attachment?.tip.x ?? 0)) / scale,
      y + (attachment?.tip.y ?? 0) + (pointY - (attachment?.tip.y ?? 0)) / scale,
    ];
    for (const command of commands) {
      const [pointX, pointY] = command.kind === "closePath"
        ? [0, 0]
        : point(command.x, command.y);
      if (command.kind === "moveTo") context.moveTo(pointX, pointY);
      else if (command.kind === "lineTo") context.lineTo(pointX, pointY);
      else if (command.kind === "quadraticCurveTo") {
        const [controlX, controlY] = point(command.cpx, command.cpy);
        context.quadraticCurveTo(
          controlX,
          controlY,
          pointX,
          pointY,
        );
      } else if (command.kind === "bezierCurveTo") {
        const [control1X, control1Y] = point(command.cp1x, command.cp1y);
        const [control2X, control2Y] = point(command.cp2x, command.cp2y);
        context.bezierCurveTo(
          control1X,
          control1Y,
          control2X,
          control2Y,
          pointX,
          pointY,
        );
      } else context.closePath();
    }
  }
  context.fillStyle = strokeStyle;
  context.fill("nonzero");
}

interface CanvasAffineTransform {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

type TextFragmentCollector = (
  fragment: Omit<RenderedTextFragment, "transform">,
  transform: CanvasAffineTransform,
) => void;

function semanticTextFragments(
  fragments: readonly RenderedTextFragment[],
  textByObject: ReadonlyMap<string, string>,
): readonly RenderedTextFragment[] {
  const visualLengths = new Map<string, number>();
  for (const fragment of fragments) {
    visualLengths.set(
      fragment.objectId,
      (visualLengths.get(fragment.objectId) ?? 0) + [...fragment.text].length,
    );
  }
  const offsets = new Map<string, number>();
  return fragments.map((fragment) => {
    const semantic = textByObject.get(fragment.objectId);
    const offset = offsets.get(fragment.objectId) ?? 0;
    const length = [...fragment.text].length;
    offsets.set(fragment.objectId, offset + length);
    const total = visualLengths.get(fragment.objectId) ?? 0;
    if (semantic === undefined || total === 0) return fragment;
    const characters = [...semantic];
    const start = Math.round(offset / total * characters.length);
    const end = Math.round((offset + length) / total * characters.length);
    return { ...fragment, text: characters.slice(start, end).join("") };
  });
}

interface StrokePathPoint {
  readonly x: number;
  readonly y: number;
}

interface PixelSnappedStrokePath {
  readonly points: readonly StrokePathPoint[];
  readonly lineWidth: number;
}

function canvasTransform(
  context: OffscreenCanvasRenderingContext2D,
): CanvasAffineTransform | undefined {
  const getTransform = (context as OffscreenCanvasRenderingContext2D & {
    getTransform?: () => CanvasAffineTransform;
  }).getTransform;
  if (typeof getTransform !== "function") return undefined;
  try {
    const transform = getTransform.call(context);
    const values = [
      transform.a,
      transform.b,
      transform.c,
      transform.d,
      transform.e,
      transform.f,
    ];
    if (!values.every(Number.isFinite)) return undefined;
    const magnitude = Math.max(1, ...values.slice(0, 4).map(Math.abs));
    const determinant = transform.a * transform.d - transform.b * transform.c;
    return Math.abs(determinant) > Number.EPSILON * magnitude * magnitude
      ? transform
      : undefined;
  } catch {
    return undefined;
  }
}

function currentCanvasTransform(
  context: OffscreenCanvasRenderingContext2D,
): CanvasAffineTransform | undefined {
  const transform = canvasTransform(context);
  if (transform === undefined) return undefined;
  const magnitude = Math.max(1, ...[
    transform.a,
    transform.b,
    transform.c,
    transform.d,
  ].map(Math.abs));
  const axisTolerance = magnitude * 1e-7;
  const axisAligned = (Math.abs(transform.b) <= axisTolerance
      && Math.abs(transform.c) <= axisTolerance)
    || (Math.abs(transform.a) <= axisTolerance
      && Math.abs(transform.d) <= axisTolerance);
  return axisAligned ? transform : undefined;
}

function effectiveStrokeWidth(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  authoredWidth: number,
): number {
  if (object.source.format !== "pdf") return authoredWidth;
  const transform = canvasTransform(context);
  if (transform === undefined) return authoredWidth > 0 ? authoredWidth : 1;
  const determinant = Math.abs(transform.a * transform.d - transform.b * transform.c);
  const singlePixelWidth = Math.max(
    Math.hypot(transform.a, transform.b),
    Math.hypot(transform.c, transform.d),
  ) / determinant;
  if (!Number.isFinite(singlePixelWidth) || singlePixelWidth <= Number.EPSILON) {
    return authoredWidth > 0 ? authoredWidth : 1;
  }
  // Match PDF.js' rescaleAndStroke contract: every PDF stroke, including a
  // non-zero authored hairline, must cover at least one device pixel after the
  // current transform. Canvas otherwise turns engineering-drawing strokes into
  // faint fractional-coverage lines.
  return Math.max(authoredWidth, singlePixelWidth);
}

function pixelSnappedStrokePath(
  points: readonly StrokePathPoint[],
  transform: CanvasAffineTransform,
  lineWidth: number,
  closed: boolean,
): PixelSnappedStrokePath | undefined {
  if (points.length < 2) return undefined;
  const devicePoints = points.map(({ x, y }) => ({
    x: transform.a * x + transform.c * y + transform.e,
    y: transform.b * x + transform.d * y + transform.f,
  }));
  const snapped = devicePoints.map(({ x, y }) => ({ x, y }));
  let snappedLineWidth: number | undefined;
  const segmentCount = closed ? points.length : points.length - 1;
  for (let index = 0; index < segmentCount; index += 1) {
    const nextIndex = (index + 1) % points.length;
    const start = devicePoints[index]!;
    const end = devicePoints[nextIndex]!;
    const deltaX = end.x - start.x;
    const deltaY = end.y - start.y;
    const length = Math.hypot(deltaX, deltaY);
    const axisTolerance = Math.max(1, length) * 1e-7;
    const horizontal = Math.abs(deltaY) <= axisTolerance && Math.abs(deltaX) > axisTolerance;
    const vertical = Math.abs(deltaX) <= axisTolerance && Math.abs(deltaY) > axisTolerance;
    if (!horizontal && !vertical) return undefined;

    const userDeltaX = points[nextIndex]!.x - points[index]!.x;
    const userDeltaY = points[nextIndex]!.y - points[index]!.y;
    const normalScale = Math.abs(userDeltaX) >= Math.abs(userDeltaY)
      ? Math.hypot(transform.c, transform.d)
      : Math.hypot(transform.a, transform.b);
    const deviceLineWidth = lineWidth * normalScale;
    const widthTolerance = Math.max(1, deviceLineWidth) * 1e-3;
    if (normalScale <= Number.EPSILON
      || deviceLineWidth <= 0
      || deviceLineWidth >= 1 - widthTolerance) return undefined;
    const segmentLineWidth = 1 / normalScale;
    const lineWidthTolerance = Math.max(1, segmentLineWidth) * 1e-3;
    if (snappedLineWidth !== undefined
      && Math.abs(segmentLineWidth - snappedLineWidth) > lineWidthTolerance) return undefined;
    snappedLineWidth ??= segmentLineWidth;

    const pixelCenter = 0.5;
    if (horizontal) {
      const coordinate = Math.round((start.y + end.y) / 2 - pixelCenter) + pixelCenter;
      snapped[index]!.y = coordinate;
      snapped[nextIndex]!.y = coordinate;
    } else {
      const coordinate = Math.round((start.x + end.x) / 2 - pixelCenter) + pixelCenter;
      snapped[index]!.x = coordinate;
      snapped[nextIndex]!.x = coordinate;
    }
  }

  const determinant = transform.a * transform.d - transform.b * transform.c;
  const userPoints = snapped.map(({ x, y }) => {
    const translatedX = x - transform.e;
    const translatedY = y - transform.f;
    return {
      x: (transform.d * translatedX - transform.c * translatedY) / determinant,
      y: (-transform.b * translatedX + transform.a * translatedY) / determinant,
    };
  });
  return { points: userPoints, lineWidth: snappedLineWidth ?? lineWidth };
}

function traceStrokePoints(
  context: OffscreenCanvasRenderingContext2D,
  points: readonly StrokePathPoint[],
  closed: boolean,
): void {
  context.moveTo(points[0]!.x, points[0]!.y);
  for (const point of points.slice(1)) context.lineTo(point.x, point.y);
  if (closed) context.closePath();
}

function linearStrokePathPoints(
  commands: readonly ScenePathCommand[],
): readonly StrokePathPoint[] | undefined {
  const points: StrokePathPoint[] = [];
  for (const command of commands) {
    if (command.kind !== "moveTo" && command.kind !== "lineTo") return undefined;
    points.push({ x: command.x, y: command.y });
  }
  return points.length >= 2 ? points : undefined;
}

function retraceStrokeGeometry(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  geometry: SceneGeometry,
  groupedScale = 1,
  snapHairlines = true,
): boolean {
  const transform = snapHairlines ? currentCanvasTransform(context) : undefined;
  const { x, y, width, height } = object.bounds;
  if (transform !== undefined && geometry === "line") {
    const path = pixelSnappedStrokePath(
      [{ x, y }, { x: x + width, y: y + height }],
      transform,
      context.lineWidth,
      false,
    );
    if (path !== undefined) {
      context.lineWidth = path.lineWidth;
      context.beginPath();
      traceStrokePoints(context, path.points, false);
      return true;
    }
  }
  if (transform !== undefined && geometry === "rectangle") {
    const path = pixelSnappedStrokePath(
      [
        { x, y },
        { x: x + width, y },
        { x: x + width, y: y + height },
        { x, y: y + height },
      ],
      transform,
      context.lineWidth,
      true,
    );
    if (path !== undefined) {
      context.lineWidth = path.lineWidth;
      context.beginPath();
      traceStrokePoints(context, path.points, true);
      return true;
    }
  }
  if (typeof geometry !== "object" || geometry.kind !== "path") return false;
  const subpaths = geometrySubpaths(geometry.commands);
  const closedLineEnds = attachedClosedLineEnds(subpaths);
  const strokeSubpaths = subpaths.filter((commands) => !closedLineEnds.includes(commands));
  if (strokeSubpaths.length === 0) return false;

  const scale = Number.isFinite(groupedScale) && groupedScale > 0 ? groupedScale : 1;
  const lineEnds = closedLineEnds
    .map(closedLineEndAttachment)
    .filter((lineEnd) => lineEnd !== undefined)
    .map(({ tip, x, y }) => ({
      tip,
      x: tip.x + (x - tip.x) / scale,
      y: tip.y + (y - tip.y) / scale,
    }));
  const attachmentAt = (pointX: number, pointY: number): { readonly x: number; readonly y: number } | undefined =>
    lineEnds.find(({ tip }) => Math.abs(tip.x - pointX) < 0.01 && Math.abs(tip.y - pointY) < 0.01);
  const strokePointSubpaths = strokeSubpaths.map((commands) => {
    const points = linearStrokePathPoints(commands);
    return points?.map((point, index) => {
      const attached = index === 0 || index === points.length - 1
        ? attachmentAt(point.x, point.y) ?? point
        : point;
      return { x: x + attached.x, y: y + attached.y };
    });
  });
  if (transform !== undefined && strokePointSubpaths.every((points) => points !== undefined)) {
    const snappedSubpaths = strokePointSubpaths.map((points) => pixelSnappedStrokePath(
      points!, transform, context.lineWidth, false,
    ));
    const snappedLineWidth = snappedSubpaths[0]?.lineWidth;
    if (snappedLineWidth !== undefined
      && snappedSubpaths.every((path) => path !== undefined
        && Math.abs(path.lineWidth - snappedLineWidth) <= Math.max(1, snappedLineWidth) * 1e-3)) {
      context.lineWidth = snappedLineWidth;
      context.beginPath();
      for (const path of snappedSubpaths) traceStrokePoints(context, path!.points, false);
      return true;
    }
  }

  if (closedLineEnds.length === 0) return false;
  context.beginPath();
  for (const commands of strokeSubpaths) {
    for (const [index, command] of commands.entries()) {
      const isEndpoint = index === 0 || index === commands.length - 1;
      const attachment = isEndpoint && command.kind !== "closePath"
        ? attachmentAt(command.x, command.y)
        : undefined;
      const commandX = attachment?.x ?? (command.kind === "closePath" ? 0 : command.x);
      const commandY = attachment?.y ?? (command.kind === "closePath" ? 0 : command.y);
      if (command.kind === "moveTo") context.moveTo(x + commandX, y + commandY);
      else if (command.kind === "lineTo") context.lineTo(x + commandX, y + commandY);
      else if (command.kind === "quadraticCurveTo") {
        context.quadraticCurveTo(x + command.cpx, y + command.cpy, x + commandX, y + commandY);
      } else if (command.kind === "bezierCurveTo") {
        context.bezierCurveTo(
          x + command.cp1x,
          y + command.cp1y,
          x + command.cp2x,
          y + command.cp2y,
          x + commandX,
          y + commandY,
        );
      } else context.closePath();
    }
  }
  return true;
}

type ScenePathFillMode = Extract<SceneGeometry, { readonly kind: "layered-path" }>["layers"][number]["fill"];

function pathFillColor(
  value: number,
  mode: ScenePathFillMode,
): number {
  const transform = (channel: number): number => {
    if (mode === "darken") return Math.round(channel * 0.6);
    if (mode === "darken-less") return Math.round(channel * 205 / 255);
    if (mode === "lighten") return Math.round(channel + (255 - channel) * 0.4);
    if (mode === "lighten-less") return Math.round(channel + (255 - channel) * 50 / 255);
    return channel;
  };
  const red = transform((value >>> 24) & 0xff);
  const green = transform((value >>> 16) & 0xff);
  const blue = transform((value >>> 8) & 0xff);
  return ((red << 24) | (green << 16) | (blue << 8) | (value & 0xff)) >>> 0;
}

function pathFillPaint(paint: ScenePaint, mode: ScenePathFillMode): ScenePaint {
  if (mode === "normal") return paint;
  if (paint.kind === "mapped-gradient") return { ...paint, paint: pathFillPaint(paint.paint, mode) };
  if (paint.kind === "solid") return { ...paint, color: pathFillColor(paint.color, mode) };
  if (
    paint.kind === "linear-gradient"
    || paint.kind === "radial-gradient"
    || paint.kind === "rect-gradient"
    || paint.kind === "circle-gradient"
    || paint.kind === "shape-gradient"
  ) {
    return {
      ...paint,
      stops: paint.stops.map((stop) => ({ ...stop, color: pathFillColor(stop.color, mode) })),
    };
  }
  if (paint.kind === "pattern") {
    return {
      ...paint,
      foreground: pathFillColor(paint.foreground, mode),
      background: pathFillColor(paint.background, mode),
    };
  }
  return paint;
}

function drawLegacyGeometry(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  visual: Extract<SceneVisual, { kind: "shape" | "text" }>,
  strokeStyle?: SceneStrokeStyle,
  groupedScale = 1,
): void {
  if (typeof visual.geometry === "object" && visual.geometry.kind === "layered-path") {
    // Paint the entire composite surface before any outlines.
    for (const strokePass of [false, true]) {
      for (const layer of visual.geometry.layers) {
        const geometry: SceneGeometry = {
          kind: "path",
          fillRule: layer.fillRule,
          commands: layer.commands,
        };
        if (!strokePass && layer.fill !== "none") {
          drawLegacyGeometry(context, object, {
            ...visual,
            geometry,
            fill: pathFillColor(visual.fill, layer.fill),
            stroke: 0,
          }, strokeStyle, groupedScale);
        }
        if (strokePass && layer.stroke) {
          drawLegacyGeometry(context, object, { ...visual, geometry, fill: 0 }, strokeStyle, groupedScale);
        }
      }
    }
    return;
  }
  const fillRule = traceGeometry(context, object, visual.geometry);
  if (visible(visual.fill) && visual.geometry !== "line") {
    context.fillStyle = color(visual.fill);
    context.fill(fillRule);
  }
  const strokeWidth = effectiveStrokeWidth(context, object, visual.strokeWidth / groupedScale);
  if (visible(visual.stroke) && strokeWidth > 0) {
    const strokeColor = color(visual.stroke);
    strokePaintedGeometry(
      context,
      object,
      visual.geometry,
      fillRule,
      strokeColor,
      strokeWidth,
      strokeStyle,
      undefined,
      groupedScale,
    );
    if (object.source.format !== "pdf" && object.source.format !== "xps") {
      fillStrokeColoredLineEnds(context, object, visual.geometry, strokeColor, groupedScale);
    }
  }
}

function paintStyle(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  paint: ScenePaint,
  originX = 0,
  originY = 0,
  geometry: SceneGeometry = "rectangle",
): string | CanvasGradient | CanvasPattern | undefined {
  if (paint.kind === "mapped-gradient") return mappedGradientPattern(context, object, paint, geometry, originX, originY);
  if (paint.kind === "none") return undefined;
  if (paint.kind === "solid") return color(paint.color);
  if (paint.kind === "pattern") {
    const checker = /^pdf:checker:([0-9.]+):[0-9.]+:(-?[0-9.]+):(-?[0-9.]+)$/u.exec(paint.preset);
    if (checker === null) return drawingmlPattern(context, paint, object.bounds.x - originX, object.bounds.y - originY);
    const spacing = Math.max(1, Math.min(64, Number(checker[1])));
    const tileSize = Math.max(2, Math.ceil(spacing * 2));
    const tile = new OffscreenCanvas(tileSize, tileSize);
    const tileContext = tile.getContext("2d");
    if (tileContext === null) return undefined;
    tileContext.fillStyle = color(paint.background);
    tileContext.fillRect(0, 0, tileSize, tileSize);
    tileContext.fillStyle = color(paint.foreground);
    tileContext.fillRect(0, 0, spacing, spacing);
    tileContext.fillRect(spacing, spacing, spacing, spacing);
    const pattern = context.createPattern(tile, "repeat");
    pattern?.setTransform({ e: Number(checker[2]), f: Number(checker[3]) });
    return pattern ?? undefined;
  }
  if (
    paint.kind === "image"
    || paint.kind === "visual"
    || paint.kind === "xps-gradient"
    || paint.kind === "rect-gradient"
    || paint.kind === "circle-gradient"
    || paint.kind === "shape-gradient"
  ) return undefined;
  const x = object.bounds.x - originX;
  const y = object.bounds.y - originY;
  const gradient = paint.kind === "linear-gradient"
    ? context.createLinearGradient(
        x + paint.start.x,
        y + paint.start.y,
        x + paint.end.x,
        y + paint.end.y,
      )
    : context.createRadialGradient(
        x + paint.start.x,
        y + paint.start.y,
        paint.start.radius,
        x + paint.end.x,
        y + paint.end.y,
        paint.end.radius,
      );
  for (const stop of paint.stops) gradient.addColorStop(stop.offset, color(stop.color));
  return gradient;
}

function reversedConicRadialPaintStyle(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  paint: Extract<ScenePaint, { kind: "radial-gradient" }>,
): CanvasGradient | undefined {
  const centerDistance = Math.hypot(
    paint.start.x - paint.end.x,
    paint.start.y - paint.end.y,
  );
  if (centerDistance + paint.end.radius <= paint.start.radius
    || centerDistance + paint.start.radius <= paint.end.radius) return undefined;
  const { x, y } = object.bounds;
  const gradient = context.createRadialGradient(
    x + paint.end.x,
    y + paint.end.y,
    paint.end.radius,
    x + paint.start.x,
    y + paint.start.y,
    paint.start.radius,
  );
  for (const stop of [...paint.stops].reverse()) {
    gradient.addColorStop(1 - stop.offset, color(stop.color));
  }
  return gradient;
}

function mappedGradientPattern(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  paint: Extract<ScenePaint, { kind: "mapped-gradient" }>,
  geometry: SceneGeometry,
  originX: number,
  originY: number,
): CanvasPattern | undefined {
  const { width, height } = object.bounds;
  if (!(width > 0 && height > 0)) return undefined;
  const transform = canvasTransform(context) ?? IDENTITY_TRANSFORM;
  const raster = gradientRasterSize(width, height, Math.hypot(transform.a, transform.b), Math.hypot(transform.c, transform.d));
  const canvas = new OffscreenCanvas(raster.width, raster.height);
  const target = canvas.getContext("2d");
  if (target === null) return undefined;
  target.scale(raster.scaleX, raster.scaleY);
  const local = { ...object, bounds: { x: 0, y: 0, width, height } };
  drawPaintedGeometry(target, local, { kind: "painted-shape", geometry: paint.paint.kind === "shape-gradient" ? geometry : "rectangle", fill: paint.paint, stroke: { kind: "none" }, strokeWidth: 0 });
  return imagePaintPattern(context, { ...object, bounds: { ...object.bounds, x: object.bounds.x - originX, y: object.bounds.y - originY } }, {
    kind: "image", bytes: new Uint8Array(), mediaType: "image/png",
    cropLeft: 0, cropTop: 0, cropRight: 0, cropBottom: 0, tile: true,
    tileWidth: width * (1 - paint.tile.left - paint.tile.right),
    tileHeight: height * (1 - paint.tile.top - paint.tile.bottom),
    mapping: { scaleX: 1, scaleY: 1, offsetX: width * paint.tile.left, offsetY: height * paint.tile.top,
      alignmentX: 0, alignmentY: 0, flip: paint.flip, rotateWithShape: paint.rotateWithShape,
      left: 0, top: 0, right: 0, bottom: 0, dpi: 0 },
  }, canvas);
}

function imagePixelDensity(bytes: Uint8Array): { x: number; y: number } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const density = (x: number, y: number, unit: number) => ({ x: x > 0 ? x * unit : 96, y: y > 0 ? y * unit : 96 });
  if (bytes.length >= 8 && view.getUint32(0) === 0x89504e47) {
    for (let offset = 8; offset + 12 <= bytes.length;) {
      const length = view.getUint32(offset);
      if (length > bytes.length - offset - 12) break;
      if (view.getUint32(offset + 4) === 0x70485973 && length === 9 && bytes[offset + 16] === 1) {
        return density(view.getUint32(offset + 8), view.getUint32(offset + 12), .0254);
      }
      offset += length + 12;
    }
  } else if (bytes.length >= 4 && view.getUint16(0) === 0xffd8) {
    for (let offset = 2; offset + 4 <= bytes.length;) {
      if (bytes[offset] !== 0xff) break;
      const marker = bytes[offset + 1];
      if (marker === 0xda || marker === 0xd9) break;
      const length = view.getUint16(offset + 2);
      if (length < 2 || length > bytes.length - offset - 2) break;
      if (marker === 0xe0 && length >= 16 && view.getUint32(offset + 4) === 0x4a464946) {
        const unit = bytes[offset + 11];
        if (unit === 1 || unit === 2) return density(view.getUint16(offset + 12), view.getUint16(offset + 14), unit === 1 ? 1 : 2.54);
      }
      offset += length + 2;
    }
  } else if (bytes.length >= 46 && view.getUint16(0) === 0x424d) {
    return density(view.getInt32(38, true), view.getInt32(42, true), .0254);
  }
  return { x: 96, y: 96 };
}

function drawImageRegion(
  context: OffscreenCanvasRenderingContext2D,
  image: ImageBitmap | OffscreenCanvas,
  sx: number, sy: number, sw: number, sh: number,
  dx: number, dy: number, dw: number, dh: number,
): void {
  if (!(sw > 0 && sh > 0)) return;
  // WebKit drops ImageBitmap draws whose source extends beyond the bitmap.
  // Intersect both rectangles proportionally so negative Office crops keep their padding.
  if (sx < 0 || sy < 0 || sx + sw > image.width || sy + sh > image.height) {
    const left = Math.max(0, sx), top = Math.max(0, sy);
    const right = Math.min(image.width, sx + sw), bottom = Math.min(image.height, sy + sh);
    if (right <= left || bottom <= top) return;
    dx += (left - sx) / sw * dw;
    dy += (top - sy) / sh * dh;
    dw *= (right - left) / sw;
    dh *= (bottom - top) / sh;
    sx = left; sy = top; sw = right - left; sh = bottom - top;
  }
  context.drawImage(image, sx, sy, sw, sh, dx, dy, dw, dh);
}

function imagePaintPattern(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  paint: Extract<ScenePaint, { kind: "image" }>,
  bitmap: ImageBitmap | OffscreenCanvas,
  groupedScale = 1,
): CanvasPattern | undefined {
  const mapping = paint.mapping;
  const sourceWidth = bitmap.width * (1 - paint.cropLeft - paint.cropRight);
  const sourceHeight = bitmap.height * (1 - paint.cropTop - paint.cropBottom);
  if (!(sourceWidth > 0 && sourceHeight > 0)) return undefined;
  const flipX = mapping?.flip === "flip-x" || mapping?.flip === "flip-xy";
  const flipY = mapping?.flip === "flip-y" || mapping?.flip === "flip-xy";
  const columns = flipX ? 2 : 1, rows = flipY ? 2 : 1;
  const raster = gradientRasterSize(sourceWidth * columns, sourceHeight * rows, 1, 1);
  const tile = new OffscreenCanvas(raster.width, raster.height);
  const target = tile.getContext("2d");
  if (target === null) return undefined;
  const cellWidth = tile.width / columns, cellHeight = tile.height / rows;
  for (let row = 0; row < rows; row++) {
    for (let column = 0; column < columns; column++) {
      target.save();
      target.translate(column === 1 ? cellWidth * 2 : 0, row === 1 ? cellHeight * 2 : 0);
      target.scale(column === 1 ? -1 : 1, row === 1 ? -1 : 1);
      drawImageRegion(target, bitmap, bitmap.width * paint.cropLeft, bitmap.height * paint.cropTop,
        sourceWidth, sourceHeight, 0, 0, cellWidth, cellHeight);
      target.restore();
    }
  }
  const density = mapping?.dpi ? { x: mapping.dpi, y: mapping.dpi } : imagePixelDensity(paint.bytes);
  const width = paint.tile
    ? (paint.tileWidth ?? sourceWidth * 96 / density.x * (mapping?.scaleX ?? 1)) / groupedScale
    : object.bounds.width * (1 - (mapping?.left ?? 0) - (mapping?.right ?? 0));
  const height = paint.tile
    ? (paint.tileHeight ?? sourceHeight * 96 / density.y * (mapping?.scaleY ?? 1)) / groupedScale
    : object.bounds.height * (1 - (mapping?.top ?? 0) - (mapping?.bottom ?? 0));
  if (!(width > 0 && height > 0)) return undefined;
  const x = object.bounds.x + (paint.tile
    ? (object.bounds.width - width) * (mapping?.alignmentX ?? 0) + (mapping?.offsetX ?? 0) / groupedScale
    : object.bounds.width * (mapping?.left ?? 0));
  const y = object.bounds.y + (paint.tile
    ? (object.bounds.height - height) * (mapping?.alignmentY ?? 0) + (mapping?.offsetY ?? 0) / groupedScale
    : object.bounds.height * (mapping?.top ?? 0));
  let transform: SceneAffineTransform = { a: width / cellWidth, b: 0, c: 0, d: height / cellHeight, e: x, f: y };
  if (mapping?.rotateWithShape === false) {
    const own = visualTransform(object.visual);
    const angle = -Math.atan2(own.b, own.a), cosine = Math.cos(angle), sine = Math.sin(angle);
    const cx = object.bounds.x + object.bounds.width / 2, cy = object.bounds.y + object.bounds.height / 2;
    transform = multiplyTransform({ a: cosine, b: sine, c: -sine, d: cosine,
      e: cx - cosine * cx + sine * cy, f: cy - sine * cx - cosine * cy }, transform);
  }
  const pattern = context.createPattern(tile, paint.tile ? "repeat" : "no-repeat");
  pattern?.setTransform(transform);
  return pattern ?? undefined;
}

function drawImagePaint(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  geometry: SceneGeometry,
  fillRule: CanvasFillRule,
  paint: Extract<ScenePaint, { kind: "image" }>,
  bitmap: ImageBitmap,
  groupedScale = 1,
): void {
  if (paint.mapping !== undefined) {
    const pattern = imagePaintPattern(context, object, paint, bitmap, groupedScale);
    if (pattern !== undefined) { context.fillStyle = pattern; context.fill(fillRule); }
    return;
  }
  const { cropLeft, cropTop, cropRight, cropBottom } = paint;
  const sourceX = bitmap.width * cropLeft;
  const sourceY = bitmap.height * cropTop;
  const sourceWidth = bitmap.width * (1 - cropLeft - cropRight);
  const sourceHeight = bitmap.height * (1 - cropTop - cropBottom);
  if (sourceWidth <= 0 || sourceHeight <= 0) return;
  context.save();
  try {
    traceGeometry(context, object, geometry);
    context.clip(fillRule);
    if (!paint.tile) {
      drawImageRegion(context,
        bitmap,
        sourceX,
        sourceY,
        sourceWidth,
        sourceHeight,
        object.bounds.x,
        object.bounds.y,
        object.bounds.width,
        object.bounds.height,
      );
      return;
    }
    const tileWidth = Math.max(1, paint.tileWidth ?? sourceWidth);
    const tileHeight = Math.max(1, paint.tileHeight ?? sourceHeight);
    for (let y = object.bounds.y; y < object.bounds.y + object.bounds.height; y += tileHeight) {
      for (let x = object.bounds.x; x < object.bounds.x + object.bounds.width; x += tileWidth) {
        drawImageRegion(context,
          bitmap,
          sourceX,
          sourceY,
          sourceWidth,
          sourceHeight,
          x,
          y,
          tileWidth + 1,
          tileHeight + 1,
        );
      }
    }
  } finally {
    context.restore();
  }
}

function progressivelyDownscaleImage(
  bitmap: ImageBitmap,
  renderedWidth: number,
  renderedHeight: number,
): { readonly source: ImageBitmap | OffscreenCanvas; readonly width: number; readonly height: number } {
  let source: ImageBitmap | OffscreenCanvas = bitmap;
  let width = bitmap.width;
  let height = bitmap.height;
  let widthScale = width / Math.max(renderedWidth, 1);
  let heightScale = height / Math.max(renderedHeight, 1);
  while ((widthScale > 2 && width > 1) || (heightScale > 2 && height > 1)) {
    const nextWidth = widthScale > 2 && width > 1 ? Math.ceil(width / 2) : width;
    const nextHeight = heightScale > 2 && height > 1 ? Math.ceil(height / 2) : height;
    const canvas = new OffscreenCanvas(nextWidth, nextHeight);
    const context = canvas.getContext("2d");
    if (context === null) break;
    context.drawImage(source, 0, 0, width, height, 0, 0, nextWidth, nextHeight);
    widthScale /= width / nextWidth;
    heightScale /= height / nextHeight;
    source = canvas;
    width = nextWidth;
    height = nextHeight;
  }
  return { source, width, height };
}

function drawPdfImageAtIntegerCoordinates(
  context: OffscreenCanvasRenderingContext2D,
  source: CanvasImageSource,
  sourceX: number,
  sourceY: number,
  sourceWidth: number,
  sourceHeight: number,
  destinationX: number,
  destinationY: number,
  destinationWidth: number,
  destinationHeight: number,
): void {
  const transform = currentCanvasTransform(context);
  if (transform === undefined
    || typeof (context as { setTransform?: unknown }).setTransform !== "function") {
    context.drawImage(
      source,
      sourceX,
      sourceY,
      sourceWidth,
      sourceHeight,
      destinationX,
      destinationY,
      destinationWidth,
      destinationHeight,
    );
    return;
  }
  const { a, b, c, d, e, f } = transform;
  if (b === 0 && c === 0) {
    const transformedWidth = Math.abs(destinationWidth * a);
    const transformedHeight = Math.abs(destinationHeight * d);
    if (transformedWidth < 1 || transformedHeight < 1) {
      // Preserve subpixel coverage for image-mask hairlines. Snapping either
      // axis to one opaque pixel makes a fractional PDF line much too dark.
      context.drawImage(
        source,
        sourceX,
        sourceY,
        sourceWidth,
        sourceHeight,
        destinationX,
        destinationY,
        destinationWidth,
        destinationHeight,
      );
      return;
    }
    const x = Math.round(destinationX * a + e);
    const y = Math.round(destinationY * d + f);
    const width = Math.abs(Math.round((destinationX + destinationWidth) * a + e) - x) || 1;
    const height = Math.abs(Math.round((destinationY + destinationHeight) * d + f) - y) || 1;
    context.setTransform(Math.sign(a), 0, 0, Math.sign(d), x, y);
    context.drawImage(source, sourceX, sourceY, sourceWidth, sourceHeight, 0, 0, width, height);
    context.setTransform(a, b, c, d, e, f);
    return;
  }
  if (a === 0 && d === 0) {
    const x = Math.round(destinationY * c + e);
    const y = Math.round(destinationX * b + f);
    const width = Math.abs(Math.round((destinationY + destinationHeight) * c + e) - x) || 1;
    const height = Math.abs(Math.round((destinationX + destinationWidth) * b + f) - y) || 1;
    context.setTransform(0, Math.sign(b), Math.sign(c), 0, x, y);
    context.drawImage(source, sourceX, sourceY, sourceWidth, sourceHeight, 0, 0, height, width);
    context.setTransform(a, b, c, d, e, f);
    return;
  }
  context.drawImage(
    source,
    sourceX,
    sourceY,
    sourceWidth,
    sourceHeight,
    destinationX,
    destinationY,
    destinationWidth,
    destinationHeight,
  );
}

function writeGradientColor(
  stops: readonly { readonly offset: number; readonly color: number }[],
  offset: number,
  output: Uint8ClampedArray,
  index: number,
): void {
  const first = stops[0]!;
  const last = stops[stops.length - 1]!;
  let rightIndex = 0;
  while (rightIndex < stops.length && stops[rightIndex]!.offset < offset) rightIndex += 1;
  const right = rightIndex >= stops.length ? last : stops[rightIndex]!;
  const left = rightIndex <= 0 ? first : stops[rightIndex - 1]!;
  const span = right.offset - left.offset;
  const ratio = span <= Number.EPSILON ? 0 : Math.max(0, Math.min(1, (offset - left.offset) / span));
  const inverse = 1 - ratio;
  output[index] = Math.round(((left.color >>> 24) & 0xff) * inverse + ((right.color >>> 24) & 0xff) * ratio);
  output[index + 1] = Math.round(((left.color >>> 16) & 0xff) * inverse + ((right.color >>> 16) & 0xff) * ratio);
  output[index + 2] = Math.round(((left.color >>> 8) & 0xff) * inverse + ((right.color >>> 8) & 0xff) * ratio);
  output[index + 3] = Math.round((left.color & 0xff) * inverse + (right.color & 0xff) * ratio);
}

function gradientRasterSize(width: number, height: number, scaleX: number, scaleY: number): {
  readonly width: number;
  readonly height: number;
  readonly scaleX: number;
  readonly scaleY: number;
} {
  let pixelWidth = Math.max(1, Math.ceil(width * scaleX));
  let pixelHeight = Math.max(1, Math.ceil(height * scaleY));
  const maximumPixels = 8_000_000;
  const pixels = pixelWidth * pixelHeight;
  if (pixels > maximumPixels) {
    const reduction = Math.sqrt(maximumPixels / pixels);
    pixelWidth = Math.max(1, Math.floor(pixelWidth * reduction));
    pixelHeight = Math.max(1, Math.floor(pixelHeight * reduction));
  }
  return {
    width: pixelWidth,
    height: pixelHeight,
    scaleX: pixelWidth / width,
    scaleY: pixelHeight / height,
  };
}

function drawShapeGradientPaint(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  geometry: SceneGeometry,
  fillRule: CanvasFillRule,
  paint: Extract<ScenePaint, { kind: "shape-gradient" }>,
): void {
  const { x, y, width, height } = object.bounds;
  if (width <= 0 || height <= 0 || paint.stops.length === 0) return;
  const transform = canvasTransform(context) ?? IDENTITY_TRANSFORM;
  const raster = gradientRasterSize(width, height,
    Math.hypot(transform.a, transform.b), Math.hypot(transform.c, transform.d));
  const canvas = new OffscreenCanvas(raster.width, raster.height);
  const target = canvas.getContext("2d", { willReadFrequently: true });
  if (target === null) return;
  target.scale(raster.scaleX, raster.scaleY);
  target.translate(-x, -y);
  traceGeometry(target, object, geometry);
  target.clip(fillRule);
  // Store the ramp in an opaque mask first, so translucent stops do not
  // accumulate opacity as the authored contour contracts toward fillToRect.
  for (let step = 255; step >= 0; step -= 1) {
    const contraction = 1 - Math.sqrt(1 - step / 255);
    const inward = 1 - contraction;
    target.save();
    target.translate(x + width * paint.focus.left * inward, y + height * paint.focus.top * inward);
    target.scale(Math.max(1e-6, 1 - (paint.focus.left + paint.focus.right) * inward),
      Math.max(1e-6, 1 - (paint.focus.top + paint.focus.bottom) * inward));
    target.translate(-x, -y);
    traceGeometry(target, object, geometry);
    target.fillStyle = `rgb(${step},${step},${step})`;
    target.fill(fillRule);
    target.restore();
  }
  const pixels = target.getImageData(0, 0, raster.width, raster.height);
  for (let index = 0; index < pixels.data.length; index += 4) {
    const coverage = pixels.data[index + 3]! / 255;
    writeGradientColor(paint.stops, pixels.data[index]! / 255, pixels.data, index);
    pixels.data[index + 3] = Math.round(pixels.data[index + 3]! * coverage);
  }
  target.putImageData(pixels, 0, 0);
  context.drawImage(canvas, x, y, width, height);
}

function drawRectGradientPaint(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  geometry: SceneGeometry,
  fillRule: CanvasFillRule,
  paint: Extract<ScenePaint, { kind: "rect-gradient" }>,
): void {
  const { x, y, width, height } = object.bounds;
  if (width <= 0 || height <= 0 || paint.stops.length === 0) return;
  const transform = typeof context.getTransform === "function"
    ? context.getTransform()
    : { a: 1, b: 0, c: 0, d: 1 };
  const raster = gradientRasterSize(
    width,
    height,
    Math.max(1, Math.hypot(transform.a, transform.b)),
    Math.max(1, Math.hypot(transform.c, transform.d)),
  );
  const { width: pixelWidth, height: pixelHeight, scaleX, scaleY } = raster;
  const canvas = new OffscreenCanvas(pixelWidth, pixelHeight);
  const paintContext = canvas.getContext("2d", { alpha: true });
  if (paintContext === null) return;
  const image = paintContext.createImageData(pixelWidth, pixelHeight);
  const leftSpan = Math.max(Number.EPSILON, paint.center.x);
  const rightSpan = Math.max(Number.EPSILON, width - paint.center.x);
  const topSpan = Math.max(Number.EPSILON, paint.center.y);
  const bottomSpan = Math.max(Number.EPSILON, height - paint.center.y);
  for (let pixelY = 0; pixelY < pixelHeight; pixelY += 1) {
    const localY = (pixelY + 0.5) / scaleY;
    const vertical = localY <= paint.center.y
      ? (paint.center.y - localY) / topSpan
      : (localY - paint.center.y) / bottomSpan;
    for (let pixelX = 0; pixelX < pixelWidth; pixelX += 1) {
      const localX = (pixelX + 0.5) / scaleX;
      const horizontal = localX <= paint.center.x
        ? (paint.center.x - localX) / leftSpan
        : (localX - paint.center.x) / rightSpan;
      const distance = Math.max(0, Math.min(1, Math.max(horizontal, vertical)));
      const offset = distance;
      const index = (pixelY * pixelWidth + pixelX) * 4;
      writeGradientColor(paint.stops, offset, image.data, index);
    }
  }
  paintContext.putImageData(image, 0, 0);
  context.save();
  try {
    traceGeometry(context, object, geometry);
    context.clip(fillRule);
    context.drawImage(canvas, x, y, width, height);
  } finally {
    context.restore();
  }
}

function drawCircleGradientPaint(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  geometry: SceneGeometry,
  fillRule: CanvasFillRule,
  paint: Extract<ScenePaint, { kind: "circle-gradient" }>,
): void {
  const { x, y, width, height } = object.bounds;
  if (width <= 0 || height <= 0 || paint.stops.length === 0) return;
  const transform = typeof context.getTransform === "function"
    ? context.getTransform()
    : { a: 1, b: 0, c: 0, d: 1 };
  const raster = gradientRasterSize(
    width,
    height,
    Math.max(1, Math.hypot(transform.a, transform.b)),
    Math.max(1, Math.hypot(transform.c, transform.d)),
  );
  const { width: pixelWidth, height: pixelHeight, scaleX, scaleY } = raster;
  const canvas = new OffscreenCanvas(pixelWidth, pixelHeight);
  const paintContext = canvas.getContext("2d", { alpha: true });
  if (paintContext === null) return;
  const image = paintContext.createImageData(pixelWidth, pixelHeight);
  const centerDeltaX = paint.end.x - paint.start.x;
  const centerDeltaY = paint.end.y - paint.start.y;
  const radiusDelta = paint.end.radius - paint.start.radius;
  const quadratic = centerDeltaX * centerDeltaX
    + centerDeltaY * centerDeltaY
    - radiusDelta * radiusDelta;
  const focusInsideOuterCircle = Math.hypot(centerDeltaX, centerDeltaY) <= paint.end.radius;
  for (let pixelY = 0; pixelY < pixelHeight; pixelY += 1) {
    const localY = (pixelY + 0.5) / scaleY;
    for (let pixelX = 0; pixelX < pixelWidth; pixelX += 1) {
      const localX = (pixelX + 0.5) / scaleX;
      const pointX = localX - paint.start.x;
      const pointY = localY - paint.start.y;
      const linear = -2 * (
        pointX * centerDeltaX
        + pointY * centerDeltaY
        + paint.start.radius * radiusDelta
      );
      const constant = pointX * pointX + pointY * pointY
        - paint.start.radius * paint.start.radius;
      let scale = 1;
      if (Math.abs(quadratic) <= Number.EPSILON) {
        if (Math.abs(linear) > Number.EPSILON) scale = -constant / linear;
      } else {
        const discriminant = Math.max(0, linear * linear - 4 * quadratic * constant);
        const root = Math.sqrt(discriminant);
        const first = (-linear - root) / (2 * quadratic);
        const second = (-linear + root) / (2 * quadratic);
        if (first >= 0 && second >= 0) scale = Math.min(first, second);
        else if (first >= 0) scale = first;
        else if (second >= 0) scale = second;
      }
      scale = Math.max(0, Math.min(1, scale));
      const offset = focusInsideOuterCircle ? scale : scale * scale;
      const index = (pixelY * pixelWidth + pixelX) * 4;
      writeGradientColor(paint.stops, offset, image.data, index);
    }
  }
  paintContext.putImageData(image, 0, 0);
  context.save();
  try {
    traceGeometry(context, object, geometry);
    context.clip(fillRule);
    context.drawImage(canvas, x, y, width, height);
  } finally {
    context.restore();
  }
}

const DRAWINGML_STIPPLE_THRESHOLDS = [
  0, 48, 12, 60, 3, 51, 15, 63,
  32, 16, 44, 28, 35, 19, 47, 31,
  8, 56, 4, 52, 11, 59, 7, 55,
  40, 24, 36, 20, 43, 27, 39, 23,
  2, 50, 14, 62, 1, 49, 13, 61,
  34, 18, 46, 30, 33, 17, 45, 29,
  10, 58, 6, 54, 9, 57, 5, 53,
  42, 26, 38, 22, 41, 25, 37, 21,
] as const;

// Standard hatch bitmap data, cross-checked against native GDI+ pixel expectations:
// https://github.com/wine-mirror/wine/blob/master/dlls/gdiplus/tests/brush.c (test_hatchBrush).
const DRAWINGML_PATTERN_ROWS: Readonly<Record<string, readonly number[]>> = {
  pct5: [128, 0, 0, 0, 8, 0, 0, 0],
  pct10: [128, 0, 8, 0, 128, 0, 8, 0],
  pct20: [136, 0, 34, 0, 136, 0, 34, 0],
  pct25: [136, 34, 136, 34, 136, 34, 136, 34],
  pct30: [170, 68, 170, 17, 170, 68, 170, 17],
  pct40: [170, 85, 170, 81, 170, 85, 170, 21],
  pct50: [170, 85, 170, 85, 170, 85, 170, 85],
  pct60: [238, 85, 187, 85, 238, 85, 187, 85],
  pct70: [119, 221, 119, 221, 119, 221, 119, 221],
  pct75: [119, 255, 221, 255, 119, 255, 221, 255],
  pct80: [239, 255, 254, 255, 239, 255, 254, 255],
  pct90: [255, 255, 255, 247, 255, 255, 255, 127],
  horz: [255, 0, 0, 0, 0, 0, 0, 0], vert: [128, 128, 128, 128, 128, 128, 128, 128],
  ltHorz: [255, 0, 0, 0, 255, 0, 0, 0], ltVert: [136, 136, 136, 136, 136, 136, 136, 136],
  narHorz: [255, 0, 255, 0, 255, 0, 255, 0], narVert: [85, 85, 85, 85, 85, 85, 85, 85],
  dkHorz: [255, 255, 0, 0, 255, 255, 0, 0], dkVert: [204, 204, 204, 204, 204, 204, 204, 204],
  dnDiag: [1, 2, 4, 8, 16, 32, 64, 128], upDiag: [128, 64, 32, 16, 8, 4, 2, 1],
  ltDnDiag: [136, 68, 34, 17, 136, 68, 34, 17], ltUpDiag: [17, 34, 68, 136, 17, 34, 68, 136],
  wdDnDiag: [193, 224, 112, 56, 28, 14, 7, 131], wdUpDiag: [131, 7, 14, 28, 56, 112, 224, 193],
  dkDnDiag: [204, 102, 51, 153, 204, 102, 51, 153], dkUpDiag: [51, 102, 204, 153, 51, 102, 204, 153],
  dashHorz: [240, 0, 0, 0, 15, 0, 0, 0], dashVert: [128, 128, 128, 128, 8, 8, 8, 8],
  dashDnDiag: [0, 0, 136, 68, 34, 17, 0, 0], dashUpDiag: [0, 0, 17, 34, 68, 136, 0, 0],
  cross: [255, 128, 128, 128, 128, 128, 128, 128],
  diagCross: [129, 66, 36, 24, 24, 36, 66, 129],
  smGrid: [255, 136, 136, 136, 255, 136, 136, 136], lgGrid: [255, 128, 128, 128, 128, 128, 128, 128],
  smCheck: [153, 102, 102, 153, 153, 102, 102, 153], lgCheck: [240, 240, 240, 240, 15, 15, 15, 15],
  dotGrid: [170, 0, 128, 0, 128, 0, 128, 0], dotDmnd: [128, 0, 34, 0, 8, 0, 34, 0],
  smConfetti: [128, 8, 64, 2, 16, 1, 32, 4], lgConfetti: [177, 48, 3, 27, 216, 192, 12, 141],
  horzBrick: [255, 128, 128, 128, 255, 8, 8, 8], diagBrick: [1, 2, 4, 8, 24, 36, 66, 129],
  openDmnd: [130, 68, 40, 16, 40, 68, 130, 1], solidDmnd: [16, 56, 124, 254, 124, 56, 16, 0],
  plaid: [170, 85, 170, 85, 240, 240, 240, 240],
  weave: [136, 84, 34, 69, 136, 20, 34, 81],
  shingle: [3, 132, 72, 48, 12, 2, 1, 1],
  trellis: [255, 102, 255, 153, 255, 102, 255, 153],
  divot: [0, 16, 8, 16, 0, 128, 1, 128],
  zigZag: [129, 66, 36, 24, 129, 66, 36, 24],
  wave: [0, 24, 37, 192, 0, 24, 37, 192],
  sphere: [119, 137, 143, 143, 119, 152, 248, 248],
};

// SpreadsheetML cell patterns retain their own density and hatch geometry.
const SPREADSHEET_PATTERN_ROWS: Readonly<Record<string, readonly number[]>> = {
  mediumGray: [170,85,170,85,170,85,170,85], darkGray: [255,170,255,170,255,170,255,170],
  lightGray: [170,0,170,0,170,0,170,0], gray125: [136,0,34,0,136,0,34,0], gray0625: [128,0,8,0,128,0,8,0],
  darkHorizontal: [255,255,0,0,255,255,0,0], darkVertical: [204,204,204,204,204,204,204,204],
  darkDown: [204,102,51,153,204,102,51,153], darkUp: [153,51,102,204,153,51,102,204],
  darkGrid: [255,136,136,136,255,136,136,136], darkTrellis: [221,119,221,119,221,119,221,119],
  lightHorizontal: [255,0,0,0,255,0,0,0], lightVertical: [136,136,136,136,136,136,136,136],
  lightDown: [136,68,34,17,136,68,34,17], lightUp: [17,34,68,136,17,34,68,136],
  lightGrid: [255,128,128,128,128,128,128,128], lightTrellis: [129,66,36,24,24,36,66,129],
};

// One repeating DrawingML cell shared by shape fills, outlines and text.
function drawingmlPattern(
  context: OffscreenCanvasRenderingContext2D,
  paint: Extract<ScenePaint, { kind: "pattern" }>,
  x: number,
  y: number,
  scale = 1,
): CanvasPattern | undefined {
  const percentage = /^pct(\d+)$/u.exec(paint.preset)?.[1];

  const mask = paint.preset.startsWith("xlsx:")
    ? SPREADSHEET_PATTERN_ROWS[paint.preset.slice(5)] : DRAWINGML_PATTERN_ROWS[paint.preset];
  if (percentage === undefined && mask === undefined) return undefined;
  const tile = new OffscreenCanvas(8, 8);
  const target = tile.getContext("2d");
  if (target === null) return undefined;
  target.fillStyle = color(paint.background);
  target.fillRect(0, 0, 8, 8);
  target.fillStyle = color(paint.foreground);
  const covered = Math.round(Math.min(100, Number(percentage ?? 0)) * 64 / 100);
  for (let index = 0; index < 64; index++) {
    if (mask !== undefined
      ? ((mask[index >> 3] ?? 0) & (128 >> (index % 8))) !== 0
      : DRAWINGML_STIPPLE_THRESHOLDS[index]! < covered) {
      target.fillRect(index % 8, index >> 3, 1, 1);
    }
  }
  const pattern = context.createPattern(tile, "repeat");
  pattern?.setTransform({ a: 1 / scale, d: 1 / scale, e: x, f: y });
  return pattern ?? undefined;
}

function drawPatternPaint(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  geometry: SceneGeometry,
  fillRule: CanvasFillRule,
  paint: Extract<ScenePaint, { kind: "pattern" }>,
  groupedScale: number,
): void {
  if (!paint.preset.startsWith("pdf:")) {
    const pattern = drawingmlPattern(context, paint, object.bounds.x, object.bounds.y, groupedScale);
    if (pattern !== undefined) {
      context.fillStyle = pattern;
      context.fill(fillRule);
    }
    return;
  }
  // PDF hatches keep their authored local spacing, phase and subpixel width.
  const pdf = /^pdf:(up|down|cross|vert|horz|dot|checker):([0-9.]+):([0-9.]+)(?::(-?[0-9.]+):(-?[0-9.]+))?$/u.exec(paint.preset);
  if (pdf === null) return;
  if (visible(paint.background)) {
    context.fillStyle = color(paint.background);
    context.fill(fillRule);
  }
  if (!visible(paint.foreground)) return;
  context.save();
  try {
    traceGeometry(context, object, geometry);
    context.clip(fillRule);
    const { x, y, width, height } = object.bounds;
    const spacing = Math.max(1, Math.min(64, Number(pdf[2])));
    const latticeStart = (minimum: number, origin: string | undefined): number =>
      origin === undefined ? minimum : minimum + ((Number(origin) - minimum) % spacing + spacing) % spacing;
    context.strokeStyle = context.fillStyle = color(paint.foreground);
    context.lineWidth = effectiveStrokeWidth(context, object, Math.max(0.1, Math.min(8, Number(pdf[3]))));
    context.setLineDash?.([]);
    if (pdf[1] === "checker" || pdf[1] === "dot") {
      const checker = pdf[1] === "checker";
      const startX = checker ? latticeStart(x - spacing, pdf[4]) : x;
      const startY = checker ? latticeStart(y - spacing, pdf[5]) : y;
      for (let row = 0, py = startY; py <= y + height + spacing; row++, py += spacing) {
        for (let column = 0, px = startX; px <= x + width + spacing; column++, px += spacing) {
          if (checker) {
            if ((row + column) % 2 === 0) context.fillRect(px, py, spacing, spacing);
          } else {
            context.beginPath();
            const radius = Math.max(0.5, spacing / 6);
            context.ellipse(px, py, radius, radius, 0, 0, Math.PI * 2);
            context.fill();
          }
        }
      }
    } else {
      context.beginPath();
      if (pdf[1] === "horz") {
        for (let py = latticeStart(y, pdf[5]); py <= y + height + spacing; py += spacing) {
          context.moveTo(x, py); context.lineTo(x + width, py);
        }
      } else if (pdf[1] === "vert") {
        for (let px = latticeStart(x, pdf[4]); px <= x + width + spacing; px += spacing) {
          context.moveTo(px, y); context.lineTo(px, y + height);
        }
      } else {
        for (let offset = -height; offset <= width + height; offset += spacing) {
          if (pdf[1] !== "down") {
            context.moveTo(x + offset, y + height); context.lineTo(x + offset + height, y);
          }
          if (pdf[1] !== "up") {
            context.moveTo(x + offset, y); context.lineTo(x + offset + height, y + height);
          }
        }
      }
      context.stroke();
    }
  } finally { context.restore(); }
}

function applyStrokeStyle(
  context: OffscreenCanvasRenderingContext2D,
  style: SceneStrokeStyle | undefined,
): void {
  if (style === undefined) return;
  context.lineCap = style.cap === "flat" ? "butt" : style.cap;
  context.lineJoin = style.join;
  context.miterLimit = style.miterLimit;
  context.lineDashOffset = style.dashOffset ?? 0;
  if (typeof context.setLineDash === "function") context.setLineDash([...style.dash]);
}

function strokePaintedGeometry(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  geometry: SceneGeometry,
  fillRule: CanvasFillRule,
  stroke: string | CanvasGradient | CanvasPattern,
  width: number,
  style: SceneStrokeStyle | undefined,
  gap: string | CanvasGradient | CanvasPattern | undefined,
  groupedScale = 1,
): void {
  const strokeOnce = (
    lineWidth: number,
    strokeStyle: string | CanvasGradient | CanvasPattern,
  ) => {
    context.strokeStyle = strokeStyle;
    context.lineWidth = lineWidth;
    if (object.source.format === "pdf" || object.source.format === "xps") {
      traceGeometry(context, object, geometry);
    } else {
      retraceStrokeGeometry(context, object, geometry, groupedScale, (style?.compound ?? "single") === "single");
    }
    context.stroke();
  };
  context.save();
  try {
    applyStrokeStyle(context, style);
    if (style?.alignment === "inset" && geometry !== "line") {
      traceGeometry(context, object, geometry);
      context.clip(fillRule);
    }
    const compound = style?.compound ?? "single";
    if (compound === "single") {
      strokeOnce(width, stroke);
    } else if (compound === "double") {
      strokeOnce(width, stroke);
      strokeOnce(width / 3, gap ?? "rgba(255, 255, 255, 1)");
    } else if (compound === "triple") {
      strokeOnce(width * 1.8, stroke);
      strokeOnce(width * 1.2, gap ?? "rgba(255, 255, 255, 1)");
      strokeOnce(width * 0.8, stroke);
      strokeOnce(width * 0.4, gap ?? "rgba(255, 255, 255, 1)");
      strokeOnce(width * 0.2, stroke);
    } else {
      const outerScale = compound === "thick-thin" ? 1.8 : 1.5;
      const innerScale = compound === "thick-thin" ? 0.25 : 0.55;
      strokeOnce(width * outerScale, stroke);
      strokeOnce(width * 0.85, gap ?? "rgba(255, 255, 255, 1)");
      strokeOnce(width * innerScale, stroke);
    }
  } finally {
    context.restore();
  }
}

function drawPaintedGeometry(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  visual: Extract<SceneVisual, { kind: "painted-shape" | "rich-text" }>,
  fillImage?: ImageBitmap,
  strokeStyle?: SceneStrokeStyle,
  fillVisualPattern?: CanvasPattern,
  strokeVisualPattern?: CanvasPattern,
  groupedScale = 1,
): void {
  if (typeof visual.geometry === "object" && visual.geometry.kind === "layered-path") {
    // Paint the entire composite surface before any outlines.
    for (const strokePass of [false, true]) {
      for (const layer of visual.geometry.layers) {
        const geometry: SceneGeometry = {
          kind: "path",
          fillRule: layer.fillRule,
          commands: layer.commands,
        };
        if (!strokePass && layer.fill !== "none") {
          drawPaintedGeometry(
            context,
            object,
            {
              ...visual,
              geometry,
              fill: pathFillPaint(visual.fill, layer.fill),
              stroke: { kind: "none" },
            },
            fillImage,
            strokeStyle,
            fillVisualPattern,
            strokeVisualPattern,
            groupedScale,
          );
        }
        if (strokePass && layer.stroke) {
          drawPaintedGeometry(
            context,
            object,
            { ...visual, geometry, fill: { kind: "none" } },
            fillImage,
            strokeStyle,
            fillVisualPattern,
            strokeVisualPattern,
            groupedScale,
          );
        }
      }
    }
    return;
  }
  const fillRule = traceGeometry(context, object, visual.geometry);
  const strokeWidth = effectiveStrokeWidth(context, object, visual.strokeWidth / groupedScale);
  if (visual.fill.kind === "image" && fillImage !== undefined && visual.geometry !== "line") {
    drawImagePaint(context, object, visual.geometry, fillRule, visual.fill, fillImage, groupedScale);
  } else if ((visual.fill.kind === "visual" || visual.fill.kind === "xps-gradient")
    && fillVisualPattern !== undefined && visual.geometry !== "line") {
    context.fillStyle = fillVisualPattern;
    context.fill(fillRule);
  } else if (visual.fill.kind === "pattern" && visual.geometry !== "line") {
    drawPatternPaint(context, object, visual.geometry, fillRule, visual.fill, groupedScale);
    if (visual.stroke.kind !== "none" && strokeWidth > 0) {
      traceGeometry(context, object, visual.geometry);
    }
  } else if (visual.fill.kind === "shape-gradient" && visual.geometry !== "line") {
    drawShapeGradientPaint(context, object, visual.geometry, fillRule, visual.fill);
  } else if (visual.fill.kind === "rect-gradient" && visual.geometry !== "line") {
    drawRectGradientPaint(context, object, visual.geometry, fillRule, visual.fill);
  } else if (visual.fill.kind === "circle-gradient" && visual.geometry !== "line") {
    drawCircleGradientPaint(context, object, visual.geometry, fillRule, visual.fill);
  } else {
    const fill = paintStyle(context, object, visual.fill, 0, 0, visual.geometry);
    if (fill !== undefined && visual.geometry !== "line") {
      if (object.source.format === "pdf" && visual.fill.kind === "radial-gradient") {
        const reversed = reversedConicRadialPaintStyle(context, object, visual.fill);
        if (reversed !== undefined) {
          context.fillStyle = reversed;
          context.fill(fillRule);
        }
      }
      context.fillStyle = fill;
      context.fill(fillRule);
    }
  }
  const stroke = visual.stroke.kind === "pattern"
    ? (visual.stroke.preset.startsWith("pdf:") ? color(visual.stroke.foreground)
      : drawingmlPattern(context, visual.stroke, object.bounds.x, object.bounds.y, groupedScale))
    : visual.stroke.kind === "image"
      ? undefined
      : visual.stroke.kind === "visual" || visual.stroke.kind === "xps-gradient"
        ? strokeVisualPattern
    : paintStyle(context, object, visual.stroke);
  if (visual.geometry !== "line" && visual.fill.kind !== "none") {
    // The fill already cast the outer shadow. Do not let the later stroke cast
    // another copy across the filled shape's near edges.
    context.shadowColor = "transparent";
    context.shadowBlur = 0;
    context.shadowOffsetX = 0;
    context.shadowOffsetY = 0;
  }
  if (stroke !== undefined && strokeWidth > 0) {
    const gap = visual.fill.kind === "pattern"
      ? color(visual.fill.background)
      : visual.fill.kind === "image"
        ? undefined
        : visual.fill.kind === "visual" || visual.fill.kind === "xps-gradient"
          ? fillVisualPattern
        : paintStyle(context, object, visual.fill);
    strokePaintedGeometry(
      context,
      object,
      visual.geometry,
      fillRule,
      stroke,
      strokeWidth,
      strokeStyle,
      gap,
      groupedScale,
    );
    if (object.source.format !== "pdf" && object.source.format !== "xps") {
      fillStrokeColoredLineEnds(context, object, visual.geometry, stroke, groupedScale);
    }
  }
}

export interface TextLayoutOptions {
  readonly fillCharacter?: SceneTextLayout["fillCharacter"];
  /** Rectangular exclusions relative to the text origin. */
  readonly wrapRegions?: readonly Rect[];
  readonly fixedLineHeight?: boolean;
  readonly compressPunctuation?: boolean;
  /** The final paragraph continues in another frame, so its last line may justify. */
  readonly continuesAfter?: boolean;
  readonly maxWidth: number;
  readonly maxHeight?: number;
  readonly lineHeight?: number;
  readonly align?: SceneTextAlign;
  readonly direction?: "ltr" | "rtl";
  readonly orientation?: "horizontal" | "vertical-rl" | "vertical-lr" | "rotated-90" | "rotated-270" | "stacked-rl" | "stacked-lr";
  readonly verticalAlign?: "top" | "center" | "bottom";
  readonly autoFit?: "none" | "shrink" | "fit-frame";
  /** Smallest font scale accepted by shrink-to-fit. */
  readonly minScale?: number;
  readonly fontScale?: number;
  readonly lineSpacingReduction?: number;
  readonly columns?: number;
  readonly columnSpacing?: number;
  readonly wrap?: boolean;
  /** Literal bullet or numbering text rendered before the first run. */
  readonly prefix?: string;
  /** Absolute positions from the text box's leading edge. Numeric entries are start tabs. */
  readonly tabStops?: readonly (number | SceneTextTabStop)[];
  readonly defaultTabStop?: number;
  /** Leading position for lines after the first. Defaults to the prefix width. */
  readonly hangingIndent?: number;
  /** Additional vertical distance after an explicit paragraph break. */
  readonly paragraphSpacing?: number;
  readonly firstLineIndent?: number;
  readonly paragraphs?: readonly SceneTextParagraphLayout[];
  /** Internal document-scoped exact-face resolution used for measurement and paint. */
  readonly resolveFontFamily?: (family: string, run: SceneTextRun) => string;
}

export interface TextLayoutRun {
  /** Independent glyph baseline for a drop cap, relative to its ink-box top. */
  readonly baselineOffset?: number;
  /** Line-box top before paragraph leading shifts the glyph origin. */
  readonly lineTop?: number;
  /** Whitespace removed from this punctuation glyph's advance, not its ink. */
  readonly compression?: number;
  readonly text: string;
  /** UTF-16 range in the concatenated logical input runs. */
  readonly start?: number;
  readonly end?: number;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly line: number;
  readonly style: SceneTextRun;
  /** Paragraph-specific line box height. */
  readonly height?: number;
  /** Additional layout-owned spacing between graphemes, in rendered pixels. */
  readonly tracking?: number;
  /** Horizontal-script text painted sideways within a vertical line. */
  readonly sideways?: boolean;
  /** Leader painted across a tab's layout-owned advance. */
  readonly leader?: SceneTextTabStop["leader"];
}

export interface TextLayout {
  readonly runs: readonly TextLayoutRun[];
  readonly rules: readonly TextLayoutRule[];
  readonly width: number;
  readonly height: number;
  readonly lineCount: number;
  readonly lineHeight: number;
  readonly scale: number;
}

interface TextLayoutRule {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly color: number;
  readonly strokeWidth: number;
}

type TextMeasureContext = Pick<OffscreenCanvasRenderingContext2D, "measureText"> & {
  textBaseline?: CanvasTextBaseline;
  font: string;
  letterSpacing: string;
};

function applyRunFont(
  context: TextMeasureContext,
  run: SceneTextRun,
  scale = 1,
  resolveFontFamily: (family: string, run: SceneTextRun) => string = (family) => family,
): void {
  const font = fontShorthand(
    resolveFontFamily(run.fontFamily, run),
    run.fontSize * scale,
    run.italic,
    run.bold,
  );
  if (context.font !== font) context.font = font;
  const letterSpacing = `${run.letterSpacing * scale}px`;
  if (context.letterSpacing !== letterSpacing) context.letterSpacing = letterSpacing;
}

function runHorizontalScale(run: SceneTextRun): number {
  return Number.isFinite(run.horizontalScale) && (run.horizontalScale ?? 0) > 0
    ? run.horizontalScale ?? 1
    : 1;
}

function measureRunText(context: TextMeasureContext, text: string, run: SceneTextRun): number {
  return context.measureText(text).width * runHorizontalScale(run);
}

function guardedFallbackSpaceLetterSpacing(
  context: TextMeasureContext,
  run: SceneTextRun,
  renderedFamily: string,
  advanceEm: number,
): number {
  if (!/^[\t ]+$/u.test(run.text)) return run.letterSpacing;
  const previousFont = context.font;
  const previousLetterSpacing = context.letterSpacing;
  try {
    context.font = fontShorthand(renderedFamily, run.fontSize, run.italic, run.bold);
    context.letterSpacing = "0px";
    const correction = run.fontSize * advanceEm - context.measureText(" ").width;
    return Number.isFinite(correction) ? run.letterSpacing + correction : run.letterSpacing;
  } finally {
    context.font = previousFont;
    context.letterSpacing = previousLetterSpacing;
  }
}

function fitSubstitutedSongGlyph(
  context: TextMeasureContext,
  run: SceneTextRun,
  resolution: FontResolution,
): SceneTextRun {
  // Positioned half-width SimSun glyphs must not overrun their authored cells.
  // Exact/embedded fonts and proportional multi-character runs keep their metrics.
  if (!/^(?:宋体|新宋体|simsun|nsimsun)$/iu.test(run.fontFamily.trim())
    || !/^[!-~]$/u.test(run.text)
    || !(resolution.source === "fallback" || resolution.spaceAdvanceEm === .5)) return run;
  const previousFont = context.font;
  const previousLetterSpacing = context.letterSpacing;
  try {
    context.font = fontShorthand(resolution.family, run.fontSize, run.italic, run.bold);
    context.letterSpacing = "0px";
    const width = context.measureText(run.text).width;
    if (!Number.isFinite(width) || width <= 0) return run;
    const advance = run.fontSize / 2;
    return {
      ...run,
      horizontalScale: runHorizontalScale(run) * Math.min(1, advance / width),
      letterSpacing: run.letterSpacing + Math.max(0, advance - width),
    };
  } finally {
    context.font = previousFont;
    context.letterSpacing = previousLetterSpacing;
  }
}

const CJK_SCRIPT = /\p{Script=Han}|\p{Script=Hiragana}|\p{Script=Katakana}|\p{Script=Hangul}/u;
const EAST_ASIAN_AUTO_SPACING_EM = 0.25;
const ARABIC_SCRIPT = /\p{Script=Arabic}/u;
const THAI_SCRIPT = /\p{Script=Thai}/u;
const COMPLEX_SHAPING_SCRIPT = /\p{Script=Arabic}|\p{Script=Thai}/u;
const DISTRIBUTED_JUSTIFY_SCRIPT =
  /\p{Script=Han}|\p{Script=Hiragana}|\p{Script=Katakana}|\p{Script=Hangul}|\p{Script=Thai}|\p{Script=Arabic}/u;
const OPENING_PUNCTUATION = /^(?:\p{Ps}|\p{Pi})\p{M}*$/u;
const CLOSING_PUNCTUATION = /^(?:\p{Pe}|\p{Pf})\p{M}*$/u;
const EAST_ASIAN_PUNCTUATION = /[\u3000-\u303f\ufe10-\ufe1f\ufe30-\ufe4f\uff01-\uff65]/u;
const TRAILING_PUNCTUATION = /^[\p{Pe}\p{Pf}\p{Po}]\p{M}*$/u;
const GRAPHEME_SEGMENTER = new Intl.Segmenter(undefined, { granularity: "grapheme" });

function punctuationCompression(text: string, advance: number): number {
  return /^[、。，．！？：；）］】》〉」』〕”’（［【《〈「『〔“‘]$/u.test(text)
    ? Math.max(0, advance) / 2 : 0;
}

function isJustifiedTextAlign(align: SceneTextAlign): boolean {
  return align === "justify"
    || align === "distribute"
    || align === "medium-kashida"
    || align === "high-kashida"
    || align === "low-kashida"
    || align === "thai-distribute";
}

function isTrackingCharacter(align: SceneTextAlign, text: string): boolean {
  if (align === "distribute") return true;
  if (align === "medium-kashida" || align === "high-kashida" || align === "low-kashida") {
    return ARABIC_SCRIPT.test(text);
  }
  if (align === "thai-distribute") return THAI_SCRIPT.test(text);
  return align === "justify" && DISTRIBUTED_JUSTIFY_SCRIPT.test(text);
}

function isTrackingBoundary(align: SceneTextAlign, left: string, right: string): boolean {
  return isTrackingCharacter(align, left) && isTrackingCharacter(align, right);
}

function trackingPieces(text: string, align: SceneTextAlign): readonly string[] {
  if (!isJustifiedTextAlign(align) || align === "distribute") return [text];
  const textGraphemes = graphemes(text);
  if (textGraphemes.length < 2) return [text];
  const eligibility = textGraphemes.map((grapheme) => isTrackingCharacter(align, grapheme));
  if (eligibility.every((value) => value === eligibility[0])) return [text];

  const result: string[] = [];
  let start = 0;
  for (let index = 1; index <= textGraphemes.length; index += 1) {
    if (index < textGraphemes.length && eligibility[index] === eligibility[start]) continue;
    result.push(textGraphemes.slice(start, index).join(""));
    start = index;
  }
  return result;
}

function sameTextRunStyle(
  left: SceneTextRun,
  right: SceneTextRun,
  includeLetterSpacing = true,
): boolean {
  return left.fontFamily === right.fontFamily
    && left.fontSize === right.fontSize
    && left.color === right.color
    && left.paint === right.paint
    && left.textEffect?.stroke === right.textEffect?.stroke
    && left.textEffect?.strokeWidth === right.textEffect?.strokeWidth
    && left.textEffect?.glow === right.textEffect?.glow
    && left.textEffect?.reflection === right.textEffect?.reflection
    && left.textEffect?.fillToText === right.textEffect?.fillToText
    && left.bold === right.bold
    && left.italic === right.italic
    && left.underline === right.underline
    && left.strikethrough === right.strikethrough
    && left.highlight === right.highlight
    && left.baselineShift === right.baselineShift
    && sameShadow(left.shadow, right.shadow)
    && sameShadow(left.innerShadow, right.innerShadow)
    && left.textEffect?.shadowScaleX === right.textEffect?.shadowScaleX
    && left.textEffect?.shadowScaleY === right.textEffect?.shadowScaleY
    && left.textEffect?.shadowSkewX === right.textEffect?.shadowSkewX
    && left.textEffect?.shadowSkewY === right.textEffect?.shadowSkewY
    && left.textEffect?.shadowAlignment === right.textEffect?.shadowAlignment
    && (left.eastAsianLineBreaks !== false) === (right.eastAsianLineBreaks !== false)
    && runHorizontalScale(left) === runHorizontalScale(right)
    && (!includeLetterSpacing || left.letterSpacing === right.letterSpacing);
}

function sameShadow(left: SceneTextRun["shadow"], right: SceneTextRun["shadow"]): boolean {
  return left === right || (left !== undefined && right !== undefined
    && left.color === right.color
    && left.blur === right.blur
    && left.offsetX === right.offsetX
    && left.offsetY === right.offsetY);
}

function graphemes(text: string): readonly string[] {
  return Array.from(
    GRAPHEME_SEGMENTER.segment(text),
    ({ segment }) => segment,
  );
}

function isEastAsianTrailingPunctuation(text: string): boolean {
  return (EAST_ASIAN_PUNCTUATION.test(text) && TRAILING_PUNCTUATION.test(text))
    || /^[”’]$/u.test(text);
}

function isLineStartProhibited(text: string): boolean {
  const segments = graphemes(text);
  return segments.length > 0 && segments.every(segment =>
    isEastAsianTrailingPunctuation(segment) || /^[,.;:!?)\]}…]$/u.test(segment));
}

function compoundTextTokens(text: string, eastAsianLineBreaks: boolean): readonly string[] {
  const textGraphemes = graphemes(text);
  let closingStart = textGraphemes.length;
  while (eastAsianLineBreaks && closingStart > 0 && CLOSING_PUNCTUATION.test(textGraphemes[closingStart - 1] ?? "")) {
    closingStart -= 1;
  }
  const result: string[] = [];
  let compound = "";
  const flushCompound = (): void => {
    if (compound !== "") result.push(compound);
    compound = "";
  };
  for (const segment of textGraphemes.slice(0, closingStart)) {
    if (eastAsianLineBreaks && EAST_ASIAN_PUNCTUATION.test(segment) && OPENING_PUNCTUATION.test(segment)) {
      flushCompound();
      compound = segment;
    } else if (CJK_SCRIPT.test(segment)) {
      if (compound !== "") {
        if (graphemes(compound).every((value) => (
          EAST_ASIAN_PUNCTUATION.test(value) && OPENING_PUNCTUATION.test(value)
        ))) {
          result.push(compound + segment);
          compound = "";
        } else {
          flushCompound();
          result.push(segment);
        }
      } else {
        result.push(segment);
      }
    } else if (EAST_ASIAN_PUNCTUATION.test(segment)
      || (compound === "" && isLineStartProhibited(segment))) {
      flushCompound();
      result.push(segment);
    } else {
      compound += segment;
    }
  }
  flushCompound();
  const closing = textGraphemes.slice(closingStart).join("");
  if (closing !== "") result.push(closing);
  return result;
}

function latinLineBreakTokens(text: string): readonly string[] {
  const result: string[] = [];
  let token = "";
  for (const segment of graphemes(text)) {
    token += segment;
    if (segment !== "-" && segment !== "\u2010") continue;
    result.push(token);
    token = "";
  }
  if (token !== "") result.push(token);
  return result.length === 0 ? [text] : result;
}

function textTokens(text: string, eastAsianLineBreaks: boolean): readonly string[] {
  const tokens = text.match(/\r\n|[\n\u2028]|[^\S\r\n\u2028]+|[^\s]+/gu) ?? [];
  return tokens.flatMap((token) => {
    if (token.includes("\t")) return token.split(/(\t)/u).filter((part) => part !== "");
    if (token === "\n" || token === "\r\n" || token === "\u2028" || /^\s+$/u.test(token)) return [token];
    return compoundTextTokens(token, eastAsianLineBreaks);
  });
}

interface IndexedTextPart {
  readonly text: string;
  readonly start?: number;
  readonly end?: number;
}

function indexedTextParts(
  parts: readonly string[],
  start: number,
  sourceEnd?: number,
): readonly IndexedTextPart[] {
  let cursor = start;
  return parts.map((text) => {
    const part = { text, start: cursor, end: sourceEnd === start ? cursor : cursor + text.length };
    cursor = part.end;
    return part;
  });
}

function isSoftWrapSpace(token: string): boolean {
  return token.length > 0 && Array.from(token).every((character) => character === " ");
}

function wrappingWidth(
  context: TextMeasureContext,
  token: string,
  width: number,
  hangingPunctuation: boolean,
  run: SceneTextRun,
): number {
  if (!hangingPunctuation) return width;
  const tokenGraphemes = graphemes(token);
  let punctuationStart = tokenGraphemes.length;
  while (punctuationStart > 0
    && isEastAsianTrailingPunctuation(tokenGraphemes[punctuationStart - 1] ?? "")) {
    punctuationStart -= 1;
  }
  const punctuation = tokenGraphemes.slice(punctuationStart).join("");
  return punctuation === ""
    ? width
    : Math.max(0, width - measureRunText(context, punctuation, run));
}

interface ResolvedTextParagraph {
  readonly align: SceneTextAlign;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly firstLineIndent: number;
  readonly defaultTabStop: number;
  readonly continuationX: number;
  readonly lineHeight: number;
  readonly spaceBefore: number;
  readonly spaceAfter: number;
  readonly latinLineBreak: boolean;
  readonly hangingPunctuation: boolean;
  readonly authored: boolean;
  readonly index: number;
  readonly ruleAbove?: SceneTextParagraphLayout["ruleAbove"];
  readonly ruleBelow?: SceneTextParagraphLayout["ruleBelow"];
  readonly dropCap?: SceneTextParagraphLayout["dropCap"];
}

interface HorizontalTextLayoutState {
  readonly fixedLineHeight: boolean;
  readonly placements: readonly TextLayoutRun[];
  readonly lineWidths: readonly number[];
  readonly lineStarts: readonly number[];
  readonly lineOffsets: readonly number[];
  readonly lineHeights: readonly number[];
  readonly lineColumns: readonly number[];
  readonly lineJustifiable: readonly boolean[];
  readonly lineHasTab: readonly boolean[];
  readonly lineParagraphs: readonly ResolvedTextParagraph[];
  readonly fallbackParagraph: ResolvedTextParagraph;
  readonly lineCount: number;
  readonly columns: number;
  readonly columnSpacing: number;
  readonly lineHeight: number;
  readonly scale: number;
  readonly maxWidth: number;
  readonly maxHeight: number;
  readonly align: SceneTextAlign;
  readonly direction: "ltr" | "rtl";
  readonly verticalAlign: TextLayoutOptions["verticalAlign"];
  readonly hasTrailingEmptyParagraph: boolean;
}

function verticalTextOffset(
  align: TextLayoutOptions["verticalAlign"],
  remaining: number,
): number {
  if (align === "center") return remaining / 2;
  if (align === "bottom") return remaining;
  return 0;
}

function horizontalTextOffset(
  align: SceneTextAlign,
  direction: "ltr" | "rtl",
  remaining: number,
): number {
  if (align === "center") return remaining / 2;
  if (((align === "start" || isJustifiedTextAlign(align)) && direction === "rtl")
    || (align === "end" && direction === "ltr")) {
    return remaining;
  }
  return 0;
}

function finalizeHorizontalTextLayout(state: HorizontalTextLayoutState): TextLayout {
  const {
    placements,
    lineWidths,
    lineStarts,
    lineOffsets,
    lineHeights,
    lineColumns,
    lineJustifiable,
    lineHasTab,
    lineParagraphs,
    fallbackParagraph,
    lineCount,
    columns,
    columnSpacing,
    lineHeight,
    scale,
    maxWidth,
    maxHeight,
    align,
    direction,
    verticalAlign,
    hasTrailingEmptyParagraph,
  } = state;
  const columnHeights = Array.from({ length: columns }, () => 0);
  for (let lineIndex = 0; lineIndex < lineCount; lineIndex += 1) {
    const column = lineColumns[lineIndex] ?? 0;
    columnHeights[column] = Math.max(
      columnHeights[column] ?? 0,
      (lineOffsets[lineIndex] ?? 0) + (lineHeights[lineIndex] ?? lineHeight),
    );
  }
  const firstContent = Array.from({ length: lineCount }, () => -1);
  const lastContent = Array.from({ length: lineCount }, () => -1);
  placements.forEach((placement, index) => {
    if (placement.baselineOffset !== undefined) return;
    if (placement.text.trim() === "") return;
    if ((firstContent[placement.line] ?? -1) < 0) firstContent[placement.line] = index;
    lastContent[placement.line] = index;
  });
  const collapsesTrailingWhitespace = (lineIndex: number): boolean => (
    lineJustifiable[lineIndex] === true
    || (lineParagraphs[lineIndex]?.align ?? align) === "distribute"
  );
  const visibleLineWidths = lineWidths.map((width, lineIndex) => {
    if (!collapsesTrailingWhitespace(lineIndex)) return width;
    const lastPlacement = placements[lastContent[lineIndex] ?? -1];
    return lastPlacement === undefined
      ? lineStarts[lineIndex] ?? 0
      : lastPlacement.x + lastPlacement.width;
  });
  const naturalLineHeights = Array.from({ length: lineCount }, () => 0);
  for (const placement of placements) {
    if (placement.baselineOffset !== undefined) continue;
    naturalLineHeights[placement.line] = Math.max(
      naturalLineHeights[placement.line] ?? 0,
      placement.style.fontSize * 1.2 * scale,
    );
  }
  const placementShift = Array.from({ length: placements.length }, () => 0);
  const placementExtraWidth = Array.from({ length: placements.length }, () => 0);
  const placementTracking = Array.from({ length: placements.length }, () => 0);
  const stretchedLines = Array.from({ length: lineCount }, () => false);
  for (let lineIndex = 0; lineIndex < lineCount; lineIndex += 1) {
    const lineParagraph = lineParagraphs[lineIndex] ?? fallbackParagraph;
    const paragraphAlign = lineParagraph.align ?? align;
    const first = firstContent[lineIndex] ?? -1;
    const last = lastContent[lineIndex] ?? -1;
    const lineWidth = visibleLineWidths[lineIndex] ?? 0;
    const remaining = maxWidth - lineParagraph.marginRight - lineWidth;
    if (!isJustifiedTextAlign(paragraphAlign)
      || (lineJustifiable[lineIndex] !== true && paragraphAlign !== "distribute")
      || lineHasTab[lineIndex] === true
      || first < 0
      || last < first
      || remaining <= 0) continue;

    const whitespaceGaps: number[] = [];
    if (paragraphAlign === "justify") {
      for (let index = first + 1; index < last; index += 1) {
        if (placements[index]?.text.trim() === "") whitespaceGaps.push(index);
      }
    }
    if (whitespaceGaps.length > 0) {
      const extra = remaining / whitespaceGaps.length;
      let preceding = 0;
      let gapCursor = 0;
      for (let index = first; index <= last; index += 1) {
        while ((whitespaceGaps[gapCursor] ?? Number.POSITIVE_INFINITY) < index) {
          preceding += 1;
          gapCursor += 1;
        }
        placementShift[index] = preceding * extra;
        if (whitespaceGaps[gapCursor] === index) placementExtraWidth[index] = extra;
      }
      stretchedLines[lineIndex] = true;
      continue;
    }

    const lineGraphemes: Array<{ placementIndex: number; text: string }> = [];
    for (let placementIndex = first; placementIndex <= last; placementIndex += 1) {
      const placement = placements[placementIndex];
      if (placement?.line !== lineIndex) continue;
      for (const text of graphemes(placement.text)) {
        lineGraphemes.push({ placementIndex, text });
      }
    }
    const eligibleBoundaries: number[] = [];
    for (let index = 0; index + 1 < lineGraphemes.length; index += 1) {
      const left = lineGraphemes[index];
      const right = lineGraphemes[index + 1];
      if (left !== undefined
        && right !== undefined
        && isTrackingBoundary(paragraphAlign, left.text, right.text)) {
        eligibleBoundaries.push(index);
      }
    }
    if (eligibleBoundaries.length === 0) continue;
    const extra = remaining / eligibleBoundaries.length;
    let preceding = 0;
    let boundaryCursor = 0;
    for (const [graphemeIndex, grapheme] of lineGraphemes.entries()) {
      const placementIndex = grapheme.placementIndex;
      if (graphemeIndex === 0
        || lineGraphemes[graphemeIndex - 1]?.placementIndex !== placementIndex) {
        placementShift[placementIndex] = preceding * extra;
      }
      if (eligibleBoundaries[boundaryCursor] === graphemeIndex) {
        const rightPlacementIndex = lineGraphemes[graphemeIndex + 1]?.placementIndex;
        placementExtraWidth[placementIndex] =
          (placementExtraWidth[placementIndex] ?? 0) + extra;
        placementTracking[placementIndex] = extra;
        if (rightPlacementIndex !== undefined) {
          placementTracking[rightPlacementIndex] = extra;
        }
        preceding += 1;
        boundaryCursor += 1;
      }
    }
    stretchedLines[lineIndex] = true;
  }
  const alignedPlacements = placements.map((placement, placementIndex) => {
    const lineWidth = visibleLineWidths[placement.line] ?? 0;
    const lineParagraph = lineParagraphs[placement.line] ?? fallbackParagraph;
    const remaining = maxWidth - lineParagraph.marginRight - lineWidth;
    const paragraphAlign = lineParagraph.align ?? align;
    const stretched = stretchedLines[placement.line] === true;
    const effectiveLineWidth = stretched ? lineWidth + remaining : lineWidth;
    const collapsedTrailingWhitespace = collapsesTrailingWhitespace(placement.line)
      && ((lastContent[placement.line] ?? -1) < 0
        || placementIndex > (lastContent[placement.line] ?? -1))
      && placement.text.trim() === "";
    const placementWidth = collapsedTrailingWhitespace
      ? 0
      : placement.width + (placementExtraWidth[placementIndex] ?? 0);
    const logicalPlacementX = collapsedTrailingWhitespace
      ? effectiveLineWidth
      : placement.x + (placementShift[placementIndex] ?? 0);
    const alignmentRemaining = stretched ? 0 : remaining;
    const lineOffset = horizontalTextOffset(paragraphAlign, direction, alignmentRemaining);
    const logicalX = direction === "rtl"
      ? effectiveLineWidth - logicalPlacementX - placementWidth
      : logicalPlacementX;
    const column = lineColumns[placement.line] ?? 0;
    const verticalRemaining = Number.isFinite(maxHeight)
      ? maxHeight - (columnHeights[column] ?? 0)
      : 0;
    const anchoredVerticalRemaining = hasTrailingEmptyParagraph
      ? Math.max(0, verticalRemaining)
      : verticalRemaining;
    const verticalOffset = verticalTextOffset(verticalAlign, anchoredVerticalRemaining);
    // DrawingML places the extra leading from authored multiple line spacing
    // before the glyph box. All runs on a line share the largest natural box,
    // so smaller runs keep their baseline instead of being lowered separately.
    const leading = placement.baselineOffset === undefined && !state.fixedLineHeight && lineParagraphs[placement.line]?.authored === true
      ? Math.max(
          0,
          (lineHeights[placement.line] ?? lineHeight)
            - (naturalLineHeights[placement.line] ?? lineHeight),
        )
      : 0;
    return {
      ...placement,
      width: placementWidth,
      x: column * (maxWidth + columnSpacing) + lineOffset + logicalX,
      lineTop: placement.y + verticalOffset,
      y: placement.y + verticalOffset + leading,
      tracking: placementTracking[placementIndex] ?? 0,
    };
  });
  const rules = Array.from({ length: lineCount }, (_, lineIndex) => {
    const paragraph = lineParagraphs[lineIndex];
    if (paragraph === undefined) return [];
    const column = lineColumns[lineIndex] ?? 0;
    const verticalRemaining = Number.isFinite(maxHeight)
      ? maxHeight - (columnHeights[column] ?? 0)
      : 0;
    const verticalOffset = verticalTextOffset(verticalAlign, hasTrailingEmptyParagraph
      ? Math.max(0, verticalRemaining) : verticalRemaining);
    return [
      { rule: paragraph.ruleAbove, boundary: lineIndex === 0 || lineParagraphs[lineIndex - 1]?.index !== paragraph.index, after: 0 },
      { rule: paragraph.ruleBelow, boundary: lineIndex + 1 === lineCount || lineParagraphs[lineIndex + 1]?.index !== paragraph.index, after: lineHeights[lineIndex] ?? lineHeight },
    ].flatMap(({ rule, boundary, after }) => {
      if (rule === undefined || !boundary) return [];
      const offsetX = rule.offsetX * scale;
      return [{
        x: column * (maxWidth + columnSpacing) + offsetX,
        y: (lineOffsets[lineIndex] ?? 0) + after + rule.offsetY * scale + verticalOffset,
        width: maxWidth * rule.width - 2 * offsetX,
        color: rule.color,
        strokeWidth: rule.strokeWidth * scale,
      }];
    });
  }).flat();
  const minimumX = alignedPlacements.reduce(
    (minimum, placement) => Math.min(minimum, placement.x),
    0,
  );
  const maximumX = alignedPlacements.reduce(
    (maximum, placement) => Math.max(maximum, placement.x + placement.width),
    0,
  );
  return {
    runs: alignedPlacements,
    rules,
    width: maximumX - minimumX,
    height: columnHeights.reduce((maximum, value) => Math.max(maximum, value), 0),
    lineCount,
    lineHeight: lineHeights.slice(0, lineCount).reduce(
      (maximum, value) => Math.max(maximum, value),
      lineHeight,
    ),
    scale,
  };
}

export function layoutTextRuns(
  context: TextMeasureContext,
  runs: readonly SceneTextRun[],
  options: TextLayoutOptions,
): TextLayout {
  const maxWidth = Math.max(1, options.maxWidth);
  if (options.fillCharacter !== undefined) {
    const { offset, character } = options.fillCharacter;
    let cursor = 0;
    const index = runs.findIndex((run) => {
      if (offset >= cursor && offset <= cursor + run.text.length) return true;
      cursor += run.text.length;
      return false;
    });
    const run = runs[index];
    if (run !== undefined && Number.isSafeInteger(offset) && offset >= 0
      && [...character].length === 1 && !/[\r\n\t]/u.test(character)) {
      const measure = (text: string, item: SceneTextRun): number => {
        applyRunFont(context, item, options.fontScale ?? 1, options.resolveFontFamily);
        return measureRunText(context, text, item);
      };
      const width = runs.reduce((sum, item) => sum + measure(item.text, item), 0);
      const glyphWidth = measure(character, run);
      // ponytail: cap decorative expansion at 4096 glyphs; use tiled painting for larger cells.
      const count = glyphWidth > 0 && Number.isFinite(glyphWidth)
        ? Math.max(0, Math.min(4096, Math.floor((maxWidth - width) / glyphWidth))) : 0;
      if (count > 0) {
        const split = offset - cursor;
        const start = run.sourceStart ?? cursor;
        const end = run.sourceEnd ?? start + run.text.length;
        runs = [...runs.slice(0, index),
          { ...run, text: run.text.slice(0, split), sourceStart: start, sourceEnd: start + split },
          { ...run, text: character.repeat(count), sourceStart: start + split, sourceEnd: start + split },
          { ...run, text: run.text.slice(split), sourceStart: start + split, sourceEnd: end },
          ...runs.slice(index + 1)];
      }
    }
  }
  const baseLineHeight = options.lineHeight !== undefined && options.lineHeight > 0
    ? options.lineHeight
    : Math.max(1, ...runs.map((run) => run.fontSize * 1.2));
  const lineSpacingReduction = options.lineSpacingReduction !== undefined
    && Number.isFinite(options.lineSpacingReduction)
    ? Math.min(0.99, Math.max(0, options.lineSpacingReduction))
    : 0;
  const reducedLineHeight = baseLineHeight * (1 - lineSpacingReduction);
  const layoutMaxHeight = options.maxHeight !== undefined
    && Number.isFinite(options.maxHeight)
    && options.maxHeight > 0
    ? options.maxHeight
    : Number.POSITIVE_INFINITY;
  let sourceCursor = 0;
  const preparedRuns = runs.map((run) => {
    const start = run.sourceStart ?? sourceCursor;
    const end = run.sourceEnd ?? start + run.text.length;
    sourceCursor = end;
    const exactOffsets = end - start === run.text.length;
    return {
      run,
      horizontalTokens: indexedTextParts(
        textTokens(run.text, run.eastAsianLineBreaks !== false),
        exactOffsets ? start : 0,
      ).map((part) => exactOffsets ? part
        : start === end || part.start === 0 && part.end === run.text.length ? { text: part.text, start, end }
          : { text: part.text }),
      verticalTokens: indexedTextParts(graphemes(run.text), exactOffsets ? start : 0)
        .map((part) => exactOffsets ? part
          : start === end || part.start === 0 && part.end === run.text.length ? { text: part.text, start, end }
            : { text: part.text }),
    };
  });
  const tabStops = (options.tabStops ?? [])
    .map((stop): SceneTextTabStop => typeof stop === "number"
      ? { position: stop, align: "start", leader: "none" }
      : stop.align === "end" && stop.position > maxWidth
        ? { ...stop, position: Math.max(1, maxWidth - 1) }
        : stop)
    .filter((stop) => Number.isFinite(stop.position) && stop.position > 0)
    .slice()
    .sort((left, right) => left.position - right.position);
  let hasTrailingEmptyParagraph = false;
  for (let index = runs.length - 1; index >= 0; index -= 1) {
    const text = runs[index]?.text;
    if (text === undefined || text.length === 0) continue;
    hasTrailingEmptyParagraph = /(?:\r\n|[\n\u2028])$/u.test(text);
    break;
  }
  const prefixGraphemes = options.prefix === undefined ? [] : graphemes(options.prefix);
  const baseParagraphSpacing = options.paragraphSpacing !== undefined
    && Number.isFinite(options.paragraphSpacing)
    && options.paragraphSpacing > 0
    ? options.paragraphSpacing
    : 0;
  const layoutAtScale = (scale: number): TextLayout => {
    const lineHeight = reducedLineHeight * scale;
    const paragraphSpacing = baseParagraphSpacing * scale;
    const orientation = options.orientation ?? "horizontal";
    if (orientation === "vertical-rl" || orientation === "vertical-lr"
      || orientation === "stacked-rl" || orientation === "stacked-lr") {
      const placements: TextLayoutRun[] = [];
      const columnHeights: number[] = [0];
      const defaultTabStop = options.defaultTabStop !== undefined
        && Number.isFinite(options.defaultTabStop)
        && options.defaultTabStop > 0
        ? options.defaultTabStop
        : 36;
      let column = 0;
      let y = 0;
      const newColumn = (): void => {
        columnHeights[column] = y;
        column += 1;
        columnHeights[column] = 0;
        y = 0;
      };
      const mixed = orientation === "vertical-rl" || orientation === "vertical-lr";
      const place = (
        part: Readonly<{ text: string; start?: number; end?: number }>,
        style: SceneTextRun,
        sideways = false,
      ): void => {
        const token = part.text;
        if (token === "\n" || token === "\r\n" || token === "\u2028") {
          newColumn();
          return;
        }
        if (token === "\t") {
          y = tabStops.find((stop) => stop.position > y)?.position
            ?? ((Math.floor(y / defaultTabStop) + 1) * defaultTabStop);
          columnHeights[column] = y;
          return;
        }
        if (!mixed && token.trim().length === 0) {
          if (y > 0) newColumn();
          return;
        }
        applyRunFont(context, style, scale, options.resolveFontFamily);
        const width = measureRunText(context, token, style);
        const advance = sideways ? width : lineHeight;
        if (y > 0 && y + advance > layoutMaxHeight) newColumn();
        placements.push({
          text: token,
          ...(part.start === undefined || part.end === undefined ? {} : {
            start: part.start,
            end: part.end,
          }),
          x: 0,
          y,
          width,
          line: column,
          style,
          sideways,
        });
        y += advance;
        columnHeights[column] = y;
      };
      const prefixStyle = runs[0];
      if (prefixStyle !== undefined && options.prefix !== undefined) {
        for (const token of prefixGraphemes) place({ text: token }, prefixStyle);
      }
      for (const prepared of preparedRuns) {
        for (const part of mixed ? prepared.horizontalTokens : prepared.verticalTokens) {
          place(
            part,
            prepared.run,
            mixed && !CJK_SCRIPT.test(part.text) && !EAST_ASIAN_PUNCTUATION.test(part.text),
          );
        }
      }
      const lineCount = placements.length === 0 && column === 0 ? 0 : column + 1;
      const alignedPlacements = placements.map((placement) => {
        const columnHeight = columnHeights[placement.line] ?? 0;
        const remaining = Number.isFinite(layoutMaxHeight)
          ? layoutMaxHeight - columnHeight
          : 0;
        // Authored paragraphs, including empty ones, share the overflow anchor.
        // Preserve the historical clamp only for unstructured text callers.
        const anchoredRemaining = hasTrailingEmptyParagraph
          && (options.paragraphs?.length ?? 0) === 0
          ? Math.max(0, remaining)
          : remaining;
        const yOffset = verticalTextOffset(options.verticalAlign, anchoredRemaining);
        const x = orientation === "vertical-rl" || orientation === "stacked-rl"
          ? maxWidth - (placement.line + 1) * lineHeight
          : placement.line * lineHeight;
        return { ...placement, x, y: placement.y + yOffset };
      });
      return {
        runs: alignedPlacements,
        rules: [],
        width: lineCount * lineHeight,
        height: columnHeights.slice(0, lineCount).reduce((maximum, value) => Math.max(maximum, value), 0),
        lineCount,
        lineHeight,
        scale,
      };
    }
    const placements: TextLayoutRun[] = [];
    const lineWidths: number[] = [0];
    const lineStarts: number[] = [0];
    const lineOffsets: number[] = [0];
    const lineHeights: number[] = [lineHeight];
    const lineColumns: number[] = [0];
    const lineJustifiable: boolean[] = [false];
    const lineCompressionCapacity: number[] = [0];
    const lineHasTab: boolean[] = [false];
    const columns = options.columns !== undefined && Number.isInteger(options.columns)
      ? Math.min(64, Math.max(1, options.columns))
      : 1;
    const columnSpacing = options.columnSpacing !== undefined
      && Number.isFinite(options.columnSpacing)
      ? Math.max(0, options.columnSpacing)
      : 0;
    const lineParagraphs: ResolvedTextParagraph[] = [];
    const authoredParagraphs = options.paragraphs ?? [];
    const legacyFirstLineIndent = options.firstLineIndent !== undefined
      && Number.isFinite(options.firstLineIndent)
      ? options.firstLineIndent
      : 0;
    const legacyDefaultTabStop = options.defaultTabStop !== undefined
      && Number.isFinite(options.defaultTabStop)
      && options.defaultTabStop > 0
      ? options.defaultTabStop
      : 36;
    let legacyContinuationX = 0;
    const paragraphAt = (index: number) => {
      const authored = authoredParagraphs[index];
      if (authored === undefined) {
        return {
          align: options.align ?? "start",
          marginLeft: 0,
          marginRight: 0,
          firstLineIndent: index === 0 ? legacyFirstLineIndent : legacyContinuationX,
          defaultTabStop: legacyDefaultTabStop,
          continuationX: legacyContinuationX,
          lineHeight,
          spaceBefore: 0,
          spaceAfter: paragraphSpacing,
          // Preserve the historical renderer behavior for non-DrawingML callers.
          latinLineBreak: true,
          hangingPunctuation: false,
          authored: false,
          index,
        } as const;
      }
      const marginLeft = Number.isFinite(authored.marginLeft) ? Math.max(0, authored.marginLeft) : 0;
      return {
        align: authored.align,
        marginLeft,
        marginRight: Number.isFinite(authored.marginRight) ? Math.max(0, authored.marginRight) : 0,
        firstLineIndent: Number.isFinite(authored.firstLineIndent) ? authored.firstLineIndent : 0,
        defaultTabStop: Number.isFinite(authored.defaultTabStop) && authored.defaultTabStop > 0
          ? authored.defaultTabStop
          : legacyDefaultTabStop,
        continuationX: marginLeft,
        lineHeight: Number.isFinite(authored.lineHeight) && (authored.lineHeight ?? 0) > 0
          ? (authored.lineHeight ?? lineHeight) * (1 - lineSpacingReduction) * scale
          : lineHeight,
        spaceBefore: Number.isFinite(authored.spaceBefore)
          ? Math.max(0, authored.spaceBefore ?? 0) * scale
          : 0,
        spaceAfter: Number.isFinite(authored.spaceAfter)
          ? Math.max(0, authored.spaceAfter ?? 0) * scale
          : 0,
        // Office defaults from [MS-OE376] §5.1.5.4.13 also cover old protocols.
        latinLineBreak: authored.latinLineBreak ?? false,
        hangingPunctuation: authored.hangingPunctuation ?? true,
        authored: true,
        index,
        ruleAbove: authored.ruleAbove,
        ruleBelow: authored.ruleBelow,
        dropCap: authored.dropCap,
      } as const;
    };
    let line = 0;
    let paragraphIndex = 0;
    let softBreakPending = false;
    let paragraph = paragraphAt(paragraphIndex);
    lineParagraphs[0] = paragraph;
    lineOffsets[0] = paragraph.spaceBefore;
    lineHeights[0] = paragraph.lineHeight;
    let x = Math.max(0, paragraph.marginLeft + paragraph.firstLineIndent);
    lineWidths[0] = x;
    lineStarts[0] = x;
    const lineBaseStarts: number[] = [x];
    const dropCapRegions: Rect[] = [];
    let dropCapRemaining = paragraph.dropCap?.characters ?? 0;
    let dropCapX = x;
    let dropCapBottom = 0;
    const applyLineWrap = (minimumWidth = 1): void => {
      const regions = dropCapRegions.length === 0 ? options.wrapRegions ?? []
        : [...(options.wrapRegions ?? []), ...dropCapRegions];
      if (regions.length === 0) return;
      const columnX = (lineColumns[line] ?? 0) * (maxWidth + columnSpacing);
      const startX = lineBaseStarts[line] ?? x;
      minimumWidth = Math.min(minimumWidth, maxWidth - paragraph.marginRight - startX);
      let top = lineOffsets[line] ?? 0;
      // ponytail: rectangular bands choose the widest free interval; curved
      // contours need polygon intersections in this same layout pass.
      for (let attempt = 0; attempt <= regions.length; attempt += 1) {
        let intervals = [[startX, maxWidth - paragraph.marginRight]];
        let nextBottom = Number.POSITIVE_INFINITY;
        for (const region of regions) {
          const left = region.x - columnX;
          const right = left + region.width;
          if (top + paragraph.lineHeight <= region.y || top >= region.y + region.height
            || right <= startX || left >= maxWidth - paragraph.marginRight) continue;
          nextBottom = Math.min(nextBottom, region.y + region.height);
          intervals = intervals.flatMap(([start = 0, end = 0]) => {
            if (right <= start || left >= end) return [[start, end]];
            return [...(left > start ? [[start, left]] : []), ...(right < end ? [[right, end]] : [])];
          });
        }
        const best = intervals.reduce<number[] | undefined>((best, interval) =>
          best === undefined || interval[1]! - interval[0]! > best[1]! - best[0]! ? interval : best, undefined);
        if (best !== undefined && best[1]! - best[0]! >= minimumWidth) {
          lineParagraphs[line] = {...paragraph,
            marginLeft: paragraph.marginLeft + best[0]! - startX,
            marginRight: maxWidth - best[1]!};
          x = best[0]!;
          lineStarts[line] = x;
          lineWidths[line] = x;
          lineOffsets[line] = top;
          return;
        }
        if (!Number.isFinite(nextBottom) || nextBottom <= top) return;
        top = nextBottom;
      }
    };
    const lineRight = (): number => maxWidth - (lineParagraphs[line]?.marginRight ?? paragraph.marginRight);
    applyLineWrap();
    const horizontalItems = preparedRuns.flatMap(({ run, horizontalTokens }) => (
      horizontalTokens.map((part) => ({ ...part, token: part.text, run }))
    ));
    const listMarkerRun = paragraph.authored && runs.length > 1 && runs[0]?.text.endsWith("\t")
      && /\S/u.test(runs[0].text)
      ? runs[0]
      : undefined;
    const followingTabWidth = (startIndex: number): number => {
      let width = 0;
      for (let index = startIndex; index < horizontalItems.length; index += 1) {
        const item = horizontalItems[index];
        if (item === undefined || item.token === "\t" || item.token === "\n" || item.token === "\r\n" || item.token === "\u2028") {
          break;
        }
        applyRunFont(context, item.run, scale, options.resolveFontFamily);
        const tokens = paragraph.latinLineBreak ? latinLineBreakTokens(item.token) : [item.token];
        for (const token of tokens) {
          width += trackingPieces(token, paragraph.align)
            .reduce((total, piece) => total + measureRunText(context, piece, item.run), 0);
        }
      }
      return width;
    };
    const crossRunWordWidth = (startIndex: number): number | undefined => {
      const startsWord = (token: string): boolean => token !== "\t"
        && token !== "\n"
        && token !== "\r\n"
        && token !== "\u2028"
        && !isSoftWrapSpace(token)
        && !isLineStartProhibited(token)
        && !CJK_SCRIPT.test(token);
      const previous = horizontalItems[startIndex - 1];
      if (previous !== undefined
        && startsWord(previous.token)
        && !/[-\u2010]$/u.test(previous.token)) return undefined;
      let width = 0;
      let lastWidth = 0;
      let lastItem: (typeof horizontalItems)[number] | undefined;
      let crossedRun = false;
      for (let index = startIndex; index < horizontalItems.length; index += 1) {
        const item = horizontalItems[index];
        if (item === undefined || !startsWord(item.token)) break;
        if (lastItem !== undefined && lastItem.run !== item.run) crossedRun = true;
        applyRunFont(context, item.run, scale, options.resolveFontFamily);
        lastWidth = trackingPieces(item.token, paragraph.align)
          .reduce((total, piece) => total + measureRunText(context, piece, item.run), 0);
        width += lastWidth;
        lastItem = item;
        if (/[-\u2010]$/u.test(item.token)) break;
      }
      if (!crossedRun || lastItem === undefined) return undefined;
      return width - lastWidth + wrappingWidth(
        context,
        lastItem.token,
        lastWidth,
        paragraph.hangingPunctuation,
        lastItem.run,
      );
    };
    const advanceTab = (
      style: SceneTextRun,
      contentWidth: number,
      generatedListMarker: boolean,
      start?: number,
      end?: number,
    ): void => {
      // Preserve the body anchor unless a wide marker needs a normal space to clear it.
      const startX = x;
      const stop = tabStops.find((candidate) => candidate.position > x);
      const alignedTab = stop !== undefined && stop.align !== "start";
      const anchorsHangingBody = paragraph.authored
        && paragraph.firstLineIndent < 0
        && (lineStarts[line] ?? paragraph.marginLeft) < paragraph.marginLeft
        && !lineHasTab[line];
      lineHasTab[line] = true;
      lineCompressionCapacity[line] = 0;
      // Explicit aligned fields (for example TOC page numbers) are not list-marker tabs.
      if (!alignedTab && (anchorsHangingBody || (paragraph.authored && x < paragraph.marginLeft))) {
        applyRunFont(context, style, scale, options.resolveFontFamily);
        x = x < paragraph.marginLeft ? paragraph.marginLeft
          : x + measureRunText(context, " ", style);
      } else if (generatedListMarker && stop === undefined) {
        x = Math.max(x, paragraph.marginLeft + paragraph.defaultTabStop);
      } else {
        if (stop === undefined) {
          x = (Math.floor(x / paragraph.defaultTabStop) + 1) * paragraph.defaultTabStop;
        } else {
          // An overfull aligned tab extends forward instead of overwriting preceding text.
          x = Math.max(startX, stop.align === "center"
            ? stop.position - contentWidth / 2
            : stop.align === "end" ? stop.position - contentWidth : stop.position);
        }
      }
      if (x > startX) {
        placements.push({
          text: "\t",
          ...(start === undefined || end === undefined ? {} : { start, end }),
          x: startX,
          y: lineOffsets[line] ?? 0,
          width: x - startX,
          line,
          style,
          height: lineHeights[line] ?? lineHeight,
          ...(stop?.leader === undefined || stop.leader === "none" ? {} : { leader: stop.leader }),
        });
      }
      lineWidths[line] = x;
    };
    const prefixStyle = runs[0];
    let prefixWidth = 0;
    if (prefixStyle !== undefined && options.prefix !== undefined && options.prefix !== "") {
      applyRunFont(context, prefixStyle, scale, options.resolveFontFamily);
      prefixWidth = measureRunText(context, options.prefix.replace(/\t$/u, ""), prefixStyle);
      placements.push({
        text: options.prefix.replace(/\t$/u, ""),
        x,
        y: lineOffsets[0] ?? 0,
        width: prefixWidth,
        line: 0,
        style: prefixStyle,
        height: lineHeights[0],
      });
      x += prefixWidth;
      if (options.prefix.endsWith("\t")) advanceTab(prefixStyle, 0, true);
      lineWidths[0] = x;
    }
    legacyContinuationX = options.hangingIndent === undefined
      ? prefixWidth
      : Number.isFinite(options.hangingIndent) ? Math.max(0, options.hangingIndent) : prefixWidth;
    if (!paragraph.authored) {
      paragraph = paragraphAt(paragraphIndex);
      lineParagraphs[0] = {...paragraph,
        marginLeft: lineParagraphs[0]?.marginLeft ?? paragraph.marginLeft,
        marginRight: lineParagraphs[0]?.marginRight ?? paragraph.marginRight};
    }
    const newLine = (paragraphBreak: boolean): void => {
      lineWidths[line] = x;
      lineJustifiable[line] = !paragraphBreak;
      const previousParagraph = paragraph;
      const previousLineHeight = lineHeights[line] ?? lineHeight;
      if (paragraphBreak) paragraphIndex += 1;
      const nextParagraph = paragraphAt(paragraphIndex);
      let nextOffset = (lineOffsets[line] ?? 0)
        + previousLineHeight
        + (paragraphBreak ? previousParagraph.spaceAfter + nextParagraph.spaceBefore : 0);
      if (paragraphBreak) {
        nextOffset = Math.max(nextOffset, dropCapBottom + previousParagraph.spaceAfter + nextParagraph.spaceBefore);
        dropCapBottom = 0;
        dropCapRemaining = nextParagraph.dropCap?.characters ?? 0;
        dropCapX = Math.max(0, nextParagraph.marginLeft + nextParagraph.firstLineIndent);
      }
      let nextColumn = lineColumns[line] ?? 0;
      if (columns > 1
        && Number.isFinite(layoutMaxHeight)
        && nextOffset + nextParagraph.lineHeight > layoutMaxHeight
        && nextColumn + 1 < columns) {
        nextColumn += 1;
        nextOffset = paragraphBreak ? nextParagraph.spaceBefore : 0;
      }
      line += 1;
      paragraph = nextParagraph;
      lineParagraphs[line] = paragraph;
      lineHeights[line] = paragraph.lineHeight;
      lineColumns[line] = nextColumn;
      lineJustifiable[line] = false;
      lineCompressionCapacity[line] = 0;
      lineHasTab[line] = false;
      x = paragraphBreak
        ? Math.max(0, paragraph.marginLeft + paragraph.firstLineIndent)
        : paragraph.continuationX;
      lineWidths[line] = x;
      lineStarts[line] = x;
      lineOffsets[line] = nextOffset;
      lineBaseStarts[line] = x;
      applyLineWrap();
    };
    for (const [itemIndex, item] of horizontalItems.entries()) {
      let rawToken = item.token;
      let tokenStart = item.start;
      const run = item.run;
      const cap = paragraph.dropCap;
      if (cap !== undefined && dropCapRemaining > 0 && !/[\n\r]/u.test(rawToken)) {
        // Relative-size initials (Pages) use live glyph bounds. DOCX framed
        // initials already carry an absolute font size and baseline from Core.
        if (dropCapRemaining === cap.characters) dropCapX = x;
        const text = [...rawToken].slice(0, dropCapRemaining).join("");
        dropCapRemaining -= [...text].length;
        applyRunFont(context, run, scale, options.resolveFontFamily);
        context.textBaseline = "alphabetic";
        const bodyMetrics = context.measureText("H");
        const bodyAscent = bodyMetrics.actualBoundingBoxAscent || run.fontSize * scale * .7;
        const glyph = context.measureText(text);
        const glyphHeight = (glyph.actualBoundingBoxAscent || bodyAscent) + (glyph.actualBoundingBoxDescent || 0);
        const height = (cap.lines - 1) * paragraph.lineHeight + bodyAscent;
        const style = { ...run, fontSize: run.fontSize * height / glyphHeight, text };
        applyRunFont(context, style, scale, options.resolveFontFamily);
        const measured = context.measureText(text);
        const ascent = measured.actualBoundingBoxAscent || height;
        const descent = measured.actualBoundingBoxDescent || 0;
        const width = measureRunText(context, text, style);
        const top = lineOffsets[line] ?? 0;
        const leading = Math.max(0, paragraph.lineHeight - run.fontSize * 1.2 * scale);
        const inkTop = top + leading + (bodyMetrics.fontBoundingBoxAscent || run.fontSize * scale) - bodyAscent
          - cap.raisedLines * paragraph.lineHeight;
        placements.push({ text, ...(tokenStart === undefined ? {} : { start: tokenStart, end: tokenStart + text.length }),
          x: dropCapX - cap.outdent * scale, y: inkTop, width, line, style,
          height: ascent + descent, baselineOffset: ascent });
        dropCapX += width;
        const columnX = (lineColumns[line] ?? 0) * (maxWidth + columnSpacing);
        dropCapRegions.push({ x: columnX + paragraph.marginLeft - cap.outdent * scale, y: top,
          width: Math.max(0, dropCapX - paragraph.marginLeft + cap.padding * scale),
          height: cap.lines * paragraph.lineHeight });
        dropCapBottom = Math.max(dropCapBottom, top + cap.lines * paragraph.lineHeight);
        applyLineWrap();
        rawToken = rawToken.slice(text.length);
        if (tokenStart !== undefined) tokenStart += text.length;
        if (rawToken === "") continue;
      }
        // Core flow paragraphs keep a fitting URL together; only oversized URLs
        // split into graphemes. DrawingML retains its explicit Latin break policy.
        const tokenTexts = paragraph.latinLineBreak
          && (paragraph.authored || !/^(?:https?:\/\/|www\.)/iu.test(rawToken))
          ? latinLineBreakTokens(rawToken) : [rawToken];
        const tokens: readonly IndexedTextPart[] = item.start === undefined || item.end === undefined
          ? tokenTexts.map((text) => ({ text }))
          : indexedTextParts(tokenTexts, tokenStart ?? item.start, item.end);
        for (const tokenPart of tokens) {
          const token = tokenPart.text;
          if (token === "\n" || token === "\r\n" || token === "\u2028") {
            softBreakPending = false;
            const paragraphBreak = token !== "\u2028" && (authoredParagraphs.length === 0
              || paragraphIndex + 1 < authoredParagraphs.length);
            newLine(paragraphBreak);
            continue;
          }
          if (token === "\t") {
            advanceTab(run, followingTabWidth(itemIndex + 1), run === listMarkerRun, item.start, item.end);
            continue;
          }
          const softWrapSpace = isSoftWrapSpace(token);
          if (softBreakPending && !softWrapSpace && !isLineStartProhibited(token)) {
            newLine(false);
            softBreakPending = false;
          }
          const usableWidth = Math.max(1, lineRight() - paragraph.continuationX);
          const styledWordWidth = crossRunWordWidth(itemIndex);
          if (options.wrap !== false
            && styledWordWidth !== undefined
            && styledWordWidth <= usableWidth
            && x > (lineStarts[line] ?? paragraph.marginLeft)
            && x + styledWordWidth - (lineCompressionCapacity[line] ?? 0) > lineRight()) {
            newLine(false);
          }
          applyRunFont(context, run, scale, options.resolveFontFamily);
          const trackedPieces = options.compressPunctuation === true
            && graphemes(token).some((text) => punctuationCompression(text, 1) > 0)
            ? graphemes(token) : trackingPieces(token, paragraph.align);
          const trackedPieceWidths = trackedPieces.map((piece) => measureRunText(context, piece, run));
          const tokenWidth = trackedPieceWidths.reduce((total, width) => total + width, 0);
          const tokenCompression = options.compressPunctuation === true && !lineHasTab[line]
            ? trackedPieces.reduce((total, piece, index) => total + punctuationCompression(piece, trackedPieceWidths[index] ?? 0), 0)
            : 0;
          const tokenWrappingWidth = wrappingWidth(
            context,
            token,
            tokenWidth,
            paragraph.hangingPunctuation,
            run,
          );
          // Whole-word wrapping still splits words wider than a full body line.
          const oversized = options.wrap !== false
            && tokenWrappingWidth > usableWidth
            && token.trim() !== "";
          const pieces = oversized ? graphemes(token) : trackedPieces;
          const indexedPieces: readonly IndexedTextPart[] = tokenPart.start === undefined || tokenPart.end === undefined
            ? pieces.map((text) => ({ text }))
            : indexedTextParts(pieces, tokenPart.start, tokenPart.end);
          const pieceWidths = oversized
            ? pieces.map((piece) => measureRunText(context, piece, run))
            : trackedPieceWidths;
          if (options.wrap !== false
            && softWrapSpace
            && x > paragraph.marginLeft
            && x + tokenWidth > lineRight()) {
            placements.push({
              text: token,
              ...(tokenPart.start === undefined || tokenPart.end === undefined ? {} : {
                start: tokenPart.start,
                end: tokenPart.end,
              }),
              x,
              y: lineOffsets[line] ?? 0,
              width: 0,
              line,
              style: run,
              height: lineHeights[line] ?? lineHeight,
            });
            lineWidths[line] = x;
            softBreakPending = true;
            continue;
          }
          if (options.wrap !== false
            && x > (lineStarts[line] ?? paragraph.marginLeft)
            && (x + tokenWrappingWidth - (lineCompressionCapacity[line] ?? 0) - tokenCompression > lineRight() + (options.compressPunctuation === true ? 0.001 : 0)
              || (options.compressPunctuation === true && x > lineRight() + 0.001))
            && !oversized) {
            if (!isLineStartProhibited(token)) {
              newLine(false);
            }
          }
          if (!softWrapSpace && !oversized && x === lineStarts[line]) {
            applyLineWrap(Math.max(tokenWrappingWidth, styledWordWidth ?? 0));
          }
          for (const [pieceIndex, piecePart] of indexedPieces.entries()) {
            const piece = piecePart.text;
            const width = pieceWidths[pieceIndex] ?? 0;
            const pieceWrappingWidth = wrappingWidth(
              context,
              piece,
              width,
              paragraph.hangingPunctuation,
              run,
            );
            if (options.wrap !== false
              && oversized
              && x > (lineStarts[line] ?? paragraph.marginLeft)
              && x + pieceWrappingWidth > lineRight()) {
              newLine(false);
            }
            if (x === lineStarts[line] && piece.trim()) applyLineWrap(width);
            placements.push({
              text: piece,
              ...(piecePart.start === undefined || piecePart.end === undefined ? {} : {
                start: piecePart.start,
                end: piecePart.end,
              }),
              x,
              y: lineOffsets[line] ?? 0,
              width,
              line,
              style: run,
              height: lineHeights[line] ?? lineHeight,
            });
            x += width;
            if (options.compressPunctuation === true && !lineHasTab[line]) {
              lineCompressionCapacity[line] = (lineCompressionCapacity[line] ?? 0) + punctuationCompression(piece, width);
            }
            lineWidths[line] = x;
          }
        }
    }
    if (options.continuesAfter === true) lineJustifiable[line] = true;
    if (options.compressPunctuation === true) {
      let firstPlacement = 0;
      for (let lineIndex = 0; lineIndex <= line; lineIndex += 1) {
        let end = firstPlacement;
        while (end < placements.length && placements[end]?.line === lineIndex) end += 1;
        const capacity = lineHasTab[lineIndex] ? 0 : lineCompressionCapacity[lineIndex] ?? 0;
        const overflow = Math.max(0, (lineWidths[lineIndex] ?? 0) - maxWidth + (lineParagraphs[lineIndex]?.marginRight ?? 0));
        const ratio = capacity > 0 ? Math.min(1, overflow / capacity) : 0;
        let shift = 0;
        if (ratio > 0) {
          for (let index = firstPlacement; index < end; index += 1) {
            const placement = placements[index]!;
            const compression = punctuationCompression(placement.text, placement.width) * ratio;
            placements[index] = {...placement, x: placement.x - shift, width: placement.width - compression,
              ...(compression > 0 ? {compression} : {})};
            shift += compression;
          }
          lineWidths[lineIndex] = (lineWidths[lineIndex] ?? 0) - shift;
        }
        firstPlacement = end;
      }
    }
    const lineCount = placements.length === 0 && line === 0 ? 0 : line + 1;
    const align = options.align ?? "start";
    const direction = options.direction ?? "ltr";
    return finalizeHorizontalTextLayout({
      fixedLineHeight: options.fixedLineHeight ?? false,
      placements,
      lineWidths,
      lineStarts,
      lineOffsets,
      lineHeights,
      lineColumns,
      lineJustifiable,
      lineHasTab,
      lineParagraphs,
      fallbackParagraph: paragraphAt(0),
      lineCount,
      columns,
      columnSpacing,
      lineHeight,
      scale,
      // Fit the natural text block before alignment; outer alignment space must
      // not become part of the width subsequently stretched to the frame.
      maxWidth: options.autoFit === "fit-frame"
        ? lineWidths.reduce((width, line) => Math.max(width, line), 0)
        : maxWidth,
      maxHeight: layoutMaxHeight,
      align,
      direction,
      verticalAlign: options.verticalAlign,
      hasTrailingEmptyParagraph: hasTrailingEmptyParagraph && authoredParagraphs.length === 0,
    });
  };

  const fontScale = options.fontScale !== undefined && Number.isFinite(options.fontScale)
    ? Math.min(1, Math.max(0.01, options.fontScale))
    : 1;
  const unscaled = layoutAtScale(fontScale);
  if (options.autoFit === "fit-frame"
    && Number.isFinite(layoutMaxHeight)
    && unscaled.height > 0) {
    return layoutAtScale(fontScale * layoutMaxHeight / unscaled.height);
  }
  if (options.autoFit !== "shrink") return unscaled;
  const totalWidth = maxWidth * (options.columns ?? 1)
    + Math.max(0, (options.columns ?? 1) - 1) * (options.columnSpacing ?? 0);
  const fits = (layout: TextLayout): boolean => (
    layout.width <= totalWidth + 0.001 && layout.height <= layoutMaxHeight + 0.001
  );
  if (fits(unscaled)) return unscaled;
  const minScale = options.minScale !== undefined && Number.isFinite(options.minScale)
    ? Math.min(1, Math.max(0.01, options.minScale))
    : 0.1;
  let low = Math.min(minScale, fontScale);
  let high = fontScale;
  let best = layoutAtScale(low);
  for (let iteration = 0; iteration < 16; iteration += 1) {
    const middle = (low + high) / 2;
    const candidate = layoutAtScale(middle);
    if (fits(candidate)) {
      low = middle;
      best = candidate;
    } else high = middle;
  }
  return best;
}

export interface SheetAutoFitCell {
  readonly text: string;
  readonly width: number;
  readonly wrap: boolean;
  readonly fontRuns?: readonly DocumentFontRun[];
}

const SHEET_AUTO_FIT_DEFAULT_FONT_SIZE = 11 * 96 / 72;

function sheetAutoFitTextRuns(cell: SheetAutoFitCell): readonly SceneTextRun[] {
  const runs: SceneTextRun[] = [];
  let cursor = 0;
  const append = (text: string, font?: DocumentFontRun): void => {
    if (text.length === 0) return;
    runs.push({
      text,
      fontFamily: font?.renderedFamily ?? "Arial",
      fontSize: font?.fontSize ?? SHEET_AUTO_FIT_DEFAULT_FONT_SIZE,
      color: 0x0000_00ff,
      bold: font?.bold ?? false,
      italic: font?.italic ?? false,
      underline: false,
      strikethrough: false,
      highlight: 0,
      baselineShift: 0,
      letterSpacing: 0,
    });
  };
  for (const font of cell.fontRuns ?? []) {
    if (font.start < cursor || font.end <= font.start || font.end > cell.text.length) continue;
    append(cell.text.slice(cursor, font.start));
    append(cell.text.slice(font.start, font.end), font);
    cursor = font.end;
  }
  append(cell.text.slice(cursor));
  return runs;
}

export function sheetAutoFitColumnWidth(
  context: TextMeasureContext,
  cells: readonly SheetAutoFitCell[],
  minimumWidth: number,
): number {
  let width = minimumWidth;
  for (const cell of cells) {
    const runs = sheetAutoFitTextRuns(cell);
    const lineHeight = Math.max(SHEET_AUTO_FIT_DEFAULT_FONT_SIZE,
      ...runs.map((run) => run.fontSize)) * 1.2;
    const layout = layoutTextRuns(context, runs, {
      maxWidth: Number.MAX_SAFE_INTEGER,
      lineHeight,
      wrap: false,
    });
    width = Math.max(width, Math.ceil(layout.width + 5));
  }
  return width;
}

export function sheetAutoFitRowHeight(
  context: TextMeasureContext,
  cells: readonly SheetAutoFitCell[],
  minimumHeight: number,
): number {
  let height = minimumHeight;
  for (const cell of cells) {
    const runs = sheetAutoFitTextRuns(cell);
    const lineHeight = Math.max(SHEET_AUTO_FIT_DEFAULT_FONT_SIZE,
      ...runs.map((run) => run.fontSize)) * 1.2;
    const layout = layoutTextRuns(context, runs, {
      maxWidth: Math.max(1, cell.width - 4),
      lineHeight,
      wrap: cell.wrap,
    });
    height = Math.max(height, Math.ceil(layout.height + 2));
  }
  return height;
}

function visualTextRuns(
  object: SceneObject,
  visual: Extract<SceneVisual, { kind: "text" | "rich-text" }>,
): readonly SceneTextRun[] {
  if (visual.kind === "rich-text") return visual.runs;
  const text = object.text ?? "";
  return text === "" ? [] : [{
    text,
    fontFamily: visual.fontFamily,
    fontSize: visual.fontSize,
    color: visual.color,
    bold: visual.bold,
    italic: visual.italic,
    underline: false,
    strikethrough: false,
    highlight: 0,
    baselineShift: 0,
    letterSpacing: 0,
  }];
}

function positionedPdfRun(
  object: SceneObject,
  visual: Extract<SceneVisual, { kind: "text" | "rich-text" }>,
  textLayout: SceneTextLayout | undefined,
  runs: readonly SceneTextRun[],
): SceneTextRun | undefined {
  const run = runs.length === 1 ? runs[0] : undefined;
  return object.source.format === "pdf"
    && run !== undefined
    && visual.align === "start"
    && textLayout?.orientation === "horizontal"
    && textLayout.wrap === false
    && textLayout.columnCount === 1
    && textLayout.horizontalOverflow === "overflow"
    && textLayout.verticalOverflow === "overflow"
    && textLayout.rotationDegrees === 0
    && textLayout.warp === undefined
    && textLayout.prefix === undefined
    && (textLayout.paragraphs?.length ?? 0) === 0
    && !run.underline
    && !run.strikethrough
    && (run.highlight & 0xff) === 0
    && (textLayout.textBaseline ?? 0) > 0
    ? run
    : undefined;
}

function hasInvisibleTextGeometry(
  visual: Extract<SceneVisual, { kind: "text" | "rich-text" }>,
): boolean {
  return visual.kind === "text"
    ? (visual.fill & 0xff) === 0 && (visual.stroke & 0xff) === 0
    : visual.fill.kind === "none" && visual.stroke.kind === "none";
}

function drawPositionedPdfRun(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  run: SceneTextRun,
  textLayout: SceneTextLayout,
  fontFamily: string,
  strokeStyle?: SceneStrokeStyle,
  useEmbeddedAdvances = false,
  collectTextFragment?: TextFragmentCollector,
): void {
  const width = object.bounds.width;
  const insetLeft = Math.min(textLayout.insetLeft, width / 2);
  const insetRight = Math.min(textLayout.insetRight, width / 2);
  const marginLeft = Math.max(0, textLayout.marginLeft);
  const marginRight = Math.max(0, textLayout.marginRight);
  const availableWidth = Math.max(1, width - insetLeft - insetRight - marginLeft - marginRight);
  context.save();
  try {
    applyStrokeStyle(context, strokeStyle);
    context.textAlign = "left";
    context.direction = resolvedTextDirection([run], textLayout.direction);
    applyRunFont(context, run, 1, () => fontFamily);
    context.fontKerning = "none";
    context.textBaseline = "alphabetic";
    const textX = object.bounds.x + insetLeft + marginLeft;
    const textY = object.bounds.y + textLayout.insetTop
      + (textLayout.textBaseline ?? 0) - run.baselineShift;
    context.translate(textX, textY);
    const glyphPaint = run.paint ?? textLayout.textPaint;
    context.fillStyle = glyphPaint === undefined ? color(run.color)
      : paintStyle(context, object, glyphPaint, textX, textY) ?? color(run.color);
    const textStroke = textStrokeStyle(context, object, textLayout, textX, textY);
    const horizontalScale = runHorizontalScale(run);
    let measuredWidth: number | undefined;
    if ((textLayout.textScaleToFit ?? false)
      && (!useEmbeddedAdvances || (textLayout.textMatrixScaleToFit ?? false))) {
      measuredWidth = context.measureText(run.text).width;
      if (measuredWidth > 0) context.scale(availableWidth / measuredWidth, 1);
    } else if (horizontalScale !== 1) context.scale(horizontalScale, 1);
    const transform = collectTextFragment === undefined ? undefined : canvasTransform(context);
    if (transform !== undefined && collectTextFragment !== undefined) {
      measuredWidth ??= context.measureText(run.text).width;
    }
    if (transform !== undefined && collectTextFragment !== undefined
      && measuredWidth !== undefined && measuredWidth > 0) {
      const metrics = context.measureText("Mg");
      const ascent = Number.isFinite(metrics.fontBoundingBoxAscent) && metrics.fontBoundingBoxAscent > 0
        ? metrics.fontBoundingBoxAscent
        : run.fontSize * 0.8;
      const descent = Number.isFinite(metrics.fontBoundingBoxDescent) && metrics.fontBoundingBoxDescent >= 0
        ? metrics.fontBoundingBoxDescent
        : run.fontSize * 0.2;
      collectTextFragment({
        objectId: object.id,
        text: run.text,
        start: 0,
        end: run.text.length,
        width: measuredWidth,
        height: ascent + descent,
        line: 0,
        font: context.font,
        letterSpacing: run.letterSpacing,
        direction: context.direction,
      }, multiplyTransform(transform, {
        a: 1,
        b: 0,
        c: 0,
        d: 1,
        e: 0,
        f: -ascent,
      }));
    }
    if (textLayout.textFill ?? true) context.fillText(run.text, 0, 0);
    if (textStroke !== undefined) {
      context.strokeStyle = textStroke;
      context.lineWidth = effectiveStrokeWidth(
        context,
        object,
        textLayout.textStrokeWidth ?? 0,
      );
      context.strokeText(run.text, 0, 0);
    }
  } finally {
    context.restore();
  }
}

function textStrokeStyle(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  textLayout: SceneTextLayout | undefined,
  originX = 0,
  originY = 0,
): string | CanvasGradient | CanvasPattern | undefined {
  if (textLayout?.textStrokePaint !== undefined) {
    return paintStyle(context, object, textLayout.textStrokePaint, originX, originY);
  }
  return ((textLayout?.textStrokeColor ?? 0) & 0xff) === 0
    ? undefined
    : color(textLayout?.textStrokeColor ?? 0);
}

function resolvedTextDirection(
  runs: readonly SceneTextRun[],
  direction: SceneTextLayout["direction"] | undefined,
): "ltr" | "rtl" {
  if (direction === "ltr" || direction === "rtl") return direction;
  for (const run of runs) {
    for (const character of run.text) {
      if (/\p{Script=Hebrew}|\p{Script=Arabic}/u.test(character)) return "rtl";
      if (/\p{Letter}/u.test(character)) return "ltr";
    }
  }
  return "ltr";
}

function coalescePaintTextRuns(
  runs: readonly TextLayoutRun[],
  direction: "ltr" | "rtl",
  mergeWholeLine = false,
): readonly TextLayoutRun[] {
  const result: TextLayoutRun[] = [];
  for (const placement of runs) {
    const previous = result.at(-1);
    if (placement.baselineOffset !== undefined || previous?.baselineOffset !== undefined
      || placement.start !== undefined && placement.start === placement.end
      || previous?.start !== undefined && previous.start === previous.end) {
      result.push(placement);
      continue;
    }
    const tracking = placement.tracking ?? 0;
    const previousTracking = previous?.tracking ?? 0;
    if (previous?.sideways === true
      && placement.sideways === true
      && previous.line === placement.line
      && Math.abs(previous.x - placement.x) < 0.001
      && Math.abs(previous.y + previous.width - placement.y) < 0.001
      && previousTracking === tracking
      && sameTextRunStyle(previous.style, placement.style)) {
      const { start: previousStart, end: previousEnd, ...base } = previous;
      result[result.length - 1] = {
        ...base,
        text: previous.text + placement.text,
        width: previous.width + placement.width,
        ...(previousEnd !== undefined && placement.end !== undefined && previousEnd === placement.start
          ? { start: previousStart, end: placement.end }
          : {}),
      };
      continue;
    }
    const eastAsianAutoSpacingContinuation = previous !== undefined
      && CJK_SCRIPT.test(previous.text)
      && CJK_SCRIPT.test(placement.text)
      && graphemes(placement.text).length === 1
      && sameTextRunStyle(previous.style, placement.style, false)
      && Math.abs(
        placement.style.letterSpacing
          - previous.style.letterSpacing
          - placement.style.fontSize * EAST_ASIAN_AUTO_SPACING_EM,
      ) < 0.001;
    const contiguous = previous !== undefined && (direction === "rtl"
      ? Math.abs(placement.x + placement.width - previous.x) < 0.001
      : Math.abs(previous.x + previous.width - placement.x) < 0.001);
    if (previous !== undefined
      && previous.compression === undefined && placement.compression === undefined
      && contiguous
      && previous.line === placement.line
      && previous.y === placement.y
      && previousTracking === tracking
      && ((mergeWholeLine || previous.style.textEffect === placement.style.textEffect
          && (placement.style.textEffect?.glow !== undefined || placement.style.textEffect?.reflection !== undefined))
          && sameTextRunStyle(previous.style, placement.style)
        || previous.text.trim() !== ""
          && placement.text.trim() !== ""
          && ((previous.style === placement.style
            && (/\p{P}$/u.test(previous.text) || /^\p{P}/u.test(placement.text)))
            || ((CJK_SCRIPT.test(previous.text) && CJK_SCRIPT.test(placement.text)
              || COMPLEX_SHAPING_SCRIPT.test(previous.text + placement.text))
              && (sameTextRunStyle(previous.style, placement.style)
                || eastAsianAutoSpacingContinuation))))) {
      const x = Math.min(previous.x, placement.x);
      const right = Math.max(previous.x + previous.width, placement.x + placement.width);
      const { start: previousStart, end: previousEnd, ...base } = previous;
      result[result.length - 1] = {
        ...base,
        text: previous.text + placement.text,
        x,
        width: right - x,
        ...(previousEnd !== undefined && placement.end !== undefined && previousEnd === placement.start
          ? { start: previousStart, end: placement.end }
          : {}),
      };
    } else {
      result.push(placement);
    }
  }
  return result;
}

// Two open rails describe a text envelope, independent of the source format.
// ponytail: single line/cubic rails; add piecewise rails when a real document needs them.
function drawTextEnvelope(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  geometry: SceneGeometry,
  runs: readonly SceneTextRun[],
  textLayout: SceneTextLayout,
  resolveFontFamily?: (family: string, run: SceneTextRun) => string,
  paintStyles?: ReadonlyMap<ScenePaint, string | CanvasGradient | CanvasPattern>,
): boolean {
  if (typeof geometry !== "object" || geometry.kind !== "layered-path"
    || geometry.layers.length !== 2) return false;
  const rails = geometry.layers.map(layer => layer.commands);
  if (!rails.every(commands => commands.length === 2 && commands[0]?.kind === "moveTo"
    && (commands[1]?.kind === "lineTo" || commands[1]?.kind === "bezierCurveTo"))) return false;
  const point = (rail: number, t: number): { x: number; y: number } => {
    const start = rails[rail]?.[0];
    const end = rails[rail]?.[1];
    if (start?.kind !== "moveTo" || end === undefined || end.kind === "closePath") return { x: 0, y: 0 };
    const u = 1 - t;
    return end.kind === "bezierCurveTo" ? {
      x: u ** 3 * start.x + 3 * u * u * t * end.cp1x + 3 * u * t * t * end.cp2x + t ** 3 * end.x,
      y: u ** 3 * start.y + 3 * u * u * t * end.cp1y + 3 * u * t * t * end.cp2y + t ** 3 * end.y,
    } : { x: u * start.x + t * end.x, y: u * start.y + t * end.y };
  };
  const layout = layoutTextRuns(context, runs, { maxWidth: object.bounds.width,
    wrap: false, align: "start",
    ...(resolveFontFamily === undefined ? {} : { resolveFontFamily }) });
  const lines = Array.from({ length: layout.lineCount }, () => ({ left: Infinity, right: -Infinity }));
  let top = Infinity;
  let bottom = -Infinity;
  for (const placement of layout.runs) {
    applyRunFont(context, placement.style, 1, resolveFontFamily);
    context.textBaseline = "alphabetic";
    const metrics = context.measureText(placement.text);
    const horizontalScale = runHorizontalScale(placement.style);
    const line = lines[placement.line];
    if (line !== undefined) {
      line.left = Math.min(line.left, placement.x - metrics.actualBoundingBoxLeft * horizontalScale);
      line.right = Math.max(line.right, placement.x + metrics.actualBoundingBoxRight * horizontalScale);
    }
    top = Math.min(top, placement.y - metrics.actualBoundingBoxAscent);
    bottom = Math.max(bottom, placement.y + metrics.actualBoundingBoxDescent);
  }
  const naturalWidth = lines.reduce((width, line) => Math.max(width, line.right - line.left), 0);
  if (!(naturalWidth > 0) || !Number.isFinite(naturalWidth) || !(bottom > top) || !Number.isFinite(bottom - top)) return false;
  const transform = context.getTransform();
  const width = Math.min(2048, Math.max(256, Math.ceil(object.bounds.width * Math.hypot(transform.a, transform.b))));
  const height = Math.min(2048, Math.max(64, Math.ceil(width * (bottom - top) / naturalWidth)));
  const canvas = new OffscreenCanvas(width, height);
  const target = canvas.getContext("2d");
  if (target === null) return false;
  for (const placement of layout.runs) {
    const line = lines[placement.line];
    if (line === undefined || !(line.right > line.left)) continue;
    const lineWidth = line.right - line.left;
    const scaleX = width / (textLayout.autoFit === "fit-frame" ? lineWidth : naturalWidth);
    const offsetX = textLayout.autoFit === "fit-frame" ? 0 : (width - lineWidth * scaleX) / 2;
    target.save();
    target.translate(offsetX, 0);
    target.scale(scaleX, height / (bottom - top));
    applyRunFont(target, placement.style, 1, resolveFontFamily);
    target.textBaseline = "alphabetic";
    const glyphPaint = placement.style.paint ?? textLayout.textPaint;
    target.fillStyle = glyphPaint === undefined
      ? color(placement.style.color) : paintStyles?.get(glyphPaint) ?? paintStyle(target, object, glyphPaint) ?? color(placement.style.color);
    target.translate(placement.x - line.left, placement.y - top);
    target.scale(runHorizontalScale(placement.style), 1);
    if (textLayout.textFill ?? true) target.fillText(placement.text, 0, 0);
    const stroke = textStrokeStyle(target, object, textLayout);
    if (stroke !== undefined) {
      target.strokeStyle = stroke;
      target.lineWidth = textLayout.textStrokeWidth ?? 0;
      target.strokeText(placement.text, 0, 0);
    }
    target.restore();
  }
  // Rasterize authored text once, then map bounded strips between the actual rails.
  for (let column = 0; column < width; column += 1) {
    const a = point(0, column / width);
    const b = point(0, (column + 1) / width);
    const c = point(1, column / width);
    context.save();
    context.transform(b.x - a.x, b.y - a.y, c.x - a.x, c.y - a.y,
      object.bounds.x + a.x, object.bounds.y + a.y);
    context.drawImage(canvas, column, 0, 1, height, 0, 0, 1.01, 1);
    context.restore();
  }
  return true;
}

function drawText(
  context: OffscreenCanvasRenderingContext2D,
  object: SceneObject,
  visual: Extract<SceneVisual, { kind: "text" | "rich-text" }>,
  textLayout?: SceneTextLayout,
  resolveFontFamily?: (family: string, run: SceneTextRun) => string,
  prepareRun?: (run: SceneTextRun) => SceneTextRun,
  expandRuns?: (runs: readonly SceneTextRun[]) => readonly SceneTextRun[],
  layoutCache?: Map<string, TextLayout>,
  layoutCacheKey?: string,
  groupedTextScaleX = 1,
  groupedTextScaleY = groupedTextScaleX,
  strokeStyle?: SceneStrokeStyle,
  textEffects?: readonly SceneTextEffect[],
  collectTextFragment?: TextFragmentCollector,
  paintStyles?: ReadonlyMap<ScenePaint, string | CanvasGradient | CanvasPattern>,
): void {
  let sourceOffset = 0;
  const sourceRuns = visualTextRuns(object, visual).map((run, index) => {
    const sourceStart = run.sourceStart ?? sourceOffset;
    const sourceEnd = run.sourceEnd ?? sourceStart + run.text.length;
    sourceOffset = sourceEnd;
    const effect = textEffects?.[index];
    const indexedRun = { ...run, sourceStart, sourceEnd };
    return effect === undefined ? indexedRun : {
      ...indexedRun,
      ...(effect.shadow === undefined ? {} : { shadow: effect.shadow }),
      ...(effect.innerShadow === undefined ? {} : { innerShadow: effect.innerShadow }),
      textEffect: effect,
    };
  });
  const expandedRuns = expandRuns === undefined ? sourceRuns : expandRuns(sourceRuns);
  const preparedRuns = prepareRun === undefined ? expandedRuns : expandedRuns.map(prepareRun);
  if (textLayout?.warp === "text-envelope" && visual.kind === "rich-text") {
    context.save();
    try {
      if (drawTextEnvelope(context, object, visual.geometry, preparedRuns, textLayout, resolveFontFamily, paintStyles)) return;
    } finally { context.restore(); }
  }
  const textScaleX = Number.isFinite(groupedTextScaleX) && groupedTextScaleX > 0
    ? 1 / groupedTextScaleX
    : 1;
  const textScaleY = Number.isFinite(groupedTextScaleY) && groupedTextScaleY > 0
    ? 1 / groupedTextScaleY
    : 1;
  const runs = preparedRuns;
  if (runs.length === 0) return;
  const x = object.bounds.x / textScaleX;
  const y = object.bounds.y / textScaleY;
  const boundsWidth = object.bounds.width / textScaleX;
  const boundsHeight = object.bounds.height / textScaleY;
  const orientation = textLayout?.orientation ?? "horizontal";
  const rotated = orientation === "rotated-90" || orientation === "rotated-270";
  const width = rotated ? boundsHeight : boundsWidth;
  const height = rotated ? boundsWidth : boundsHeight;
  const insetLeft = Math.min(textLayout?.insetLeft ?? 4, width / 2);
  const insetRight = Math.min(textLayout?.insetRight ?? 4, width / 2);
  const insetTop = textLayout?.insetTop ?? Math.min(4, height / 2);
  const insetBottom = textLayout?.insetBottom ?? Math.min(4, height / 2);
  const marginLeft = Math.max(0, textLayout?.marginLeft ?? 0);
  const marginRight = Math.max(0, textLayout?.marginRight ?? 0);
  const availableWidth = Math.max(1, width - insetLeft - insetRight - marginLeft - marginRight);
  const availableHeight = Math.max(1, height - insetTop - insetBottom);
  const columns = Math.max(1, textLayout?.columnCount ?? 1);
  const columnSpacing = Math.max(0, textLayout?.columnSpacing ?? 0);
  const columnWidth = Math.max(1, (availableWidth - columnSpacing * (columns - 1)) / columns);
  const rotationDegrees = Number.isFinite(textLayout?.rotationDegrees)
    ? textLayout?.rotationDegrees ?? 0
    : 0;
  const horizontalOverflow = textLayout?.horizontalOverflow ?? "overflow";
  const verticalOverflow = textLayout?.verticalOverflow ?? "overflow";
  const originX = (rotated ? 0 : x) + insetLeft + marginLeft;
  // Overconstrained insets collapse the text rectangle at its midpoint.
  const authoredHeight = height - insetTop - insetBottom;
  const originY = (rotated ? 0 : y) + insetTop + Math.min(0, authoredHeight) / 2
    + verticalTextOffset(textLayout?.verticalAlign, Math.max(0, authoredHeight) - availableHeight);
  context.save();
  context.scale(textScaleX, textScaleY);
  applyStrokeStyle(context, strokeStyle);
  if (horizontalOverflow === "clip" || verticalOverflow !== "overflow") {
    context.beginPath();
    context.rect(x, y, boundsWidth, boundsHeight);
    context.clip();
  }
  if (rotationDegrees !== 0) {
    const centerX = x + boundsWidth / 2;
    const centerY = y + boundsHeight / 2;
    context.translate(centerX, centerY);
    context.rotate(rotationDegrees * Math.PI / 180);
    context.translate(-centerX, -centerY);
  }
  if (orientation === "rotated-90") {
    context.translate(x + boundsWidth, y);
    context.rotate(Math.PI / 2);
  } else if (orientation === "rotated-270") {
    context.translate(x, y + boundsHeight);
    context.rotate(-Math.PI / 2);
  }
  // Layout coordinates are physical left edges, including for RTL paragraphs.
  context.textAlign = "left";
  const direction = resolvedTextDirection(runs, textLayout?.direction);
  context.direction = direction;
  const directRun = positionedPdfRun(object, visual, textLayout, runs);
  if (directRun !== undefined && textLayout !== undefined) {
    context.restore();
    drawPositionedPdfRun(
      context,
      object,
      directRun,
      textLayout,
      resolveFontFamily?.(directRun.fontFamily, directRun) ?? directRun.fontFamily,
      strokeStyle,
      false,
      collectTextFragment,
    );
    return;
  }
  let layout = layoutCacheKey === undefined ? undefined : layoutCache?.get(layoutCacheKey);
  if (layout !== undefined && layoutCache !== undefined && layoutCacheKey !== undefined) {
    layoutCache.delete(layoutCacheKey);
    layoutCache.set(layoutCacheKey, layout);
  }
  const paragraphs = textLayout?.paragraphs;
  layout ??= layoutTextRuns(context, runs, {
    ...(textLayout?.fillCharacter === undefined ? {} : { fillCharacter: textLayout.fillCharacter }),
    // An authored baseline already includes the line's leading.
    fixedLineHeight: textLayout?.fixedLineHeight === true || (textLayout?.textBaseline ?? 0) > 0,
    compressPunctuation: textLayout?.compressPunctuation ?? false,
    continuesAfter: textLayout?.continuesAfter ?? false,
    wrapRegions: textLayout?.wrapRegions ?? [],
    maxWidth: columnWidth,
    maxHeight: availableHeight,
    lineHeight: visual.kind === "rich-text" ? visual.lineHeight : visual.fontSize * 1.2,
    align: visual.align,
    direction,
    orientation: rotated ? "horizontal" : orientation,
    verticalAlign: textLayout?.verticalAlign ?? "top",
    autoFit: (textLayout?.fontScale ?? 1) < 1 ? "none" : textLayout?.autoFit ?? "none",
    minScale: textLayout?.minScale ?? 0.1,
    fontScale: textLayout?.fontScale ?? 1,
    lineSpacingReduction: textLayout?.lineSpacingReduction ?? 0,
    columns,
    columnSpacing,
    wrap: textLayout?.wrap ?? true,
    ...(textLayout?.prefix === undefined ? {} : { prefix: textLayout.prefix }),
    tabStops: textLayout?.tabStops ?? [],
    defaultTabStop: textLayout?.defaultTabStop ?? 36,
    hangingIndent: textLayout?.hangingIndent ?? 0,
    paragraphSpacing: textLayout?.paragraphSpacing ?? 0,
    firstLineIndent: textLayout?.firstLineIndent ?? 0,
    paragraphs: paragraphs ?? [],
    ...(resolveFontFamily === undefined ? {} : { resolveFontFamily }),
  });
  if (layoutCache !== undefined && layoutCacheKey !== undefined && !layoutCache.has(layoutCacheKey)) {
    layoutCache.set(layoutCacheKey, layout);
    while (layoutCache.size > 4_096) {
      const oldest = layoutCache.keys().next().value;
      if (oldest === undefined) break;
      layoutCache.delete(oldest);
    }
  }
  for (const rule of layout.rules) {
    context.save();
    context.strokeStyle = color(rule.color);
    context.lineWidth = rule.strokeWidth;
    context.beginPath();
    context.moveTo(originX + rule.x, originY + rule.y);
    context.lineTo(originX + rule.x + rule.width, originY + rule.y);
    context.stroke();
    context.restore();
  }
  if (textLayout?.autoFit === "fit-frame" && layout.width > 0) {
    context.translate(originX, 0);
    context.scale(availableWidth / layout.width, 1);
    context.translate(-originX, 0);
  }
  const visibleRuns = verticalOverflow === "overflow"
    ? layout.runs
    : layout.runs.filter((placement) => (
        (placement.lineTop ?? placement.y) + (placement.height ?? layout.lineHeight) <= availableHeight + 0.001
      ));
  const displayedRuns = verticalOverflow === "ellipsis"
    && visibleRuns.length < layout.runs.length
    && visibleRuns.length > 0
    ? visibleRuns.map((placement, index) => (
        index === visibleRuns.length - 1
          ? { ...placement, text: `${placement.text.replace(/…$/u, "")}…` }
          : placement
      ))
    : visibleRuns;
  const paintRuns = coalescePaintTextRuns(
    displayedRuns,
    direction,
    textLayout?.textScaleToFit ?? false,
  );
  const wordParagraph = object.source.format === "docx" && object.type === "paragraph"
    && (textLayout?.paragraphs?.length ?? 0) === 0;
  const fontBaselines = new Map<string, {
    readonly offset: number | null;
    readonly descent: number;
    readonly strikeAscent: number;
    readonly strikeDescent: number;
    readonly fontBoxHeight: number | null;
  }>();
  const measuredRuns = paintRuns.map((placement) => {
    applyRunFont(context, placement.style, layout.scale, resolveFontFamily);
    const fontSize = placement.style.fontSize * layout.scale;
    const naturalLineHeight = fontSize * 1.2;
    const lineBoxHeight = placement.height ?? layout.lineHeight;
    // Pages' sub-single line spacing changes baseline advance, not font ascent.
    const lineCompression = object.type === "paragraph" || object.source.format === "pages"
      ? 1
      : Math.min(1, Math.max(0, lineBoxHeight / naturalLineHeight));
    const extraLeading = (textLayout?.paragraphs?.length ?? 0) > 0
      ? Math.max(0, lineBoxHeight - naturalLineHeight)
      : 0;
    const baselineKey = `${context.font}\0${lineCompression}\0${extraLeading}\0${lineBoxHeight}`;
    let baseline = fontBaselines.get(baselineKey);
    if (baseline === undefined) {
      context.textBaseline = "alphabetic";
      const alphabeticMetrics = context.measureText("Mg");
      context.textBaseline = "top";
      const topMetrics = context.measureText("Mg");
      const ascent = alphabeticMetrics.fontBoundingBoxAscent;
      const descent = alphabeticMetrics.fontBoundingBoxDescent;
      const topAscent = topMetrics.fontBoundingBoxAscent;
      if (Number.isFinite(ascent) && ascent > 0
        && Number.isFinite(descent) && descent >= 0
        && Number.isFinite(topAscent)) {
        // DrawingML positions glyphs from the font box. Canvas' `top` baseline is
        // the em-box top instead, which is visibly high for fonts such as Segoe UI.
        // Condensed line boxes compress the ascent; authored extra leading was
        // already placed before the glyph by layoutTextRuns, so consume it first.
        // Top-relative font ascent can be negative for CJK faces; it is a signed distance.
        const canvasTopBaseline = Math.max(0, ascent - topAscent);
        const drawingMlBaseline = ascent * lineCompression - extraLeading;
        const actualAscent = alphabeticMetrics.actualBoundingBoxAscent;
        const actualDescent = alphabeticMetrics.actualBoundingBoxDescent;
        baseline = {
          // Word puts natural line leading above the glyph box. DrawingML's
          // authored paragraph leading is already applied by layoutTextRuns.
          offset: Math.max(
            canvasTopBaseline,
            drawingMlBaseline + (wordParagraph
              ? Math.max(0, lineBoxHeight - ascent - descent)
              : 0),
          ),
          descent,
          strikeAscent: Number.isFinite(actualAscent) && actualAscent > 0
            ? actualAscent
            : ascent * 0.7,
          strikeDescent: Number.isFinite(actualDescent) && actualDescent >= 0
            ? actualDescent
            : 0,
          fontBoxHeight: ascent + descent,
        };
      } else {
        baseline = {
          offset: null,
          descent: 0,
          strikeAscent: 0,
          strikeDescent: 0,
          fontBoxHeight: null,
        };
      }
      fontBaselines.set(baselineKey, baseline);
    }
    context.textBaseline = "alphabetic";
    const ink = context.measureText(placement.text);
    return {
      placement,
      inkAscent: ink.actualBoundingBoxAscent,
      inkDescent: ink.actualBoundingBoxDescent,
      font: context.font,
      letterSpacing: `${
        placement.style.letterSpacing * layout.scale + (placement.tracking ?? 0)
      }px`,
      baseline,
    };
  });
  // A typographic line owns one baseline even when its runs use different fonts.
  const lineBaselineOffsets = new Map<number, number>();
  const fixedLineMetrics = new Map<number, { ascent: number; descent: number; height: number }>();
  for (const { placement, baseline } of measuredRuns) {
    if (placement.baselineOffset !== undefined) continue;
    if (baseline.offset === null) continue;
    if (textLayout?.fixedLineHeight && baseline.fontBoxHeight !== null) {
      const previous = fixedLineMetrics.get(placement.line);
      fixedLineMetrics.set(placement.line, {
        ascent: Math.max(previous?.ascent ?? 0, baseline.fontBoxHeight - baseline.descent),
        descent: Math.max(previous?.descent ?? 0, baseline.descent),
        height: placement.height ?? layout.lineHeight,
      });
    }
    lineBaselineOffsets.set(
      placement.line,
      Math.max(lineBaselineOffsets.get(placement.line) ?? 0, baseline.offset),
    );
  }
  for (const [line, { ascent, descent, height }] of fixedLineMetrics) {
    const remaining = height - ascent - descent;
    lineBaselineOffsets.set(line, ascent + (remaining >= 0 ? remaining / 2 : remaining));
  }
  const lineInkBounds = new Map<number, { top: number; bottom: number }>();
  for (const { placement, inkAscent, inkDescent } of measuredRuns) {
    if (!Number.isFinite(inkAscent) || !Number.isFinite(inkDescent) || inkAscent + inkDescent <= 0) continue;
    const baseline = originY + placement.y - placement.style.baselineShift * layout.scale
      + (placement.baselineOffset ?? lineBaselineOffsets.get(placement.line) ?? 0);
    const previous = lineInkBounds.get(placement.line);
    lineInkBounds.set(placement.line, {
      top: Math.min(previous?.top ?? Infinity, baseline - inkAscent),
      bottom: Math.max(previous?.bottom ?? -Infinity, baseline + inkDescent),
    });
  }
  const lineUnderlineOffsets = new Map<string, number>();
  for (const { placement, baseline } of measuredRuns) {
    if (!placement.style.underline || baseline.offset === null) continue;
    const key = `${
      placement.line
    }\0${placement.style.fontSize}\0${placement.style.baselineShift}\0${placement.style.color}`;
    const offset = Math.max(
      1,
      placement.style.fontSize * layout.scale / 14,
      baseline.descent / 2,
    );
    lineUnderlineOffsets.set(key, Math.max(lineUnderlineOffsets.get(key) ?? 0, offset));
  }
  for (const [index, measured] of measuredRuns.entries()) {
    const { placement, font, letterSpacing, baseline } = measured;
    const lineY = originY + placement.y;
    const runFont = typeof font === "string"
      ? font
      : fontShorthand(
        resolveFontFamily?.(placement.style.fontFamily, placement.style)
          ?? placement.style.fontFamily,
        placement.style.fontSize * layout.scale,
        placement.style.italic,
        placement.style.bold,
      );
    context.font = runFont;
    context.letterSpacing = letterSpacing;
    const scaleToFit = textLayout?.textScaleToFit ?? false;
    const physicalRunX = originX + (scaleToFit ? 0 : placement.x);
    const sideways = placement.sideways === true;
    if (sideways) {
      context.save();
      context.translate(physicalRunX + (placement.height ?? layout.lineHeight), lineY);
      context.rotate(Math.PI / 2);
    }
    const punctuationOffset = OPENING_PUNCTUATION.test(placement.text)
      ? -(placement.compression ?? 0)
      : /^[：；！？]$/u.test(placement.text) ? -(placement.compression ?? 0) / 2 : 0;
    const selectionX = sideways ? 0 : physicalRunX;
    const runX = selectionX + (sideways ? 0 : punctuationOffset);
    const textTopY = (sideways ? 0 : lineY) - placement.style.baselineShift * layout.scale;
    const fontSize = placement.style.fontSize * layout.scale;
    const authoredBaseline = textLayout?.textBaseline ?? 0;
    const baselineOffset = placement.baselineOffset ?? (authoredBaseline > 0
      ? authoredBaseline
      : lineBaselineOffsets.get(placement.line) ?? null);
    const textY = baselineOffset === null ? textTopY : textTopY + baselineOffset;
    const runTextBaseline = baselineOffset === null ? "top" : "alphabetic";
    context.textBaseline = runTextBaseline;
    const lineBoxHeight = placement.height ?? layout.lineHeight;
    const selectionHeight = object.source.format === "docx"
      && baseline.fontBoxHeight !== null
      ? Math.min(lineBoxHeight, baseline.fontBoxHeight)
      : lineBoxHeight;
    const selectionY = wordParagraph && baseline.fontBoxHeight !== null
      ? textY - baseline.fontBoxHeight + baseline.descent
      : textTopY;
    if (placement.width > 0 && (placement.style.highlight & 0xff) !== 0) {
      context.fillStyle = color(placement.style.highlight);
      context.fillRect(selectionX, selectionY, placement.width, selectionHeight);
    }
    const currentTransform = collectTextFragment === undefined || textLayout?.warp !== undefined
      ? undefined
      : canvasTransform(context);
    if (currentTransform !== undefined && collectTextFragment !== undefined && placement.text.length > 0) {
      collectTextFragment({
        objectId: object.id,
        text: placement.text,
        ...(placement.start === undefined || placement.end === undefined ? {} : {
          start: placement.start,
          end: placement.end,
        }),
        width: placement.width,
        height: selectionHeight,
        line: placement.line,
        font: runFont,
        letterSpacing: placement.style.letterSpacing * layout.scale + (placement.tracking ?? 0),
        direction,
      }, multiplyTransform(currentTransform, {
        a: 1,
        b: 0,
        c: 0,
        d: 1,
        // Compressed punctuation shifts its ink, not its logical selection advance.
        e: selectionX,
        f: selectionY,
      }));
    }
    const glyphPaint = placement.style.paint ?? textLayout?.textPaint;
    const effect = placement.style.textEffect;
    const metrics = context.measureText(placement.text);
    const glyphBounds = {
      x: runX - (metrics.actualBoundingBoxLeft || 0),
      y: textY - (metrics.actualBoundingBoxAscent || 0),
      width: Math.max(1, (metrics.actualBoundingBoxLeft || 0) + (metrics.actualBoundingBoxRight || placement.width)),
      height: Math.max(1, (metrics.actualBoundingBoxAscent || 0) + (metrics.actualBoundingBoxDescent || 0)),
    };
    // DrawingML run gradients span the glyphs on each laid-out line, whereas
    // page-space paints (e.g. PDF) keep their authored coordinates.
    const glyphStyle = (paint: ScenePaint): string | CanvasGradient | CanvasPattern | undefined => {
      if (paint.kind === "linear-gradient") {
        const scaleX = glyphBounds.width / Math.max(1, object.bounds.width);
        const scaleY = glyphBounds.height / Math.max(1, object.bounds.height);
        return paintStyle(context, { ...object, bounds: glyphBounds }, {
          ...paint,
          start: { x: paint.start.x * scaleX, y: paint.start.y * scaleY },
          end: { x: paint.end.x * scaleX, y: paint.end.y * scaleY },
        });
      }
      return paintStyles?.get(paint) ?? paintStyle(context, object, paint);
    };
    const runFillStyle = glyphPaint === undefined ? color(placement.style.color)
      : glyphPaint.kind === "none" ? "rgba(0,0,0,0)"
        : (effect?.fillToText ? glyphStyle(glyphPaint) : paintStyles?.get(glyphPaint) ?? paintStyle(context, object, glyphPaint))
          ?? color(placement.style.color);
    context.fillStyle = runFillStyle;
    const textStroke = effect?.stroke === undefined ? textStrokeStyle(context, object, textLayout)
      : effect.stroke.kind === "none" ? undefined : glyphStyle(effect.stroke);
    const textStrokeWidth = effect?.strokeWidth ?? textLayout?.textStrokeWidth ?? 0;
    const fitTextPath = (
      target: OffscreenCanvasRenderingContext2D,
      text: string,
      glyphX: number,
      glyphY: number,
    ): boolean => {
      if (textLayout?.warp !== "text-path-fit") return false;
      const metrics = target.measureText(text);
      const inkWidth = metrics.actualBoundingBoxLeft + metrics.actualBoundingBoxRight;
      const inkHeight = metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent;
      if (!(inkWidth > 0) || !(inkHeight > 0)) return false;
      target.translate(originX, originY);
      target.scale(availableWidth / inkWidth, availableHeight / inkHeight);
      target.translate(
        -glyphX + metrics.actualBoundingBoxLeft,
        -glyphY + metrics.actualBoundingBoxAscent,
      );
      return true;
    };
    const paintGlyphs = (
      target: OffscreenCanvasRenderingContext2D,
      text: string,
      glyphX: number,
      glyphY: number,
    ): void => {
      if (textLayout?.warp === "text-path-fit") {
        target.save();
        const fitted = fitTextPath(target, text, glyphX, glyphY);
        if (fitted) {
          if (textLayout.textFill ?? true) target.fillText(text, glyphX, glyphY);
          if (textStroke !== undefined) {
            target.strokeStyle = textStroke;
            target.lineWidth = effectiveStrokeWidth(
              target,
              object,
              textStrokeWidth,
            );
            target.strokeText(text, glyphX, glyphY);
          }
        }
        target.restore();
        if (fitted) return;
      }
      const scaleX = (textLayout?.textScaleToFit ?? false) && placement.width > 0
        ? runHorizontalScale(placement.style) * availableWidth / placement.width
        : runHorizontalScale(placement.style);
      if (Math.abs(scaleX - 1) < 0.000_001) {
        if (textLayout?.textFill ?? true) target.fillText(text, glyphX, glyphY);
        if (textStroke !== undefined) {
          target.strokeStyle = textStroke;
          target.lineWidth = effectiveStrokeWidth(
            target,
            object,
            textStrokeWidth,
          );
          target.strokeText(text, glyphX, glyphY);
        }
        return;
      }
      target.save();
      target.translate(glyphX, glyphY);
      target.scale(scaleX, 1);
      if (textLayout?.textFill ?? true) target.fillText(text, 0, 0);
      if (textStroke !== undefined) {
        target.strokeStyle = textStroke;
        target.lineWidth = effectiveStrokeWidth(
          target,
          object,
          textStrokeWidth,
        );
        target.strokeText(text, 0, 0);
      }
      target.restore();
    };
    const paintShadowGlyphs = (
      target: OffscreenCanvasRenderingContext2D,
      text: string,
      glyphX: number,
      glyphY: number,
      shadow: NonNullable<SceneTextRun["shadow"]>,
    ): void => {
      const effect = placement.style.textEffect;
      if (effect === undefined) return;
      const horizontal = effect.shadowAlignment % 3;
      const vertical = Math.floor(effect.shadowAlignment / 3);
      const anchorX = scaleToFit
        ? originX + availableWidth * horizontal / 2
        : runX + placement.width * horizontal / 2;
      const anchorY = scaleToFit
        ? originY + availableHeight * vertical / 2
        : textTopY + (placement.height ?? layout.lineHeight) * vertical / 2;
      const scaleX = effect.shadowScaleX;
      const scaleY = effect.shadowScaleY;
      const skewX = Math.tan(effect.shadowSkewX * Math.PI / 180);
      const skewY = Math.tan(effect.shadowSkewY * Math.PI / 180);
      target.save();
      try {
        target.translate(
          anchorX + shadow.offsetX * layout.scale,
          anchorY + shadow.offsetY * layout.scale,
        );
        target.transform(scaleX, skewY, skewX, scaleY, 0, 0);
        target.translate(-anchorX, -anchorY);
        // DrawingML's blur radius spans roughly two CSS blur standard deviations.
        target.filter = `blur(${shadow.blur * layout.scale / 2}px)`;
        target.fillStyle = color(shadow.color);
        if (fitTextPath(target, text, glyphX, glyphY)) {
          target.fillText(text, glyphX, glyphY);
          return;
        }
        const textScaleX = (textLayout?.textScaleToFit ?? false) && placement.width > 0
          ? runHorizontalScale(placement.style) * availableWidth / placement.width
          : runHorizontalScale(placement.style);
        if (Math.abs(textScaleX - 1) < 0.000_001) target.fillText(text, glyphX, glyphY);
        else {
          target.translate(glyphX, glyphY);
          target.scale(textScaleX, 1);
          target.fillText(text, 0, 0);
        }
      } finally {
        target.restore();
      }
    };
    const paintRunGlyphs = (text: string, glyphX: number, glyphY: number): void => {
      const paint = (target: OffscreenCanvasRenderingContext2D): void => {
        target.font = runFont;
        target.letterSpacing = letterSpacing;
        target.textBaseline = runTextBaseline;
        target.fillStyle = runFillStyle;
        if (placement.style.shadow !== undefined) {
          paintShadowGlyphs(target, text, glyphX, glyphY, placement.style.shadow);
        }
        paintGlyphs(target, text, glyphX, glyphY);
        // ponytail: Canvas 2D cannot isolate this from the painted slide; keep
        // inner shadows parsed until text runs render through a scratch layer.
      };
      const reflection = effect?.reflection;
      const glow = effect?.glow;
      if (reflection === undefined && glow === undefined) {
        paint(context);
        return;
      }
      const lineInk = sideways ? undefined : lineInkBounds.get(placement.line);
      drawWithAdvancedEffects(
        context,
        { ...object, bounds: reflection === undefined || lineInk === undefined
          ? glyphBounds : { ...glyphBounds, y: lineInk.top, height: lineInk.bottom - lineInk.top } },
        [{ kind: "advanced-effect", ...(reflection === undefined ? {} : { reflection }),
          ...(glow === undefined ? {} : { glow }), visual: { kind: "none" } }],
        paint,
      );
    };
    if (placement.text === "\t") {
      const leader = placement.leader === "dot" ? "."
        : placement.leader === "hyphen" ? "-"
          : placement.leader === "underscore" ? "_"
            : placement.leader === "middle-dot" ? "·" : undefined;
      if (leader !== undefined && placement.width > 0) {
        const leaderWidth = Math.max(1, measureRunText(context, leader, placement.style));
        for (let leaderX = runX + leaderWidth / 2;
          leaderX + leaderWidth / 2 <= runX + placement.width;
          leaderX += leaderWidth) {
          paintRunGlyphs(leader, leaderX, textY);
        }
      }
    } else if (textLayout?.warp !== undefined && textLayout.warp !== "text-path-fit") {
      const phase = paintRuns.length <= 1 ? 0 : index / (paintRuns.length - 1);
      context.save();
      context.translate(runX, textY);
      context.rotate(Math.sin(phase * Math.PI * 2) * Math.PI / 36);
      context.translate(-runX, -textY);
      paintRunGlyphs(placement.text, runX, textY);
      context.restore();
    } else {
      paintRunGlyphs(placement.text, runX, textY);
    }
    if (placement.width > 0 && (placement.style.underline || placement.style.strikethrough)) {
      context.strokeStyle = color(placement.style.color);
      context.lineWidth = Math.max(1, placement.style.fontSize * layout.scale / 48);
      const decoration = (decorationY: number): void => {
        context.beginPath();
        context.moveTo(runX, decorationY);
        context.lineTo(runX + placement.width, decorationY);
        context.stroke();
      };
      const wavyDecoration = (decorationY: number): void => {
        const amplitude = context.lineWidth;
        const halfWave = amplitude * 2;
        const end = runX + placement.width;
        let x = runX;
        let direction = -1;
        context.beginPath();
        context.moveTo(x, decorationY);
        while (x < end) {
          const width = Math.min(halfWave, end - x);
          context.bezierCurveTo(
            x + width / 3,
            decorationY + direction * amplitude,
            x + width * 2 / 3,
            decorationY + direction * amplitude,
            x + width,
            decorationY,
          );
          x += width;
          direction *= -1;
        }
        context.stroke();
      };
      // Canvas TextMetrics are relative to the alphabetic baseline: place the
      // underline in the descent and the strike through the visible glyph body.
      // The legacy top-baseline heuristic remains only when metrics are absent.
      if (placement.style.underline) {
        context.save();
        try {
          if (placement.style.textEffect?.heavyUnderline === true) context.lineWidth *= 2;
          if (placement.style.textEffect?.dotDashUnderline === true) {
            const width = context.lineWidth;
            context.setLineDash([8 * width, 4 * width, 2 * width, 4 * width]);
          } else if (placement.style.textEffect?.dottedUnderline === true) {
            context.lineCap = "round";
            context.setLineDash([0, context.lineWidth * 2]);
          }
          const underlineKey = `${
            placement.line
          }\0${placement.style.fontSize}\0${placement.style.baselineShift}\0${placement.style.color}`;
          const underlineY = baselineOffset === null
            ? textY + fontSize * 1.05
            : textY + (lineUnderlineOffsets.get(underlineKey)
              ?? Math.max(context.lineWidth, baseline.descent / 2));
          const paintUnderline = placement.style.textEffect?.wavyUnderline === true
            ? wavyDecoration
            : decoration;
          paintUnderline(underlineY);
          if (placement.style.textEffect?.doubleUnderline === true) {
            paintUnderline(underlineY + Math.max(1.5, context.lineWidth * 2));
          }
        } finally {
          context.restore();
        }
      }
      if (placement.style.strikethrough) {
        const centerY = baselineOffset === null
          ? textY + fontSize * 0.55
          : textY + (baseline.strikeDescent - baseline.strikeAscent) / 2;
        if (placement.style.textEffect?.doubleStrikethrough === true) {
          const halfGap = Math.max(0.75, context.lineWidth);
          decoration(centerY - halfGap);
          decoration(centerY + halfGap);
        } else {
          decoration(centerY);
        }
      }
    }
    if (sideways) context.restore();
  }
  context.restore();
}

interface RenderControl {
  readonly isCancelled: () => boolean;
  readonly yieldEvery?: number;
}

async function yieldRenderTurn(): Promise<void> {
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
}

export class SceneRenderer {
  readonly #objectsByUnit = new Map<number, readonly SceneObject[]>();
  readonly #preparedUnits = new Map<number, PreparedSceneUnit>();
  readonly #limits: ResourceLimits;
  #imageFonts: readonly OfficeImageCodecFont[];
  #fonts: RuntimeFontSet | RuntimeFontResolver | undefined;
  readonly #images = new Map<string, CachedSceneImage>();
  readonly #renderedFrames = new Map<string, CachedRenderedFrame>();
  readonly #imageFingerprints = new WeakMap<Uint8Array, string>();
  readonly #imageSources = new Map<string, Uint8Array[]>();
  readonly #fontAvailability = new Map<string, boolean>();
  readonly #fontResolutions = new Map<string, FontResolution>();
  readonly #textLayouts = new Map<string, TextLayout>();
  #clipPathsByGeometry = new WeakMap<
    object,
    { readonly x: number; readonly y: number; readonly path: Path2D }
  >();
  readonly #imagePrefetches = new Map<Promise<DecodedSceneImage>, number>();
  readonly #imagePrefetchWaiters = new Set<() => void>();
  readonly #imagePins = new Map<Promise<DecodedSceneImage>, number>();
  readonly #imageDecodeQueue: {
    readonly pixels: number;
    readonly resolve: (release: () => void) => void;
    readonly reject: (cause: unknown) => void;
  }[] = [];
  #imagePixels = 0;
  #retainedImagePixels = 0;
  #retainedFramePixels = 0;
  #imagePrefetchPixels = 0;
  #activeImageDecodes = 0;
  #activeImageDecodePixels = 0;
  #activeRenders = 0;
  readonly #retiredImages = new Set<ImageBitmap>();
  #closed = false;

  constructor(
    objects: readonly SceneObject[],
    limits: ResourceLimits,
    fonts?: RuntimeFontSet | RuntimeFontResolver,
    imageFonts: readonly OfficeImageCodecFont[] = [],
  ) {
    const objectsByUnit = new Map<number, SceneObject[]>();
    for (const object of objects) {
      const unitObjects = objectsByUnit.get(object.unitIndex);
      if (unitObjects === undefined) objectsByUnit.set(object.unitIndex, [object]);
      else unitObjects.push(object);
    }
    for (const [unitIndex, unitObjects] of objectsByUnit) {
      this.#objectsByUnit.set(unitIndex, unitObjects);
    }
    this.#limits = limits;
    this.#fonts = fonts;
    this.#imageFonts = imageFonts;
  }

  appendObjects(objects: readonly SceneObject[]): void {
    this.#assertOpen();
    const additions = new Map<number, SceneObject[]>();
    for (const object of objects) {
      const unitObjects = additions.get(object.unitIndex);
      if (unitObjects === undefined) additions.set(object.unitIndex, [object]);
      else unitObjects.push(object);
    }
    for (const [unitIndex, unitObjects] of additions) {
      const existing = this.#objectsByUnit.get(unitIndex) ?? [];
      this.#objectsByUnit.set(unitIndex, [...existing, ...unitObjects]);
      this.#preparedUnits.delete(unitIndex);
      this.#invalidateRenderedFrames(unitIndex);
    }
  }

  replaceObjects(objects: readonly SceneObject[], imageFonts: readonly OfficeImageCodecFont[]): void {
    this.#assertOpen();
    // A scene refresh changes geometry/text, not the document's decoded rasters.
    if (imageFonts.length !== this.#imageFonts.length
      || imageFonts.some((font, index) => font.family !== this.#imageFonts[index]?.family
        || font.bytes !== this.#imageFonts[index]?.bytes)) {
      for (const [key, entry] of this.#images) {
        if (!key.startsWith("image/svg+xml:")) continue;
        this.#images.delete(key);
        this.#retainedImagePixels -= entry.pixels ?? 0;
        this.#imagePixels -= entry.pixels ?? 0;
        void entry.promise.then(({ bitmap }) => bitmap.close(), () => undefined);
      }
    }
    this.#imageFonts = imageFonts;
    this.#objectsByUnit.clear();
    this.#preparedUnits.clear();
    this.#fontAvailability.clear();
    this.#fontResolutions.clear();
    this.#textLayouts.clear();
    this.#clipPathsByGeometry = new WeakMap();
    for (const frame of this.#renderedFrames.values()) frame.bitmap.close();
    this.#renderedFrames.clear();
    this.#retainedFramePixels = 0;
    this.appendObjects(objects);
  }

  async render(
    unit: UnitDescriptor,
    request: RenderRequest,
    control?: RenderControl,
  ): Promise<RenderResult> {
    this.#activeRenders += 1;
    try {
      return await this.#render(unit, request, control);
    } finally {
      this.#activeRenders -= 1;
      if (this.#activeRenders === 0) {
        for (const bitmap of this.#retiredImages) {
          this.#imagePixels = Math.max(0, this.#imagePixels - bitmap.width * bitmap.height);
          bitmap.close();
        }
        this.#retiredImages.clear();
      }
    }
  }

  async #render(
    unit: UnitDescriptor,
    request: RenderRequest,
    control?: RenderControl,
  ): Promise<RenderResult> {
    this.#assertOpen();
    const assertActive = (): void => {
      this.#assertOpen();
      if (control?.isCancelled() === true) {
        throw new OfficeEngineError("OPERATION_ABORTED", "Rendering was superseded or cancelled");
      }
    };
    assertActive();
    const mappers = sheetMappers(unit, request);
    const sheetAxes = unit.type === "sheet"
      ? mappers ?? {
        rows: new SheetAxisMapper(unit.rowAxis, unit.rows),
        columns: new SheetAxisMapper(unit.columnAxis, unit.columns),
      }
      : undefined;
    const scale = request.scale ?? 1;
    const pixelRatio = request.pixelRatio ?? 1;
    validateNumber(scale, "scale");
    validateNumber(pixelRatio, "pixelRatio");
    if (request.watermark !== undefined
      && (typeof request.watermark !== "string"
        || request.watermark.trim().length === 0
        || request.watermark.length > 256)) {
      throw new OfficeEngineError("INVALID_RENDER_REQUEST", "watermark must contain 1 through 256 characters");
    }
    if (request.includeTextFragments !== undefined
      && typeof request.includeTextFragments !== "boolean") {
      throw new OfficeEngineError("INVALID_RENDER_REQUEST", "includeTextFragments must be a boolean");
    }
    const watermark = request.watermark?.trim();
    const viewport = request.viewport ?? {
      x: 0,
      y: 0,
      width: mappers?.columns.total ?? unit.width,
      height: mappers?.rows.total ?? unit.height,
    };
    validateNumber(viewport.x, "viewport.x", true);
    validateNumber(viewport.y, "viewport.y", true);
    validateNumber(viewport.width, "viewport.width");
    validateNumber(viewport.height, "viewport.height");
    const pixelWidth = Math.ceil(viewport.width * scale * pixelRatio);
    const pixelHeight = Math.ceil(viewport.height * scale * pixelRatio);
    if (pixelWidth * pixelHeight > this.#limits.renderPixels) {
      throw new OfficeEngineError(
        "RENDER_PIXEL_LIMIT",
        `Requested render is ${pixelWidth * pixelHeight} pixels; limit is ${this.#limits.renderPixels}`,
      );
    }
    const unitObjects = this.#objectsByUnit.get(unit.index);
    const pdfPage = unitObjects?.some(({ source }) => source.format === "pdf") === true;
    const pdfUnit = pdfPage && unitObjects?.some(({ source, visual }) => (
      source.format === "pdf" && visualNeedsLowResolutionSupersampling(visual)
    )) === true;
    const requestedDeviceScale = scale * pixelRatio;
    // PDF page coordinates are converted from 72 pt/in to 96 CSS px/in in the
    // format pack. A device scale of 0.75 is therefore already a 1:1 PDF-pixel
    // render and must not be resampled a second time.
    const pdfPixelScale = 0.75;
    const pdfSupersample = pdfUnit && requestedDeviceScale < pdfPixelScale
      ? Math.min(2, pdfPixelScale / requestedDeviceScale)
      : 1;
    const supersampledWidth = Math.ceil(viewport.width * requestedDeviceScale * pdfSupersample);
    const supersampledHeight = Math.ceil(viewport.height * requestedDeviceScale * pdfSupersample);
    const supersample = supersampledWidth * supersampledHeight <= this.#limits.renderPixels
      ? pdfSupersample
      : 1;
    const deviceScale = scale * pixelRatio * supersample;
    const canvasWidth = Math.ceil(viewport.width * deviceScale);
    const canvasHeight = Math.ceil(viewport.height * deviceScale);
    const roundingDistortion = Math.max(
      canvasWidth / (viewport.width * deviceScale) - 1,
      canvasHeight / (viewport.height * deviceScale) - 1,
    );
    // PDF.js keeps one uniform viewport scale and leaves any fractional
    // remainder transparent at the right/bottom canvas edge. Stretching a PDF
    // page independently to each rounded canvas dimension changes image and
    // glyph sampling, especially on very small pages.
    const useExactCanvasScale = !pdfPage && roundingDistortion > 0.01;
    const canvasScaleX = useExactCanvasScale ? canvasWidth / viewport.width : deviceScale;
    const canvasScaleY = useExactCanvasScale ? canvasHeight / viewport.height : deviceScale;
    const renderedFrameKey = JSON.stringify([
      unit.index,
      viewport.x,
      viewport.y,
      viewport.width,
      viewport.height,
      scale,
      pixelRatio,
      request.background ?? "#ffffff",
      watermark ?? "",
      request.includeTextFragments === true,
      unit.type === "sheet" ? sheetSizeCacheKey(request) : "",
    ]);
    const cachedFrame = this.#renderedFrames.get(renderedFrameKey);
    if (cachedFrame !== undefined && typeof createImageBitmap === "function") {
      let bitmap: ImageBitmap | undefined;
      try {
        bitmap = await createImageBitmap(cachedFrame.bitmap);
        assertActive();
        this.#renderedFrames.delete(renderedFrameKey);
        this.#renderedFrames.set(renderedFrameKey, cachedFrame);
        return { ...cachedFrame.result, bitmap };
      } catch (cause) {
        bitmap?.close();
        if (cause instanceof OfficeEngineError) throw cause;
        this.#renderedFrames.delete(renderedFrameKey);
        this.#retainedFramePixels -= cachedFrame.pixels;
        cachedFrame.bitmap.close();
      }
    }
    if (typeof OffscreenCanvas === "undefined") {
      throw new OfficeEngineError("UNSUPPORTED_ENVIRONMENT", "OffscreenCanvas is required for rendering");
    }
    const canvas = new OffscreenCanvas(canvasWidth, canvasHeight);
    const context = canvas.getContext("2d", { alpha: true });
    if (context === null) {
      throw new OfficeEngineError("UNSUPPORTED_ENVIRONMENT", "A Canvas 2D context is required for rendering");
    }
    context.imageSmoothingEnabled = true;
    context.imageSmoothingQuality = "high";

    const diagnostics: Diagnostic[] = [];
    const reportedFonts = new Set<string>();
    const prepared = this.#prepareUnit(unit, request, mappers);
    const first = firstMaximumBottomAfter(prepared.maximumBottoms, viewport.y);
    const end = firstObjectStartingAtOrAfter(prepared.objects, viewport.y + viewport.height);
    const visibleObjects: { object: SceneObject; state: VisualState }[] = [];
    for (let index = first; index < end; index += 1) {
      const candidate = prepared.objects[index]!;
      if (!intersects(candidate.bounds, viewport)) continue;
      visibleObjects.push({
        object: candidate.object,
        state: candidate.state,
      });
    }
    visibleObjects.sort(({ object: left }, { object: right }) => comparePaintOrder(left, right));
    const media = visibleObjects.flatMap(({ object, state }) => state.media === undefined ? [] : [{
      objectId: object.id,
      kind: state.media.mediaKind,
      mediaType: state.media.mediaType,
      bytes: state.media.bytes.slice(),
      bounds: { ...object.bounds },
      transform: { ...state.transform },
    }]);
    const textFragments: RenderedTextFragment[] = [];
    const collectTextFragment: TextFragmentCollector | undefined = request.includeTextFragments === true
      ? (fragment, transform) => {
          textFragments.push({
            ...fragment,
            transform: multiplyTransform({
              a: 1 / canvasScaleX,
              b: 0,
              c: 0,
              d: 1 / canvasScaleY,
              e: viewport.x,
              f: viewport.y,
            }, transform),
          });
        }
      : undefined;
    const imagePrefetch = this.#prefetchImages(
      visibleObjects,
      () => control?.isCancelled() !== true,
      deviceScale,
    );
    let renderedObjectCount = 0;

    context.save();
    try {
      context.scale(canvasScaleX, canvasScaleY);
      context.translate(-viewport.x, -viewport.y);
      let gridPending = unit.type === "sheet";
      const drawPendingGrid = () => {
        if (gridPending && unit.type === "sheet" && sheetAxes !== undefined) {
          drawSheetGridLines(context, unit, viewport, sheetAxes.rows, sheetAxes.columns);
        }
        gridPending = false;
      };

      let processedObjects = 0;
      const sharedClipKeys: Path2D[] = [];
      const drawDirectPdfText = ({ object, state }: {
        readonly object: SceneObject;
        readonly state: VisualState;
      }): boolean => {
        const { visual } = state;
        if ((visual.kind !== "text" && visual.kind !== "rich-text")
          || state.operations.length !== 0
          || state.groupedTextScale !== 1
          || !hasInvisibleTextGeometry(visual)
          || state.textLayout === undefined) return false;
        const run = positionedPdfRun(object, visual, state.textLayout, visualTextRuns(object, visual));
        if (run === undefined) return false;
        const resolution = this.#resolveFont(run, context);
        if (resolution.source === "fallback") return false;
        drawPositionedPdfRun(
          context,
          object,
          run,
          state.textLayout,
          resolution.family,
          state.strokeStyle,
          resolution.source === "embedded",
          collectTextFragment,
        );
        return true;
      };
      for (let visibleIndex = 0; visibleIndex < visibleObjects.length; visibleIndex += 1) {
        const { object, state } = visibleObjects[visibleIndex]!;
        if (xlsxPaintLayer(object) >= 0) drawPendingGrid();
        processedObjects += 1;
        if (control !== undefined && processedObjects % (control.yieldEvery ?? 1_024) === 0) {
          await yieldRenderTurn();
        }
        assertActive();
        const { visual } = state;
        if (visual.kind === "none") continue;
        const batch = batchablePdfPath({ object, state });
        if (batch !== undefined) {
          let batchEnd = visibleIndex + 1;
          while (batchEnd < visibleObjects.length) {
            const next = visibleObjects[batchEnd]!;
            const nextBatch = batchablePdfPath(next);
            if (nextBatch === undefined
              || nextBatch.geometry.fillRule !== batch.geometry.fillRule
              || !sameBatchPaint(object, batch.fill, next.object, nextBatch.fill)
              || visibleObjects
                .slice(visibleIndex, batchEnd)
                .some((current) => intersects(current.object.bounds, next.object.bounds))) break;
            batchEnd += 1;
          }
          if (batchEnd > visibleIndex + 1) {
            const path = new Path2D();
            for (let index = visibleIndex; index < batchEnd; index += 1) {
              const current = visibleObjects[index]!;
              const geometry = (current.state.visual as BatchablePdfPath).geometry;
              tracePathCommands(path, current.object, geometry.commands);
            }
            const fill = paintStyle(context, object, batch.fill);
            if (fill !== undefined) {
              context.fillStyle = fill;
              context.fill(path, batch.geometry.fillRule);
            }
            const batchLength = batchEnd - visibleIndex;
            renderedObjectCount += batchLength;
            processedObjects += batchLength - 1;
            visibleIndex = batchEnd - 1;
            continue;
          }
        }
        const objectClipKeys: Path2D[] = [];
        for (const operation of state.operations) {
          const key = this.#shareableClipKey(operation);
          if (key === undefined) break;
          objectClipKeys.push(key);
        }
        let sharedClipCount = 0;
        while (sharedClipCount < sharedClipKeys.length
          && sharedClipKeys[sharedClipCount] === objectClipKeys[sharedClipCount]) {
          sharedClipCount += 1;
        }
        while (sharedClipKeys.length > sharedClipCount) {
          context.restore();
          sharedClipKeys.pop();
        }
        while (sharedClipKeys.length < objectClipKeys.length) {
          const operation = state.operations[sharedClipKeys.length]!;
          context.save();
          this.#applyShareableClip(context, operation);
          sharedClipKeys.push(objectClipKeys[sharedClipKeys.length]!);
        }
        if (drawDirectPdfText(visibleObjects[visibleIndex]!)) {
          let batchEnd = visibleIndex + 1;
          while (batchEnd < visibleObjects.length) {
            if (control !== undefined
              && (processedObjects + 1) % (control.yieldEvery ?? 1_024) === 0) {
              await yieldRenderTurn();
            }
            assertActive();
            if (!drawDirectPdfText(visibleObjects[batchEnd]!)) break;
            processedObjects += 1;
            batchEnd += 1;
          }
          renderedObjectCount += batchEnd - visibleIndex;
          visibleIndex = batchEnd - 1;
          continue;
        }
        context.save();
        try {
          const advancedSurfaceOperationIndex = state.operations.findIndex(({ visual }) =>
            visual.kind === "advanced-effect"
              && (visual.reflection !== undefined || visual.threeD !== undefined));
          const deferredOperationIndexes = new Set(
            state.operations
              .map(({ visual }, index) => index > advancedSurfaceOperationIndex
                && advancedSurfaceOperationIndex >= 0
                && (visual.kind === "layer" || visual.kind === "effect") ? index : -1)
              .filter((index) => index >= 0),
          );
          const applyContextOperation = (
            target: OffscreenCanvasRenderingContext2D,
            operation: VisualState["operations"][number],
          ): void => {
            if (operation.visual.kind === "layer") {
              const { a, b, c, d, e, f } = operation.visual.transform;
              target.transform(a, b, c, d, e, f);
              target.globalAlpha *= operation.visual.opacity;
              target.globalCompositeOperation = operation.visual.blendMode ?? "source-over";
            } else if (operation.visual.kind === "effect") {
              if (operation.visual.shadow !== undefined) {
                target.shadowColor = color(operation.visual.shadow.color);
                target.shadowBlur = operation.visual.shadow.blur;
                target.shadowOffsetX = operation.visual.shadow.offsetX;
                target.shadowOffsetY = operation.visual.shadow.offsetY;
              }
              if (operation.visual.clip !== undefined) {
                this.#clip(target, operation.owner, operation.visual.clip);
              }
            }
          };
          const prepareAdvancedSurface = deferredOperationIndexes.size === 0
            ? undefined
            : (target: OffscreenCanvasRenderingContext2D): void => {
              for (const index of deferredOperationIndexes) {
                applyContextOperation(target, state.operations[index]!);
              }
            };
          for (let operationIndex = objectClipKeys.length;
            operationIndex < state.operations.length;
            operationIndex += 1) {
            if (deferredOperationIndexes.has(operationIndex)) continue;
            const operation = state.operations[operationIndex]!;
            applyContextOperation(context, operation);
          }
          const advancedEffects = advancedEffectsFor(state);
          if (visual.kind === "group" || visual.kind === "opacity-mask") {
            let textOffset = 0;
            await this.#drawVisualBrushChild(
              context, object, object.bounds, visual, 0, state.strokeStyle, state.textLayout, state.textEffects,
              collectTextFragment === undefined ? undefined : (fragment, transform) => {
                // Child layouts use local ranges; selection addresses the owning object's text.
                const start = textOffset;
                textOffset += fragment.text.length;
                collectTextFragment({ ...fragment, start, end: textOffset }, transform);
              },
            );
            renderedObjectCount += 1;
          } else if (visual.kind === "shape") {
            drawWithAdvancedEffects(context, object, advancedEffects, (target) => {
              drawLegacyGeometry(target, object, visual, state.strokeStyle, state.groupedTextScale);
            }, "full", prepareAdvancedSurface);
            renderedObjectCount += 1;
          } else if (visual.kind === "painted-shape") {
            let fillImage: ImageBitmap | undefined;
            if (visual.fill.kind === "image") {
              try {
                fillImage = (await (imagePrefetch.take(state) ?? this.#paintImage(
                  visual.fill,
                  vectorRasterTargetFor(
                    visual.fill.mediaType,
                    transformedBounds(object, state.transform),
                    deviceScale,
                    this.#limits.imagePixels,
                    visual.fill,
                  ),
                  imageAdjustmentsFor(state),
                ))).bitmap;
                assertActive();
              } catch (cause) {
                diagnostics.push({
                  code: cause instanceof OfficeEngineError ? cause.code : "IMAGE_DECODE_FAILED",
                  severity: "warning",
                  fidelity: "not-rendered",
                  phase: "render",
                  message: cause instanceof Error ? cause.message : "The browser could not decode this image fill",
                  objectId: object.id,
                  part: object.source.part,
                  details: { feature: "shape-image-fill" },
                });
              }
            }
            const fillVisualPattern = visual.fill.kind === "visual"
              ? await this.#visualBrushPattern(context, object, visual.fill)
              : visual.fill.kind === "xps-gradient"
                ? await this.#xpsGradientPattern(context, object, visual.fill)
              : undefined;
            const strokeVisualPattern = visual.stroke.kind === "visual"
              ? await this.#visualBrushPattern(context, object, visual.stroke)
              : visual.stroke.kind === "xps-gradient"
                ? await this.#xpsGradientPattern(context, object, visual.stroke)
              : undefined;
            assertActive();
            drawWithAdvancedEffects(context, object, advancedEffects, (target) => {
              drawPaintedGeometry(
                target,
                object,
                visual,
                fillImage,
                state.strokeStyle,
                fillVisualPattern,
                strokeVisualPattern,
                state.groupedTextScale,
              );
            }, "full", prepareAdvancedSurface, visual.fill.kind === "none" && state.strokeStyle === undefined
              ? effectiveStrokeWidth(context, object, visual.strokeWidth / state.groupedTextScale) : undefined);
            if (visual.fill.kind === "image") imagePrefetch.consume(state);
            renderedObjectCount += 1;
          } else if (visual.kind === "text" || visual.kind === "rich-text") {
            let fillImage: ImageBitmap | undefined;
            if (visual.kind === "rich-text" && visual.fill.kind === "image") {
              try {
                fillImage = (await (imagePrefetch.take(state) ?? this.#paintImage(
                  visual.fill,
                  vectorRasterTargetFor(
                    visual.fill.mediaType,
                    transformedBounds(object, state.transform),
                    deviceScale,
                    this.#limits.imagePixels,
                    visual.fill,
                  ),
                  imageAdjustmentsFor(state),
                ))).bitmap;
                assertActive();
              } catch (cause) {
                diagnostics.push({
                  code: cause instanceof OfficeEngineError ? cause.code : "IMAGE_DECODE_FAILED",
                  severity: "warning",
                  fidelity: "not-rendered",
                  phase: "render",
                  message: cause instanceof Error ? cause.message : "The browser could not decode this image fill",
                  objectId: object.id,
                  part: object.source.part,
                  details: { feature: "shape-image-fill" },
                });
              }
            }
            const resolveTextRun = (run: SceneTextRun): FontResolution => {
              const resolution = this.#resolveFont(run, context);
              const compatibleFamily = resolution.source === "fallback"
                ? compatibleFallbackFamily(object.source.format, run.fontFamily)
                : undefined;
              return compatibleFamily === undefined
                ? resolution
                : { family: compatibleFamily, source: resolution.source };
            };
            for (const run of visualTextRuns(object, visual)) {
              const resolution = resolveTextRun(run);
              const reportKey = run.fontFamily.trim().toLowerCase();
              if (resolution.source === "fallback" && !reportedFonts.has(reportKey)) {
                reportedFonts.add(reportKey);
                diagnostics.push({
                  code: "FONT_SUBSTITUTED",
                  severity: "warning",
                  fidelity: "approximate",
                  phase: "render",
                  message: `Font ${run.fontFamily} was not resolved by the configured font sources; ${resolution.family} was used as an approximate substitute`,
                  objectId: object.id,
                  part: object.source.part,
                  details: {
                    family: run.fontFamily,
                    fallbackFamily: resolution.family,
                  },
                });
              }
            }
            const drawShape = (target: OffscreenCanvasRenderingContext2D): void => {
              if (visual.kind === "text") {
                drawLegacyGeometry(target, object, visual, state.strokeStyle, state.groupedTextScale);
              }
              else drawPaintedGeometry(
                target,
                object,
                visual,
                fillImage,
                state.strokeStyle,
                undefined,
                undefined,
                state.groupedTextScale,
              );
            };
            const glyphPaintStyles = await this.#textPaintStyles(context, object, visual, state.textLayout, diagnostics);
            const drawOwnText = (target: OffscreenCanvasRenderingContext2D): void => drawText(
              target,
              object,
              visual,
              state.textLayout,
              (_family, run) => resolveTextRun(run).family,
              (run) => {
                const resolution = resolveTextRun(run);
                run = fitSubstitutedSongGlyph(target, run, resolution);
                const spaceAdvanceEm = resolution.spaceAdvanceEm
                  ?? (resolution.source === "fallback"
                    ? fallbackSpaceAdvanceEm(object.source.format, run.fontFamily) : undefined);
                const letterSpacing = spaceAdvanceEm === undefined ? run.letterSpacing
                  : guardedFallbackSpaceLetterSpacing(target, run, resolution.family, spaceAdvanceEm);
                if (resolution.source !== "fallback") {
                  const text = symbolFontGlyphText(run.text, run.fontFamily, {
                    family: resolution.family,
                    style: run.italic ? "italic" : "normal",
                    weight: run.bold ? 700 : 400,
                    stretch: "normal",
                  }, target);
                  return text === run.text && letterSpacing === run.letterSpacing ? run : { ...run, text, letterSpacing };
                }
                const text = fallbackGlyphText(run.text, run.fontFamily);
                const fontSizeFactor =
                  fallbackFontSizeFactor(object.source.format, run.fontFamily) ?? 1;
                const fontSize = run.fontSize * fontSizeFactor;
                const baselineShift = run.baselineShift
                  + (fallbackBaselineShift(
                    object.source.format,
                    run.fontFamily,
                    run.fontSize,
                  ) ?? 0);
                return text === run.text
                    && letterSpacing === run.letterSpacing
                    && fontSize === run.fontSize
                    && baselineShift === run.baselineShift
                  ? run
                  : { ...run, text, letterSpacing, fontSize, baselineShift };
              },
              (runs) => runs.flatMap((run) => {
                const resolution = resolveTextRun(run);
                const segments = resolution.source === "fallback"
                  ? fontFallbackTextSegments(run.text).map(segment => segment.text) : [run.text];
                const preserveSpaces = resolution.spaceAdvanceEm !== undefined
                  || (resolution.source === "fallback"
                    && fallbackSpaceAdvanceEm(object.source.format, run.fontFamily) !== undefined);
                let cursor = run.sourceStart;
                return segments.flatMap(text => preserveSpaces ? text.split(/( +)/u).filter(Boolean) : [text])
                  .map((text) => {
                    const sourceStart = cursor;
                    const sourceEnd = sourceStart === undefined ? undefined : sourceStart + text.length;
                    cursor = sourceEnd;
                    return {
                      ...run,
                      text,
                      ...(sourceStart === undefined || sourceEnd === undefined ? {} : {
                        sourceStart,
                        sourceEnd,
                      }),
                    };
                  });
              }),
              this.#textLayouts,
              `${object.numericId}:${object.bounds.width}:${object.bounds.height}`,
              state.groupedTextScaleX,
              state.groupedTextScaleY,
              state.strokeStyle,
              state.textEffects,
              collectTextFragment,
              glyphPaintStyles,
            );
            if (advancedEffects.some(({ threeD }) => threeD !== undefined)) {
              drawWithAdvancedEffects(context, object, advancedEffects, drawShape, "full", prepareAdvancedSurface);
              const textThreeD = advancedEffects.some(({ threeD }) => threeD?.appliesToText === true);
              const textEffects = !textThreeD && (visual.kind === "rich-text"
                ? opaquePaint(visual.fill) : (visual.fill & 255) === 255)
                ? advancedEffects.map(({ outerShadow, innerShadow, ...effect }) => effect)
                : advancedEffects;
              drawWithAdvancedEffects(context, object, textEffects, drawOwnText, textThreeD ? "full" : "surface", prepareAdvancedSurface);
            } else {
              drawWithAdvancedEffects(context, object, advancedEffects, (target) => {
                drawShape(target);
                drawOwnText(target);
              }, "full", prepareAdvancedSurface);
            }
            if (visual.kind === "rich-text" && visual.fill.kind === "image") {
              imagePrefetch.consume(state);
            }
            renderedObjectCount += 1;
          } else if (visual.kind === "image") {
            try {
              const decoded = await (imagePrefetch.take(state)
                ?? this.#image(
                  visual,
                  imageColorChangesFor(state),
                  imageAdjustmentsFor(state),
                  vectorRasterTarget(object, state, deviceScale, this.#limits.imagePixels),
                  object.source.format === "ppt" ? 9 : 0,
                ));
              assertActive();
              const bitmap = decoded.bitmap;
              const { cropLeft, cropTop, cropRight, cropBottom } = visual;
              drawWithAdvancedEffects(context, object, advancedEffects, (target) => {
                const previousSmoothing = target.imageSmoothingEnabled;
                const authoredInterpolationDisabled = visual.mediaType.includes(";interpolate=false");
                const transform = canvasTransform(target)
                  ?? { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
                const renderedWidth = Math.hypot(transform.a, transform.b) * object.bounds.width;
                const renderedHeight = Math.hypot(transform.c, transform.d) * object.bounds.height;
                // PDF Interpolate controls magnification. Like pdf.js/Firefox, retain
                // smoothing while reducing an image even when the flag is false.
                target.imageSmoothingEnabled = !authoredInterpolationDisabled
                  || (renderedWidth <= bitmap.width + 1e-3
                    && renderedHeight <= bitmap.height + 1e-3);
                try {
                    const scaled = object.source.format === "pdf"
                      ? progressivelyDownscaleImage(bitmap, renderedWidth, renderedHeight)
                      : { source: bitmap, width: bitmap.width, height: bitmap.height };
                  if (cropLeft === 0 && cropTop === 0 && cropRight === 0 && cropBottom === 0) {
                    // PowerPoint preserves native bitmap pixels when the authored extent differs
                    // from the decoded size only by a subpixel rounding remainder. Avoid making
                    // Canvas resample an otherwise 1:1 image for that harmless remainder.
                    const preserveNativeExtent = Math.abs(object.bounds.width * deviceScale - bitmap.width) <= 0.125
                      && Math.abs(object.bounds.height * deviceScale - bitmap.height) <= 0.125;
                    const width = preserveNativeExtent ? bitmap.width / deviceScale : object.bounds.width;
                    const height = preserveNativeExtent ? bitmap.height / deviceScale : object.bounds.height;
                    if (object.source.format === "pdf") {
                      drawPdfImageAtIntegerCoordinates(
                        target,
                        scaled.source,
                        0,
                        0,
                        scaled.width,
                        scaled.height,
                        object.bounds.x,
                        object.bounds.y,
                        width,
                        height,
                      );
                    } else {
                      target.drawImage(scaled.source, object.bounds.x, object.bounds.y, width, height);
                    }
                  } else {
                    const sourceWidth = scaled.width * (1 - cropLeft - cropRight);
                    const sourceHeight = scaled.height * (1 - cropTop - cropBottom);
                    if (sourceWidth > 0 && sourceHeight > 0) {
                      if (object.source.format === "pdf") {
                        drawPdfImageAtIntegerCoordinates(
                          target,
                          scaled.source,
                          scaled.width * cropLeft,
                          scaled.height * cropTop,
                          sourceWidth,
                          sourceHeight,
                          object.bounds.x,
                          object.bounds.y,
                          object.bounds.width,
                          object.bounds.height,
                        );
                      } else {
                        drawImageRegion(target,
                          scaled.source,
                          scaled.width * cropLeft,
                          scaled.height * cropTop,
                          sourceWidth,
                          sourceHeight,
                          object.bounds.x,
                          object.bounds.y,
                          object.bounds.width,
                          object.bounds.height,
                        );
                      }
                    }
                  }
                } finally {
                  target.imageSmoothingEnabled = previousSmoothing;
                }
              }, "full", prepareAdvancedSurface);
              renderedObjectCount += 1;
              if (decoded.usedFallback) {
                diagnostics.push({
                  code: "IMAGE_FALLBACK_USED",
                  severity: "warning",
                  fidelity: "approximate",
                  phase: "render",
                  message: `The preferred ${visual.mediaType} image could not be decoded; its embedded ${decoded.mediaType} fallback was rendered`,
                  objectId: object.id,
                  part: object.source.part,
                  details: {
                    preferredMediaType: visual.mediaType,
                    fallbackMediaType: decoded.mediaType,
                  },
                });
              }
              if (decoded.approximate) {
                diagnostics.push({
                  code: "IMAGE_FIDELITY_APPROXIMATE",
                  severity: "warning",
                  fidelity: "approximate",
                  phase: "render",
                  message: approximateImageMessage(decoded.mediaType),
                  objectId: object.id,
                  part: object.source.part,
                });
              }
            } catch (cause) {
              if (this.#closed) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
              diagnostics.push({
                code: cause instanceof OfficeEngineError ? cause.code : "IMAGE_DECODE_FAILED",
                severity: "warning",
                fidelity: "not-rendered",
                phase: "render",
                message: cause instanceof Error ? cause.message : "The browser could not decode this image",
                objectId: object.id,
                part: object.source.part,
              });
            } finally {
              imagePrefetch.consume(state);
            }
          }
        } finally {
          context.restore();
        }
      }
      while (sharedClipKeys.length > 0) {
        context.restore();
        sharedClipKeys.pop();
      }
      drawPendingGrid();
      // The requested background is the display medium, not part of the
      // document backdrop. Paint it underneath after document compositing so
      // blend modes see transparency outside authored page content.
      context.save();
      try {
        context.globalAlpha = 1;
        context.globalCompositeOperation = "destination-over";
        context.fillStyle = request.background ?? "#ffffff";
        context.fillRect(
          viewport.x,
          viewport.y,
          canvasWidth / canvasScaleX,
          canvasHeight / canvasScaleY,
        );
      } finally {
        context.restore();
      }
    } finally {
      context.restore();
      imagePrefetch.release();
    }
    assertActive();

    let bitmap: ImageBitmap;
    if (supersample === 1) {
      if (watermark !== undefined) drawWatermark(context, pixelWidth, pixelHeight, watermark, pixelRatio);
      bitmap = canvas.transferToImageBitmap();
    } else {
      const output = new OffscreenCanvas(pixelWidth, pixelHeight);
      const outputContext = output.getContext("2d", { alpha: true });
      if (outputContext === null) {
        throw new OfficeEngineError("UNSUPPORTED_ENVIRONMENT", "A Canvas 2D context is required for rendering");
      }
      outputContext.imageSmoothingEnabled = true;
      outputContext.imageSmoothingQuality = "high";
      outputContext.drawImage(canvas, 0, 0, pixelWidth, pixelHeight);
      if (watermark !== undefined) drawWatermark(outputContext, pixelWidth, pixelHeight, watermark, pixelRatio);
      bitmap = output.transferToImageBitmap();
    }
    let selectableTextFragments: readonly RenderedTextFragment[] | undefined;
    if (request.includeTextFragments === true) {
      const semanticText = new Map((unitObjects ?? []).flatMap((object) =>
        object.text === undefined ? [] : [[object.id, object.text] as const]));
      selectableTextFragments = semanticTextFragments(textFragments, semanticText);
    }
    const result: Omit<RenderResult, "bitmap"> = {
      viewport: { ...viewport },
      pixelWidth,
      pixelHeight,
      renderedObjectCount,
      media,
      ...(selectableTextFragments === undefined ? {} : { textFragments: selectableTextFragments }),
      diagnostics,
    };
    const pixels = pixelWidth * pixelHeight;
    if (typeof createImageBitmap !== "function"
      || bitmap.width !== pixelWidth
      || bitmap.height !== pixelHeight
      || pixels > this.#renderedFramePixelBudget()) return { ...result, bitmap };
    this.#evictRenderedFrames(pixels);
    if (this.#retainedFramePixels + pixels > this.#renderedFramePixelBudget()) {
      return { ...result, bitmap };
    }
    const cached: CachedRenderedFrame = { unitIndex: unit.index, bitmap, pixels, result };
    this.#renderedFrames.set(renderedFrameKey, cached);
    this.#retainedFramePixels += pixels;
    let output: ImageBitmap | undefined;
    try {
      output = await createImageBitmap(bitmap);
      assertActive();
      return { ...result, bitmap: output };
    } catch (cause) {
      output?.close();
      this.#renderedFrames.delete(renderedFrameKey);
      this.#retainedFramePixels -= pixels;
      if (cause instanceof OfficeEngineError) {
        bitmap.close();
        throw cause;
      }
      return { ...result, bitmap };
    }
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    const prefetchWaiters = [...this.#imagePrefetchWaiters];
    this.#imagePrefetchWaiters.clear();
    const queuedDecodes = this.#imageDecodeQueue.splice(0);
    for (const image of this.#images.values()) {
      void image.promise.then(({ bitmap }) => bitmap.close(), () => undefined);
    }
    for (const frame of this.#renderedFrames.values()) frame.bitmap.close();
    for (const bitmap of this.#retiredImages) bitmap.close();
    this.#retiredImages.clear();
    this.#images.clear();
    this.#renderedFrames.clear();
    this.#imageSources.clear();
    this.#fontAvailability.clear();
    this.#fontResolutions.clear();
    this.#textLayouts.clear();
    this.#clipPathsByGeometry = new WeakMap();
    this.#preparedUnits.clear();
    this.#objectsByUnit.clear();
    this.#imagePrefetches.clear();
    this.#imagePins.clear();
    this.#imagePixels = 0;
    this.#retainedImagePixels = 0;
    this.#retainedFramePixels = 0;
    this.#imagePrefetchPixels = 0;
    this.#fonts = undefined;
    for (const resume of prefetchWaiters) resume();
    const closeError = new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    for (const queued of queuedDecodes) queued.reject(closeError);
  }

  #prepareUnit(
    unit: UnitDescriptor,
    request: RenderRequest,
    mappers: { readonly rows: SheetAxisMapper; readonly columns: SheetAxisMapper } | undefined,
  ): PreparedSceneUnit {
    const key = unit.type === "sheet" ? sheetSizeCacheKey(request) : "";
    const cached = this.#preparedUnits.get(request.unitIndex);
    if (cached?.key === key) {
      this.#preparedUnits.delete(request.unitIndex);
      this.#preparedUnits.set(request.unitIndex, cached);
      return cached;
    }

    const sceneObjects = (this.#objectsByUnit.get(request.unitIndex) ?? [])
      .map((object) => {
        const mapped = mappers === undefined
          ? object
          : mapSheetObject(object, unit.index, mappers.rows, mappers.columns);
        return unit.type === "sheet"
          ? collapseXlsxCellBorder(
            mapped,
            mappers?.columns.total ?? unit.width,
            mappers?.rows.total ?? unit.height,
          )
          : mapped;
      });
    const objectsById = new Map(sceneObjects.map((object) => [object.numericId, object]));
    const objects = cullInvisibleObjects(sceneObjects
      .flatMap(expandPositionedTextBatch)
      .filter((object) => isSceneObjectVisible(object, objectsById))
      .map((object) => {
        const state = resolveVisualState(object, objectsById);
        return { object, bounds: paintedBounds(object, state), state };
      })
      .sort(({ object: left }, { object: right }) => (
        left.z - right.z || left.numericId - right.numericId
      )))
      .sort((left, right) => (
        left.bounds.y - right.bounds.y || left.object.numericId - right.object.numericId
      ));
    const maximumBottoms = new Float64Array(objects.length);
    let maximumBottom = Number.NEGATIVE_INFINITY;
    for (let index = 0; index < objects.length; index += 1) {
      const bounds = objects[index]!.bounds;
      maximumBottom = Math.max(maximumBottom, bounds.y + bounds.height);
      maximumBottoms[index] = maximumBottom;
    }
    const prepared = { key, objects, maximumBottoms };
    this.#preparedUnits.set(request.unitIndex, prepared);
    while (this.#preparedUnits.size > 32) {
      const oldest = this.#preparedUnits.keys().next().value;
      if (oldest === undefined) break;
      this.#preparedUnits.delete(oldest);
    }
    return prepared;
  }

  #clip(
    context: OffscreenCanvasRenderingContext2D,
    owner: SceneObject,
    geometry: SceneGeometry,
  ): void {
    if (typeof geometry !== "object" || geometry.kind !== "path") {
      context.clip(traceGeometry(context, owner, geometry));
      return;
    }
    context.clip(this.#clipPath(owner, geometry), geometry.fillRule);
  }

  #clipPath(
    owner: SceneObject,
    geometry: Extract<SceneGeometry, { readonly kind: "path" }>,
  ): Path2D {
    const cached = this.#clipPathsByGeometry.get(geometry);
    if (cached?.x === owner.bounds.x && cached.y === owner.bounds.y) return cached.path;
    const path = path2d(owner, geometry);
    this.#clipPathsByGeometry.set(geometry, { x: owner.bounds.x, y: owner.bounds.y, path });
    return path;
  }

  #shareableClipKey(operation: VisualState["operations"][number]): Path2D | undefined {
    const { owner, visual } = operation;
    if (visual.kind !== "effect"
      || visual.shadow !== undefined
      || typeof visual.clip !== "object"
      || visual.clip.kind !== "path") {
      return undefined;
    }
    return this.#clipPath(owner, visual.clip);
  }

  #applyShareableClip(
    context: OffscreenCanvasRenderingContext2D,
    operation: VisualState["operations"][number],
  ): void {
    const { owner, visual } = operation;
    if (visual.kind === "effect" && typeof visual.clip === "object" && visual.clip.kind === "path") {
      this.#clip(context, owner, visual.clip);
    }
  }

  #resolveFont(
    run: SceneTextRun,
    context?: FontProbeContext,
  ): FontResolution {
    const { fontFamily: family, fontSize: size, italic, bold } = run;
    const exactPdfFaceKey = / PDF \d+ \d+$/u.test(family)
      ? ["pdf", family, String(size), italic ? "italic" : "normal", bold ? "700" : "400"].join("\0")
      : undefined;
    const exactPdfResolution = exactPdfFaceKey === undefined
      ? undefined
      : this.#fontResolutions.get(exactPdfFaceKey);
    if (exactPdfResolution !== undefined) return exactPdfResolution;
    const cacheKey = [family, String(size), italic ? "italic" : "normal", bold ? "700" : "400", run.text].join("\0");
    const cached = this.#fontResolutions.get(cacheKey);
    if (cached !== undefined) return cached;
    const codePoints = runCodePoints(semanticSymbolFontText(run.text, family));
    const fonts = this.#fonts;
    if (fonts !== undefined && "resolveFace" in fonts && typeof fonts.resolveFace === "function") {
      const face = fonts.resolveFace({
        family,
        style: italic ? "italic" : "normal",
        weight: bold ? 700 : 400,
        stretch: "normal",
        codePoints,
      });
      if ("source" in face && typeof face.source === "string") {
        const resolution = { family: face.family, source: face.source as FontResolution["source"],
          ...("spaceAdvanceEm" in face && typeof face.spaceAdvanceEm === "number"
            ? { spaceAdvanceEm: face.spaceAdvanceEm } : {}) };
        this.#fontResolutions.set(cacheKey, resolution);
        if (exactPdfFaceKey !== undefined && resolution.source === "embedded") {
          this.#fontResolutions.set(exactPdfFaceKey, resolution);
        }
        return resolution;
      }
    }
    if (fonts !== undefined && "resolve" in fonts && typeof fonts.resolve === "function") {
      const resolution = fonts.resolve(family);
      this.#fontResolutions.set(cacheKey, resolution);
      return resolution;
    }
    const font = fontShorthand(family, size, italic, bold);
    if (this.#fontUnavailable(font, family, italic, bold, run.text, codePoints, context)) {
      const resolution = { family: approximateFontFamily(family, codePoints), source: "fallback" } as const;
      this.#fontResolutions.set(cacheKey, resolution);
      return resolution;
    }
    const resolution = { family, source: "browser" } as const;
    this.#fontResolutions.set(cacheKey, resolution);
    return resolution;
  }

  #fontUnavailable(
    font: string,
    family: string,
    italic: boolean,
    bold: boolean,
    text: string,
    codePoints: readonly number[],
    context: FontProbeContext | undefined,
  ): boolean {
    const cacheKey = `${font}\0${codePoints.join(",")}`;
    const cached = this.#fontAvailability.get(cacheKey);
    if (cached !== undefined) return cached;
    const fonts = this.#fonts;
    if (fonts === undefined || !("check" in fonts) || typeof fonts.check !== "function") return false;
    let unavailable = false;
    try {
      unavailable = !fonts.check(font, text);
    } catch {
      return false;
    }
    if (!unavailable) {
      unavailable = probeBrowserFontFace({
        family,
        style: italic ? "italic" : "normal",
        weight: bold ? 700 : 400,
        stretch: "normal",
      }, context, text) === false;
    }
    this.#fontAvailability.set(cacheKey, unavailable);
    return unavailable;
  }

  #image(
    visual: Extract<SceneVisual, { kind: "image" }>,
    colorChanges: readonly SceneImageColorChange[],
    adjustments: readonly SceneImageAdjustment[] = [],
    target?: ImageRasterTarget,
    colorChangeTolerance = 0,
  ): Promise<DecodedSceneImage> {
    if (this.#closed) {
      return Promise.reject(new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed"));
    }
    const fallbackKey = visual.fallback === undefined
      ? ""
      : `:fallback:${this.#imageResourceKey(visual.fallback.bytes, visual.fallback.mediaType)}`;
    const changesKey = colorChanges.length === 0
      ? ""
      : `:changes:${colorChangeTolerance}:${colorChanges.map((change) => `${change.from}:${change.to}:${change.useAlpha ? 1 : 0}`).join(",")}`;
    const adjustmentsKey = adjustments.length === 0
      ? ""
      : `:adjustments:${imageAdjustmentKey(adjustments)}`;
    const maskKey = visual.alphaMask === undefined
      ? ""
      : `:mask:${visual.alphaMask.width}x${visual.alphaMask.height}:${this.#imageResourceKey(visual.alphaMask.bytes, visual.alphaMask.mediaType)}`;
    const targetKey = target === undefined ? "" : `:target:${target.width}x${target.height}`;
    const key = `${this.#imageResourceKey(visual.bytes, visual.mediaType)}${fallbackKey}${maskKey}${changesKey}${adjustmentsKey}${targetKey}`;
    const cached = this.#reuseImage(key);
    if (cached !== undefined) return cached;
    const source = this.#decodeImageCandidate(
      visual.bytes,
      visual.mediaType,
      key,
      target,
    ).then(
      (image): DecodedSceneImage => ({ ...image, mediaType: visual.mediaType, usedFallback: false }),
      async (preferredCause): Promise<DecodedSceneImage> => {
        if (this.#closed) {
          throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
        }
        if (visual.fallback === undefined) throw preferredCause;
        try {
          const image = await this.#decodeImageCandidate(
            visual.fallback.bytes,
            visual.fallback.mediaType,
            key,
            target,
          );
          return { ...image, mediaType: visual.fallback.mediaType, usedFallback: true };
        } catch (fallbackCause) {
          const preferredMessage = preferredCause instanceof Error
            ? preferredCause.message
            : "unknown preferred-image decode failure";
          const fallbackMessage = fallbackCause instanceof Error
            ? fallbackCause.message
            : "unknown fallback decode failure";
          throw new OfficeEngineError(
            fallbackCause instanceof OfficeEngineError ? fallbackCause.code : "IMAGE_DECODE_FAILED",
            `Preferred image decode failed (${preferredMessage}); raster fallback also failed (${fallbackMessage})`,
            { cause: fallbackCause },
          );
        }
      },
    );
    const decoded = (visual.alphaMask === undefined
      ? source
      : Promise.allSettled([
        source,
        this.#decodeImageCandidate(
          visual.alphaMask.bytes,
          visual.alphaMask.mediaType,
          key,
        ),
      ]).then((results): DecodedSceneImage => {
        const [imageResult, maskResult] = results;
        if (imageResult.status === "rejected" || maskResult.status === "rejected") {
          if (imageResult.status === "fulfilled") {
            const pixels = imageResult.value.bitmap.width * imageResult.value.bitmap.height;
            imageResult.value.bitmap.close();
            this.#imagePixels = Math.max(0, this.#imagePixels - pixels);
          }
          if (maskResult.status === "fulfilled") {
            const pixels = maskResult.value.bitmap.width * maskResult.value.bitmap.height;
            maskResult.value.bitmap.close();
            this.#imagePixels = Math.max(0, this.#imagePixels - pixels);
          }
          if (imageResult.status === "rejected") throw imageResult.reason;
          if (maskResult.status === "rejected") throw maskResult.reason;
        }
        const image = imageResult.value;
        const maskImage = maskResult.value;
        try {
          const pixels = image.bitmap.width * image.bitmap.height;
          this.#evictImages(pixels * 3, key);
          const maskPixels = maskImage.bitmap.width * maskImage.bitmap.height;
          const bitmap = applyImageAlphaMask(
            image.bitmap,
            maskImage.bitmap,
            visual.alphaMask?.mediaType.includes(";pdf-stencil-painted=0") === true,
            visual.alphaMask?.mediaType.includes(";pdf-stencil-flip-y=true") === true,
            this.#limits.imagePixels,
            this.#retainedImagePixels + pixels + maskPixels,
            this.#limits.totalImagePixels,
          );
          this.#imagePixels = Math.max(0, this.#imagePixels - maskPixels);
          return {
            ...image,
            bitmap,
          };
        } catch (cause) {
          const decodedPixels = image.bitmap.width * image.bitmap.height;
          image.bitmap.close();
          const maskPixels = maskImage.bitmap.width * maskImage.bitmap.height;
          maskImage.bitmap.close();
          this.#imagePixels = Math.max(0, this.#imagePixels - decodedPixels - maskPixels);
          throw cause;
        }
    })).then((image): DecodedSceneImage => {
      if (colorChanges.length === 0 && adjustments.length === 0) return image;
      try {
        const pixels = image.bitmap.width * image.bitmap.height;
        this.#evictImages(pixels * 3, key);
        const colorChanged = applyImageColorChanges(
          image.bitmap,
          colorChanges,
          colorChangeTolerance,
          this.#limits.imagePixels,
          this.#retainedImagePixels + pixels,
          this.#limits.totalImagePixels,
        );
        return {
          ...image,
          bitmap: applyImageAdjustments(
            colorChanged,
            adjustments,
            this.#limits.imagePixels,
            this.#retainedImagePixels + pixels,
            this.#limits.totalImagePixels,
          ),
        };
      } catch (cause) {
        const decodedPixels = image.bitmap.width * image.bitmap.height;
        image.bitmap.close();
        this.#imagePixels = Math.max(0, this.#imagePixels - decodedPixels);
        throw cause;
      }
    });
    const entry: CachedSceneImage = { promise: decoded };
    this.#images.set(key, entry);
    void decoded.then((image) => {
      if (this.#closed || this.#images.get(key) !== entry) return;
      entry.image = image;
      entry.pixels = image.bitmap.width * image.bitmap.height;
      this.#retainedImagePixels += entry.pixels;
      this.#evictImages(0, key);
    }, () => {
      if (this.#images.get(key) === entry) this.#images.delete(key);
    });
    return decoded;
  }

  #prefetchImages(
    objects: readonly { readonly object: SceneObject; readonly state: VisualState }[],
    isActive: () => boolean,
    deviceScale: number,
  ): ImagePrefetch {
    const tasks: {
      readonly state: VisualState;
      readonly pixels: number;
      readonly start: () => Promise<DecodedSceneImage>;
    }[] = [];
    const imagePixels = (bytes: Uint8Array, mediaType: string): number => {
      try {
        return identifyImage(bytes, mediaType, this.#limits.imagePixels).pixels;
      } catch {
        return this.#limits.imagePixels;
      }
    };
    for (const { object, state } of objects) {
      const { visual } = state;
      if (visual.kind === "image") {
        const changes = imageColorChangesFor(state);
        const adjustments = imageAdjustmentsFor(state);
        const target = vectorRasterTarget(
          object,
          state,
          deviceScale,
          this.#limits.imagePixels,
        );
        const fallbackPixels = visual.fallback === undefined
          ? 0
          : imagePixels(visual.fallback.bytes, visual.fallback.mediaType);
        const maskPixels = visual.alphaMask === undefined
          ? 0
          : visual.alphaMask.width * visual.alphaMask.height;
        const sourcePixels = Math.max(
          target === undefined ? imagePixels(visual.bytes, visual.mediaType) : target.width * target.height,
          fallbackPixels,
        );
        tasks.push({
          state,
          pixels: sourcePixels * (visual.alphaMask !== undefined || changes.length > 0 || adjustments.length > 0 ? 3 : 1)
            + maskPixels,
          start: () => this.#image(
            visual,
            changes,
            adjustments,
            target,
            // OfficeArt pictureTransparent uses a 9-level RGB tolerance (MS DFF import).
            // DrawingML color replacement retains exact matching.
            object.source.format === "ppt" ? 9 : 0,
          ),
        });
      } else if ((visual.kind === "painted-shape" || visual.kind === "rich-text")
        && visual.fill.kind === "image") {
        const fill = visual.fill;
        const target = vectorRasterTargetFor(
          fill.mediaType,
          transformedBounds(object, state.transform),
          deviceScale,
          this.#limits.imagePixels,
          fill,
        );
        tasks.push({
          state,
          pixels: target === undefined ? imagePixels(fill.bytes, fill.mediaType) : target.width * target.height,
          start: () => this.#paintImage(fill, target, imageAdjustmentsFor(state)),
        });
      }
    }
    const tasksByState = new Map(tasks.map((task) => [task.state, task]));
    const promises = new Map<VisualState, Promise<DecodedSceneImage>>();
    const pinnedPromises = new Map<VisualState, Promise<DecodedSceneImage>>();
    const passed = new Set<VisualState>();
    if (tasks.length === 0) {
      return {
        take: () => undefined,
        consume: () => undefined,
        release: () => undefined,
      };
    }

    const pixelBudget = Math.min(
      IMAGE_PREFETCH_PIXEL_BUDGET,
      Math.floor(this.#limits.totalImagePixels / 4),
    );
    let next = 0;
    let released = false;
    let waitingForSlot = false;
    const pendingDemands = new Map<VisualState, {
      readonly resume: () => void;
      readonly reject: (cause: unknown) => void;
    }>();
    const schedulingError = (): OfficeEngineError => this.#closed
      ? new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed")
      : new OfficeEngineError("OPERATION_ABORTED", "Rendering was superseded or cancelled");
    const canStart = (pixels: number): boolean => (
      this.#imagePrefetches.size < IMAGE_PREFETCH_CONCURRENCY
      && (pixels > pixelBudget
        ? this.#imagePrefetches.size === 0
        : this.#imagePrefetchPixels + pixels <= pixelBudget)
    );
    const launch = (task: typeof tasks[number]): Promise<DecodedSceneImage> => {
      const promise = task.start();
      pinnedPromises.set(task.state, promise);
      this.#pinImage(promise);
      this.#trackImagePrefetch(promise, task.pixels);
      return promise;
    };
    const resume = (): void => {
      waitingForSlot = false;
      pump();
    };
    const pump = (): void => {
      while (!released && isActive() && !this.#closed && next < tasks.length) {
        const task = tasks[next]!;
        if (passed.has(task.state)) {
          next += 1;
          continue;
        }
        if (!canStart(task.pixels)) {
          if (!waitingForSlot) {
            waitingForSlot = true;
            this.#imagePrefetchWaiters.add(resume);
          }
          return;
        }
        next += 1;
        const promise = launch(task);
        promises.set(task.state, promise);
      }
    };
    const scheduleDemand = (task: typeof tasks[number]): Promise<DecodedSceneImage> => (
      new Promise<DecodedSceneImage>((resolve, reject) => {
        const resumeDemand = (): void => {
          this.#imagePrefetchWaiters.delete(resumeDemand);
          if (released || this.#closed || !isActive()) {
            pendingDemands.delete(task.state);
            reject(schedulingError());
            return;
          }
          if (!canStart(task.pixels)) {
            this.#imagePrefetchWaiters.add(resumeDemand);
            return;
          }
          pendingDemands.delete(task.state);
          let decoded: Promise<DecodedSceneImage>;
          try {
            decoded = launch(task);
          } catch (cause) {
            reject(cause);
            return;
          }
          void decoded.then(resolve, reject);
          pump();
        };
        pendingDemands.set(task.state, { resume: resumeDemand, reject });
        resumeDemand();
      })
    );
    pump();
    return {
      take: (state) => {
        const promise = promises.get(state);
        if (promise !== undefined) return promise;
        const task = tasksByState.get(state);
        if (task === undefined) return undefined;
        passed.add(state);
        if (waitingForSlot) {
          waitingForSlot = false;
          this.#imagePrefetchWaiters.delete(resume);
        }
        const demanded = scheduleDemand(task);
        promises.set(state, demanded);
        return demanded;
      },
      consume: (state) => {
        passed.add(state);
        const pinned = pinnedPromises.get(state);
        if (pinned !== undefined) this.#releaseImagePrefetch(pinned);
        promises.delete(state);
        pump();
      },
      release: () => {
        if (released) return;
        released = true;
        waitingForSlot = false;
        this.#imagePrefetchWaiters.delete(resume);
        const cause = schedulingError();
        for (const pending of pendingDemands.values()) {
          this.#imagePrefetchWaiters.delete(pending.resume);
          pending.reject(cause);
        }
        pendingDemands.clear();
        for (const promise of pinnedPromises.values()) this.#unpinImage(promise);
        pinnedPromises.clear();
        promises.clear();
      },
    };
  }

  #pinImage(promise: Promise<DecodedSceneImage>): void {
    this.#imagePins.set(promise, (this.#imagePins.get(promise) ?? 0) + 1);
  }

  #unpinImage(promise: Promise<DecodedSceneImage>): void {
    const pins = this.#imagePins.get(promise);
    if (pins === undefined) return;
    if (pins === 1) {
      this.#imagePins.delete(promise);
      this.#releaseImagePrefetch(promise);
    } else {
      this.#imagePins.set(promise, pins - 1);
    }
    this.#evictImages(0, "");
  }

  #trackImagePrefetch(promise: Promise<DecodedSceneImage>, pixels: number): void {
    if (this.#imagePrefetches.has(promise)) return;
    this.#imagePrefetches.set(promise, pixels);
    this.#imagePrefetchPixels += pixels;
  }

  async #textPaintStyles(
    context: OffscreenCanvasRenderingContext2D,
    object: SceneObject,
    visual: Extract<SceneVisual, { kind: "text" | "rich-text" }>,
    layout?: SceneTextLayout,
    diagnostics?: Diagnostic[],
  ): Promise<Map<ScenePaint, string | CanvasGradient | CanvasPattern>> {
    const result = new Map<ScenePaint, string | CanvasGradient | CanvasPattern>();
    const paints = new Set([layout?.textPaint, ...visualTextRuns(object, visual).map(run => run.paint)]);
    for (const paint of paints) {
      if (paint === undefined) continue;
      if (paint.kind === "none") { result.set(paint, "rgba(0,0,0,0)"); continue; }
      try {
        let style = paintStyle(context, object, paint);
        if (paint.kind === "image") {
          style = imagePaintPattern(context, object, paint, (await this.#paintImage(paint)).bitmap);
        } else if (paint.kind === "visual") {
          style = await this.#visualBrushPattern(context, object, paint);
        } else if (paint.kind === "xps-gradient") {
          style = await this.#xpsGradientPattern(context, object, paint);
        } else if (paint.kind === "rect-gradient" || paint.kind === "circle-gradient" || paint.kind === "shape-gradient") {
          const transform = canvasTransform(context) ?? IDENTITY_TRANSFORM;
          const raster = gradientRasterSize(object.bounds.width, object.bounds.height,
            Math.hypot(transform.a, transform.b), Math.hypot(transform.c, transform.d));
          const canvas = new OffscreenCanvas(raster.width, raster.height);
          const target = canvas.getContext("2d");
          if (target === null) continue;
          target.scale(raster.scaleX, raster.scaleY);
          const localObject = { ...object, bounds: { ...object.bounds, x: 0, y: 0 } };
          traceGeometry(target, localObject, "rectangle");
          if (paint.kind === "rect-gradient") drawRectGradientPaint(target, localObject, "rectangle", "nonzero", paint);
          else if (paint.kind === "circle-gradient") drawCircleGradientPaint(target, localObject, "rectangle", "nonzero", paint);
          else drawShapeGradientPaint(target, localObject, "rectangle", "nonzero", paint);
          const pattern = context.createPattern(canvas, "no-repeat");
          pattern?.setTransform({ a: 1 / raster.scaleX, d: 1 / raster.scaleY, e: object.bounds.x, f: object.bounds.y });
          style = pattern ?? undefined;
        }
        if (style !== undefined) result.set(paint, style);
      } catch (cause) {
        diagnostics?.push({ code: "IMAGE_DECODE_FAILED", severity: "warning", fidelity: "approximate", phase: "render",
          message: cause instanceof Error ? cause.message : "Could not render text fill",
          objectId: object.id, part: object.source.part, details: { feature: "text-fill" } });
      }
    }
    return result;
  }

  async #visualBrushPattern(
    context: OffscreenCanvasRenderingContext2D,
    object: SceneObject,
    paint: Extract<ScenePaint, { kind: "visual" }>,
    depth = 0,
  ): Promise<CanvasPattern | undefined> {
    if (depth > 64 || paint.children.length === 0 || paint.opacity === 0) return undefined;
    const content = paint.children.reduce((bounds, child) => {
      const childObject = { ...object, bounds: child.bounds, visual: child.visual };
      const childBounds = transformedBounds(childObject, visualTransform(child.visual));
      if (bounds === undefined) return { ...childBounds };
      const right = Math.max(bounds.x + bounds.width, childBounds.x + childBounds.width);
      const bottom = Math.max(bounds.y + bounds.height, childBounds.y + childBounds.height);
      bounds.x = Math.min(bounds.x, childBounds.x);
      bounds.y = Math.min(bounds.y, childBounds.y);
      bounds.width = right - bounds.x;
      bounds.height = bottom - bounds.y;
      return bounds;
    }, undefined as { x: number; y: number; width: number; height: number } | undefined);
    if (content === undefined || content.width <= 0 || content.height <= 0) return undefined;
    const viewbox = paint.viewboxRelative ? {
      x: content.x + paint.viewbox.x * content.width,
      y: content.y + paint.viewbox.y * content.height,
      width: paint.viewbox.width * content.width,
      height: paint.viewbox.height * content.height,
    } : paint.viewbox;
    const viewport = paint.viewportRelative ? {
      x: object.bounds.x + paint.viewport.x * object.bounds.width,
      y: object.bounds.y + paint.viewport.y * object.bounds.height,
      width: paint.viewport.width * object.bounds.width,
      height: paint.viewport.height * object.bounds.height,
    } : paint.viewport;
    if (viewbox.width <= 0 || viewbox.height <= 0 || viewport.width <= 0 || viewport.height <= 0) {
      return undefined;
    }
    const current = context.getTransform();
    const brushTransform = multiplyTransform(
      relativeBrushTransform(object.bounds, paint.relativeTransform),
      paint.transform,
    );
    const brushScale = Math.max(
      Math.hypot(brushTransform.a, brushTransform.b),
      Math.hypot(brushTransform.c, brushTransform.d),
      1,
    );
    const pixelScale = Math.max(
      Math.hypot(current.a, current.b),
      Math.hypot(current.c, current.d),
      1,
    ) * brushScale;
    const tiled = paint.tileMode !== "none";
    const sourceBounds = viewport;
    const flipX = paint.tileMode === "flip-x" || paint.tileMode === "flip-xy";
    const flipY = paint.tileMode === "flip-y" || paint.tileMode === "flip-xy";
    const baseWidth = Math.max(1, Math.ceil(viewport.width * pixelScale));
    const baseHeight = Math.max(1, Math.ceil(viewport.height * pixelScale));
    const outputWidth = baseWidth * (tiled && flipX ? 2 : 1);
    const outputHeight = baseHeight * (tiled && flipY ? 2 : 1);
    if (!Number.isSafeInteger(outputWidth)
      || !Number.isSafeInteger(outputHeight)
      || outputWidth * outputHeight > this.#limits.imagePixels) {
      throw new OfficeEngineError("IMAGE_PIXEL_LIMIT", "XPS VisualBrush tile exceeds the configured image pixel limit");
    }
    const tile = new OffscreenCanvas(baseWidth, baseHeight);
    const tileContext = tile.getContext("2d");
    if (tileContext === null) throw new OfficeEngineError("RENDER_FAILED", "Could not create an XPS VisualBrush surface");
    const rawScaleX = viewport.width / viewbox.width;
    const rawScaleY = viewport.height / viewbox.height;
    const scale = paint.stretch === "none"
      ? { x: 1, y: 1 }
      : paint.stretch === "fill"
        ? { x: rawScaleX, y: rawScaleY }
        : paint.stretch === "uniform"
          ? { x: Math.min(rawScaleX, rawScaleY), y: Math.min(rawScaleX, rawScaleY) }
          : { x: Math.max(rawScaleX, rawScaleY), y: Math.max(rawScaleX, rawScaleY) };
    const alignedX = (viewport.width - viewbox.width * scale.x) * paint.alignmentX;
    const alignedY = (viewport.height - viewbox.height * scale.y) * paint.alignmentY;
    tileContext.scale(pixelScale, pixelScale);
    tileContext.translate(alignedX, alignedY);
    tileContext.scale(scale.x, scale.y);
    tileContext.translate(-viewbox.x, -viewbox.y);
    tileContext.globalAlpha = paint.opacity;
    for (const child of paint.children) {
      await this.#drawVisualBrushChild(tileContext, object, child.bounds, child.visual, depth + 1);
    }
    const output = new OffscreenCanvas(outputWidth, outputHeight);
    const outputContext = output.getContext("2d");
    if (outputContext === null) throw new OfficeEngineError("RENDER_FAILED", "Could not create an XPS VisualBrush pattern");
    if (tiled) {
      outputContext.drawImage(tile, 0, 0);
      if (flipX) {
        outputContext.save();
        outputContext.translate(baseWidth * 2, 0);
        outputContext.scale(-1, 1);
        outputContext.drawImage(tile, 0, 0);
        outputContext.restore();
      }
      if (flipY) {
        outputContext.save();
        outputContext.translate(0, baseHeight * 2);
        outputContext.scale(1, -1);
        outputContext.drawImage(output, 0, 0, outputWidth, baseHeight, 0, 0, outputWidth, baseHeight);
        outputContext.restore();
      }
    } else {
      outputContext.drawImage(tile, 0, 0);
    }
    const pattern = context.createPattern(output, tiled ? "repeat" : "no-repeat");
    if (pattern === null) throw new OfficeEngineError("RENDER_FAILED", "Could not create an XPS VisualBrush pattern");
    const base = { a: 1 / pixelScale, b: 0, c: 0, d: 1 / pixelScale, e: sourceBounds.x, f: sourceBounds.y };
    pattern.setTransform(multiplyTransform(brushTransform, base));
    return pattern;
  }

  async #drawVisualBrushChild(
    context: OffscreenCanvasRenderingContext2D,
    owner: SceneObject,
    bounds: SceneObject["bounds"],
    visual: SceneVisual,
    depth: number,
    strokeStyle?: SceneStrokeStyle,
    textLayout?: SceneTextLayout,
    textEffects?: readonly SceneTextEffect[],
    collectTextFragment?: TextFragmentCollector,
  ): Promise<void> {
    if (depth > 64 || visual.kind === "none") return;
    if (owner.source.format === "xlsx" && owner.source.kind === "cell"
      && visual.kind === "painted-shape" && visual.geometry === "rectangle") {
      bounds = alignXlsxCellBorderToDevicePixels(context, bounds);
    }
    const object = { ...owner, bounds, visual };
    if (visual.kind === "group") {
      for (const child of visual.children) {
        await this.#drawVisualBrushChild(
          context,
          owner,
          child.bounds,
          child.visual,
          depth + 1,
          strokeStyle,
          textLayout,
          textEffects,
          collectTextFragment,
        );
      }
      return;
    }
    if (visual.kind === "opacity-mask") {
      await this.#drawOpacityMaskedVisual(context, owner, bounds, visual, depth + 1, strokeStyle, textLayout, textEffects, collectTextFragment);
      return;
    }
    if (visual.kind === "layer") {
      context.save();
      try {
        const { a, b, c, d, e, f } = visual.transform;
        context.transform(a, b, c, d, e, f);
        context.globalAlpha *= visual.opacity;
        context.globalCompositeOperation = visual.blendMode ?? "source-over";
        await this.#drawVisualBrushChild(
          context,
          owner,
          bounds,
          visual.visual,
          depth + 1,
          strokeStyle,
          textLayout,
          textEffects,
          collectTextFragment,
        );
      } finally {
        context.restore();
      }
      return;
    }
    if (visual.kind === "effect") {
      context.save();
      try {
        if (visual.shadow !== undefined) {
          context.shadowColor = color(visual.shadow.color);
          context.shadowBlur = visual.shadow.blur;
          context.shadowOffsetX = visual.shadow.offsetX;
          context.shadowOffsetY = visual.shadow.offsetY;
        }
        if (visual.clip !== undefined) context.clip(traceGeometry(context, object, visual.clip));
        await this.#drawVisualBrushChild(
          context,
          owner,
          bounds,
          visual.visual,
          depth + 1,
          strokeStyle,
          textLayout,
          textEffects,
          collectTextFragment,
        );
      } finally {
        context.restore();
      }
      return;
    }
    if (visual.kind === "stroke-style") {
      await this.#drawVisualBrushChild(
        context,
        owner,
        bounds,
        visual.visual,
        depth + 1,
        visual.style,
        textLayout,
        textEffects,
        collectTextFragment,
      );
      return;
    }
    if (visual.kind === "text-layout") {
      await this.#drawVisualBrushChild(
        context,
        owner,
        bounds,
        visual.visual,
        depth + 1,
        strokeStyle,
        visual.layout,
        textEffects,
        collectTextFragment,
      );
      return;
    }
    if (visual.kind === "text-effects") {
      await this.#drawVisualBrushChild(
        context,
        owner,
        bounds,
        visual.visual,
        depth + 1,
        strokeStyle,
        textLayout,
        visual.effects,
        collectTextFragment,
      );
      return;
    }
    if (visual.kind === "text" || visual.kind === "rich-text") {
      drawText(
        context,
        object,
        visual,
        textLayout,
        (_family, run) => this.#resolveFont(run, context).family,
        (run) => fitSubstitutedSongGlyph(context, run, this.#resolveFont(run, context)),
        undefined,
        this.#textLayouts,
        `${object.numericId}:${bounds.x}:${bounds.y}:${bounds.width}:${bounds.height}`,
        1,
        1,
        strokeStyle,
        textEffects,
        collectTextFragment,
        await this.#textPaintStyles(context, object, visual, textLayout),
      );
      return;
    }
    if (visual.kind === "image") {
      const bitmap = (await this.#image(
        visual,
        [],
        [],
        nestedVectorRasterTarget(context, bounds, visual.mediaType, this.#limits.imagePixels, visual),
      )).bitmap;
      const { cropLeft, cropTop, cropRight, cropBottom } = visual;
      drawImageRegion(context,
        bitmap,
        bitmap.width * cropLeft,
        bitmap.height * cropTop,
        bitmap.width * (1 - cropLeft - cropRight),
        bitmap.height * (1 - cropTop - cropBottom),
        bounds.x,
        bounds.y,
        bounds.width,
        bounds.height,
      );
      return;
    }
    if (visual.kind === "color-managed-image") {
      if (visual.visual.kind !== "image") {
        throw new OfficeEngineError("RENDER_FAILED", "XPS color management requires an image visual");
      }
      const bitmap = await this.#colorManagedBitmap(visual);
      const image = visual.visual;
      drawImageRegion(context,
        bitmap,
        bitmap.width * image.cropLeft,
        bitmap.height * image.cropTop,
        bitmap.width * (1 - image.cropLeft - image.cropRight),
        bitmap.height * (1 - image.cropTop - image.cropBottom),
        bounds.x, bounds.y, bounds.width, bounds.height,
      );
      bitmap.close();
      return;
    }
    if (visual.kind === "shape") {
      drawLegacyGeometry(context, object, visual, strokeStyle);
      return;
    }
    if (visual.kind !== "painted-shape") {
      throw new OfficeEngineError("RENDER_FAILED", `Unsupported XPS VisualBrush child ${visual.kind}`);
    }
    const fillImage = visual.fill.kind === "image"
      ? (await this.#paintImage(
          visual.fill,
          nestedVectorRasterTarget(
            context,
            bounds,
            visual.fill.mediaType,
            this.#limits.imagePixels,
            visual.fill,
          ),
        )).bitmap
      : undefined;
    const fillVisual = visual.fill.kind === "visual"
      ? await this.#visualBrushPattern(context, object, visual.fill, depth + 1)
      : visual.fill.kind === "xps-gradient"
        ? await this.#xpsGradientPattern(context, object, visual.fill)
      : undefined;
    const strokeVisual = visual.stroke.kind === "visual"
      ? await this.#visualBrushPattern(context, object, visual.stroke, depth + 1)
      : visual.stroke.kind === "xps-gradient"
        ? await this.#xpsGradientPattern(context, object, visual.stroke)
      : undefined;
    drawPaintedGeometry(context, object, visual, fillImage, strokeStyle, fillVisual, strokeVisual);
  }

  async #drawOpacityMaskedVisual(
    context: OffscreenCanvasRenderingContext2D,
    owner: SceneObject,
    bounds: SceneObject["bounds"],
    masked: Extract<SceneVisual, { kind: "opacity-mask" }>,
    depth: number,
    strokeStyle?: SceneStrokeStyle,
    textLayout?: SceneTextLayout,
    textEffects?: readonly SceneTextEffect[],
    collectTextFragment?: TextFragmentCollector,
  ): Promise<void> {
    // The destination-in rectangle limits the result to the mask bounds. Keep
    // the original origin: translating changes WebKit pattern rasterization.
    // Trim the right/bottom edges, retaining an antialiasing margin.
    const transform = context.getTransform();
    const deviceBounds = transformedBounds({ ...owner, bounds }, transform);
    const left = Math.max(0, Math.floor(deviceBounds.x) - 2);
    const top = Math.max(0, Math.floor(deviceBounds.y) - 2);
    const right = Math.min(context.canvas.width, Math.ceil(deviceBounds.x + deviceBounds.width) + 2);
    const bottom = Math.min(context.canvas.height, Math.ceil(deviceBounds.y + deviceBounds.height) + 2);
    if (right <= left || bottom <= top) return;
    const canvas = new OffscreenCanvas(right, bottom);
    const isolated = canvas.getContext("2d");
    if (isolated === null) throw new OfficeEngineError("RENDER_FAILED", "Could not create an XPS opacity-mask surface");
    isolated.setTransform(transform);
    await this.#drawVisualBrushChild(isolated, owner, bounds, masked.visual, depth + 1, strokeStyle, textLayout, textEffects, collectTextFragment);
    isolated.globalCompositeOperation = "destination-in";
    const maskObject = { ...owner, bounds };
    const maskVisual = {
      kind: "painted-shape" as const,
      geometry: "rectangle" as const,
      fill: masked.mask,
      stroke: { kind: "none" as const },
      strokeWidth: 0,
    };
    const maskImage = masked.mask.kind === "image" ? (await this.#paintImage(masked.mask)).bitmap : undefined;
    const maskPattern = masked.mask.kind === "visual"
      ? await this.#visualBrushPattern(isolated, maskObject, masked.mask, depth + 1)
      : masked.mask.kind === "xps-gradient"
        ? await this.#xpsGradientPattern(isolated, maskObject, masked.mask)
      : undefined;
    drawPaintedGeometry(isolated, maskObject, maskVisual, maskImage, undefined, maskPattern);
    context.save();
    try {
      context.setTransform(1, 0, 0, 1, 0, 0);
      context.drawImage(canvas, 0, 0);
    } finally {
      context.restore();
    }
  }

  async #xpsGradientPattern(
    context: OffscreenCanvasRenderingContext2D,
    object: SceneObject,
    paint: Extract<ScenePaint, { kind: "xps-gradient" }>,
  ): Promise<CanvasPattern> {
    const current = context.getTransform();
    const scale = Math.max(Math.hypot(current.a, current.b), Math.hypot(current.c, current.d), 1);
    const width = Math.max(1, Math.ceil(object.bounds.width * scale));
    const height = Math.max(1, Math.ceil(object.bounds.height * scale));
    if (width * height > this.#limits.imagePixels) {
      throw new OfficeEngineError("IMAGE_DIMENSION_LIMIT", "XPS gradient surface exceeds the configured pixel limit");
    }
    const surface = new OffscreenCanvas(width, height);
    const gradient = surface.getContext("2d");
    if (gradient === null) throw new OfficeEngineError("RENDER_FAILED", "Could not create an XPS gradient surface");
    const image = gradient.createImageData(width, height);
    const bounds = { x: 0, y: 0, width: object.bounds.width, height: object.bounds.height };
    const combined = multiplyTransform(paint.transform, relativeBrushTransform(bounds, paint.relativeTransform));
    const determinant = combined.a * combined.d - combined.b * combined.c;
    if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-12) {
      throw new OfficeEngineError("RENDER_FAILED", "XPS gradient transform is singular");
    }
    const inverse = {
      a: combined.d / determinant, b: -combined.b / determinant,
      c: -combined.c / determinant, d: combined.a / determinant,
      e: (combined.c * combined.f - combined.d * combined.e) / determinant,
      f: (combined.b * combined.e - combined.a * combined.f) / determinant,
    };
    const sx = paint.relative ? paint.start.x * bounds.width : paint.start.x;
    const sy = paint.relative ? paint.start.y * bounds.height : paint.start.y;
    const ex = paint.relative ? paint.end.x * bounds.width : paint.end.x;
    const ey = paint.relative ? paint.end.y * bounds.height : paint.end.y;
    const rx = paint.relative ? paint.radiusX * bounds.width : paint.radiusX;
    const ry = paint.relative ? paint.radiusY * bounds.height : paint.radiusY;
    const linear = (value: number) => value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
    const encoded = (value: number) => value <= 0.0031308 ? value * 12.92 : 1.055 * value ** (1 / 2.4) - 0.055;
    const spread = (value: number) => paint.spread === "pad"
      ? Math.max(0, Math.min(1, value))
      : paint.spread === "repeat"
        ? value - Math.floor(value)
        : 1 - Math.abs((value - Math.floor(value / 2) * 2) - 1);
    const stops = await Promise.all(paint.stops.map(async (stop) => {
      const color = await this.#xpsColor(stop.color);
      return {
        offset: stop.offset,
        channels: [24, 16, 8, 0].map((shift) => {
          const value = ((color >>> shift) & 0xff) / 255;
          return paint.linearRgb && shift !== 0 ? linear(value) : value;
        }),
      };
    }));
    const sample = (value: number, offset: number, extend = false) => {
      const t = extend ? 1 : spread(value);
      let right = 0;
      while (right < stops.length - 1 && !(stops[right]!.offset >= t)) right += 1;
      const left = Math.max(0, right - 1);
      const a = stops[left]!;
      const b = stops[right]!;
      const mix = a === b || b.offset === a.offset ? 0 : (t - a.offset) / (b.offset - a.offset);
      for (let channel = 0; channel < 4; channel += 1) {
        const av = a.channels[channel]!;
        const bv = b.channels[channel]!;
        const value = av + (bv - av) * mix;
        image.data[offset + channel] = Math.round(Math.max(0, Math.min(1,
          paint.linearRgb && channel !== 3 ? encoded(value) : value,
        )) * 255);
      }
    };
    for (let py = 0, offset = 0; py < height; py += 1) {
      for (let px = 0; px < width; px += 1, offset += 4) {
        const x = (px + 0.5) / scale;
        const y = (py + 0.5) / scale;
        const gx = inverse.a * x + inverse.c * y + inverse.e;
        const gy = inverse.b * x + inverse.d * y + inverse.f;
        let t: number;
        if (!paint.radial) {
          const dx = ex - sx;
          const dy = ey - sy;
          t = ((gx - sx) * dx + (gy - sy) * dy) / (dx * dx + dy * dy || 1);
        } else {
          const dx = gx - sx;
          const dy = gy - sy;
          const ox = sx - ex;
          const oy = sy - ey;
          const aa = dx * dx / (rx * rx || 1) + dy * dy / (ry * ry || 1);
          const bb = 2 * (ox * dx / (rx * rx || 1) + oy * dy / (ry * ry || 1));
          const cc = ox * ox / (rx * rx || 1) + oy * oy / (ry * ry || 1) - 1;
          const discriminant = bb * bb - 4 * aa * cc;
          if (discriminant < 0) { sample(1, offset, true); continue; }
          const root = (-bb + Math.sqrt(discriminant)) / (2 * aa || 1);
          if (root <= 0) { sample(1, offset, true); continue; }
          t = 1 / root;
        }
        sample(t, offset);
      }
    }
    gradient.putImageData(image, 0, 0);
    const pattern = context.createPattern(surface, "no-repeat");
    if (pattern === null) throw new OfficeEngineError("RENDER_FAILED", "Could not create an XPS gradient pattern");
    pattern.setTransform({ a: 1 / scale, b: 0, c: 0, d: 1 / scale, e: object.bounds.x, f: object.bounds.y });
    return pattern;
  }

  async #xpsColor(color: SceneXpsColor): Promise<number> {
    if (color.kind === "rgba") return color.color;
    const [cms, formatter, formats] = await Promise.all([
      import("@kittl/little-cms"),
      import("@kittl/little-cms/formatter"),
      import("@kittl/little-cms/formats"),
    ]);
    (await cms.initWasm(new URL("./lcms.wasm", import.meta.url).href)).valueOrThrow;
    const source = cms.cmsOpenProfileFromMem(color.profile).valueOrThrow;
    const destination = cms.cmsCreate_sRGBProfile().valueOrThrow;
    try {
      const space = cms.cmsGetColorSpace(source).valueOrThrow;
      const standard = space === cms.IccColorSpaceMap.GRAY
        ? formatter.ColorSpaceCode.PT_GRAY
        : space === cms.IccColorSpaceMap.RGB
          ? formatter.ColorSpaceCode.PT_RGB
          : space === cms.IccColorSpaceMap.CMY
            ? formatter.ColorSpaceCode.PT_CMY
            : space === cms.IccColorSpaceMap.CMYK
              ? formatter.ColorSpaceCode.PT_CMYK
              : undefined;
      const sourceSpace = standard ?? formatter.ColorSpaceCode.PT_MCH1 + color.channels.length - 1;
      let sourceFormat = formatter.iniFormatter();
      sourceFormat = formatter.setFloat(sourceFormat, true);
      sourceFormat = formatter.setBytesPerSample(sourceFormat, 4);
      sourceFormat = formatter.setChannels(sourceFormat, color.channels.length);
      sourceFormat = formatter.setColorSpace(
        sourceFormat,
        sourceSpace as Parameters<typeof formatter.setColorSpace>[1],
      );
      const transform = cms.cmsCreateTransform(
        source,
        sourceFormat,
        destination,
        formats.TYPE_RGB_8,
        cms.CmsIntent.RelativeColorimetric,
        0,
      ).valueOrThrow;
      try {
        // XPS normalizes CMYK to 0..1; LittleCMS float CMYK uses percent ink.
        const input = new Float32Array(color.channels.map(value =>
          space === cms.IccColorSpaceMap.CMYK ? value * 100 : value));
        const output = cms.cmsDoTransform(
          transform,
          new Uint8Array(input.buffer, input.byteOffset, input.byteLength),
          1,
        ).valueOrThrow;
        // The CMS wrapper allocates one byte per output channel, so request
        // RGB8 instead of a float output that would overrun that allocation.
        const alpha = Math.round(Math.max(0, Math.min(1, color.alpha)) * 255);
        return (output[0]! << 24) | (output[1]! << 16) | (output[2]! << 8) | alpha;
      } finally {
        cms.cmsDeleteTransform(transform);
      }
    } finally {
      cms.cmsCloseProfile(source);
      cms.cmsCloseProfile(destination);
    }
  }

  async #colorManagedBitmap(
    visual: Extract<SceneVisual, { kind: "color-managed-image" }>,
  ): Promise<ImageBitmap> {
    const cms = await import("@kittl/little-cms");
    if (visual.visual.kind !== "image") throw new OfficeEngineError("RENDER_FAILED", "Invalid XPS color-managed visual");
    const source = (await this.#image(visual.visual, [])).bitmap;
    const surface = new OffscreenCanvas(source.width, source.height);
    const context = surface.getContext("2d", { willReadFrequently: true });
    if (context === null) throw new OfficeEngineError("RENDER_FAILED", "Could not create an XPS ICC surface");
    context.drawImage(source, 0, 0);
    const pixels = context.getImageData(0, 0, source.width, source.height);
    (await cms.initWasm(new URL("./lcms.wasm", import.meta.url).href)).valueOrThrow;
    const sourceProfile = cms.cmsOpenProfileFromMem(visual.sourceProfile).valueOrThrow;
    const destinationProfile = visual.destinationProfile === undefined
      ? cms.cmsCreate_sRGBProfile().valueOrThrow
      : cms.cmsOpenProfileFromMem(visual.destinationProfile).valueOrThrow;
    const srgb = cms.cmsCreate_sRGBProfile().valueOrThrow;
    try {
      if (cms.cmsGetColorSpace(sourceProfile).valueOrThrow !== cms.IccColorSpaceMap.RGB
        || cms.cmsGetColorSpace(destinationProfile).valueOrThrow !== cms.IccColorSpaceMap.RGB) {
        throw new OfficeEngineError("RENDER_FAILED", "XPS ColorConvertedBitmap requires raw non-RGB channel decoding");
      }
      const first = cms.cmsCreateTransform(
        sourceProfile, LCMS_TYPE_RGBA_8, destinationProfile, LCMS_TYPE_RGBA_8,
        cms.CmsIntent.RelativeColorimetric, LCMS_FLAGS_COPY_ALPHA,
      ).valueOrThrow;
      try {
        const input = new Uint8Array(pixels.data.buffer, pixels.data.byteOffset, pixels.data.byteLength);
        const converted = cms.cmsDoTransform(first, input, source.width * source.height).valueOrThrow;
        if (visual.destinationProfile !== undefined) {
          const second = cms.cmsCreateTransform(
            destinationProfile, LCMS_TYPE_RGBA_8, srgb, LCMS_TYPE_RGBA_8,
            cms.CmsIntent.RelativeColorimetric, LCMS_FLAGS_COPY_ALPHA,
          ).valueOrThrow;
          try {
            pixels.data.set(cms.cmsDoTransform(second, converted, source.width * source.height).valueOrThrow);
          } finally { cms.cmsDeleteTransform(second); }
        } else pixels.data.set(converted);
      } finally { cms.cmsDeleteTransform(first); }
    } finally {
      cms.cmsCloseProfile(sourceProfile);
      cms.cmsCloseProfile(destinationProfile);
      cms.cmsCloseProfile(srgb);
    }
    context.putImageData(pixels, 0, 0);
    return createImageBitmap(surface);
  }

  #releaseImagePrefetch(promise: Promise<DecodedSceneImage>): void {
    const pixels = this.#imagePrefetches.get(promise);
    if (pixels === undefined) return;
    this.#imagePrefetches.delete(promise);
    this.#imagePrefetchPixels = Math.max(0, this.#imagePrefetchPixels - pixels);
    const waiters = [...this.#imagePrefetchWaiters];
    this.#imagePrefetchWaiters.clear();
    for (const resume of waiters) resume();
  }

  #paintImage(
    paint: Extract<ScenePaint, { kind: "image" }>,
    target?: ImageRasterTarget,
    adjustments: readonly SceneImageAdjustment[] = [],
  ): Promise<DecodedSceneImage> {
    if (this.#closed) {
      return Promise.reject(new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed"));
    }
    const resourceKey = this.#imageResourceKey(paint.bytes, paint.mediaType);
    const adjustmentKey = adjustments.length === 0
      ? ""
      : `:adjust:${imageAdjustmentKey(adjustments)}`;
    const key = `${target === undefined ? resourceKey : `${resourceKey}:target:${target.width}x${target.height}`}${adjustmentKey}`;
    const cached = this.#reuseImage(key);
    if (cached !== undefined) return cached;
    const decoded = this.#decodeImageCandidate(paint.bytes, paint.mediaType, key, target)
      .then((image): DecodedSceneImage => ({
        ...image,
        bitmap: applyImageAdjustments(
          image.bitmap,
          adjustments,
          this.#limits.imagePixels,
          this.#retainedImagePixels,
          this.#limits.totalImagePixels,
        ),
        mediaType: paint.mediaType,
        usedFallback: false,
      }));
    const entry: CachedSceneImage = { promise: decoded };
    this.#images.set(key, entry);
    void decoded.then((image) => {
      if (this.#closed || this.#images.get(key) !== entry) return;
      entry.image = image;
      entry.pixels = image.bitmap.width * image.bitmap.height;
      this.#retainedImagePixels += entry.pixels;
      this.#evictImages(0, key);
    }, () => {
      if (this.#images.get(key) === entry) this.#images.delete(key);
    });
    return decoded;
  }

  async #decodeImageCandidate(
    bytes: Uint8Array,
    mediaType: string,
    protectedKey: string,
    target?: ImageRasterTarget,
  ): Promise<DecodedOfficeImage> {
    const reducedJpeg2000 = mediaType.startsWith("image/jp2") || mediaType.startsWith("image/j2k");
    const imageInfo = identifyImage(
      bytes,
      mediaType,
      this.#limits.imagePixels,
      reducedJpeg2000,
    );
    const vectorTarget = RERASTERIZED_VECTOR_FORMATS.has(imageInfo.format)
      ? target
      : undefined;
    const rasterDimensions = reducedJpeg2000
      ? reducedJpeg2000Dimensions(imageInfo.width, imageInfo.height, this.#limits.imagePixels)
      : imageInfo;
    const expectedPixels = vectorTarget === undefined
      ? rasterDimensions.width * rasterDimensions.height
      : vectorTarget.width * vectorTarget.height;
    const decodeSlot = this.#acquireImageDecode(expectedPixels);
    const releaseDecodeSlot = typeof decodeSlot === "function" ? decodeSlot : await decodeSlot;
    let reservedPixels = 0;
    const releaseReservation = () => {
      if (reservedPixels === 0) return;
      this.#imagePixels = Math.max(0, this.#imagePixels - reservedPixels);
      reservedPixels = 0;
    };
    try {
      this.#assertOpen();
      this.#evictImages(expectedPixels, protectedKey);
      const totalPixels = this.#imagePixels + expectedPixels;
      if (!Number.isSafeInteger(totalPixels) || totalPixels > this.#limits.totalImagePixels) {
        throw new OfficeEngineError(
          "IMAGE_TOTAL_PIXEL_LIMIT",
          `Embedded images require ${totalPixels} decoded pixels; limit is ${this.#limits.totalImagePixels}`,
        );
      }
      this.#imagePixels = totalPixels;
      reservedPixels = expectedPixels;
      let imageDecode: Promise<DecodedOfficeImage>;
      try {
        imageDecode = decodeOfficeImage(
          bytes,
          imageInfo,
          this.#limits,
          this.#imageFonts,
          vectorTarget,
        );
      } catch (cause) {
        releaseReservation();
        throw cause;
      }
      const decoded = new Promise<DecodedOfficeImage>((resolve, reject) => {
        let settled = false;
        const timer = setTimeout(() => {
          settled = true;
          releaseReservation();
          reject(new OfficeEngineError(
            "IMAGE_DECODE_TIMEOUT",
            `Browser image decode exceeded ${IMAGE_DECODE_TIMEOUT_MS}ms`,
          ));
        }, IMAGE_DECODE_TIMEOUT_MS);
        void imageDecode.then((image) => {
          const { bitmap } = image;
          if (settled) {
            bitmap.close();
            return;
          }
          settled = true;
          clearTimeout(timer);
          if (this.#closed) {
            bitmap.close();
            releaseReservation();
            reject(new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed"));
            return;
          }
          const decodedPixels = bitmap.width * bitmap.height;
          if (!Number.isSafeInteger(decodedPixels) || decodedPixels > this.#limits.imagePixels) {
            bitmap.close();
            releaseReservation();
            reject(new OfficeEngineError(
              "IMAGE_DIMENSION_LIMIT",
              "Decoded image dimensions exceed the configured pixel limit",
            ));
            return;
          }
          if (decodedPixels > reservedPixels) {
            const additionalPixels = decodedPixels - reservedPixels;
            const decodedTotalPixels = this.#imagePixels + additionalPixels;
            if (!Number.isSafeInteger(decodedTotalPixels)
              || decodedTotalPixels > this.#limits.totalImagePixels) {
              bitmap.close();
              releaseReservation();
              reject(new OfficeEngineError(
                "IMAGE_TOTAL_PIXEL_LIMIT",
                "Decoded images exceed the configured total pixel limit",
              ));
              return;
            }
            this.#imagePixels = decodedTotalPixels;
            reservedPixels = decodedPixels;
          } else if (decodedPixels < reservedPixels) {
            this.#imagePixels -= reservedPixels - decodedPixels;
            reservedPixels = decodedPixels;
          }
          resolve(image);
        }, (cause) => {
          if (settled) return;
          settled = true;
          clearTimeout(timer);
          releaseReservation();
          reject(cause);
        });
      });
      const image = await decoded;
      if (this.#closed) {
        image.bitmap.close();
        releaseReservation();
        throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
      }
      return image;
    } finally {
      releaseDecodeSlot();
    }
  }

  #acquireImageDecode(pixels: number): (() => void) | Promise<() => void> {
    if (this.#closed) {
      return Promise.reject(new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed"));
    }
    let acquired: (() => void) | undefined;
    const pending = new Promise<() => void>((resolve, reject) => {
      this.#imageDecodeQueue.push({
        pixels,
        resolve: (release) => {
          acquired = release;
          resolve(release);
        },
        reject,
      });
      this.#drainImageDecodeQueue();
    });
    return acquired ?? pending;
  }

  #drainImageDecodeQueue(): void {
    if (this.#closed) {
      const closeError = new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
      for (const queued of this.#imageDecodeQueue.splice(0)) queued.reject(closeError);
      return;
    }
    const pixelBudget = Math.min(
      IMAGE_PREFETCH_PIXEL_BUDGET,
      Math.floor(this.#limits.totalImagePixels / 4),
    );
    while (this.#imageDecodeQueue.length > 0) {
      const queued = this.#imageDecodeQueue[0]!;
      const fits = queued.pixels > pixelBudget
        ? this.#activeImageDecodes === 0
        : this.#activeImageDecodePixels + queued.pixels <= pixelBudget;
      if (this.#activeImageDecodes >= IMAGE_PREFETCH_CONCURRENCY || !fits) return;
      this.#imageDecodeQueue.shift();
      this.#activeImageDecodes += 1;
      this.#activeImageDecodePixels += queued.pixels;
      let released = false;
      queued.resolve(() => {
        if (released) return;
        released = true;
        this.#activeImageDecodes = Math.max(0, this.#activeImageDecodes - 1);
        this.#activeImageDecodePixels = Math.max(0, this.#activeImageDecodePixels - queued.pixels);
        this.#drainImageDecodeQueue();
      });
    }
  }

  #imageResourceKey(bytes: Uint8Array, mediaType: string): string {
    let fingerprint = this.#imageFingerprints.get(bytes);
    if (fingerprint === undefined) {
      fingerprint = byteFingerprint(bytes);
      this.#imageFingerprints.set(bytes, fingerprint);
    }
    const baseKey = `${mediaType}:${fingerprint}`;
    const sources = this.#imageSources.get(baseKey);
    if (sources === undefined) {
      this.#imageSources.set(baseKey, [bytes]);
      return baseKey;
    }
    for (let index = 0; index < sources.length; index += 1) {
      const source = sources[index]!;
      if (source === bytes || bytesEqual(source, bytes)) {
        return index === 0 ? baseKey : `${baseKey}:collision:${index}`;
      }
    }
    sources.push(bytes);
    return `${baseKey}:collision:${sources.length - 1}`;
  }

  #reuseImage(key: string): Promise<DecodedSceneImage> | undefined {
    const entry = this.#images.get(key);
    if (entry === undefined) return undefined;
    this.#images.delete(key);
    this.#images.set(key, entry);
    const image = entry.image;
    if (image !== undefined && typeof ImageBitmap !== "undefined" && image.bitmap instanceof ImageBitmap
      && typeof navigator !== "undefined" && /Chrom(?:e|ium)\//u.test(navigator.userAgent)
      && !this.#imagePins.has(entry.promise) && this.#activeRenders === 1) {
      const pixels = image.bitmap.width * image.bitmap.height;
      this.#evictImages(pixels, key);
      if (this.#imagePixels + pixels > this.#limits.totalImagePixels) {
        this.#images.delete(key);
        this.#retainedImagePixels -= pixels;
        this.#imagePixels -= pixels;
        image.bitmap.close();
        return undefined;
      }
      // Chromium retains scale-dependent sampling state on a painted bitmap.
      // Clone decoded pixels, never encoded bytes, to keep cold-render sampling.
      const bitmap = structuredClone(image.bitmap);
      this.#retiredImages.add(image.bitmap);
      this.#imagePixels += pixels;
      entry.image = { ...image, bitmap };
      entry.promise = Promise.resolve(entry.image);
    }
    return entry.promise;
  }

  #evictImages(requiredPixels: number, protectedKey: string): void {
    // Frames are cheap to repaint; retain decoded sources until the document budget requires eviction.
    const neededPixels = this.#imagePixels + requiredPixels;
    for (const [key, frame] of this.#renderedFrames) {
      if (neededPixels + this.#retainedFramePixels <= this.#limits.totalImagePixels) break;
      this.#renderedFrames.delete(key);
      this.#retainedFramePixels -= frame.pixels;
      frame.bitmap.close();
    }
    const cachePixels = Math.max(0, this.#limits.totalImagePixels - this.#retainedFramePixels);
    if (this.#imagePixels + requiredPixels <= cachePixels) return;
    for (const [key, entry] of this.#images) {
      if (key === protectedKey || entry.pixels === undefined || this.#imagePins.has(entry.promise)) continue;
      this.#images.delete(key);
      this.#retainedImagePixels = Math.max(0, this.#retainedImagePixels - entry.pixels);
      this.#imagePixels = Math.max(0, this.#imagePixels - entry.pixels);
      void entry.promise.then(({ bitmap }) => bitmap.close(), () => undefined);
      if (this.#imagePixels + requiredPixels <= cachePixels) break;
    }
  }

  #renderedFramePixelBudget(): number {
    return Math.min(
      Math.floor(this.#limits.totalImagePixels / 4),
      Math.max(0, this.#limits.totalImagePixels - this.#retainedImagePixels),
      32_000_000,
    );
  }

  #evictRenderedFrames(requiredPixels: number): void {
    const budget = this.#renderedFramePixelBudget();
    for (const [key, frame] of this.#renderedFrames) {
      if (this.#retainedFramePixels + requiredPixels <= budget) break;
      this.#renderedFrames.delete(key);
      this.#retainedFramePixels -= frame.pixels;
      frame.bitmap.close();
    }
  }

  #invalidateRenderedFrames(unitIndex: number): void {
    for (const [key, frame] of this.#renderedFrames) {
      if (frame.unitIndex !== unitIndex) continue;
      this.#renderedFrames.delete(key);
      this.#retainedFramePixels -= frame.pixels;
      frame.bitmap.close();
    }
  }

  #assertOpen(): void {
    if (this.#closed) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
  }
}
