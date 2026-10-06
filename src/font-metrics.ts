import {
  compatibleFallbackFamily,
  fontFallbackTextSegments,
  fontShorthand,
  semanticSymbolFontText,
  symbolFontGlyphText,
  type FontResolution,
  type PreparedFontAsset,
  type RuntimeFontResolver,
} from "./font.js";
import { nestedFontVisual, type SceneObject, type SceneTextRun, type SceneVisual } from "./scene.js";
import type { Diagnostic, DocumentFormat, FontStretch, FontStyle } from "./types.js";
import { OfficeEngineError } from "./types.js";

const MAGIC = [0x4f, 0x56, 0x46, 0x4d] as const; // OVFM
const HEADER_BYTES = 20;
const FACE_RECORD_BYTES_V1 = 16;
const FACE_RECORD_BYTES_V2 = 28;
const FACE_RECORD_BYTES_V3 = 32;
const METRIC_RECORD_BYTES = 8;
const MEASURE_FONT_SIZE = 1_000;
const MAX_ADVANCE_EM = 1_024;
const MAX_FAMILY_UTF8_BYTES = 1_024;

export const FONT_METRIC_TABLE_MAGIC = "OVFM";
export const FONT_METRIC_TABLE_VERSION = 3;

const STYLES: readonly FontStyle[] = ["normal", "italic", "oblique"];
const STRETCHES: readonly FontStretch[] = [
  "ultra-condensed",
  "extra-condensed",
  "condensed",
  "semi-condensed",
  "normal",
  "semi-expanded",
  "expanded",
  "extra-expanded",
  "ultra-expanded",
];
const GENERIC_FAMILIES = new Set([
  "serif",
  "sans-serif",
  "monospace",
  "cursive",
  "fantasy",
  "system-ui",
  "ui-serif",
  "ui-sans-serif",
  "ui-monospace",
  "ui-rounded",
]);

export interface FontMetricFace {
  readonly family: string;
  readonly style: FontStyle;
  readonly weight: number;
  readonly stretch: FontStretch;
}

/** RegisteredFonts satisfies this interface; exact face resolvers may override resolveFace. */
export interface FontMetricResolver extends RuntimeFontResolver {
  resolveFace?(
    face: FontMetricFace & { readonly codePoints?: readonly number[] },
  ): FontMetricFace & Partial<FontResolution>;
  imageCodecFonts?(): readonly PreparedFontAsset[];
}

export interface FontMeasureContext {
  font: string;
  measureText(text: string): {
    readonly width: number;
    readonly fontBoundingBoxAscent?: number;
    readonly fontBoundingBoxDescent?: number;
  };
}

export interface FontMetricMeasureOptions {
  /** Test/host seam. The default creates an OffscreenCanvas 2D context. */
  readonly createContext?: () => FontMeasureContext | null | undefined;
  /** Maximum encoded OVFM bytes materialized for one document. */
  readonly maxBytes?: number;
}

export interface FontAdvanceMetric {
  readonly codePoint: number;
  readonly advanceEm: number;
}

export interface FontVerticalMetrics {
  readonly ascentEm: number;
  readonly descentEm: number;
  readonly lineGapEm: number;
  /** Canvas font boxes do not expose external leading. */
  readonly lineGapUnknown?: boolean;
}

export interface FontMetricFaceTable extends FontMetricFace {
  /** Sparse advances measured at this CSS size; absent means scalable metrics. */
  readonly fontSize?: number;
  readonly metrics: readonly FontAdvanceMetric[];
  readonly verticalMetrics?: FontVerticalMetrics;
  readonly fallback?: boolean;
}

export interface DecodedFontMetricTable {
  readonly version: 1 | 2 | 3;
  readonly faces: readonly FontMetricFaceTable[];
}

export interface FontMetricMeasureResult {
  readonly table: Uint8Array;
  readonly faceCount: number;
  readonly metricCount: number;
  readonly diagnostics: readonly Diagnostic[];
}

export function fontMetricLayoutObjects(
  objects: readonly SceneObject[],
  format: DocumentFormat,
): readonly SceneObject[] {
  if (format === "doc" || format === "docx" || format === "rtf" || format === "pages" || format === "odt") return objects;
  if (format !== "pptx" && format !== "odp") return [];
  const children = new Map<number, SceneObject[]>();
  const included = new Set<number>();
  const pending: SceneObject[] = [];
  for (const object of objects) {
    if (object.parentNumericId !== undefined) {
      const siblings = children.get(object.parentNumericId);
      if (siblings === undefined) children.set(object.parentNumericId, [object]);
      else siblings.push(object);
    }
    if (object.type === "table" || object.type === "cell") pending.push(object);
  }
  while (pending.length > 0) {
    const object = pending.pop()!;
    if (included.has(object.numericId)) continue;
    included.add(object.numericId);
    pending.push(...(children.get(object.numericId) ?? []));
  }
  return included.size === 0 ? [] : objects.filter((object) => included.has(object.numericId));
}

interface MutableDemand {
  readonly fontSize?: number;
  /** Scene-facing lookup key encoded into OVFM for the Rust layout engine. */
  readonly authoredFace: FontMetricFace;
  /** Resolved face used to measure each authored Unicode scalar. */
  readonly measuredFaces: Map<number, ResolvedMeasuredFace>;
}

interface DemandBudget {
  remaining: number;
  truncated: boolean;
}

interface ResolvedMeasuredFace {
  readonly face: FontMetricFace;
  readonly fallback: boolean;
  readonly spaceAdvanceEm?: number;
}

function invalidTable(message: string): never {
  throw new OfficeEngineError("INVALID_FONT_METRIC_TABLE", message);
}

function isScalarValue(value: number): boolean {
  return Number.isInteger(value)
    && value >= 0
    && value <= 0x10ffff
    && !(value >= 0xd800 && value <= 0xdfff);
}

function validateFace(face: FontMetricFace, label: string): FontMetricFace {
  if (face === null || typeof face !== "object") invalidTable(`${label} must be an object`);
  if (typeof face.family !== "string") invalidTable(`${label}.family must be a string`);
  const family = face.family.trim();
  if (family.length === 0 || family.length > 256 || /[\0-\x1f\x7f]/u.test(family)) {
    invalidTable(`${label}.family must be 1-256 printable characters`);
  }
  if (!STYLES.includes(face.style)) invalidTable(`${label}.style is invalid`);
  if (!Number.isInteger(face.weight) || face.weight < 1 || face.weight > 1_000) {
    invalidTable(`${label}.weight must be an integer from 1 through 1000`);
  }
  if (!STRETCHES.includes(face.stretch)) invalidTable(`${label}.stretch is invalid`);
  return Object.freeze({
    family,
    style: face.style,
    weight: face.weight,
    stretch: face.stretch,
  });
}

function faceKey(face: FontMetricFace): string {
  const family = face.family.trim().replace(/^(['"])(.*)\1$/u, "$2").toLowerCase();
  return [family, face.style, String(face.weight), face.stretch].join("\0");
}

function faceFamilyKey(face: Pick<FontMetricFace, "family">): string {
  return face.family.trim().replace(/^(['"])(.*)\1$/u, "$2").toLowerCase();
}

function sfntLineGapEm(bytes: ArrayBuffer): number | undefined {
  if (bytes.byteLength < 12) return undefined;
  const view = new DataView(bytes);
  const signature = view.getUint32(0);
  if (signature !== 0x0001_0000 && signature !== 0x4f54_544f) return undefined;
  const tableCount = view.getUint16(4);
  if (12 + tableCount * 16 > bytes.byteLength) return undefined;
  const tables = new Map<number, { readonly offset: number; readonly length: number }>();
  for (let index = 0; index < tableCount; index += 1) {
    const record = 12 + index * 16;
    const offset = view.getUint32(record + 8);
    const length = view.getUint32(record + 12);
    if (offset <= bytes.byteLength && length <= bytes.byteLength - offset) {
      tables.set(view.getUint32(record), { offset, length });
    }
  }
  const head = tables.get(0x6865_6164);
  if (head === undefined || head.length < 20) return undefined;
  const unitsPerEm = view.getUint16(head.offset + 18);
  if (unitsPerEm === 0) return undefined;
  const hhea = tables.get(0x6868_6561);
  const os2 = tables.get(0x4f53_2f32);
  if ((hhea?.length ?? 0) < 10 && (os2?.length ?? 0) < 74) return undefined;
  const lineGap = Math.max(
    0,
    hhea !== undefined && hhea.length >= 10 ? view.getInt16(hhea.offset + 8) : 0,
    os2 !== undefined && os2.length >= 74 ? view.getInt16(os2.offset + 72) : 0,
  ) / unitsPerEm;
  return Number.isFinite(lineGap) && lineGap <= MAX_ADVANCE_EM ? lineGap : undefined;
}

function fontLineGaps(resolver: FontMetricResolver): {
  readonly exact: ReadonlyMap<string, number>;
  readonly families: ReadonlyMap<string, number>;
} {
  const exact = new Map<string, number>();
  const families = new Map<string, number>();
  for (const asset of resolver.imageCodecFonts?.() ?? []) {
    const lineGap = sfntLineGapEm(asset.bytes);
    if (lineGap === undefined) continue;
    const resolved = resolveFace({
      family: asset.family,
      style: asset.style,
      weight: asset.weight,
      stretch: asset.stretch,
    }, resolver, []);
    const key = faceKey(resolved.face);
    const family = faceFamilyKey(resolved.face);
    exact.set(key, Math.max(exact.get(key) ?? 0, lineGap));
    families.set(family, Math.max(families.get(family) ?? 0, lineGap));
  }
  return { exact, families };
}

function compareFaces(left: FontMetricFace, right: FontMetricFace): number {
  if (left.family !== right.family) return left.family < right.family ? -1 : 1;
  const style = STYLES.indexOf(left.style) - STYLES.indexOf(right.style);
  if (style !== 0) return style;
  if (left.weight !== right.weight) return left.weight - right.weight;
  return STRETCHES.indexOf(left.stretch) - STRETCHES.indexOf(right.stretch);
}

function authoredFace(run: Pick<SceneTextRun, "fontFamily" | "bold" | "italic">): FontMetricFace {
  return {
    family: run.fontFamily,
    style: run.italic ? "italic" : "normal",
    weight: run.bold ? 700 : 400,
    stretch: "normal",
  };
}

function resolveFace(
  face: FontMetricFace,
  resolver: FontMetricResolver,
  codePoints: readonly number[],
): ResolvedMeasuredFace {
  if (resolver.resolveFace !== undefined) {
    const resolution = resolver.resolveFace({ ...face, codePoints });
    return Object.freeze({
      face: validateFace(resolution, "Resolved font face"),
      fallback: resolution.source === "fallback",
      ...(resolution.spaceAdvanceEm === undefined ? {} : { spaceAdvanceEm: resolution.spaceAdvanceEm }),
    });
  }
  const resolution = resolver.resolve(face.family);
  return Object.freeze({
    face: validateFace({ ...face, family: resolution.family }, "Resolved font face"),
    fallback: resolution.source === "fallback",
    ...(resolution.spaceAdvanceEm === undefined ? {} : { spaceAdvanceEm: resolution.spaceAdvanceEm }),
  });
}

function addText(
  demands: Map<string, MutableDemand>,
  text: string,
  authored: FontMetricFace,
  resolver: FontMetricResolver,
  budget: DemandBudget,
  format: DocumentFormat,
  fontSize?: number,
): void {
  if (text.length === 0) return;
  if (fontSize !== undefined) {
    fontSize = Math.fround(fontSize);
    if (!Number.isFinite(fontSize) || fontSize <= 0) return;
  }
  const authoredFace = validateFace(authored, "Authored font face");
  const key = `${faceKey(authoredFace)}\0${fontSize ?? 0}`;
  let demand = demands.get(key);
  if (demand === undefined) {
    demand = { authoredFace, ...(fontSize === undefined ? {} : { fontSize }), measuredFaces: new Map() };
    demands.set(key, demand);
  }
  for (const segment of fontFallbackTextSegments(text)) {
    const segmentCodePoints = [...new Set([...segment.text].map((scalar) => scalar.codePointAt(0)!))]
      .filter(isScalarValue);
    const resolved = resolveFace(authoredFace, resolver, segmentCodePoints);
    const compatibleFamily = resolved.fallback ? compatibleFallbackFamily(format, authoredFace.family) : undefined;
    const measuredFace = compatibleFamily === undefined ? resolved
      : { ...resolved, face: { ...resolved.face, family: compatibleFamily } };
    for (const codePoint of segmentCodePoints) {
      const previousFace = demand.measuredFaces.get(codePoint);
      if (previousFace !== undefined) {
        if (compareFaces(measuredFace.face, previousFace.face) < 0) {
          demand.measuredFaces.set(codePoint, measuredFace);
        }
        continue;
      }
      if (budget.remaining === 0) {
        budget.truncated = true;
        continue;
      }
      demand.measuredFaces.set(codePoint, measuredFace);
      budget.remaining -= 1;
    }
  }
}

function collectVisual(
  object: SceneObject,
  visual: SceneVisual,
  resolver: FontMetricResolver,
  demands: Map<string, MutableDemand>,
  budget: DemandBudget,
  depth: number,
  sized = false,
): void {
  if (depth > 64) return;
  if (visual.kind === "text") {
    if (object.text !== undefined) addText(demands, object.text, {
      family: visual.fontFamily,
      style: visual.italic ? "italic" : "normal",
      weight: visual.bold ? 700 : 400,
      stretch: "normal",
    }, resolver, budget, object.source.format, sized ? visual.fontSize : undefined);
    return;
  }
  if (visual.kind === "rich-text") {
    for (const run of visual.runs) addText(demands, run.text, authoredFace(run), resolver, budget, object.source.format, sized ? run.fontSize : undefined);
    return;
  }
  const nested = nestedFontVisual(visual);
  if (nested !== undefined) collectVisual(object, nested, resolver, demands, budget, depth + 1, sized);
}

function collectDemands(
  objects: readonly SceneObject[],
  resolver: FontMetricResolver,
  maximumMetrics: number,
): { readonly demands: readonly MutableDemand[]; readonly truncated: boolean } {
  const demands = new Map<string, MutableDemand>();
  const budget: DemandBudget = { remaining: maximumMetrics, truncated: false };
  for (const object of objects) collectVisual(object, object.visual, resolver, demands, budget, 0);
  for (const object of objects) collectVisual(object, object.visual, resolver, demands, budget, 0, true);
  return {
    demands: [...demands.values()].sort((left, right) =>
      (left.fontSize === undefined ? 0 : 1) - (right.fontSize === undefined ? 0 : 1)
      || compareFaces(left.authoredFace, right.authoredFace)
      || (left.fontSize ?? 0) - (right.fontSize ?? 0)),
    truncated: budget.truncated,
  };
}

function cssFont(face: FontMetricFace, size = MEASURE_FONT_SIZE): string {
  if (face.stretch === "normal"
    && (face.style === "normal" || face.style === "italic")
    && (face.weight === 400 || face.weight === 700)) {
    return fontShorthand(
      face.family,
      size,
      face.style === "italic",
      face.weight === 700,
    );
  }
  const normalized = face.family.toLowerCase();
  const family = GENERIC_FAMILIES.has(normalized) ? normalized : JSON.stringify(face.family);
  return `${face.style} ${face.weight} ${face.stretch} ${size}px ${family}`;
}

function offscreenContext(): FontMeasureContext | undefined {
  if (typeof OffscreenCanvas === "undefined") return undefined;
  try {
    return new OffscreenCanvas(1, 1).getContext("2d") ?? undefined;
  } catch {
    return undefined;
  }
}

function diagnostic(
  code: string,
  message: string,
  details: Readonly<Record<string, string | number | boolean>>,
): Diagnostic {
  return Object.freeze({
    code,
    severity: "warning",
    fidelity: "approximate",
    phase: "layout",
    message,
    details: Object.freeze({ ...details }),
  });
}

function encodeFontMetricTable(faces: readonly FontMetricFaceTable[]): Uint8Array {
  const version = faces.some(face => face.fontSize !== undefined) ? FONT_METRIC_TABLE_VERSION : 2;
  const faceRecordBytes = version === 2 ? FACE_RECORD_BYTES_V2 : FACE_RECORD_BYTES_V3;
  const encoder = new TextEncoder();
  const families = faces.map((face, index) => {
    validateFace(face, `faces[${index}]`);
    const bytes = encoder.encode(face.family);
    if (bytes.length === 0 || bytes.length > MAX_FAMILY_UTF8_BYTES) {
      return invalidTable(`faces[${index}].family UTF-8 length is invalid`);
    }
    return bytes;
  });
  const metricCount = faces.reduce((total, face) => total + face.metrics.length, 0);
  const faceBytes = families.reduce((total, family) => total + faceRecordBytes + family.length, 0);
  const metricsOffset = HEADER_BYTES + faceBytes;
  const byteLength = metricsOffset + metricCount * METRIC_RECORD_BYTES;
  if (!Number.isSafeInteger(byteLength) || byteLength > 0xffffffff || faces.length > 0xffffffff) {
    return invalidTable("Font metric table exceeds the binary format limits");
  }

  const bytes = new Uint8Array(byteLength);
  const view = new DataView(bytes.buffer);
  MAGIC.forEach((value, index) => view.setUint8(index, value));
  view.setUint16(4, version, true);
  view.setUint16(6, HEADER_BYTES, true);
  view.setUint32(8, faces.length, true);
  view.setUint32(12, metricCount, true);
  view.setUint32(16, metricsOffset, true);

  let faceOffset = HEADER_BYTES;
  let metricIndex = 0;
  for (let index = 0; index < faces.length; index += 1) {
    const face = faces[index]!;
    const family = families[index]!;
    view.setUint16(faceOffset, family.length, true);
    view.setUint8(faceOffset + 2, STYLES.indexOf(face.style));
    view.setUint8(faceOffset + 3, STRETCHES.indexOf(face.stretch));
    view.setUint16(faceOffset + 4, face.weight, true);
    view.setUint16(faceOffset + 6, (face.fallback === true ? 1 : 0) | (face.verticalMetrics?.lineGapUnknown === true ? 2 : 0), true);
    view.setUint32(faceOffset + 8, metricIndex, true);
    view.setUint32(faceOffset + 12, face.metrics.length, true);
    const vertical = face.verticalMetrics;
    view.setFloat32(faceOffset + 16, vertical?.ascentEm ?? 0, true);
    view.setFloat32(faceOffset + 20, vertical?.descentEm ?? 0, true);
    view.setFloat32(faceOffset + 24, vertical?.lineGapEm ?? 0, true);
    if (version >= 3) view.setFloat32(faceOffset + 28, face.fontSize ?? 0, true);
    bytes.set(family, faceOffset + faceRecordBytes);
    faceOffset += faceRecordBytes + family.length;

    for (const metric of face.metrics) {
      if (!isScalarValue(metric.codePoint)
        || !Number.isFinite(metric.advanceEm)
        || metric.advanceEm < 0
        || metric.advanceEm > MAX_ADVANCE_EM) {
        return invalidTable(`faces[${index}] contains an invalid advance metric`);
      }
      const metricOffset = metricsOffset + metricIndex * METRIC_RECORD_BYTES;
      view.setUint32(metricOffset, metric.codePoint, true);
      view.setFloat32(metricOffset + 4, metric.advanceEm, true);
      metricIndex += 1;
    }
  }
  return bytes;
}

/**
 * Measures scalable advances and sparse actual-size corrections. Browser font
 * quantization is not necessarily linear; vertical metrics remain scalable.
 */
export function measureSceneFontMetrics(
  objects: readonly SceneObject[],
  resolver: FontMetricResolver,
  options: FontMetricMeasureOptions = {},
): FontMetricMeasureResult {
  if (options.maxBytes !== undefined
    && (!Number.isSafeInteger(options.maxBytes) || options.maxBytes < 0)) {
    throw new OfficeEngineError("INVALID_FONT_METRIC_LIMIT", "maxBytes must be a non-negative safe integer");
  }
  const maxBytes = Math.max(HEADER_BYTES, options.maxBytes ?? Number.MAX_SAFE_INTEGER);
  const maximumDemandMetrics = options.maxBytes === undefined
    ? Number.MAX_SAFE_INTEGER
    : Math.floor(Math.max(0, options.maxBytes - HEADER_BYTES) / METRIC_RECORD_BYTES);
  const collected = collectDemands(objects, resolver, maximumDemandMetrics);
  const demands = collected.demands;
  const requestedMetricCount = demands.reduce((total, demand) => total + demand.measuredFaces.size, 0);
  if (requestedMetricCount === 0) {
    const diagnostics = collected.truncated
      ? [diagnostic(
          "FONT_METRICS_LIMIT",
          "Browser font demand exceeded the document font byte limit before measurement",
          { limit: options.maxBytes ?? maxBytes, demandTruncated: true },
        )]
      : [];
    return Object.freeze({
      table: encodeFontMetricTable([]),
      faceCount: 0,
      metricCount: 0,
      diagnostics: Object.freeze(diagnostics),
    });
  }

  let context: FontMeasureContext | null | undefined;
  try {
    context = options.createContext?.() ?? offscreenContext();
  } catch {
    context = undefined;
  }
  if (context === undefined || context === null) {
    return Object.freeze({
      table: encodeFontMetricTable([]),
      faceCount: 0,
      metricCount: 0,
      diagnostics: Object.freeze([diagnostic(
        "FONT_METRICS_UNAVAILABLE",
        "OffscreenCanvas 2D metrics are unavailable; text layout remains approximate",
        { requestedFaceCount: demands.length, requestedMetricCount },
      )]),
    });
  }

  let invalidMetricCount = 0;
  let omittedMetricCount = 0;
  let encodedBytes = HEADER_BYTES;
  let sizedTable = false;
  const encoder = new TextEncoder();
  const faces: FontMetricFaceTable[] = [];
  const scalable = new Map<string, Map<number, number>>();
  const lineGaps = fontLineGaps(resolver);
  for (const demand of demands) {
    const measureSize = demand.fontSize ?? MEASURE_FONT_SIZE;
    if (!Number.isFinite(measureSize) || measureSize <= 0) continue;
    const baseline = scalable.get(faceKey(demand.authoredFace));
    if (demand.fontSize !== undefined && baseline === undefined) continue;
    const metrics: FontAdvanceMetric[] = [];
    let ascentEm = 0;
    let descentEm = 0;
    let lineGapEm = 0;
    let lineGapUnknown = false;
    const codePoints = [...demand.measuredFaces.keys()].sort((left, right) => left - right);
    const sized = demand.fontSize !== undefined;
    const upgradeBytes = sized && !sizedTable ? faces.length * (FACE_RECORD_BYTES_V3 - FACE_RECORD_BYTES_V2) : 0;
    const faceBytes = (sized || sizedTable ? FACE_RECORD_BYTES_V3 : FACE_RECORD_BYTES_V2)
      + encoder.encode(demand.authoredFace.family).byteLength;
    let faceReserved = false;
    let activeMeasuredFace = "";
    for (let index = 0; index < codePoints.length; index += 1) {
      const codePoint = codePoints[index]!;
      if (sized && !baseline?.has(codePoint)) continue;
      const requiredBytes = (faceReserved ? 0 : faceBytes + upgradeBytes) + METRIC_RECORD_BYTES;
      if (encodedBytes + requiredBytes > maxBytes) {
        omittedMetricCount += codePoints.length - index;
        break;
      }
      const measuredFace = demand.measuredFaces.get(codePoint)!;
      const measuredLineGap = lineGaps.exact.get(faceKey(measuredFace.face))
        ?? lineGaps.families.get(faceFamilyKey(measuredFace.face));
      lineGapUnknown ||= measuredLineGap === undefined;
      lineGapEm = Math.max(lineGapEm, measuredLineGap ?? 0);
      const measuredFaceKey = faceKey(measuredFace.face);
      if (measuredFaceKey !== activeMeasuredFace) {
        context.font = cssFont(measuredFace.face, measureSize);
        activeMeasuredFace = measuredFaceKey;
      }
      let measured: ReturnType<FontMeasureContext["measureText"]>;
      try {
        const glyph = String.fromCodePoint(codePoint);
        const text = measuredFace.fallback
          ? semanticSymbolFontText(glyph, demand.authoredFace.family)
          : symbolFontGlyphText(glyph, demand.authoredFace.family, measuredFace.face, context);
        measured = context.measureText(text);
      } catch {
        invalidMetricCount += 1;
        continue;
      }
      const width = codePoint === 32 && measuredFace.spaceAdvanceEm !== undefined
        ? measuredFace.spaceAdvanceEm * measureSize : measured.width;
      const advanceEm = Math.fround(width / measureSize);
      if (!Number.isFinite(width)
        || width < 0
        || !Number.isFinite(advanceEm)
        || advanceEm < 0
        || advanceEm > MAX_ADVANCE_EM) {
        invalidMetricCount += 1;
        continue;
      }
      if (demand.fontSize !== undefined) {
        const base = baseline?.get(codePoint);
        // Ignore floating-point noise, not browser pixel quantization. Missing
        // sizes/scalars keep the scalable record and the existing format fallback.
        if (base === undefined || Math.abs(advanceEm - base) <= 0.000_001) continue;
      }
      if (!faceReserved) {
        encodedBytes += faceBytes + upgradeBytes;
        sizedTable ||= sized;
        faceReserved = true;
      }
      encodedBytes += METRIC_RECORD_BYTES;
      metrics.push(Object.freeze({ codePoint, advanceEm }));
      const measuredAscentEm = Math.fround((measured.fontBoundingBoxAscent ?? 0) / MEASURE_FONT_SIZE);
      const measuredDescentEm = Math.fround((measured.fontBoundingBoxDescent ?? 0) / MEASURE_FONT_SIZE);
      if (Number.isFinite(measuredAscentEm) && measuredAscentEm >= 0) {
        ascentEm = Math.max(ascentEm, measuredAscentEm);
      }
      if (Number.isFinite(measuredDescentEm) && measuredDescentEm >= 0) {
        descentEm = Math.max(descentEm, measuredDescentEm);
      }
    }
    if (metrics.length !== 0) {
      if (demand.fontSize === undefined) {
        scalable.set(faceKey(demand.authoredFace), new Map(metrics.map(metric => [metric.codePoint, metric.advanceEm])));
      }
      const verticalMetrics = demand.fontSize === undefined && ascentEm + descentEm > 0
        ? Object.freeze({
            ascentEm: Math.fround(ascentEm),
            descentEm: Math.fround(descentEm),
            lineGapEm: Math.fround(lineGapEm),
            ...(lineGapUnknown ? { lineGapUnknown: true } : {}),
          })
        : undefined;
      faces.push(Object.freeze({
        ...demand.authoredFace,
        ...(demand.fontSize === undefined ? {} : { fontSize: demand.fontSize }),
        metrics: Object.freeze(metrics),
        ...(verticalMetrics === undefined ? {} : { verticalMetrics }),
        ...([...demand.measuredFaces.values()].some((measuredFace) => measuredFace.fallback)
          ? { fallback: true }
          : {}),
      }));
    }
  }
  faces.sort((left, right) => compareFaces(left, right) || (left.fontSize ?? 0) - (right.fontSize ?? 0));
  const metricCount = faces.reduce((total, face) => total + face.metrics.length, 0);
  const diagnostics: Diagnostic[] = [];
  const incompleteFaceCount = faces.filter(face => face.verticalMetrics?.lineGapUnknown).length;
  if (incompleteFaceCount !== 0) {
    diagnostics.push(diagnostic(
      "FONT_METRIC_VERTICAL_APPROXIMATE",
      "Local font external leading is unavailable; natural line spacing retains the format estimate",
      { incompleteFaceCount },
    ));
  }
  if (invalidMetricCount !== 0) {
    diagnostics.push(diagnostic(
      "FONT_METRIC_INVALID",
      "Canvas returned invalid font advances; affected characters remain approximate",
      { invalidMetricCount, requestedMetricCount, metricCount },
    ));
  }
  if (omittedMetricCount !== 0 || collected.truncated) {
    diagnostics.push(diagnostic(
      "FONT_METRICS_LIMIT",
      "Browser-measured font advances were truncated to the document font byte limit",
      {
        limit: options.maxBytes ?? maxBytes,
        requestedMetricCount,
        metricCount,
        omittedMetricCount,
        demandTruncated: collected.truncated,
      },
    ));
  }
  return Object.freeze({
    table: encodeFontMetricTable(faces),
    faceCount: faces.length,
    metricCount,
    diagnostics: Object.freeze(diagnostics),
  });
}

function asBytes(source: ArrayBuffer | Uint8Array): Uint8Array {
  if (source instanceof Uint8Array) {
    return new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  }
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  return invalidTable("Font metric table must be an ArrayBuffer or Uint8Array");
}

/** Decodes and validates every byte; trailing, non-canonical, and unsafe records are rejected. */
export function decodeFontMetricTable(source: ArrayBuffer | Uint8Array): DecodedFontMetricTable {
  const bytes = asBytes(source);
  if (bytes.byteLength < HEADER_BYTES) return invalidTable("Font metric table header is truncated");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let index = 0; index < MAGIC.length; index += 1) {
    if (view.getUint8(index) !== MAGIC[index]) return invalidTable("Font metric table magic is invalid");
  }
  const version = view.getUint16(4, true);
  if (version !== 1 && version !== 2 && version !== FONT_METRIC_TABLE_VERSION) {
    return invalidTable(`Unsupported font metric table version ${version}`);
  }
  const faceRecordBytes = version === 1 ? FACE_RECORD_BYTES_V1 : version === 2 ? FACE_RECORD_BYTES_V2 : FACE_RECORD_BYTES_V3;
  if (view.getUint16(6, true) !== HEADER_BYTES) return invalidTable("Font metric table header size is invalid");
  const faceCount = view.getUint32(8, true);
  const metricCount = view.getUint32(12, true);
  const metricsOffset = view.getUint32(16, true);
  if (metricsOffset < HEADER_BYTES || metricsOffset > bytes.byteLength) {
    return invalidTable("Font metric table metric offset is invalid");
  }
  const expectedBytes = metricsOffset + metricCount * METRIC_RECORD_BYTES;
  if (!Number.isSafeInteger(expectedBytes) || expectedBytes !== bytes.byteLength) {
    return invalidTable("Font metric table length is invalid");
  }
  if (faceCount > Math.floor((metricsOffset - HEADER_BYTES) / faceRecordBytes)) {
    return invalidTable("Font metric table face count is invalid");
  }

  const decoder = new TextDecoder("utf-8", { fatal: true });
  const faces: FontMetricFaceTable[] = [];
  let faceOffset = HEADER_BYTES;
  let expectedMetricStart = 0;
  let previousFace: (FontMetricFace & { fontSize?: number }) | undefined;
  for (let index = 0; index < faceCount; index += 1) {
    if (faceOffset + faceRecordBytes > metricsOffset) return invalidTable("Font metric face is truncated");
    const familyLength = view.getUint16(faceOffset, true);
    const style = STYLES[view.getUint8(faceOffset + 2)];
    const stretch = STRETCHES[view.getUint8(faceOffset + 3)];
    const weight = view.getUint16(faceOffset + 4, true);
    const reserved = view.getUint16(faceOffset + 6, true);
    const metricStart = view.getUint32(faceOffset + 8, true);
    const faceMetricCount = view.getUint32(faceOffset + 12, true);
    if (familyLength === 0 || familyLength > MAX_FAMILY_UTF8_BYTES) return invalidTable("Font metric family length is invalid");
    if (style === undefined
      || stretch === undefined
      || (version === 1 ? reserved !== 0 : (reserved & ~3) !== 0)) {
      return invalidTable("Font metric face descriptor is invalid");
    }
    if (faceOffset + faceRecordBytes + familyLength > metricsOffset) return invalidTable("Font metric family is truncated");
    if (metricStart !== expectedMetricStart || metricStart + faceMetricCount > metricCount) {
      return invalidTable("Font metric ranges are not contiguous");
    }
    let family: string;
    try {
      family = decoder.decode(bytes.subarray(
        faceOffset + faceRecordBytes,
        faceOffset + faceRecordBytes + familyLength,
      ));
    } catch {
      return invalidTable("Font metric family is not valid UTF-8");
    }
    const fontSize = version >= 3 ? view.getFloat32(faceOffset + 28, true) : 0;
    if (!Number.isFinite(fontSize) || fontSize < 0) return invalidTable("Font metric size is invalid");
    const face = { ...validateFace({ family, style, weight, stretch }, `faces[${index}]`), ...(fontSize === 0 ? {} : { fontSize }) };
    if (previousFace !== undefined && (compareFaces(previousFace, face) || (previousFace.fontSize ?? 0) - fontSize) >= 0) {
      return invalidTable("Font metric faces are not uniquely sorted");
    }

    const metrics: FontAdvanceMetric[] = [];
    let previousCodePoint = -1;
    for (let relative = 0; relative < faceMetricCount; relative += 1) {
      const metricIndex = metricStart + relative;
      const offset = metricsOffset + metricIndex * METRIC_RECORD_BYTES;
      const codePoint = view.getUint32(offset, true);
      const advanceEm = view.getFloat32(offset + 4, true);
      if (!isScalarValue(codePoint)
        || codePoint <= previousCodePoint
        || !Number.isFinite(advanceEm)
        || advanceEm < 0
        || advanceEm > MAX_ADVANCE_EM) {
        return invalidTable("Font metric record is invalid or not uniquely sorted");
      }
      metrics.push(Object.freeze({ codePoint, advanceEm }));
      previousCodePoint = codePoint;
    }
    const verticalMetrics = version >= 2
      ? {
          ascentEm: view.getFloat32(faceOffset + 16, true),
          descentEm: view.getFloat32(faceOffset + 20, true),
          lineGapEm: view.getFloat32(faceOffset + 24, true),
          ...((reserved & 2) !== 0 ? { lineGapUnknown: true } : {}),
        }
      : undefined;
    if (verticalMetrics !== undefined
      && (!Number.isFinite(verticalMetrics.ascentEm) || verticalMetrics.ascentEm < 0
        || !Number.isFinite(verticalMetrics.descentEm) || verticalMetrics.descentEm < 0
        || !Number.isFinite(verticalMetrics.lineGapEm) || verticalMetrics.lineGapEm < 0
        || verticalMetrics.ascentEm + verticalMetrics.descentEm + verticalMetrics.lineGapEm > MAX_ADVANCE_EM)) {
      return invalidTable("Font vertical metrics are invalid");
    }
    faces.push(Object.freeze({
      ...face,
      metrics: Object.freeze(metrics),
      ...(verticalMetrics === undefined
        || verticalMetrics.ascentEm + verticalMetrics.descentEm + verticalMetrics.lineGapEm === 0
        ? {}
        : { verticalMetrics: Object.freeze(verticalMetrics) }),
      ...(version >= 2 && (reserved & 1) !== 0 ? { fallback: true } : {}),
    }));
    previousFace = face;
    expectedMetricStart += faceMetricCount;
    faceOffset += faceRecordBytes + familyLength;
  }
  if (faceOffset !== metricsOffset || expectedMetricStart !== metricCount) {
    return invalidTable("Font metric table sections do not exactly match the header");
  }
  return Object.freeze({ version: version as 1 | 2 | 3, faces: Object.freeze(faces) });
}
