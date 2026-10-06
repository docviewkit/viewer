import type {
  Diagnostic,
  DocumentFontRun,
  DocumentFormat,
  FontAsset,
  FontPolicy,
  FontProvider,
  FontProviderDiagnostic,
  FontProviderResult,
  FontRequest,
  FontStretch,
  FontStyle,
} from "./types.js";
import { eotToTtf, parseEotMetadata } from "./mtx-decompressor.js";
import { immutableDiagnostics, OfficeEngineError } from "./types.js";
import { copyByteSource, isArrayBuffer, isUint8Array } from "./bytes.js";
import { nestedFontVisual, type SceneDocument, type SceneEmbeddedFont, type SceneObject, type SceneVisual } from "./scene.js";
import { semanticSymbolFontText as semanticSymbolFontTextBase } from "./symbol-font-mappings.js";

/** Base size of a mixed text box, weighted by authored character coverage. */
export function dominantDocumentFontSize(runs: readonly DocumentFontRun[]): number | undefined {
  const coverage = new Map<number, number>();
  for (const { fontSize, start, end } of runs) {
    if (fontSize === undefined || !Number.isFinite(fontSize) || fontSize <= 0
      || !Number.isFinite(start) || !Number.isFinite(end) || end <= start) continue;
    coverage.set(fontSize, (coverage.get(fontSize) ?? 0) + end - start);
  }
  let dominant: number | undefined;
  let maximum = 0;
  for (const [size, count] of coverage) {
    if (count > maximum) {
      dominant = size;
      maximum = count;
    }
  }
  return dominant;
}

export function semanticSymbolFontText(text: string, fontFamily: string): string {
  const sourceFamily = pdfObjectSourceFamily(fontFamily) ?? fontFamily;
  const normalizedFamily = sourceFamily.trim().replace(/^(['"])(.*)\1$/u, "$2");
  const semanticFamily = /^wingdings-regular$/iu.test(normalizedFamily)
    ? "Wingdings"
    : /^symbolmt$/iu.test(normalizedFamily) ? "Symbol" : sourceFamily;
  if (semanticFamily.trim().replace(/^(['"])(.*)\1$/u, "$2").toLowerCase() === "wingdings") {
    let normalized = "";
    for (const character of text) {
      const codePoint = character.codePointAt(0)!;
      const source = codePoint <= 0xff
        ? codePoint
        : codePoint >= 0xf000 && codePoint <= 0xf0ff ? codePoint & 0xff : undefined;
      normalized += source === 0x6c
        ? "●"
        : source === 0x70 ? "□" : source === 0xd8 ? "➢"
          : source === 0xdf ? "←" : source === 0xe0 ? "→"
            : semanticSymbolFontTextBase(character, "Wingdings");
    }
    return normalized;
  }
  if (semanticFamily.trim().replace(/^(['"])(.*)\1$/u, "$2").toLowerCase() === "webdings") {
    return [...text].map((character) => {
      const codePoint = character.codePointAt(0)!;
      const source = codePoint <= 0xff
        ? codePoint
        : codePoint >= 0xf000 && codePoint <= 0xf0ff ? codePoint & 0xff : undefined;
      return source === 0x4e ? "👁" : character;
    }).join("");
  }
  return semanticSymbolFontTextBase(text, semanticFamily);
}

const DEFAULT_FONT_LOAD_TIMEOUT_MS = 10_000;
const MAX_FONT_REQUEST_CODE_POINTS = 65_536;
const MAX_FONT_PROBE_CODE_POINTS = 1_024;
const MAX_PROVIDER_FACES_PER_CALL = 128;
const MAX_FONT_LOAD_DIAGNOSTICS = 128;
const FONT_LOAD_DIAGNOSTICS_TRUNCATED_KEY = "\0font-load-diagnostics-truncated";
const REGISTERED_FONT_DIAGNOSTICS = new WeakMap<object, Map<string, Diagnostic>>();
const TEXT_DOCUMENT_FALLBACKS = new Map([
  ["liberation sans", "Arial"],
  ["dejavu sans", "Verdana"],
]);
const LEGACY_DOC_SHUSONG_FALLBACKS = new Map([
  ["汉仪书宋二kw", "Hiragino Mincho ProN"],
  ["shus-sc", "Hiragino Mincho ProN"],
  ["书宋-简", "Hiragino Mincho ProN"],
]);
const PRESENTATION_CJK_FALLBACKS = new Map([
  ["等线 light", "PingFang SC Light"],
  ["dengxian light", "PingFang SC Light"],
  ["等线", "PingFang SC Light"],
  ["dengxian", "PingFang SC Light"],
]);
const PRESENTATION_DENGXIAN_LIGHT_SIZE_FACTOR = 0.955;
const PRESENTATION_DENGXIAN_LIGHT_BASELINE_SHIFT_FACTOR = 0.09;
const FONT_STYLES = new Set<FontStyle>(["normal", "italic", "oblique"]);
const FONT_STRETCHES = new Set<FontStretch>([
  "ultra-condensed",
  "extra-condensed",
  "condensed",
  "semi-condensed",
  "normal",
  "semi-expanded",
  "expanded",
  "extra-expanded",
  "ultra-expanded",
]);

export interface PreparedFontAsset {
  readonly family: string;
  readonly bytes: ArrayBuffer;
  readonly sha256?: string;
  readonly style: FontStyle;
  readonly weight: number;
  readonly stretch: FontStretch;
}

export interface RuntimeFontFace {
  readonly family?: string;
  load(): Promise<unknown>;
}

export interface RuntimeFontSet {
  add(face: RuntimeFontFace): unknown;
  delete(face: RuntimeFontFace): boolean;
  check(font: string, text?: string): boolean;
}

export interface FontRuntime {
  readonly FontFace: new (
    family: string,
    source: ArrayBuffer | string,
    descriptors: { readonly style: string; readonly weight: string; readonly stretch: string },
  ) => RuntimeFontFace;
  readonly fontSet: RuntimeFontSet;
}

export interface FontProbeContext {
  font: string;
  measureText(text: string): {
    readonly width: number;
    readonly actualBoundingBoxLeft?: number;
    readonly actualBoundingBoxRight?: number;
    readonly actualBoundingBoxAscent?: number;
    readonly actualBoundingBoxDescent?: number;
    readonly fontBoundingBoxAscent?: number;
    readonly fontBoundingBoxDescent?: number;
  };
}

const FONT_PROBE_BASELINES = ["monospace", "serif", "sans-serif"] as const;
const FONT_PROBE_TEXT = Object.freeze([
  "BESbswy 0123456789 MWil1",
  "漢字かなカナ한글",
  "БГДЖЙΩψ",
  "العربيةעברית",
  "देवनागरीไทย",
  "✓★→∑∞♥♣\ue000\uf021",
]);
const FONT_PROBE_METRIC_KEYS = [
  "width",
  "actualBoundingBoxLeft",
  "actualBoundingBoxRight",
  "actualBoundingBoxAscent",
  "actualBoundingBoxDescent",
  "fontBoundingBoxAscent",
  "fontBoundingBoxDescent",
] as const;

function ambientFontProbeContext(): FontProbeContext | undefined {
  if (typeof OffscreenCanvas === "undefined") return undefined;
  try {
    return new OffscreenCanvas(1, 1).getContext("2d") ?? undefined;
  } catch {
    return undefined;
  }
}

function probeFontShorthand(
  family: string,
  fallback: string,
  face: Pick<FontFaceDescriptor, "style" | "weight">,
): string {
  return `${exactProbeFontShorthand(family, face)}, ${fallback}`;
}

function exactProbeFontShorthand(
  family: string,
  face: Pick<FontFaceDescriptor, "style" | "weight">,
): string {
  return `${face.style === "normal" ? "" : `${face.style} `}${face.weight} 72px ${JSON.stringify(family)}`;
}

function fallbackProbeShorthand(
  fallback: string,
  face: Pick<FontFaceDescriptor, "style" | "weight">,
): string {
  return `${face.style === "normal" ? "" : `${face.style} `}${face.weight} 72px ${fallback}`;
}

function metricDiffers(
  left: ReturnType<FontProbeContext["measureText"]>,
  right: ReturnType<FontProbeContext["measureText"]>,
): boolean {
  for (const key of FONT_PROBE_METRIC_KEYS) {
    const leftValue = left[key];
    const rightValue = right[key];
    if (typeof leftValue !== "number" || typeof rightValue !== "number") continue;
    if (!Number.isFinite(leftValue) || !Number.isFinite(rightValue)) continue;
    if (Math.abs(leftValue - rightValue) > 0.01) return true;
  }
  return false;
}

/** Use a font's Unicode cmap only when its legacy glyph is the missing-glyph box. */
export function symbolFontGlyphText(
  text: string,
  authoredFamily: string,
  face: FontFaceDescriptor,
  context: FontProbeContext | undefined,
): string {
  if (context === undefined || semanticSymbolFontText(text, authoredFamily) === text) return text;
  const previousFont = context.font;
  try {
    context.font = exactProbeFontShorthand(face.family, face);
    // ponytail: metric-only coverage probe; use cmap coverage when font bytes are available.
    const missing = context.measureText("\uffff");
    return [...text].map((glyph) => {
      const semantic = semanticSymbolFontText(glyph, authoredFamily);
      return semantic !== glyph
          && !metricDiffers(context.measureText(glyph), missing)
          && metricDiffers(context.measureText(semantic), missing)
        ? semantic
        : glyph;
    }).join("");
  } catch {
    return text;
  } finally {
    context.font = previousFont;
  }
}

function fontFaceMetricsMatch(
  leftFamily: string,
  rightFamily: string,
  face: Pick<FontFaceDescriptor, "style" | "weight">,
  context: FontProbeContext | undefined,
  text: string,
): boolean | undefined {
  if (context === undefined) return undefined;
  const previousFont = context.font;
  try {
    const probeTexts = text.length === 0 ? FONT_PROBE_TEXT : [text];
    for (const probeText of probeTexts) {
      context.font = exactProbeFontShorthand(leftFamily, face);
      const left = context.measureText(probeText);
      context.font = exactProbeFontShorthand(rightFamily, face);
      const right = context.measureText(probeText);
      if (metricDiffers(left, right)) return false;
    }
    return true;
  } catch {
    return undefined;
  } finally {
    context.font = previousFont;
  }
}

/**
 * Distinguishes an exact browser face from CSS fallback selection. FontFaceSet
 * and FontFace.load() may both succeed when the browser silently substitutes a
 * fallback; identical metrics against three unrelated generic baselines are a
 * strong negative signal. An unavailable Canvas leaves the result unknown so
 * callers can preserve the usable-font path instead of rejecting real faces.
 */
export function probeBrowserFontFace(
  face: FontFaceResolutionRequest,
  context: FontProbeContext | undefined = ambientFontProbeContext(),
  requestText?: string,
): boolean | undefined {
  if (context === undefined) return undefined;
  const normalizedFamily = normalizeFamily(face.family);
  if (GENERIC_FAMILIES.has(normalizedFamily)) return true;
  const previousFont = context.font;
  try {
    const requestedProbeText = boundedFontProbeText(face.codePoints, requestText);
    const probeTexts = requestedProbeText.length === 0 ? FONT_PROBE_TEXT : [requestedProbeText];
    for (const fallback of FONT_PROBE_BASELINES) {
      for (const text of probeTexts) {
        context.font = fallbackProbeShorthand(fallback, face);
        const baseline = context.measureText(text);
        context.font = probeFontShorthand(face.family, fallback, face);
        const candidate = context.measureText(text);
        if (metricDiffers(candidate, baseline)) return true;
      }
    }
    return false;
  } catch {
    return undefined;
  } finally {
    context.font = previousFont;
  }
}

function boundedFontProbeText(
  codePoints: readonly number[] | undefined,
  requestText?: string,
): string {
  const selected: number[] = [];
  const seen = new Set<number>();
  const add = (codePoint: number): void => {
    if (selected.length >= MAX_FONT_PROBE_CODE_POINTS
      || seen.has(codePoint)
      || !Number.isInteger(codePoint)
      || codePoint < 0
      || codePoint > 0x10ffff
      || codePoint >= 0xd800 && codePoint <= 0xdfff) return;
    seen.add(codePoint);
    selected.push(codePoint);
  };
  if (requestText !== undefined) {
    const requestLimit = codePoints === undefined
      ? MAX_FONT_PROBE_CODE_POINTS
      : Math.floor(MAX_FONT_PROBE_CODE_POINTS / 2);
    for (const character of requestText) {
      if (selected.length === requestLimit) break;
      add(character.codePointAt(0)!);
    }
  }
  if (codePoints !== undefined && selected.length < MAX_FONT_PROBE_CODE_POINTS) {
    if (codePoints.length <= MAX_FONT_PROBE_CODE_POINTS - selected.length) {
      for (const codePoint of codePoints) add(codePoint);
    } else {
      const remaining = MAX_FONT_PROBE_CODE_POINTS - selected.length;
      for (let index = 0; index < remaining; index += 1) {
        const sourceIndex = remaining === 1
          ? 0
          : Math.round(index * (codePoints.length - 1) / (remaining - 1));
        add(codePoints[sourceIndex]!);
      }
    }
  }
  return String.fromCodePoint(...selected);
}

interface FontCandidate extends Omit<PreparedFontAsset, "bytes"> {
  readonly source: ArrayBuffer | Uint8Array;
}

function invalid(message: string): never {
  throw new OfficeEngineError("INVALID_FONT_ASSET", message);
}

function validateFamily(value: unknown, index: number): string {
  if (typeof value !== "string") invalid(`fonts[${index}].family must be a string`);
  const family = value.trim();
  if (family.length === 0 || family.length > 256 || /[\0-\x1f\x7f]/u.test(family)) {
    invalid(`fonts[${index}].family must be 1-256 printable characters`);
  }
  return family;
}

function validateSource(value: unknown, index: number): ArrayBuffer | Uint8Array {
  if (isArrayBuffer(value)) {
    if (value.byteLength === 0) invalid(`fonts[${index}].bytes must not be empty`);
    return value;
  }
  if (isUint8Array(value)) {
    if (value.byteLength === 0) invalid(`fonts[${index}].bytes must not be empty`);
    return value;
  }
  return invalid(`fonts[${index}].bytes must be an ArrayBuffer or Uint8Array`);
}

function validateAssetSha256(value: unknown, index: number): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || !/^[a-f0-9]{64}$/iu.test(value)) {
    invalid(`fonts[${index}].sha256 must be a 64-character hexadecimal digest`);
  }
  return value.toLowerCase();
}

function copySource(source: ArrayBuffer | Uint8Array): ArrayBuffer {
  try {
    return copyByteSource(source);
  } catch (cause) {
    throw new OfficeEngineError("FONT_COPY_FAILED", "Could not copy a host-provided font asset", { cause });
  }
}

/** Validates the complete byte budget before copying any font data. */
export function prepareFontAssets(
  assets: readonly FontAsset[] | undefined,
  byteLimit: number,
): readonly PreparedFontAsset[] {
  if (assets === undefined) return Object.freeze([]);
  if (!Array.isArray(assets)) invalid("fonts must be an array");

  let totalBytes = 0;
  const candidates: FontCandidate[] = [];
  for (let index = 0; index < assets.length; index += 1) {
    const asset = assets[index];
    if (asset === null || typeof asset !== "object") invalid(`fonts[${index}] must be an object`);
    const source = validateSource(asset.bytes, index);
    const style = asset.style ?? "normal";
    if (!FONT_STYLES.has(style)) invalid(`fonts[${index}].style is invalid`);
    const weight = asset.weight ?? 400;
    if (!Number.isInteger(weight) || weight < 1 || weight > 1000) {
      invalid(`fonts[${index}].weight must be an integer from 1 through 1000`);
    }
    const stretch = asset.stretch ?? "normal";
    if (!FONT_STRETCHES.has(stretch)) invalid(`fonts[${index}].stretch is invalid`);
    totalBytes += source.byteLength;
    if (!Number.isSafeInteger(totalBytes) || totalBytes > byteLimit) {
      throw new OfficeEngineError(
        "FONT_BYTES_LIMIT",
        `Host fonts require ${totalBytes} bytes; limit is ${byteLimit}`,
      );
    }
    const sha256 = validateAssetSha256(asset.sha256, index);
    candidates.push({
      family: validateFamily(asset.family, index),
      source,
      ...(sha256 === undefined ? {} : { sha256 }),
      style,
      weight,
      stretch,
    });
  }

  return Object.freeze(candidates.map((candidate) => Object.freeze({
    family: candidate.family,
    bytes: copySource(candidate.source),
    ...(candidate.sha256 === undefined ? {} : { sha256: candidate.sha256 }),
    style: candidate.style,
    weight: candidate.weight,
    stretch: candidate.stretch,
  })));
}

/** Creates per-document copies; Worker callers may transfer every returned buffer. */
export function copyPreparedFonts(assets: readonly PreparedFontAsset[]): PreparedFontAsset[] {
  return assets.map((asset) => ({ ...asset, bytes: asset.bytes.slice(0) }));
}

export function prepareEmbeddedFontAssets(
  assets: readonly SceneEmbeddedFont[],
): readonly PreparedFontAsset[] {
  return Object.freeze(assets.map((asset) => Object.freeze({
    family: asset.family,
    bytes: browserFontBytes(asset.bytes),
    style: asset.style,
    weight: asset.weight,
    stretch: "normal" as const,
  })));
}

function browserFontBytes(source: Uint8Array): ArrayBuffer {
  const copied = copySource(source);
  const bytes = new Uint8Array(copied);
  try {
    parseEotMetadata(bytes);
    const decompressed = normalizeBrowserSfnt(eotToTtf(bytes));
    return decompressed.buffer.slice(
      decompressed.byteOffset,
      decompressed.byteOffset + decompressed.byteLength,
    ) as ArrayBuffer;
  } catch {
    return normalizeBrowserSfnt(bytes).buffer as ArrayBuffer;
  }
}

function normalizeBrowserSfnt(font: Uint8Array): Uint8Array {
  if (font.byteLength < 12
    || ![0x00010000, 0x4f54544f, 0x74727565].includes(readU32BigEndian(font, 0))) return font;
  const tableCount = readU16BigEndian(font, 4);
  if (tableCount > 4096 || 12 + tableCount * 16 > font.byteLength) return font;
  const records = Array.from({ length: tableCount }, (_, index) => font.slice(12 + index * 16, 28 + index * 16));
  const lowercaseOs2 = records.filter((record) => readU32BigEndian(record, 0) === 0x6f732f32);
  let changed = false;
  if (lowercaseOs2.length === 1 && !records.some((record) => readU32BigEndian(record, 0) === 0x4f532f32)) {
    // Some embedded subsets use "os/2"; browser sanitizers require the exact SFNT tag.
    writeU32BigEndian(lowercaseOs2[0]!, 0, 0x4f532f32);
    changed = true;
  }
  if (changed || records.some((record, index) => index > 0
    && readU32BigEndian(records[index - 1]!, 0) > readU32BigEndian(record, 0))) {
    records.sort((left, right) => readU32BigEndian(left, 0) - readU32BigEndian(right, 0));
    records.forEach((record, index) => font.set(record, 12 + index * 16));
    changed = true;
  }
  let cmapRecord = -1;
  let headRecord = -1;
  let postRecord = -1;
  let maxpRecord = -1;
  for (let index = 0; index < tableCount; index += 1) {
    const record = 12 + index * 16;
    const tag = String.fromCharCode(font[record]!, font[record + 1]!, font[record + 2]!, font[record + 3]!);
    if (tag === "cmap") cmapRecord = record;
    else if (tag === "head") headRecord = record;
    else if (tag === "post") postRecord = record;
    else if (tag === "maxp") maxpRecord = record;
  }
  if (cmapRecord < 0 || headRecord < 0) return font;
  const cmapOffset = readU32BigEndian(font, cmapRecord + 8);
  const cmapLength = readU32BigEndian(font, cmapRecord + 12);
  const headOffset = readU32BigEndian(font, headRecord + 8);
  const headLength = readU32BigEndian(font, headRecord + 12);
  if (cmapLength < 4
    || cmapOffset + cmapLength > font.byteLength
    || headLength < 12
    || headOffset + headLength > font.byteLength) return font;
  const encodingCount = readU16BigEndian(font, cmapOffset + 2);
  if (cmapOffset + 4 + encodingCount * 8 > cmapOffset + cmapLength) return font;
  if (postRecord >= 0 && maxpRecord >= 0) {
    const postOffset = readU32BigEndian(font, postRecord + 8);
    const postLength = readU32BigEndian(font, postRecord + 12);
    const maxpOffset = readU32BigEndian(font, maxpRecord + 8);
    const maxpLength = readU32BigEndian(font, maxpRecord + 12);
    if (postLength >= 34 && postOffset + postLength <= font.byteLength
      && maxpLength >= 6 && maxpOffset + maxpLength <= font.byteLength
      && readU32BigEndian(font, postOffset) === 0x00020000
      && readU16BigEndian(font, postOffset + 32) !== readU16BigEndian(font, maxpOffset + 4)) {
      // Subsets may retain the original post names. Version 3 omits only names;
      // browser glyph selection still uses the unchanged cmap and outlines.
      writeU32BigEndian(font, postOffset, 0x00030000);
      writeU32BigEndian(font, postRecord + 12, 32);
      writeU32BigEndian(font, postRecord + 4, sfntChecksum(font, postOffset, 32));
      changed = true;
    }
  }
  for (let index = 0; index < encodingCount; index += 1) {
    const encoding = cmapOffset + 4 + index * 8;
    const subtableOffset = readU32BigEndian(font, encoding + 4);
    const subtable = cmapOffset + subtableOffset;
    if (subtable + 12 > cmapOffset + cmapLength) continue;
    const format = readU16BigEndian(font, subtable);
    if (format === 4) {
      const length = readU16BigEndian(font, subtable + 2);
      const segmentCountX2 = readU16BigEndian(font, subtable + 6);
      const segments = segmentCountX2 / 2;
      if (segments > 0 && Number.isInteger(segments) && length >= 16 + segments * 8
        && subtable + length <= cmapOffset + cmapLength) {
        const end = subtable + 14 + (segments - 1) * 2;
        const start = end + segments * 2 + 2;
        const delta = start + segments * 2;
        const range = delta + segments * 2;
        if (readU16BigEndian(font, end) === 0xffff && readU16BigEndian(font, start) === 0xffff
          && readU16BigEndian(font, delta) === 0 && readU16BigEndian(font, range) === 0) {
          // The format-4 terminator must map U+FFFF to glyph zero, not glyph 65535.
          font[delta + 1] = 1;
          changed = true;
        }
      }
    }
    if ((format === 12 || format === 13) && readU32BigEndian(font, subtable + 8) !== 0) {
      writeU32BigEndian(font, subtable + 8, 0);
      changed = true;
    }
  }
  if (!changed) return font;
  writeU32BigEndian(font, cmapRecord + 4, sfntChecksum(font, cmapOffset, cmapLength));
  writeU32BigEndian(font, headOffset + 8, 0);
  const adjustment = (0xb1b0afba - sfntChecksum(font, 0, font.byteLength)) >>> 0;
  writeU32BigEndian(font, headOffset + 8, adjustment);
  return font;
}

function readU16BigEndian(bytes: Uint8Array, offset: number): number {
  return (bytes[offset]! << 8) | bytes[offset + 1]!;
}

function readU32BigEndian(bytes: Uint8Array, offset: number): number {
  return (
    bytes[offset]! * 0x1000000
    + bytes[offset + 1]! * 0x10000
    + bytes[offset + 2]! * 0x100
    + bytes[offset + 3]!
  ) >>> 0;
}

function writeU32BigEndian(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = value >>> 24;
  bytes[offset + 1] = value >>> 16;
  bytes[offset + 2] = value >>> 8;
  bytes[offset + 3] = value;
}

function sfntChecksum(bytes: Uint8Array, offset: number, length: number): number {
  let checksum = 0;
  const end = offset + length;
  for (let cursor = offset; cursor < end; cursor += 4) {
    let word = 0;
    for (let index = 0; index < 4; index += 1) {
      word = (word << 8) | (cursor + index < end ? bytes[cursor + index]! : 0);
    }
    checksum = (checksum + (word >>> 0)) >>> 0;
  }
  return checksum;
}

export function fontAssetBytes(assets: readonly PreparedFontAsset[]): number {
  return assets.reduce((total, asset) => total + asset.bytes.byteLength, 0);
}

export interface FontProviderCacheOptions {
  readonly maxBytes: number;
  readonly policy?: FontPolicy | undefined;
  readonly timeoutMs?: number | undefined;
}

function invalidRequest(message: string): never {
  throw new OfficeEngineError("INVALID_FONT_REQUEST", message);
}

function normalizeLanguage(language: string | undefined): string | undefined {
  if (language === undefined) return undefined;
  if (typeof language !== "string" || language.trim().length === 0) {
    return invalidRequest("Font request language must be a non-empty BCP 47 tag");
  }
  try {
    return Intl.getCanonicalLocales(language.trim())[0];
  } catch {
    return invalidRequest(`Invalid font request language ${language}`);
  }
}

function normalizeScript(script: string | undefined): string | undefined {
  if (script === undefined) return undefined;
  if (typeof script !== "string" || !/^[a-z]{4}$/iu.test(script.trim())) {
    return invalidRequest("Font request script must be a four-letter ISO 15924 code");
  }
  const normalized = script.trim().toLowerCase();
  return `${normalized[0]!.toUpperCase()}${normalized.slice(1)}`;
}

function normalizeExpectedSha256(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || !/^[a-f0-9]{64}$/iu.test(value)) {
    return invalidRequest("Font request expectedSha256 must be a 64-character hexadecimal digest");
  }
  return value.toLowerCase();
}

function normalizeRequest(request: FontRequest, index: number): FontRequest {
  if (request === null || typeof request !== "object") {
    return invalidRequest(`Font request ${index} must be an object`);
  }
  const family = validateFamily(request.family, index);
  if (!FONT_STYLES.has(request.style)) invalidRequest(`Font request ${index} style is invalid`);
  if (!Number.isInteger(request.weight) || request.weight < 1 || request.weight > 1000) {
    invalidRequest(`Font request ${index} weight must be an integer from 1 through 1000`);
  }
  if (!FONT_STRETCHES.has(request.stretch)) invalidRequest(`Font request ${index} stretch is invalid`);
  if (!Array.isArray(request.codePoints)) invalidRequest(`Font request ${index} codePoints must be an array`);
  const codePoints = [...new Set(request.codePoints)];
  if (codePoints.length > MAX_FONT_REQUEST_CODE_POINTS) {
    invalidRequest(`Font request ${index} codePoints must contain at most ${MAX_FONT_REQUEST_CODE_POINTS} unique values`);
  }
  for (const codePoint of codePoints) {
    if (!Number.isInteger(codePoint)
      || codePoint < 0
      || codePoint > 0x10ffff
      || codePoint >= 0xd800 && codePoint <= 0xdfff) {
      invalidRequest(`Font request ${index} contains an invalid Unicode scalar value`);
    }
  }
  codePoints.sort((left, right) => left - right);
  const language = normalizeLanguage(request.language);
  const script = normalizeScript(request.script);
  const expectedSha256 = normalizeExpectedSha256(request.expectedSha256);
  return Object.freeze({
    family,
    style: request.style,
    weight: request.weight,
    stretch: request.stretch,
    codePoints: Object.freeze(codePoints),
    ...(language === undefined ? {} : { language }),
    ...(script === undefined ? {} : { script }),
    ...(expectedSha256 === undefined ? {} : { expectedSha256 }),
  });
}

function providerFaceKey(request: FontRequest): string {
  return [
    normalizeFamily(request.family),
    request.style,
    String(request.weight),
    request.stretch,
    request.expectedSha256 ?? "*",
  ].join("\u0000");
}

function diagnosticRequestKey(request: FontRequest): string {
  return [
    normalizeFamily(request.family),
    request.style,
    String(request.weight),
    request.stretch,
    `sha256=${request.expectedSha256 ?? "any"}`,
  ].join(";");
}

function providerDiagnostic(
  diagnostic: FontProviderDiagnostic,
): FontProviderDiagnostic {
  return Object.freeze({
    ...diagnostic,
    ...(diagnostic.details === undefined
      ? {}
      : { details: Object.freeze({ ...diagnostic.details }) }),
  });
}

function errorMessage(cause: unknown): string {
  if (cause instanceof Error && cause.message.length !== 0) return cause.message;
  return typeof cause === "string" && cause.length !== 0 ? cause : "Unknown provider error";
}

function providerFailureDiagnostics(
  requests: readonly FontRequest[],
  cause: unknown,
): readonly FontProviderDiagnostic[] {
  const error = errorMessage(cause);
  return requests.map((request) => providerDiagnostic({
    code: "FONT_PROVIDER_FAILED",
    severity: "warning",
    fidelity: "approximate",
    phase: "layout",
    message: `Font provider failed while resolving ${request.family}`,
    details: {
      requestKey: diagnosticRequestKey(request),
      family: request.family,
      error,
    },
  }));
}

function exactFaceKey(face: {
  readonly family: string;
  readonly style: FontStyle;
  readonly weight: number;
  readonly stretch: FontStretch;
}): string {
  return [normalizeFamily(face.family), face.style, String(face.weight), face.stretch].join("\u0000");
}

function assetFaceKey(asset: PreparedFontAsset): string {
  return exactFaceKey(asset);
}

function mergeFontRequests(requests: readonly FontRequest[]): readonly FontRequest[] {
  interface PendingRequest {
    readonly request: FontRequest;
    readonly codePoints: Set<number>;
    language: string | undefined;
    script: string | undefined;
  }
  const merged = new Map<string, PendingRequest>();
  for (const request of requests) {
    const key = providerFaceKey(request);
    const existing = merged.get(key);
    if (existing === undefined) {
      merged.set(key, {
        request,
        codePoints: new Set(request.codePoints),
        language: request.language,
        script: request.script,
      });
      continue;
    }
    for (const codePoint of request.codePoints) {
      if (existing.codePoints.size === MAX_FONT_REQUEST_CODE_POINTS) break;
      existing.codePoints.add(codePoint);
    }
    if (existing.language !== request.language) existing.language = undefined;
    if (existing.script !== request.script) existing.script = undefined;
  }
  return Object.freeze([...merged.values()].map((entry) => {
    const { request } = entry;
    return Object.freeze({
      family: request.family,
      style: request.style,
      weight: request.weight,
      stretch: request.stretch,
      codePoints: Object.freeze([...entry.codePoints].sort((left, right) => left - right)),
      ...(entry.language === undefined ? {} : { language: entry.language }),
      ...(entry.script === undefined ? {} : { script: entry.script }),
      ...(request.expectedSha256 === undefined ? {} : { expectedSha256: request.expectedSha256 }),
    });
  }));
}

function copyProviderAsset(asset: PreparedFontAsset): FontAsset {
  return Object.freeze({
    family: asset.family,
    bytes: asset.bytes.slice(0),
    ...(asset.sha256 === undefined ? {} : { sha256: asset.sha256 }),
    style: asset.style,
    weight: asset.weight,
    stretch: asset.stretch,
  });
}

async function sha256(bytes: ArrayBuffer): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (subtle === undefined) {
    throw new OfficeEngineError(
      "FONT_HASH_UNAVAILABLE",
      "SHA-256 is unavailable in this JavaScript environment",
    );
  }
  const digest = await subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(digest)]
    .map((value) => value.toString(16).padStart(2, "0"))
    .join("");
}

/** Verifies optional manifest digests before static fonts cross a document boundary. */
export async function verifyFontAssetIntegrity(
  assets: readonly PreparedFontAsset[],
): Promise<void> {
  await Promise.all(assets.map(async (asset) => {
    if (asset.sha256 === undefined) return;
    const actualSha256 = await sha256(asset.bytes);
    if (actualSha256 === asset.sha256) return;
    const diagnostic: FontProviderDiagnostic = {
      code: "FONT_INTEGRITY_MISMATCH",
      severity: "error",
      fidelity: "not-rendered",
      phase: "security",
      message: `Static font ${asset.family} did not match its expected SHA-256`,
      details: {
        family: asset.family,
        expectedSha256: asset.sha256,
        actualSha256,
      },
    };
    throw new OfficeEngineError(
      "FONT_INTEGRITY_MISMATCH",
      diagnostic.message,
      { diagnostics: [diagnostic] },
    );
  }));
}

/** Engine-scoped lazy provider cache. Returned buffers are copies safe to transfer to Workers. */
export class FontProviderCache {
  readonly #provider: FontProvider;
  readonly #maxBytes: number;
  readonly #policy: FontPolicy;
  readonly #timeoutMs: number;
  readonly #faces = new Map<string, { readonly asset: PreparedFontAsset; readonly hash: string }>();
  readonly #content = new Map<string, ArrayBuffer>();
  readonly #inflight = new Map<string, Promise<readonly FontProviderDiagnostic[]>>();
  readonly #controllers = new Set<AbortController>();
  readonly #pins = new Map<string, number>();
  #cachedBytes = 0;
  #closed = false;

  constructor(provider: FontProvider, options: FontProviderCacheOptions) {
    if (typeof provider !== "function") invalidRequest("fontProvider must be a function");
    if (options === null || typeof options !== "object") invalidRequest("Font provider cache options are required");
    if (!Number.isSafeInteger(options.maxBytes) || options.maxBytes < 0) {
      invalidRequest("Font provider cache maxBytes must be a non-negative safe integer");
    }
    const policy = options.policy ?? "deterministic";
    if (policy !== "deterministic" && policy !== "local-first") {
      invalidRequest("fontPolicy must be deterministic or local-first");
    }
    const timeoutMs = options.timeoutMs ?? DEFAULT_FONT_LOAD_TIMEOUT_MS;
    if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
      invalidRequest("Font provider timeoutMs must be greater than zero");
    }
    this.#provider = provider;
    this.#maxBytes = options.maxBytes;
    this.#policy = policy;
    this.#timeoutMs = timeoutMs;
  }

  async resolve(requests: readonly FontRequest[], signal?: AbortSignal): Promise<FontProviderResult> {
    this.#ensureOpen();
    if (!Array.isArray(requests)) invalidRequest("Font requests must be an array");
    if (signal?.aborted === true) {
      throw new OfficeEngineError("OPERATION_ABORTED", "The operation was aborted", { cause: signal.reason });
    }
    const normalized = mergeFontRequests(requests.map((request, index) => {
      const normalizedRequest = normalizeRequest(request, index);
      const sourceFamily = pdfObjectSourceFamily(normalizedRequest.family);
      return sourceFamily === undefined
        ? normalizedRequest
        : Object.freeze({ ...normalizedRequest, family: sourceFamily });
    }));
    const pinnedKeys = normalized.map(providerFaceKey);
    this.#pin(pinnedKeys);
    try {
      const missing = Object.freeze(normalized.filter(
        (request) => !this.#faces.has(providerFaceKey(request)),
      ));
      const diagnostics: FontProviderDiagnostic[] = [];
      if (missing.length !== 0 && this.#maxBytes === 0) {
        diagnostics.push(...missing.map((request) => providerDiagnostic({
          code: "FONT_PROVIDER_CACHE_LIMIT",
          severity: "warning",
          fidelity: "approximate",
          phase: "layout",
          message: `Font provider cache has no room for ${request.family}`,
          details: {
            requestKey: diagnosticRequestKey(request),
            family: request.family,
            fontBytes: 0,
            cacheBytes: 0,
            cacheLimit: 0,
          },
        })));
      } else if (missing.length !== 0) {
        const operations = new Set<Promise<readonly FontProviderDiagnostic[]>>();
        const fresh: FontRequest[] = [];
        for (const request of missing) {
          const operation = this.#inflight.get(providerFaceKey(request));
          if (operation === undefined) fresh.push(request);
          else operations.add(operation);
        }
        if (fresh.length !== 0) {
          const batch = Object.freeze(fresh);
          const operation = this.#load(batch);
          for (const request of batch) this.#inflight.set(providerFaceKey(request), operation);
          const cleanup = () => {
            for (const request of batch) {
              const key = providerFaceKey(request);
              if (this.#inflight.get(key) === operation) this.#inflight.delete(key);
            }
          };
          operation.then(cleanup, cleanup);
          operations.add(operation);
        }
        const outcomes = await waitForFontProviderOperations([...operations], signal);
        this.#ensureOpen();
        const requestKeys = new Set(missing.map(diagnosticRequestKey));
        for (const outcome of outcomes) {
          diagnostics.push(...outcome.filter(
            (diagnostic) => typeof diagnostic.details?.requestKey === "string"
              && requestKeys.has(diagnostic.details.requestKey),
          ));
        }
      }
      this.#ensureOpen();
      const assets = normalized.flatMap((request) => {
        const key = providerFaceKey(request);
        const cached = this.#faces.get(key);
        if (cached === undefined) return [];
        this.#faces.delete(key);
        this.#faces.set(key, cached);
        return [copyProviderAsset(cached.asset)];
      });
      return Object.freeze({
        assets: Object.freeze(assets),
        diagnostics: Object.freeze(diagnostics),
      });
    } finally {
      this.#unpin(pinnedKeys);
    }
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    for (const controller of this.#controllers) controller.abort("Font provider cache closed");
    this.#controllers.clear();
    this.#inflight.clear();
    this.#faces.clear();
    this.#content.clear();
    this.#pins.clear();
    this.#cachedBytes = 0;
  }

  async #load(requests: readonly FontRequest[]): Promise<readonly FontProviderDiagnostic[]> {
    this.#ensureOpen();
    if (requests.length > MAX_PROVIDER_FACES_PER_CALL) {
      const diagnostics: FontProviderDiagnostic[] = [];
      for (let offset = 0; offset < requests.length; offset += MAX_PROVIDER_FACES_PER_CALL) {
        this.#ensureOpen();
        diagnostics.push(...await this.#loadBatch(
          Object.freeze(requests.slice(offset, offset + MAX_PROVIDER_FACES_PER_CALL)),
        ));
      }
      return Object.freeze(diagnostics);
    }
    return this.#loadBatch(requests);
  }

  async #loadBatch(requests: readonly FontRequest[]): Promise<readonly FontProviderDiagnostic[]> {
    this.#ensureOpen();
    const controller = new AbortController();
    this.#controllers.add(controller);
    try {
      let response: readonly FontAsset[];
      try {
        response = await this.#callProvider(requests, controller);
      } catch (cause) {
        if (cause instanceof OfficeEngineError && cause.code === "ENGINE_CLOSED") throw cause;
        if (cause instanceof OfficeEngineError && cause.code === "FONT_PROVIDER_TIMEOUT") {
          return Object.freeze(requests.map((request) => providerDiagnostic({
            code: "FONT_PROVIDER_TIMEOUT",
            severity: "warning",
            fidelity: "approximate",
            phase: "layout",
            message: `Font provider timed out while resolving ${request.family}`,
            details: {
              requestKey: diagnosticRequestKey(request),
              family: request.family,
              timeoutMs: this.#timeoutMs,
            },
          })));
        }
        return Object.freeze(providerFailureDiagnostics(requests, cause));
      }
      if (!Array.isArray(response)) {
        return this.#invalidResponseDiagnostics(requests, "fontProvider must return an array");
      }
      if (response.length > requests.length) {
        return this.#invalidResponseDiagnostics(
          requests,
          "fontProvider returned more faces than were requested",
        );
      }
      let prepared: readonly PreparedFontAsset[];
      try {
        prepared = prepareFontAssets(response, this.#maxBytes);
      } catch (cause) {
        const message = cause instanceof OfficeEngineError && cause.code === "FONT_BYTES_LIMIT"
          ? "The provider response exceeds the Engine font cache budget"
          : errorMessage(cause);
        return this.#invalidResponseDiagnostics(requests, message);
      }
      const hashed = await Promise.all(prepared.map(async (asset) => ({
        asset,
        hash: await sha256(asset.bytes),
      })));
      this.#ensureOpen();
      const byFace = new Map(hashed.map((entry) => [assetFaceKey(entry.asset), entry]));
      const diagnostics: FontProviderDiagnostic[] = [];
      for (const request of requests) {
        const candidate = byFace.get([
          normalizeFamily(request.family),
          request.style,
          String(request.weight),
          request.stretch,
        ].join("\u0000"));
        if (candidate === undefined) {
          diagnostics.push(providerDiagnostic({
            code: "FONT_PROVIDER_MISSING",
            severity: "warning",
            fidelity: "approximate",
            phase: "layout",
            message: `Font provider did not return ${request.family} ${request.style} ${request.weight} ${request.stretch}`,
            details: {
              requestKey: diagnosticRequestKey(request),
              family: request.family,
              style: request.style,
              weight: request.weight,
              stretch: request.stretch,
              codePointCount: request.codePoints.length,
              ...(request.language === undefined ? {} : { language: request.language }),
              ...(request.script === undefined ? {} : { script: request.script }),
            },
          }));
          continue;
        }
        const expectedSha256 = request.expectedSha256 ?? candidate.asset.sha256;
        if (expectedSha256 !== undefined && candidate.hash !== expectedSha256) {
          diagnostics.push(providerDiagnostic({
            code: "FONT_INTEGRITY_MISMATCH",
            severity: "error",
            fidelity: "not-rendered",
            phase: "security",
            message: `Font provider returned an unexpected binary for ${request.family}`,
            details: {
              requestKey: diagnosticRequestKey(request),
              family: request.family,
              expectedSha256,
              actualSha256: candidate.hash,
            },
          }));
          continue;
        }
        let bytes = this.#content.get(candidate.hash);
        if (bytes === undefined) {
          this.#evictFor(candidate.asset.bytes.byteLength);
          if (this.#cachedBytes + candidate.asset.bytes.byteLength > this.#maxBytes) {
            diagnostics.push(providerDiagnostic({
              code: "FONT_PROVIDER_CACHE_LIMIT",
              severity: "warning",
              fidelity: "approximate",
              phase: "layout",
              message: `Font provider cache has no room for ${request.family}`,
              details: {
                requestKey: diagnosticRequestKey(request),
                family: request.family,
                fontBytes: candidate.asset.bytes.byteLength,
                cacheBytes: this.#cachedBytes,
                cacheLimit: this.#maxBytes,
              },
            }));
            continue;
          }
          bytes = candidate.asset.bytes;
          this.#content.set(candidate.hash, bytes);
          this.#cachedBytes += bytes.byteLength;
        }
        const asset = Object.freeze({ ...candidate.asset, bytes });
        this.#faces.set(providerFaceKey(request), Object.freeze({ asset, hash: candidate.hash }));
      }
      return Object.freeze(diagnostics);
    } finally {
      this.#controllers.delete(controller);
    }
  }

  async #callProvider(
    requests: readonly FontRequest[],
    controller: AbortController,
  ): Promise<readonly FontAsset[]> {
    return new Promise<readonly FontAsset[]>((resolve, reject) => {
      let settled = false;
      const finish = (callback: () => void): void => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        controller.signal.removeEventListener("abort", onAbort);
        callback();
      };
      const onAbort = (): void => finish(() => reject(new OfficeEngineError(
        "ENGINE_CLOSED",
        "The engine is closed",
        { cause: controller.signal.reason },
      )));
      const timer = setTimeout(() => {
        finish(() => reject(new OfficeEngineError(
          "FONT_PROVIDER_TIMEOUT",
          `Font provider exceeded ${this.#timeoutMs}ms`,
        )));
        controller.abort("Font provider timed out");
      }, this.#timeoutMs);
      controller.signal.addEventListener("abort", onAbort, { once: true });
      Promise.resolve()
        .then(() => {
          if (settled) throw new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
          return this.#provider(requests, Object.freeze({
            signal: controller.signal,
            policy: this.#policy,
          }));
        })
        .then(
          (assets) => finish(() => resolve(assets)),
          (cause: unknown) => finish(() => reject(cause)),
        );
    });
  }

  #invalidResponseDiagnostics(
    requests: readonly FontRequest[],
    error: string,
  ): readonly FontProviderDiagnostic[] {
    return Object.freeze(requests.map((request) => providerDiagnostic({
      code: "FONT_PROVIDER_INVALID_RESPONSE",
      severity: "warning",
      fidelity: "approximate",
      phase: "layout",
      message: `Font provider returned an invalid response for ${request.family}`,
      details: {
        requestKey: diagnosticRequestKey(request),
        family: request.family,
        error,
      },
    })));
  }

  #evictFor(requiredBytes: number): void {
    while (this.#cachedBytes + requiredBytes > this.#maxBytes && this.#faces.size !== 0) {
      let oldestKey: string | undefined;
      for (const key of this.#faces.keys()) {
        if (!this.#pins.has(key)) {
          oldestKey = key;
          break;
        }
      }
      if (oldestKey === undefined) break;
      const oldest = this.#faces.get(oldestKey);
      this.#faces.delete(oldestKey);
      if (oldest === undefined) continue;
      const shared = [...this.#faces.values()].some(({ hash }) => hash === oldest.hash);
      if (!shared) {
        this.#content.delete(oldest.hash);
        this.#cachedBytes -= oldest.asset.bytes.byteLength;
      }
    }
  }

  #ensureOpen(): void {
    if (this.#closed) throw new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
  }

  #pin(keys: readonly string[]): void {
    for (const key of keys) this.#pins.set(key, (this.#pins.get(key) ?? 0) + 1);
  }

  #unpin(keys: readonly string[]): void {
    for (const key of keys) {
      const count = this.#pins.get(key);
      if (count === undefined || count === 1) this.#pins.delete(key);
      else this.#pins.set(key, count - 1);
    }
  }
}

async function waitForFontProviderOperations<T>(
  operations: readonly Promise<T>[],
  signal: AbortSignal | undefined,
): Promise<T[]> {
  if (signal === undefined) return Promise.all(operations);
  let abort: (() => void) | undefined;
  try {
    return await Promise.race([
      Promise.all(operations),
      new Promise<never>((_, reject) => {
        abort = () => reject(new OfficeEngineError(
          "OPERATION_ABORTED",
          "The operation was aborted",
          { cause: signal.reason },
        ));
        signal.addEventListener("abort", abort, { once: true });
        if (signal.aborted) abort();
      }),
    ]);
  } finally {
    if (abort !== undefined) signal.removeEventListener("abort", abort);
  }
}

function ambientFontSet(): RuntimeFontSet | undefined {
  const scope = globalThis as unknown as {
    readonly fonts?: RuntimeFontSet;
    readonly document?: { readonly fonts?: RuntimeFontSet };
  };
  return scope.fonts ?? scope.document?.fonts;
}

function ambientRuntime(): FontRuntime | undefined {
  const scope = globalThis as unknown as {
    readonly FontFace?: FontRuntime["FontFace"];
  };
  const fontSet = ambientFontSet();
  if (scope.FontFace === undefined || fontSet === undefined) return undefined;
  return { FontFace: scope.FontFace, fontSet };
}

export interface FontFaceDescriptor {
  readonly family: string;
  readonly style: FontStyle;
  readonly weight: number;
  readonly stretch: FontStretch;
}

export interface FontFaceResolutionRequest extends FontFaceDescriptor {
  readonly codePoints?: readonly number[];
}

export type FontResolutionSource = "embedded" | "host" | "provider" | "browser" | "fallback";

export interface FontResolution {
  readonly family: string;
  readonly source: FontResolutionSource;
  /** Preserve authored blank advances when a local substitute has different spaces. */
  readonly spaceAdvanceEm?: number;
}

export interface ResolvedFontFace extends FontFaceDescriptor, FontResolution {
  readonly alternateFamily?: string;
}

export interface RuntimeFontResolver {
  resolve(family: string): FontResolution;
  resolveFace?(face: FontFaceResolutionRequest): FontFaceDescriptor;
}

type BinaryFontSource = "embedded" | "host" | "provider" | "fallback";

interface RegisteredBinaryFont {
  readonly asset: PreparedFontAsset;
  readonly source: BinaryFontSource;
}

type RequestedFontFace = FontFaceResolutionRequest;

type FontScriptDemand = "latin" | "cjk" | "complex" | "symbol" | "mixed";

export interface FontFallbackTextSegment {
  readonly text: string;
  readonly start: number;
  readonly end: number;
}

export class RegisteredFonts implements RuntimeFontResolver {
  readonly fontSet: RuntimeFontSet | undefined;
  readonly #faces: RuntimeFontFace[];
  readonly #resolutions: Map<string, ResolvedFontFace>;
  readonly #familyDefaults: Map<string, ResolvedFontFace>;
  readonly #binaryAssets: RegisteredBinaryFont[];
  readonly #scriptDemands: ReadonlyMap<string, FontScriptDemand>;
  readonly #runtime: FontRuntime | undefined;
  readonly #fontAlternateNames: ReadonlyMap<string, readonly string[]>;
  #closed = false;

  constructor(
    fontSet: RuntimeFontSet | undefined,
    faces: RuntimeFontFace[],
    resolutions: Map<string, ResolvedFontFace> = new Map(),
    familyDefaults: Map<string, ResolvedFontFace> = new Map(),
    binaryAssets: RegisteredBinaryFont[] = [],
    scriptDemands: ReadonlyMap<string, FontScriptDemand> = new Map(),
    runtime?: FontRuntime,
    fontAlternateNames: ReadonlyMap<string, readonly string[]> = new Map(),
  ) {
    this.fontSet = fontSet;
    this.#faces = faces;
    this.#resolutions = resolutions;
    this.#familyDefaults = familyDefaults;
    this.#binaryAssets = binaryAssets;
    this.#scriptDemands = scriptDemands;
    this.#runtime = runtime;
    this.#fontAlternateNames = fontAlternateNames;
    REGISTERED_FONT_DIAGNOSTICS.set(this, new Map());
  }

  resolve(family: string): FontResolution {
    const resolved = this.resolveFace({
      family,
      style: "normal",
      weight: 400,
      stretch: "normal",
    });
    return {
      family: resolved.family,
      source: resolved.source,
      ...(resolved.spaceAdvanceEm === undefined ? {} : { spaceAdvanceEm: resolved.spaceAdvanceEm }),
    };
  }

  resolveFace(face: FontFaceResolutionRequest): ResolvedFontFace {
    const normalizedFamily = normalizeFamily(face.family);
    const exact = this.#resolutions.get(exactFaceKey(face));
    if (exact !== undefined && (exact.source !== "fallback" || exact.alternateFamily !== undefined)) return exact;
    const familyDefault = this.#familyDefaults.get(normalizedFamily);
    if (familyDefault !== undefined && familyDefault.source !== "fallback") {
      return Object.freeze({
        family: familyDefault.family,
        style: face.style,
        weight: face.weight,
        stretch: face.stretch,
        source: familyDefault.source,
        ...(familyDefault.spaceAdvanceEm === undefined ? {} : { spaceAdvanceEm: familyDefault.spaceAdvanceEm }),
      });
    }
    for (const family of this.#fontAlternateNames.get(normalizedFamily) ?? []) {
      const alternate = this.#resolutions.get(exactFaceKey({ ...face, family }))
        ?? this.#familyDefaults.get(normalizeFamily(family));
      if (alternate !== undefined && alternate.source !== "fallback") {
        return Object.freeze({ ...alternate, ...face, family: alternate.family, source: "fallback" });
      }
    }
    if (exact !== undefined) return exact;
    if (familyDefault !== undefined) {
      return Object.freeze({ ...familyDefault, ...face, family: familyDefault.family });
    }
    const sourceFamily = pdfObjectSourceFamily(face.family);
    const fallbackFace = sourceFamily === undefined ? face : { ...face, family: sourceFamily };
    const normalizedFallbackFamily = normalizeFamily(fallbackFace.family);
    if (sourceFamily !== undefined) {
      const source = this.#resolutions.get(exactFaceKey(fallbackFace))
        ?? this.#familyDefaults.get(normalizedFallbackFamily);
      if (source !== undefined) {
        return Object.freeze({
          ...face,
          family: source.family,
          source: "fallback",
        });
      }
    }
    if (GENERIC_FAMILIES.has(normalizedFallbackFamily)) {
      return Object.freeze({ ...face, family: normalizedFallbackFamily, source: "browser" });
    }
    if (sourceFamily !== undefined && PDF_PREVIEW_HELVETICA_FALLBACKS.has(normalizedFallbackFamily)) {
      return Object.freeze({ ...face, family: "Helvetica", source: "fallback" });
    }
    const scriptDemand = face.codePoints === undefined
      ? this.#scriptDemands.get(exactFaceKey(face)) ?? "latin"
      : classifyFontScriptDemand(face.codePoints);
    const approximateFamily = approximateFontFamilyForDemand(normalizedFallbackFamily, scriptDemand);
    const managedFallbacks = managedFallbackFamilyNames(approximateFamily, scriptDemand);
    for (const candidateFamily of managedFallbacks) {
      const candidateFace = { ...face, family: candidateFamily };
      const managed = this.#resolutions.get(exactFaceKey(candidateFace))
        ?? this.#familyDefaults.get(normalizeFamily(candidateFamily));
      if (managed !== undefined) {
        return Object.freeze({
          ...face,
          family: managed.family,
          source: "fallback",
        });
      }
    }
    return Object.freeze({
      ...face,
      family: approximateFamily,
      source: "fallback",
    });
  }

  /** True only when this exact style/weight/stretch face was registered successfully. */
  hasExactFace(face: FontFaceDescriptor): boolean {
    const resolved = this.#resolutions.get(exactFaceKey(face));
    return resolved !== undefined && resolved.source !== "fallback";
  }

  /** Binary fonts reusable by isolated codecs which cannot access FontFaceSet. */
  imageCodecFonts(): readonly PreparedFontAsset[] {
    return this.#binaryAssets.map(({ asset }) => asset);
  }

  /** Non-fatal font registration failures accumulated by initial and lazy loads. */
  diagnostics(): readonly Diagnostic[] {
    return immutableDiagnostics([...registeredFontDiagnosticStore(this).values()]);
  }

  async addEmbeddedAssets(
    assets: readonly PreparedFontAsset[],
    timeoutMs = DEFAULT_FONT_LOAD_TIMEOUT_MS,
    signal?: AbortSignal,
  ): Promise<void> {
    if (assets.length === 0) return;
    if (this.#closed) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    const runtime = this.#runtime;
    if (runtime === undefined) {
      for (const asset of assets) {
        recordFontLoadDiagnostic(
          registeredFontDiagnosticStore(this),
          "FONT_ENVIRONMENT_UNAVAILABLE",
          "embedded",
          asset,
          "FontFace and FontFaceSet are unavailable",
        );
      }
      return;
    }
    const aliases = createAliasRegistry();
    for (const asset of assets) ensureAlias(aliases, asset.family, "embedded");
    await loadBinaryFonts(
      runtime,
      assets,
      "embedded",
      false,
      this.#faces,
      this.#resolutions,
      this.#familyDefaults,
      this.#binaryAssets,
      aliases,
      timeoutMs,
      signal,
      registeredFontDiagnosticStore(this),
    );
  }

  async addLocalRequests(
    requests: readonly RequestedFontFace[],
    timeoutMs = DEFAULT_FONT_LOAD_TIMEOUT_MS,
    signal?: AbortSignal,
  ): Promise<void> {
    if (requests.length === 0) return;
    if (this.#closed) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    const runtime = this.#runtime;
    if (runtime === undefined) return;
    await resolveLocalFonts(
      runtime,
      collectRequestedFaces([], requests),
      this.#faces,
      this.#resolutions,
      this.#familyDefaults,
      createAliasRegistry(),
      timeoutMs,
      signal,
      this.#fontAlternateNames,
    );
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    const faces = this.#faces.splice(0);
    this.#binaryAssets.splice(0);
    if (this.fontSet === undefined) return;
    for (const face of faces) {
      try {
        this.fontSet.delete(face);
      } catch {
        // Realm teardown can invalidate FontFaceSet before document cleanup runs.
      }
    }
  }
}

export interface RegisterFontOptions {
  readonly providerAssets?: readonly PreparedFontAsset[];
  readonly policy?: FontPolicy;
  readonly requestedFaces?: readonly RequestedFontFace[];
  readonly fontAlternateNames?: SceneDocument["fontAlternateNames"];
}

interface FontAliasRegistry {
  readonly registrationId: number;
  readonly aliases: Map<string, string>;
  readonly counts: Record<FontResolutionSource, number>;
}

let nextFontRegistrationId = 0;

function createAliasRegistry(): FontAliasRegistry {
  const registrationId = nextFontRegistrationId;
  nextFontRegistrationId = (nextFontRegistrationId + 1) % Number.MAX_SAFE_INTEGER;
  return {
    registrationId,
    aliases: new Map(),
    counts: { embedded: 0, host: 0, provider: 0, browser: 0, fallback: 0 },
  };
}

function ensureAlias(
  registry: FontAliasRegistry,
  family: string,
  source: FontResolutionSource,
): string {
  const normalized = source === "fallback" ? `fallback:${normalizeFamily(family)}` : normalizeFamily(family);
  const existing = registry.aliases.get(normalized);
  if (existing !== undefined) return existing;
  const index = registry.counts[source];
  registry.counts[source] += 1;
  const alias = `OfficeViewer ${registry.registrationId} ${source} ${source === "fallback" ? normalizeFamily(family) === "calibri" ? "Carlito " : "Caladea " : ""}${index}`;
  registry.aliases.set(normalized, alias);
  return alias;
}

function normalizeFaceDescriptor(face: RequestedFontFace, index: number): RequestedFontFace {
  if (face === null || typeof face !== "object") invalid(`requestedFaces[${index}] must be an object`);
  const family = validateFamily(face.family, index);
  if (!FONT_STYLES.has(face.style)) invalid(`requestedFaces[${index}].style is invalid`);
  if (!Number.isInteger(face.weight) || face.weight < 1 || face.weight > 1000) {
    invalid(`requestedFaces[${index}].weight must be an integer from 1 through 1000`);
  }
  if (!FONT_STRETCHES.has(face.stretch)) invalid(`requestedFaces[${index}].stretch is invalid`);
  const codePoints = face.codePoints;
  let normalizedCodePoints: readonly number[] | undefined;
  if (codePoints !== undefined) {
    if (!Array.isArray(codePoints)) invalid(`requestedFaces[${index}].codePoints must be an array`);
    const uniqueCodePoints = new Set<number>();
    for (const codePoint of codePoints) {
      if (!Number.isInteger(codePoint)
        || codePoint < 0
        || codePoint > 0x10ffff
        || codePoint >= 0xd800 && codePoint <= 0xdfff) {
        invalid(`requestedFaces[${index}].codePoints contains an invalid Unicode scalar value`);
      }
      if (!uniqueCodePoints.has(codePoint)
        && uniqueCodePoints.size === MAX_FONT_REQUEST_CODE_POINTS) {
        invalid(`requestedFaces[${index}].codePoints must contain at most ${MAX_FONT_REQUEST_CODE_POINTS} unique values`);
      }
      uniqueCodePoints.add(codePoint);
    }
    normalizedCodePoints = Object.freeze([...uniqueCodePoints].sort((left, right) => left - right));
  }
  return Object.freeze({
    family,
    style: face.style,
    weight: face.weight,
    stretch: face.stretch,
    ...(normalizedCodePoints === undefined ? {} : { codePoints: normalizedCodePoints }),
  });
}

function collectRequestedFaces(
  requestedFamilies: readonly string[],
  requestedFaces: readonly RequestedFontFace[],
): readonly RequestedFontFace[] {
  const faces = [
    ...requestedFamilies.map((family) => ({
      family,
      style: "normal" as const,
      weight: 400,
      stretch: "normal" as const,
    })),
    ...requestedFaces,
  ].map(normalizeFaceDescriptor);
  const unique = new Map<string, {
    readonly face: RequestedFontFace;
    codePoints: Set<number> | undefined;
  }>();
  for (const face of faces) {
    const key = exactFaceKey(face);
    const existing = unique.get(key);
    if (existing === undefined) {
      unique.set(key, {
        face,
        codePoints: face.codePoints === undefined ? undefined : new Set(face.codePoints),
      });
      continue;
    }
    if (face.codePoints === undefined) continue;
    existing.codePoints ??= new Set();
    for (const codePoint of face.codePoints) {
      if (existing.codePoints.size === MAX_FONT_REQUEST_CODE_POINTS) break;
      existing.codePoints.add(codePoint);
    }
  }
  return Object.freeze([...unique.values()].map(({ face, codePoints }) => Object.freeze({
    ...face,
    ...(codePoints === undefined
      ? {}
      : { codePoints: Object.freeze([...codePoints].sort((left, right) => left - right)) }),
  })));
}

const CJK_SCRIPT_DEMAND = /\p{Script_Extensions=Han}|\p{Script_Extensions=Hiragana}|\p{Script_Extensions=Katakana}|\p{Script_Extensions=Hangul}|\p{Script_Extensions=Bopomofo}/u;
const EUROPEAN_SCRIPT_DEMAND = /\p{Script_Extensions=Latin}|\p{Script_Extensions=Cyrillic}|\p{Script_Extensions=Greek}/u;
const COMPLEX_SCRIPT_DEMAND = /\p{Script_Extensions=Arabic}|\p{Script_Extensions=Hebrew}|\p{Script_Extensions=Syriac}|\p{Script_Extensions=Thaana}|\p{Script_Extensions=Nko}|\p{Script_Extensions=Devanagari}|\p{Script_Extensions=Bengali}|\p{Script_Extensions=Gurmukhi}|\p{Script_Extensions=Gujarati}|\p{Script_Extensions=Oriya}|\p{Script_Extensions=Tamil}|\p{Script_Extensions=Telugu}|\p{Script_Extensions=Kannada}|\p{Script_Extensions=Malayalam}|\p{Script_Extensions=Sinhala}|\p{Script_Extensions=Thai}|\p{Script_Extensions=Lao}|\p{Script_Extensions=Tibetan}|\p{Script_Extensions=Myanmar}|\p{Script_Extensions=Khmer}/u;
const LETTER_DEMAND = /\p{Letter}/u;
const NUMBER_DEMAND = /\p{Number}/u;
const SYMBOL_DEMAND = /\p{Symbol}|\p{Private_Use}/u;
const VISIBLE_NEUTRAL_DEMAND = /[^\p{Separator}\p{Control}\p{Mark}]/u;

function classifyFontScriptDemand(codePoints: readonly number[]): FontScriptDemand {
  const demands = new Set<Exclude<FontScriptDemand, "mixed">>();
  let hasVisibleNeutral = false;
  for (const codePoint of codePoints) {
    if (!Number.isInteger(codePoint)
      || codePoint < 0
      || codePoint > 0x10ffff
      || codePoint >= 0xd800 && codePoint <= 0xdfff) continue;
    const character = String.fromCodePoint(codePoint);
    if (CJK_SCRIPT_DEMAND.test(character)) demands.add("cjk");
    else if (EUROPEAN_SCRIPT_DEMAND.test(character) || NUMBER_DEMAND.test(character)) demands.add("latin");
    else if (COMPLEX_SCRIPT_DEMAND.test(character) || LETTER_DEMAND.test(character)) demands.add("complex");
    else if (SYMBOL_DEMAND.test(character)) demands.add("symbol");
    else hasVisibleNeutral ||= VISIBLE_NEUTRAL_DEMAND.test(character);
    if (demands.size > 1) return "mixed";
  }
  return demands.values().next().value ?? (hasVisibleNeutral ? "symbol" : "latin");
}

function characterScriptDemand(character: string): Exclude<FontScriptDemand, "mixed"> | undefined {
  if (CJK_SCRIPT_DEMAND.test(character)) return "cjk";
  if (EUROPEAN_SCRIPT_DEMAND.test(character) || NUMBER_DEMAND.test(character)) return "latin";
  if (COMPLEX_SCRIPT_DEMAND.test(character) || LETTER_DEMAND.test(character)) return "complex";
  if (SYMBOL_DEMAND.test(character)) return "symbol";
  return undefined;
}

/**
 * Splits a mixed-script fallback run at strong script boundaries. Punctuation,
 * whitespace, controls, and combining marks stay with the preceding script so
 * layout does not manufacture extra typographic boundaries around neutral text.
 */
export function fontFallbackTextSegments(text: string): readonly FontFallbackTextSegment[] {
  if (text.length === 0) return Object.freeze([]);
  const segments: FontFallbackTextSegment[] = [];
  let start = 0;
  let demand: Exclude<FontScriptDemand, "mixed"> | undefined;
  for (let index = 0; index < text.length;) {
    const codePoint = text.codePointAt(index)!;
    const character = String.fromCodePoint(codePoint);
    const nextDemand = characterScriptDemand(character);
    if (nextDemand !== undefined) {
      if (demand !== undefined && nextDemand !== demand) {
        segments.push(Object.freeze({ text: text.slice(start, index), start, end: index }));
        start = index;
      }
      demand = nextDemand;
    }
    index += character.length;
  }
  segments.push(Object.freeze({ text: text.slice(start), start, end: text.length }));
  return Object.freeze(segments);
}

function requestedScriptDemands(
  faces: readonly RequestedFontFace[],
): ReadonlyMap<string, FontScriptDemand> {
  const demands = new Map<string, FontScriptDemand>();
  for (const face of faces) {
    demands.set(exactFaceKey(face), classifyFontScriptDemand(face.codePoints ?? []));
  }
  return demands;
}

function recordResolution(
  resolutions: Map<string, ResolvedFontFace>,
  familyDefaults: Map<string, ResolvedFontFace>,
  authoredFace: FontFaceDescriptor,
  renderedFamily: string,
  source: FontResolutionSource,
  spaceAdvanceEm?: number,
  alternateFamily?: string,
): void {
  const resolved = Object.freeze({
    ...authoredFace,
    family: renderedFamily,
    source,
    ...(spaceAdvanceEm === undefined ? {} : { spaceAdvanceEm }),
    ...(alternateFamily === undefined ? {} : { alternateFamily }),
  });
  resolutions.set(exactFaceKey(authoredFace), resolved);
  const normalizedFamily = normalizeFamily(authoredFace.family);
  if (!familyDefaults.has(normalizedFamily)) familyDefaults.set(normalizedFamily, resolved);
}

export async function registerFonts(
  assets: readonly PreparedFontAsset[],
  timeoutMs = DEFAULT_FONT_LOAD_TIMEOUT_MS,
  runtime = ambientRuntime(),
  signal?: AbortSignal,
  embeddedAssets: readonly PreparedFontAsset[] = [],
  requestedFamilies: readonly string[] = [],
  options: RegisterFontOptions = {},
): Promise<RegisteredFonts> {
  const providerAssets = options.providerAssets ?? [];
  const policy = options.policy ?? "deterministic";
  if (policy !== "deterministic" && policy !== "local-first") {
    throw new OfficeEngineError("INVALID_FONT_POLICY", `Unsupported font policy ${String(policy)}`);
  }
  const requestedFaces = collectRequestedFaces(requestedFamilies, options.requestedFaces ?? []);
  const fontAlternateNames = new Map((options.fontAlternateNames ?? [])
    .map(({ family, names }) => [normalizeFamily(family), names]));
  const scriptDemands = requestedScriptDemands(requestedFaces);
  if (signal?.aborted === true) {
    throw new OfficeEngineError("OPERATION_ABORTED", "The operation was aborted", { cause: signal.reason });
  }
  if (assets.length === 0
    && embeddedAssets.length === 0
    && providerAssets.length === 0
    && requestedFaces.length === 0) {
    return new RegisteredFonts(runtime?.fontSet, [], new Map(), new Map(), [], new Map(), runtime);
  }
  if (runtime === undefined
    || typeof runtime.fontSet.add !== "function"
    || typeof runtime.fontSet.delete !== "function") {
    if (assets.length !== 0 || providerAssets.length !== 0) {
      throw new OfficeEngineError(
        "FONT_ENVIRONMENT_UNAVAILABLE",
        "Host and provider fonts require FontFace and FontFaceSet support in the document execution realm",
      );
    }
    // Without a FontFaceSet, local faces cannot be verified or registered.
    // Treat them as unavailable: Firefox OffscreenCanvas can otherwise select
    // an unrelated cmap for an unresolved authored family.
    const registration = new RegisteredFonts(
      undefined,
      [],
      new Map(),
      new Map(),
      [],
      scriptDemands,
      undefined,
      fontAlternateNames,
    );
    const diagnostics = registeredFontDiagnosticStore(registration);
    for (const asset of embeddedAssets) {
      recordFontLoadDiagnostic(
        diagnostics,
        "FONT_ENVIRONMENT_UNAVAILABLE",
        "embedded",
        asset,
        "FontFace and FontFaceSet are unavailable",
      );
    }
    return registration;
  }

  const faces: RuntimeFontFace[] = [];
  const resolutions = new Map<string, ResolvedFontFace>();
  const familyDefaults = new Map<string, ResolvedFontFace>();
  const binaryAssets: RegisteredBinaryFont[] = [];
  const aliases = createAliasRegistry();
  for (const asset of embeddedAssets) ensureAlias(aliases, asset.family, "embedded");
  for (const asset of assets) ensureAlias(aliases, asset.family, "host");
  if (policy === "local-first") {
    for (const face of requestedFaces) ensureAlias(aliases, face.family, "browser");
  }
  for (const asset of providerAssets) ensureAlias(aliases, asset.family, "provider");
  const registration = new RegisteredFonts(
    runtime.fontSet,
    faces,
    resolutions,
    familyDefaults,
    binaryAssets,
    scriptDemands,
    runtime,
    fontAlternateNames,
  );
  const diagnostics = registeredFontDiagnosticStore(registration);
  try {
    if (embeddedAssets.length !== 0) {
      await loadBinaryFonts(runtime, embeddedAssets, "embedded", false, faces, resolutions, familyDefaults, binaryAssets, aliases, timeoutMs, signal, diagnostics);
    }
    if (assets.length !== 0) {
      await loadBinaryFonts(runtime, assets, "host", true, faces, resolutions, familyDefaults, binaryAssets, aliases, timeoutMs, signal, diagnostics);
    }
    if (policy === "local-first" && requestedFaces.length !== 0) {
      await resolveLocalFonts(runtime, requestedFaces, faces, resolutions, familyDefaults, aliases, timeoutMs, signal, fontAlternateNames);
    }
    if (providerAssets.length !== 0) {
      await loadBinaryFonts(runtime, providerAssets, "provider", false, faces, resolutions, familyDefaults, binaryAssets, aliases, timeoutMs, signal, diagnostics);
    }
    // Office-private fonts need metric-compatible fallbacks for stable line wrapping.
    const missingOfficeFaces = requestedFaces.filter((face) =>
      ["cambria", "calibri"].includes(normalizeFamily(face.family)) && !resolutions.has(exactFaceKey(face)));
    for (const requested of missingOfficeFaces) {
      const fallback = normalizeFamily(requested.family) === "calibri" ? "Carlito" : "Caladea";
      const variant = requested.weight >= 600
        ? requested.style === "normal" ? "Bold" : "BoldItalic"
        : requested.style === "normal" ? "Regular" : "Italic";
      try {
        const response = await fetch(new URL(`./${fallback}-${variant}.ttf`, import.meta.url), {
          signal: signal === undefined
            ? AbortSignal.timeout(timeoutMs)
            : AbortSignal.any([signal, AbortSignal.timeout(timeoutMs)]),
        });
        if (!response.ok) throw new Error(`Bundled font HTTP ${response.status}`);
        const bytes = await response.arrayBuffer();
        if (bytes.byteLength > (fallback === "Carlito" ? 832 : 64) * 1024) throw new Error("Bundled font exceeds its fixed size bound");
        await loadBinaryFonts(runtime, [{ ...requested, bytes }], "fallback", false,
          faces, resolutions, familyDefaults, binaryAssets, aliases, timeoutMs, signal, diagnostics);
      } catch (cause) {
        if (signal?.aborted) throw cause;
        recordFontLoadDiagnostic(diagnostics, "FONT_LOAD_FAILED", "fallback",
          { ...requested, bytes: new ArrayBuffer(0) }, cause);
      }
    }
    return registration;
  } catch (cause) {
    registration.close();
    if (cause instanceof OfficeEngineError) throw cause;
    throw new OfficeEngineError("FONT_LOAD_FAILED", "A host-provided font could not be loaded", { cause });
  }
}

interface SceneFontRun {
  readonly text: string;
  readonly fontFamily: string;
  readonly fontSize?: number;
  readonly italic: boolean;
  readonly bold: boolean;
}

function sceneVisualTextRuns(
  visual: SceneVisual,
  objectText: string | undefined,
  depth = 0,
): readonly SceneFontRun[] {
  if (depth > 64) return [];
  for (let nestedDepth = depth; nestedDepth <= 64; nestedDepth += 1) {
    if (visual.kind === "text") {
      return objectText === undefined ? [] : [{
        text: objectText,
        fontFamily: visual.fontFamily,
        fontSize: visual.fontSize,
        italic: visual.italic,
        bold: visual.bold,
      }];
    }
    if (visual.kind === "rich-text") return visual.runs;
    if (visual.kind === "group") {
      return visual.children.flatMap((child) => sceneVisualTextRuns(child.visual, undefined, nestedDepth + 1));
    }
    const nested = nestedFontVisual(visual);
    if (nested === undefined) return [];
    visual = nested;
  }
  return [];
}

function sceneTextRuns(object: SceneObject): readonly SceneFontRun[] {
  return sceneVisualTextRuns(object.visual, object.text);
}

/** Projects authored and resolved fonts without exposing the internal scene graph. */
export function documentFontRuns(
  object: SceneObject,
  resolver: RuntimeFontResolver,
): readonly DocumentFontRun[] | undefined {
  const text = object.text;
  if (text === undefined || text.length === 0) return undefined;
  const runs: DocumentFontRun[] = [];
  let cursor = 0;
  for (const run of sceneTextRuns(object)) {
    if (run.text.length === 0) continue;
    const semanticText = semanticSymbolFontText(run.text, run.fontFamily);
    const start = text.startsWith(semanticText, cursor)
      ? cursor
      : text.indexOf(semanticText, cursor);
    if (start < 0) return undefined;
    const familyResolution = resolver.resolve(run.fontFamily);
    const resolveSegment = (segmentText: string): {
      readonly family: string;
      readonly source: FontResolutionSource;
    } => {
      const codePoints = new Set<number>();
      const semanticText = semanticSymbolFontText(segmentText, run.fontFamily);
      for (const character of semanticText) {
        if (codePoints.size === MAX_FONT_REQUEST_CODE_POINTS) break;
        codePoints.add(character.codePointAt(0)!);
      }
      const resolvedFace = resolver.resolveFace?.({
        family: run.fontFamily,
        style: run.italic ? "italic" : "normal",
        weight: run.bold ? 700 : 400,
        stretch: "normal",
        codePoints: [...codePoints],
      });
      const source = resolvedFace !== undefined
      && "source" in resolvedFace
      && typeof resolvedFace.source === "string"
        ? resolvedFace.source as FontResolutionSource
        : familyResolution.source;
      return {
        family: source === "fallback"
          ? compatibleFallbackFamily(object.source.format, run.fontFamily)
            ?? resolvedFace?.family
            ?? familyResolution.family
          : resolvedFace?.family ?? familyResolution.family,
        source,
      };
    };
    const wholeResolution = resolveSegment(run.text);
    const segments = wholeResolution.source === "fallback"
      ? fontFallbackTextSegments(run.text)
      : [{ text: run.text, start: 0, end: run.text.length }];
    for (const segment of segments) {
      const semanticSegment = semanticSymbolFontText(segment.text, run.fontFamily);
      const resolution = segments.length === 1 ? wholeResolution : resolveSegment(segment.text);
      runs.push(Object.freeze({
        start: start + segment.start,
        end: start + segment.start + semanticSegment.length,
        authoredFamily: run.fontFamily,
        renderedFamily: resolution.family,
        source: resolution.source,
        ...((object.source.format === "rtf" || object.type === "cell" || object.source.format === "xlsx")
          && run.fontSize !== undefined
          ? { fontSize: run.fontSize }
          : {}),
        ...(object.type === "cell" ? { bold: run.bold, italic: run.italic } : {}),
      }));
    }
    cursor = start + semanticText.length;
  }
  return runs.length === 0 ? undefined : Object.freeze(runs);
}

export function compatibleFallbackFamily(
  format: DocumentFormat,
  authoredFamily: string,
): string | undefined {
  const normalized = authoredFamily.trim().toLowerCase();
  if (format === "rtf" || format === "odt") return TEXT_DOCUMENT_FALLBACKS.get(normalized);
  if (format === "doc") return LEGACY_DOC_SHUSONG_FALLBACKS.get(normalized);
  if (format === "ppt" || format === "odp") {
    return PRESENTATION_CJK_FALLBACKS.get(normalized);
  }
  return undefined;
}

export function fallbackSpaceAdvanceEm(
  format: DocumentFormat,
  authoredFamily: string,
): number | undefined {
  return format === "doc"
    && LEGACY_DOC_SHUSONG_FALLBACKS.has(authoredFamily.trim().toLowerCase())
    ? 0.5
    : undefined;
}

export function fallbackFontSizeFactor(
  format: DocumentFormat,
  authoredFamily: string,
): number | undefined {
  if (format !== "ppt" && format !== "odp") return undefined;
  return /^(?:等线|dengxian)\s+light$/iu.test(authoredFamily.trim())
    ? PRESENTATION_DENGXIAN_LIGHT_SIZE_FACTOR
    : undefined;
}

export function fallbackBaselineShift(
  format: DocumentFormat,
  authoredFamily: string,
  fontSize: number,
): number | undefined {
  if (
    (format !== "ppt" && format !== "odp")
    || !/^(?:等线|dengxian)\s+light$/iu.test(authoredFamily.trim())
  ) {
    return undefined;
  }
  return fontSize * PRESENTATION_DENGXIAN_LIGHT_BASELINE_SHIFT_FACTOR;
}

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
const SERIF_HINT = /(?:serif|times|cambria|georgia|garamond|caslon|nuptial(?:script)?|baskerville|didot|janson(?:text)?|nimbusrom(?:an|no9l)?|antiqua|palatino|ming|song|simsun|宋体|明朝|明體)/iu;
const MONOSPACE_HINT = /(?:mono|courier|consolas|menlo|monaco|\bcode\b|等宽|等寬)/iu;
const OFFICE_THEME_SANS_HINT = /^segoe ui\s*\((?:body|headings)\)$/iu;
const SEGOE_UI_SYMBOL_HINT = /^segoe\s*ui\s*symbol$/iu;
const CALIBRI_COMPATIBLE_HINT = /\b(?:aptos|calibri)\b/iu;
const ARIAL_COMPATIBLE_HINT = /^(?:[a-z]{6}\+)?arial(?:(?:unicode)?ms|(?:-(?:bolditalic|bold|italic))?mt)?$/iu;
const LUCIDA_SANS_HINT = /^(?:[a-z]{6}\+)?lucida\s*sans(?:[-\s].*)?$/iu;
const CENTURY_GOTHIC_HINT = /^(?:[a-z]{6}\+)?century\s*gothic\b/iu;
const BANK_GOTHIC_HINT = /^(?:[a-z]{6}\+)?bankgothic\b/iu;
const OFFICE_CJK_SONG_FALLBACK_HINT = /^avenir next(?:\s+(?:condensed|ultra light))?$/iu;
const YU_GOTHIC_HINT = /\byu gothic(?: ui)?\b/iu;
const TEX_SERIF_HINT = /^(?:[a-z]{6}\+)?cm(?:r|b|bx|ti|sl|mi|sy|ex)\d+$/iu;
const TEX_MONOSPACE_HINT = /^(?:[a-z]{6}\+)?cmtt\d+$/iu;
const TEX_SANS_HINT = /^(?:[a-z]{6}\+)?cmss(?:bx|dc|qi)?\d+$/iu;
const LOCAL_FONT_COMPATIBILITY_CANDIDATES: ReadonlyMap<string, readonly string[]> = new Map([
  ["宋体", ["SimSun", "NSimSun", "STSong", "Songti SC"]],
  ["simsun", ["NSimSun", "STSong", "Songti SC"]],
  ["nsimsun", ["SimSun", "STSong", "Songti SC"]],
  ["新宋体", ["NSimSun", "SimSun", "STSong", "Songti SC"]],
  ["楷体", ["KaiTi", "KaiTi_GB2312", "STKaiti", "Kaiti SC"]],
  ["楷体_gb2312", ["KaiTi_GB2312", "KaiTi", "STKaiti", "Kaiti SC"]],
  ["kaiti", ["KaiTi_GB2312", "STKaiti", "Kaiti SC"]],
  ["kaiti_gb2312", ["KaiTi", "STKaiti", "Kaiti SC"]],
  ["仿宋", ["FangSong", "FangSong_GB2312", "STFangsong"]],
  ["仿宋_gb2312", ["FangSong_GB2312", "FangSong", "STFangsong"]],
  ["fangsong", ["FangSong_GB2312", "STFangsong"]],
  ["黑体", ["SimHei", "STHeitiSC-Medium", "STHeiti", "Heiti SC"]],
  ["simhei", ["STHeitiSC-Medium", "STHeiti", "Heiti SC"]],
  ["微软雅黑", ["Microsoft YaHei", "PingFang SC"]],
  ["microsoft yahei", ["PingFang SC"]],
  ["等线", ["DengXian", "PingFangSC-Regular", "PingFang SC"]],
  ["dengxian", ["PingFangSC-Regular", "PingFang SC"]],
  ["新細明體", ["PMingLiU", "MingLiU", "Songti TC"]],
  ["新细明体", ["PMingLiU", "MingLiU", "Songti TC"]],
  ["pmingliu", ["MingLiU", "Songti TC"]],
  ["標楷體", ["DFKai-SB", "BiauKai", "Kaiti TC"]],
  ["标楷体", ["DFKai-SB", "BiauKai", "Kaiti TC"]],
  ["dfkai-sb", ["BiauKai", "Kaiti TC"]],
  ["微軟正黑體", ["Microsoft JhengHei", "PingFang TC"]],
  ["微软正黑体", ["Microsoft JhengHei", "PingFang TC"]],
  ["microsoft jhenghei", ["PingFang TC"]],
  ["ms gothic", ["Yu Gothic", "Hiragino Sans", "Meiryo"]],
  ["ms pgothic", ["Yu Gothic", "Hiragino Sans", "Meiryo"]],
  ["ｍｓ ゴシック", ["Yu Gothic", "Hiragino Sans", "Meiryo"]],
  ["ｍｓ ｐゴシック", ["Yu Gothic", "Hiragino Sans", "Meiryo"]],
  ["yu gothic", ["Hiragino Sans", "Meiryo"]],
  ["yu gothic ui", ["Hiragino Sans", "Meiryo UI", "Meiryo"]],
  ["meiryo", ["Hiragino Sans"]],
  ["meiryo ui", ["Hiragino Sans", "Meiryo"]],
  ["ms mincho", ["Yu Mincho", "Hiragino Mincho ProN", "Hiragino Mincho Pro"]],
  ["ms pmincho", ["Yu Mincho", "Hiragino Mincho ProN", "Hiragino Mincho Pro"]],
  ["ｍｓ 明朝", ["Yu Mincho", "Hiragino Mincho ProN", "Hiragino Mincho Pro"]],
  ["yu mincho", ["Hiragino Mincho ProN", "Hiragino Mincho Pro"]],
  ["malgun gothic", ["Apple SD Gothic Neo"]],
  ["맑은 고딕", ["Malgun Gothic", "Apple SD Gothic Neo"]],
  ["gulim", ["Apple SD Gothic Neo"]],
  ["굴림", ["Gulim", "Apple SD Gothic Neo"]],
  ["batang", ["AppleMyungjo"]],
  ["바탕", ["Batang", "AppleMyungjo"]],
  ["traditional arabic", ["Geeza Pro", "Noto Naskh Arabic"]],
  ["arabic typesetting", ["Geeza Pro", "Noto Naskh Arabic"]],
  ["simplified arabic", ["Geeza Pro", "Noto Sans Arabic"]],
  ["david", ["New Peninim MT", "Noto Serif Hebrew"]],
  ["arial hebrew", ["Noto Sans Hebrew"]],
  ["mangal", ["Kohinoor Devanagari", "Noto Sans Devanagari"]],
  ["nirmala ui", ["Kohinoor Devanagari", "Noto Sans Devanagari"]],
  ["aparajita", ["Kohinoor Devanagari", "Noto Serif Devanagari"]],
  ["leelawadee ui", ["Thonburi", "Noto Sans Thai"]],
  ["calibri", ["Carlito", "Aptos"]],
  ["calibri light", ["Carlito", "Aptos Display", "Aptos"]],
  ["aptos", ["Calibri", "Carlito"]],
  ["arial", ["Helvetica", "Liberation Sans"]],
  ["roboto condensed", ["Arial Narrow"]],
  ["times new roman", ["Times", "Liberation Serif"]],
  ["book antiqua", ["Palatino", "Palatino Linotype"]],
  ["palatino linotype", ["Palatino", "Book Antiqua"]],
  ["courier new", ["Courier", "Liberation Mono"]],
  ["segoe ui", ["Helvetica Neue", "Arial"]],
  ["futura", ["Futura-Medium"]],
  ["futura-boo", ["Futura-Medium"]],
  ["futura-dem", ["Futura-Bold"]],
  ["futura-lig", ["Futura-Medium"]],
  ["futura-bol", ["Futura-Bold"]],
  ["futuratot-demi", ["Futura-Bold"]],
  ["futuratlight", ["Futura-Medium"]],
  ["bankgothic lt bt", ["Copperplate"]],
]);
const PDF_PREVIEW_HELVETICA_FALLBACKS = new Set([
  "simsun",
  "kaiti",
  "simhei",
  "dengxian",
]);
const CALIBRI_MANAGED_FALLBACKS = ["Calibri", "Carlito", "Aptos"] as const;
const SEGOE_MANAGED_FALLBACKS = ["Segoe UI"] as const;
const JAPANESE_GOTHIC_MANAGED_FALLBACKS = [
  "Hiragino Sans",
  "Yu Gothic UI",
  "Yu Gothic",
  "游ゴシック",
  "Meiryo UI",
  "Meiryo",
] as const;
const SERIF_MANAGED_FALLBACKS = ["Cambria", "Times New Roman"] as const;
const MONOSPACE_MANAGED_FALLBACKS = ["Consolas", "Courier New"] as const;
const SYMBOL_MANAGED_FALLBACKS = ["Segoe UI Symbol", "Apple Symbols", "Noto Sans Symbols"] as const;
const SONG_MANAGED_FALLBACKS = ["宋体", "SimSun", "NSimSun", "STSong", "Songti SC"] as const;
const NO_MANAGED_FALLBACKS: readonly string[] = Object.freeze([]);

function normalizeFamily(family: string): string {
  return family.trim().replace(/^(['"])(.*)\1$/u, "$2").toLowerCase();
}

function pdfObjectSourceFamily(family: string): string | undefined {
  const match = /^(.*?)\s+PDF\s+[0-9]+\s+[0-9]+$/u.exec(family.trim());
  const source = match?.[1]?.trim();
  if (source === undefined || source === "") return undefined;
  const canonical = source.replace(/^[A-Za-z]{6}\+/u, "");
  return canonical === "" ? undefined : canonical;
}

function localFontFamilyCandidates(family: string): readonly string[] {
  const candidates = [
    family,
    ...(LOCAL_FONT_COMPATIBILITY_CANDIDATES.get(normalizeFamily(family)) ?? []),
  ];
  const seen = new Set<string>();
  return candidates.filter((candidate) => {
    const normalized = normalizeFamily(candidate);
    if (seen.has(normalized)) return false;
    seen.add(normalized);
    return true;
  });
}

function approximateFontFamilyBase(family: string): string {
  const normalizedFamily = normalizeFamily(family);
  if (OFFICE_THEME_SANS_HINT.test(normalizedFamily)) return "Segoe UI";
  // Segoe UI Symbol also contains ordinary Latin text. Keep those runs on a
  // metric-compatible sans face; symbol-only runs are routed back to the
  // symbol fallback set by `adjustApproximateFontFamily` below.
  if (SEGOE_UI_SYMBOL_HINT.test(normalizedFamily)) return "Arial";
  if (CALIBRI_COMPATIBLE_HINT.test(normalizedFamily)) return "Calibri";
  if (ARIAL_COMPATIBLE_HINT.test(normalizedFamily) || normalizedFamily === "liberation sans") return "Arial";
  if (LUCIDA_SANS_HINT.test(normalizedFamily)) return "Arial";
  if (CENTURY_GOTHIC_HINT.test(normalizedFamily)) return "Arial";
  if (BANK_GOTHIC_HINT.test(normalizedFamily)) return "Copperplate";
  if (YU_GOTHIC_HINT.test(normalizedFamily)) return "Hiragino Sans";
  if (TEX_MONOSPACE_HINT.test(normalizedFamily)) return "monospace";
  if (TEX_SERIF_HINT.test(normalizedFamily)) return "serif";
  if (TEX_SANS_HINT.test(normalizedFamily)) return "Arial";
  if (MONOSPACE_HINT.test(normalizedFamily)) return "monospace";
  if (SERIF_HINT.test(normalizedFamily)) return "serif";
  return "Calibri";
}

export function approximateFontFamily(
  family: string,
  codePoints?: readonly number[],
): string {
  const demand = codePoints === undefined ? undefined : classifyFontScriptDemand(codePoints);
  if (demand === "cjk" && OFFICE_CJK_SONG_FALLBACK_HINT.test(normalizeFamily(family))) {
    return "SimSun";
  }
  const approximateFamily = approximateFontFamilyBase(family);
  if (demand === undefined) return approximateFamily;
  return adjustApproximateFontFamily(approximateFamily, demand);
}

function adjustApproximateFontFamily(
  approximateFamily: string,
  demand: FontScriptDemand,
): string {
  if (demand === "cjk" && normalizeFamily(approximateFamily) === "calibri") return "Hiragino Sans";
  if (demand === "complex" && normalizeFamily(approximateFamily) === "calibri") return "Arial";
  if (demand === "symbol") return "Segoe UI Symbol";
  return approximateFamily;
}

function approximateFontFamilyForDemand(
  family: string,
  demand: FontScriptDemand,
): string {
  if (demand === "cjk" && OFFICE_CJK_SONG_FALLBACK_HINT.test(family)) return "SimSun";
  return adjustApproximateFontFamily(approximateFontFamilyBase(family), demand);
}

function managedFallbackFamilyNames(
  approximateFamily: string,
  demand: FontScriptDemand,
): readonly string[] {
  const normalized = normalizeFamily(approximateFamily);
  if (demand === "mixed" || demand === "complex") return NO_MANAGED_FALLBACKS;
  if (demand === "symbol") return SYMBOL_MANAGED_FALLBACKS;
  if (normalized === "calibri") return CALIBRI_MANAGED_FALLBACKS;
  if (normalized === "segoe ui") return SEGOE_MANAGED_FALLBACKS;
  if (normalized === "hiragino sans") return JAPANESE_GOTHIC_MANAGED_FALLBACKS;
  if (normalized === "simsun") return SONG_MANAGED_FALLBACKS;
  if (normalized === "serif") return SERIF_MANAGED_FALLBACKS;
  if (normalized === "monospace") return MONOSPACE_MANAGED_FALLBACKS;
  return NO_MANAGED_FALLBACKS;
}

function fontLoadErrorMessage(cause: unknown): string {
  const message = cause instanceof Error
    ? cause.message
    : typeof cause === "string"
      ? cause
      : "Unknown font loading failure";
  return message.slice(0, 512);
}

function registeredFontDiagnosticStore(registration: object): Map<string, Diagnostic> {
  const existing = REGISTERED_FONT_DIAGNOSTICS.get(registration);
  if (existing !== undefined) return existing;
  const diagnostics = new Map<string, Diagnostic>();
  REGISTERED_FONT_DIAGNOSTICS.set(registration, diagnostics);
  return diagnostics;
}

function recordFontLoadDiagnostic(
  diagnostics: Map<string, Diagnostic>,
  code: "FONT_ENVIRONMENT_UNAVAILABLE" | "FONT_LOAD_FAILED" | "FONT_LOAD_TIMEOUT",
  source: BinaryFontSource,
  asset: PreparedFontAsset,
  cause: unknown,
): void {
  const key = `${code}\0${source}\0${exactFaceKey(asset)}`;
  if (diagnostics.has(key)) return;
  if (diagnostics.size >= MAX_FONT_LOAD_DIAGNOSTICS - 1) {
    if (!diagnostics.has(FONT_LOAD_DIAGNOSTICS_TRUNCATED_KEY)) {
      const [diagnostic] = immutableDiagnostics([{
        code: "FONT_LOAD_DIAGNOSTICS_TRUNCATED",
        severity: "warning",
        fidelity: "approximate",
        phase: "render",
        message: "Additional font loading failures were omitted after the diagnostic limit was reached",
        details: {
          limit: MAX_FONT_LOAD_DIAGNOSTICS,
          retained: MAX_FONT_LOAD_DIAGNOSTICS - 1,
        },
      }]);
      if (diagnostic !== undefined) {
        diagnostics.set(FONT_LOAD_DIAGNOSTICS_TRUNCATED_KEY, diagnostic);
      }
    }
    return;
  }
  const label = `${source[0]!.toUpperCase()}${source.slice(1)}`;
  const [diagnostic] = immutableDiagnostics([{
    code,
    severity: "warning",
    fidelity: "approximate",
    phase: "render",
    message: code === "FONT_ENVIRONMENT_UNAVAILABLE"
      ? `${label} font "${asset.family}" could not be registered; this attempt was skipped`
      : `${label} font "${asset.family}" could not be loaded; this attempt was skipped`,
    details: {
      source,
      family: asset.family,
      style: asset.style,
      weight: asset.weight,
      stretch: asset.stretch,
      bytes: asset.bytes.byteLength,
      error: fontLoadErrorMessage(cause),
    },
  }]);
  if (diagnostic !== undefined) diagnostics.set(key, diagnostic);
}

async function loadBinaryFonts(
  runtime: FontRuntime,
  assets: readonly PreparedFontAsset[],
  source: BinaryFontSource,
  fatal: boolean,
  registeredFaces: RuntimeFontFace[],
  resolutions: Map<string, ResolvedFontFace>,
  familyDefaults: Map<string, ResolvedFontFace>,
  registeredAssets: RegisteredBinaryFont[],
  aliases: FontAliasRegistry,
  timeoutMs: number,
  signal: AbortSignal | undefined,
  diagnostics: Map<string, Diagnostic>,
): Promise<void> {
  if (assets.length === 0) return;
  const canRegister = (key: string): boolean => {
    const existing = resolutions.get(key);
    return existing === undefined || existing.source === "fallback" && source !== "fallback";
  };
  const seen = new Set<string>();
  const candidates = assets.flatMap((asset) => {
    const key = exactFaceKey(asset);
    if (!canRegister(key) || seen.has(key)) return [];
    seen.add(key);
    const alias = ensureAlias(aliases, asset.family, source);
    const face = new runtime.FontFace(alias, asset.bytes, {
      style: asset.style,
      weight: String(asset.weight),
      stretch: asset.stretch,
    });
    return [{ asset, key, alias, face }];
  });
  if (candidates.length === 0) return;
  for (const { face } of candidates) {
    runtime.fontSet.add(face);
    registeredFaces.push(face);
  }
  let results: PromiseSettledResult<unknown>[];
  try {
    results = await settleFontOperations(
      candidates.map(({ face }) => () => face.load()),
      timeoutMs,
      signal,
      `${source[0]!.toUpperCase()}${source.slice(1)} font loading`,
    );
  } catch (cause) {
    for (const { face } of candidates) removeFace(runtime.fontSet, registeredFaces, face);
    throw cause;
  }
  const failure = results.find((result): result is PromiseRejectedResult => result.status === "rejected");
  if (fatal && failure !== undefined) throw failure.reason;
  for (let index = 0; index < candidates.length; index += 1) {
    const candidate = candidates[index];
    if (candidate === undefined) continue;
    if (results[index]?.status !== "fulfilled") {
      removeFace(runtime.fontSet, registeredFaces, candidate.face);
      const result = results[index];
      recordFontLoadDiagnostic(
        diagnostics,
        result?.status === "rejected"
          && result.reason instanceof OfficeEngineError
          && result.reason.code === "FONT_LOAD_TIMEOUT"
          ? "FONT_LOAD_TIMEOUT"
          : "FONT_LOAD_FAILED",
        source,
        candidate.asset,
        result?.status === "rejected" ? result.reason : "Unknown font loading failure",
      );
      continue;
    }
    const { asset, key, alias } = candidate;
    registeredAssets.push({ asset, source });
    if (canRegister(key)) recordResolution(resolutions, familyDefaults, asset, alias, source);
  }
}

function removeFace(
  fontSet: RuntimeFontSet,
  registeredFaces: RuntimeFontFace[],
  face: RuntimeFontFace,
): void {
  try {
    fontSet.delete(face);
  } catch {
    // Realm teardown can race tolerant embedded-font cleanup.
  }
  const index = registeredFaces.indexOf(face);
  if (index !== -1) registeredFaces.splice(index, 1);
}

async function resolveLocalFonts(
  runtime: FontRuntime,
  requestedFaces: readonly RequestedFontFace[],
  registeredFaces: RuntimeFontFace[],
  resolutions: Map<string, ResolvedFontFace>,
  familyDefaults: Map<string, ResolvedFontFace>,
  aliases: FontAliasRegistry,
  timeoutMs: number,
  signal: AbortSignal | undefined,
  fontAlternateNames: ReadonlyMap<string, readonly string[]>,
): Promise<void> {
  const localCandidates: {
    readonly requested: RequestedFontFace;
    readonly key: string;
    readonly alias: string;
    readonly localFamily: string;
    readonly declared: boolean;
    readonly face: RuntimeFontFace;
  }[] = [];
  for (const requested of requestedFaces) {
    const key = exactFaceKey(requested);
    if (resolutions.has(key)) continue;
    const normalizedFamily = normalizeFamily(requested.family);
    const localSourceFamily = pdfObjectSourceFamily(requested.family) ?? requested.family;
    if (GENERIC_FAMILIES.has(normalizedFamily)) {
      recordResolution(resolutions, familyDefaults, requested, normalizedFamily, "browser");
      continue;
    }
    const pdfSourceFamily = pdfObjectSourceFamily(requested.family);
    const declaredNames = fontAlternateNames.get(normalizedFamily) ?? [];
    const localFamilies = pdfSourceFamily !== undefined
      && PDF_PREVIEW_HELVETICA_FALLBACKS.has(normalizeFamily(pdfSourceFamily))
      ? []
      : [...new Set([localSourceFamily, ...declaredNames, ...localFontFamilyCandidates(localSourceFamily)])];
    for (const localFamily of localFamilies) {
      const exact = normalizeFamily(localFamily) === normalizeFamily(localSourceFamily);
      const declared = !exact && declaredNames.includes(localFamily);
      const alias = ensureAlias(
        aliases,
        exact ? requested.family : `${declared ? "alternate" : "compatible"}:${localFamily}`,
        "browser",
      );
      const localFace = exact || declared
        ? requested
        : { ...requested, style: "normal" as const, weight: 400, stretch: "normal" as const };
      const face = new runtime.FontFace(alias, `local(${JSON.stringify(localFamily)})`, {
        style: localFace.style,
        weight: String(localFace.weight),
        stretch: localFace.stretch,
      });
      localCandidates.push({ requested, key, alias, localFamily, declared, face });
    }
  }
  if (localCandidates.length === 0) return;
  const results = await settleFontOperations(
    localCandidates.map(({ face }) => () => face.load()),
    timeoutMs,
    signal,
    "Local font probing",
  );
  const probeContext = ambientFontProbeContext();
  for (let index = 0; index < localCandidates.length; index += 1) {
    const candidate = localCandidates[index];
    if (candidate === undefined) continue;
    const { requested, key, alias, localFamily, declared, face } = candidate;
    if (resolutions.has(key)) continue;
    const registered = resolutions.get(exactFaceKey({ ...requested, family: localFamily }))
      ?? familyDefaults.get(normalizeFamily(localFamily));
    if (declared && registered !== undefined && ["embedded", "host", "provider"].includes(registered.source)) {
      recordResolution(resolutions, familyDefaults, requested, registered.family, "fallback", registered.spaceAdvanceEm, localFamily);
      continue;
    }
    if (results[index]?.status !== "fulfilled") {
      // local() accepts face names, while CSS selects families and traits.
      // A measured CSS face remains usable even without a loadable family alias.
      if ((declared || normalizeFamily(localFamily) === normalizeFamily(requested.family))
        && probeBrowserFontFace({ ...requested, family: localFamily }, probeContext) === true) {
        recordResolution(resolutions, familyDefaults, requested, localFamily, declared ? "fallback" : "browser", undefined, declared ? localFamily : undefined);
      }
      continue;
    }
    runtime.fontSet.add(face);
    registeredFaces.push(face);
    const requestedText = boundedFontProbeText(requested.codePoints);
    const aliasFont = exactProbeFontShorthand(alias, requested);
    const authoredFamily = declared ? localFamily : requested.family;
    const authoredFont = exactProbeFontShorthand(authoredFamily, requested);
    let aliasChecked = true;
    let authoredChecked = true;
    try {
      aliasChecked = runtime.fontSet.check(aliasFont, requestedText);
      authoredChecked = runtime.fontSet.check(authoredFont, requestedText);
    } catch {
      // Some FontFaceSet shims only implement registration, so Canvas remains
      // the authoritative false-positive probe when check() is unavailable.
    }
    const aliasAvailable = aliasChecked
      ? probeBrowserFontFace({ ...requested, family: alias }, probeContext, requestedText)
      : false;
    const exact = normalizeFamily(localFamily) === normalizeFamily(requested.family);
    const authoredAvailable = (exact || declared) && authoredChecked
      ? probeBrowserFontFace({ ...requested, family: authoredFamily }, probeContext, requestedText)
      : false;
    const compatibleSourceAvailable = exact || declared
      ? undefined
      : probeBrowserFontFace(
        { ...requested, family: localFamily, style: "normal", weight: 400, stretch: "normal" },
        probeContext,
        requestedText,
      );
    const aliasMatchesCompatibleSource = exact || declared
      ? undefined
      : fontFaceMetricsMatch(
        alias,
        localFamily,
        { style: "normal", weight: 400 },
        probeContext,
        requestedText,
      );
    const unavailable = exact || declared
      ? aliasAvailable === false && authoredAvailable === false
      : aliasAvailable === false
        || compatibleSourceAvailable === false
        || aliasMatchesCompatibleSource === false;
    if (unavailable) {
      removeFace(runtime.fontSet, registeredFaces, face);
      continue;
    }
    if (!resolutions.has(key)) {
      recordResolution(
        resolutions,
        familyDefaults,
        requested,
        (exact || declared) && authoredAvailable !== false ? authoredFamily : alias,
        declared ? "fallback" : "browser",
        // SimSun/NSimSun have half-em spaces; the macOS Song substitutes use
        // quarter-em spaces. Keep form blanks stable without stretching glyphs.
        /^(?:宋体|新宋体|simsun|nsimsun)$/iu.test(requested.family)
          && /^(?:stsong|songti sc)$/iu.test(localFamily) ? 0.5 : undefined,
        declared ? localFamily : undefined,
      );
    }
  }
}

async function settleFontOperations(
  operations: readonly (() => PromiseLike<unknown> | unknown)[],
  timeoutMs: number,
  signal: AbortSignal | undefined,
  label: string,
): Promise<PromiseSettledResult<unknown>[]> {
  let abort: (() => void) | undefined;
  let abortError: OfficeEngineError | undefined;
  const abortPending = signal === undefined
    ? undefined
    : new Promise<never>((_, reject) => {
        abort = () => {
          abortError = new OfficeEngineError(
            "OPERATION_ABORTED",
            "The operation was aborted",
            { cause: signal.reason },
          );
          reject(abortError);
        };
        signal.addEventListener("abort", abort, { once: true });
        if (signal.aborted) abort();
      });
  try {
    const results = await Promise.all(operations.map(async (operation, index) => {
      let timer: ReturnType<typeof setTimeout> | undefined;
      try {
        let operationPending: Promise<unknown>;
        try {
          operationPending = Promise.resolve(operation());
        } catch (reason) {
          operationPending = Promise.reject(reason);
        }
        const pending: Promise<unknown>[] = [
          operationPending,
          new Promise<never>((_, reject) => {
            timer = setTimeout(() => reject(new OfficeEngineError(
              "FONT_LOAD_TIMEOUT",
              `${label} face ${index + 1} exceeded ${timeoutMs}ms`,
            )), timeoutMs);
          }),
        ];
        if (abortPending !== undefined) pending.push(abortPending);
        return {
          status: "fulfilled",
          value: await Promise.race(pending),
        } as const;
      } catch (reason) {
        return { status: "rejected", reason } as const;
      } finally {
        if (timer !== undefined) clearTimeout(timer);
      }
    }));
    if (signal?.aborted === true) {
      throw abortError ?? new OfficeEngineError(
        "OPERATION_ABORTED",
        "The operation was aborted",
        { cause: signal.reason },
      );
    }
    return results;
  } finally {
    if (abort !== undefined) signal?.removeEventListener("abort", abort);
  }
}

export function collectFontFamilies(objects: readonly SceneObject[]): readonly string[] {
  const families = new Set<string>();
  for (const object of objects) {
    for (const run of sceneTextRuns(object)) families.add(run.fontFamily);
  }
  return [...families];
}

/** Collects exact face and Unicode-scalar demand before lazy provider resolution. */
export function collectFontRequests(
  objects: readonly SceneObject[],
  fontAlternateNames: SceneDocument["fontAlternateNames"] = [],
): readonly FontRequest[] {
  const requests: FontRequest[] = [];
  const add = (
    family: string,
    text: string,
    italic: boolean,
    bold: boolean,
  ): void => {
    if (text.length === 0) return;
    const codePoints = new Set<number>();
    // PUA glyphs must be probed in the authored font; their semantic substitute
    // may not exist in that font (e.g. Wingdings' wide-barbed arrows).
    const probeText = /[\uf000-\uf0ff]/u.test(text) ? text : semanticSymbolFontText(text, family);
    for (const character of probeText) {
      codePoints.add(character.codePointAt(0)!);
    }
    requests.push({
      family,
      style: italic ? "italic" : "normal",
      weight: bold ? 700 : 400,
      stretch: "normal",
      codePoints: [...codePoints],
    });
  };
  for (const object of objects) {
    for (const run of sceneTextRuns(object)) {
      add(run.fontFamily, run.text, run.italic, run.bold);
    }
  }
  const alternates = new Map(fontAlternateNames.map(({ family, names }) => [normalizeFamily(family), names]));
  return mergeFontRequests(requests.flatMap(request => [request,
    ...(alternates.get(normalizeFamily(request.family)) ?? []).map(family => ({ ...request, family })),
  ]).map(normalizeRequest));
}

export function fontShorthand(
  family: string,
  size: number,
  italic: boolean,
  bold: boolean,
): string {
  const normalized = normalizeFamily(family);
  let cssFamily: string;
  let weight = bold ? 700 : 400;
  if (GENERIC_FAMILIES.has(normalized)) cssFamily = normalized;
  else if (normalized === "pingfang sc light") {
    weight = bold ? 700 : 300;
    cssFamily = '"PingFang SC", "Microsoft YaHei UI", "Microsoft YaHei", "Hiragino Sans GB", sans-serif';
  }
  else if (normalized === "pingfang sc") {
    cssFamily = '"PingFang SC", "Microsoft YaHei UI", "Microsoft YaHei", "Hiragino Sans GB", sans-serif';
  }
  else if (normalized === "hiragino sans") {
    cssFamily = `${JSON.stringify(family)}, "Yu Gothic UI", "Yu Gothic", "Meiryo UI", "Meiryo", "Microsoft YaHei", sans-serif`;
  }
  else if (normalized === "hiragino mincho pron") {
    cssFamily = `${JSON.stringify(family)}, "Hiragino Mincho Pro", "Yu Mincho", "MS Mincho", "STSong", "Songti SC", "Noto Serif CJK SC", serif`;
  }
  else if (normalized === "calibri") {
    cssFamily = `${JSON.stringify(family)}, "Carlito", "Aptos", "Segoe UI", "Arial", "Arial Unicode MS", "Microsoft YaHei", "Yu Gothic UI", "Meiryo UI", "Hiragino Sans", "Noto Sans Arabic", "Noto Sans Hebrew", "Noto Sans Devanagari", "Noto Sans Thai", "Segoe UI Symbol", "Apple Symbols", sans-serif`;
  } else if (normalized === "segoe ui") {
    cssFamily = `${JSON.stringify(family)}, "Arial", "Arial Unicode MS", "Microsoft YaHei", "Yu Gothic UI", "Meiryo UI", "Hiragino Sans", "Noto Sans Arabic", "Noto Sans Hebrew", "Noto Sans Devanagari", "Noto Sans Thai", "Segoe UI Symbol", "Apple Symbols", sans-serif`;
  } else if (normalized === "arial") {
    cssFamily = `${JSON.stringify(family)}, "Arial Unicode MS", "Noto Sans Arabic", "Noto Sans Hebrew", "Noto Sans Devanagari", "Noto Sans Thai", "Microsoft YaHei", "Yu Gothic UI", "Meiryo UI", "Hiragino Sans", "Segoe UI Symbol", "Apple Symbols", sans-serif`;
  } else if (normalized === "segoe ui symbol") {
    cssFamily = `${JSON.stringify(family)}, "Apple Symbols", "Noto Sans Symbols", "Noto Sans Symbols 2", "Arial Unicode MS", "Segoe UI Emoji", "Apple Color Emoji", sans-serif`;
  } else cssFamily = JSON.stringify(family);
  return `${italic ? "italic " : ""}${weight} ${size}px ${cssFamily}`;
}
