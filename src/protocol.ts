import { byteFingerprint, bytesEqual } from "./bytes.js";
import type {
  Diagnostic,
  DiagnosticFidelity,
  DiagnosticPhase,
  DiagnosticSeverity,
  DocumentAction,
  DocumentActionKind,
  DocumentActionTrigger,
  DocumentFormat,
  DocumentKind,
  DocumentOutlineItem,
  DocumentObjectType,
  MappingQuality,
  SheetAxis,
  SheetPrintSettings,
  SpeakerNoteParagraph,
  SpeakerNoteRun,
  SourceRef,
  UnitDescriptor,
} from "./types.js";
import type {
  SceneDocument,
  SceneEmbeddedFont,
  SceneGeometry,
  SceneObject,
  ScenePaint,
  ScenePathCommand,
  SceneTextAlign,
  SceneVisual,
  SceneXpsColor,
} from "./scene.js";
import { OfficeEngineError } from "./types.js";

const MAGIC = 0x3144564f;
const MIN_VERSION = 3;
const VERSION = 76;
const NONE = 0xffff_ffff;
const MAX_VISUAL_DEPTH = 64;
const MAX_PATH_COMMANDS = 65_536;
const MAX_GRADIENT_STOPS = 4_096;
const MAX_VISUAL_BRUSH_CHILDREN = 65_536;
const MAX_RICH_TEXT_RUNS = 100_000;
const MAX_TAB_STOPS = 256;
const MAX_TEXT_PARAGRAPHS = 100_000;
const LEGACY_TEXT_ALIGNS: readonly SceneTextAlign[] = [
  "start",
  "center",
  "end",
  "justify",
];
const TEXT_ALIGNS: readonly SceneTextAlign[] = [
  ...LEGACY_TEXT_ALIGNS,
  "distribute",
  "medium-kashida",
  "high-kashida",
  "low-kashida",
  "thai-distribute",
];

const FORMATS: readonly (DocumentFormat | undefined)[] = [
  undefined,
  "pptx",
  "odp",
  "xlsx",
  "ods",
  "docx",
  "odt",
  undefined,
  "csv",
  undefined,
  "rtf",
  "ppt",
  "xls",
  "doc",
  "pages",
  "numbers",
  "keynote",
  "pdf",
  "xps",
  "ofd",
];
const KINDS: readonly (DocumentKind | undefined)[] = [undefined, "presentation", "spreadsheet", "text"];
const UNIT_TYPES = [undefined, "slide", "sheet", "page"] as const;
const OBJECT_TYPES: readonly (DocumentObjectType | undefined)[] = [
  undefined,
  "group",
  "text-box",
  "paragraph",
  "image",
  "shape",
  "table",
  "cell",
  "unknown",
];
const SEVERITIES: readonly DiagnosticSeverity[] = ["info", "warning", "error", "fatal"];
const FIDELITIES: readonly DiagnosticFidelity[] = ["exact", "approximate", "unsupported", "not-rendered"];
const PHASES: readonly DiagnosticPhase[] = ["identify", "package", "parse", "layout", "render", "security"];
const MAPPINGS: readonly MappingQuality[] = ["exact", "derived", "approximate"];
const LEGACY_GEOMETRIES = ["rectangle", "ellipse", "line"] as const;
const FONT_STYLES = ["normal", "italic", "oblique"] as const;

class Reader {
  readonly #view: DataView;
  readonly #bytes: Uint8Array;
  readonly #binaryResources = new Map<string, Uint8Array[]>();
  readonly #imageResources = new Map<number, Pick<Extract<SceneVisual, { kind: "image" }>, "mediaType" | "bytes">>();
  readonly #imageMasks = new Map<number, NonNullable<Extract<SceneVisual, { kind: "image" }>["alphaMask"]>>();
  #offset = 0;

  constructor(bytes: Uint8Array) {
    this.#bytes = bytes;
    this.#view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  get remaining(): number {
    return this.#bytes.length - this.#offset;
  }

  u8(): number {
    this.#need(1);
    return this.#view.getUint8(this.#offset++);
  }

  u16(): number {
    this.#need(2);
    const value = this.#view.getUint16(this.#offset, true);
    this.#offset += 2;
    return value;
  }

  u32(): number {
    this.#need(4);
    const value = this.#view.getUint32(this.#offset, true);
    this.#offset += 4;
    return value;
  }

  i32(): number {
    this.#need(4);
    const value = this.#view.getInt32(this.#offset, true);
    this.#offset += 4;
    return value;
  }

  f32(): number {
    this.#need(4);
    const value = this.#view.getFloat32(this.#offset, true);
    this.#offset += 4;
    return value;
  }

  string(): string {
    const length = this.u32();
    if (length === NONE) {
      throw this.#invalid("required string is absent");
    }
    this.#need(length);
    const value = this.#decodeString(length);
    this.#offset += length;
    return value;
  }

  optionalString(): string | undefined {
    const length = this.u32();
    if (length === NONE) return undefined;
    this.#need(length);
    const value = this.#decodeString(length);
    this.#offset += length;
    return value;
  }

  byteArray(): Uint8Array {
    const length = this.u32();
    this.#need(length);
    const source = this.#bytes.subarray(this.#offset, this.#offset + length);
    this.#offset += length;
    if (length < 4_096) return source.slice();
    const key = byteFingerprint(source);
    const candidates = this.#binaryResources.get(key);
    if (candidates !== undefined) {
      for (const candidate of candidates) {
        if (bytesEqual(candidate, source)) return candidate;
      }
    }
    const value = source.slice();
    if (candidates === undefined) this.#binaryResources.set(key, [value]);
    else candidates.push(value);
    return value;
  }

  imageMask(
    resourceId: number,
    mask: NonNullable<Extract<SceneVisual, { kind: "image" }>["alphaMask"]>,
  ): NonNullable<Extract<SceneVisual, { kind: "image" }>["alphaMask"]> {
    if (mask.bytes.length !== 0) {
      if (resourceId !== 0) this.#imageMasks.set(resourceId, mask);
      return mask;
    }
    const cached = this.#imageMasks.get(resourceId);
    if (resourceId === 0 || cached === undefined || cached.width !== mask.width
      || cached.height !== mask.height || cached.mediaType !== mask.mediaType) {
      throw this.#invalid("unresolved image alpha mask reference");
    }
    return cached;
  }

  imageResource(
    resourceId: number,
    image: Pick<Extract<SceneVisual, { kind: "image" }>, "mediaType" | "bytes">,
  ): Pick<Extract<SceneVisual, { kind: "image" }>, "mediaType" | "bytes"> {
    if (image.bytes.length !== 0) {
      if (resourceId !== 0) this.#imageResources.set(resourceId, image);
      return image;
    }
    const cached = this.#imageResources.get(resourceId);
    if (resourceId === 0 || cached === undefined || cached.mediaType !== image.mediaType) {
      throw this.#invalid("unresolved image resource reference");
    }
    return cached;
  }

  skip(length: number): void {
    this.#need(length);
    this.#offset += length;
  }

  #need(length: number): void {
    if (!Number.isSafeInteger(length) || length < 0 || length > this.remaining) {
      throw this.#invalid("truncated or oversized field");
    }
  }

  #decodeString(length: number): string {
    try {
      return new TextDecoder("utf-8", { fatal: true }).decode(
        this.#bytes.subarray(this.#offset, this.#offset + length),
      );
    } catch (cause) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid core snapshot: string is not valid UTF-8", {
        cause,
      });
    }
  }

  #invalid(reason: string): OfficeEngineError {
    return new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid core snapshot: ${reason}`);
  }
}

function enumValue<T>(values: readonly (T | undefined)[], code: number, field: string): T {
  const value = values[code];
  if (value === undefined) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid ${field} code ${code}`);
  }
  return value;
}

function readSheetAxis(reader: Reader): SheetAxis {
  const defaultSize = reader.f32();
  const count = reader.u32();
  if (count > Math.floor(reader.remaining / 12)) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid core snapshot: oversized sheet axis");
  }
  const spans: { start: number; end: number; size: number }[] = [];
  for (let index = 0; index < count; index += 1) {
    spans.push({ start: reader.u32(), end: reader.u32(), size: reader.f32() });
  }
  return { defaultSize, spans };
}

function uniformSheetAxis(count: number, totalSize: number): SheetAxis {
  return { defaultSize: count === 0 ? 0 : totalSize / count, spans: [] };
}

function textAlignValue(reader: Reader, version: number, field: string): SceneTextAlign {
  return enumValue(
    version >= 18 ? TEXT_ALIGNS : LEGACY_TEXT_ALIGNS,
    reader.u8(),
    field,
  );
}

function readTextRun(reader: Reader, version: number, depth = 0): SpeakerNoteRun & { paint?: ScenePaint } {
  const text = reader.string();
  const fontFamily = reader.string();
  const fontSize = reader.f32();
  const color = reader.u32();
  const flags = reader.u8();
  reader.skip(3);
  const letterSpacing = reader.f32();
  const highlight = reader.u32();
  const baselineShift = reader.f32();
  const horizontalScale = version >= 56 ? reader.f32() : 1;
  const paint = version >= 70 ? readPaint(reader, version, depth + 1) : undefined;
  return {
    text,
    fontFamily,
    fontSize,
    color,
    bold: (flags & 1) !== 0,
    italic: (flags & 2) !== 0,
    underline: (flags & 4) !== 0,
    strikethrough: (flags & 8) !== 0,
    highlight,
    baselineShift,
    letterSpacing,
    horizontalScale,
    ...(paint === undefined || (flags & 32) === 0 ? {} : { paint }),
    ...(version >= 63 && (flags & 16) !== 0 ? { eastAsianLineBreaks: false } : {}),
  };
}

function readSpeakerNoteParagraph(reader: Reader, version: number): SpeakerNoteParagraph {
  const align = textAlignValue(reader, version, "speaker-note paragraph alignment");
  reader.skip(3);
  const marginLeft = reader.f32();
  const marginRight = reader.f32();
  const firstLineIndent = reader.f32();
  const defaultTabStop = reader.f32();
  const lineHeight = reader.f32();
  const spaceBefore = reader.f32();
  const spaceAfter = reader.f32();
  const flags = reader.u8();
  reader.skip(3);
  const runCount = reader.u32();
  if (runCount > MAX_RICH_TEXT_RUNS || runCount > Math.floor(reader.remaining / 32)) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid speaker-note run count");
  }
  const runs: SpeakerNoteRun[] = [];
  for (let index = 0; index < runCount; index += 1) runs.push(readTextRun(reader, version));
  return {
    align,
    marginLeft,
    marginRight,
    firstLineIndent,
    defaultTabStop,
    lineHeight,
    spaceBefore,
    spaceAfter,
    latinLineBreak: (flags & 1) !== 0,
    hangingPunctuation: (flags & 2) !== 0,
    runs,
  };
}

function readFields(reader: Reader): ReadonlyMap<string, string | number> {
  const count = reader.u8();
  const fields = new Map<string, string | number>();
  for (let index = 0; index < count; index += 1) {
    const key = reader.string();
    if (fields.has(key)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Duplicate protocol field ${key}`);
    }
    const type = reader.u8();
    if (type === 0) fields.set(key, reader.string());
    else if (type === 1) fields.set(key, reader.u32());
    else if (type === 2) fields.set(key, reader.i32());
    else throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid source field type ${type}`);
  }
  return fields;
}

function optionalStringField(
  fields: ReadonlyMap<string, string | number>,
  key: string,
  locator: string,
): string | undefined {
  const value = fields.get(key);
  if (value === undefined) return undefined;
  if (typeof value !== "string") {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid ${locator} source field ${key}`);
  }
  return value;
}

function requiredStringField(
  fields: ReadonlyMap<string, string | number>,
  key: string,
  locator: string,
): string {
  const value = optionalStringField(fields, key, locator);
  if (value === undefined) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Missing ${locator} source field ${key}`);
  }
  return value;
}

function optionalUnsignedField(
  fields: ReadonlyMap<string, string | number>,
  key: string,
  locator: string,
): number | undefined {
  const value = fields.get(key);
  if (value === undefined) return undefined;
  if (typeof value !== "number" || !Number.isSafeInteger(value) || value < 0 || value > 0xffff_ffff) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid ${locator} source field ${key}`);
  }
  return value;
}

function requiredUnsignedField(
  fields: ReadonlyMap<string, string | number>,
  key: string,
  locator: string,
): number {
  const value = optionalUnsignedField(fields, key, locator);
  if (value === undefined) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Missing ${locator} source field ${key}`);
  }
  return value;
}

function optionalTextRange(
  fields: ReadonlyMap<string, string | number>,
  locator: string,
): readonly [start: number, end: number] | undefined {
  const start = optionalUnsignedField(fields, "rangeStart", locator);
  const end = optionalUnsignedField(fields, "rangeEnd", locator);
  if ((start === undefined) !== (end === undefined)) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Incomplete ${locator} source text range`);
  }
  if (start === undefined || end === undefined) return undefined;
  if (end < start) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid ${locator} source text range`);
  }
  return [start, end];
}

function readAction(
  fields: ReadonlyMap<string, string | number>,
  trigger: DocumentActionTrigger,
  locator: string,
): DocumentAction | undefined {
  const prefix = trigger === "click" ? "click" : "hover";
  const kind = optionalStringField(fields, `${prefix}Kind`, locator) as DocumentActionKind | undefined;
  if (kind === undefined) return undefined;
  if (!["hyperlink", "slide", "command", "unknown"].includes(kind)) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid ${locator} source field ${prefix}Kind`);
  }
  const action = optionalStringField(fields, `${prefix}Action`, locator);
  const target = optionalStringField(fields, `${prefix}Target`, locator);
  const tooltip = optionalStringField(fields, `${prefix}Tooltip`, locator);
  return {
    trigger,
    kind,
    ...(action === undefined ? {} : { action }),
    ...(target === undefined ? {} : { target }),
    ...(tooltip === undefined ? {} : { tooltip }),
  };
}

function readSource(
  reader: Reader,
  format: DocumentFormat,
  mapping: MappingQuality,
): SourceRef {
  const part = reader.string();
  const kind = reader.string();
  const fields = readFields(reader);
  const locator = `${format} ${kind}`;

  switch (format) {
    case "pptx": {
      if (kind !== "shape") break;
      const row = optionalUnsignedField(fields, "row", locator);
      const column = optionalUnsignedField(fields, "column", locator);
      const textRange = optionalTextRange(fields, locator);
      const name = optionalStringField(fields, "name", locator);
      const title = optionalStringField(fields, "title", locator);
      const description = optionalStringField(fields, "description", locator);
      const hidden = optionalUnsignedField(fields, "hidden", locator);
      if (hidden !== undefined && hidden > 1) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid ${locator} source field hidden`);
      }
      const actions = [
        readAction(fields, "click", locator),
        readAction(fields, "hover", locator),
      ].filter((action): action is DocumentAction => action !== undefined);
      return {
        format,
        part,
        kind,
        shapeId: requiredUnsignedField(fields, "shapeId", locator),
        ...(row === undefined ? {} : { row }),
        ...(column === undefined ? {} : { column }),
        ...(textRange === undefined ? {} : { textRange }),
        ...(name === undefined ? {} : { name }),
        ...(title === undefined ? {} : { title }),
        ...(description === undefined ? {} : { description }),
        ...(hidden === undefined ? {} : { hidden: hidden === 1 }),
        ...(actions.length === 0 ? {} : { actions }),
        mapping,
      };
    }
    case "odp": {
      if (kind !== "element") break;
      const elementId = optionalStringField(fields, "elementId", locator);
      const row = optionalUnsignedField(fields, "row", locator);
      const column = optionalUnsignedField(fields, "column", locator);
      return {
        format,
        part,
        kind,
        ...(elementId === undefined ? {} : { elementId }),
        path: requiredStringField(fields, "path", locator),
        ...(row === undefined ? {} : { row }),
        ...(column === undefined ? {} : { column }),
        mapping,
      };
    }
    case "xlsx": {
      if (kind !== "cell" && kind !== "drawing") break;
      const address = optionalStringField(fields, "address", locator);
      const formula = optionalStringField(fields, "formula", locator);
      const drawingId = optionalUnsignedField(fields, "drawingId", locator);
      return {
        format,
        part,
        kind,
        sheetName: requiredStringField(fields, "sheetName", locator),
        ...(address === undefined ? {} : { address }),
        ...(formula === undefined ? {} : { formula }),
        ...(drawingId === undefined ? {} : { drawingId }),
        mapping,
      };
    }
    case "ods": {
      if (kind !== "cell" && kind !== "element") break;
      const row = optionalUnsignedField(fields, "row", locator);
      const column = optionalUnsignedField(fields, "column", locator);
      const elementId = optionalStringField(fields, "elementId", locator);
      return {
        format,
        part,
        kind,
        tableName: requiredStringField(fields, "tableName", locator),
        ...(row === undefined ? {} : { row }),
        ...(column === undefined ? {} : { column }),
        ...(elementId === undefined ? {} : { elementId }),
        path: requiredStringField(fields, "path", locator),
        mapping,
      };
    }
    case "docx": {
      if (kind !== "paragraph" && kind !== "drawing" && kind !== "table" && kind !== "table-cell") break;
      const paragraphId = optionalStringField(fields, "paragraphId", locator);
      const paragraphIndex = optionalUnsignedField(fields, "paragraphIndex", locator);
      const drawingId = optionalUnsignedField(fields, "drawingId", locator);
      const row = optionalUnsignedField(fields, "row", locator);
      const column = optionalUnsignedField(fields, "column", locator);
      const textRange = optionalTextRange(fields, locator);
      const actions = [readAction(fields, "click", locator)]
        .filter((action): action is DocumentAction => action !== undefined);
      return {
        format,
        part,
        kind,
        ...(paragraphId === undefined ? {} : { paragraphId }),
        ...(paragraphIndex === undefined ? {} : { paragraphIndex }),
        ...(drawingId === undefined ? {} : { drawingId }),
        ...(row === undefined ? {} : { row }),
        ...(column === undefined ? {} : { column }),
        ...(textRange === undefined ? {} : { textRange }),
        ...(actions.length === 0 ? {} : { actions }),
        mapping,
      };
    }
    case "odt": {
      if (kind !== "element" && kind !== "text-range" && kind !== "table-cell") break;
      const elementId = optionalStringField(fields, "elementId", locator);
      const row = optionalUnsignedField(fields, "row", locator);
      const column = optionalUnsignedField(fields, "column", locator);
      const textRange = optionalTextRange(fields, locator);
      return {
        format,
        part,
        kind,
        ...(elementId === undefined ? {} : { elementId }),
        path: requiredStringField(fields, "path", locator),
        ...(row === undefined ? {} : { row }),
        ...(column === undefined ? {} : { column }),
        ...(textRange === undefined ? {} : { textRange }),
        mapping,
      };
    }
    case "rtf": {
      if (kind === "paragraph") {
        const textRange = optionalTextRange(fields, locator);
        return {
          format,
          part,
          kind,
          paragraphIndex: requiredUnsignedField(fields, "index", locator),
          ...(textRange === undefined ? {} : { textRange }),
          mapping,
        };
      }
      if (kind === "picture") {
        return {
          format,
          part,
          kind,
          pictureIndex: requiredUnsignedField(fields, "index", locator),
          mapping,
        };
      }
      break;
    }
    case "csv": {
      if (kind !== "cell") break;
      const textRange = optionalTextRange(fields, locator);
      return {
        format,
        part,
        kind,
        row: requiredUnsignedField(fields, "row", locator),
        column: requiredUnsignedField(fields, "column", locator),
        ...(textRange === undefined ? {} : { textRange }),
        mapping,
      };
    }
    case "ppt": {
      if (kind !== "text" && kind !== "drawing") break;
      const stream = requiredStringField(fields, "stream", locator);
      if (stream !== "PowerPoint Document") break;
      const recordOffset = optionalUnsignedField(fields, "recordOffset", locator);
      const textRange = kind === "text" ? optionalTextRange(fields, locator) : undefined;
      return {
        format,
        part,
        kind,
        stream,
        ...(recordOffset === undefined ? {} : { recordOffset }),
        ...(textRange === undefined ? {} : { textRange }),
        mapping,
      };
    }
    case "xls": {
      if (kind !== "cell" && kind !== "drawing") break;
      const stream = requiredStringField(fields, "stream", locator);
      if (stream !== "Workbook" && stream !== "Book") break;
      const recordOffset = optionalUnsignedField(fields, "recordOffset", locator);
      const row = optionalUnsignedField(fields, "row", locator);
      const column = optionalUnsignedField(fields, "column", locator);
      const textRange = optionalTextRange(fields, locator);
      return {
        format,
        part,
        kind,
        stream,
        ...(recordOffset === undefined ? {} : { recordOffset }),
        ...(row === undefined ? {} : { row }),
        ...(column === undefined ? {} : { column }),
        ...(textRange === undefined ? {} : { textRange }),
        mapping,
      };
    }
    case "doc": {
      if (kind !== "paragraph") break;
      const stream = requiredStringField(fields, "stream", locator);
      if (stream !== "WordDocument") break;
      const recordOffset = optionalUnsignedField(fields, "recordOffset", locator);
      const textRange = optionalTextRange(fields, locator);
      return {
        format,
        part,
        kind,
        stream,
        ...(recordOffset === undefined ? {} : { recordOffset }),
        ...(textRange === undefined ? {} : { textRange }),
        mapping,
      };
    }
    case "pages": {
      if (kind !== "preview" && kind !== "body" && kind !== "inline-image"
        && kind !== "shape" && kind !== "image" && kind !== "media" && kind !== "text-box"
        && kind !== "table" && kind !== "table-grid" && kind !== "table-cell"
        && kind !== "chart" && kind !== "chart-series") break;
      return {
        format,
        part,
        kind,
        component: requiredStringField(fields, "component", locator),
        mapping,
      };
    }
    case "numbers": {
      if (kind !== "preview" && kind !== "cell" && kind !== "chart" && kind !== "chart-series") break;
      return {
        format,
        part,
        kind,
        component: requiredStringField(fields, "component", locator),
        mapping,
      };
    }
    case "keynote": {
      if (kind !== "preview" && kind !== "slide-preview" && kind !== "slide-background"
        && kind !== "shape" && kind !== "image" && kind !== "media" && kind !== "text-box"
        && kind !== "table" && kind !== "table-grid" && kind !== "table-cell"
        && kind !== "chart" && kind !== "chart-series") break;
      return {
        format,
        part,
        kind,
        component: requiredStringField(fields, "component", locator),
        mapping,
      };
    }
    case "pdf": {
      if (kind !== "text" && kind !== "path" && kind !== "image" && kind !== "form"
        && kind !== "annotation" && kind !== "knockout-backdrop") break;
      const objectNumber = optionalUnsignedField(fields, "objectNumber", locator);
      const byteOffset = optionalUnsignedField(fields, "byteOffset", locator);
      const actions = [
        readAction(fields, "click", locator),
        readAction(fields, "hover", locator),
      ].filter((action): action is DocumentAction => action !== undefined);
      return {
        format,
        part,
        kind,
        ...(objectNumber === undefined ? {} : { objectNumber }),
        ...(byteOffset === undefined ? {} : { byteOffset }),
        ...(actions.length === 0 ? {} : { actions }),
        mapping,
      };
    }
    case "xps": {
      if (kind !== "canvas" && kind !== "glyphs" && kind !== "path") break;
      return {
        format,
        part,
        kind,
        path: requiredStringField(fields, "path", locator),
        mapping,
      };
    }
    case "ofd": {
      if (kind !== "text" && kind !== "path" && kind !== "image" && kind !== "seal") break;
      return {
        format,
        part,
        kind,
        path: requiredStringField(fields, "path", locator),
        mapping,
      };
    }
  }

  throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid ${format} source locator ${kind}`);
}

function readPathCommands(reader: Reader): readonly ScenePathCommand[] {
  const count = reader.u32();
    if (count > MAX_PATH_COMMANDS || count > Math.floor(reader.remaining / 4)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid path command count");
    }
    const commands: ScenePathCommand[] = [];
    for (let index = 0; index < count; index += 1) {
      const command = reader.u8();
      reader.skip(3);
      if (command === 0) commands.push({ kind: "moveTo", x: reader.f32(), y: reader.f32() });
      else if (command === 1) commands.push({ kind: "lineTo", x: reader.f32(), y: reader.f32() });
      else if (command === 2) {
        commands.push({
          kind: "quadraticCurveTo",
          cpx: reader.f32(),
          cpy: reader.f32(),
          x: reader.f32(),
          y: reader.f32(),
        });
      } else if (command === 3) {
        commands.push({
          kind: "bezierCurveTo",
          cp1x: reader.f32(),
          cp1y: reader.f32(),
          cp2x: reader.f32(),
          cp2y: reader.f32(),
          x: reader.f32(),
          y: reader.f32(),
        });
      } else if (command === 4) commands.push({ kind: "closePath" });
      else throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid path command ${command}`);
    }
  return commands;
}

function readGeometry(reader: Reader, version: number): SceneGeometry {
  const code = reader.u8();
  if (code <= 2) return enumValue(LEGACY_GEOMETRIES, code, "geometry");
  if (version < 4) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid geometry code ${code}`);
  }
  if (code === 3) {
    return { kind: "rounded-rectangle", radiusX: reader.f32(), radiusY: reader.f32() };
  }
  if (code === 4) {
    const fillRule = enumValue(["nonzero", "evenodd"] as const, reader.u8(), "path fill rule");
    reader.skip(3);
    return { kind: "path", fillRule, commands: readPathCommands(reader) };
  }
  if (code === 5 && version >= 21) {
    const count = reader.u32();
    if (count > MAX_PATH_COMMANDS || count > Math.floor(reader.remaining / 8)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid layered path count");
    }
    const layers = [];
    for (let index = 0; index < count; index += 1) {
      const fillRule = enumValue(["nonzero", "evenodd"] as const, reader.u8(), "path fill rule");
      const fill = enumValue(
        ["normal", "none", "darken", "darken-less", "lighten", "lighten-less"] as const,
        reader.u8(),
        "path fill mode",
      );
      const stroke = reader.u8();
      if (stroke > 1) throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid path stroke flag");
      reader.skip(1);
      layers.push({ fillRule, fill, stroke: stroke === 1, commands: readPathCommands(reader) });
    }
    return { kind: "layered-path", layers };
  }
  throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid geometry code ${code}`);
}

function readXpsColor(reader: Reader): SceneXpsColor {
  const kind = reader.u8();
  reader.skip(3);
  if (kind === 0) return { kind: "rgba", color: reader.u32() };
  if (kind !== 1) throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid XPS color kind ${kind}`);
  const alpha = reader.f32();
  const profile = reader.byteArray();
  const count = reader.u32();
  if (count === 0 || count > 15 || count > Math.floor(reader.remaining / 4)) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid XPS ContextColor channel count");
  }
  const channels = Array.from({ length: count }, () => reader.f32());
  return { kind: "context", alpha, profile, channels };
}

function readPaint(reader: Reader, version: number, depth = 0): ScenePaint {
  if (depth > MAX_VISUAL_DEPTH) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Visual brush nesting exceeds the protocol limit");
  }
  const kind = reader.u8();
  reader.skip(3);
  if (kind === 0) return { kind: "none" };
  if (kind === 1) return { kind: "solid", color: reader.u32() };
  if (kind === 2) {
    const start = { x: reader.f32(), y: reader.f32() };
    const end = { x: reader.f32(), y: reader.f32() };
    const count = reader.u32();
    if (count > MAX_GRADIENT_STOPS || count > Math.floor(reader.remaining / 8)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid gradient stop count");
    }
    const stops: { offset: number; color: number }[] = [];
    for (let index = 0; index < count; index += 1) {
      stops.push({ offset: reader.f32(), color: reader.u32() });
    }
    return { kind: "linear-gradient", start, end, stops };
  }
  if (kind === 3 && version >= 5) {
    const start = { x: reader.f32(), y: reader.f32(), radius: reader.f32() };
    const end = { x: reader.f32(), y: reader.f32(), radius: reader.f32() };
    const count = reader.u32();
    if (count > MAX_GRADIENT_STOPS || count > Math.floor(reader.remaining / 8)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid gradient stop count");
    }
    const stops: { offset: number; color: number }[] = [];
    for (let index = 0; index < count; index += 1) {
      stops.push({ offset: reader.f32(), color: reader.u32() });
    }
    return { kind: "radial-gradient", start, end, stops };
  }
  if (kind === 4 && version >= 12) {
    return {
      kind: "pattern",
      preset: reader.string(),
      foreground: reader.u32(),
      background: reader.u32(),
    };
  }
  if (kind === 5 && version >= 13) {
    const cropLeft = reader.f32();
    const cropTop = reader.f32();
    const cropRight = reader.f32();
    const cropBottom = reader.f32();
    const tileWidth = version >= 36 ? reader.f32() : 0;
    const tileHeight = version >= 36 ? reader.f32() : 0;
    const flags = reader.u8();
    const tile = (flags & 1) !== 0;
    reader.skip(3);
    let mapping: Extract<ScenePaint, {kind: "image"}>["mapping"];
    if (version >= 70 && (flags & 2) !== 0) {
      const scaleX = reader.f32(), scaleY = reader.f32(), offsetX = reader.f32(), offsetY = reader.f32();
      const alignmentX = reader.f32(), alignmentY = reader.f32();
      const left = reader.f32(), top = reader.f32(), right = reader.f32(), bottom = reader.f32(), dpi = reader.f32();
      const flip = ["none", "tile", "flip-x", "flip-y", "flip-xy"][reader.u8()] as NonNullable<typeof mapping>["flip"];
      const rotateWithShape = reader.u8() !== 0;
      reader.skip(2);
      if (flip === undefined || ![scaleX, scaleY, offsetX, offsetY, alignmentX, alignmentY, left, top, right, bottom, dpi].every(Number.isFinite)
        || scaleX <= 0 || scaleY <= 0 || dpi < 0) throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid image fill mapping");
      mapping = { scaleX, scaleY, offsetX, offsetY, alignmentX, alignmentY, left, top, right, bottom, dpi, flip, rotateWithShape };
    }
    return {
      kind: "image",
      cropLeft,
      cropTop,
      cropRight,
      cropBottom,
      tile,
      ...(mapping === undefined ? {} : { mapping }),
      ...(tileWidth > 0 ? { tileWidth } : {}),
      ...(tileHeight > 0 ? { tileHeight } : {}),
      mediaType: reader.string(),
      bytes: reader.byteArray(),
    };
  }
  if (kind === 11 && version >= 70) {
    const tile = { left: reader.f32(), top: reader.f32(), right: reader.f32(), bottom: reader.f32() };
    const flipCode = reader.u8(); const rotate = reader.u8(); reader.u16();
    const flip = (["none", "tile", "flip-x", "flip-y", "flip-xy"] as const)[flipCode];
    if (flip === undefined || rotate > 1 || !Object.values(tile).every(Number.isFinite)
      || tile.left + tile.right >= 1 || tile.top + tile.bottom >= 1) throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid gradient mapping");
    return { kind: "mapped-gradient", tile, flip, rotateWithShape: rotate === 1, paint: readPaint(reader, version, depth + 1) };
  }
  if (kind === 10 && version >= 70) {
    const focus = { left: reader.f32(), top: reader.f32(), right: reader.f32(), bottom: reader.f32() };
    const count = reader.u32();
    if (count > MAX_GRADIENT_STOPS || count > Math.floor(reader.remaining / 8)
      || !Object.values(focus).every(Number.isFinite)) throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid shape gradient");
    const stops = Array.from({ length: count }, () => ({ offset: reader.f32(), color: reader.u32() }));
    return { kind: "shape-gradient", focus, stops };
  }
  if (kind === 6 && version >= 25) {
    const center = { x: reader.f32(), y: reader.f32() };
    const count = reader.u32();
    if (count > MAX_GRADIENT_STOPS || count > Math.floor(reader.remaining / 8)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid gradient stop count");
    }
    const stops: { offset: number; color: number }[] = [];
    for (let index = 0; index < count; index += 1) {
      stops.push({ offset: reader.f32(), color: reader.u32() });
    }
    return { kind: "rect-gradient", center, stops };
  }
  if (kind === 7 && version >= 26) {
    const start = { x: reader.f32(), y: reader.f32(), radius: reader.f32() };
    const end = { x: reader.f32(), y: reader.f32(), radius: reader.f32() };
    const count = reader.u32();
    if (count > MAX_GRADIENT_STOPS || count > Math.floor(reader.remaining / 8)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid gradient stop count");
    }
    const stops: { offset: number; color: number }[] = [];
    for (let index = 0; index < count; index += 1) {
      stops.push({ offset: reader.f32(), color: reader.u32() });
    }
    return { kind: "circle-gradient", start, end, stops };
  }
  if (kind === 8 && version >= 42) {
    const readRect = () => ({
      x: reader.f32(),
      y: reader.f32(),
      width: reader.f32(),
      height: reader.f32(),
    });
    const readTransform = () => ({
      a: reader.f32(),
      b: reader.f32(),
      c: reader.f32(),
      d: reader.f32(),
      e: reader.f32(),
      f: reader.f32(),
    });
    const viewbox = readRect();
    const viewport = readRect();
    const transform = readTransform();
    const relativeTransform = readTransform();
    const opacity = reader.f32();
    const alignmentX = reader.f32();
    const alignmentY = reader.f32();
    const tileMode = enumValue(
      ["none", "tile", "flip-x", "flip-y", "flip-xy"] as const,
      reader.u8(),
      "visual brush tile mode",
    );
    const stretch = enumValue(
      ["none", "fill", "uniform", "uniform-to-fill"] as const,
      reader.u8(),
      "visual brush stretch",
    );
    const flags = reader.u8();
    if ((flags & ~0b11) !== 0) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid visual brush flags");
    }
    reader.skip(1);
    const count = reader.u32();
    if (count > MAX_VISUAL_BRUSH_CHILDREN || count > Math.floor(reader.remaining / 20)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid visual brush child count");
    }
    const children = [];
    for (let index = 0; index < count; index += 1) {
      const bounds = readRect();
      const visualKind = reader.u8();
      reader.skip(3);
      children.push({ bounds, visual: readVisual(reader, visualKind, version, depth + 1) });
    }
    return {
      kind: "visual",
      viewbox,
      viewport,
      viewboxRelative: (flags & 1) !== 0,
      viewportRelative: (flags & 2) !== 0,
      tileMode,
      stretch,
      alignmentX,
      alignmentY,
      transform,
      relativeTransform,
      opacity,
      children,
    };
  }
  if (kind === 9 && version >= 44) {
    const start = { x: reader.f32(), y: reader.f32() };
    const end = { x: reader.f32(), y: reader.f32() };
    const radiusX = reader.f32();
    const radiusY = reader.f32();
    const readTransform = () => ({
      a: reader.f32(), b: reader.f32(), c: reader.f32(), d: reader.f32(), e: reader.f32(), f: reader.f32(),
    });
    const transform = readTransform();
    const relativeTransform = readTransform();
    const flags = reader.u8();
    if ((flags & ~0b111) !== 0) throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid XPS gradient flags");
    const spread = enumValue(["pad", "reflect", "repeat"] as const, reader.u8(), "XPS gradient spread");
    reader.skip(2);
    const count = reader.u32();
    const stopBytes = version >= 47 ? 12 : 8;
    if (count > MAX_GRADIENT_STOPS || count > Math.floor(reader.remaining / stopBytes)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid XPS gradient stop count");
    }
    const stops = [];
    for (let index = 0; index < count; index += 1) {
      const offset = reader.f32();
      const color = version >= 47 ? readXpsColor(reader) : { kind: "rgba" as const, color: reader.u32() };
      stops.push({ offset, color });
    }
    return {
      kind: "xps-gradient", radial: (flags & 1) !== 0, relative: (flags & 2) !== 0,
      linearRgb: (flags & 4) !== 0, start, end, radiusX, radiusY,
      transform, relativeTransform, spread, stops,
    };
  }
  throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid paint kind ${kind}`);
}

function readVisual(reader: Reader, kind: number, version: number, depth = 0): SceneVisual {
  if (depth > MAX_VISUAL_DEPTH) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Visual layer nesting exceeds the protocol limit");
  }
  if (kind === 0) return { kind: "none" };
  if (kind === 1) {
    return {
      kind: "shape",
      geometry: readGeometry(reader, version),
      fill: reader.u32(),
      stroke: reader.u32(),
      strokeWidth: reader.f32(),
    };
  }
  if (kind === 2) {
    const geometry = readGeometry(reader, version);
    const fill = reader.u32();
    const stroke = reader.u32();
    const strokeWidth = reader.f32();
    const fontFamily = reader.string();
    const fontSize = reader.f32();
    const color = reader.u32();
    const flags = reader.u8();
    const align = textAlignValue(reader, version, "text alignment");
    reader.skip(2);
    return {
      kind: "text",
      geometry,
      fill,
      stroke,
      strokeWidth,
      fontFamily,
      fontSize,
      color,
      bold: (flags & 1) !== 0,
      italic: (flags & 2) !== 0,
      align,
    };
  }
  if (kind === 3) {
    return {
      kind: "image",
      cropLeft: reader.f32(),
      cropTop: reader.f32(),
      cropRight: reader.f32(),
      cropBottom: reader.f32(),
      mediaType: reader.string(),
      bytes: reader.byteArray(),
    };
  }
  if (kind === 4) {
    const cropLeft = reader.f32();
    const cropTop = reader.f32();
    const cropRight = reader.f32();
    const cropBottom = reader.f32();
    const mediaType = reader.string();
    const bytes = reader.byteArray();
    const fallbackMediaType = reader.string();
    const fallbackBytes = reader.byteArray();
    return {
      kind: "image",
      cropLeft,
      cropTop,
      cropRight,
      cropBottom,
      mediaType,
      bytes,
      fallback: { mediaType: fallbackMediaType, bytes: fallbackBytes },
    };
  }
  if (version < 4) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid visual kind ${kind}`);
  }
  if (kind === 5) {
    const transform = {
      a: reader.f32(),
      b: reader.f32(),
      c: reader.f32(),
      d: reader.f32(),
      e: reader.f32(),
      f: reader.f32(),
    };
    const opacity = reader.f32();
    const nestedKind = reader.u8();
    const blendMode = version >= 39
      ? enumValue([
          "source-over", "multiply", "screen", "overlay", "darken", "lighten",
          "color-dodge", "color-burn", "hard-light", "soft-light", "difference",
          "exclusion", "hue", "saturation", "color", "luminosity",
          "destination-out",
        ] as const, reader.u8(), "blend mode")
      : "source-over";
    reader.skip(version >= 39 ? 2 : 3);
    return {
      kind: "layer",
      transform,
      opacity,
      ...(blendMode === "source-over" ? {} : { blendMode }),
      visual: readVisual(reader, nestedKind, version, depth + 1),
    };
  }
  if (kind === 6) {
    return {
      kind: "painted-shape",
      geometry: readGeometry(reader, version),
      fill: readPaint(reader, version, depth),
      stroke: readPaint(reader, version, depth),
      strokeWidth: reader.f32(),
    };
  }
  if (kind === 7) {
    const geometry = readGeometry(reader, version);
    const fill = readPaint(reader, version, depth);
    const stroke = readPaint(reader, version, depth);
    const strokeWidth = reader.f32();
    const align = textAlignValue(reader, version, "text alignment");
    reader.skip(3);
    const lineHeight = reader.f32();
    const count = reader.u32();
    if (count > MAX_RICH_TEXT_RUNS || count > Math.floor(reader.remaining / 32)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid rich-text run count");
    }
    const runs = [];
    for (let index = 0; index < count; index += 1) {
      runs.push(readTextRun(reader, version, depth));
    }
    return { kind: "rich-text", geometry, fill, stroke, strokeWidth, align, lineHeight, runs };
  }
  if (version < 5) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid visual kind ${kind}`);
  }
  if (kind === 8) {
    const flags = reader.u8();
    reader.skip(3);
    if ((flags & ~3) !== 0) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid visual effect flags");
    }
    const shadow = (flags & 1) === 0 ? undefined : {
      color: reader.u32(),
      blur: reader.f32(),
      offsetX: reader.f32(),
      offsetY: reader.f32(),
    };
    const clip = (flags & 2) === 0 ? undefined : readGeometry(reader, version);
    const nestedKind = reader.u8();
    reader.skip(3);
    return {
      kind: "effect",
      ...(shadow === undefined ? {} : { shadow }),
      ...(clip === undefined ? {} : { clip }),
      visual: readVisual(reader, nestedKind, version, depth + 1),
    };
  }
  if (kind === 9) {
    const direction = enumValue(["auto", "ltr", "rtl"] as const, reader.u8(), "text direction");
    const orientation = enumValue(
      ["horizontal", "vertical-rl", "vertical-lr", "rotated-90", "rotated-270", "stacked-rl", "stacked-lr"] as const,
      reader.u8(),
      "text orientation",
    );
    const autoFit = enumValue(["none", "shrink", "fit-frame"] as const, reader.u8(), "text autofit");
    const flags = reader.u8();
    if ((flags & ~(version >= 64 ? 0b111111 : version >= 62 ? 0b11111 : version >= 61 ? 0b1111 : 0b111)) !== 0 || ((flags >> 1) & 0b11) === 0b11) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid text-layout flags");
    }
    const verticalAlign = enumValue(
      ["top", "center", "bottom"] as const,
      (flags >> 1) & 0b11,
      "text vertical alignment",
    );
    const defaultTabStop = reader.f32();
    const hangingIndent = reader.f32();
    const minScale = reader.f32();
    const paragraphSpacing = version >= 7 ? reader.f32() : 0;
    const insetLeft = version >= 8 ? reader.f32() : 4;
    const insetRight = version >= 8 ? reader.f32() : 4;
    const insetTop = version >= 8 ? reader.f32() : 4;
    const insetBottom = version >= 8 ? reader.f32() : 4;
    const marginLeft = version >= 8 ? reader.f32() : 0;
    const marginRight = version >= 8 ? reader.f32() : 0;
    const firstLineIndent = version >= 8 ? reader.f32() : 0;
    const columnCount = version >= 16 ? reader.u32() : 1;
    const columnSpacing = version >= 16 ? reader.f32() : 0;
    const rotationDegrees = version >= 16 ? reader.f32() : 0;
    const fontScale = version >= 16 ? reader.f32() : 1;
    const lineSpacingReduction = version >= 16 ? reader.f32() : 0;
    const horizontalOverflow = version >= 16
      ? enumValue(["overflow", "clip"] as const, reader.u8(), "horizontal text overflow")
      : "overflow";
    const verticalOverflow = version >= 16
      ? enumValue(["overflow", "clip", "ellipsis"] as const, reader.u8(), "vertical text overflow")
      : "overflow";
    const wrapValue = version >= 16 ? reader.u8() : 1;
    const warpValue = version >= 16 ? reader.u8() : 0;
    if (wrapValue > 1 || warpValue > 1) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid text layout extension flags");
    }
    const textPaintFlags = version >= 32 ? reader.u8() : 1;
    if (version >= 32) reader.skip(3);
    const textStrokeColor = version >= 32 ? reader.u32() : 0;
    const textStrokeWidth = version >= 32 ? reader.f32() : 0;
    const textBaseline = version >= 32 ? reader.f32() : 0;
    const textPaint = version >= 40 ? readPaint(reader, version, depth) : undefined;
    const textStrokePaint = version >= 41 ? readPaint(reader, version, depth) : undefined;
    const validTextPaintFlags = version >= 41 ? 0b1111 : version >= 34 ? 0b111 : 0b11;
    if ((textPaintFlags & ~validTextPaintFlags) !== 0 || !Number.isFinite(textStrokeWidth) || textStrokeWidth < 0
      || !Number.isFinite(textBaseline) || textBaseline < 0) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid text paint fields");
    }
    const tabCount = reader.u32();
    const tabStopBytes = version >= 28 ? 8 : 4;
    if (tabCount > MAX_TAB_STOPS || tabCount > Math.floor(reader.remaining / tabStopBytes)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid text tab-stop count");
    }
    const tabStops = [];
    for (let index = 0; index < tabCount; index += 1) {
      const position = reader.f32();
      if (version >= 28) {
        const align = enumValue(
          ["start", "center", "end"] as const,
          reader.u8(),
          "text tab alignment",
        );
        const leader = enumValue(
          ["none", "dot", "hyphen", "underscore", "middle-dot"] as const,
          reader.u8(),
          "text tab leader",
        );
        reader.skip(2);
        tabStops.push({ position, align, leader });
      } else {
        tabStops.push({ position, align: "start" as const, leader: "none" as const });
      }
    }
    const prefix = (flags & 1) === 0 ? undefined : reader.string();
    const warp = warpValue === 0 ? undefined : reader.string();
    const paragraphs = [];
    if (version >= 9) {
      const paragraphCount = reader.u32();
      const paragraphBytes = version >= 68 ? 104 : version >= 66 ? 84 : version >= 65 ? 60 : version >= 24 ? 36 : version >= 19 ? 32 : 20;
      if (paragraphCount > MAX_TEXT_PARAGRAPHS || paragraphCount > Math.floor(reader.remaining / paragraphBytes)) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid text paragraph count");
      }
      for (let index = 0; index < paragraphCount; index += 1) {
        const align = textAlignValue(reader, version, "paragraph alignment");
        reader.skip(3);
        const paragraph = {
          align,
          marginLeft: reader.f32(),
          marginRight: reader.f32(),
          firstLineIndent: reader.f32(),
          defaultTabStop: reader.f32(),
        };
        const paragraphSpacing = version >= 19 ? {
          ...paragraph,
          lineHeight: reader.f32(),
          spaceBefore: reader.f32(),
          spaceAfter: reader.f32(),
        } : paragraph;
        if (version >= 24) {
          const paragraphFlags = reader.u8();
          reader.skip(3);
          if ((paragraphFlags & ~(version >= 68 ? 0b11111 : version >= 66 ? 0b1111 : version >= 65 ? 0b111 : 0b11)) !== 0) {
            throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid text paragraph flags");
          }
          const ruleAbove = version >= 65 ? {
            color: reader.u32(),
            strokeWidth: reader.f32(),
            offsetX: reader.f32(),
            offsetY: reader.f32(),
            width: reader.f32(),
          } : undefined;
          if (version >= 65) reader.skip(4);
          const ruleBelow = version >= 66 ? {
            color: reader.u32(), strokeWidth: reader.f32(),
            offsetX: reader.f32(), offsetY: reader.f32(), width: reader.f32(),
          } : undefined;
          if (version >= 66) reader.skip(4);
          const dropCap = version >= 68 ? {
            characters: reader.u32(), lines: reader.u32(), raisedLines: reader.u32(),
            padding: reader.f32(), outdent: reader.f32(),
          } : undefined;
          paragraphs.push({
            ...paragraphSpacing,
            latinLineBreak: (paragraphFlags & 1) !== 0,
            hangingPunctuation: (paragraphFlags & 2) !== 0,
            ...((paragraphFlags & 4) === 0 || ruleAbove === undefined ? {} : { ruleAbove }),
            ...((paragraphFlags & 8) === 0 || ruleBelow === undefined ? {} : { ruleBelow }),
            ...((paragraphFlags & 16) === 0 || dropCap === undefined ? {} : { dropCap }),
          });
        } else {
          paragraphs.push({
            ...paragraphSpacing,
            // Paragraph flags were added in v24; Office defaults apply to old payloads.
            latinLineBreak: false,
            hangingPunctuation: true,
          });
        }
      }
    }
    const wrapRegions = [];
    if (version >= 67) {
      const count = reader.u32();
      if (count > 1024 || count > Math.floor(reader.remaining / 16)) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid text wrap region count");
      }
      for (let index = 0; index < count; index += 1) {
        wrapRegions.push({x: reader.f32(), y: reader.f32(), width: reader.f32(), height: reader.f32()});
      }
    }
    const fillOffset = version >= 72 ? reader.u32() : 0xffff_ffff;
    const fillCodePoint = version >= 72 ? reader.u32() : 0;
    if (fillOffset !== 0xffff_ffff && (fillCodePoint < 32 || fillCodePoint > 0x10ffff
      || fillCodePoint >= 0x7f && fillCodePoint <= 0x9f
      || fillCodePoint >= 0xd800 && fillCodePoint <= 0xdfff)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid text fill character");
    }
    const nestedKind = reader.u8();
    reader.skip(3);
    return {
      kind: "text-layout",
      layout: {
        ...(fillOffset === 0xffff_ffff ? {} : { fillCharacter: {
          offset: fillOffset, character: String.fromCodePoint(fillCodePoint),
        } }),
        ...((flags & 8) === 0 ? {} : { continuesAfter: true }),
        ...((flags & 16) === 0 ? {} : { compressPunctuation: true }),
        ...((flags & 32) === 0 ? {} : { fixedLineHeight: true }),
        direction,
        orientation,
        autoFit,
        verticalAlign,
        ...(prefix === undefined ? {} : { prefix }),
        tabStops,
        defaultTabStop,
        hangingIndent,
        paragraphSpacing,
        insetLeft,
        insetRight,
        insetTop,
        insetBottom,
        marginLeft,
        marginRight,
        firstLineIndent,
        columnCount,
        columnSpacing,
        rotationDegrees,
        fontScale,
        lineSpacingReduction,
        horizontalOverflow,
        verticalOverflow,
        wrap: wrapValue === 1,
        ...(warp === undefined ? {} : { warp }),
        ...(version < 32 ? {} : {
          textFill: (textPaintFlags & 1) !== 0,
          textScaleToFit: (textPaintFlags & 2) !== 0,
          ...(version < 41 ? {} : { textMatrixScaleToFit: (textPaintFlags & 8) !== 0 }),
          ...(version < 34 ? {} : { lowResolutionSupersample: (textPaintFlags & 4) !== 0 }),
          textStrokeColor,
          textStrokeWidth,
          textBaseline,
          ...(textPaint === undefined || textPaint.kind === "none" ? {} : { textPaint }),
          ...(textStrokePaint === undefined || textStrokePaint.kind === "none"
            ? {} : { textStrokePaint }),
        }),
        ...(paragraphs.length === 0 ? {} : { paragraphs }),
        ...(wrapRegions.length === 0 ? {} : { wrapRegions }),
        minScale,
      },
      visual: readVisual(reader, nestedKind, version, depth + 1),
    };
  }
  if (kind === 10 && version >= 14) {
    const cap = enumValue(["flat", "round", "square"] as const, reader.u8(), "line cap");
    const join = enumValue(["miter", "round", "bevel"] as const, reader.u8(), "line join");
    const compound = enumValue(
      ["single", "double", "thick-thin", "thin-thick", "triple"] as const,
      reader.u8(),
      "compound line",
    );
    const alignment = enumValue(["center", "inset"] as const, reader.u8(), "line alignment");
    const miterLimit = reader.f32();
    const dashOffset = version >= 71 ? reader.f32() : 0;
    const dashCount = reader.u32();
    if (dashCount > 256 || dashCount > Math.floor(reader.remaining / 4)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid line dash count");
    }
    const dash: number[] = [];
    for (let index = 0; index < dashCount; index += 1) dash.push(reader.f32());
    const nestedKind = reader.u8();
    reader.skip(3);
    return {
      kind: "stroke-style",
      style: { cap, join, compound, alignment, miterLimit, dash, dashOffset },
      visual: readVisual(reader, nestedKind, version, depth + 1),
    };
  }
  if (kind === 11 && version >= 15) {
    const flags = reader.u8();
    reader.skip(3);
    const supportedFlags = version >= 57 ? 0b11_1111 : version >= 20 ? 0b1_1111 : 0b1111;
    if (flags === 0 || (flags & ~supportedFlags) !== 0) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid advanced-effect flags");
    }
    let outerShadow;
    if ((flags & 32) !== 0) {
      const color = reader.u32();
      const blur = reader.f32();
      const offsetX = reader.f32();
      const offsetY = reader.f32();
      const scaleX = reader.f32();
      const scaleY = reader.f32();
      const skewX = reader.f32();
      const skewY = reader.f32();
      const alignment = reader.u8();
      reader.skip(3);
      outerShadow = { color, blur, offsetX, offsetY, scaleX, scaleY, skewX, skewY, alignment };
    }
    const innerShadow = (flags & 1) === 0 ? undefined : {
      color: reader.u32(),
      blur: reader.f32(),
      offsetX: reader.f32(),
      offsetY: reader.f32(),
    };
    const glow = (flags & 2) === 0 ? undefined : {
      color: reader.u32(),
      radius: reader.f32(),
    };
    let reflection;
    if ((flags & 4) !== 0) {
      if (version >= 49) {
        reflection = {
          startOpacity: reader.f32(),
          endOpacity: reader.f32(),
          startPosition: reader.f32(),
          endPosition: reader.f32(),
          directionDegrees: reader.f32(),
          blur: reader.f32(),
          distance: reader.f32(),
          scaleX: reader.f32(),
          scaleY: reader.f32(),
        };
      } else {
        const opacity = reader.f32();
        reflection = {
          startOpacity: opacity,
          endOpacity: opacity,
          startPosition: 0,
          endPosition: 1,
          directionDegrees: 90,
          blur: reader.f32(),
          distance: reader.f32(),
          scaleX: reader.f32(),
          scaleY: reader.f32(),
        };
      }
    }
    const softEdge = (flags & 8) === 0 ? undefined : reader.f32();
    let threeD;
    if ((flags & 16) !== 0 && version >= 20) {
      const cameraPreset = reader.string();
      const lightRig = reader.string();
      const lightDirection = reader.string();
      const material = reader.string();
      const cameraFov = reader.f32();
      const cameraZoom = reader.f32();
      const cameraLatitude = reader.f32();
      const cameraLongitude = reader.f32();
      const cameraRevolution = reader.f32();
      const lightLatitude = reader.f32();
      const lightLongitude = reader.f32();
      const lightRevolution = reader.f32();
      const z = reader.f32();
      const extrusionHeight = reader.f32();
      const contourWidth = reader.f32();
      const threeDFlags = reader.u8();
      reader.skip(3);
      const supportedThreeDFlags = version >= 58 ? 0b111_1111 : version >= 21 ? 0b11_1111 : 0b1111;
      if ((threeDFlags & ~supportedThreeDFlags) !== 0) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid 3D effect flags");
      }
      const readBevel = () => ({
        width: reader.f32(),
        height: reader.f32(),
        preset: reader.string(),
      });
      const bevelTop = (threeDFlags & 1) === 0 ? undefined : readBevel();
      const bevelBottom = (threeDFlags & 2) === 0 ? undefined : readBevel();
      const extrusionColor = (threeDFlags & 4) === 0 ? undefined : reader.u32();
      const contourColor = (threeDFlags & 8) === 0 ? undefined : reader.u32();
      const backdrop = (threeDFlags & 16) === 0 ? undefined : {
        anchorX: reader.f32(),
        anchorY: reader.f32(),
        anchorZ: reader.f32(),
        normalX: reader.f32(),
        normalY: reader.f32(),
        normalZ: reader.f32(),
        upX: reader.f32(),
        upY: reader.f32(),
        upZ: reader.f32(),
      };
      const flatTextZ = (threeDFlags & 32) === 0 ? undefined : reader.f32();
      const appliesToText = (threeDFlags & 64) !== 0;
      threeD = {
        cameraPreset,
        cameraFov,
        cameraZoom,
        cameraLatitude,
        cameraLongitude,
        cameraRevolution,
        lightRig,
        lightDirection,
        lightLatitude,
        lightLongitude,
        lightRevolution,
        z,
        extrusionHeight,
        contourWidth,
        material,
        ...(bevelTop === undefined ? {} : { bevelTop }),
        ...(bevelBottom === undefined ? {} : { bevelBottom }),
        ...(extrusionColor === undefined ? {} : { extrusionColor }),
        ...(contourColor === undefined ? {} : { contourColor }),
        ...(backdrop === undefined ? {} : { backdrop }),
        ...(flatTextZ === undefined ? {} : { flatTextZ }),
        ...(appliesToText ? { appliesToText } : {}),
      };
    }
    const nestedKind = reader.u8();
    reader.skip(3);
    return {
      kind: "advanced-effect",
      ...(outerShadow === undefined ? {} : { outerShadow }),
      ...(innerShadow === undefined ? {} : { innerShadow }),
      ...(glow === undefined ? {} : { glow }),
      ...(reflection === undefined ? {} : { reflection }),
      ...(softEdge === undefined ? {} : { softEdge }),
      ...(threeD === undefined ? {} : { threeD }),
      visual: readVisual(reader, nestedKind, version, depth + 1),
    };
  }
  if (kind === 12 && version >= 22) {
    const from = reader.u32();
    const to = reader.u32();
    const useAlpha = version >= 23
      ? enumValue([false, true] as const, reader.u8(), "image color-change useA")
      : true;
    const nestedKind = reader.u8();
    reader.skip(version >= 23 ? 2 : 3);
    return {
      kind: "image-color-change",
      from,
      to,
      useAlpha,
      visual: readVisual(reader, nestedKind, version, depth + 1),
    };
  }
  if (kind === 18 && version >= 48) {
    const flags = reader.u8();
    const nestedKind = reader.u8();
    reader.skip(2);
    const bilevelThreshold = reader.f32();
    const brightness = reader.f32();
    const contrast = reader.f32();
    const first = reader.u32();
    const second = reader.u32();
    if ((flags & ~0b111) !== 0) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid image adjustment flags");
    }
    return {
      kind: "image-adjustment",
      grayscale: (flags & 1) !== 0,
      ...((flags & 2) === 0 ? {} : { bilevelThreshold }),
      brightness,
      contrast,
      ...((flags & 4) === 0 ? {} : { duotone: [first, second] as const }),
      visual: readVisual(reader, nestedKind, version, depth + 1),
    };
  }
  if (kind === 13 && version >= 29) {
    const mediaKind = enumValue(["audio", "video"] as const, reader.u8(), "embedded media kind");
    reader.skip(3);
    const mediaType = reader.string();
    const bytes = reader.byteArray();
    const nestedKind = reader.u8();
    reader.skip(3);
    return {
      kind: "media",
      mediaKind,
      mediaType,
      bytes,
      visual: readVisual(reader, nestedKind, version, depth + 1),
    };
  }
  if (kind === 14 && version >= 33) {
    const cropLeft = reader.f32();
    const cropTop = reader.f32();
    const cropRight = reader.f32();
    const cropBottom = reader.f32();
    const mediaType = reader.string();
    const bytes = reader.byteArray();
    const resourceId = version >= 35 ? reader.u32() : 0;
    const maskResourceId = reader.u32();
    const width = reader.u32();
    const height = reader.u32();
    const maskMediaType = reader.string();
    const maskBytes = reader.byteArray();
    const pixels = width * height;
    if (!Number.isSafeInteger(pixels) || pixels <= 0 || maskMediaType.length === 0) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid image alpha mask");
    }
    const alphaMask = reader.imageMask(maskResourceId, {
      width,
      height,
      mediaType: maskMediaType,
      bytes: maskBytes,
    });
    const image = reader.imageResource(resourceId, { mediaType, bytes });
    return {
      kind: "image",
      cropLeft,
      cropTop,
      cropRight,
      cropBottom,
      ...image,
      alphaMask,
    };
  }
  if (kind === 15 && version >= 43) {
    const count = reader.u32();
    if (count > MAX_VISUAL_BRUSH_CHILDREN || count > Math.floor(reader.remaining / 20)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid visual group child count");
    }
    const children = [];
    for (let index = 0; index < count; index += 1) {
      const bounds = {
        x: reader.f32(), y: reader.f32(), width: reader.f32(), height: reader.f32(),
      };
      const childKind = reader.u8();
      reader.skip(3);
      children.push({ bounds, visual: readVisual(reader, childKind, version, depth + 1) });
    }
    return { kind: "group", children };
  }
  if (kind === 16 && version >= 43) {
    const mask = readPaint(reader, version, depth);
    const nestedKind = reader.u8();
    reader.skip(3);
    return { kind: "opacity-mask", mask, visual: readVisual(reader, nestedKind, version, depth + 1) };
  }
  if (kind === 17 && version >= 45) {
    const sourceProfile = reader.byteArray();
    const hasDestination = enumValue([false, true] as const, reader.u8(), "ICC destination profile flag");
    reader.skip(3);
    const destinationProfile = hasDestination ? reader.byteArray() : undefined;
    const nestedKind = reader.u8();
    reader.skip(3);
    return {
      kind: "color-managed-image", sourceProfile,
      ...(destinationProfile === undefined ? {} : { destinationProfile }),
      visual: readVisual(reader, nestedKind, version, depth + 1),
    };
  }
  if (kind === 19 && version >= 46) {
    const count = reader.u32();
    if (count === 0 || count > MAX_RICH_TEXT_RUNS || count > Math.floor(reader.remaining / 24)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid text-effect count");
    }
    const effects = [];
    for (let index = 0; index < count; index += 1) {
      const flags = reader.u8();
      reader.skip(3);
      const supportedFlags = version >= 53 ? 255
        : version >= 52 ? 127 : version >= 51 ? 63 : version >= 50 ? 31 : 7;
      if ((flags & ~supportedFlags) !== 0) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid text-effect flags");
      }
      const shadowScaleX = reader.f32();
      const shadowScaleY = reader.f32();
      const shadowSkewX = reader.f32();
      const shadowSkewY = reader.f32();
      const shadowAlignment = reader.u8();
      const extraFlags = version >= 60 ? reader.u8() : 0;
      if ((extraFlags & ~(version >= 75 ? 15 : 1)) !== 0) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid extended text-effect flags");
      }
      reader.skip(version >= 60 ? 2 : 3);
      const readShadow = () => ({
        color: reader.u32(),
        blur: reader.f32(),
        offsetX: reader.f32(),
        offsetY: reader.f32(),
      });
      const shadow = (flags & 1) === 0 ? undefined : readShadow();
      const innerShadow = (flags & 2) === 0 ? undefined : readShadow();
      const reflection = (extraFlags & 1) === 0 ? undefined : {
        startOpacity: reader.f32(),
        endOpacity: reader.f32(),
        startPosition: reader.f32(),
        endPosition: reader.f32(),
        directionDegrees: reader.f32(),
        blur: reader.f32(),
        distance: reader.f32(),
        scaleX: reader.f32(),
        scaleY: reader.f32(),
      };
      const glow = (extraFlags & 2) === 0 ? undefined : { color: reader.u32(), radius: reader.f32() };
      const strokeWidth = (extraFlags & 4) === 0 ? undefined : reader.f32();
      const stroke = strokeWidth === undefined ? undefined : readPaint(reader, version, depth + 1);
      effects.push({
        ...(glow === undefined ? {} : { glow }),
        ...(stroke === undefined ? {} : { stroke, strokeWidth: strokeWidth! }),
        ...((extraFlags & 8) === 0 ? {} : { fillToText: true }),
        wavyUnderline: (flags & 4) !== 0,
        dottedUnderline: version >= 50 && (flags & 8) !== 0,
        heavyUnderline: version >= 50 && (flags & 16) !== 0,
        doubleUnderline: version >= 51 && (flags & 32) !== 0,
        dotDashUnderline: version >= 52 && (flags & 64) !== 0,
        doubleStrikethrough: version >= 53 && (flags & 128) !== 0,
        shadowScaleX,
        shadowScaleY,
        shadowSkewX,
        shadowSkewY,
        shadowAlignment,
        ...(shadow === undefined ? {} : { shadow }),
        ...(innerShadow === undefined ? {} : { innerShadow }),
        ...(reflection === undefined ? {} : { reflection }),
      });
    }
    const nestedKind = reader.u8();
    reader.skip(3);
    return {
      kind: "text-effects",
      effects,
      visual: readVisual(reader, nestedKind, version, depth + 1),
    };
  }
  throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid visual kind ${kind}`);
}

function readDiagnostic(reader: Reader): Diagnostic {
  const code = reader.string();
  const severity = enumValue(SEVERITIES, reader.u8(), "diagnostic severity");
  const fidelity = enumValue(FIDELITIES, reader.u8(), "diagnostic fidelity");
  const phase = enumValue(PHASES, reader.u8(), "diagnostic phase");
  reader.skip(1);
  const message = reader.string();
  const part = reader.optionalString();
  const objectId = reader.optionalString();
  const fields = readFields(reader);
  const details = Object.fromEntries(fields);
  return {
    code,
    severity,
    fidelity,
    phase,
    message,
    ...(part === undefined ? {} : { part }),
    ...(objectId === undefined ? {} : { objectId }),
    ...(fields.size === 0 ? {} : { details }),
  };
}

function validateGeometry(
  geometry: SceneGeometry,
  label: string,
  invalid: (reason: string) => never,
): void {
  if (typeof geometry === "string") return;
  if (geometry.kind === "rounded-rectangle") {
    if (![geometry.radiusX, geometry.radiusY].every(Number.isFinite)
      || geometry.radiusX < 0 || geometry.radiusY < 0) {
      invalid(`${label} has invalid rounded-rectangle radii`);
    }
    return;
  }
  const commandSets = geometry.kind === "path"
    ? [geometry.commands]
    : geometry.layers.map(({ commands }) => commands);
  for (const commands of commandSets) {
    for (const command of commands) {
      const coordinates = command.kind === "closePath"
        ? []
        : command.kind === "moveTo" || command.kind === "lineTo"
          ? [command.x, command.y]
          : command.kind === "quadraticCurveTo"
            ? [command.cpx, command.cpy, command.x, command.y]
            : [command.cp1x, command.cp1y, command.cp2x, command.cp2y, command.x, command.y];
      if (!coordinates.every(Number.isFinite)) invalid(`${label} has an invalid path coordinate`);
    }
  }
}

function validatePaint(
  paint: ScenePaint,
  label: string,
  invalid: (reason: string) => never,
  depth: number,
): void {
  if (depth > MAX_VISUAL_DEPTH) invalid(`${label} has too many nested visual brushes`);
  if (paint.kind === "none" || paint.kind === "solid") return;
  if (paint.kind === "mapped-gradient") {
    if (!Object.values(paint.tile).every(Number.isFinite) || paint.tile.left + paint.tile.right >= 1
      || paint.tile.top + paint.tile.bottom >= 1) invalid(`${label} has invalid gradient mapping`);
    validatePaint(paint.paint, label, invalid, depth + 1);
    return;
  }
  if (paint.kind === "pattern") {
    if (paint.preset.length === 0 || paint.preset.length > 64) {
      invalid(`${label} has an invalid pattern fill`);
    }
    return;
  }
  if (paint.kind === "image") {
    const { cropLeft, cropTop, cropRight, cropBottom } = paint;
    if (paint.mediaType.length === 0
      || paint.bytes.length === 0
      || ![cropLeft, cropTop, cropRight, cropBottom].every((value) => (
        Number.isFinite(value) && value >= -1 && value <= 1
      ))
      || Math.fround(cropLeft + cropRight) > 1
      || Math.fround(cropTop + cropBottom) > 1
      || paint.tileWidth !== undefined && (!Number.isFinite(paint.tileWidth) || paint.tileWidth <= 0)
      || paint.tileHeight !== undefined && (!Number.isFinite(paint.tileHeight) || paint.tileHeight <= 0)) {
      invalid(`${label} has an invalid image fill`);
    }
    return;
  }
  if (paint.kind === "visual") {
    const rects = [paint.viewbox, paint.viewport];
    const transforms = [paint.transform, paint.relativeTransform];
    if (rects.some((rect) => ![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite)
        || rect.width <= 0 || rect.height <= 0)
      || transforms.some(({ a, b, c, d, e, f }) => ![a, b, c, d, e, f].every(Number.isFinite))
      || !Number.isFinite(paint.opacity) || paint.opacity < 0 || paint.opacity > 1
      || !Number.isFinite(paint.alignmentX) || paint.alignmentX < 0 || paint.alignmentX > 1
      || !Number.isFinite(paint.alignmentY) || paint.alignmentY < 0 || paint.alignmentY > 1
      || paint.children.length > MAX_VISUAL_BRUSH_CHILDREN) {
      invalid(`${label} has an invalid visual brush`);
    }
    for (const child of paint.children) {
      if (![child.bounds.x, child.bounds.y, child.bounds.width, child.bounds.height].every(Number.isFinite)
        || child.bounds.width < 0 || child.bounds.height < 0) {
        invalid(`${label} has invalid visual brush bounds`);
      }
      validateVisual(child.visual, `${label} visual brush child`, invalid, depth + 1);
    }
    return;
  }
  if (paint.kind === "xps-gradient") {
    const values = [paint.start.x, paint.start.y, paint.end.x, paint.end.y, paint.radiusX, paint.radiusY,
      ...Object.values(paint.transform), ...Object.values(paint.relativeTransform)];
    if (!values.every(Number.isFinite) || paint.radiusX < 0 || paint.radiusY < 0) {
      invalid(`${label} has an invalid XPS gradient`);
    }
    let previous = -Infinity;
    for (const stop of paint.stops) {
      if (!Number.isFinite(stop.offset) || stop.offset < previous) {
        invalid(`${label} has invalid gradient stops`);
      }
      if (stop.color.kind === "context"
        && (!Number.isFinite(stop.color.alpha) || stop.color.alpha < 0 || stop.color.alpha > 1
          || stop.color.profile.length === 0 || stop.color.channels.length === 0
          || stop.color.channels.length > 15 || !stop.color.channels.every(Number.isFinite))) {
        invalid(`${label} has an invalid XPS ContextColor`);
      }
      previous = stop.offset;
    }
    return;
  }
  const coordinates = paint.kind === "linear-gradient"
    ? [paint.start.x, paint.start.y, paint.end.x, paint.end.y]
    : paint.kind === "shape-gradient"
      ? Object.values(paint.focus)
      : paint.kind === "rect-gradient"
      ? [paint.center.x, paint.center.y]
      : [
        paint.start.x,
        paint.start.y,
        paint.start.radius,
        paint.end.x,
        paint.end.y,
        paint.end.radius,
      ];
  if (!coordinates.every(Number.isFinite)
    || (paint.kind === "radial-gradient" || paint.kind === "circle-gradient")
      && (paint.start.radius < 0 || paint.end.radius < 0)
    || paint.stops.length === 0) {
    invalid(`${label} has an invalid gradient`);
  }
  let previous = -1;
  for (const stop of paint.stops) {
    if (!Number.isFinite(stop.offset) || stop.offset < 0 || stop.offset > 1 || stop.offset < previous) {
      invalid(`${label} has invalid gradient stops`);
    }
    previous = stop.offset;
  }
}

function validateVisual(
  visual: SceneVisual,
  label: string,
  invalid: (reason: string) => never,
  depth = 0,
): void {
  if (depth > MAX_VISUAL_DEPTH) invalid(`${label} has too many nested visual layers`);
  if (visual.kind === "none") return;
  if (visual.kind === "group") {
    if (visual.children.length > MAX_VISUAL_BRUSH_CHILDREN) invalid(`${label} has too many group children`);
    for (const child of visual.children) {
      const { x, y, width, height } = child.bounds;
      if (![x, y, width, height].every(Number.isFinite) || width < 0 || height < 0) {
        invalid(`${label} has invalid group child bounds`);
      }
      validateVisual(child.visual, label, invalid, depth + 1);
    }
    return;
  }
  if (visual.kind === "opacity-mask") {
    validatePaint(visual.mask, `${label} opacity mask`, invalid, depth);
    validateVisual(visual.visual, label, invalid, depth + 1);
    return;
  }
  if (visual.kind === "color-managed-image") {
    if (visual.sourceProfile.length === 0 || visual.destinationProfile?.length === 0) {
      invalid(`${label} has an invalid ICC profile`);
    }
    validateVisual(visual.visual, label, invalid, depth + 1);
    return;
  }
  if (visual.kind === "layer") {
    const { a, b, c, d, e, f } = visual.transform;
    if (![a, b, c, d, e, f].every(Number.isFinite)
      || !Number.isFinite(visual.opacity) || visual.opacity < 0 || visual.opacity > 1) {
      invalid(`${label} has an invalid compositing layer`);
    }
    validateVisual(visual.visual, label, invalid, depth + 1);
    return;
  }
  if (visual.kind === "effect") {
    if (visual.shadow === undefined && visual.clip === undefined) {
      invalid(`${label} has an empty visual effect`);
    }
    if (visual.shadow !== undefined
      && (![visual.shadow.blur, visual.shadow.offsetX, visual.shadow.offsetY].every(Number.isFinite)
        || visual.shadow.blur < 0)) {
      invalid(`${label} has an invalid shadow`);
    }
    if (visual.clip !== undefined) validateGeometry(visual.clip, `${label} clip`, invalid);
    validateVisual(visual.visual, label, invalid, depth + 1);
    return;
  }
  if (visual.kind === "text-layout") {
    const layout = visual.layout;
    const fill = layout.fillCharacter;
    if (fill !== undefined && (!Number.isInteger(fill.offset) || fill.offset < 0 || fill.offset >= 0xffff_ffff
      || [...fill.character].length !== 1 || /[\p{Cc}\p{Cs}]/u.test(fill.character))) {
      invalid(`${label} has an invalid fill character`);
    }
    if (!Number.isFinite(layout.defaultTabStop)
      || layout.defaultTabStop <= 0
      || !Number.isFinite(layout.hangingIndent)
      || layout.hangingIndent < 0
      || layout.paragraphSpacing !== undefined
        && (!Number.isFinite(layout.paragraphSpacing) || layout.paragraphSpacing < 0)
      || !Number.isFinite(layout.insetLeft)
      || layout.insetLeft < 0
      || !Number.isFinite(layout.insetRight)
      || layout.insetRight < 0
      || !Number.isFinite(layout.insetTop)
      || layout.insetTop < 0
      || !Number.isFinite(layout.insetBottom)
      || layout.insetBottom < 0
      || !Number.isFinite(layout.marginLeft)
      || layout.marginLeft < 0
      || !Number.isFinite(layout.marginRight)
      || layout.marginRight < 0
      || !Number.isFinite(layout.firstLineIndent)
      || !Number.isFinite(layout.minScale)
      || layout.minScale <= 0
      || layout.minScale > 1
      || !Number.isInteger(layout.columnCount)
      || layout.columnCount <= 0
      || layout.columnCount > 64
      || !Number.isFinite(layout.columnSpacing)
      || layout.columnSpacing < 0
      || !Number.isFinite(layout.rotationDegrees)
      || !Number.isFinite(layout.fontScale)
      || layout.fontScale <= 0
      || layout.fontScale > 1
      || !Number.isFinite(layout.lineSpacingReduction)
      || layout.lineSpacingReduction < 0
      || layout.lineSpacingReduction >= 1
      || (layout.wrapRegions?.length ?? 0) > 1024
      || layout.wrapRegions?.some(rect => !Number.isFinite(rect.x) || !Number.isFinite(rect.y)
        || !Number.isFinite(rect.width) || !Number.isFinite(rect.height) || rect.width <= 0 || rect.height <= 0)
      || (layout.paragraphs?.length ?? 0) > MAX_TEXT_PARAGRAPHS
      || layout.tabStops.some((stop, index) => (
        !Number.isFinite(stop.position)
        || stop.position <= 0
        || index > 0 && stop.position <= (layout.tabStops[index - 1]?.position ?? 0)
      ))
      || layout.paragraphs?.some((paragraph) => (
        !Number.isFinite(paragraph.marginLeft)
        || paragraph.marginLeft < 0
        || !Number.isFinite(paragraph.marginRight)
        || paragraph.marginRight < 0
        || !Number.isFinite(paragraph.firstLineIndent)
        || !Number.isFinite(paragraph.defaultTabStop)
        || paragraph.defaultTabStop <= 0
        || !Number.isFinite(paragraph.lineHeight ?? 0)
        || (paragraph.lineHeight ?? 0) < 0
        || !Number.isFinite(paragraph.spaceBefore ?? 0)
        || (paragraph.spaceBefore ?? 0) < 0
        || !Number.isFinite(paragraph.spaceAfter ?? 0)
        || (paragraph.spaceAfter ?? 0) < 0
        || (paragraph.dropCap !== undefined && (
          !Number.isInteger(paragraph.dropCap.characters) || paragraph.dropCap.characters < 1 || paragraph.dropCap.characters > 32
          || !Number.isInteger(paragraph.dropCap.lines) || paragraph.dropCap.lines < 1 || paragraph.dropCap.lines > 32
          || !Number.isInteger(paragraph.dropCap.raisedLines) || paragraph.dropCap.raisedLines < 0 || paragraph.dropCap.raisedLines > 32
          || !Number.isFinite(paragraph.dropCap.padding) || paragraph.dropCap.padding < 0 || paragraph.dropCap.padding > 2048
          || !Number.isFinite(paragraph.dropCap.outdent) || Math.abs(paragraph.dropCap.outdent) > 2048
        ))
        || [paragraph.ruleAbove, paragraph.ruleBelow].some(rule => rule !== undefined && (
          !Number.isFinite(rule.strokeWidth) || rule.strokeWidth <= 0
          || !Number.isFinite(rule.offsetX) || !Number.isFinite(rule.offsetY)
          || !Number.isFinite(rule.width) || rule.width <= 0
        ))
      ))) {
      invalid(`${label} has invalid text layout settings`);
    }
    validateVisual(visual.visual, label, invalid, depth + 1);
    return;
  }
  if (visual.kind === "stroke-style") {
    if (!Number.isFinite(visual.style.dashOffset ?? 0) || !Number.isFinite(visual.style.miterLimit)
      || visual.style.miterLimit <= 0
      || visual.style.dash.length > 256
      || visual.style.dash.some((value) => !Number.isFinite(value) || value < 0)
      || (visual.style.dash.length > 0 && !visual.style.dash.some((value) => value > 0))) {
      invalid(`${label} has invalid line styling`);
    }
    validateVisual(visual.visual, label, invalid, depth + 1);
    return;
  }
  if (visual.kind === "text-effects") {
    const shadows = visual.effects.flatMap((effect) => [effect.shadow, effect.innerShadow])
      .filter((shadow) => shadow !== undefined);
    if (visual.effects.length === 0
      || visual.effects.length > MAX_RICH_TEXT_RUNS
      || shadows.length === 0 && !visual.effects.some((effect) => effect.reflection !== undefined
        || effect.stroke !== undefined || effect.glow !== undefined || effect.fillToText === true
        || effect.wavyUnderline
        || effect.dottedUnderline || effect.heavyUnderline || effect.doubleUnderline
        || effect.dotDashUnderline || effect.doubleStrikethrough)
      || visual.effects.some((effect) => ![
        effect.shadowScaleX,
        effect.shadowScaleY,
        effect.shadowSkewX,
        effect.shadowSkewY,
      ].every(Number.isFinite) || effect.shadowAlignment < 0 || effect.shadowAlignment > 8)
      || shadows.some((shadow) => ![shadow.blur, shadow.offsetX, shadow.offsetY]
        .every(Number.isFinite) || shadow.blur < 0)
      || visual.effects.some((effect) => effect.reflection !== undefined
        && ![
          effect.reflection.startOpacity,
          effect.reflection.endOpacity,
          effect.reflection.startPosition,
          effect.reflection.endPosition,
          effect.reflection.directionDegrees,
          effect.reflection.blur,
          effect.reflection.distance,
          effect.reflection.scaleX,
          effect.reflection.scaleY,
        ].every(Number.isFinite))) {
      invalid(`${label} has invalid text effects`);
    }
    for (const effect of visual.effects) {
      if (effect.stroke !== undefined) {
        validatePaint(effect.stroke, label, invalid, depth + 1);
        if (!Number.isFinite(effect.strokeWidth) || (effect.strokeWidth ?? -1) < 0) invalid(`${label} has invalid text outline`);
      }
      if (effect.glow !== undefined && (!Number.isFinite(effect.glow.radius) || effect.glow.radius < 0)) invalid(`${label} has invalid text glow`);
    }
    validateVisual(visual.visual, label, invalid, depth + 1);
    return;
  }
  if (visual.kind === "advanced-effect") {
    if (visual.outerShadow === undefined
      && visual.innerShadow === undefined
      && visual.glow === undefined
      && visual.reflection === undefined
      && visual.softEdge === undefined
      && visual.threeD === undefined) {
      invalid(`${label} has an empty advanced effect`);
    }
    if (visual.outerShadow !== undefined
      && (![visual.outerShadow.blur, visual.outerShadow.offsetX, visual.outerShadow.offsetY,
        visual.outerShadow.scaleX, visual.outerShadow.scaleY, visual.outerShadow.skewX,
        visual.outerShadow.skewY].every(Number.isFinite)
        || visual.outerShadow.blur < 0
        || visual.outerShadow.alignment < 0 || visual.outerShadow.alignment > 8)
      || visual.innerShadow !== undefined
      && (![visual.innerShadow.blur, visual.innerShadow.offsetX, visual.innerShadow.offsetY]
        .every(Number.isFinite)
        || visual.innerShadow.blur < 0)
      || visual.glow !== undefined
        && (!Number.isFinite(visual.glow.radius) || visual.glow.radius < 0)
      || visual.reflection !== undefined
        && (![visual.reflection.startOpacity, visual.reflection.endOpacity,
          visual.reflection.startPosition, visual.reflection.endPosition,
          visual.reflection.directionDegrees, visual.reflection.blur, visual.reflection.distance,
          visual.reflection.scaleX, visual.reflection.scaleY].every(Number.isFinite)
          || visual.reflection.startOpacity < 0 || visual.reflection.startOpacity > 1
          || visual.reflection.endOpacity < 0 || visual.reflection.endOpacity > 1
          || visual.reflection.startPosition < 0 || visual.reflection.startPosition > 1
          || visual.reflection.endPosition < 0 || visual.reflection.endPosition > 1
          || visual.reflection.startPosition > visual.reflection.endPosition
          || visual.reflection.blur < 0
          || visual.reflection.scaleX <= 0
          || visual.reflection.scaleY <= 0)
      || visual.softEdge !== undefined
        && (!Number.isFinite(visual.softEdge) || visual.softEdge < 0)
      || visual.threeD !== undefined
        && (![visual.threeD.cameraFov, visual.threeD.cameraZoom,
          visual.threeD.cameraLatitude, visual.threeD.cameraLongitude,
          visual.threeD.cameraRevolution, visual.threeD.lightLatitude,
          visual.threeD.lightLongitude, visual.threeD.lightRevolution,
          visual.threeD.z, visual.threeD.extrusionHeight, visual.threeD.contourWidth,
          visual.threeD.flatTextZ ?? 0]
          .every(Number.isFinite)
          || visual.threeD.cameraFov < 0
          || visual.threeD.cameraZoom <= 0
          || visual.threeD.extrusionHeight < 0
          || visual.threeD.contourWidth < 0
          || [visual.threeD.cameraPreset, visual.threeD.lightRig,
            visual.threeD.lightDirection, visual.threeD.material]
            .some((value) => value.length === 0)
          || visual.threeD.backdrop !== undefined
            && (![visual.threeD.backdrop.anchorX, visual.threeD.backdrop.anchorY,
              visual.threeD.backdrop.anchorZ, visual.threeD.backdrop.normalX,
              visual.threeD.backdrop.normalY, visual.threeD.backdrop.normalZ,
              visual.threeD.backdrop.upX, visual.threeD.backdrop.upY,
              visual.threeD.backdrop.upZ].every(Number.isFinite)
              || Math.abs(visual.threeD.backdrop.normalX)
                + Math.abs(visual.threeD.backdrop.normalY)
                + Math.abs(visual.threeD.backdrop.normalZ) === 0
              || Math.abs(visual.threeD.backdrop.upX)
                + Math.abs(visual.threeD.backdrop.upY)
                + Math.abs(visual.threeD.backdrop.upZ) === 0)
          || [visual.threeD.bevelTop, visual.threeD.bevelBottom]
            .some((bevel) => bevel !== undefined
              && (![bevel.width, bevel.height].every(Number.isFinite)
                || bevel.width < 0 || bevel.height < 0 || bevel.preset.length === 0)))) {
      invalid(`${label} has invalid advanced effects`);
    }
    validateVisual(visual.visual, label, invalid, depth + 1);
    return;
  }
  if (visual.kind === "image-color-change") {
    validateVisual(visual.visual, label, invalid, depth + 1);
    return;
  }
  if (visual.kind === "image-adjustment") {
    if (!Number.isFinite(visual.brightness) || visual.brightness < -1 || visual.brightness > 1
      || !Number.isFinite(visual.contrast) || visual.contrast < -1 || visual.contrast > 1
      || visual.bilevelThreshold !== undefined
        && (!Number.isFinite(visual.bilevelThreshold)
          || visual.bilevelThreshold < 0 || visual.bilevelThreshold > 1)) {
      invalid(`${label} has invalid image adjustments`);
    }
    validateVisual(visual.visual, label, invalid, depth + 1);
    return;
  }
  if (visual.kind === "media") {
    if (visual.mediaType.length === 0 || visual.bytes.length === 0) {
      invalid(`${label} has invalid embedded media`);
    }
    validateVisual(visual.visual, label, invalid, depth + 1);
    return;
  }
  if (visual.kind === "image") {
    const { cropLeft, cropTop, cropRight, cropBottom } = visual;
    if (![cropLeft, cropTop, cropRight, cropBottom].every((value) => (
      Number.isFinite(value) && value >= -1 && value <= 1
    )) || Math.fround(cropLeft + cropRight) > 1 || Math.fround(cropTop + cropBottom) > 1) {
      invalid(`${label} has an invalid image crop`);
    }
    return;
  }
  validateGeometry(visual.geometry, label, invalid);
  if (!Number.isFinite(visual.strokeWidth) || visual.strokeWidth < 0) {
    invalid(`${label} has an invalid stroke width`);
  }
  if (visual.kind === "shape" || visual.kind === "text") {
    if (visual.kind === "text" && (!Number.isFinite(visual.fontSize) || visual.fontSize <= 0)) {
      invalid(`${label} has an invalid font size`);
    }
    return;
  }
  validatePaint(visual.fill, label, invalid, depth);
  validatePaint(visual.stroke, label, invalid, depth);
  if (visual.kind === "rich-text") {
    if (!Number.isFinite(visual.lineHeight) || visual.lineHeight < 0) {
      invalid(`${label} has an invalid line height`);
    }
    for (const run of visual.runs) {
      if (run.paint !== undefined) validatePaint(run.paint, label, invalid, depth + 1);
      if (!Number.isFinite(run.fontSize)
        || run.fontSize <= 0
        || !Number.isFinite(run.letterSpacing)
        || !Number.isFinite(run.baselineShift)) {
        invalid(`${label} has an invalid rich-text run`);
      }
    }
  }
}

function validateScene(
  fatal: boolean,
  format: DocumentFormat | undefined,
  kind: DocumentKind | undefined,
  units: readonly UnitDescriptor[],
  outline: readonly DocumentOutlineItem[],
  objects: readonly SceneObject[],
): void {
  const invalid = (reason: string): never => {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `Invalid core snapshot: ${reason}`);
  };
  if (format === undefined || kind === undefined) {
    if (!fatal || units.length !== 0 || objects.length !== 0) invalid("scene lacks document identity");
    return;
  }
  const expectedKind: DocumentKind = format === "pptx" || format === "odp" || format === "ppt" || format === "keynote"
    ? "presentation"
    : format === "xlsx" || format === "ods" || format === "csv" || format === "xls" || format === "numbers"
      ? "spreadsheet"
      : "text";
  if (kind !== expectedKind) invalid(`format ${format} cannot have document kind ${kind}`);
  const expectedUnitType = kind === "presentation" ? "slide" : kind === "spreadsheet" ? "sheet" : "page";
  for (let index = 0; index < units.length; index += 1) {
    const unit = units[index]!;
    if (unit.index !== index) invalid("unit indexes must be contiguous and ordered");
    if (unit.type !== expectedUnitType) invalid(`document kind ${kind} cannot contain a ${unit.type} unit`);
    if (!Number.isFinite(unit.width) || unit.width <= 0 || !Number.isFinite(unit.height) || unit.height <= 0) {
      invalid(`unit ${index} has invalid dimensions`);
    }
    if (unit.type === "sheet" && (
      !Number.isInteger(unit.frozenRows)
      || unit.frozenRows < 0
      || unit.frozenRows > unit.rows
      || !Number.isInteger(unit.frozenColumns)
      || unit.frozenColumns < 0
      || unit.frozenColumns > unit.columns
      || !Number.isFinite(unit.frozenWidth)
      || unit.frozenWidth < 0
      || unit.frozenWidth > unit.width
      || !Number.isFinite(unit.frozenHeight)
      || unit.frozenHeight < 0
      || unit.frozenHeight > unit.height
    )) invalid(`sheet ${index} has invalid frozen panes`);
    if (unit.type === "sheet") {
      for (const [name, axis, count] of [
        ["row", unit.rowAxis, unit.rows],
        ["column", unit.columnAxis, unit.columns],
      ] as const) {
        if (!Number.isFinite(axis.defaultSize) || axis.defaultSize < 0) {
          invalid(`sheet ${index} has an invalid ${name} axis`);
        }
        let next = 0;
        for (const span of axis.spans) {
          if (!Number.isInteger(span.start) || !Number.isInteger(span.end)
            || span.start < next || span.start > span.end || span.end >= count
            || !Number.isFinite(span.size) || span.size < 0) {
            invalid(`sheet ${index} has an invalid ${name} axis span`);
          }
          next = span.end + 1;
        }
      }
    }
    if (unit.type === "slide" && unit.speakerNoteParagraphs !== undefined) {
      if (unit.speakerNoteParagraphs.length > MAX_TEXT_PARAGRAPHS) {
        invalid(`slide ${index} has too many speaker-note paragraphs`);
      }
      for (const paragraph of unit.speakerNoteParagraphs) {
        if (![paragraph.marginLeft, paragraph.marginRight, paragraph.firstLineIndent,
          paragraph.defaultTabStop, paragraph.lineHeight, paragraph.spaceBefore, paragraph.spaceAfter]
          .every(Number.isFinite)
          || paragraph.marginLeft < 0
          || paragraph.marginRight < 0
          || paragraph.defaultTabStop <= 0
          || paragraph.lineHeight < 0
          || paragraph.spaceBefore < 0
          || paragraph.spaceAfter < 0
          || paragraph.runs.length > MAX_RICH_TEXT_RUNS) {
          invalid(`slide ${index} has an invalid speaker-note paragraph`);
        }
        for (const run of paragraph.runs) {
          if (!Number.isFinite(run.fontSize) || run.fontSize <= 0
            || !Number.isFinite(run.letterSpacing)
            || !Number.isFinite(run.baselineShift)) {
            invalid(`slide ${index} has an invalid speaker-note run`);
          }
        }
      }
    }
  }
  for (const item of outline) {
    if (item.title.length === 0
      || !Number.isInteger(item.unitIndex)
      || item.unitIndex < 0
      || item.unitIndex >= units.length
      || !Number.isInteger(item.level)
      || item.level < 0) {
      invalid("document outline contains an invalid item");
    }
  }

  const byNumericId = new Map<number, SceneObject>();
  const stableIds = new Set<string>();
  for (const object of objects) {
    if (byNumericId.has(object.numericId)) invalid(`duplicate object numeric ID ${object.numericId}`);
    if (stableIds.has(object.id)) invalid(`duplicate object ID ${object.id}`);
    byNumericId.set(object.numericId, object);
    stableIds.add(object.id);
    const unit = units[object.unitIndex];
    if (unit === undefined || unit.index !== object.unitIndex) {
      invalid(`object ${object.id} refers to a missing unit`);
    }
    const { x, y, width, height } = object.bounds;
    if (![x, y, width, height].every(Number.isFinite) || width < 0 || height < 0) {
      invalid(`object ${object.id} has invalid bounds`);
    }
    validateVisual(object.visual, `object ${object.id}`, invalid);
  }

  for (const object of objects) {
    if (object.parentNumericId === undefined) {
      if (object.parentId !== undefined) invalid(`object ${object.id} has only a string parent ID`);
      continue;
    }
    const parent = byNumericId.get(object.parentNumericId);
    if (parent === undefined) invalid(`object ${object.id} refers to a missing parent`);
    if (object.parentId !== parent?.id) invalid(`object ${object.id} has inconsistent parent IDs`);
  }

  const state = new Map<number, 1 | 2>();
  for (const object of objects) {
    if (state.get(object.numericId) === 2) continue;
    const path: number[] = [];
    let current: SceneObject | undefined = object;
    while (current !== undefined) {
      const currentState = state.get(current.numericId);
      if (currentState === 1) invalid(`object ${current.id} participates in a parent cycle`);
      if (currentState === 2) break;
      state.set(current.numericId, 1);
      path.push(current.numericId);
      current = current.parentNumericId === undefined ? undefined : byNumericId.get(current.parentNumericId);
    }
    for (const numericId of path) state.set(numericId, 2);
  }
}

export function decodeSnapshot(bytes: Uint8Array): SceneDocument {
  const reader = new Reader(bytes);
  const magic = reader.u32();
  const version = reader.u16();
  if (magic !== MAGIC || version < MIN_VERSION || version > VERSION) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Unsupported core snapshot header");
  }
  const fatal = reader.u8() !== 0;
  const formatCode = reader.u8();
  const kindCode = reader.u8();
  reader.skip(3);

  const diagnosticCount = reader.u32();
  const diagnostics: Diagnostic[] = [];
  for (let index = 0; index < diagnosticCount; index += 1) {
    diagnostics.push(readDiagnostic(reader));
  }

  const unitCount = reader.u32();
  const units: UnitDescriptor[] = [];
  for (let index = 0; index < unitCount; index += 1) {
    const type = enumValue(UNIT_TYPES, reader.u8(), "unit type");
    reader.skip(3);
    const unitIndex = reader.u32();
    const id = reader.string();
    const name = reader.string();
    const width = reader.f32();
    const height = reader.f32();
    const rows = reader.u32();
    const columns = reader.u32();
    const frozenRows = version >= 5 ? reader.u32() : 0;
    const frozenColumns = version >= 5 ? reader.u32() : 0;
    const frozenWidth = version >= 5 ? reader.f32() : 0;
    const frozenHeight = version >= 5 ? reader.f32() : 0;
    const rowAxis = version >= 27 ? readSheetAxis(reader) : uniformSheetAxis(rows, height);
    const columnAxis = version >= 27 ? readSheetAxis(reader) : uniformSheetAxis(columns, width);
    const showGridLines = version >= 55 ? reader.u8() !== 0 : false;
    if (version >= 55) reader.skip(3);
    let tabColor: number | undefined;
    if (version >= 73) {
      const hasTabColor = reader.u8() !== 0;
      reader.skip(3);
      const color = reader.u32();
      if (hasTabColor) tabColor = color;
    }
    const sourceId = version >= 11 ? reader.optionalString() : undefined;
    const sourcePart = version >= 11 ? reader.optionalString() : undefined;
    const speakerNotes = version >= 17 ? reader.optionalString() : undefined;
    const speakerNotesPart = version >= 17 ? reader.optionalString() : undefined;
    const speakerNoteParagraphs: SpeakerNoteParagraph[] = [];
    if (version >= 30) {
      const paragraphCount = reader.u32();
      if (paragraphCount > MAX_TEXT_PARAGRAPHS || paragraphCount > Math.floor(reader.remaining / 40)) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid speaker-note paragraph count");
      }
      for (let paragraphIndex = 0; paragraphIndex < paragraphCount; paragraphIndex += 1) {
        speakerNoteParagraphs.push(readSpeakerNoteParagraph(reader, version));
      }
    }
    const slideNumber = version >= 11 ? reader.u32() : unitIndex + 1;
    const hidden = version >= 11 ? reader.u8() !== 0 : false;
    if (version >= 11) reader.skip(3);
    let printSettings: SheetPrintSettings | undefined;
    if (version >= 38) {
      const hasPrintSettings = reader.u8() !== 0;
      const orientationCode = reader.u8();
      const printFlags = reader.u8();
      reader.skip(1);
      const paperSizeValue = reader.u32();
      const scaleValue = reader.u32();
      const fitToWidthValue = reader.u32();
      const fitToHeightValue = reader.u32();
      const marginValues = Array.from({ length: 6 }, () => reader.f32());
      const printArea = reader.optionalString();
      const oddHeader = reader.optionalString();
      const oddFooter = reader.optionalString();
      const evenHeader = reader.optionalString();
      const evenFooter = reader.optionalString();
      const firstHeader = reader.optionalString();
      const firstFooter = reader.optionalString();
      const printTitles = version >= 74 ? reader.optionalString() : undefined;
      const readBreaks = (limit: number, perpendicular: number): readonly (readonly [number, number, number])[] => {
        if (version < 74) return [];
        const count = reader.u32();
        if (count > Math.floor(reader.remaining / 12)) throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid page break count");
        return Array.from({ length: count }, () => {
          const id = reader.u32(), min = reader.u32(), max = reader.u32();
          if (id === 0 || id >= limit || min > max || max >= perpendicular) throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid page break");
          return [id, min, max] as const;
        });
      };
      const rowBreaks = readBreaks(1_048_576, 16_384);
      const columnBreaks = readBreaks(16_384, 1_048_576);
      if (orientationCode > 2 || (printFlags & ~(version >= 59 ? 0x1f : 0x07)) !== 0
        || (printFlags & 0x18) === 0x18) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid sheet print settings");
      }
      for (const margin of marginValues) {
        if (!Number.isNaN(margin) && (!Number.isFinite(margin) || margin < 0)) {
          throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid sheet print margin");
        }
      }
      if (hasPrintSettings) {
        const [left, right, top, bottom, header, footer] = marginValues;
        printSettings = {
          ...((printFlags & 0x08) !== 0 ? { viewMode: "pageLayout" as const } : {}),
          ...((printFlags & 0x10) !== 0 ? { viewMode: "pageBreakPreview" as const } : {}),
          ...(printArea === undefined ? {} : { printArea }),
          ...(printTitles === undefined ? {} : { printTitles }),
          ...(rowBreaks.length === 0 ? {} : { rowBreaks }),
          ...(columnBreaks.length === 0 ? {} : { columnBreaks }),
          ...(paperSizeValue === NONE ? {} : { paperSize: paperSizeValue }),
          ...(orientationCode === 0 ? {} : {
            orientation: orientationCode === 1 ? "portrait" as const : "landscape" as const,
          }),
          ...(scaleValue === NONE ? {} : { scale: scaleValue }),
          ...(fitToWidthValue === NONE ? {} : { fitToWidth: fitToWidthValue }),
          ...(fitToHeightValue === NONE ? {} : { fitToHeight: fitToHeightValue }),
          fitToPage: (printFlags & 0x01) !== 0,
          margins: {
            ...(Number.isNaN(left) ? {} : { left }),
            ...(Number.isNaN(right) ? {} : { right }),
            ...(Number.isNaN(top) ? {} : { top }),
            ...(Number.isNaN(bottom) ? {} : { bottom }),
            ...(Number.isNaN(header) ? {} : { header }),
            ...(Number.isNaN(footer) ? {} : { footer }),
          },
          differentOddEven: (printFlags & 0x02) !== 0,
          differentFirst: (printFlags & 0x04) !== 0,
          ...(oddHeader === undefined ? {} : { oddHeader }),
          ...(oddFooter === undefined ? {} : { oddFooter }),
          ...(evenHeader === undefined ? {} : { evenHeader }),
          ...(evenFooter === undefined ? {} : { evenFooter }),
          ...(firstHeader === undefined ? {} : { firstHeader }),
          ...(firstFooter === undefined ? {} : { firstFooter }),
        };
      }
    }
    if (type === "sheet") units.push({
      type, index: unitIndex, id, name, width, height, rows, columns,
      frozenRows, frozenColumns, frozenWidth, frozenHeight, rowAxis, columnAxis, showGridLines,
      ...(tabColor === undefined ? {} : { tabColor }),
      ...(printSettings === undefined ? {} : { printSettings }),
    });
    else if (type === "slide") units.push({
      type,
      index: unitIndex,
      id,
      name,
      width,
      height,
      ...(sourceId === undefined ? {} : { sourceId }),
      ...(sourcePart === undefined ? {} : { sourcePart }),
      ...(speakerNotes === undefined ? {} : { speakerNotes }),
      ...(speakerNotesPart === undefined ? {} : { speakerNotesPart }),
      ...(speakerNoteParagraphs.length === 0 ? {} : { speakerNoteParagraphs }),
      slideNumber,
      hidden,
    });
    else units.push({ type, index: unitIndex, id, name, width, height });
  }

  const format = formatCode === 0 ? undefined : enumValue(FORMATS, formatCode, "format");
  const kind = kindCode === 0 ? undefined : enumValue(KINDS, kindCode, "document kind");
  const outline: DocumentOutlineItem[] = [];
  if (version >= 37) {
    const outlineCount = reader.u32();
    if (outlineCount > Math.floor(reader.remaining / 12)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid document outline count");
    }
    for (let index = 0; index < outlineCount; index += 1) {
      outline.push({
        title: reader.string(),
        unitIndex: reader.u32(),
        level: reader.u32(),
      });
    }
  }
  const embeddedFonts: SceneEmbeddedFont[] = [];
  if (version >= 6) {
    const fontCount = reader.u32();
    for (let index = 0; index < fontCount; index += 1) {
      const style = enumValue(FONT_STYLES, reader.u8(), "font style");
      reader.skip(3);
      const weight = reader.u32();
      const family = reader.string();
      const fontBytes = reader.byteArray();
      if (family.length === 0 || family.length > 256 || weight < 1 || weight > 1000 || fontBytes.length === 0) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid embedded font descriptor");
      }
      embeddedFonts.push({ family, bytes: fontBytes, style, weight });
    }
  }
  const fontAlternateNames: NonNullable<SceneDocument["fontAlternateNames"]>[number][] = [];
  if (version >= 76) {
    const count = reader.u32();
    if (count > 1024 || count > Math.floor(reader.remaining / 12)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid font alternate-name count");
    }
    for (let index = 0; index < count; index += 1) {
      const family = reader.string();
      const nameCount = reader.u32();
      if (!family || family.length > 128 || /[\u0000-\u001f\u007f]/u.test(family)
        || nameCount === 0 || nameCount > 64 || nameCount > Math.floor(reader.remaining / 5)) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid font alternate-name descriptor");
      }
      const names: string[] = [];
      for (let nameIndex = 0; nameIndex < nameCount; nameIndex += 1) {
        const name = reader.string();
        if (!name || name.length > 128 || /[\u0000-\u001f\u007f]/u.test(name)) {
          throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid font alternate name");
        }
        names.push(name);
      }
      fontAlternateNames.push({ family, names });
    }
  }
  const objectCount = reader.u32();
  const objects: SceneObject[] = [];
  for (let index = 0; index < objectCount; index += 1) {
    if (format === undefined) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Object exists without a document format");
    }
    const numericId = reader.u32();
    const parent = reader.i32();
    const unitIndex = reader.u32();
    const type = enumValue(OBJECT_TYPES, reader.u8(), "object type");
    const visualKind = reader.u8();
    const mapping = enumValue(MAPPINGS, reader.u8(), "mapping quality");
    reader.skip(1);
    const z = reader.i32();
    const bounds = { x: reader.f32(), y: reader.f32(), width: reader.f32(), height: reader.f32() };
    const id = reader.string();
    const parentId = reader.optionalString();
    const text = reader.optionalString();
    const source = readSource(reader, format, mapping);
    const visual = readVisual(reader, visualKind, version);
    objects.push({
      numericId,
      ...(parent < 0 ? {} : { parentNumericId: parent }),
      id,
      type,
      unitIndex,
      bounds,
      ...(parentId === undefined ? {} : { parentId }),
      ...(text === undefined ? {} : { text }),
      ...(source.format !== "pptx" || source.name === undefined ? {} : { name: source.name }),
      ...(source.format !== "pptx" || source.title === undefined ? {} : { title: source.title }),
      ...(source.format !== "pptx" || source.description === undefined
        ? {}
        : { description: source.description }),
      hidden: source.format === "pptx" && source.hidden === true,
      ...(!("actions" in source) || source.actions === undefined ? {} : { actions: source.actions }),
      source,
      z,
      visual,
    });
  }

  if (reader.remaining !== 0) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Core snapshot has trailing data");
  }
  if (!fatal && (format === undefined || kind === undefined)) {
    throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Successful snapshot lacks document identity");
  }
  validateScene(fatal, format, kind, units, outline, objects);

  return {
    fatal,
    ...(format === undefined || kind === undefined ? {} : { info: { format, kind, units, outline } }),
    diagnostics,
    embeddedFonts,
    ...(fontAlternateNames.length === 0 ? {} : { fontAlternateNames }),
    objects,
  };
}
