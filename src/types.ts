import type { FormatPackFactory } from "./format-pack.js";

export type DocumentFormat =
  | "pptx"
  | "odp"
  | "xlsx"
  | "ods"
  | "docx"
  | "odt"
  | "csv"
  | "rtf"
  | "ppt"
  | "xls"
  | "doc"
  | "pages"
  | "numbers"
  | "keynote"
  | "pdf"
  | "xps"
  | "ofd";
export type DocumentKind = "presentation" | "spreadsheet" | "text";
export type MappingQuality = "exact" | "derived" | "approximate";

/**
 * Document-space rectangle in CSS pixels at 96 DPI. The origin is the unit's
 * top-left corner; render scale and pixelRatio never change these coordinates.
 */
export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface SheetAxisSpan {
  readonly start: number;
  readonly end: number;
  readonly size: number;
}

export interface SheetAxis {
  readonly defaultSize: number;
  readonly spans: readonly SheetAxisSpan[];
}

export interface SheetMargins {
  /** Authored margin in inches. */
  readonly left?: number;
  readonly right?: number;
  readonly top?: number;
  readonly bottom?: number;
  readonly header?: number;
  readonly footer?: number;
}

export interface SheetPrintSettings {
  /** Authored worksheet view when it exposes printed-page boundaries. */
  readonly viewMode?: "pageLayout" | "pageBreakPreview";
  /** Authored Excel print-area formula, when present. */
  readonly printArea?: string;
  readonly printTitles?: string;
  /** Boundary id, first and last perpendicular cell indexes (inclusive). */
  readonly rowBreaks?: readonly (readonly [number, number, number])[];
  readonly columnBreaks?: readonly (readonly [number, number, number])[];
  /** ECMA-376 paper-size code; 9 is A4. */
  readonly paperSize?: number;
  readonly orientation?: "portrait" | "landscape";
  readonly scale?: number;
  readonly fitToWidth?: number;
  readonly fitToHeight?: number;
  readonly fitToPage: boolean;
  readonly margins: SheetMargins;
  readonly differentOddEven: boolean;
  readonly differentFirst: boolean;
  readonly oddHeader?: string;
  readonly oddFooter?: string;
  readonly evenHeader?: string;
  readonly evenFooter?: string;
  readonly firstHeader?: string;
  readonly firstFooter?: string;
}

export type SpeakerNoteTextAlign =
  | "start"
  | "center"
  | "end"
  | "justify"
  | "distribute"
  | "medium-kashida"
  | "high-kashida"
  | "low-kashida"
  | "thai-distribute";

export interface SpeakerNoteRun {
  /** False disables East Asian punctuation restrictions for this run. */
  readonly eastAsianLineBreaks?: boolean;
  readonly text: string;
  readonly fontFamily: string;
  readonly fontSize: number;
  /** RGBA packed as 0xRRGGBBAA. */
  readonly color: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly strikethrough: boolean;
  /** RGBA highlight; an alpha byte of zero means no highlight. */
  readonly highlight: number;
  readonly baselineShift: number;
  readonly letterSpacing: number;
  readonly horizontalScale: number;
}

export interface SpeakerNoteParagraph {
  readonly align: SpeakerNoteTextAlign;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly firstLineIndent: number;
  readonly defaultTabStop: number;
  readonly lineHeight: number;
  readonly spaceBefore: number;
  readonly spaceAfter: number;
  readonly latinLineBreak: boolean;
  readonly hangingPunctuation: boolean;
  readonly runs: readonly SpeakerNoteRun[];
}

/** Unit width and height use the same 96-DPI document-space coordinates as Rect. */
export type UnitDescriptor =
  | {
      readonly type: "slide";
      readonly index: number;
      readonly id: string;
      readonly name: string;
      readonly width: number;
      readonly height: number;
      readonly sourceId?: string;
      readonly sourcePart?: string;
      /** Plain-text compatibility view of the slide's notesSlide body. */
      readonly speakerNotes?: string;
      readonly speakerNotesPart?: string;
      /** Authored paragraph and run formatting from the slide's notesSlide body. */
      readonly speakerNoteParagraphs?: readonly SpeakerNoteParagraph[];
      readonly slideNumber: number;
      readonly hidden: boolean;
    }
  | {
      readonly type: "sheet";
      readonly index: number;
      readonly id: string;
      readonly name: string;
      readonly width: number;
      readonly height: number;
      readonly rows: number;
      readonly columns: number;
      readonly frozenRows: number;
      readonly frozenColumns: number;
      readonly frozenWidth: number;
      readonly frozenHeight: number;
      readonly rowAxis: SheetAxis;
      readonly columnAxis: SheetAxis;
      readonly showGridLines: boolean;
      /** Authored worksheet tab color, RGBA packed as 0xRRGGBBAA. */
      readonly tabColor?: number;
      readonly printSettings?: SheetPrintSettings;
    }
  | {
      readonly type: "page";
      readonly index: number;
      readonly id: string;
      readonly name: string;
      readonly width: number;
      readonly height: number;
    };

export type SourceRef =
  | {
      readonly format: "pptx";
      readonly part: string;
      readonly kind: "shape";
      readonly shapeId: number;
      readonly row?: number;
      readonly column?: number;
      readonly textRange?: readonly [start: number, end: number];
      readonly name?: string;
      readonly title?: string;
      readonly description?: string;
      readonly hidden?: boolean;
      readonly actions?: readonly DocumentAction[];
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "odp";
      readonly part: string;
      readonly kind: "element";
      readonly elementId?: string;
      readonly path: string;
      readonly row?: number;
      readonly column?: number;
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "xlsx";
      readonly part: string;
      readonly kind: "cell" | "drawing";
      readonly sheetName: string;
      readonly address?: string;
      readonly formula?: string;
      readonly drawingId?: number;
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "ods";
      readonly part: string;
      readonly kind: "cell" | "element";
      readonly tableName: string;
      readonly row?: number;
      readonly column?: number;
      readonly elementId?: string;
      readonly path: string;
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "docx";
      readonly part: string;
      readonly kind: "paragraph" | "drawing" | "table" | "table-cell";
      readonly paragraphId?: string;
      readonly paragraphIndex?: number;
      readonly drawingId?: number;
      readonly row?: number;
      readonly column?: number;
      readonly textRange?: readonly [start: number, end: number];
      readonly actions?: readonly DocumentAction[];
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "odt";
      readonly part: string;
      readonly kind: "element" | "text-range" | "table-cell";
      readonly elementId?: string;
      readonly path: string;
      readonly row?: number;
      readonly column?: number;
      readonly textRange?: readonly [start: number, end: number];
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "rtf";
      readonly part: string;
      readonly kind: "paragraph";
      readonly paragraphIndex: number;
      readonly textRange?: readonly [start: number, end: number];
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "rtf";
      readonly part: string;
      readonly kind: "picture";
      readonly pictureIndex: number;
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "csv";
      readonly part: string;
      readonly kind: "cell";
      readonly row: number;
      readonly column: number;
      readonly textRange?: readonly [start: number, end: number];
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "ppt";
      readonly part: string;
      readonly kind: "text" | "drawing";
      readonly stream: "PowerPoint Document";
      readonly recordOffset?: number;
      readonly textRange?: readonly [start: number, end: number];
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "xls";
      readonly part: string;
      readonly kind: "cell" | "drawing";
      readonly stream: "Workbook" | "Book";
      readonly recordOffset?: number;
      readonly row?: number;
      readonly column?: number;
      readonly textRange?: readonly [start: number, end: number];
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "doc";
      readonly part: string;
      readonly kind: "paragraph";
      readonly stream: "WordDocument";
      readonly recordOffset?: number;
      readonly textRange?: readonly [start: number, end: number];
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "pages";
      readonly part: string;
      readonly kind:
        | "preview"
        | "body"
        | "inline-image"
        | "shape"
        | "image"
        | "media"
        | "text-box"
        | "table"
        | "table-grid"
        | "table-cell"
        | "chart"
        | "chart-series";
      readonly component: string;
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "numbers";
      readonly part: string;
      readonly kind: "preview" | "cell" | "chart" | "chart-series";
      readonly component: string;
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "keynote";
      readonly part: string;
      readonly kind:
        | "preview"
        | "slide-preview"
        | "slide-background"
        | "shape"
        | "image"
        | "media"
        | "text-box"
        | "table"
        | "table-grid"
        | "table-cell"
        | "chart"
        | "chart-series";
      readonly component: string;
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "pdf";
      readonly part: string;
      readonly kind: "text" | "path" | "image" | "form" | "annotation" | "knockout-backdrop";
      readonly objectNumber?: number;
      readonly byteOffset?: number;
      readonly actions?: readonly DocumentAction[];
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "xps";
      readonly part: string;
      readonly kind: "canvas" | "glyphs" | "path";
      readonly path: string;
      readonly mapping: MappingQuality;
    }
  | {
      readonly format: "ofd";
      readonly part: string;
      readonly kind: "text" | "path" | "image" | "seal";
      readonly path: string;
      readonly mapping: MappingQuality;
    };

export type DocumentObjectType =
  | "group"
  | "text-box"
  | "paragraph"
  | "image"
  | "shape"
  | "table"
  | "cell"
  | "unknown";

export type DocumentActionTrigger = "click" | "hover";
export type DocumentActionKind = "hyperlink" | "slide" | "command" | "unknown";

export interface DocumentAction {
  readonly trigger: DocumentActionTrigger;
  readonly kind: DocumentActionKind;
  readonly action?: string;
  readonly target?: string;
  readonly tooltip?: string;
}

export type DocumentFontSource = "embedded" | "host" | "provider" | "browser" | "fallback";

/** Font resolution for a UTF-16 text range within DocumentObject.text. */
export interface DocumentFontRun {
  readonly start: number;
  readonly end: number;
  /** Font family declared by the document. */
  readonly authoredFamily: string;
  /** CSS font family used by the renderer after document, host, and browser resolution. */
  readonly renderedFamily: string;
  readonly source: DocumentFontSource;
  /** Authored CSS-pixel font size when exposed by the format adapter. */
  readonly fontSize?: number;
  /** Authored spreadsheet cell font weight when exposed by the format adapter. */
  readonly bold?: boolean;
  /** Authored spreadsheet cell font style when exposed by the format adapter. */
  readonly italic?: boolean;
}

export interface DocumentObject {
  readonly id: string;
  readonly type: DocumentObjectType;
  readonly unitIndex: number;
  readonly bounds: Rect;
  readonly parentId?: string;
  readonly text?: string;
  readonly fontRuns?: readonly DocumentFontRun[];
  /** Whether spreadsheet cell text uses authored wrapping. */
  readonly wrapText?: boolean;
  readonly name?: string;
  readonly title?: string;
  readonly description?: string;
  readonly hidden: boolean;
  readonly actions?: readonly DocumentAction[];
  readonly source: SourceRef;
}

export type DiagnosticSeverity = "info" | "warning" | "error" | "fatal";
export type DiagnosticFidelity = "exact" | "approximate" | "unsupported" | "not-rendered";
export type DiagnosticPhase = "identify" | "package" | "parse" | "layout" | "render" | "security";

export interface Diagnostic {
  readonly code: string;
  readonly severity: DiagnosticSeverity;
  readonly fidelity: DiagnosticFidelity;
  readonly phase: DiagnosticPhase;
  readonly message: string;
  readonly part?: string;
  readonly objectId?: string;
  readonly details?: Readonly<Record<string, string | number | boolean>>;
}

export interface DocumentOutlineItem {
  readonly title: string;
  readonly unitIndex: number;
  readonly level: number;
}

export interface DocumentInfo {
  readonly format: DocumentFormat;
  readonly kind: DocumentKind;
  readonly units: readonly UnitDescriptor[];
  readonly outline: readonly DocumentOutlineItem[];
}

/** A document-space crop rectangle using the same coordinates as object bounds. */
export interface Viewport extends Rect {}

export interface RenderRequest {
  readonly unitIndex: number;
  /** Document-space crop; output pixels are viewport dimensions × scale × pixelRatio. */
  readonly viewport?: Viewport;
  readonly scale?: number;
  readonly pixelRatio?: number;
  readonly background?: string;
  /** Repeated diagonal text painted over the rendered bitmap without modifying document objects. */
  readonly watermark?: string;
  /** Includes the exact text placements used by this render for selection and accessibility overlays. */
  readonly includeTextFragments?: boolean;
  /** View-only row and column sizes. The source document is never modified. */
  readonly sheetSizes?: {
    readonly rows?: readonly SheetSizeOverride[];
    readonly columns?: readonly SheetSizeOverride[];
  };
}

export type RenderPriority = "interactive" | "visible" | "prefetch";

export interface OperationOptions {
  readonly signal?: AbortSignal;
  /** Hard deadline in Worker mode. Expiry terminates the document Worker. */
  readonly timeoutMs?: number;
}

export interface RenderOptions extends OperationOptions {
  /** Higher-priority work may interrupt lower-priority rendering in Worker mode. */
  readonly priority?: RenderPriority;
  /** Only the newest queued render with the same key is retained. */
  readonly supersedeKey?: string;
}

/** @internal */
export function validateOperationOptions(options: OperationOptions): void {
  if (options === null || typeof options !== "object") {
    throw new OfficeEngineError("INVALID_OPERATION_OPTIONS", "Operation options must be an object");
  }
  if (options.timeoutMs !== undefined
    && (!Number.isFinite(options.timeoutMs) || options.timeoutMs <= 0)) {
    throw new OfficeEngineError("INVALID_TIMEOUT", "timeoutMs must be a finite positive number");
  }
}

/** @internal */
export function validateRenderOptions(options: RenderOptions): void {
  validateOperationOptions(options);
  if (options === null || typeof options !== "object") {
    throw new OfficeEngineError("INVALID_RENDER_OPTIONS", "Render options must be an object");
  }
  if (options.priority !== undefined
    && options.priority !== "interactive"
    && options.priority !== "visible"
    && options.priority !== "prefetch") {
    throw new OfficeEngineError("INVALID_RENDER_OPTIONS", "Render priority is invalid");
  }
  if (options.supersedeKey !== undefined
    && (typeof options.supersedeKey !== "string"
      || options.supersedeKey.length === 0
      || options.supersedeKey.length > 256)) {
    throw new OfficeEngineError(
      "INVALID_RENDER_OPTIONS",
      "supersedeKey must contain 1 through 256 characters",
    );
  }
}

export interface SheetSizeOverride {
  readonly index: number;
  readonly size: number;
}

export interface RenderResult {
  readonly bitmap: ImageBitmap;
  readonly viewport: Viewport;
  readonly pixelWidth: number;
  readonly pixelHeight: number;
  readonly renderedObjectCount: number;
  /** Locally embedded audio/video intersecting this render viewport. */
  readonly media: readonly RenderedMedia[];
  /** Final renderer-owned text placements, present only when requested. */
  readonly textFragments?: readonly RenderedTextFragment[];
  readonly diagnostics: readonly Diagnostic[];
}

/** One selectable text fragment positioned by the same layout pass that painted the bitmap. */
export interface RenderedTextFragment {
  readonly objectId: string;
  readonly text: string;
  /** UTF-16 range in the object's logical text when the renderer can preserve it exactly. */
  readonly start?: number;
  readonly end?: number;
  /** Fragment-local dimensions before the affine transform is applied. */
  readonly width: number;
  readonly height: number;
  readonly line: number;
  readonly font: string;
  readonly letterSpacing: number;
  readonly direction: "ltr" | "rtl";
  /** Maps fragment-local coordinates to unscaled document space. */
  readonly transform: Readonly<{ a: number; b: number; c: number; d: number; e: number; f: number }>;
}

export interface RenderedMedia {
  readonly objectId: string;
  readonly kind: "audio" | "video";
  readonly mediaType: string;
  readonly bytes: Uint8Array;
  readonly bounds: Rect;
  /** Authored document-space affine transform applied around the media object. */
  readonly transform: Readonly<{ a: number; b: number; c: number; d: number; e: number; f: number }>;
}

export interface HitTestRequest {
  readonly unitIndex: number;
  /** Unscaled document-space X coordinate in 96-DPI CSS pixels. */
  readonly x: number;
  /** Unscaled document-space Y coordinate in 96-DPI CSS pixels. */
  readonly y: number;
  readonly limit?: number;
}

export interface HitResult {
  readonly object: DocumentObject;
  readonly ancestors: readonly DocumentObject[];
}

export interface TextSearchRequest {
  readonly query: string;
  readonly unitIndex?: number;
  readonly matchCase?: boolean;
  readonly wholeWord?: boolean;
  /** Maximum matching objects returned; each object can contain multiple ranges. */
  readonly limit?: number;
}

export interface TextSearchResult {
  readonly object: DocumentObject;
  /** UTF-16 offsets into object.text, matching JavaScript string indexing. */
  readonly ranges: readonly (readonly [start: number, end: number])[];
}

export interface ObjectListRequest {
  readonly unitIndex?: number;
  readonly types?: readonly DocumentObjectType[];
  /** Return only objects with non-empty text, suitable for a selectable text layer. */
  readonly textOnly?: boolean;
  /** Return only objects whose document-space bounds overlap this viewport. */
  readonly viewport?: Viewport;
  /** Maximum objects returned. Defaults to 10,000. */
  readonly limit?: number;
}

export interface ResourceLimits {
  readonly inputBytes: number;
  readonly zipEntries: number;
  readonly inflatedBytes: number;
  readonly entryBytes: number;
  readonly compressionRatio: number;
  readonly xmlBytes: number;
  readonly xmlDepth: number;
  readonly xmlNodes: number;
  readonly relationshipEdges: number;
  readonly documentObjects: number;
  readonly renderPixels: number;
  /** Maximum decoded pixels in one embedded image. */
  readonly imagePixels: number;
  /** Maximum simultaneously retained or working decoded image pixels in one document. */
  readonly totalImagePixels: number;
  /** Maximum encoded bytes in host-provided or document-embedded font assets. */
  readonly fontBytes: number;
}

export type WasmSource = ArrayBuffer | Uint8Array | WebAssembly.Module | URL | string;

export type FontStyle = "normal" | "italic" | "oblique";
export type FontStretch =
  | "ultra-condensed"
  | "extra-condensed"
  | "condensed"
  | "semi-condensed"
  | "normal"
  | "semi-expanded"
  | "expanded"
  | "extra-expanded"
  | "ultra-expanded";

/** A local binary font supplied by the host. URL and CSS sources are not accepted. */
export interface FontAsset {
  readonly family: string;
  readonly bytes: ArrayBuffer | Uint8Array;
  /** Expected SHA-256 of bytes; recommended for manifest/provider assets. */
  readonly sha256?: string;
  readonly style?: FontStyle;
  /** CSS numeric weight, from 1 through 1000. */
  readonly weight?: number;
  readonly stretch?: FontStretch;
}

/** Controls whether managed fonts stay deterministic or exact requested local faces may be probed. */
export type FontPolicy = "deterministic" | "local-first";

/** One complete font-face demand collected from document text. */
export interface FontRequest {
  readonly family: string;
  readonly style: FontStyle;
  readonly weight: number;
  readonly stretch: FontStretch;
  /** Sorted Unicode scalar values are supplied to providers after normalization. */
  readonly codePoints: readonly number[];
  readonly language?: string;
  /** ISO 15924 script code, for example Latn, Hans, or Arab. */
  readonly script?: string;
  /** Optional lowercase SHA-256 hex digest required by deterministic deployments. */
  readonly expectedSha256?: string;
}

export interface FontProviderContext {
  readonly signal: AbortSignal;
  readonly policy: FontPolicy;
}

/** Host-owned lazy font source. Implementations return complete binary faces, never URLs. */
export type FontProvider = (
  requests: readonly FontRequest[],
  context: FontProviderContext,
) => Promise<readonly FontAsset[]> | readonly FontAsset[];

export type FontProviderDiagnosticCode =
  | "FONT_PROVIDER_FAILED"
  | "FONT_PROVIDER_TIMEOUT"
  | "FONT_PROVIDER_INVALID_RESPONSE"
  | "FONT_PROVIDER_MISSING"
  | "FONT_INTEGRITY_MISMATCH"
  | "FONT_PROVIDER_CACHE_LIMIT";

export interface FontProviderDiagnostic extends Omit<Diagnostic, "code"> {
  readonly code: FontProviderDiagnosticCode;
}

export interface FontProviderResult {
  readonly assets: readonly FontAsset[];
  readonly diagnostics: readonly FontProviderDiagnostic[];
}

export interface EngineOptions {
  readonly wasm?: WasmSource;
  /** Optional spreadsheet calculator. Loaded only when formula caches are missing. */
  readonly calculationWasm?: WasmSource | false;
  readonly workerUrl?: URL | string;
  readonly workerFactory?: (url: URL | string) => Worker;
  readonly execution?: "worker" | "inline";
  readonly limits?: Partial<ResourceLimits>;
  readonly fonts?: readonly FontAsset[];
  /** Defaults to local-first. Use deterministic for environment-independent rendering. */
  readonly fontPolicy?: FontPolicy;
  readonly fontProvider?: FontProvider;
  /** Per-provider-call deadline. Defaults to 10 seconds. */
  readonly fontProviderTimeoutMs?: number;
  /** Optional format packs. Omit or set false to keep them unsupported and unloaded. */
  readonly formatPack?: false | FormatPackFactory;
  /** @deprecated Ignored; Engine access requires no runtime license. */
  readonly license?: import("./license.js").LicenseOptions;
}

export interface OpenOptions {
  readonly signal?: AbortSignal;
  readonly timeoutMs?: number;
  /** Password for a PDF protected by the Standard Security Handler. It is never retained after open(). */
  readonly password?: string;
  /** Optional display name used only to isolate structurally recognized WPS-family inputs. */
  readonly fileName?: string;
  /**
   * Allows the engine to reuse the input; Worker mode transfers an ArrayBuffer.
   * The caller must not read or mutate the buffer after calling open().
   */
  readonly transferInput?: boolean;
}

export interface OfficeDocument {
  readonly info: DocumentInfo;
  diagnostics(): readonly Diagnostic[];
  render(request: RenderRequest, options?: RenderOptions): Promise<RenderResult>;
  hitTest(request: HitTestRequest, options?: OperationOptions): Promise<readonly HitResult[]>;
  searchText(request: TextSearchRequest, options?: OperationOptions): Promise<readonly TextSearchResult[]>;
  listObjects(request?: ObjectListRequest, options?: OperationOptions): Promise<readonly DocumentObject[]>;
  getObject(id: string, options?: OperationOptions): Promise<DocumentObject | undefined>;
  close(): void;
}

const DOCUMENT_OBJECT_TYPES = new Set<DocumentObjectType>([
  "group",
  "text-box",
  "paragraph",
  "image",
  "shape",
  "table",
  "cell",
  "unknown",
]);

/** @internal Shared by inline and Worker-backed documents; not re-exported by the package entrypoint. */
export function listDocumentObjects(
  objects: Iterable<DocumentObject>,
  units: readonly UnitDescriptor[],
  request: ObjectListRequest = {},
): readonly DocumentObject[] {
  if (request === null || typeof request !== "object") {
    throw new OfficeEngineError("INVALID_OBJECT_LIST", "Object list request must be an object");
  }
  if (request.unitIndex !== undefined) {
    const unit = units[request.unitIndex];
    if (!Number.isInteger(request.unitIndex)
      || request.unitIndex < 0
      || unit === undefined
      || unit.index !== request.unitIndex) {
      throw new OfficeEngineError("INVALID_UNIT", `Document has no unit at index ${request.unitIndex}`);
    }
  }
  if (request.textOnly !== undefined && typeof request.textOnly !== "boolean") {
    throw new OfficeEngineError("INVALID_OBJECT_LIST", "textOnly must be a boolean");
  }
  if (request.viewport !== undefined) {
    if (request.viewport === null || typeof request.viewport !== "object") {
      throw new OfficeEngineError("INVALID_OBJECT_LIST", "viewport must be an object");
    }
    const { x, y, width, height } = request.viewport;
    if (![x, y, width, height].every(Number.isFinite) || width <= 0 || height <= 0) {
      throw new OfficeEngineError("INVALID_OBJECT_LIST", "viewport must have finite coordinates and positive dimensions");
    }
  }
  const limit = request.limit ?? 10_000;
  if (!Number.isInteger(limit) || limit < 1 || limit > 100_000) {
    throw new OfficeEngineError("INVALID_OBJECT_LIST", "Object list limit must be an integer from 1 through 100000");
  }
  let types: ReadonlySet<DocumentObjectType> | undefined;
  if (request.types !== undefined) {
    if (!Array.isArray(request.types) || request.types.length === 0) {
      throw new OfficeEngineError("INVALID_OBJECT_LIST", "types must be a non-empty array");
    }
    const selected = new Set<DocumentObjectType>();
    for (const type of request.types) {
      if (!DOCUMENT_OBJECT_TYPES.has(type)) {
        throw new OfficeEngineError("INVALID_OBJECT_LIST", `Unsupported document object type ${String(type)}`);
      }
      selected.add(type);
    }
    types = selected;
  }
  const result: DocumentObject[] = [];
  for (const object of objects) {
    if (request.unitIndex !== undefined && object.unitIndex !== request.unitIndex) continue;
    if (types !== undefined && !types.has(object.type)) continue;
    if (request.textOnly === true && (object.text === undefined || object.text.length === 0)) continue;
    if (request.viewport !== undefined && !(
      object.bounds.x < request.viewport.x + request.viewport.width
      && object.bounds.x + object.bounds.width > request.viewport.x
      && object.bounds.y < request.viewport.y + request.viewport.height
      && object.bounds.y + object.bounds.height > request.viewport.y
    )) continue;
    result.push(object);
    if (result.length === limit) break;
  }
  return Object.freeze(result);
}

export interface OfficeEngine {
  open(bytes: ArrayBuffer | Uint8Array, options?: OpenOptions): Promise<OfficeDocument>;
  close(): void;
}

export interface OfficeEngineErrorOptions extends ErrorOptions {
  readonly diagnostics?: readonly Diagnostic[];
}

export class OfficeEngineError extends Error {
  readonly code: string;
  readonly diagnostics: readonly Diagnostic[];

  constructor(code: string, message: string, options: OfficeEngineErrorOptions = {}) {
    super(message, options);
    this.name = "OfficeEngineError";
    this.code = code;
    this.diagnostics = immutableDiagnostics(options.diagnostics ?? []);
  }
}

function immutableTextRange(value: readonly [number, number] | undefined) {
  return value === undefined
    ? {}
    : { textRange: Object.freeze([value[0], value[1]] as [number, number]) };
}

function immutableActions(actions: readonly DocumentAction[] | undefined) {
  return actions === undefined
    ? {}
    : { actions: Object.freeze(actions.map((action) => Object.freeze({ ...action }))) };
}

function immutableSource(source: SourceRef): SourceRef {
  switch (source.format) {
    case "pptx":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        shapeId: source.shapeId,
        ...(source.row === undefined ? {} : { row: source.row }),
        ...(source.column === undefined ? {} : { column: source.column }),
        ...immutableTextRange(source.textRange),
        ...(source.name === undefined ? {} : { name: source.name }),
        ...(source.title === undefined ? {} : { title: source.title }),
        ...(source.description === undefined ? {} : { description: source.description }),
        ...(source.hidden === undefined ? {} : { hidden: source.hidden }),
        ...immutableActions(source.actions),
        mapping: source.mapping,
      });
    case "odp":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        ...(source.elementId === undefined ? {} : { elementId: source.elementId }),
        path: source.path,
        ...(source.row === undefined ? {} : { row: source.row }),
        ...(source.column === undefined ? {} : { column: source.column }),
        mapping: source.mapping,
      });
    case "xlsx":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        sheetName: source.sheetName,
        ...(source.address === undefined ? {} : { address: source.address }),
        ...(source.formula === undefined ? {} : { formula: source.formula }),
        ...(source.drawingId === undefined ? {} : { drawingId: source.drawingId }),
        mapping: source.mapping,
      });
    case "ods":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        tableName: source.tableName,
        ...(source.row === undefined ? {} : { row: source.row }),
        ...(source.column === undefined ? {} : { column: source.column }),
        ...(source.elementId === undefined ? {} : { elementId: source.elementId }),
        path: source.path,
        mapping: source.mapping,
      });
    case "docx":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        ...(source.paragraphId === undefined ? {} : { paragraphId: source.paragraphId }),
        ...(source.paragraphIndex === undefined ? {} : { paragraphIndex: source.paragraphIndex }),
        ...(source.drawingId === undefined ? {} : { drawingId: source.drawingId }),
        ...(source.row === undefined ? {} : { row: source.row }),
        ...(source.column === undefined ? {} : { column: source.column }),
        ...immutableTextRange(source.textRange),
        ...immutableActions(source.actions),
        mapping: source.mapping,
      });
    case "odt":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        ...(source.elementId === undefined ? {} : { elementId: source.elementId }),
        path: source.path,
        ...(source.row === undefined ? {} : { row: source.row }),
        ...(source.column === undefined ? {} : { column: source.column }),
        ...immutableTextRange(source.textRange),
        mapping: source.mapping,
      });
    case "rtf":
      if (source.kind === "paragraph") {
        return Object.freeze({
          format: source.format,
          part: source.part,
          kind: source.kind,
          paragraphIndex: source.paragraphIndex,
          ...immutableTextRange(source.textRange),
          mapping: source.mapping,
        });
      }
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        pictureIndex: source.pictureIndex,
        mapping: source.mapping,
      });
    case "csv":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        row: source.row,
        column: source.column,
        ...immutableTextRange(source.textRange),
        mapping: source.mapping,
      });
    case "ppt":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        stream: source.stream,
        ...(source.recordOffset === undefined ? {} : { recordOffset: source.recordOffset }),
        ...immutableTextRange(source.textRange),
        mapping: source.mapping,
      });
    case "xls":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        stream: source.stream,
        ...(source.recordOffset === undefined ? {} : { recordOffset: source.recordOffset }),
        ...(source.row === undefined ? {} : { row: source.row }),
        ...(source.column === undefined ? {} : { column: source.column }),
        ...immutableTextRange(source.textRange),
        mapping: source.mapping,
      });
    case "doc":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        stream: source.stream,
        ...(source.recordOffset === undefined ? {} : { recordOffset: source.recordOffset }),
        ...immutableTextRange(source.textRange),
        mapping: source.mapping,
      });
    case "pages":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        component: source.component,
        mapping: source.mapping,
      });
    case "numbers":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        component: source.component,
        mapping: source.mapping,
      });
    case "keynote":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        component: source.component,
        mapping: source.mapping,
      });
    case "pdf":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        ...(source.objectNumber === undefined ? {} : { objectNumber: source.objectNumber }),
        ...(source.byteOffset === undefined ? {} : { byteOffset: source.byteOffset }),
        ...immutableActions(source.actions),
        mapping: source.mapping,
      });
    case "xps":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        path: source.path,
        mapping: source.mapping,
      });
    case "ofd":
      return Object.freeze({
        format: source.format,
        part: source.part,
        kind: source.kind,
        path: source.path,
        mapping: source.mapping,
      });
  }
  throw new OfficeEngineError("INVALID_SOURCE_REF", "Document object has an unsupported source format");
}

function immutableUnit(unit: UnitDescriptor): UnitDescriptor {
  if (unit.type === "sheet") {
    return Object.freeze({
      type: unit.type,
      index: unit.index,
      id: unit.id,
      name: unit.name,
      width: unit.width,
      height: unit.height,
      rows: unit.rows,
      columns: unit.columns,
      frozenRows: unit.frozenRows,
      frozenColumns: unit.frozenColumns,
      frozenWidth: unit.frozenWidth,
      frozenHeight: unit.frozenHeight,
      rowAxis: Object.freeze({
        defaultSize: unit.rowAxis.defaultSize,
        spans: Object.freeze(unit.rowAxis.spans.map((span) => Object.freeze({ ...span }))),
      }),
      columnAxis: Object.freeze({
        defaultSize: unit.columnAxis.defaultSize,
        spans: Object.freeze(unit.columnAxis.spans.map((span) => Object.freeze({ ...span }))),
      }),
      showGridLines: unit.showGridLines,
      ...(unit.tabColor === undefined ? {} : { tabColor: unit.tabColor }),
      ...(unit.printSettings === undefined ? {} : {
        printSettings: Object.freeze({
          ...unit.printSettings,
          margins: Object.freeze({ ...unit.printSettings.margins }),
          ...(unit.printSettings.rowBreaks === undefined ? {} : {
            rowBreaks: Object.freeze(unit.printSettings.rowBreaks.map(value => Object.freeze([...value] as const))),
          }),
          ...(unit.printSettings.columnBreaks === undefined ? {} : {
            columnBreaks: Object.freeze(unit.printSettings.columnBreaks.map(value => Object.freeze([...value] as const))),
          }),
        }),
      }),
    });
  }
  if (unit.type === "slide") {
    return Object.freeze({
      type: unit.type,
      index: unit.index,
      id: unit.id,
      name: unit.name,
      width: unit.width,
      height: unit.height,
      ...(unit.sourceId === undefined ? {} : { sourceId: unit.sourceId }),
      ...(unit.sourcePart === undefined ? {} : { sourcePart: unit.sourcePart }),
      ...(unit.speakerNotes === undefined ? {} : { speakerNotes: unit.speakerNotes }),
      ...(unit.speakerNotesPart === undefined ? {} : { speakerNotesPart: unit.speakerNotesPart }),
      ...(unit.speakerNoteParagraphs === undefined ? {} : {
        speakerNoteParagraphs: Object.freeze(unit.speakerNoteParagraphs.map((paragraph) => Object.freeze({
          ...paragraph,
          runs: Object.freeze(paragraph.runs.map((run) => Object.freeze({ ...run }))),
        }))),
      }),
      slideNumber: unit.slideNumber,
      hidden: unit.hidden,
    });
  }
  return Object.freeze({
    type: unit.type,
    index: unit.index,
    id: unit.id,
    name: unit.name,
    width: unit.width,
    height: unit.height,
  });
}

export function immutableDocumentInfo(info: DocumentInfo): DocumentInfo {
  const units = Object.freeze(info.units.map(immutableUnit));
  const outline = Object.freeze((info.outline ?? []).map((item) => Object.freeze({ ...item })));
  return Object.freeze({ format: info.format, kind: info.kind, units, outline });
}

export function immutableDocumentObject(object: DocumentObject): DocumentObject {
  return Object.freeze({
    id: object.id,
    type: object.type,
    unitIndex: object.unitIndex,
    bounds: Object.freeze({
      x: object.bounds.x,
      y: object.bounds.y,
      width: object.bounds.width,
      height: object.bounds.height,
    }),
    ...(object.parentId === undefined ? {} : { parentId: object.parentId }),
    ...(object.text === undefined ? {} : { text: object.text }),
    ...(object.fontRuns === undefined
      ? {}
      : { fontRuns: Object.freeze(object.fontRuns.map((run) => Object.freeze({ ...run }))) }),
    ...(object.wrapText === undefined ? {} : { wrapText: object.wrapText }),
    ...(object.name === undefined ? {} : { name: object.name }),
    ...(object.title === undefined ? {} : { title: object.title }),
    ...(object.description === undefined ? {} : { description: object.description }),
    hidden: object.hidden,
    ...immutableActions(object.actions),
    source: immutableSource(object.source),
  });
}

export function immutableDiagnostics(diagnostics: readonly Diagnostic[]): readonly Diagnostic[] {
  return Object.freeze(diagnostics.map((diagnostic) => Object.freeze({
    code: diagnostic.code,
    severity: diagnostic.severity,
    fidelity: diagnostic.fidelity,
    phase: diagnostic.phase,
    message: diagnostic.message,
    ...(diagnostic.part === undefined ? {} : { part: diagnostic.part }),
    ...(diagnostic.objectId === undefined ? {} : { objectId: diagnostic.objectId }),
    ...(diagnostic.details === undefined
      ? {}
      : { details: Object.freeze({ ...diagnostic.details }) }),
  })));
}

export function immutableHitResults(results: readonly HitResult[]): readonly HitResult[] {
  return Object.freeze(results.map((result) => Object.freeze({
    object: result.object,
    ancestors: Object.freeze([...result.ancestors]),
  })));
}

const WORD_CHARACTER = /[\p{L}\p{N}_]/u;

function adjacentCodePoint(text: string, offset: number, direction: -1 | 1): string | undefined {
  if (direction < 0) return Array.from(text.slice(0, offset)).at(-1);
  return Array.from(text.slice(offset))[0];
}

/** @internal Shared by inline and Worker-backed documents; not re-exported by the package entrypoint. */
export function searchDocumentText(
  objects: Iterable<DocumentObject>,
  units: readonly UnitDescriptor[],
  request: TextSearchRequest,
): readonly TextSearchResult[] {
  if (request === null || typeof request !== "object") {
    throw new OfficeEngineError("INVALID_TEXT_SEARCH", "Text search request must be an object");
  }
  if (typeof request.query !== "string" || request.query.length === 0 || request.query.length > 1_024) {
    throw new OfficeEngineError("INVALID_TEXT_SEARCH", "Text search query must contain 1 through 1024 characters");
  }
  if (request.matchCase !== undefined && typeof request.matchCase !== "boolean") {
    throw new OfficeEngineError("INVALID_TEXT_SEARCH", "matchCase must be a boolean");
  }
  if (request.wholeWord !== undefined && typeof request.wholeWord !== "boolean") {
    throw new OfficeEngineError("INVALID_TEXT_SEARCH", "wholeWord must be a boolean");
  }
  if (request.unitIndex !== undefined) {
    const unit = units[request.unitIndex];
    if (!Number.isInteger(request.unitIndex)
      || request.unitIndex < 0
      || unit === undefined
      || unit.index !== request.unitIndex) {
      throw new OfficeEngineError("INVALID_UNIT", `Document has no unit at index ${request.unitIndex}`);
    }
  }
  const limit = request.limit ?? 100;
  if (!Number.isInteger(limit) || limit < 1 || limit > 10_000) {
    throw new OfficeEngineError("INVALID_TEXT_SEARCH", "Text search limit must be an integer from 1 through 10000");
  }
  const escaped = request.query.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
  const expression = new RegExp(escaped, request.matchCase === true ? "gu" : "giu");
  const results: TextSearchResult[] = [];
  for (const object of objects) {
    if (request.unitIndex !== undefined && object.unitIndex !== request.unitIndex) continue;
    const text = object.text;
    if (text === undefined || text === "") continue;
    const ranges: (readonly [number, number])[] = [];
    for (const match of text.matchAll(expression)) {
      const start = match.index;
      const end = start + match[0].length;
      if (request.wholeWord === true) {
        const previous = adjacentCodePoint(text, start, -1);
        const next = adjacentCodePoint(text, end, 1);
        if ((previous !== undefined && WORD_CHARACTER.test(previous))
          || (next !== undefined && WORD_CHARACTER.test(next))) continue;
      }
      ranges.push(Object.freeze([start, end] as [number, number]));
    }
    if (ranges.length === 0) continue;
    results.push(Object.freeze({ object, ranges: Object.freeze(ranges) }));
    if (results.length === limit) break;
  }
  return Object.freeze(results);
}

export function immutableRenderResult(result: RenderResult): RenderResult {
  return Object.freeze({
    bitmap: result.bitmap,
    viewport: Object.freeze({
      x: result.viewport.x,
      y: result.viewport.y,
      width: result.viewport.width,
      height: result.viewport.height,
    }),
    pixelWidth: result.pixelWidth,
    pixelHeight: result.pixelHeight,
    renderedObjectCount: result.renderedObjectCount,
    media: Object.freeze(result.media.map((media) => Object.freeze({
      objectId: media.objectId,
      kind: media.kind,
      mediaType: media.mediaType,
      bytes: media.bytes.slice(),
      bounds: Object.freeze({ ...media.bounds }),
      transform: Object.freeze({ ...media.transform }),
    }))),
    ...(result.textFragments === undefined ? {} : {
      textFragments: Object.freeze(result.textFragments.map((fragment) => Object.freeze({
        ...fragment,
        transform: Object.freeze({ ...fragment.transform }),
      }))),
    }),
    diagnostics: immutableDiagnostics(result.diagnostics),
  });
}
