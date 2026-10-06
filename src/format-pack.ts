import type {
  OfficeDocument,
  OfficeEngine,
  OpenOptions,
  WasmSource,
} from "./types.js";
import { OfficeEngineError } from "./types.js";
import { copyByteSource, isArrayBuffer, isUint8Array } from "./bytes.js";

/** Closed optional format families which are intentionally kept out of the base runtime. */
export type FormatPackCandidate = "odf" | "iwork" | "legacy-office" | "wps" | "pdf" | "xps" | "ofd";

/** A lazily imported pack resolves a native Wasm module or a browser-native format engine. */
export interface FormatPack {
  load(candidate: FormatPackCandidate): Promise<WasmSource | OfficeEngine>;
}

export type FormatPackFactory = () => Promise<FormatPack>;

/**
 * Integration seam for engine.ts. The factory must preserve the parent Engine's
 * execution, Worker, resource-limit, font, and provider options, replace only
 * `wasm`, and force `formatPack: false` to prevent recursive fallback.
 */
export type CreateFormatPackEngine = (
  candidate: FormatPackCandidate,
  wasm: WasmSource,
) => Promise<OfficeEngine>;

const ZIP_LOCAL_FILE = 0x04034b50;
const ZIP_CENTRAL_FILE = 0x02014b50;
const ZIP_END = 0x06054b50;
const ZIP_END_BYTES = 22;
const ZIP_MAX_COMMENT_BYTES = 0xffff;
const ZIP_MAX_ENTRIES = 50_000;
const DEFAULT_OPEN_TIMEOUT_MS = 60_000;
const DEFAULT_INPUT_BYTES = 128 * 1024 * 1024;
const OLE_MAGIC = Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1);
const PDF_HEADER_SCAN_BYTES = 1_024;
const PDF_RECOVERY_SCAN_BYTES = 16 * 1024 * 1024;
const PDF_BINARY_MARKER = Uint8Array.of(0x25, 0xe2, 0xe3, 0xcf, 0xd3);
const FLAT_ODF_SCAN_BYTES = 64 * 1024;
const FLAT_PRESENTATION_MIMETYPE = "application/vnd.oasis.opendocument.presentation";
const FLAT_GRAPHICS_MIMETYPE = "application/vnd.oasis.opendocument.graphics";
const FLAT_SPREADSHEET_MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet";
const FLAT_TEXT_MIMETYPE = "application/vnd.oasis.opendocument.text";
const WPS_FILE_NAME = /(?:^|[/\\])[^/\\]+\.(?:wps|et|dps)$/iu;

function sourceBytes(source: ArrayBuffer | Uint8Array): Uint8Array {
  if (isArrayBuffer(source)) return new Uint8Array(source);
  if (isUint8Array(source)) {
    return new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  }
  throw new OfficeEngineError("INVALID_INPUT", "Document bytes must be an ArrayBuffer or Uint8Array");
}

function hasBytesAt(bytes: Uint8Array, offset: number, expected: Uint8Array): boolean {
  if (offset < 0 || offset + expected.byteLength > bytes.byteLength) return false;
  for (let index = 0; index < expected.byteLength; index += 1) {
    if (bytes[offset + index] !== expected[index]) return false;
  }
  return true;
}

function hasAsciiAt(bytes: Uint8Array, offset: number, text: string): boolean {
  if (offset < 0 || offset + text.length > bytes.byteLength) return false;
  for (let index = 0; index < text.length; index += 1) {
    if (bytes[offset + index] !== text.charCodeAt(index)) return false;
  }
  return true;
}

function hasAsciiWithin(bytes: Uint8Array, text: string, limit: number): boolean {
  const end = Math.min(bytes.byteLength, limit);
  for (let offset = 0; offset + text.length <= end; offset += 1) {
    if (hasAsciiAt(bytes, offset, text)) return true;
  }
  return false;
}

function looksLikeFlatOdf(bytes: Uint8Array): boolean {
  let offset = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  while (offset < bytes.byteLength
    && (bytes[offset] === 0x09
      || bytes[offset] === 0x0a
      || bytes[offset] === 0x0d
      || bytes[offset] === 0x20)) {
    offset += 1;
  }
  return bytes[offset] === 0x3c
    && (hasAsciiWithin(bytes, FLAT_PRESENTATION_MIMETYPE, FLAT_ODF_SCAN_BYTES)
      || hasAsciiWithin(bytes, FLAT_GRAPHICS_MIMETYPE, FLAT_ODF_SCAN_BYTES)
      || hasAsciiWithin(bytes, FLAT_SPREADSHEET_MIMETYPE, FLAT_ODF_SCAN_BYTES)
      || hasAsciiWithin(bytes, FLAT_TEXT_MIMETYPE, FLAT_ODF_SCAN_BYTES));
}

function isAsciiName(bytes: Uint8Array, offset: number, length: number, name: string): boolean {
  return length === name.length && hasAsciiAt(bytes, offset, name);
}

function isIwaName(bytes: Uint8Array, offset: number, length: number): boolean {
  return length > "Index/.iwa".length
    && hasAsciiAt(bytes, offset, "Index/")
    && hasAsciiAt(bytes, offset + length - 4, ".iwa");
}

function hasNestedAsciiSuffix(
  bytes: Uint8Array,
  offset: number,
  length: number,
  suffix: string,
): boolean {
  return length > suffix.length && hasAsciiAt(bytes, offset + length - suffix.length, suffix);
}

function hasAsciiSuffix(bytes: Uint8Array, offset: number, length: number, suffix: string): boolean {
  return length >= suffix.length && hasAsciiAt(bytes, offset + length - suffix.length, suffix);
}

function findZipEnd(view: DataView): number | undefined {
  const first = Math.max(0, view.byteLength - ZIP_END_BYTES - ZIP_MAX_COMMENT_BYTES);
  for (let offset = view.byteLength - ZIP_END_BYTES; offset >= first; offset -= 1) {
    if (view.getUint32(offset, true) !== ZIP_END) continue;
    const commentBytes = view.getUint16(offset + 20, true);
    if (offset + ZIP_END_BYTES + commentBytes === view.byteLength) return offset;
  }
  return undefined;
}

/** Reads central-directory names only; it never inflates or transforms document data. */
function detectZipFormatPackCandidate(bytes: Uint8Array): FormatPackCandidate | undefined {
  if (bytes.byteLength < ZIP_END_BYTES + 4) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  // Foxit OFD packages can put an end record and directory before local files.
  const signature = view.getUint32(0, true);
  if (signature !== ZIP_LOCAL_FILE && signature !== ZIP_END) return undefined;
  const endOffset = findZipEnd(view);
  if (endOffset === undefined) return undefined;

  const disk = view.getUint16(endOffset + 4, true);
  const centralDisk = view.getUint16(endOffset + 6, true);
  const entriesOnDisk = view.getUint16(endOffset + 8, true);
  const entryCount = view.getUint16(endOffset + 10, true);
  const centralBytes = view.getUint32(endOffset + 12, true);
  const centralOffset = view.getUint32(endOffset + 16, true);
  if (disk !== 0
    || centralDisk !== 0
    || entriesOnDisk !== entryCount
    || entryCount === 0xffff
    || entryCount > ZIP_MAX_ENTRIES
    || centralOffset + centralBytes > endOffset) {
    return undefined;
  }

  let offset = centralOffset;
  let hasOdfMimetype = false;
  let hasOdfContent = false;
  let hasOdfManifest = false;
  let hasIndex = false;
  let hasMetadata = false;
  let hasDirectoryIndex = false;
  let hasDirectoryMetadata = false;
  let hasXpsSequence = false;
  let hasOpcTypes = false;
  let hasOpcRelationships = false;
  let hasXpsDocument = false;
  let hasXpsPage = false;
  let hasOfdRoot = false;
  let hasOfdDocument = false;
  let hasOfdPage = false;
  for (let entry = 0; entry < entryCount; entry += 1) {
    if (offset + 46 > endOffset || view.getUint32(offset, true) !== ZIP_CENTRAL_FILE) return undefined;
    const nameBytes = view.getUint16(offset + 28, true);
    const extraBytes = view.getUint16(offset + 30, true);
    const commentBytes = view.getUint16(offset + 32, true);
    const nameOffset = offset + 46;
    const next = nameOffset + nameBytes + extraBytes + commentBytes;
    if (next > endOffset) return undefined;
    hasOdfMimetype ||= isAsciiName(bytes, nameOffset, nameBytes, "mimetype");
    hasOdfContent ||= isAsciiName(bytes, nameOffset, nameBytes, "content.xml");
    hasOdfManifest ||= isAsciiName(bytes, nameOffset, nameBytes, "META-INF/manifest.xml");
    hasIndex ||= isIwaName(bytes, nameOffset, nameBytes)
      || isAsciiName(bytes, nameOffset, nameBytes, "Index.zip");
    hasMetadata ||= isAsciiName(bytes, nameOffset, nameBytes, "Metadata/Properties.plist");
    hasDirectoryIndex ||= hasNestedAsciiSuffix(
      bytes,
      nameOffset,
      nameBytes,
      "/Index/Document.iwa",
    ) || hasNestedAsciiSuffix(bytes, nameOffset, nameBytes, "/Index.zip");
    hasDirectoryMetadata ||= hasNestedAsciiSuffix(
      bytes,
      nameOffset,
      nameBytes,
      "/Metadata/Properties.plist",
    );
    // OPC logical part names may end in a numbered physical piece.
    let partBytes = nameBytes;
    if (hasAsciiSuffix(bytes, nameOffset, nameBytes, ".piece")) {
      for (let i = nameBytes - 1; i > 0; i -= 1) {
        if (bytes[nameOffset + i] === 0x2f) { partBytes = i; break; }
      }
    }
    hasOpcTypes ||= isAsciiName(bytes, nameOffset, nameBytes, "[Content_Types].xml");
    hasOpcRelationships ||= isAsciiName(bytes, nameOffset, nameBytes, "_rels/.rels");
    hasXpsSequence ||= hasAsciiSuffix(bytes, nameOffset, partBytes, ".fdseq")
      || hasAsciiSuffix(bytes, nameOffset, partBytes, "FixedDocSeq.xaml");
    hasXpsDocument ||= hasAsciiSuffix(bytes, nameOffset, partBytes, ".fdoc")
      || hasAsciiSuffix(bytes, nameOffset, partBytes, "FixedDoc.xaml");
    hasXpsPage ||= hasAsciiSuffix(bytes, nameOffset, partBytes, ".fpage")
      || hasAsciiSuffix(bytes, nameOffset, partBytes, ".xaml");
    hasOfdRoot ||= isAsciiName(bytes, nameOffset, nameBytes, "OFD.xml");
    hasOfdDocument ||= hasNestedAsciiSuffix(bytes, nameOffset, nameBytes, "/Document.xml");
    hasOfdPage ||= hasNestedAsciiSuffix(bytes, nameOffset, nameBytes, "/Content.xml");
    if (hasOdfMimetype && hasOdfContent && hasOdfManifest) return "odf";
    if ((hasIndex && hasMetadata) || (hasDirectoryIndex && hasDirectoryMetadata)) return "iwork";
    if (hasXpsPage && ((hasXpsSequence && hasXpsDocument)
      || (hasOpcTypes && hasOpcRelationships && (hasXpsSequence || hasXpsDocument)))) return "xps";
    if (hasOfdRoot && hasOfdDocument && hasOfdPage) return "ofd";
    offset = next;
  }
  return undefined;
}

/**
 * Cheap, bounded pre-dispatch used to preserve candidate bytes before the base
 * engine runs. It never authorizes fallback by itself: the base error is still
 * checked against a narrow allowlist. File names, MIME types, and extensions
 * are never trusted.
 */
export function detectFormatPackCandidate(
  source: ArrayBuffer | Uint8Array,
  fileName?: string,
): FormatPackCandidate | undefined {
  const bytes = sourceBytes(source);
  if (hasBytesAt(bytes, 0, OLE_MAGIC)) {
    // WPS, ET, and DPS reuse the DOC, BIFF, and PowerPoint CFB record families.
    // The extension only chooses between two equally bounded CFB adapters; the
    // selected native parser must still validate the matching main stream and
    // its record-level signature before accepting the document.
    if (typeof fileName === "string"
      && fileName.length <= 1_024
      && WPS_FILE_NAME.test(fileName.trim())) {
      return "wps";
    }
    return "legacy-office";
  }
  const headerLimit = Math.min(bytes.byteLength, PDF_HEADER_SCAN_BYTES);
  for (let offset = 0; offset + 5 <= headerLimit; offset += 1) {
    if (hasAsciiAt(bytes, offset, "%PDF-")) return "pdf";
  }
  if (hasBytesAt(bytes, 0, PDF_BINARY_MARKER)
    && hasAsciiWithin(bytes, " obj", PDF_RECOVERY_SCAN_BYTES)
    && hasAsciiWithin(bytes, "trailer", PDF_RECOVERY_SCAN_BYTES)
    && hasAsciiWithin(bytes, "/Root", PDF_RECOVERY_SCAN_BYTES)) {
    return "pdf";
  }
  const packaged = detectZipFormatPackCandidate(bytes);
  if (packaged !== undefined) return packaged;
  if (looksLikeFlatOdf(bytes)) return "odf";
  return undefined;
}

function isOfficeEngine(value: WasmSource | OfficeEngine): value is OfficeEngine {
  return value !== null
    && typeof value === "object"
    && "open" in value
    && typeof value.open === "function"
    && "close" in value
    && typeof value.close === "function";
}

function isNativeFallback(cause: unknown, candidate: FormatPackCandidate): boolean {
  if (cause === null
    || typeof cause !== "object"
    || !("code" in cause)
    || typeof cause.code !== "string") {
    return false;
  }
  if (cause.code === "UNSUPPORTED_FORMAT") return true;
  if (candidate !== "iwork") return false;
  if (cause.code === "PACKAGE_ZIP64_UNSUPPORTED") return true;
  return cause.code === "PACKAGE_ZIP_INVALID"
    && "message" in cause
    && cause.message === "non-ASCII ZIP path requires the UTF-8 flag";
}

function now(): number {
  return globalThis.performance?.now() ?? Date.now();
}

function aborted(signal: AbortSignal): OfficeEngineError {
  return new OfficeEngineError("OPERATION_ABORTED", "The operation was aborted", { cause: signal.reason });
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted === true) throw aborted(signal);
}

function timedOut(timeoutMs: number): OfficeEngineError {
  return new OfficeEngineError("OPERATION_TIMEOUT", `Operation exceeded ${timeoutMs}ms`);
}

function closedError(): OfficeEngineError {
  return new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
}

function waitForControlled<T>(
  promise: Promise<T>,
  signal: AbortSignal | undefined,
  lifecycle: AbortSignal,
  deadline: number | undefined,
  timeoutMs: number | undefined,
): Promise<T> {
  if (signal?.aborted === true) return Promise.reject(aborted(signal));
  if (lifecycle.aborted) return Promise.reject(closedError());
  const delay = deadline === undefined ? undefined : deadline - now();
  if (delay !== undefined && delay <= 0) return Promise.reject(timedOut(timeoutMs!));

  return new Promise<T>((resolve, reject) => {
    let settled = false;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const cleanup = () => {
      if (timer !== undefined) clearTimeout(timer);
      signal?.removeEventListener("abort", onAbort);
      lifecycle.removeEventListener("abort", onClose);
    };
    const succeed = (value: T) => {
      if (settled) return;
      if (deadline !== undefined && now() >= deadline) {
        settled = true;
        cleanup();
        reject(timedOut(timeoutMs!));
        return;
      }
      settled = true;
      cleanup();
      resolve(value);
    };
    const fail = (cause: unknown) => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(cause);
    };
    const onAbort = () => fail(aborted(signal!));
    const onClose = () => fail(closedError());
    signal?.addEventListener("abort", onAbort, { once: true });
    lifecycle.addEventListener("abort", onClose, { once: true });
    if (delay !== undefined) timer = setTimeout(() => fail(timedOut(timeoutMs!)), delay);
    promise.then(succeed, fail);
  });
}

/**
 * Native-first optional fallback. For a recognized pack candidate, this wrapper
 * gives both engines the same byte-for-byte input snapshot and returns the pack
 * document unchanged. It never produces or consumes an intermediate document.
 *
 * Integration requirements:
 * - types.ts adds `formatPack?: false | FormatPackFactory` to EngineOptions;
 * - engine.ts constructs this wrapper only around the already-created base Engine;
 * - index.ts exports the types/runtime used by the public SDK;
 * - package.json exports `./extended-formats`, and the build publishes both pack Wasm assets.
 */
export class FormatPackRuntime implements OfficeEngine {
  readonly #nativeEngine: OfficeEngine;
  #formatPackFactory: FormatPackFactory | undefined;
  #createPackEngine: CreateFormatPackEngine | undefined;
  readonly #maxInputBytes: number;
  readonly #execution: "worker" | "inline";
  readonly #enginePromises = new Map<FormatPackCandidate, Promise<OfficeEngine>>();
  readonly #packEngines = new Set<OfficeEngine>();
  readonly #lifecycle = new AbortController();
  #formatPackPromise: Promise<FormatPack> | undefined;
  #closed = false;

  constructor(
    nativeEngine: OfficeEngine,
    formatPack: false | FormatPackFactory | undefined,
    createPackEngine: CreateFormatPackEngine,
    maxInputBytes = DEFAULT_INPUT_BYTES,
    execution: "worker" | "inline" = "worker",
  ) {
    this.#nativeEngine = nativeEngine;
    this.#formatPackFactory = formatPack === false ? undefined : formatPack;
    this.#createPackEngine = createPackEngine;
    this.#maxInputBytes = maxInputBytes;
    this.#execution = execution;
  }

  async open(
    bytes: ArrayBuffer | Uint8Array,
    options: OpenOptions = {},
  ): Promise<OfficeDocument> {
    this.#assertOpen();
    const started = now();
    const candidate = this.#formatPackFactory !== undefined && bytes.byteLength <= this.#maxInputBytes
      ? detectFormatPackCandidate(bytes, options.fileName)
      : undefined;
    const stableBytes = candidate === undefined
      ? bytes
      : new Uint8Array(copyByteSource(bytes));
    if (candidate === "odf" && looksLikeFlatOdf(sourceBytes(stableBytes))) {
      return this.#openPack(candidate, stableBytes, options, started);
    }

    let nativeFailure: unknown;
    try {
      return await this.#nativeEngine.open(stableBytes, options);
    } catch (cause) {
      nativeFailure = cause;
    }

    if (candidate === undefined || !isNativeFallback(nativeFailure, candidate)) {
      throw nativeFailure;
    }
    throwIfAborted(options.signal);
    this.#assertOpen();

    return this.#openPack(candidate, stableBytes, options, started);
  }

  async #openPack(
    candidate: FormatPackCandidate,
    stableBytes: ArrayBuffer | Uint8Array,
    options: OpenOptions,
    started: number,
  ): Promise<OfficeDocument> {
    const timeoutMs = options.timeoutMs ?? DEFAULT_OPEN_TIMEOUT_MS;
    const deadline = started + timeoutMs;
    const packEngine = await waitForControlled(
      this.#engine(candidate),
      options.signal,
      this.#lifecycle.signal,
      deadline,
      timeoutMs,
    );
    this.#assertOpen();
    throwIfAborted(options.signal);

    const remaining = deadline - now();
    if (remaining <= 0) throw timedOut(timeoutMs);
    const packOptions = options.timeoutMs === undefined && this.#execution === "inline"
      ? options
      : { ...options, timeoutMs: remaining };
    const opened = packEngine.open(stableBytes, packOptions);
    void opened.then((document) => {
      if (this.#closed
        || options.signal?.aborted === true
        || now() >= deadline) {
        document.close();
      }
    }, () => {});
    return waitForControlled(
      opened,
      options.signal,
      this.#lifecycle.signal,
      deadline,
      timeoutMs,
    );
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    this.#lifecycle.abort(closedError());
    this.#formatPackFactory = undefined;
    this.#createPackEngine = undefined;
    this.#enginePromises.clear();
    this.#formatPackPromise = undefined;
    let failure: unknown;
    for (const engine of [this.#nativeEngine, ...this.#packEngines]) {
      try {
        engine.close();
      } catch (cause) {
        failure ??= cause;
      }
    }
    this.#packEngines.clear();
    if (failure !== undefined) throw failure;
  }

  #engine(candidate: FormatPackCandidate): Promise<OfficeEngine> {
    const cached = this.#enginePromises.get(candidate);
    if (cached !== undefined) return cached;
    const pending = this.#create(candidate);
    this.#enginePromises.set(candidate, pending);
    void pending.catch(() => {
      if (this.#enginePromises.get(candidate) === pending) {
        this.#enginePromises.delete(candidate);
      }
    });
    return pending;
  }

  async #create(candidate: FormatPackCandidate): Promise<OfficeEngine> {
    const pack = await this.#pack();
    if (this.#closed) throw closedError();
    if (pack === null || typeof pack !== "object" || typeof pack.load !== "function") {
      throw new OfficeEngineError("INVALID_FORMAT_PACK", "The format pack factory returned an invalid pack");
    }
    const payload = await pack.load(candidate);
    const createPackEngine = this.#createPackEngine;
    if (this.#closed || createPackEngine === undefined) throw closedError();
    const engine = isOfficeEngine(payload)
      ? payload
      : await createPackEngine(candidate, payload);
    if (this.#closed) {
      engine.close();
      throw closedError();
    }
    this.#packEngines.add(engine);
    return engine;
  }

  #pack(): Promise<FormatPack> {
    const cached = this.#formatPackPromise;
    if (cached !== undefined) return cached;
    const factory = this.#formatPackFactory;
    if (factory === undefined) return Promise.reject(closedError());
    const pending = Promise.resolve().then(factory);
    this.#formatPackPromise = pending;
    void pending.catch(() => {
      if (this.#formatPackPromise === pending) this.#formatPackPromise = undefined;
    });
    return pending;
  }

  #assertOpen(): void {
    if (this.#closed) throw closedError();
  }
}
