import { decodeSnapshot } from "./protocol.js";
import { copyByteSource, isUint8Array } from "./bytes.js";
import type { SceneDocument } from "./scene.js";
import type { ResourceLimits, WasmSource } from "./types.js";
import { OfficeEngineError } from "./types.js";

const CORE_ABI_VERSION = 11;
const CALC_ABI_VERSION = 3;
const UNIT_ALREADY_LOADED = 0xffff_ffff;
const UNIT_LOAD_ALL_REQUIRED = 0xffff_fffe;
const UNIT_REPLACE_SNAPSHOT_REQUIRED = 0xffff_fffd;
const compiledModules = new Map<string, Promise<WebAssembly.Module>>();
const HARD_LIMITS: ResourceLimits = Object.freeze({
  inputBytes: 512 * 1024 * 1024,
  zipEntries: 50_000,
  inflatedBytes: 1024 * 1024 * 1024,
  entryBytes: 256 * 1024 * 1024,
  compressionRatio: 1_000,
  xmlBytes: 256 * 1024 * 1024,
  xmlDepth: 256,
  xmlNodes: 10_000_000,
  relationshipEdges: 200_000,
  documentObjects: 4_000_000,
  renderPixels: 256_000_000,
  imagePixels: 256_000_000,
  totalImagePixels: 512_000_000,
  fontBytes: 256 * 1024 * 1024,
});

export const DEFAULT_LIMITS: ResourceLimits = Object.freeze({
  inputBytes: 128 * 1024 * 1024,
  zipEntries: 20_000,
  inflatedBytes: 512 * 1024 * 1024,
  entryBytes: 128 * 1024 * 1024,
  compressionRatio: 200,
  xmlBytes: 64 * 1024 * 1024,
  xmlDepth: 128,
  xmlNodes: 2_000_000,
  relationshipEdges: 50_000,
  documentObjects: 1_000_000,
  renderPixels: 32_000_000,
  imagePixels: 64_000_000,
  totalImagePixels: 128_000_000,
  fontBytes: 64 * 1024 * 1024,
});

export function assertInputSize(byteLength: number, limit: number): void {
  if (byteLength > limit) {
    throw new OfficeEngineError(
      "INPUT_SIZE_LIMIT",
      `Input is ${byteLength} bytes; limit is ${limit}`,
    );
  }
}

export function encodePdfPassword(password: string | undefined): Uint8Array | undefined {
  if (password === undefined) return undefined;
  if (typeof password !== "string") {
    throw new OfficeEngineError("INVALID_PASSWORD", "PDF password must be a string");
  }
  const bytes = new TextEncoder().encode(password);
  if (bytes.byteLength > 1_024) {
    throw new OfficeEngineError("INVALID_PASSWORD", "PDF password must encode to at most 1024 bytes");
  }
  return bytes;
}

export function resolveLimits(overrides: Partial<ResourceLimits> | undefined): ResourceLimits {
  for (const key of Object.keys(overrides ?? {})) {
    if (!Object.hasOwn(HARD_LIMITS, key)) {
      throw new OfficeEngineError("INVALID_LIMITS", `Unknown resource limit ${key}`);
    }
  }
  const limits = { ...DEFAULT_LIMITS, ...overrides };
  for (const key of Object.keys(limits) as (keyof ResourceLimits)[]) {
    const value = limits[key];
    if (!Number.isSafeInteger(value) || value <= 0 || value > HARD_LIMITS[key]) {
      throw new OfficeEngineError(
        "INVALID_LIMITS",
        `${key} must be a positive integer no greater than ${HARD_LIMITS[key]}`,
      );
    }
  }
  if (limits.entryBytes > limits.inflatedBytes) {
    throw new OfficeEngineError("INVALID_LIMITS", "entryBytes cannot exceed inflatedBytes");
  }
  if (limits.xmlBytes > limits.entryBytes) {
    throw new OfficeEngineError("INVALID_LIMITS", "xmlBytes cannot exceed entryBytes");
  }
  if (limits.imagePixels > limits.totalImagePixels) {
    throw new OfficeEngineError("INVALID_LIMITS", "imagePixels cannot exceed totalImagePixels");
  }
  return Object.freeze(limits);
}

interface CoreExports extends WebAssembly.Exports {
  readonly memory: WebAssembly.Memory;
  readonly ov_core_abi_version: () => number;
  readonly ov_alloc: (length: number) => number;
  readonly ov_free: (pointer: number, length: number) => void;
  readonly ov_document_open: (
    inputPointer: number,
    inputLength: number,
    limitsPointer: number,
    limitsCount: number,
    passwordPointer: number,
    passwordLength: number,
    fieldDate: number,
    fieldTime: number,
  ) => number;
  readonly ov_document_open_with_font_metrics: (
    inputPointer: number,
    inputLength: number,
    limitsPointer: number,
    limitsCount: number,
    metricsPointer: number,
    metricsLength: number,
    passwordPointer: number,
    passwordLength: number,
    fieldDate: number,
    fieldTime: number,
  ) => number;
  readonly ov_document_snapshot: (handle: number) => number;
  readonly ov_result_pointer: () => number;
  readonly ov_result_clear: () => void;
  readonly ov_document_hit_test: (
    handle: number,
    unitIndex: number,
    x: number,
    y: number,
    outputPointer: number,
    capacity: number,
  ) => number;
  readonly ov_document_load_unit: (handle: number, unitIndex: number) => number;
  readonly ov_document_load_unit_region: (
    handle: number,
    unitIndex: number,
    maxRow: number,
    maxColumn: number,
  ) => number;
  readonly ov_document_load_all: (handle: number) => number;
  readonly ov_document_requires_calculation: (handle: number) => number;
  readonly ov_document_apply_calculation: (handle: number, pointer: number, length: number) => number;
  readonly ov_document_close: (handle: number) => void;
}

interface CalculationExports extends WebAssembly.Exports {
  readonly memory: WebAssembly.Memory;
  readonly ov_calc_abi_version: () => number;
  readonly ov_alloc: (length: number) => number;
  readonly ov_free: (pointer: number, length: number) => void;
  readonly ov_calculate: (
    input: number,
    length: number,
    limits: number,
    limitCount: number,
    localTimeMillis: number,
  ) => number;
  readonly ov_result_pointer: () => number;
  readonly ov_result_clear: () => void;
}

function isCoreExports(exports: WebAssembly.Exports): exports is CoreExports {
  const candidate = exports as Partial<CoreExports>;
  return candidate.memory instanceof WebAssembly.Memory
    && typeof candidate.ov_core_abi_version === "function"
    && typeof candidate.ov_alloc === "function"
    && typeof candidate.ov_free === "function"
    && typeof candidate.ov_document_open === "function"
    && typeof candidate.ov_document_open_with_font_metrics === "function"
    && typeof candidate.ov_document_snapshot === "function"
    && typeof candidate.ov_result_pointer === "function"
    && typeof candidate.ov_result_clear === "function"
    && typeof candidate.ov_document_hit_test === "function"
    && typeof candidate.ov_document_load_unit === "function"
    && typeof candidate.ov_document_load_unit_region === "function"
    && typeof candidate.ov_document_load_all === "function"
    && typeof candidate.ov_document_requires_calculation === "function"
    && typeof candidate.ov_document_apply_calculation === "function"
    && typeof candidate.ov_document_close === "function";
}

export function coreImports(module: WebAssembly.Module): WebAssembly.Imports {
  const imports = WebAssembly.Module.imports(module);
  if (imports.length === 0) return {};
  const placeholder: Record<string, WebAssembly.ImportValue> = {};
  const externref: Record<string, WebAssembly.ImportValue> = {};
  const allowedPlaceholderPrefixes = [
    "__wbg_ic_tz_parts_",
    "__wbg_length_",
    "__wbg_get_",
    "__wbindgen_object_drop_ref",
    "__wbg_ic_tz_validate_",
    "__wbg_ic_tz_all_",
    "__wbg_get_unchecked_",
    "__wbindgen_describe",
    "__wbg_random_",
    "__wbg___wbindgen_number_get_",
    "__wbg___wbindgen_throw_",
    "__wbg___wbindgen_string_get_",
  ];
  for (const entry of imports) {
    if (entry.kind !== "function") {
      throw new OfficeEngineError("CORE_LOAD_FAILED", "The Wasm core requests a forbidden host import");
    }
    if (entry.module === "__wbindgen_placeholder__"
      && allowedPlaceholderPrefixes.some((prefix) => entry.name.startsWith(prefix))) {
      if (entry.name.startsWith("__wbg_ic_tz_validate_")) placeholder[entry.name] = () => 1;
      else if (entry.name.startsWith("__wbg_random_")) placeholder[entry.name] = () => Number.NaN;
      else if (entry.name.startsWith("__wbg___wbindgen_throw_")) {
        placeholder[entry.name] = () => { throw new Error("Unsupported formula runtime operation"); };
      } else placeholder[entry.name] = () => 0;
      continue;
    }
    if (entry.module === "__wbindgen_externref_xform__"
      && (entry.name === "__wbindgen_externref_table_grow"
        || entry.name === "__wbindgen_externref_table_set_null")) {
      externref[entry.name] = () => 0;
      continue;
    }
    throw new OfficeEngineError("CORE_LOAD_FAILED", "The Wasm core requests a forbidden host import");
  }
  return {
    ...(Object.keys(placeholder).length === 0 ? {} : { __wbindgen_placeholder__: placeholder }),
    ...(Object.keys(externref).length === 0 ? {} : { __wbindgen_externref_xform__: externref }),
  };
}

export async function compileCore(source: WasmSource | undefined): Promise<WebAssembly.Module> {
  if (source instanceof WebAssembly.Module) return source;
  if (source === undefined || typeof source === "string" || source instanceof URL) {
    const location = source ?? new URL("./office-viewer-core.wasm", import.meta.url);
    const key = location instanceof URL ? location.href : location;
    const cached = compiledModules.get(key);
    if (cached !== undefined) return cached;
    const compiling = (async () => {
      let response: Response;
      try {
        response = await fetch(location);
      } catch (cause) {
        throw new OfficeEngineError("CORE_LOAD_FAILED", "Could not load the Wasm core", { cause });
      }
      if (!response.ok) {
        throw new OfficeEngineError("CORE_LOAD_FAILED", `Could not load Wasm core: HTTP ${response.status}`);
      }
      const fallback = response.clone();
      if (typeof WebAssembly.compileStreaming === "function") {
        try {
          return await WebAssembly.compileStreaming(response);
        } catch {
          // Some servers omit application/wasm; buffered compilation remains compatible.
        }
      }
      try {
        return await WebAssembly.compile(await fallback.arrayBuffer());
      } catch (cause) {
        throw new OfficeEngineError("CORE_LOAD_FAILED", "Could not compile the Wasm core", { cause });
      }
    })();
    compiledModules.set(key, compiling);
    try {
      return await compiling;
    } catch (cause) {
      if (compiledModules.get(key) === compiling) compiledModules.delete(key);
      if (cause instanceof OfficeEngineError) throw cause;
      throw new OfficeEngineError("CORE_LOAD_FAILED", "Could not load or compile the Wasm core", { cause });
    }
  }
  let bytes: ArrayBuffer;
  try {
    if (isUint8Array(source)) {
      bytes = copyByteSource(source);
    } else {
      bytes = source.slice(0);
    }
  } catch (cause) {
    if (cause instanceof OfficeEngineError) throw cause;
    throw new OfficeEngineError("CORE_LOAD_FAILED", "Could not load the Wasm core", { cause });
  }
  try {
    return await WebAssembly.compile(bytes);
  } catch (cause) {
    throw new OfficeEngineError("CORE_LOAD_FAILED", "Could not compile the Wasm core", { cause });
  }
}

export function limitWords(limits: ResourceLimits): Uint32Array {
  return Uint32Array.of(
    limits.inputBytes,
    limits.zipEntries,
    limits.entryBytes,
    limits.inflatedBytes,
    limits.compressionRatio,
    limits.xmlBytes,
    limits.xmlDepth,
    limits.xmlNodes,
    limits.relationshipEdges,
    limits.documentObjects,
    limits.renderPixels,
    limits.fontBytes,
  );
}

export async function calculateSpreadsheet(
  source: WasmSource | undefined,
  bytes: Uint8Array,
  limits: ResourceLimits,
): Promise<Uint8Array> {
  const module = await compileCore(source ?? new URL("./office-viewer-calc.wasm", import.meta.url));
  const instance = await WebAssembly.instantiate(module, coreImports(module));
  const exports = instance.exports as Partial<CalculationExports>;
  if (!(exports.memory instanceof WebAssembly.Memory)
    || typeof exports.ov_calc_abi_version !== "function"
    || exports.ov_calc_abi_version() !== CALC_ABI_VERSION
    || typeof exports.ov_alloc !== "function"
    || typeof exports.ov_free !== "function"
    || typeof exports.ov_calculate !== "function"
    || typeof exports.ov_result_pointer !== "function"
    || typeof exports.ov_result_clear !== "function") {
    throw new OfficeEngineError("CORE_ABI_MISMATCH", "The calculation Wasm does not expose the expected ABI");
  }
  const calc = exports as CalculationExports;
  const words = limitWords(limits);
  const inputLength = Math.max(bytes.byteLength, 1);
  const input = calc.ov_alloc(inputLength) >>> 0;
  const limitPointer = calc.ov_alloc(words.byteLength) >>> 0;
  const rangeIsValid = (pointer: number, length: number, alignment = 1) => (
    pointer !== 0
    && pointer % alignment === 0
    && pointer + length <= calc.memory.buffer.byteLength
  );
  if (!rangeIsValid(input, inputLength, 8) || !rangeIsValid(limitPointer, words.byteLength, 8)) {
    if (input !== 0) calc.ov_free(input, inputLength);
    if (limitPointer !== 0) calc.ov_free(limitPointer, words.byteLength);
    throw new OfficeEngineError("CORE_ALLOCATION_FAILED", "The calculation Wasm could not allocate input memory");
  }
  try {
    new Uint8Array(calc.memory.buffer, input, bytes.byteLength).set(bytes);
    new Uint32Array(calc.memory.buffer, limitPointer, words.length).set(words);
    const now = new Date();
    const localTimeMillis = Date.UTC(
      now.getFullYear(),
      now.getMonth(),
      now.getDate(),
      now.getHours(),
      now.getMinutes(),
      now.getSeconds(),
      now.getMilliseconds(),
    );
    const length = calc.ov_calculate(
      input,
      bytes.byteLength,
      limitPointer,
      words.length,
      localTimeMillis,
    ) >>> 0;
    const pointer = calc.ov_result_pointer() >>> 0;
    if (length === 0 || !rangeIsValid(pointer, length)) {
      throw new OfficeEngineError("CORE_OPEN_FAILED", "The calculation module could not evaluate this workbook");
    }
    return new Uint8Array(calc.memory.buffer.slice(pointer, pointer + length));
  } finally {
    calc.ov_result_clear();
    calc.ov_free(input, inputLength);
    calc.ov_free(limitPointer, words.byteLength);
  }
}

function localFieldDateTime(date = new Date()): readonly [number, number] {
  return [
    date.getFullYear() * 10_000 + (date.getMonth() + 1) * 100 + date.getDate(),
    date.getHours() * 100 + date.getMinutes(),
  ];
}

export class Core {
  #exports: CoreExports | undefined;
  readonly #limits: ResourceLimits;
  #documentCount = 0;
  #closed = false;

  private constructor(exports: CoreExports, limits: ResourceLimits) {
    this.#exports = exports;
    this.#limits = limits;
  }

  static async create(source: WasmSource | undefined, limits: ResourceLimits): Promise<Core> {
    const module = await compileCore(source);
    let instance: WebAssembly.Instance;
    try {
      instance = await WebAssembly.instantiate(module, coreImports(module));
    } catch (cause) {
      throw new OfficeEngineError("CORE_LOAD_FAILED", "Could not instantiate the Wasm core", { cause });
    }
    if (!isCoreExports(instance.exports)) {
      throw new OfficeEngineError("CORE_ABI_MISMATCH", "The Wasm core does not expose the expected ABI");
    }
    let abiVersion: number;
    try {
      abiVersion = instance.exports.ov_core_abi_version();
    } catch (cause) {
      throw new OfficeEngineError("CORE_ABI_MISMATCH", "Could not read the Wasm core ABI version", { cause });
    }
    if (abiVersion !== CORE_ABI_VERSION) {
      throw new OfficeEngineError(
        "CORE_ABI_MISMATCH",
        `Wasm core ABI ${abiVersion} is incompatible with adapter ABI ${CORE_ABI_VERSION}`,
      );
    }
    return new Core(instance.exports, limits);
  }

  get limits(): ResourceLimits {
    return this.#limits;
  }

  open(
    bytes: Uint8Array,
    fontMetrics?: Uint8Array,
    password?: Uint8Array,
    fieldDateTime = localFieldDateTime(),
  ): CoreDocument {
    this.#assertOpen();
    const exports = this.#getExports();
    assertInputSize(bytes.byteLength, this.#limits.inputBytes);
    if (fontMetrics !== undefined && !isUint8Array(fontMetrics)) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Font metrics must be a Uint8Array");
    }
    if ((fontMetrics?.byteLength ?? 0) > this.#limits.fontBytes) {
      throw new OfficeEngineError(
        "FONT_BYTES_LIMIT",
        `Font metrics are ${fontMetrics?.byteLength ?? 0} bytes; limit is ${this.#limits.fontBytes}`,
      );
    }
    if (password !== undefined && !isUint8Array(password)) {
      throw new OfficeEngineError("INVALID_PASSWORD", "PDF password must encode to a Uint8Array");
    }
    const inputAllocationLength = Math.max(bytes.byteLength, 1);
    const input = this.#allocate(inputAllocationLength);
    const limits = limitWords(this.#limits);
    let limitAllocation = 0;
    let metricAllocation = 0;
    let passwordAllocation = 0;
    let passwordAllocationLength = 0;
    let handle = 0;
    let inputConsumed = false;
    try {
      limitAllocation = this.#allocate(limits.byteLength);
      if (fontMetrics !== undefined && fontMetrics.byteLength !== 0) {
        metricAllocation = this.#allocate(fontMetrics.byteLength);
      }
      new Uint8Array(exports.memory.buffer, input, bytes.byteLength).set(bytes);
      new Uint32Array(exports.memory.buffer, limitAllocation, limits.length).set(limits);
      if (metricAllocation !== 0 && fontMetrics !== undefined) {
        new Uint8Array(exports.memory.buffer, metricAllocation, fontMetrics.byteLength).set(fontMetrics);
      }
      if (password !== undefined) {
        passwordAllocationLength = Math.max(password.byteLength, 1);
        passwordAllocation = this.#allocate(passwordAllocationLength);
        new Uint8Array(exports.memory.buffer, passwordAllocation, password.byteLength).set(password);
      }
      const opened = metricAllocation === 0
        ? exports.ov_document_open(
            input,
            bytes.byteLength,
            limitAllocation,
            limits.length,
            passwordAllocation,
            password?.byteLength ?? 0,
            fieldDateTime[0],
            fieldDateTime[1],
          )
        : exports.ov_document_open_with_font_metrics(
            input,
            bytes.byteLength,
            limitAllocation,
            limits.length,
            metricAllocation,
            fontMetrics?.byteLength ?? 0,
            passwordAllocation,
            password?.byteLength ?? 0,
            fieldDateTime[0],
            fieldDateTime[1],
          );
      inputConsumed = true;
      handle = this.#u32(opened, "document handle");
    } finally {
      if (!inputConsumed) exports.ov_free(input, inputAllocationLength);
      if (limitAllocation !== 0) exports.ov_free(limitAllocation, limits.byteLength);
      if (metricAllocation !== 0 && fontMetrics !== undefined) {
        exports.ov_free(metricAllocation, fontMetrics.byteLength);
      }
      if (passwordAllocation !== 0 && password !== undefined) {
        new Uint8Array(exports.memory.buffer, passwordAllocation, passwordAllocationLength).fill(0);
        exports.ov_free(passwordAllocation, passwordAllocationLength);
      }
    }
    if (handle === 0) {
      throw new OfficeEngineError("CORE_OPEN_FAILED", "The Wasm core could not allocate a document handle");
    }
    let scene: SceneDocument;
    try {
      scene = this.snapshot(handle);
    } catch (cause) {
      try {
        exports.ov_document_close(handle);
      } catch {
        // Preserve the snapshot failure while still attempting to release ownership.
      }
      throw cause;
    }
    this.#documentCount += 1;
    return new CoreDocument(this, handle, scene);
  }

  snapshot(handle: number): SceneDocument {
    const exports = this.#getExports();
    const length = this.#u32(exports.ov_document_snapshot(handle), "snapshot length");
    return this.#decodeResult(length);
  }

  #decodeResult(length: number): SceneDocument {
    const exports = this.#getExports();
    try {
      const pointer = this.#u32(exports.ov_result_pointer(), "snapshot pointer");
      if (length === 0 || pointer === 0) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "The Wasm core returned an empty snapshot");
      }
      this.#assertMemoryRange(pointer, length);
      return decodeSnapshot(new Uint8Array(exports.memory.buffer, pointer, length));
    } finally {
      exports.ov_result_clear();
    }
  }

  hitTest(handle: number, unitIndex: number, x: number, y: number, limit: number): readonly number[] {
    this.#assertOpen();
    if (!Number.isSafeInteger(unitIndex) || unitIndex < 0 || unitIndex > 0xffff_ffff) {
      throw new OfficeEngineError("INVALID_UNIT", "Hit-test unit index is outside the supported range");
    }
    if (!Number.isFinite(x) || !Number.isFinite(y)) {
      throw new OfficeEngineError("INVALID_HIT_TEST", "Hit-test coordinates must be finite numbers");
    }
    if (!Number.isInteger(limit) || limit < 1 || limit > 256) {
      throw new OfficeEngineError("INVALID_HIT_TEST", "Hit-test limit must be an integer from 1 through 256");
    }
    const capacity = limit;
    const pointer = this.#allocate(capacity * 4);
    const exports = this.#getExports();
    try {
      const count = this.#u32(
        exports.ov_document_hit_test(handle, unitIndex, x, y, pointer, capacity),
        "hit-test result count",
      );
      if (count > capacity) {
        throw new OfficeEngineError(
          "CORE_PROTOCOL_INVALID",
          `The Wasm core returned ${count} hit-test results for capacity ${capacity}`,
        );
      }
      return Array.from(new Uint32Array(exports.memory.buffer, pointer, count));
    } finally {
      exports.ov_free(pointer, capacity * 4);
    }
  }

  loadUnit(handle: number, unitIndex: number): { readonly scene: SceneDocument; readonly replace: boolean } | undefined {
    this.#assertOpen();
    if (!Number.isSafeInteger(unitIndex) || unitIndex < 0 || unitIndex > 0xffff_ffff) {
      throw new OfficeEngineError("INVALID_UNIT", "Unit index is outside the supported range");
    }
    const result = this.#u32(
      this.#getExports().ov_document_load_unit(handle, unitIndex),
      "unit load result",
    );
    if (result === UNIT_ALREADY_LOADED) return undefined;
    if (result === UNIT_LOAD_ALL_REQUIRED) {
      const scene = this.loadAll(handle);
      if (scene === undefined) {
        throw new OfficeEngineError("CORE_OPEN_FAILED", `The Wasm core did not materialize unit ${unitIndex}`);
      }
      return { scene, replace: true };
    }
    if (result === UNIT_REPLACE_SNAPSHOT_REQUIRED) {
      return { scene: this.snapshot(handle), replace: true };
    }
    if (result === 0) {
      throw new OfficeEngineError("CORE_OPEN_FAILED", `The Wasm core could not load unit ${unitIndex}`);
    }
    return { scene: this.#decodeResult(result), replace: false };
  }

  loadUnitRegion(
    handle: number,
    unitIndex: number,
    maxRow: number,
    maxColumn: number,
  ): { readonly scene: SceneDocument; readonly replace: boolean } | undefined {
    this.#assertOpen();
    if (!Number.isSafeInteger(unitIndex) || unitIndex < 0 || unitIndex > 0xffff_ffff) {
      throw new OfficeEngineError("INVALID_UNIT", "Unit index is outside the supported range");
    }
    if (!Number.isSafeInteger(maxRow) || maxRow < 1 || maxRow > 0x7fff_ffff
      || !Number.isSafeInteger(maxColumn) || maxColumn < 1 || maxColumn > 0x7fff_ffff) {
      throw new OfficeEngineError("INVALID_ARGUMENT", "XLSX formula viewport bounds are invalid");
    }
    const result = this.#u32(
      this.#getExports().ov_document_load_unit_region(handle, unitIndex, maxRow, maxColumn),
      "unit region load result",
    );
    if (result === UNIT_ALREADY_LOADED) return undefined;
    if (result === UNIT_LOAD_ALL_REQUIRED) {
      const scene = this.loadAll(handle);
      if (scene === undefined) {
        throw new OfficeEngineError("CORE_OPEN_FAILED", `The Wasm core did not upgrade unit ${unitIndex}`);
      }
      return { scene, replace: true };
    }
    if (result === UNIT_REPLACE_SNAPSHOT_REQUIRED) {
      return { scene: this.snapshot(handle), replace: true };
    }
    if (result === 0) {
      throw new OfficeEngineError("CORE_OPEN_FAILED", `The Wasm core could not load unit ${unitIndex}`);
    }
    return { scene: this.#decodeResult(result), replace: false };
  }

  loadAll(handle: number): SceneDocument | undefined {
    this.#assertOpen();
    const result = this.#u32(this.#getExports().ov_document_load_all(handle), "document load result");
    if (result === 2) return undefined;
    if (result !== 1) {
      throw new OfficeEngineError("CORE_OPEN_FAILED", "The Wasm core could not load the complete document");
    }
    return this.snapshot(handle);
  }

  requiresCalculation(handle: number): boolean {
    return this.#u32(
      this.#getExports().ov_document_requires_calculation(handle),
      "calculation requirement",
    ) === 1;
  }

  applyCalculation(handle: number, results: Uint8Array): SceneDocument {
    this.#assertOpen();
    const length = Math.max(results.byteLength, 1);
    const pointer = this.#allocate(length);
    try {
      new Uint8Array(this.#getExports().memory.buffer, pointer, results.byteLength).set(results);
      if (this.#getExports().ov_document_apply_calculation(handle, pointer, results.byteLength) !== 1) {
        throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "The Wasm core rejected calculation results");
      }
      return this.snapshot(handle);
    } finally {
      this.#getExports().ov_free(pointer, length);
    }
  }

  closeDocument(handle: number): void {
    const exports = this.#exports;
    if (exports === undefined) return;
    try {
      exports.ov_document_close(handle);
    } finally {
      this.#documentCount = Math.max(0, this.#documentCount - 1);
      this.#releaseExportsIfIdle();
    }
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    this.#releaseExportsIfIdle();
  }

  #allocate(length: number): number {
    if (!Number.isSafeInteger(length) || length <= 0 || length > 0xffff_ffff) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", "Invalid Wasm allocation length");
    }
    const pointer = this.#u32(this.#getExports().ov_alloc(length), "allocation pointer");
    if (pointer === 0) {
      throw new OfficeEngineError("CORE_ALLOCATION_FAILED", `The Wasm core could not allocate ${length} bytes`);
    }
    this.#assertMemoryRange(pointer, length, 8);
    return pointer;
  }

  #assertMemoryRange(pointer: number, length: number, alignment = 1): void {
    const end = pointer + length;
    if (!Number.isSafeInteger(pointer)
      || !Number.isSafeInteger(length)
      || pointer === 0
      || length < 0
      || pointer % alignment !== 0
      || !Number.isSafeInteger(end)
      || end > this.#getExports().memory.buffer.byteLength) {
      throw new OfficeEngineError(
        "CORE_PROTOCOL_INVALID",
        `The Wasm core returned an invalid memory range (${pointer}, ${length})`,
      );
    }
  }

  #u32(value: number, field: string): number {
    if (typeof value !== "number" || !Number.isInteger(value) || value < -0x8000_0000 || value > 0xffff_ffff) {
      throw new OfficeEngineError("CORE_PROTOCOL_INVALID", `The Wasm core returned an invalid ${field}`);
    }
    return value >>> 0;
  }

  #assertOpen(): void {
    if (this.#closed) throw new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
  }

  #getExports(): CoreExports {
    const exports = this.#exports;
    if (exports === undefined) throw new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
    return exports;
  }

  #releaseExportsIfIdle(): void {
    if (this.#closed && this.#documentCount === 0) this.#exports = undefined;
  }
}

export class CoreDocument {
  #core: Core | undefined;
  #scene: SceneDocument | undefined;
  readonly handle: number;
  #closed = false;

  constructor(core: Core, handle: number, scene: SceneDocument) {
    this.#core = core;
    this.handle = handle;
    this.#scene = scene;
  }

  get scene(): SceneDocument {
    const scene = this.#scene;
    if (scene === undefined) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    return scene;
  }

  get limits(): ResourceLimits {
    const core = this.#core;
    if (core === undefined) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    return core.limits;
  }

  requiresCalculation(): boolean {
    if (this.#closed || this.#core === undefined) return false;
    return this.#core.requiresCalculation(this.handle);
  }

  applyCalculation(results: Uint8Array): void {
    if (this.#closed || this.#core === undefined) {
      throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    }
    this.#scene = this.#core.applyCalculation(this.handle, results);
  }

  hitTest(unitIndex: number, x: number, y: number, limit: number): readonly number[] {
    if (this.#closed) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    const core = this.#core;
    if (core === undefined) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    return core.hitTest(this.handle, unitIndex, x, y, limit);
  }

  loadUnit(unitIndex: number): { readonly scene: SceneDocument; readonly replace: boolean } | undefined {
    if (this.#closed) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    const core = this.#core;
    if (core === undefined) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    const unit = this.scene.info?.units[unitIndex];
    if (!Number.isInteger(unitIndex)
      || unitIndex < 0
      || unit === undefined
      || unit.index !== unitIndex) {
      throw new OfficeEngineError("INVALID_UNIT", `Document has no unit at index ${unitIndex}`);
    }
    const loaded = core.loadUnit(this.handle, unitIndex);
    if (loaded === undefined) return undefined;
    const { scene, replace } = loaded;
    if (replace) {
      this.#scene = scene;
      return loaded;
    }
    const current = this.scene;
    this.#scene = {
      fatal: current.fatal,
      ...(current.info === undefined ? {} : { info: current.info }),
      diagnostics: [...current.diagnostics, ...scene.diagnostics],
      embeddedFonts: [...current.embeddedFonts, ...scene.embeddedFonts],
      ...(current.fontAlternateNames === undefined && scene.fontAlternateNames === undefined ? {} : {
        fontAlternateNames: [...(current.fontAlternateNames ?? []), ...(scene.fontAlternateNames ?? [])],
      }),
      objects: [...current.objects, ...scene.objects],
    };
    return loaded;
  }

  loadUnitRegion(
    unitIndex: number,
    maxRow: number,
    maxColumn: number,
  ): { readonly scene: SceneDocument; readonly replace: boolean } | undefined {
    if (this.#closed) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    const core = this.#core;
    if (core === undefined) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    const unit = this.scene.info?.units[unitIndex];
    if (!Number.isInteger(unitIndex)
      || unitIndex < 0
      || unit === undefined
      || unit.index !== unitIndex) {
      throw new OfficeEngineError("INVALID_UNIT", `Document has no unit at index ${unitIndex}`);
    }
    const loaded = core.loadUnitRegion(this.handle, unitIndex, maxRow, maxColumn);
    if (loaded === undefined) return undefined;
    const { scene, replace } = loaded;
    if (replace) {
      this.#scene = scene;
      return loaded;
    }
    const current = this.scene;
    this.#scene = {
      fatal: current.fatal,
      ...(current.info === undefined ? {} : { info: current.info }),
      diagnostics: [...current.diagnostics, ...scene.diagnostics],
      embeddedFonts: [...current.embeddedFonts, ...scene.embeddedFonts],
      ...(current.fontAlternateNames === undefined && scene.fontAlternateNames === undefined ? {} : {
        fontAlternateNames: [...(current.fontAlternateNames ?? []), ...(scene.fontAlternateNames ?? [])],
      }),
      objects: [...current.objects, ...scene.objects],
    };
    return loaded;
  }

  loadAll(): SceneDocument | undefined {
    if (this.#closed) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    const core = this.#core;
    if (core === undefined) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    const scene = core.loadAll(this.handle);
    if (scene !== undefined) this.#scene = scene;
    return scene;
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    const core = this.#core;
    this.#core = undefined;
    this.#scene = undefined;
    core?.closeDocument(this.handle);
  }
}
